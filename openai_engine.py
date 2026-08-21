# -*- coding: utf-8 -*-
"""
Geracao do texto interpretativo de cada secao do relatorio usando a API paga
da OpenAI (GPT). Mesma interface e mesmo formato de saida que ollama_engine.py
(generate_full_sections -> dict `sections` para report_engine.build_pages),
trocando apenas o "motor" de geracao: em vez de um modelo local via Ollama,
usa um modelo GPT via API, o que rende textos mais ricos, coerentes e capazes
de sustentar paragrafos mais longos.

Requer a variavel de ambiente OPENAI_API_KEY (lida automaticamente de um
arquivo .env na mesma pasta, via python-dotenv).
"""
import json
import os
import re

from dotenv import load_dotenv
from openai import OpenAI

load_dotenv()

# gpt-4o escreve em portugues com muito mais nuance, coerencia entre paragrafos
# e vocabulario variado do que os modelos locais rodados via Ollama - e sustenta
# paragrafos longos (~250-320 palavras) sem perder o fio. Para testes rapidos e
# mais baratos, passe model="gpt-4o-mini" (qualidade menor, mas ainda bem acima
# do llama3:8b local).
DEFAULT_MODEL = "gpt-4o"

STYLE_GUIDE = """Você é um(a) astrólogo(a) profissional que escreve laudos de mapa astral em
português do Brasil, em tom sofisticado, caloroso, perspicaz e prático — nunca
determinista. Fale diretamente com a pessoa usando "você".

Exemplo do registro e do TAMANHO exato de parágrafo esperado (não copie o
conteúdo, apenas o estilo e o comprimento):
"O Sol em Escorpião sugere uma identidade marcada por profundidade, percepção
psicológica e necessidade de autenticidade. Você pode tender a desconfiar do
que é superficial, a investigar o que está por trás das aparências e a viver
as experiências com forte carga emocional e simbólica.

Este posicionamento indica uma natureza que se transforma ao longo da vida.
Você pode passar por fases de recolhimento, depuração e renascimento, como se
a sua força verdadeira surgisse justamente quando algo precisa ser
reconstruído com mais verdade."

Regras obrigatórias:
- Escreva em parágrafos CURTOS de 3 a 5 frases (cerca de 60 a 110 palavras
  cada). NÃO escreva um único parágrafo longo e denso quando o pedido pede
  vários parágrafos — cada um deve ser uma unidade de pensamento fechada e
  direta, não um bloco corrido de ideias emendadas.
- Seja concreto e específico (cite o signo, o planeta, a casa) em vez de
  frases vagas que serviriam para qualquer pessoa — mas não tente encher o
  parágrafo com exemplos demais só para alongar; prefira precisão a volume.
  NUNCA repita palavra por palavra uma frase ou abertura só para bater a
  contagem de palavras.
- Varie as aberturas de frase e a estrutura sintática ao longo da resposta.
- Evite clichês de horóscopo de revista e repetição de vocabulário entre
  parágrafos; varie sinônimos e imagens.
- Nunca use markdown (sem **, #, listas com traço).
- Escreva apenas parágrafos corridos, no português informado, sem saudação, sem
  introdução do tipo "claro, aqui está", sem repetir o pedido.
- A contagem de palavras pedida em cada bloco é uma referência aproximada de
  tamanho (não um mínimo rígido a ser inflado com enchimento) — fique perto
  dela, nem muito curto nem exageradamente mais longo.
- Responda SOMENTE com o conteúdo pedido, nada mais."""

PET_SYNASTRY_STYLE_GUIDE = """Você é um(a) astrólogo(a) profissional que escreve relatórios de sinastria entre
tutor(a) e pet, em português do Brasil. Ao falar do(a) tutor(a), use a 2ª pessoa
("você") e o mesmo registro sofisticado, perspicaz e denso de um laudo individual.
Ao falar do pet, mantenha um tom caloroso, leve e com humor sutil, e NUNCA atribua
cognição ou motivação humana literal ao animal — descreva sempre respeitando a
natureza animal (instinto, vínculo, comportamento, energia). O relatório deve ler
como a leitura de uma RELAÇÃO entre os dois, não dois textos colados — sempre que
possível, relacione o traço do tutor ao traço do pet.

Regras obrigatórias:
- Escreva em parágrafos CURTOS de 3 a 5 frases (cerca de 60 a 110 palavras
  cada), com exemplos concretos do cotidiano (rotina, comportamento,
  convivência), em vez de um bloco denso e longo ou de frases genéricas.
- Varie as aberturas de frase e a estrutura sintática. NÃO comece mais de um
  parágrafo com "Esta posição", "Essa posição" ou "Isso pode indicar/sugerir" —
  use no máximo uma vez por resposta.
- Prefira frases concretas e específicas (cite o signo, o planeta, a casa) a
  frases vagas que serviriam para qualquer dupla. Não repita frases só para
  bater a contagem de palavras.
- Evite clichês de horóscopo de revista e repetição de vocabulário.
- Nunca use markdown (sem **, #, listas com traço).
- Escreva apenas parágrafos corridos, sem saudação, sem introdução tipo "claro,
  aqui está", sem repetir o pedido.
- A contagem de palavras pedida em cada bloco é uma referência aproximada de
  tamanho, não um mínimo rígido para encher com enchimento.
- Responda SOMENTE com o conteúdo pedido, nada mais."""

def _get_client():
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError(
            "OPENAI_API_KEY nao encontrada. Configure-a no arquivo .env desta pasta "
            "(OPENAI_API_KEY=sk-...) ou como variavel de ambiente."
        )
    return OpenAI(api_key=api_key)

def _call_openai(prompt, model=DEFAULT_MODEL, temperature=0.8, max_tokens=4000):
    client = _get_client()
    resp = client.chat.completions.create(
        model=model,
        messages=[{"role": "user", "content": prompt}],
        temperature=temperature,
        max_tokens=max_tokens,
    )
    return (resp.choices[0].message.content or "").strip()

def _extract_word_targets(prompt, expected_keys):
    """Le as instrucoes '~N palavras' logo apos cada marcador ###CHAVE### no
    prompt, para saber quantas palavras cobrar de cada bloco gerado."""
    targets = {}
    pattern = r'###\s*(' + '|'.join(re.escape(k) for k in expected_keys) + r')\s*###(.*?)(?=###|\Z)'
    for m in re.finditer(pattern, prompt, flags=re.DOTALL):
        key, instr = m.group(1), m.group(2)
        wm = re.search(r'~\s*(\d+)\s*palavras', instr)
        targets[key] = int(wm.group(1)) if wm else 12
    return targets

def _meets_target(text, target_words):
    text = (text or "").strip()
    if not text:
        return False
    n = len(text.split())
    return n >= max(6, round(target_words * 0.6))

def _ensure_leadin(text, leadin):
    """Garante que o paragrafo comece com o lead-in em caixa-alta esperado
    (ex: 'SOL EM TOURO NA 6A CASA.'), independente do modelo ter seguido a
    instrucao — mais confiavel do que confiar no LLM para formatacao exata."""
    text = (text or "").strip()
    if not text:
        return text
    if text.upper().startswith(leadin.upper()[:20]):
        return text
    return f"{leadin} {text}"

def _clean(text):
    text = re.sub(r'\*\*|##+|^-\s+', '', text, flags=re.MULTILINE)
    text = text.strip().strip('"')
    return text

def _split_markers(text, keys):
    """Extrai blocos delimitados por linhas '###CHAVE###' do texto gerado."""
    out = {}
    pattern = r'###\s*(' + '|'.join(re.escape(k) for k in keys) + r')\s*###'
    parts = re.split(pattern, text)
    # parts[0] eh lixo antes do primeiro marcador; depois alterna chave, conteudo, chave, conteudo...
    it = iter(parts[1:])
    for key, content in zip(it, it):
        out[key] = _clean(content)
    return out

def _planet_line(chart):
    lines = []
    for p in chart["planets"]:
        lines.append(f'{p["name"]}: {p["deg"]}°{p["min"]:02d} {p["sign"]} (casa {p["house"]})')
    return "\n".join(lines)

def _aspects_line(aspects, harmonic):
    rows = [a for a in aspects if a["harmonico"] == harmonic][:7]
    return "\n".join(f'{a["aspecto"]} (orbe {a["orbe"]})' for a in rows) if rows else "nenhum aspecto relevante encontrado"

_ELEMENT_OF_SIGN = {
    "Áries": "Fogo", "Leão": "Fogo", "Sagitário": "Fogo",
    "Touro": "Terra", "Virgem": "Terra", "Capricórnio": "Terra",
    "Gêmeos": "Ar", "Libra": "Ar", "Aquário": "Ar",
    "Câncer": "Água", "Escorpião": "Água", "Peixes": "Água",
}

def _element_balance_line(chart):
    """Resume o balanço de elementos (fogo/terra/ar/agua) e a casa com mais
    planetas, para dar ao texto da Roda Natal algo concreto para comentar
    sobre o desenho geral do mapa, em vez de uma legenda generica fixa."""
    core = [p for p in chart["planets"] if p["name"] not in ("Meio do Céu", "Ascendente")]
    counts = {}
    for p in core:
        el = _ELEMENT_OF_SIGN.get(p["sign"])
        if el:
            counts[el] = counts.get(el, 0) + 1
    elements_line = ", ".join(f"{el}: {n}" for el, n in sorted(counts.items(), key=lambda x: -x[1]))

    house_counts = {}
    for p in core:
        h = p.get("house")
        if h:
            house_counts[h] = house_counts.get(h, 0) + 1
    top_houses = sorted(house_counts.items(), key=lambda x: -x[1])[:2]
    houses_line = ", ".join(f"casa {h} ({n} planetas)" for h, n in top_houses) if top_houses else "sem concentração marcante"

    return f"Elementos: {elements_line}.\nCasas com mais planetas: {houses_line}."

def _house_overlay_line(overlay):
    rows = overlay[:6]
    return "\n".join(f'{o["planet"]} na casa {o["house_in_b"]}' for o in rows) if rows else "nenhuma sobreposição relevante encontrada"

def _run_steps(steps, model, progress_cb=None):
    """Executa cada (key, prompt) de `steps` via _call_openai, tentando ate 2
    temperaturas e ficando com a que mais atingiu a contagem de palavras pedida
    em cada bloco ###MARCADOR###. Retorna {step_key: {MARCADOR: texto}}."""
    total = len(steps)
    results = {}
    for i, (key, prompt) in enumerate(steps, start=1):
        if progress_cb:
            progress_cb(i, total, key)
        expected = re.findall(r'###\s*([A-Z0-9_]+)\s*###', prompt)
        word_targets = _extract_word_targets(prompt, expected)
        parsed = {}
        best_score = (-1, -1)
        temperatures = [0.8, 0.6]
        for attempt in range(len(temperatures)):
            raw = _call_openai(prompt, model=model, temperature=temperatures[attempt], max_tokens=4000)
            found_keys = re.findall(r'###\s*([A-Z0-9_]+)\s*###', raw)
            new_parsed = _split_markers(raw, found_keys)
            satisfied = sum(1 for k in expected if _meets_target(new_parsed.get(k, ""), word_targets.get(k, 12)))
            total_words = sum(len(new_parsed.get(k, "").split()) for k in expected)
            score = (satisfied, total_words)
            if score > best_score:
                parsed, best_score = new_parsed, score
            if satisfied == len(expected):
                break
        results[key] = parsed
    return results

def generate_full_sections(chart, model=DEFAULT_MODEL, progress_cb=None):
    """chart: dict retornado por chart_engine.compute_chart (inclui 'aspects').
    progress_cb(step, total, label): opcional, chamado a cada bloco concluido.
    Retorna dict pronto para report_engine (`data['sections']`)."""
    name = chart["name"]
    birth = chart["birth"]
    planets_txt = _planet_line(chart)

    def pget(pname):
        return next(p for p in chart["planets"] if p["name"] == pname)

    sol, lua, asc = pget("Sol"), pget("Lua"), pget("Ascendente")

    sections = {}
    steps = []

    # --- bloco 1: introducao ---
    steps.append(("intro", f"""{STYLE_GUIDE}

Dados: {name}, nascido(a) em {birth['date']} às {birth['time']}, em {birth['place']}.
Tríade principal: Sol em {sol['sign']}, Lua em {lua['sign']}, Ascendente em {asc['sign']}.

Escreva 4 blocos de texto separados por marcadores exatamente neste formato:

###OBSERVACAO###
(1 parágrafo curto, ~90 palavras, explicando que este relatório usa astrologia ocidental tropical com casas Placidus, é simbólico e não determinístico)
###P1###
(2 parágrafos curtos, ~80 palavras cada, dizendo que a pessoa não é só o signo solar, combinando os traços da tríade acima de forma específica)
###P2###
(1 parágrafo curto, ~80 palavras, convidando a pessoa a ler o relatório como uma paisagem simbólica, não uma verdade fechada)
###COMO_APROVEITAR###
(1 parágrafo curto, ~70 palavras, sugerindo como aproveitar melhor o relatório)
"""))

    elem_txt = _element_balance_line(chart)
    steps.append(("dados", f"""{STYLE_GUIDE}

Ascendente de {name}: {asc['sign']} {asc['deg']}°{asc['min']:02d}. Meio do Céu: {birth['midheaven_label']}.

###DADOS_TEXTO###
(1 parágrafo curto, ~120 palavras, interpretando o que o Ascendente em {asc['sign']} e o Meio do Céu em {pget('Meio do Céu')['sign']} sugerem sobre a primeira impressão, a postura social e a vocação pública da pessoa)
"""))

    steps.append(("roda_natal", f"""{STYLE_GUIDE}

Mapa de {name}. Balanço de elementos e casas mais ativadas:
{elem_txt}

###MAPA_TECNICO_TEXTO###
(1 parágrafo curto, ~100 palavras, comentando o formato geral do mapa — o predomínio de elementos/casas acima e o que esse desenho geral sugere sobre o temperamento da pessoa, como uma leitura de conjunto da roda natal antes de entrar em cada posição individualmente)
"""))

    steps.append(("triade", f"""{STYLE_GUIDE}

Sol em {sol['sign']} na casa {sol['house']}. Lua em {lua['sign']} na casa {lua['house']}. Ascendente em {asc['sign']}.

Para cada posição abaixo, escreva primeiro uma linha de 3 a 5 palavras-chave
(sem repetir o nome do planeta nem do signo, só as palavras-chave — ex.:
"essência, intensidade e verdade interior") e depois o parágrafo.

###SOL_HEAD###
(3 a 5 palavras-chave sobre o Sol em {sol['sign']})
###SOL_TEXT###
(2 parágrafos curtos, ~85 palavras cada, interpretando o Sol em {sol['sign']} na casa {sol['house']}, com 1 exemplo concreto do dia a dia)
###LUA_HEAD###
(3 a 5 palavras-chave sobre a Lua em {lua['sign']})
###LUA_TEXT###
(2 parágrafos curtos, ~85 palavras cada, interpretando a Lua em {lua['sign']} na casa {lua['house']}, com 1 exemplo concreto)
###ASC_HEAD###
(3 a 5 palavras-chave sobre o Ascendente em {asc['sign']})
###ASC_TEXT###
(1 a 2 parágrafos curtos, ~85 palavras ao todo, interpretando o Ascendente em {asc['sign']})
###SINTESE###
(1 parágrafo curto, ~90 palavras, sintetizando a combinação dos três elementos/temperamentos envolvidos)
"""))

    steps.append(("posicoes", f"""{STYLE_GUIDE}

Posições planetárias de {name}:
{planets_txt}

###POSICOES_TEXTO###
(1 parágrafo curto, ~120 palavras, comentando quais casas concentram mais planetas e o que isso sugere sobre os temas de vida mais ativados)
"""))

    steps.append(("casas", f"""{STYLE_GUIDE}

Ascendente {asc['sign']}, Descendente oposto. Meio do Céu {pget('Meio do Céu')['sign']}, Fundo do Céu oposto.

###CASAS_TEXTO###
(1 parágrafo curto, ~100 palavras, introduzindo o que as casas mostram sobre onde a energia da pessoa ganha palco, citando quais casas do mapa dela têm mais peso)
###EIXO1_TEXT###
(1 parágrafo curto, ~100 palavras, sobre o eixo Ascendente-Descendente e o que ele revela sobre presença e relações)
###EIXO2_TEXT###
(1 parágrafo curto, ~100 palavras, sobre o eixo Fundo do Céu-Meio do Céu e o que ele revela sobre raízes e vocação pública)
"""))

    topic_specs = [
        ("personalidade", "PERSONALIDADE", 4, "personalidade e traços centrais, combinando Sol, Lua e Ascendente"),
        ("emocional", "EMOCIONAL", 3, "mundo emocional e relacionamentos, a partir da Lua e de Vênus"),
        ("comunicacao", "COMUNICACAO", 3, "comunicação, pensamento e forma de decidir, a partir de Mercúrio"),
        ("afeto", "AFETO", 3, "amor, afeto, desejo e valores pessoais, a partir de Vênus e Marte"),
        ("carreira", "CARREIRA", 3, "carreira, imagem pública e propósito, a partir do Meio do Céu e Marte"),
    ]
    for key, marker, n_paras, desc in topic_specs:
        paras_block = "\n".join(f"###{marker}_P{i+1}###\n(1 parágrafo curto, ~80 palavras, específico — cite signos, planetas e casas)" for i in range(n_paras))
        steps.append((key, f"""{STYLE_GUIDE}

Mapa de {name}: Sol {sol['sign']}, Lua {lua['sign']}, Ascendente {asc['sign']}.
Planetas relevantes:
{planets_txt}

Escreva sobre: {desc}, em {n_paras} parágrafos curtos e independentes (não um
bloco corrido — cada parágrafo deve poder ser lido isoladamente).

{paras_block}
###{marker}_B1###
(1 frase direta e prática, ~18 palavras, um insight-resumo sobre esse tema — sem começar com "Esta posição" ou "Essa posição")
###{marker}_B2###
(1 frase direta e prática, ~18 palavras, outro insight-resumo diferente do anterior)
###{marker}_B3###
(1 frase direta e prática, ~18 palavras, outro insight-resumo diferente dos anteriores)
"""))

    steps.append(("talentos", f"""{STYLE_GUIDE}

Planetas de {name}:
{planets_txt}

Liste 6 talentos naturais, cada um em uma linha no formato exato:
###T1_LEAD### (3-6 palavras, tipo título do talento)
###T1_TEXT### (1 a 2 frases, ~30 palavras, explicando com um exemplo concreto)
###T2_LEAD###
###T2_TEXT###
###T3_LEAD###
###T3_TEXT###
###T4_LEAD###
###T4_TEXT###
###T5_LEAD###
###T5_TEXT###
###T6_LEAD###
###T6_TEXT###
###TALENTOS_CALLOUT###
(1 frase de destaque, ~25 palavras, sobre o talento central da pessoa)
"""))

    steps.append(("crescimento", f"""{STYLE_GUIDE}

Planetas de {name}:
{planets_txt}

Liste 4 oportunidades de crescimento pessoal, formato exato:
###G1_TITLE### (título curto em caixa alta, 3-6 palavras)
###G1_TEXT### (1 parágrafo curto, ~90 palavras, com 1 exemplo concreto)
###G2_TITLE###
###G2_TEXT###
###G3_TITLE###
###G3_TEXT###
###G4_TITLE###
###G4_TEXT###
"""))

    harm_list = [a for a in chart["aspects"] if a["harmonico"]][:7]
    tense_list = [a for a in chart["aspects"] if not a["harmonico"]][:6]
    harm_numbered = "\n".join(f"{i+1}. {a['aspecto']} (orbe {a['orbe']})" for i, a in enumerate(harm_list)) or "nenhum"
    tense_numbered = "\n".join(f"{i+1}. {a['aspecto']} (orbe {a['orbe']})" for i, a in enumerate(tense_list)) or "nenhum"
    harm_markers = "\n".join(f"###AH{i+1}_TEXT### (2 a 3 frases, ~45 palavras, comentando especificamente o aspecto {i+1} da lista de harmônicos)" for i in range(len(harm_list)))
    tense_markers = "\n".join(f"###AT{i+1}_TEXT### (2 a 3 frases, ~45 palavras, comentando especificamente o aspecto {i+1} da lista de tensos, sem tom de alarme, com um caminho prático)" for i in range(len(tense_list)))
    steps.append(("aspectos", f"""{STYLE_GUIDE}

Aspectos harmônicos de {name}, nesta ordem:
{harm_numbered}

Aspectos tensos, nesta ordem:
{tense_numbered}

Para cada aspecto, escreva um comentário curto e independente (não repita o
nome do aspecto na frase — comece direto pelo que ele sugere).

{harm_markers}
{tense_markers}
"""))

    steps.append(("previsao", f"""{STYLE_GUIDE}

Mapa de {name}: Sol {sol['sign']}, Lua {lua['sign']}, Ascendente {asc['sign']}.

Escreva a visão geral da previsão simbólica para os próximos 12 meses, em 4 blocos:

###PREV_P1###
(1 parágrafo curto, ~90 palavras, começando com "EXPANSÃO E OPORTUNIDADES." sobre o tema de crescimento do ano)
###PREV_P2###
(1 parágrafo curto, ~90 palavras, começando com "ESTRUTURA E RESPONSABILIDADE." sobre consolidar bases)
###PREV_P3###
(1 parágrafo curto, ~90 palavras, começando com "REVISÃO E SENSIBILIDADE." sobre um momento de questionamento)
###PREV_P4###
(1 parágrafo curto, ~90 palavras, começando com "TRANSFORMAÇÃO PESSOAL." sobre mudança de identidade ou postura)
"""))

    quarters = [
        ("q1", "1º trimestre, primeiros 3 meses à frente"),
        ("q2", "2º trimestre, meses 4 a 6"),
        ("q3", "3º trimestre, meses 7 a 9"),
        ("q4", "4º trimestre, meses 10 a 12"),
    ]
    for qkey, qdesc in quarters:
        steps.append((qkey, f"""{STYLE_GUIDE}

Mapa de {name}: Sol {sol['sign']}, Lua {lua['sign']}, Ascendente {asc['sign']}.
Escreva a previsão simbólica para o {qdesc}.

###{qkey.upper()}_P1###
(1 parágrafo curto, ~90 palavras)
###{qkey.upper()}_P2###
(1 parágrafo curto, ~90 palavras)
###{qkey.upper()}_OPP_TEXT###
(2 a 3 frases, ~50 palavras, sobre a oportunidade do período)
###{qkey.upper()}_DEC_TEXT###
(2 a 3 frases, ~50 palavras, sobre uma decisão prática recomendada)
"""))

    steps.append(("sintese", f"""{STYLE_GUIDE}

Mapa de {name}: Sol {sol['sign']}, Lua {lua['sign']}, Ascendente {asc['sign']}.

###SINTESE_P1###
(1 parágrafo curto, ~100 palavras, sobre a combinação de elementos do mapa)
###SINTESE_P2###
(1 parágrafo curto, ~100 palavras, sobre como a pessoa concilia forças aparentemente opostas)
###SINTESE_P3###
(1 parágrafo curto, ~90 palavras, sobre os próximos meses)
###SINTESE_QUOTE###
(1 frase poética e memorável, entre aspas, resumindo a essência da pessoa em até 20 palavras)
"""))

    steps.append(("capa", f"""{STYLE_GUIDE}

Mapa de {name}: Sol {sol['sign']}, Lua {lua['sign']}, Ascendente {asc['sign']}.

Escreva 5 textos curtos para a CAPA do relatório (cartão de resumo visual, não o corpo do texto — frases diretas, sem "esta posição", sem repetir a palavra "mapa"):

###CAPA_ABERTURA###
(1 frase inspiradora sobre o propósito de vida da pessoa, entre aspas, no mínimo ~20 palavras, combinando Sol/Lua/Ascendente — tom de destaque de capa, não de laudo)
###CAPA_PERSONALIDADE###
(2 frases curtas, no mínimo ~22 palavras, resumindo os traços centrais de personalidade a partir de Sol e Ascendente)
###CAPA_RELACIONAMENTOS###
(2 frases curtas, no mínimo ~22 palavras, resumindo o estilo afetivo e de relacionamento a partir de Vênus e Lua)
###CAPA_CRESCIMENTO###
(2 frases curtas, no mínimo ~22 palavras, sobre a principal oportunidade de crescimento pessoal)
###CAPA_FECHAMENTO###
(1 frase poética de fechamento em 2 linhas curtas, separadas por uma quebra de linha, no máximo ~16 palavras no total, sem aspas)
"""))

    steps.append(("nota_final", f"""{STYLE_GUIDE}

###NOTA_P1###
(1 parágrafo curto, ~100 palavras, sobre astrologia oferecer símbolos e a vida oferecer a experiência real)
###NOTA_P2###
(1 parágrafo curto, ~90 palavras, mensagem de encerramento inspiradora e gentil)
"""))

    results = _run_steps(steps, model, progress_cb)

    def gv(step, marker, default=""):
        return results.get(step, {}).get(marker, default)

    sections["intro_subtitle"] = "Uma leitura simbólica para orientar reflexão, não um roteiro fechado sobre o futuro."
    sections["intro_observacao"] = gv("intro", "OBSERVACAO")
    sections["intro_p1"] = gv("intro", "P1")
    sections["intro_p2"] = gv("intro", "P2")
    sections["intro_como_aproveitar"] = gv("intro", "COMO_APROVEITAR")

    sections["dados_nascimento_texto"] = gv("dados", "DADOS_TEXTO") or (
        f"O Ascendente em {asc['sign']} marca a porta de entrada da sua energia no mundo — a primeira "
        f"impressão que você causa e o estilo com que se apresenta em ambientes novos. O Meio do Céu em "
        f"{pget('Meio do Céu')['sign']} aponta para a direção da sua vocação pública, a forma como você "
        f"deseja ser reconhecido(a) profissionalmente. Juntos, esses dois pontos organizam a leitura de "
        f"todo o restante do mapa que segue nas próximas páginas.")

    sections["mapa_tecnico_texto"] = gv("roda_natal", "MAPA_TECNICO_TEXTO") or (
        "A roda natal organiza os doze signos no anel externo e as doze casas no anel interno, calculadas "
        "a partir do Ascendente. Os pontos coloridos indicam a posição simbólica de cada planeta, e a "
        "concentração deles em determinadas regiões da roda já sugere onde a sua energia se organiza com mais força.")

    def _head(step, marker, planet, sign):
        h = gv(step, marker).strip().rstrip(".")
        return f"{planet} em {sign} — {h}" if h else f"{planet} em {sign}"

    sections["triade_items"] = [
        (_head("triade", "SOL_HEAD", "Sol", sol['sign']), gv("triade", "SOL_TEXT")),
        (_head("triade", "LUA_HEAD", "Lua", lua['sign']), gv("triade", "LUA_TEXT")),
        (_head("triade", "ASC_HEAD", "Ascendente", asc['sign']), gv("triade", "ASC_TEXT")),
    ]
    sections["triade_sintese"] = gv("triade", "SINTESE")

    sections["posicoes_texto"] = gv("posicoes", "POSICOES_TEXTO")

    sections["casas_texto"] = gv("casas", "CASAS_TEXTO")
    sections["eixo1_title"] = f"Eixo Ascendente–Descendente: {asc['sign']}"
    sections["eixo1_text"] = gv("casas", "EIXO1_TEXT")
    sections["eixo2_title"] = "Eixo Fundo do Céu–Meio do Céu"
    sections["eixo2_text"] = gv("casas", "EIXO2_TEXT")

    for key, marker, n_paras, desc in topic_specs:
        paras = [gv(key, f"{marker}_P{i+1}") for i in range(n_paras)]
        if not any(p.strip() for p in paras):
            paras[0] = f"Seu mapa sugere traços marcantes ligados a {desc}, com nuances que a leitura das próximas seções aprofunda."
        sections[key] = paras
        bullets = [gv(key, f"{marker}_B{i}") for i in (1, 2, 3)]
        if not any(b.strip() for b in bullets):
            bullets = ["Observe como esses temas aparecem nas suas próprias experiências recentes."]
        sections[f"{key}_bullets"] = bullets

    talentos = []
    for i in range(1, 7):
        talentos.append({"lead": gv("talentos", f"T{i}_LEAD"), "text": gv("talentos", f"T{i}_TEXT")})
    sections["talentos"] = talentos
    sections["talentos_callout"] = gv("talentos", "TALENTOS_CALLOUT")

    crescimento = []
    for i in range(1, 5):
        crescimento.append({"title": gv("crescimento", f"G{i}_TITLE"), "text": gv("crescimento", f"G{i}_TEXT")})
    sections["crescimento"] = crescimento

    sections["aspectos_harmonicos"] = [{"aspecto": a["aspecto"], "orbe": a["orbe"]} for a in harm_list]
    sections["aspectos_harmonicos_items"] = [
        (f'{a["aspecto"]} — orbe {a["orbe"]}', gv("aspectos", f"AH{i+1}_TEXT"))
        for i, a in enumerate(harm_list)
    ]
    sections["aspectos_tensos"] = [{"aspecto": a["aspecto"], "orbe": a["orbe"]} for a in tense_list]
    sections["aspectos_tensos_items"] = [
        (f'{a["aspecto"]} — orbe {a["orbe"]}', gv("aspectos", f"AT{i+1}_TEXT"))
        for i, a in enumerate(tense_list)
    ]

    sections["previsao_subtitle"] = "Próximos 12 meses: leitura simbólica dos ciclos em curso."
    _prev_fallback = ("A leitura de trânsitos descreve climas simbólicos, não acontecimentos inevitáveis. "
        "No próximo ano, seu mapa tende a ser mobilizado por ciclos de expansão, estrutura, revisão e renovação pessoal.")
    _prev_leadins = ["EXPANSÃO E OPORTUNIDADES.", "ESTRUTURA E RESPONSABILIDADE.",
                      "REVISÃO E SENSIBILIDADE.", "TRANSFORMAÇÃO PESSOAL."]
    _prev_paras = [_ensure_leadin(gv("previsao", f"PREV_P{i}"), _prev_leadins[i-1]) for i in range(1, 5)]
    if not any(p.strip() for p in _prev_paras):
        _prev_paras[0] = _prev_fallback
    sections["previsao_overview"] = _prev_paras

    quarter_labels = [
        ("PREVISÃO - 1º TRIMESTRE", "PRÓXIMOS 3 MESES"),
        ("PREVISÃO - 2º TRIMESTRE", "MESES 4 A 6"),
        ("PREVISÃO - 3º TRIMESTRE", "MESES 7 A 9"),
        ("PREVISÃO - 4º TRIMESTRE", "MESES 10 A 12"),
    ]
    trimestres = []
    for (qkey, _qdesc), (eyebrow, subtitle) in zip(quarters, quarter_labels):
        trimestres.append({
            "eyebrow": eyebrow, "subtitle": subtitle,
            "p1": gv(qkey, f"{qkey.upper()}_P1"), "p2": gv(qkey, f"{qkey.upper()}_P2"),
            "opp_title": "Oportunidade do período", "opp_text": gv(qkey, f"{qkey.upper()}_OPP_TEXT"),
            "dec_title": "Decisão prática", "dec_text": gv(qkey, f"{qkey.upper()}_DEC_TEXT"),
        })
    sections["trimestres"] = trimestres

    _sint_paras = [gv("sintese", "SINTESE_P1"), gv("sintese", "SINTESE_P2"), gv("sintese", "SINTESE_P3")]
    if not any(p.strip() for p in _sint_paras):
        _sint_paras[0] = (f"O mapa de {name} reúne a combinação de Sol em {sol['sign']}, Lua em {lua['sign']} e "
            f"Ascendente em {asc['sign']} — uma assinatura própria de identidade, emoção e presença.")
    sections["sintese_final"] = _sint_paras
    sections["sintese_quote"] = gv("sintese", "SINTESE_QUOTE",
        f"“Você transforma o que sente em direção, e o que aprende em uma forma mais precisa de ser {name.split()[0]}.”")

    sections["capa_abertura"] = gv("capa", "CAPA_ABERTURA",
        f"“Você nasceu para transformar {sol['sign'].lower()} em direção e {lua['sign'].lower()} em presença.”")
    sections["capa_personalidade"] = gv("capa", "CAPA_PERSONALIDADE",
        f"Sol em {sol['sign']} e Ascendente em {asc['sign']}: uma combinação própria de identidade e presença.")
    sections["capa_relacionamentos"] = gv("capa", "CAPA_RELACIONAMENTOS",
        "Você valoriza vínculos verdadeiros, com espaço para individualidade e crescimento em conjunto.")
    sections["capa_crescimento"] = gv("capa", "CAPA_CRESCIMENTO",
        "Desenvolver constância e presença ajuda a transformar grandes visões em realizações duradouras.")
    sections["capa_fechamento"] = gv("capa", "CAPA_FECHAMENTO",
        f"Sua mente cria possibilidades,\nseu coração escolhe o que vale a pena realizar.")

    sections["apendice_texto"] = ("Os dados abaixo registram a estrutura técnica que serviu de base à interpretação. "
        "Pequenas variações podem ocorrer entre programas astrológicos conforme efemérides e convenções de cálculo.")

    _nota_paras = [gv("nota_final", "NOTA_P1"), gv("nota_final", "NOTA_P2")]
    if not any(p.strip() for p in _nota_paras):
        _nota_paras[0] = ("A astrologia oferece símbolos. A sua vida oferece a história, as escolhas, os vínculos, "
            "os aprendizados e a liberdade de transformar qualquer tendência em consciência.")
    sections["nota_final"] = _nota_paras
    sections["nota_quote"] = sections["sintese_quote"]

    return sections

def generate_pet_synastry_sections(owner_chart, pet_chart, cross_aspects, house_overlay_owner,
                                    house_overlay_pet, model=DEFAULT_MODEL, progress_cb=None):
    """owner_chart/pet_chart: dicts de chart_engine.compute_chart (pet_chart
    acrescido de 'breed', 'color' e 'time_estimated' pelo chamador).
    cross_aspects: chart_engine.compute_synastry_aspects(owner planets, pet planets).
    house_overlay_owner: compute_house_overlay(owner planets, pet cusps) -- planetas
    do tutor caindo nas casas do pet. house_overlay_pet: o inverso.
    Retorna dict pronto para report_engine.build_pet_synastry_pages (`data['sections']`)."""
    owner_name = owner_chart["name"]
    pet_name = pet_chart["name"]
    breed = pet_chart.get("breed", "") or "sem raça definida"
    color = pet_chart.get("color", "") or "sem cor informada"
    time_estimated = bool(pet_chart.get("time_estimated"))

    def opget(pname):
        return next(p for p in owner_chart["planets"] if p["name"] == pname)

    def ppget(pname):
        return next(p for p in pet_chart["planets"] if p["name"] == pname)

    o_sol, o_lua, o_asc = opget("Sol"), opget("Lua"), opget("Ascendente")
    p_sol, p_lua, p_asc = ppget("Sol"), ppget("Lua"), ppget("Ascendente")

    pet_planets_txt = _planet_line(pet_chart)
    cross_harm_txt = _aspects_line(cross_aspects, True)
    cross_tense_txt = _aspects_line(cross_aspects, False)
    overlay_owner_txt = _house_overlay_line(house_overlay_owner)
    overlay_pet_txt = _house_overlay_line(house_overlay_pet)

    disclaimer_note = (
        f"O horário de nascimento de {pet_name} não foi informado com exatidão; foi usado meio-dia como "
        f"estimativa, então o Ascendente e as casas de {pet_name} devem ser lidos como aproximados."
    ) if time_estimated else ""
    disclaimer_hint = f" Observação: {disclaimer_note}" if disclaimer_note else ""

    steps = []

    steps.append(("intro", f"""{PET_SYNASTRY_STYLE_GUIDE}

Dupla: {owner_name} (tutor) e {pet_name} (pet, raça {breed}, cor {color}).
Tríade do tutor: Sol em {o_sol['sign']}, Lua em {o_lua['sign']}, Ascendente em {o_asc['sign']}.
Tríade do pet: Sol em {p_sol['sign']}, Lua em {p_lua['sign']}, Ascendente em {p_asc['sign']}.{disclaimer_hint}

Escreva 3 blocos de texto separados por marcadores exatamente neste formato:

###OBSERVACAO###
(1 parágrafo curto, ~90 palavras, explicando que este relatório cruza a astrologia ocidental tropical (casas Placidus) do tutor com uma leitura simbólica do temperamento do pet, é simbólico e não determinístico{"; mencione que o horário de nascimento do pet foi estimado e por isso ascendente/casas dele são aproximados" if disclaimer_note else ""})
###APRESENTACAO###
(1 parágrafo curto, ~100 palavras, apresentando a dupla {owner_name} e {pet_name} com leveza e calor, situando o que este relatório revela sobre o vínculo entre os dois)
###COMO_APROVEITAR###
(1 parágrafo curto, ~70 palavras, sugerindo como aproveitar melhor este relatório)
"""))

    steps.append(("owner_essencia", f"""{PET_SYNASTRY_STYLE_GUIDE}

Sol de {owner_name} em {o_sol['sign']} na casa {o_sol['house']}. Lua em {o_lua['sign']} na casa {o_lua['house']}. Ascendente em {o_asc['sign']}.

###OWNER_SOL###
(2 parágrafos curtos, ~80 palavras cada, começando com "SOL EM {o_sol['sign'].upper()} NA {o_sol['house'].upper()} CASA." interpretando essa posição do tutor, com 1 exemplo concreto)
###OWNER_LUA###
(2 parágrafos curtos, ~80 palavras cada, começando com "LUA EM {o_lua['sign'].upper()} NA {o_lua['house'].upper()} CASA." interpretando essa posição do tutor)
###OWNER_ASCENDENTE###
(1 a 2 parágrafos curtos, ~80 palavras ao todo, começando com "ASCENDENTE EM {o_asc['sign'].upper()}." interpretando essa posição do tutor)
###OWNER_SINTESE###
(1 parágrafo curto, ~90 palavras, sintetizando a tríade do tutor)
"""))

    steps.append(("pet_essencia", f"""{PET_SYNASTRY_STYLE_GUIDE}

{pet_name} é da raça {breed}, cor {color}. Sol em {p_sol['sign']}, Lua em {p_lua['sign']}, Ascendente em {p_asc['sign']}.{disclaimer_hint}

###PET_SOL###
(1 a 2 parágrafos curtos, ~80 palavras ao todo, começando com "SOL EM {p_sol['sign'].upper()}." interpretando essa posição de {pet_name} em registro animal — temperamento, energia, instinto — sem atribuir cognição humana literal)
###PET_LUA###
(1 a 2 parágrafos curtos, ~80 palavras ao todo, começando com "LUA EM {p_lua['sign'].upper()}." sobre as necessidades emocionais e de segurança de {pet_name})
###PET_ASCENDENTE###
(1 parágrafo curto, ~70 palavras, começando com "ASCENDENTE EM {p_asc['sign'].upper()}." sobre a primeira impressão que {pet_name} causa e sua presença no ambiente)
"""))

    steps.append(("pet_posicoes", f"""{PET_SYNASTRY_STYLE_GUIDE}

Posições planetárias de {pet_name}:
{pet_planets_txt}

###PET_POSICOES_TEXTO###
(1 parágrafo curto, ~100 palavras, comentando as demais posições de {pet_name} — rotina, sociabilidade com outros animais e pessoas, nível de energia — em registro animal, com 1 exemplo concreto de comportamento)
"""))

    steps.append(("vinculo", f"""{PET_SYNASTRY_STYLE_GUIDE}

Tutor {owner_name}: Sol {o_sol['sign']}, Lua {o_lua['sign']}, Ascendente {o_asc['sign']}.
Pet {pet_name}: Sol {p_sol['sign']}, Lua {p_lua['sign']}, Ascendente {p_asc['sign']}.

###VINCULO_TEXTO###
(2 parágrafos curtos, ~90 palavras cada, sobre o que aproxima {owner_name} e {pet_name} instintivamente, cruzando as tríades dos dois)
###VINCULO_CALLOUT###
(1 frase de destaque, ~25 palavras, sobre a essência do vínculo entre os dois)
"""))

    steps.append(("aspectos_cruzados", f"""{PET_SYNASTRY_STYLE_GUIDE}

Aspectos cruzados harmônicos entre {owner_name} e {pet_name}:
{cross_harm_txt}

Aspectos cruzados tensos:
{cross_tense_txt}

###ASPECTOS_CRUZADOS_HARM_TEXTO###
(1 parágrafo curto, ~100 palavras, comentando os aspectos harmônicos acima entre os dois mapas)
###ASPECTOS_CRUZADOS_TENSOS_TEXTO###
(1 parágrafo curto, ~100 palavras, comentando os aspectos tensos acima, sem tom de alarme, com sugestões práticas)
"""))

    steps.append(("casas_cruzadas", f"""{PET_SYNASTRY_STYLE_GUIDE}

Planetas de {owner_name} caindo nas casas do mapa de {pet_name}:
{overlay_owner_txt}

Planetas de {pet_name} caindo nas casas do mapa de {owner_name}:
{overlay_pet_txt}

###CASAS_CRUZADAS_TEXTO###
(1 parágrafo curto, ~100 palavras, interpretando o que essas sobreposições de casas revelam sobre as áreas de vida que os dois ativam um no outro)
"""))

    steps.append(("rotina", f"""{PET_SYNASTRY_STYLE_GUIDE}

Tutor {owner_name} (Lua {o_lua['sign']}) e pet {pet_name} (Lua {p_lua['sign']}).

###ROTINA_TEXTO###
(2 parágrafos curtos, ~85 palavras cada, sobre como o vínculo aparece no cotidiano — cuidado, rotina, formas de comunicação entre os dois, com exemplos concretos do dia a dia)
"""))

    steps.append(("desafios", f"""{PET_SYNASTRY_STYLE_GUIDE}

Aspectos tensos entre {owner_name} e {pet_name}:
{cross_tense_txt}

Liste 3 desafios/áreas de atenção da dupla, formato exato:
###D1_TITLE### (título curto em caixa alta, 3-6 palavras)
###D1_TEXT### (1 parágrafo curto, ~90 palavras, com sugestão prática)
###D2_TITLE###
###D2_TEXT###
###D3_TITLE###
###D3_TEXT###
"""))

    steps.append(("previsao", f"""{PET_SYNASTRY_STYLE_GUIDE}

Dupla {owner_name} e {pet_name}.

###PREVISAO_TEXTO###
(2 parágrafos curtos, ~85 palavras cada, sobre os ciclos simbólicos dos próximos 6 a 12 meses para o vínculo entre os dois — oportunidades, fases de aprofundamento, cuidados)
"""))

    steps.append(("sintese", f"""{PET_SYNASTRY_STYLE_GUIDE}

Dupla {owner_name} e {pet_name}.{disclaimer_hint}

###SINTESE_P1###
(1 parágrafo curto, ~90 palavras, sintetizando a essência do vínculo entre os dois)
###SINTESE_P2###
(1 parágrafo curto, ~80 palavras, mensagem de encerramento inspiradora e gentil{"; retome brevemente que o horário do pet foi estimado" if disclaimer_note else ""})
###SINTESE_QUOTE###
(1 frase poética e memorável, entre aspas, resumindo a essência da dupla em até 20 palavras)
"""))

    results = _run_steps(steps, model, progress_cb)

    def gv(step, marker, default=""):
        return results.get(step, {}).get(marker, default)

    sections = {}
    sections["intro_subtitle"] = f"Uma leitura simbólica do vínculo entre {owner_name} e {pet_name}."
    sections["intro_observacao"] = gv("intro", "OBSERVACAO")
    sections["intro_apresentacao"] = gv("intro", "APRESENTACAO")
    sections["intro_como_aproveitar"] = gv("intro", "COMO_APROVEITAR")
    sections["time_estimated_disclaimer"] = disclaimer_note

    sections["owner_sol"] = _ensure_leadin(gv("owner_essencia", "OWNER_SOL"), f"SOL EM {o_sol['sign'].upper()} NA {o_sol['house'].upper()} CASA.")
    sections["owner_lua"] = _ensure_leadin(gv("owner_essencia", "OWNER_LUA"), f"LUA EM {o_lua['sign'].upper()} NA {o_lua['house'].upper()} CASA.")
    sections["owner_ascendente"] = _ensure_leadin(gv("owner_essencia", "OWNER_ASCENDENTE"), f"ASCENDENTE EM {o_asc['sign'].upper()}.")
    sections["owner_sintese"] = gv("owner_essencia", "OWNER_SINTESE")

    sections["pet_sol"] = _ensure_leadin(gv("pet_essencia", "PET_SOL"), f"SOL EM {p_sol['sign'].upper()}.")
    sections["pet_lua"] = _ensure_leadin(gv("pet_essencia", "PET_LUA"), f"LUA EM {p_lua['sign'].upper()}.")
    sections["pet_ascendente"] = _ensure_leadin(gv("pet_essencia", "PET_ASCENDENTE"), f"ASCENDENTE EM {p_asc['sign'].upper()}.")

    sections["pet_posicoes_texto"] = gv("pet_posicoes", "PET_POSICOES_TEXTO")

    sections["vinculo_texto"] = gv("vinculo", "VINCULO_TEXTO")
    sections["vinculo_callout"] = gv("vinculo", "VINCULO_CALLOUT")

    sections["aspectos_cruzados_harmonicos"] = [{"aspecto": a["aspecto"], "orbe": a["orbe"]}
                                                 for a in cross_aspects if a["harmonico"]][:7]
    sections["aspectos_cruzados_harmonicos_texto"] = gv("aspectos_cruzados", "ASPECTOS_CRUZADOS_HARM_TEXTO")
    sections["aspectos_cruzados_tensos"] = [{"aspecto": a["aspecto"], "orbe": a["orbe"]}
                                             for a in cross_aspects if not a["harmonico"]][:6]
    sections["aspectos_cruzados_tensos_texto"] = gv("aspectos_cruzados", "ASPECTOS_CRUZADOS_TENSOS_TEXTO")

    sections["casas_cruzadas_texto"] = gv("casas_cruzadas", "CASAS_CRUZADAS_TEXTO")

    sections["rotina_texto"] = gv("rotina", "ROTINA_TEXTO")

    desafios = []
    for i in range(1, 4):
        desafios.append({"title": gv("desafios", f"D{i}_TITLE"), "text": gv("desafios", f"D{i}_TEXT")})
    sections["desafios"] = desafios

    sections["previsao_texto"] = gv("previsao", "PREVISAO_TEXTO")

    _sint_paras = [gv("sintese", "SINTESE_P1"), gv("sintese", "SINTESE_P2")]
    if not any(p.strip() for p in _sint_paras):
        _sint_paras[0] = (f"O vínculo entre {owner_name} e {pet_name} reúne temperamentos distintos que se "
            "equilibram no cotidiano — uma combinação própria de cuidado, companhia e reconhecimento mútuo.")
    sections["sintese_final"] = _sint_paras
    sections["sintese_quote"] = gv("sintese", "SINTESE_QUOTE",
        f"“Entre {owner_name} e {pet_name}, o vínculo fala uma língua que dispensa palavras.”")

    sections["nota_final"] = [
        f"Este relatório cruza a astrologia simbólica com a leitura afetiva do vínculo entre {owner_name} e "
        f"{pet_name}. Ele não substitui orientação veterinária ou comportamental profissional."
    ]

    return sections
