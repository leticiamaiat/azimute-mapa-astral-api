# -*- coding: utf-8 -*-
"""
Geracao do texto interpretativo de cada secao do relatorio usando um modelo
local via Ollama (gratuito, sem chave de API). Recebe o mapa ja calculado
(chart_engine.compute_chart) e devolve o dicionario `sections` no formato
que report_engine.build_pages espera.
"""
import json
import re
import urllib.request

OLLAMA_URL = "http://127.0.0.1:11500/api/generate"
# qwen2.5:14b escreve melhor mas e ~5x mais lento neste hardware (sem GPU forte o
# suficiente para 14B); com paragrafos de ~200 palavras isso vira ~25min/relatorio.
# Ate termos uma maquina com mais VRAM, o padrao fica no 8B (rapido o bastante para
# uso pratico). Para gerar com mais qualidade em troca de tempo, passe model="qwen2.5:14b".
DEFAULT_MODEL = "llama3:8b"

STYLE_GUIDE = """Você é um(a) astrólogo(a) profissional que escreve laudos de mapa astral em
português do Brasil, em tom sofisticado, caloroso, perspicaz e prático — nunca
determinista. Fale diretamente com a pessoa usando "você".

Exemplo do registro exato a seguir (não copie o conteúdo, apenas o estilo):
"Seu mundo emocional pode ser mais reservado do que sua comunicação sugere. Você
tende a processar sentimentos com discrição, responsabilidade e autocontrole.
Na 12ª casa, existe uma vida interior rica, que pede momentos de silêncio,
recolhimento e elaboração."

Regras obrigatórias:
- Varie as aberturas de frase. NÃO comece mais de um parágrafo com "Esta posição",
  "Essa posição" ou "Isso pode indicar/sugerir" — use no máximo uma vez por resposta.
- Prefira frases concretas e específicas (cite o signo, o planeta, a casa) a frases
  vagas que serviriam para qualquer pessoa.
- Nunca use markdown (sem **, #, listas com traço).
- Escreva apenas parágrafos corridos, no português informado, sem saudação, sem
  introdução do tipo "claro, aqui está", sem repetir o pedido.
- Responda SOMENTE com o conteúdo pedido, nada mais."""

def _call_ollama(prompt, model=DEFAULT_MODEL, temperature=0.7, num_predict=700):
    payload = {
        "model": model,
        "prompt": prompt,
        "stream": False,
        "options": {"temperature": temperature, "num_predict": num_predict},
    }
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(OLLAMA_URL, data=data, headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=180) as resp:
        out = json.loads(resp.read().decode("utf-8"))
    return out.get("response", "").strip()

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

def generate_full_sections(chart, model=DEFAULT_MODEL, progress_cb=None):
    """chart: dict retornado por chart_engine.compute_chart (inclui 'aspects').
    progress_cb(step, total, label): opcional, chamado a cada bloco concluido.
    Retorna dict pronto para report_engine (`data['sections']`)."""
    name = chart["name"]
    birth = chart["birth"]
    planets_txt = _planet_line(chart)
    harm_txt = _aspects_line(chart["aspects"], True)
    tense_txt = _aspects_line(chart["aspects"], False)

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
(1 parágrafo, ~130 palavras, explicando que este relatório usa astrologia ocidental tropical com casas Placidus, é simbólico e não determinístico)
###P1###
(1 parágrafo, ~200 palavras, dizendo que a pessoa não é só o signo solar, combinando os traços da tríade acima de forma sofisticada e específica)
###P2###
(1 parágrafo, ~130 palavras, convidando a pessoa a ler o relatório como uma paisagem simbólica, não uma verdade fechada)
###COMO_APROVEITAR###
(1 parágrafo, ~120 palavras, sugerindo como aproveitar melhor o relatório)
"""))

    steps.append(("dados", f"""{STYLE_GUIDE}

Ascendente de {name}: {asc['sign']} {asc['deg']}°{asc['min']:02d}. Meio do Céu: {birth['midheaven_label']}.

###DADOS_TEXTO###
(1 parágrafo, ~200 palavras, interpretando em grande profundidade o que o Ascendente em {asc['sign']} e o Meio do Céu em {pget('Meio do Céu')['sign']} sugerem sobre a primeira impressão, a postura social e a vocação pública da pessoa — desenvolva com vários exemplos concretos)
"""))

    steps.append(("triade", f"""{STYLE_GUIDE}

Sol em {sol['sign']} na casa {sol['house']}. Lua em {lua['sign']} na casa {lua['house']}. Ascendente em {asc['sign']}.

###SOL###
(1 parágrafo, ~200 palavras, começando com "SOL EM {sol['sign'].upper()} NA {sol['house'].upper()} CASA." interpretando essa posição em grande profundidade, com vários exemplos concretos de como isso aparece no dia a dia)
###LUA###
(1 parágrafo, ~200 palavras, começando com "LUA EM {lua['sign'].upper()} NA {lua['house'].upper()} CASA." interpretando essa posição em grande profundidade, com vários exemplos concretos)
###ASCENDENTE###
(1 parágrafo, ~180 palavras, começando com "ASCENDENTE EM {asc['sign'].upper()}." interpretando essa posição em grande profundidade, com exemplos concretos)
###SINTESE###
(1 parágrafo, ~150 palavras, sintetizando a combinação dos três elementos/temperamentos envolvidos)
"""))

    steps.append(("posicoes", f"""{STYLE_GUIDE}

Posições planetárias de {name}:
{planets_txt}

###POSICOES_TEXTO###
(1 parágrafo, ~200 palavras, comentando em detalhe quais casas concentram mais planetas e o que isso sugere sobre os temas de vida mais ativados)
"""))

    steps.append(("casas", f"""{STYLE_GUIDE}

Ascendente {asc['sign']}, Descendente oposto. Meio do Céu {pget('Meio do Céu')['sign']}, Fundo do Céu oposto.

###CASAS_TEXTO###
(1 parágrafo, ~180 palavras, introduzindo o que as casas mostram sobre onde a energia da pessoa ganha palco, citando quais casas do mapa dela têm mais peso)
###EIXO1_TEXT###
(1 parágrafo, ~170 palavras, sobre o eixo Ascendente-Descendente e o que ele revela sobre presença e relações)
###EIXO2_TEXT###
(1 parágrafo, ~170 palavras, sobre o eixo Fundo do Céu-Meio do Céu e o que ele revela sobre raízes e vocação pública)
"""))

    topic_specs = [
        ("personalidade", "PERSONALIDADE", 4, "personalidade e traços centrais, combinando Sol, Lua e Ascendente"),
        ("emocional", "EMOCIONAL", 3, "mundo emocional e relacionamentos, a partir da Lua e de Vênus"),
        ("comunicacao", "COMUNICACAO", 3, "comunicação, pensamento e forma de decidir, a partir de Mercúrio"),
        ("afeto", "AFETO", 3, "amor, afeto, desejo e valores pessoais, a partir de Vênus e Marte"),
        ("carreira", "CARREIRA", 3, "carreira, imagem pública e propósito, a partir do Meio do Céu e Marte"),
    ]
    for key, marker, n_paras, desc in topic_specs:
        paras_block = "\n".join(f"###{marker}_P{i+1}###\n(1 parágrafo, ~200 palavras, denso e específico — cite signos, planetas e casas, com exemplos concretos)" for i in range(n_paras))
        steps.append((key, f"""{STYLE_GUIDE}

Mapa de {name}: Sol {sol['sign']}, Lua {lua['sign']}, Ascendente {asc['sign']}.
Planetas relevantes:
{planets_txt}

Escreva sobre: {desc}.

{paras_block}
###{marker}_CALLOUT###
(1 frase de destaque, ~30 palavras, tipo conclusão inspiradora sobre esse tema)
"""))

    steps.append(("talentos", f"""{STYLE_GUIDE}

Planetas de {name}:
{planets_txt}

Liste 6 talentos naturais, cada um em uma linha no formato exato:
###T1_LEAD### (3-6 palavras, tipo título do talento)
###T1_TEXT### (1 frase, ~20 palavras, explicando)
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
(1 frase de destaque, ~30 palavras, sobre o talento central da pessoa)
"""))

    steps.append(("crescimento", f"""{STYLE_GUIDE}

Planetas de {name}:
{planets_txt}

Liste 4 oportunidades de crescimento pessoal, formato exato:
###G1_TITLE### (título curto em caixa alta, 3-6 palavras)
###G1_TEXT### (1 parágrafo, ~150 palavras, desenvolvido e com exemplos concretos)
###G2_TITLE###
###G2_TEXT###
###G3_TITLE###
###G3_TEXT###
###G4_TITLE###
###G4_TEXT###
"""))

    steps.append(("aspectos", f"""{STYLE_GUIDE}

Aspectos harmônicos de {name}:
{harm_txt}

Aspectos tensos:
{tense_txt}

###ASPECTOS_HARM_TEXTO###
(1 parágrafo, ~200 palavras, comentando os aspectos harmônicos acima em grande detalhe)
###ASPECTOS_TENSOS_TEXTO###
(1 parágrafo, ~200 palavras, comentando os aspectos tensos acima em grande detalhe, sem tom de alarme)
"""))

    steps.append(("previsao", f"""{STYLE_GUIDE}

Mapa de {name}: Sol {sol['sign']}, Lua {lua['sign']}, Ascendente {asc['sign']}.

Escreva a visão geral da previsão simbólica para os próximos 12 meses, em 4 blocos:

###PREV_P1###
(1 parágrafo, ~150 palavras, começando com "EXPANSÃO E OPORTUNIDADES." sobre o tema de crescimento do ano)
###PREV_P2###
(1 parágrafo, ~150 palavras, começando com "ESTRUTURA E RESPONSABILIDADE." sobre consolidar bases)
###PREV_P3###
(1 parágrafo, ~150 palavras, começando com "REVISÃO E SENSIBILIDADE." sobre um momento de questionamento)
###PREV_P4###
(1 parágrafo, ~150 palavras, começando com "TRANSFORMAÇÃO PESSOAL." sobre mudança de identidade ou postura)
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
(1 parágrafo, ~180 palavras)
###{qkey.upper()}_P2###
(1 parágrafo, ~180 palavras)
###{qkey.upper()}_OPP_TEXT###
(2 a 3 frases, ~70 palavras, sobre a oportunidade do período)
###{qkey.upper()}_DEC_TEXT###
(2 a 3 frases, ~70 palavras, sobre uma decisão prática recomendada)
"""))

    steps.append(("sintese", f"""{STYLE_GUIDE}

Mapa de {name}: Sol {sol['sign']}, Lua {lua['sign']}, Ascendente {asc['sign']}.

###SINTESE_P1###
(1 parágrafo, ~200 palavras, sobre a combinação de elementos do mapa)
###SINTESE_P2###
(1 parágrafo, ~200 palavras, sobre como a pessoa concilia forças aparentemente opostas)
###SINTESE_P3###
(1 parágrafo, ~150 palavras, sobre os próximos meses)
###SINTESE_QUOTE###
(1 frase poética e memorável, entre aspas, resumindo a essência da pessoa em até 20 palavras)
"""))

    steps.append(("nota_final", f"""{STYLE_GUIDE}

###NOTA_P1###
(1 parágrafo, ~180 palavras, sobre astrologia oferecer símbolos e a vida oferecer a experiência real)
###NOTA_P2###
(1 parágrafo, ~150 palavras, mensagem de encerramento inspiradora e gentil)
"""))

    total = len(steps)
    results = {}
    for i, (key, prompt) in enumerate(steps, start=1):
        if progress_cb:
            progress_cb(i, total, key)
        expected = re.findall(r'###\s*([A-Z0-9_]+)\s*###', prompt)
        word_targets = _extract_word_targets(prompt, expected)
        parsed = {}
        best_score = (-1, -1)
        temperatures = [0.65, 0.5, 0.3]
        for attempt in range(3):
            raw = _call_ollama(prompt, model=model, temperature=temperatures[attempt], num_predict=2400)
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

    def gv(step, marker, default=""):
        return results.get(step, {}).get(marker, default)

    sections["intro_subtitle"] = "Uma leitura simbólica para orientar reflexão, não um roteiro fechado sobre o futuro."
    sections["intro_observacao"] = gv("intro", "OBSERVACAO")
    sections["intro_p1"] = gv("intro", "P1")
    sections["intro_p2"] = gv("intro", "P2")
    sections["intro_como_aproveitar"] = gv("intro", "COMO_APROVEITAR")

    sections["dados_nascimento_texto"] = gv("dados", "DADOS_TEXTO")
    sections["mapa_tecnico_texto"] = ("A roda natal organiza os doze signos no anel externo e as doze casas no anel "
        "interno, calculadas a partir do Ascendente. Os pontos coloridos indicam a posição simbólica de cada planeta.")

    sections["triade_sol"] = _ensure_leadin(gv("triade", "SOL"), f"SOL EM {sol['sign'].upper()} NA {sol['house'].upper()} CASA.")
    sections["triade_lua"] = _ensure_leadin(gv("triade", "LUA"), f"LUA EM {lua['sign'].upper()} NA {lua['house'].upper()} CASA.")
    sections["triade_ascendente"] = _ensure_leadin(gv("triade", "ASCENDENTE"), f"ASCENDENTE EM {asc['sign'].upper()}.")
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
        sections[f"{key}_callout"] = gv(key, f"{marker}_CALLOUT",
            "Observe como esses temas aparecem nas suas próprias experiências recentes.")

    talentos = []
    for i in range(1, 7):
        talentos.append({"lead": gv("talentos", f"T{i}_LEAD"), "text": gv("talentos", f"T{i}_TEXT")})
    sections["talentos"] = talentos
    sections["talentos_callout"] = gv("talentos", "TALENTOS_CALLOUT")

    crescimento = []
    for i in range(1, 5):
        crescimento.append({"title": gv("crescimento", f"G{i}_TITLE"), "text": gv("crescimento", f"G{i}_TEXT")})
    sections["crescimento"] = crescimento

    sections["aspectos_harmonicos"] = [{"aspecto": a["aspecto"], "orbe": a["orbe"]}
                                        for a in chart["aspects"] if a["harmonico"]][:7]
    sections["aspectos_harmonicos_texto"] = gv("aspectos", "ASPECTOS_HARM_TEXTO")
    sections["aspectos_tensos"] = [{"aspecto": a["aspecto"], "orbe": a["orbe"]}
                                    for a in chart["aspects"] if not a["harmonico"]][:6]
    sections["aspectos_tensos_texto"] = gv("aspectos", "ASPECTOS_TENSOS_TEXTO")

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

    sections["apendice_texto"] = ("Os dados abaixo registram a estrutura técnica que serviu de base à interpretação. "
        "Pequenas variações podem ocorrer entre programas astrológicos conforme efemérides e convenções de cálculo.")

    _nota_paras = [gv("nota_final", "NOTA_P1"), gv("nota_final", "NOTA_P2")]
    if not any(p.strip() for p in _nota_paras):
        _nota_paras[0] = ("A astrologia oferece símbolos. A sua vida oferece a história, as escolhas, os vínculos, "
            "os aprendizados e a liberdade de transformar qualquer tendência em consciência.")
    sections["nota_final"] = _nota_paras
    sections["nota_quote"] = sections["sintese_quote"]

    return sections
