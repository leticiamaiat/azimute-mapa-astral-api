# API — Gerador de Mapa Astral (Azimute)

Documentação de contrato para o time de frontend. Enquanto a API não está no endereço
definitivo, trate a URL base como uma única variável de configuração — no dia da troca,
o único ajuste necessário do lado do frontend é essa variável.

```
API_BASE_URL = http://<host-atual>:5000   # hoje (rede local / máquina de teste)
API_BASE_URL = https://api.azimute.com.br # depois (máquina definitiva)
```

Nenhuma rota, nome de campo ou formato de resposta muda entre um ambiente e outro.

## Fluxo geral

1. `POST {API_BASE_URL}/start` com os dados do formulário → recebe um `job_id`
2. `GET {API_BASE_URL}/status/{job_id}` a cada ~1.5s até `status` virar `done` ou `error`
3. Quando `done`: `GET {API_BASE_URL}/download/{job_id}` baixa o PDF final

A geração é assíncrona e demora alguns minutos (o texto de cada seção é escrito por
IA); por isso o fluxo é "inicia job → faz polling", não uma chamada síncrona única.

---

## `POST /start`

Inicia a geração do relatório.

**Content-Type:** `multipart/form-data`
**Tamanho máximo do corpo:** 25 MB (por causa da foto de capa)

| Campo   | Tipo   | Obrigatório | Descrição                                                  |
|---------|--------|-------------|--------------------------------------------------------------|
| `report_type` | string | não (default `"individual"`) | `"individual"` (Mapa Astral Completo) ou `"sinastria_pet"` (Sinastria com Pet) |
| `name`  | string | sim         | Nome completo do tutor/pessoa                                |
| `date`  | string | sim         | Data de nascimento, formato `YYYY-MM-DD`                     |
| `time`  | string | sim         | Horário de nascimento, formato `HH:MM` (24h)                 |
| `place` | string | sim         | Local de nascimento, ex: `"Cidade, Estado, País"` — usado para geocodificar lat/lon e fuso histórico automaticamente |
| `email` | string | não         | Não usado no cálculo; só repassado/armazenado                |
| `phone` | string | não         | Não usado no cálculo; só repassado/armazenado                |
| `cover` | file (imagem) | não  | Foto do tutor para a capa personalizada. Se ausente, a coluna correspondente na capa usa uma roda natal no lugar da foto |

Campos adicionais, **obrigatórios apenas quando `report_type` é `"sinastria_pet"`**:

| Campo   | Tipo   | Obrigatório | Descrição                                                  |
|---------|--------|-------------|--------------------------------------------------------------|
| `pet_name`  | string | sim (se `sinastria_pet`) | Nome do pet |
| `pet_breed` | string | não | Raça do pet — usada como cor/flavor no texto, não no cálculo |
| `pet_color` | string | não | Cor do pet — mesma finalidade que `pet_breed` |
| `pet_date`  | string | sim (se `sinastria_pet`) | Data de nascimento/adoção do pet, formato `YYYY-MM-DD` |
| `pet_time`  | string | **não** | Horário de nascimento do pet, formato `HH:MM` (24h). Se omitido, o cálculo usa `12:00` como estimativa e o relatório avisa que o ascendente/casas do pet são aproximados |
| `pet_place` | string | sim (se `sinastria_pet`) | Local de nascimento ou onde o pet foi encontrado/adotado |
| `pet_cover` | file (imagem) | não | Foto do pet para a capa personalizada, mesma lógica de `cover` |

**Resposta 200:**
```json
{ "job_id": "dd4078b23516" }
```

**Resposta 400** (faltou campo obrigatório):
```json
{ "error": "Preencha nome, data, hora e local." }
```
Se `report_type` for `"sinastria_pet"` e faltar `pet_name`, `pet_date` ou `pet_place`:
```json
{ "error": "Preencha nome, data e local do pet." }
```

> Se `place`/`pet_place` não forem reconhecidos pelo geocodificador, o erro
> correspondente aparece depois, no polling do `/status` (com `status: "error"`), não
> nesta chamada — a geocodificação roda em background junto com o cálculo do mapa.

---

## `GET /status/{job_id}`

Consulta o andamento de um job. Recomendado fazer polling a cada 1.5s (é o intervalo
usado hoje na página embutida da própria API).

**Resposta 200** — o campo `status` indica a fase atual:

| `status`    | Significado                                          | Outros campos presentes           |
|-------------|-------------------------------------------------------|------------------------------------|
| `queued`    | Job criado, ainda não começou                         | —                                   |
| `chart`     | Calculando posições planetárias e casas                | —                                   |
| `styling`   | Estilizando a foto enviada no padrão do modelo de referência (só ocorre se uma foto foi enviada) | — |
| `running`   | Escrevendo o texto de cada seção do relatório           | `progress: [i, n]`, `step: "..."`  |
| `rendering` | Montando o PDF final                                    | —                                   |
| `done`      | Pronto para download                                    | `filename: "Relatorio_....pdf"`, `has_image: bool` |
| `error`     | Falhou                                                  | `error: "mensagem"`                 |

`has_image` (presente em toda resposta, não só em `done`) indica se este job vai gerar
também a imagem estilizada separada — `true` só quando uma foto foi enviada em `/start`
**e** a estilização por IA teve sucesso. Se a estilização falhar (API fora do ar, etc.),
o relatório é gerado normalmente com a foto original e `has_image` fica `false`.

`has_pet_image` — mesmo conceito de `has_image`, só que para a foto do pet (`pet_cover`),
e só é relevante quando `report_type` foi `"sinastria_pet"`.

Exemplo durante `running`:
```json
{ "status": "running", "progress": [3, 12], "step": "Casas", "error": null, "has_image": true }
```

Exemplo `done`:
```json
{ "status": "done", "filename": "Relatorio_Mapa_Astral_Carla_Maria.pdf", "has_image": true }
```

Exemplo `error`:
```json
{ "status": "error", "error": "Não foi possível localizar 'Cidade Inexistente'. Tente um formato como 'Cidade, Estado, País'." }
```

**Resposta 404** (job_id inexistente ou expirado — os jobs ficam só em memória, se a API
reiniciar todos os jobs em andamento se perdem):
```json
{ "status": "error", "error": "job não encontrado" }
```

---

## `GET /download/{job_id}`

Baixa o PDF pronto. Só funciona quando `/status` retornou `"status": "done"`.

**Resposta 200:** binário `application/pdf`, com `Content-Disposition: attachment` e o
nome de arquivo já vindo em `filename` no `/status`.

**Resposta 400** (ainda não terminou):
```
Relatório ainda não está pronto.
```
(texto puro, não é JSON — trate pelo status code, não pelo corpo)

---

## `GET /download/{job_id}/imagem`

Baixa só a foto estilizada (PNG), separada do PDF. Só existe quando `/status` retornou
`"has_image": true`; se nenhuma foto foi enviada ou a estilização falhou, dá 404.

**Resposta 200:** binário `image/png`, com `Content-Disposition: attachment`.

**Resposta 400** (job ainda não terminou) ou **404** (sem imagem para este job):
```
Relatório ainda não está pronto.
Nenhuma imagem estilizada disponível para este pedido.
```
(texto puro, não é JSON — trate pelo status code, não pelo corpo)

---

## `GET /download/{job_id}/imagem_pet`

Mesma coisa que `/download/{job_id}/imagem`, mas para a foto estilizada do **pet**.
Só existe em jobs `report_type: "sinastria_pet"` e quando `/status` retornou
`"has_pet_image": true`.

**Resposta 200:** binário `image/png`, com `Content-Disposition: attachment`.

**Resposta 400** (job ainda não terminou) ou **404** (sem imagem para este job): mesmo
formato de `/download/{job_id}/imagem`.

---

## `GET /result/{job_id}`

Devolve os dados do relatório em JSON (em vez do PDF pronto), para quem quiser montar
o próprio layout/PDF no frontend. Só funciona quando `/status` retornou `"status": "done"`.

**Resposta 200** (estrutura real, produzida por `chart_engine.py` + `openai_engine.py`):
```json
{
  "name": "Carla Maria de Albuquerque Strafacci",
  "birth": {
    "full_name": "Carla Maria de Albuquerque Strafacci",
    "date": "1990-05-12", "time": "14:30", "place": "São Paulo, SP, Brasil",
    "zodiac_system": "Astrologia ocidental tropical", "house_system": "Placidus",
    "ascendant_label": "12°24' Libra", "midheaven_label": "03°10' Câncer",
    "timezone": "America/Sao_Paulo", "lat": -23.55, "lon": -46.63
  },
  "ascendant": { "sign": "Libra", "deg": 12, "min": 24 },
  "planets": [
    { "name": "Sol", "sign": "Touro", "deg": 21, "min": 42, "house": "8ª", "key": "...", "longitude": 51.7 },
    "... (Ascendente, Sol, Lua, Mercúrio ... Nodo Norte, Meio do Céu)"
  ],
  "houses": [ { "n": 1, "sign": "Libra", "deg": 12, "min": 24 }, "... (1 a 12)" ],
  "aspects": [ { "aspecto": "...", "orbe": "...", "...": "harmônico/tenso conforme compute_aspects" } ],
  "sections": {
    "intro_observacao": "texto gerado por IA...",
    "intro_p1": "...", "intro_p2": "...", "intro_como_aproveitar": "...",
    "triade_items": [ ["Sol em Touro — ...", "..."], ["Lua em ...", "..."], ["Ascendente em ...", "..."] ], "triade_sintese": "...",
    "casas_texto": "...", "eixo1_text": "...", "eixo2_text": "...",
    "personalidade_bullets": ["...", "...", "..."],
    "aspectos_harmonicos_items": [ ["Vênus sextil Nodo Norte — orbe 0.4°", "..."] ], "aspectos_tensos_items": "... (mesmo formato)",
    "previsao_overview": "...", "trimestres": "...", "sintese_final": "...", "nota_final": "...",
    "...": "chaves completas em openai_engine.py, função generate_full_sections"
  }
}
```
Esse é o mesmo conteúdo hoje usado para montar o PDF — cru, sem layout — então o
nome exato de cada chave de `sections` pode mudar se o texto do relatório for
reestruturado; a lista completa e atual sempre está em `generate_full_sections`
(`openai_engine.py`).

**Formato diferente para `report_type: "sinastria_pet"`** — em vez do formato plano
acima, a resposta vem aninhada em `owner`/`pet`:
```json
{
  "owner": { "name": "...", "birth": {...}, "ascendant": {...}, "planets": [...], "houses": [...], "cusps": {...} },
  "pet": { "name": "...", "breed": "...", "color": "...", "time_estimated": false, "birth": {...}, "ascendant": {...}, "planets": [...], "houses": [...], "cusps": {...} },
  "cross_aspects": [ { "aspecto": "Sol (A) trígono Lua (B)", "orbe": "2.1°", "harmonico": true } ],
  "house_overlay_owner": [ { "planet": "Vênus", "house_in_b": 7 } ],
  "house_overlay_pet": [ { "planet": "Lua", "house_in_b": 4 } ],
  "sections": { "owner_sol": "...", "pet_sol": "...", "vinculo_texto": "...", "...": "chaves completas em generate_pet_synastry_sections (openai_engine.py)" }
}
```
`cross_aspects` são os aspectos ENTRE os dois mapas (não os aspectos internos de cada
mapa). `house_overlay_owner` mostra em qual casa do pet cada planeta do tutor cai;
`house_overlay_pet` é o inverso.

**Resposta 400** (ainda não terminou ou job não gerou dados):
```json
{ "error": "Relatório ainda não está pronto." }
```

**Resposta 404** (job_id inexistente): mesmo formato do `/download`.

---

## `GET /health`

Healthcheck simples, sem parâmetros.

**Resposta 200:**
```json
{ "status": "ok" }
```

---

## CORS

A API já está preparada para ser chamada de uma origem diferente (frontend em outro
domínio/porta). Por padrão libera qualquer origem (`*`) para não travar a integração
antes do domínio final estar definido; quando o frontend tiver domínio fixo, dá pra
restringir via variável de ambiente do lado da API (`ALLOWED_ORIGINS`) sem mudar nada
no frontend.

## O que muda quando a API for para a máquina definitiva

- **Muda:** o valor de `API_BASE_URL`.
- **Não muda:** rotas, nomes de campos, formatos de request/response, fluxo de polling.

Se o domínio final for fixo (recomendado — ver `README.md`), nem isso precisa mudar de
novo no futuro, mesmo que a máquina por trás do domínio seja trocada outra vez.

## Autenticação

Não há autenticação hoje — qualquer um que souber o `job_id` (string aleatória de 12
caracteres) consegue consultar o status e baixar o PDF daquele job. Aceitável para o
estágio atual, mas vale revisar antes de ter tráfego real/pago passando por aqui.
