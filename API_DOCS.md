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
| `name`  | string | sim         | Nome completo                                                |
| `date`  | string | sim         | Data de nascimento, formato `YYYY-MM-DD`                     |
| `time`  | string | sim         | Horário de nascimento, formato `HH:MM` (24h)                 |
| `place` | string | sim         | Local de nascimento, ex: `"Cidade, Estado, País"` — usado para geocodificar lat/lon e fuso histórico automaticamente |
| `email` | string | não         | Não usado no cálculo; só repassado/armazenado                |
| `phone` | string | não         | Não usado no cálculo; só repassado/armazenado                |
| `cover` | file (imagem) | não  | Foto para a capa personalizada. Se ausente, gera uma capa padrão navy/dourada |

**Resposta 200:**
```json
{ "job_id": "dd4078b23516" }
```

**Resposta 400** (faltou campo obrigatório):
```json
{ "error": "Preencha nome, data, hora e local." }
```

> Se `place` não for reconhecido pelo geocodificador, o erro correspondente aparece
> depois, no polling do `/status` (com `status: "error"`), não nesta chamada — a
> geocodificação roda em background junto com o cálculo do mapa.

---

## `GET /status/{job_id}`

Consulta o andamento de um job. Recomendado fazer polling a cada 1.5s (é o intervalo
usado hoje na página embutida da própria API).

**Resposta 200** — o campo `status` indica a fase atual:

| `status`    | Significado                                          | Outros campos presentes           |
|-------------|-------------------------------------------------------|------------------------------------|
| `queued`    | Job criado, ainda não começou                         | —                                   |
| `chart`     | Calculando posições planetárias e casas                | —                                   |
| `running`   | Escrevendo o texto de cada seção do relatório           | `progress: [i, n]`, `step: "..."`  |
| `rendering` | Montando o PDF final                                    | —                                   |
| `done`      | Pronto para download                                    | `filename: "Relatorio_....pdf"`    |
| `error`     | Falhou                                                  | `error: "mensagem"`                 |

Exemplo durante `running`:
```json
{ "status": "running", "progress": [3, 12], "step": "Casas", "error": null }
```

Exemplo `done`:
```json
{ "status": "done", "filename": "Relatorio_Mapa_Astral_Carla_Maria.pdf" }
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
