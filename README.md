# Azimute — Gerador de Mapa Astral

API/aplicação Flask que recebe nome, data/hora/local de nascimento (e opcionalmente uma foto),
calcula o mapa astral (Skyfield + sistema de casas Placidus) e monta o relatório em PDF
pronto para download. O formulário de entrada é servido pela própria aplicação (rota `/`) —
não há frontend separado.

## Requisitos

- Windows com **Microsoft Word instalado** (a conversão para PDF usa automação COM do Word via `docx2pdf`/`pythoncom`)
- Python 3.11+

## Setup local

```bash
pip install -r requirements.txt
copy .env.example .env   # depois edite e coloque sua OPENAI_API_KEY
python app.py
```

Abre em `http://localhost:5000` (ou o host/porta definidos em `HOST`/`PORT`, ver abaixo).

## Variáveis de ambiente

| Variável         | Obrigatória | Default   | Descrição                        |
|------------------|-------------|-----------|-----------------------------------|
| `OPENAI_API_KEY` | sim         | —         | Chave da API da OpenAI (GPT-4o)   |
| `HOST`           | não         | `0.0.0.0` | Endereço em que o servidor escuta |
| `PORT`           | não         | `5000`    | Porta em que o servidor escuta    |

## Rotas

- `GET /` — página com o formulário
- `POST /start` — inicia a geração (multipart/form-data: `name`, `email`, `phone`, `date`, `time`, `place`, `cover`), retorna `{"job_id": "..."}`
- `GET /status/<job_id>` — status do job (`queued` / `chart` / `running` / `rendering` / `done` / `error`)
- `GET /download/<job_id>` — baixa o PDF pronto
- `GET /health` — healthcheck

## Rodando 24/7 na máquina definitiva

O código não precisa mudar para ir de "teste local" para "sempre ligado" — só o *modo* de execução:

1. **Servidor de produção em vez do modo dev do Flask**: `python app.py` usa o servidor
   embutido do Flask, que não é recomendado para carga contínua. Prefira
   [`waitress`](https://pypi.org/project/waitress/) (Windows-friendly, sem compilação):

   ```bash
   pip install waitress
   waitress-serve --host=0.0.0.0 --port=5000 app:app
   ```

2. **Manter de pé sozinho**: registre esse comando como serviço do Windows (ex. via
   [NSSM](https://nssm.cc/)) ou como Tarefa Agendada "ao iniciar o sistema" com reinício
   automático em caso de falha.

3. **Endereço público estável**: para o link usado por quem acessa o formulário nunca
   precisar mudar (mesmo que a máquina ou o IP mudem no futuro), aponte um domínio fixo
   (ex. `api.azimute.com.br`) para a máquina — via encaminhamento de porta + DNS, ou um
   Cloudflare Tunnel **nomeado** (diferente do quick tunnel usado em testes, que gera uma
   URL aleatória a cada execução). Assim, trocar de máquina no futuro vira só repontar o
   DNS/túnel, sem tocar em código nem em nenhum link já divulgado.

## Observações

- `de421.bsp` (efemérides JPL usadas pelo Skyfield) está versionado no repositório para não
  depender de download em tempo de execução.
- `output/` e `uploads/` guardam arquivos gerados em runtime (contêm dados pessoais de
  clientes) e não são versionados — ver `.gitignore`.
