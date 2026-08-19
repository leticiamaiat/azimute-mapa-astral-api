# -*- coding: utf-8 -*-
"""
Gerador de Mapas Astrais - fluxo completo e automatico.
O cliente preenche nome, data/hora/local de nascimento e envia uma foto;
o app calcula o mapa (Skyfield + Placidus), gera o texto de cada secao com
a API da OpenAI (GPT), e monta o .docx pronto no layout da Barbara.

Rodar: python app.py   ->  http://localhost:5000
"""
import io
import os
import re
import tempfile
import threading
import time
import traceback
import uuid

import pythoncom
from flask import Flask, request, render_template_string, send_file, jsonify
from flask_cors import CORS
from docx2pdf import convert as _docx_to_pdf

from chart_engine import compute_chart
from openai_engine import generate_full_sections, DEFAULT_MODEL
from report_engine import build_docx_bytes

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024

# ALLOWED_ORIGINS: lista separada por virgula (ex: "https://azimute.com.br,https://staging.azimute.com.br").
# Default "*" para nao travar a integracao enquanto o dominio final do frontend nao esta definido.
_allowed_origins = os.environ.get("ALLOWED_ORIGINS", "*")
CORS(app, origins=[o.strip() for o in _allowed_origins.split(",")])

JOBS = {}  # job_id -> {"status": "...", "step": "", "progress": (i,n), "error": None, "filename": None}

# O Word COM (usado pelo docx2pdf) nao aguenta duas conversoes simultaneas -
# uma segunda chamada concorrente derruba a primeira com "remote procedure
# call failed". Serializa todas as conversoes por este lock.
_PDF_LOCK = threading.Lock()

def _docx_to_pdf_safe(docx_path, pdf_path, retries=3):
    last_err = None
    with _PDF_LOCK:
        for attempt in range(retries):
            try:
                _docx_to_pdf(docx_path, pdf_path)
                return
            except Exception as e:
                last_err = e
                time.sleep(2 + attempt * 2)
    raise last_err

PAGE = """
<!doctype html>
<html lang="pt-br">
<head>
<meta charset="utf-8">
<title>Azimute — Mapa Astral Personalizado</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
  :root{
    --navy:#0a1730; --navy-2:#0f2140; --gold:#c9a24b; --gold-soft:#e9c98a;
    --cream:#f4ead0; --ink:#e9ecf2; --sub:#9fb0c9; --line:#2a3c5c;
  }
  *{box-sizing:border-box;}
  html{ scroll-behavior:smooth; }
  body{
    font-family: Georgia, 'Times New Roman', serif;
    background: radial-gradient(ellipse at top, var(--navy-2), var(--navy) 70%);
    margin:0; color:var(--ink); min-height:100vh;
  }
  .sans{ font-family: 'Segoe UI', system-ui, Arial, sans-serif; }
  nav{
    position:sticky; top:0; z-index:10; display:flex; align-items:center; justify-content:space-between;
    padding:16px 6vw; background:rgba(10,23,48,.88); backdrop-filter:blur(6px); border-bottom:1px solid var(--line);
  }
  nav .brand{ font-size:19px; letter-spacing:3px; color:var(--gold-soft); font-weight:bold; }
  nav a{ color:var(--sub); text-decoration:none; font-size:13px; margin-left:22px; font-family:'Segoe UI',system-ui,sans-serif; }
  nav a:hover{ color:var(--gold-soft); }

  .hero{ text-align:center; padding:64px 6vw 34px; }
  .hero .eyebrow{ font-family:'Segoe UI',system-ui,sans-serif; letter-spacing:4px; font-size:12px; color:var(--gold); text-transform:uppercase; }
  .hero h1{ font-size:clamp(28px,4.4vw,44px); margin:14px 0 10px; color:#fff; font-weight:normal; letter-spacing:.5px; }
  .hero p{ font-family:'Segoe UI',system-ui,sans-serif; color:var(--sub); max-width:560px; margin:0 auto; font-size:15px; line-height:1.6; }
  .stars{ font-size:13px; letter-spacing:14px; color:var(--gold); margin:22px 0 0; }

  main{ max-width:720px; margin:0 auto; padding:10px 6vw 80px; }
  section{ margin-top:56px; }
  h2{ font-size:22px; color:#fff; font-weight:normal; text-align:center; margin:0 0 6px; }
  .section-sub{ font-family:'Segoe UI',system-ui,sans-serif; text-align:center; color:var(--sub); font-size:13.5px; margin-bottom:28px; }

  .plans{ display:grid; grid-template-columns:repeat(auto-fit,minmax(180px,1fr)); gap:16px; }
  .plan{ background:linear-gradient(180deg, rgba(201,162,75,.07), rgba(201,162,75,.02)); border:1px solid var(--line);
    border-radius:10px; padding:20px 18px; text-align:center; font-family:'Segoe UI',system-ui,sans-serif; }
  .plan.active{ border-color:var(--gold); box-shadow:0 0 0 1px var(--gold) inset; }
  .plan .name{ font-family:Georgia,serif; font-size:16px; color:#fff; margin-bottom:6px; }
  .plan .price{ font-size:20px; color:var(--gold-soft); font-weight:bold; margin-bottom:8px; }
  .plan .desc{ font-size:12px; color:var(--sub); line-height:1.5; }
  .plan .tag{ display:inline-block; margin-top:10px; font-size:10.5px; letter-spacing:1px; text-transform:uppercase;
    color:var(--gold); border:1px solid var(--gold); border-radius:20px; padding:3px 10px; }
  .plan.soon{ opacity:.45; }
  .plan.soon .tag{ color:var(--sub); border-color:var(--line); }

  .card{ background:rgba(255,255,255,.03); border:1px solid var(--line); border-radius:12px; padding:30px 28px;
    font-family:'Segoe UI',system-ui,sans-serif; }
  label{ display:block; font-size:12px; font-weight:600; color:var(--gold-soft); letter-spacing:.4px;
    text-transform:uppercase; margin:16px 0 6px; }
  input[type=text], input[type=email], input[type=tel], input[type=date], input[type=time], input[type=file]{
    width:100%; padding:11px 12px; border:1px solid var(--line); border-radius:7px; font-size:14px;
    background:rgba(255,255,255,.04); color:var(--ink); font-family:'Segoe UI',system-ui,sans-serif; }
  input::placeholder{ color:#5c6d8a; }
  input[type=file]{ padding:9px 10px; }
  .row{ display:flex; gap:14px; flex-wrap:wrap; }
  .row > div{ flex:1; min-width:150px; }
  .hint{ font-size:12px; color:var(--sub); margin-top:5px; line-height:1.5; }
  .consent{ display:flex; gap:10px; align-items:flex-start; margin-top:22px; font-size:12.5px; color:var(--sub); }
  .consent input{ margin-top:3px; width:auto; }
  button{ background:linear-gradient(180deg,var(--gold-soft),var(--gold)); color:#1a1206; border:none; padding:14px 24px;
    border-radius:7px; font-weight:bold; cursor:pointer; font-size:15px; margin-top:24px; width:100%;
    letter-spacing:.3px; font-family:'Segoe UI',system-ui,sans-serif; }
  button:disabled{ opacity:.55; cursor:default; }
  .flash{ background:rgba(200,60,60,.12); border:1px solid #a33; color:#ffb4b4; padding:10px 14px; border-radius:7px;
    margin-top:16px; font-size:13px; }
  #status{ margin-top:20px; font-size:13.5px; color:var(--sub); display:none; }
  #bar{ height:8px; background:rgba(255,255,255,.08); border-radius:4px; overflow:hidden; margin-top:10px; }
  #bar-fill{ height:100%; width:0%; background:linear-gradient(90deg,var(--gold),var(--gold-soft)); transition:width .4s; }

  .trust{ display:flex; gap:22px; flex-wrap:wrap; justify-content:center; font-family:'Segoe UI',system-ui,sans-serif;
    font-size:12.5px; color:var(--sub); }
  .trust span{ color:var(--gold); margin-right:6px; }
  footer{ text-align:center; padding:36px 20px 50px; color:#5c6d8a; font-size:12px; font-family:'Segoe UI',system-ui,sans-serif; }
</style>
</head>
<body>
<nav>
  <div class="brand">AZIMUTE</div>
  <div>
    <a href="#planos">Planos</a>
    <a href="#pedido">Pedido</a>
    <a href="#lgpd">LGPD</a>
  </div>
</nav>

<div class="hero">
  <div class="eyebrow sans">Astrologia autoral · relatório simbólico</div>
  <h1>Seu céu, sua direção.</h1>
  <p>Preencha os dados de nascimento e receba um mapa astral completo, calculado com precisão
  astronômica e escrito especialmente para você.</p>
  <div class="stars">✶ ☉ ☽ ♃ ✶</div>
</div>

<main>
  <section id="planos">
    <h2>Planos</h2>
    <div class="section-sub">Escolha o relatório que combina com o seu momento.</div>
    <div class="plans">
      <div class="plan active">
        <div class="name">Mapa Astral Completo</div>
        <div class="price">R$ 49,90</div>
        <div class="desc">25 páginas, roda natal técnica, tríade, casas, aspectos e previsão anual — gerado agora mesmo.</div>
        <div class="tag">Disponível</div>
      </div>
      <div class="plan soon">
        <div class="name">Sinastria Casal</div>
        <div class="price">R$ 49,90</div>
        <div class="desc">Compatibilidade entre dois mapas natais.</div>
        <div class="tag">Em breve</div>
      </div>
      <div class="plan soon">
        <div class="name">Mapa Pet Completo</div>
        <div class="price">R$ 49,90</div>
        <div class="desc">Leitura simbólica para o seu animal de estimação.</div>
        <div class="tag">Em breve</div>
      </div>
    </div>
  </section>

  <section id="pedido">
    <h2>Seu pedido</h2>
    <div class="section-sub">Os dados de nascimento definem o cálculo do seu mapa — quanto mais precisos, melhor a leitura.</div>
    <div class="card">
      <form id="f">
        <label>Nome completo</label>
        <input type="text" name="name" required placeholder="Ex: Carla Maria de Albuquerque Strafacci">

        <div class="row">
          <div>
            <label>E-mail</label>
            <input type="email" name="email" placeholder="voce@email.com">
          </div>
          <div>
            <label>WhatsApp</label>
            <input type="tel" name="phone" placeholder="(11) 90000-0000">
          </div>
        </div>

        <div class="row">
          <div>
            <label>Data de nascimento</label>
            <input type="date" name="date" required>
          </div>
          <div>
            <label>Horário de nascimento</label>
            <input type="time" name="time" required>
          </div>
        </div>

        <label>Local de nascimento</label>
        <input type="text" name="place" required placeholder="Cidade, Estado, País">
        <div class="hint">Usado para calcular latitude/longitude e fuso horário histórico automaticamente.</div>

        <label>Fotografia para a arte personalizada (opcional)</label>
        <input type="file" name="cover" accept="image/*">
        <div class="hint">Se não enviar, uma capa navy/dourada é gerada automaticamente.</div>

        <div class="consent">
          <input type="checkbox" id="consent" required>
          <label for="consent" style="text-transform:none; font-weight:normal; color:var(--sub); margin:0;">
            Autorizo o uso dos meus dados de nascimento e da fotografia enviada exclusivamente para a geração deste relatório personalizado.
          </label>
        </div>

        <button id="btn" type="submit">Gerar meu mapa astral</button>
      </form>
      <div id="status">
        <div id="status-text">Iniciando…</div>
        <div id="bar"><div id="bar-fill"></div></div>
        <div class="hint" id="status-hint">O texto de cada seção é escrito por um modelo local — isso leva alguns minutos.</div>
      </div>
      <div id="err" class="flash" style="display:none;"></div>
    </div>
  </section>

  <section id="lgpd">
    <h2>Privacidade</h2>
    <div class="trust">
      <div><span>✓</span>Dados usados apenas para gerar o seu relatório</div>
      <div><span>✓</span>Processamento local, sem serviços de IA de terceiros</div>
      <div><span>✓</span>Conforme a LGPD</div>
    </div>
  </section>
</main>

<footer>AZIMUTE — Seu céu, sua direção.</footer>

<script>
const form = document.getElementById('f');
const btn = document.getElementById('btn');
const statusBox = document.getElementById('status');
const statusText = document.getElementById('status-text');
const barFill = document.getElementById('bar-fill');
const errBox = document.getElementById('err');

form.addEventListener('submit', async (e) => {
  e.preventDefault();
  errBox.style.display = 'none';
  btn.disabled = true;
  statusBox.style.display = 'block';
  statusText.textContent = 'Calculando o mapa astral...';

  const fd = new FormData(form);
  const startResp = await fetch('/start', { method: 'POST', body: fd });
  const startData = await startResp.json();
  if (!startResp.ok) {
    errBox.textContent = startData.error || 'Erro ao iniciar.';
    errBox.style.display = 'block';
    btn.disabled = false;
    statusBox.style.display = 'none';
    return;
  }
  const jobId = startData.job_id;
  const poll = setInterval(async () => {
    const r = await fetch('/status/' + jobId);
    const d = await r.json();
    if (d.status === 'running') {
      const [i, n] = d.progress || [0, 1];
      statusText.textContent = 'Escrevendo seção ' + i + ' de ' + n + ' (' + (d.step || '') + ')...';
      barFill.style.width = Math.round(100 * i / n) + '%';
    } else if (d.status === 'chart') {
      statusText.textContent = 'Calculando posições planetárias e casas...';
    } else if (d.status === 'rendering') {
      statusText.textContent = 'Montando o PDF final...';
      barFill.style.width = '100%';
    } else if (d.status === 'done') {
      clearInterval(poll);
      statusText.textContent = 'Pronto! Baixando...';
      barFill.style.width = '100%';
      window.location = '/download/' + jobId;
      btn.disabled = false;
    } else if (d.status === 'error') {
      clearInterval(poll);
      errBox.textContent = d.error;
      errBox.style.display = 'block';
      statusBox.style.display = 'none';
      btn.disabled = false;
    }
  }, 1500);
});
</script>
</body>
</html>
"""

def slugify(name):
    return re.sub(r"[^A-Za-z0-9]+", "_", name).strip("_") or "relatorio"

def _run_job(job_id, name, date_str, time_str, place, cover_path):
    pythoncom.CoInitialize()
    try:
        JOBS[job_id]["status"] = "chart"
        chart = compute_chart(name, date_str, time_str, place)

        def progress_cb(i, n, label):
            JOBS[job_id]["status"] = "running"
            JOBS[job_id]["progress"] = (i, n)
            JOBS[job_id]["step"] = label

        sections = generate_full_sections(chart, model=DEFAULT_MODEL, progress_cb=progress_cb)

        data = {
            "name": chart["name"],
            "birth": chart["birth"],
            "ascendant": chart["ascendant"],
            "planets": chart["planets"],
            "houses": chart["houses"],
            "sections": sections,
        }
        JOBS[job_id]["status"] = "rendering"
        docx_bytes = build_docx_bytes(data, cover_image_path=cover_path)
        base_name = slugify(name)
        docx_path = os.path.join(OUTPUT_DIR, f"{job_id}_Relatorio_Mapa_Astral_{base_name}.docx")
        with open(docx_path, "wb") as f:
            f.write(docx_bytes)

        pdf_filename = f"Relatorio_Mapa_Astral_{base_name}.pdf"
        pdf_path = os.path.join(OUTPUT_DIR, f"{job_id}_{pdf_filename}")
        _docx_to_pdf_safe(docx_path, pdf_path)
        os.remove(docx_path)

        JOBS[job_id]["status"] = "done"
        JOBS[job_id]["filename"] = pdf_filename
        JOBS[job_id]["path"] = pdf_path
    except Exception as e:
        traceback.print_exc()
        JOBS[job_id]["status"] = "error"
        JOBS[job_id]["error"] = str(e)
    finally:
        pythoncom.CoUninitialize()

@app.route("/", methods=["GET"])
def index():
    return render_template_string(PAGE)

@app.route("/start", methods=["POST"])
def start():
    name = request.form.get("name", "").strip()
    date_str = request.form.get("date", "").strip()
    time_str = request.form.get("time", "").strip()
    place = request.form.get("place", "").strip()
    if not all([name, date_str, time_str, place]):
        return jsonify({"error": "Preencha nome, data, hora e local."}), 400

    cover_path = None
    file = request.files.get("cover")
    if file and file.filename:
        tmp_dir = tempfile.mkdtemp(prefix="cover_")
        cover_path = os.path.join(tmp_dir, file.filename)
        file.save(cover_path)

    job_id = uuid.uuid4().hex[:12]
    JOBS[job_id] = {"status": "queued", "progress": (0, 1), "step": "", "error": None}
    t = threading.Thread(target=_run_job, args=(job_id, name, date_str, time_str, place, cover_path), daemon=True)
    t.start()
    return jsonify({"job_id": job_id})

@app.route("/status/<job_id>")
def status(job_id):
    job = JOBS.get(job_id)
    if not job:
        return jsonify({"status": "error", "error": "job não encontrado"}), 404
    public_fields = ("status", "progress", "step", "error", "filename")
    return jsonify({k: job[k] for k in public_fields if k in job})

@app.route("/download/<job_id>")
def download(job_id):
    job = JOBS.get(job_id)
    if not job or job.get("status") != "done":
        return "Relatório ainda não está pronto.", 400
    return send_file(job["path"], as_attachment=True, download_name=job["filename"], mimetype="application/pdf")

@app.route("/health")
def health():
    return {"status": "ok"}

if __name__ == "__main__":
    host = os.environ.get("HOST", "0.0.0.0")
    port = int(os.environ.get("PORT", "5000"))
    app.run(host=host, port=port, debug=False, threaded=True)
