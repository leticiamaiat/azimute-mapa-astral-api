# -*- coding: utf-8 -*-
"""
Estilizacao da foto enviada pelo cliente para o padrao visual do modelo de
referencia (retrato editorial, fundo azul-marinho profundo com halo/aura
dourada e luz quente, no espirito da capa da Barbara).

Fluxo:
1. Recorta/redimensiona a foto original em "cover" para o formato retrato
   usado na capa (sem distorcer).
2. Segmenta o assunto (pessoa) da foto com rembg, gerando uma mascara alfa.
3. Usa essa mascara para pedir a API de imagem da OpenAI (gpt-image-1) para
   redesenhar SOMENTE o fundo ao redor da pessoa no estilo de referencia,
   mantendo o rosto e o corpo da pessoa intocados (protegidos pela mascara).
4. Salva o resultado como PNG e devolve o caminho.

Requer OPENAI_API_KEY (mesmo .env usado por openai_engine.py). Qualquer falha
(rede, API, segmentacao) e reportada via excecao para quem chamou decidir o
fallback (ex.: usar a foto original sem estilizacao).
"""
import io
import os

from dotenv import load_dotenv
from PIL import Image, ImageOps
from openai import OpenAI

load_dotenv()

MODEL = "gpt-image-1"

# Tamanho final da imagem estilizada (retrato, mesma proporcao da caixa de
# foto usada na capa do relatorio: ~0.56 x 0.545 da pagina A4).
TARGET_W, TARGET_H = 1024, 1536

STYLE_PROMPT_PESSOA = (
    "Redesenhe apenas o fundo desta foto de retrato, mantendo a pessoa exatamente "
    "como esta (rosto, pele, cabelo, roupa e pose intocados). Novo fundo: estudio "
    "fotografico editorial em azul-marinho muito escuro e profundo, quase preto, "
    "com um halo suave de luz dourada quente atras da pessoa, como um brilho de "
    "aura, e um leve efeito de particulas/estrelas douradas desfocadas ao fundo, "
    "no estilo de uma capa de revista de astrologia sofisticada. Ilumine a cena "
    "com uma luz de contorno (rim light) dourada quente vinda de tras, sutil, "
    "sem alterar os tracos faciais da pessoa. Sem texto, sem logotipos, sem "
    "elementos graficos de roda zodiacal desenhados sobre a foto. Fotografia "
    "realista, alta qualidade, tons quentes e dramaticos."
)

STYLE_PROMPT_PET = (
    "Redesenhe apenas o fundo desta foto, mantendo o pet exatamente como esta "
    "(pelagem, cores, focinho/rosto, orelhas e pose intocados). Novo fundo: estudio "
    "fotografico editorial em azul-marinho muito escuro e profundo, quase preto, "
    "com um halo suave de luz dourada quente atras do animal, como um brilho de "
    "aura, e um leve efeito de particulas/estrelas douradas desfocadas ao fundo, "
    "no estilo de uma capa de revista de astrologia sofisticada. Ilumine a cena "
    "com uma luz de contorno (rim light) dourada quente vinda de tras, sutil, "
    "sem alterar as caracteristicas fisicas do animal. Sem texto, sem logotipos, sem "
    "elementos graficos de roda zodiacal desenhados sobre a foto. Fotografia "
    "realista, alta qualidade, tons quentes e dramaticos."
)

_STYLE_PROMPTS = {"pessoa": STYLE_PROMPT_PESSOA, "pet": STYLE_PROMPT_PET}


def _get_client():
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError(
            "OPENAI_API_KEY nao encontrada. Configure-a no arquivo .env desta pasta "
            "(OPENAI_API_KEY=sk-...) ou como variavel de ambiente."
        )
    return OpenAI(api_key=api_key)


def _cover_resize(img, target_w, target_h):
    """Recorta a imagem no formato 'cover' (preenche o retangulo alvo sem
    distorcer) e redimensiona para target_w x target_h."""
    src_w, src_h = img.size
    target_ratio, src_ratio = target_w / target_h, src_w / src_h
    if src_ratio > target_ratio:
        new_w = int(src_h * target_ratio)
        x0 = (src_w - new_w) // 2
        img = img.crop((x0, 0, x0 + new_w, src_h))
    else:
        new_h = int(src_w / target_ratio)
        y0 = max(0, (src_h - new_h) // 3)
        img = img.crop((0, y0, src_w, min(src_h, y0 + new_h)))
    return img.resize((target_w, target_h), Image.LANCZOS)


def _segment_subject_mask(img_rgb):
    """Retorna uma imagem RGBA do mesmo tamanho de img_rgb, onde o canal alfa
    e 255 sobre a pessoa (regiao a proteger) e 0 no fundo (regiao editavel
    pela IA) -- exatamente o formato de mascara que a API de edicao da OpenAI
    espera (areas transparentes = editaveis)."""
    from rembg import remove, new_session

    session = new_session("u2net")
    buf = io.BytesIO()
    img_rgb.save(buf, format="PNG")
    cutout_bytes = remove(buf.getvalue(), session=session)
    cutout = Image.open(io.BytesIO(cutout_bytes)).convert("RGBA")

    mask = Image.new("RGBA", cutout.size, (0, 0, 0, 0))
    alpha = cutout.split()[3]
    mask.putalpha(alpha)
    return mask


def stylize_cover_photo(input_path, out_path, subject="pessoa"):
    """Estiliza a foto em input_path no padrao do modelo de referencia e
    salva o PNG resultante em out_path. Devolve out_path. Lanca excecao se
    qualquer etapa falhar (segmentacao ou chamada da API).
    `subject`: "pessoa" (default, retrato humano) ou "pet" (ajusta o prompt
    de estilizacao para nao presumir tracos faciais humanos)."""
    style_prompt = _STYLE_PROMPTS.get(subject, STYLE_PROMPT_PESSOA)
    img = Image.open(input_path)
    img = ImageOps.exif_transpose(img).convert("RGB")
    img = _cover_resize(img, TARGET_W, TARGET_H)

    base_buf = io.BytesIO()
    img.save(base_buf, format="PNG")
    base_buf.name = "foto.png"
    base_buf.seek(0)

    mask = _segment_subject_mask(img)
    mask_buf = io.BytesIO()
    mask.save(mask_buf, format="PNG")
    mask_buf.name = "mascara.png"
    mask_buf.seek(0)

    client = _get_client()
    resp = client.images.edit(
        model=MODEL,
        image=base_buf,
        mask=mask_buf,
        prompt=style_prompt,
        size=f"{TARGET_W}x{TARGET_H}",
        quality="high",
        input_fidelity="high",
    )

    import base64
    b64 = resp.data[0].b64_json
    with open(out_path, "wb") as f:
        f.write(base64.b64decode(b64))
    return out_path
