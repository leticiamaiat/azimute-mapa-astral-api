# -*- coding: utf-8 -*-
"""
Motor de geracao do Mapa Astral Personalizado (.docx), no sistema visual
extraido do PDF-modelo (Barbara): Helvetica/Arial, navy #0a213b, dourado
#b79242, cabecalho/rodape centralizados, caixas de destaque com borda fina,
tabelas com cabecalho navy solido, roda natal tecnica em SVG/PNG.

Uso:
    from report_engine import build_docx_bytes
    docx_bytes = build_docx_bytes(data)   # data segue o esquema em SCHEMA.md / exemplo_carla.json
"""
import io
import math
import os
import re
import tempfile

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches

NAVY = RGBColor(0x0a, 0x21, 0x3b)
GOLD = RGBColor(0xb7, 0x92, 0x42)
GOLD_LINE = "d9c68a"
BODY_INK = RGBColor(0x30, 0x36, 0x3b)
SUB_INK = RGBColor(0x5a, 0x61, 0x68)
FOOTER_INK = RGBColor(0x77, 0x77, 0x77)
WHITE = RGBColor(0xff, 0xff, 0xff)
BLACK = RGBColor(0x11, 0x11, 0x11)
FONT = "Arial"

SIGNS = ["Áries", "Touro", "Gêmeos", "Câncer", "Leão", "Virgem", "Libra",
         "Escorpião", "Sagitário", "Capricórnio", "Aquário", "Peixes"]

PLANET_COLORS = {
    "Sol": "#e8a33d", "Lua": "#9fd3ef", "Mercúrio": "#7fd1ae", "Vênus": "#f2a6c9",
    "Marte": "#e15b5b", "Júpiter": "#c9a4f2", "Saturno": "#9aa0a8", "Urano": "#6fd1d1",
    "Netuno": "#6f8ee0", "Plutão": "#b06f3f", "Nodo Norte": "#d9c65a",
}

# ---------------------------------------------------------------- xml helpers

def set_cell_border(cell, color="b79242", sz=8):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:val'), 'single'); el.set(qn('w:sz'), str(sz))
        el.set(qn('w:space'), '0'); el.set(qn('w:color'), color)
        borders.append(el)
    tcPr.append(borders)

def shade_cell(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
    tcPr = cell._tc.get_or_add_tcPr()
    mar = OxmlElement('w:tcMar')
    for edge, val in (('top', top), ('bottom', bottom), ('left', left), ('right', right)):
        node = OxmlElement(f'w:{edge}')
        node.set(qn('w:w'), str(val)); node.set(qn('w:type'), 'dxa')
        mar.append(node)
    tcPr.append(mar)

def set_row_bottom_border(cell, color="e6e6e6", sz=4):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement('w:tcBorders')
    el = OxmlElement('w:bottom')
    el.set(qn('w:val'), 'single'); el.set(qn('w:sz'), str(sz))
    el.set(qn('w:space'), '0'); el.set(qn('w:color'), color)
    borders.append(el)
    tcPr.append(borders)

def para_bottom_border(paragraph, color="d9c68a", sz=6):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single'); bottom.set(qn('w:sz'), str(sz))
    bottom.set(qn('w:space'), '4'); bottom.set(qn('w:color'), color)
    pBdr.append(bottom)
    pPr.append(pBdr)

def add_page_field(paragraph):
    run = paragraph.add_run()
    fld1 = OxmlElement('w:fldChar'); fld1.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText'); instr.set(qn('xml:space'), 'preserve'); instr.text = 'PAGE'
    fld2 = OxmlElement('w:fldChar'); fld2.set(qn('w:fldCharType'), 'end')
    run._r.append(fld1); run._r.append(instr); run._r.append(fld2)
    run.font.name = FONT; run.font.size = Pt(6.3); run.font.color.rgb = FOOTER_INK

# --------------------------------------------------------------- text helpers

BOLD_RE = re.compile(r'\*\*(.+?)\*\*')

def add_richtext(paragraph, text, size=9.2, color=BODY_INK, base_bold=False, italic=False, font=FONT):
    pos = 0
    for m in BOLD_RE.finditer(text):
        if m.start() > pos:
            r = paragraph.add_run(text[pos:m.start()])
            r.font.name = font; r.font.size = Pt(size); r.font.color.rgb = color; r.bold = base_bold; r.italic = italic
        r = paragraph.add_run(m.group(1))
        r.font.name = font; r.font.size = Pt(size); r.font.color.rgb = color; r.bold = True; r.italic = italic
        pos = m.end()
    if pos < len(text):
        r = paragraph.add_run(text[pos:])
        r.font.name = font; r.font.size = Pt(size); r.font.color.rgb = color; r.bold = base_bold; r.italic = italic

def add_para(doc, text, size=9.2, color=BODY_INK, bold=False, italic=False,
             align=WD_ALIGN_PARAGRAPH.LEFT, space_after=9, space_before=0, font=FONT):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.line_spacing = 1.32
    add_richtext(p, text, size=size, color=color, base_bold=bold, italic=italic, font=font)
    return p

# ------------------------------------------------------------- block renderers

def render_head(doc, eyebrow, title, subtitle):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(2); p.paragraph_format.space_after = Pt(6)
    r = p.add_run(eyebrow); r.font.name = FONT; r.font.size = Pt(7.2); r.font.bold = True; r.font.color.rgb = GOLD

    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(title); r.font.name = FONT; r.font.size = Pt(18); r.font.bold = True; r.font.color.rgb = NAVY

    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(10)
    r = p.add_run(subtitle); r.font.name = FONT; r.font.size = Pt(8.5); r.font.italic = True; r.font.color.rgb = SUB_INK

    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(12)
    r = p.add_run("✦"); r.font.name = FONT; r.font.size = Pt(9); r.font.color.rgb = GOLD

def render_callout(doc, heading, text):
    table = doc.add_table(rows=1, cols=1); table.alignment = WD_TABLE_ALIGNMENT.CENTER; table.autofit = True
    cell = table.cell(0, 0)
    set_cell_border(cell, color="b79242", sz=8)
    set_cell_margins(cell, top=90, bottom=90, left=160, right=160)
    cell.paragraphs[0].paragraph_format.space_after = Pt(4)
    r = cell.paragraphs[0].add_run(heading.upper())
    r.font.name = FONT; r.font.size = Pt(8.3); r.font.bold = True; r.font.color.rgb = NAVY
    p2 = cell.add_paragraph(); p2.paragraph_format.line_spacing = 1.3
    add_richtext(p2, text, size=7.8, color=BODY_INK)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

def render_kv_table(doc, rows):
    table = doc.add_table(rows=0, cols=2); table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for k, v in rows:
        row = table.add_row(); c0, c1 = row.cells
        set_row_bottom_border(c0); set_row_bottom_border(c1)
        set_cell_margins(c0, top=60, bottom=60, left=0, right=100); set_cell_margins(c1, top=60, bottom=60, left=0, right=0)
        c0.width = Cm(6.5)
        r0 = c0.paragraphs[0].add_run(k); r0.font.name = FONT; r0.font.size = Pt(7.6); r0.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
        r1 = c1.paragraphs[0].add_run(v); r1.font.name = FONT; r1.font.size = Pt(7.6); r1.font.bold = True; r1.font.color.rgb = BLACK

def render_data_table(doc, headers, rows):
    table = doc.add_table(rows=1, cols=len(headers)); table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0]
    for i, h in enumerate(headers):
        cell = hdr.cells[i]; shade_cell(cell, "0a213b"); set_cell_margins(cell, top=70, bottom=70, left=110, right=110)
        r = cell.paragraphs[0].add_run(h); r.font.name = FONT; r.font.size = Pt(7.4); r.font.bold = True; r.font.color.rgb = WHITE
    for row_vals in rows:
        row = table.add_row()
        for i, val in enumerate(row_vals):
            cell = row.cells[i]; set_row_bottom_border(cell); set_cell_margins(cell, top=55, bottom=55, left=110, right=110)
            r = cell.paragraphs[0].add_run(str(val)); r.font.name = FONT; r.font.size = Pt(7.4); r.font.color.rgb = BLACK

def render_list(doc, items):
    for lead, rest in items:
        p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(6); p.paragraph_format.line_spacing = 1.25
        r = p.add_run("•  "); r.font.name = FONT; r.font.size = Pt(8.6); r.font.color.rgb = GOLD
        r = p.add_run(lead + " "); r.font.name = FONT; r.font.size = Pt(8.6); r.font.bold = True; r.font.color.rgb = NAVY
        r = p.add_run(rest); r.font.name = FONT; r.font.size = Pt(8.6); r.font.color.rgb = BODY_INK

def render_numlist(doc, items):
    for n, heading, text in items:
        p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(2)
        r = p.add_run(f"{n}. {heading}"); r.font.name = FONT; r.font.size = Pt(8); r.font.bold = True; r.font.color.rgb = GOLD
        p2 = doc.add_paragraph(); p2.paragraph_format.space_after = Pt(10); p2.paragraph_format.line_spacing = 1.3
        add_richtext(p2, text, size=9.0, color=BODY_INK)

def render_quote(doc, text):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(16); p.paragraph_format.line_spacing = 1.4
    r = p.add_run(text); r.font.name = FONT; r.font.size = Pt(10.5); r.font.italic = True; r.font.color.rgb = NAVY

def render_image(doc, path, width_cm):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(path, width=Cm(width_cm))

def render_blocks(doc, blocks):
    for b in blocks:
        kind = b[0]
        if kind == 'p':
            if b[1] and b[1].strip():
                add_para(doc, b[1])
        elif kind == 'callout':
            if b[2] and b[2].strip():
                render_callout(doc, b[1], b[2])
        elif kind == 'kv': render_kv_table(doc, b[1])
        elif kind == 'table': render_data_table(doc, b[1], b[2])
        elif kind == 'list': render_list(doc, b[1])
        elif kind == 'numlist': render_numlist(doc, b[1])
        elif kind == 'quote': render_quote(doc, b[1])
        elif kind == 'image': render_image(doc, b[1], b[2])
        elif kind == 'linelist':
            for line in b[1]:
                add_para(doc, line, size=8.6, space_after=4)

# --------------------------------------------------------------- natal wheel

def lon(sign, deg, minute):
    return SIGNS.index(sign) * 30 + deg + minute / 60.0

def _draw_wheel(ax, planets, houses, ascendant, bg="#0a213b", dot_r=0.028):
    """Desenha a roda natal dentro de `ax` (axes ja configurado, xlim/ylim
    simetricos o bastante para caber raio ~1.0-1.3)."""
    asc = lon(ascendant["sign"], ascendant["deg"], ascendant["min"])

    def xy(r, longitude):
        rel = (longitude - asc) % 360
        theta = math.radians(180 + rel)
        return r * math.cos(theta), r * math.sin(theta)

    if bg:
        ax.add_patch(mpatches.Circle((0, 0), 1.0, facecolor=bg, edgecolor="none", zorder=0))

    R_OUT, R_IN = 1.0, 0.32
    ax.add_patch(mpatches.Circle((0, 0), R_OUT, fill=False, edgecolor="#b79242", linewidth=1.4))
    ax.add_patch(mpatches.Circle((0, 0), R_IN, fill=False, edgecolor="#b79242", linewidth=1.4))
    ax.add_patch(mpatches.Circle((0, 0), (R_OUT + R_IN) / 2, fill=False, edgecolor="#b79242", linewidth=0.6, alpha=0.5))

    for h in houses:
        L = lon(h["sign"], h["deg"], h["min"])
        x1, y1 = xy(R_IN, L); x2, y2 = xy(R_OUT, L)
        ax.plot([x1, x2], [y1, y2], color="#b79242", linewidth=0.7, alpha=0.8)

    radii_cycle = [R_IN + (R_OUT - R_IN) * 0.30, R_IN + (R_OUT - R_IN) * 0.55, R_IN + (R_OUT - R_IN) * 0.80]
    used = {}
    for pl in planets:
        L = lon(pl["sign"], pl["deg"], pl["min"])
        bucket = round(L / 18)
        slot = used.get(bucket, 0); used[bucket] = slot + 1
        rr = radii_cycle[slot % len(radii_cycle)]
        x, y = xy(rr, L)
        color = PLANET_COLORS.get(pl["name"], "#cccccc")
        ax.add_patch(mpatches.Circle((x, y), dot_r, facecolor=color, edgecolor="#0a213b", linewidth=0.8, zorder=5))

def build_wheel_png(planets, houses, ascendant, out_path):
    """planets: list of dict(name, sign, deg, min); houses: list of dict(n, sign, deg, min);
    ascendant: dict(sign, deg, min)."""
    fig, ax = plt.subplots(figsize=(5.2, 5.2), dpi=300)
    fig.patch.set_facecolor("#0a213b"); ax.set_facecolor("#0a213b")
    ax.set_xlim(-1.3, 1.3); ax.set_ylim(-1.3, 1.3); ax.set_aspect("equal"); ax.axis("off")
    _draw_wheel(ax, planets, houses, ascendant, bg=None)
    fig.savefig(out_path, facecolor=fig.get_facecolor(), bbox_inches="tight", pad_inches=0.08)
    plt.close(fig)
    return out_path

_MESES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho",
          "agosto", "setembro", "outubro", "novembro", "dezembro"]

def _format_birth_line(birth):
    try:
        y, m, d = birth["date"].split("-")
        date_fmt = f"{int(d):02d} de {_MESES[int(m)-1]} de {y}"
    except Exception:
        date_fmt = birth.get("date", "")
    time_fmt = birth.get("time", "")
    return f"{date_fmt} • {time_fmt}h • {birth.get('place','')}"

def _photo_rgba_cover(photo_path, box_w_in, box_h_in, dpi, fade_right=0.0, fade_bottom=0.0, fade_left=0.0):
    """Abre a foto, corta no formato 'cover' (preenche a caixa sem distorcer)
    e aplica um degrade de transparencia nas bordas indicadas para a imagem
    se fundir organicamente no fundo navy, em vez de parecer colada."""
    from PIL import Image, ImageOps
    import numpy as np
    img = Image.open(photo_path)
    img = ImageOps.exif_transpose(img).convert("RGBA")

    target_w, target_h = int(box_w_in * dpi), int(box_h_in * dpi)
    src_w, src_h = img.size
    target_ratio, src_ratio = target_w / target_h, src_w / src_h
    if src_ratio > target_ratio:
        new_h = src_h
        new_w = int(src_ratio and target_ratio * src_h)
        x0 = (src_w - new_w) // 2
        img = img.crop((x0, 0, x0 + new_w, src_h))
    else:
        new_w = src_w
        new_h = int(src_w / target_ratio)
        y0 = max(0, (src_h - new_h) // 3)
        img = img.crop((0, y0, src_w, min(src_h, y0 + new_h)))
    img = img.resize((target_w, target_h), Image.LANCZOS)

    arr = np.array(img).astype(np.float64)
    alpha = arr[:, :, 3] / 255.0
    if fade_right > 0:
        w = arr.shape[1]
        fw = int(w * fade_right)
        if fw > 0:
            grad = np.linspace(1, 0, fw)
            alpha[:, w - fw:] *= grad[np.newaxis, :]
    if fade_left > 0:
        w = arr.shape[1]
        fw = int(w * fade_left)
        if fw > 0:
            grad = np.linspace(0, 1, fw)
            alpha[:, :fw] *= grad[np.newaxis, :]
    if fade_bottom > 0:
        h = arr.shape[0]
        fh = int(h * fade_bottom)
        if fh > 0:
            grad = np.linspace(1, 0, fh)
            alpha[h - fh:, :] *= grad[:, np.newaxis]
    arr[:, :, 3] = alpha * 255.0
    return arr / 255.0

def build_placeholder_cover_png(name, birth, planets, houses, ascendant, out_path, photo_path=None):
    """Capa navy/dourada completa: titulo, nome, data, triade, roda natal real
    e frase de fechamento. Se `photo_path` for informado, monta um layout de
    duas colunas (foto + roda natal) no espirito do modelo de referencia;
    caso contrario, usa um layout centralizado so com a roda natal."""
    fig_w, fig_h = 8.27, 11.69  # A4 portrait
    dpi = 220
    fig = plt.figure(figsize=(fig_w, fig_h), dpi=dpi)
    fig.patch.set_facecolor("#0a213b")

    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis("off")
    ax.set_facecolor("#0a213b")

    rng = __import__("random").Random(hash(name) % (2**31))
    for _ in range(110):
        sx, sy = rng.uniform(0.03, 0.97), rng.uniform(0.03, 0.97)
        ax.plot(sx, sy, marker="*", markersize=rng.uniform(1.2, 3.2),
                color="#c9a24b", alpha=rng.uniform(0.15, 0.55), zorder=1)

    ax.text(0.5, 0.955, "MAPA ASTRAL", ha="center", fontsize=27, color="#e9c98a",
            fontweight="bold", family="serif", zorder=3)
    ax.text(0.5, 0.920, "PERSONALIZADO", ha="center", fontsize=18, color="#e9c98a",
            fontweight="bold", family="serif", zorder=3)
    ax.plot([0.32, 0.68], [0.900, 0.900], color="#b79242", linewidth=0.8, alpha=0.7, zorder=3)

    ax.text(0.5, 0.868, name.upper(), ha="center", fontsize=15.5, color="#ffffff",
            family="serif", zorder=3)
    ax.text(0.5, 0.842, _format_birth_line(birth), ha="center", fontsize=10.5,
            color="#cbd3dc", family="serif", zorder=3)

    def pfind(nm):
        return next((p for p in planets if p["name"] == nm), None)
    sol, lua = pfind("Sol"), pfind("Lua")
    if sol and lua:
        triade = f'Sol em {sol["sign"]}  •  Lua em {lua["sign"]}  •  Ascendente em {ascendant["sign"]}'
        ax.text(0.5, 0.812, triade, ha="center", fontsize=11, color="#e9c98a",
                family="serif", style="italic", zorder=3)

    if photo_path:
        # ---- layout de duas colunas: foto a esquerda, roda natal a direita ----
        photo_box = (0.0, 0.235, 0.56, 0.545)  # x0, y0, w, h (fracoes da pagina)
        px0, py0, pw, ph = photo_box
        try:
            rgba = _photo_rgba_cover(photo_path, pw * fig_w, ph * fig_h, dpi, fade_right=0.35)
            ax.imshow(rgba, extent=[px0, px0 + pw, py0, py0 + ph], zorder=2, aspect="auto")
        except Exception:
            photo_path = None

        wheel_ax = fig.add_axes([0.50, 0.255, 0.47, 0.47 * fig_w / fig_h])
        wheel_ax.set_xlim(-1.18, 1.18); wheel_ax.set_ylim(-1.18, 1.18)
        wheel_ax.set_aspect("equal"); wheel_ax.axis("off"); wheel_ax.patch.set_alpha(0)
        _draw_wheel(wheel_ax, planets, houses, ascendant, bg=None, dot_r=0.038)

        # mini tabela de posicoes principais, abaixo da foto/roda
        rows = []
        for nm in ("Sol", "Lua", "Mercúrio", "Vênus", "Marte", "Ascendente", "Meio do Céu"):
            p = pfind(nm)
            if p:
                rows.append((nm, f'{p["deg"]:02d}°{p["min"]:02d} {p["sign"]}', p.get("house", "-")))
        ax.text(0.5, 0.205, "POSIÇÕES PRINCIPAIS", ha="center", fontsize=9.5, color="#e9c98a",
                family="serif", fontweight="bold", zorder=3)
        col_x = [0.14, 0.40, 0.75]
        row_y0, row_dy = 0.180, 0.0195
        for i, (nm, pos, house) in enumerate(rows):
            y = row_y0 - i * row_dy
            ax.text(col_x[0], y, nm, ha="left", fontsize=8.6, color="#ffffff", family="serif", zorder=3)
            ax.text(col_x[1], y, pos, ha="left", fontsize=8.6, color="#cbd3dc", family="serif", zorder=3)
            ax.text(col_x[2], y, f'casa {house}' if house != "-" else "—", ha="left", fontsize=8.6,
                    color="#cbd3dc", family="serif", zorder=3)

        ax.plot([0.32, 0.68], [0.052, 0.052], color="#b79242", linewidth=0.8, alpha=0.7, zorder=3)
        ax.text(0.5, 0.028, '"Um mapa não encerra uma pessoa. Ele abre uma conversa."',
                ha="center", va="center", fontsize=10, color="#e9c98a", style="italic",
                family="serif", zorder=3)
    else:
        # ---- layout centralizado (sem foto): so a roda natal, maior ----
        wheel_ax = fig.add_axes([0.145, 0.205, 0.71, 0.71 * fig_w / fig_h])
        wheel_ax.set_xlim(-1.18, 1.18); wheel_ax.set_ylim(-1.18, 1.18)
        wheel_ax.set_aspect("equal"); wheel_ax.axis("off"); wheel_ax.patch.set_alpha(0)
        _draw_wheel(wheel_ax, planets, houses, ascendant, bg=None, dot_r=0.038)

        ax.plot([0.30, 0.70], [0.135, 0.135], color="#b79242", linewidth=0.8, alpha=0.7, zorder=3)
        ax.text(0.5, 0.075, '"Um mapa não encerra uma pessoa.\nEle abre uma conversa."',
                ha="center", va="center", fontsize=12, color="#e9c98a", style="italic",
                family="serif", zorder=3, linespacing=1.8)

    fig.savefig(out_path, facecolor=fig.get_facecolor())
    plt.close(fig)
    return out_path

# -------------------------------------------------------------- page manifest

def build_pages(data):
    """Constroi a lista de 24 paginas de conteudo (sem a capa) a partir de `data`
    (ver SCHEMA.md). Campos ausentes usam texto de fallback minimo."""
    name = data["name"]
    s = data.get("sections", {})
    birth = data["birth"]
    planets = data["planets"]
    houses = data["houses"]

    def g(key, default=""):
        return s.get(key, default)

    pages = []

    pages.append(dict(eyebrow="LEITURA DO MAPA", title="ANTES DE COMEÇAR",
        subtitle=g("intro_subtitle", "Uma leitura simbólica para orientar reflexão, não um roteiro fechado sobre o futuro."),
        blocks=[
            ('callout', "Uma observação essencial", g("intro_observacao",
                "Este relatório utiliza a linguagem simbólica da astrologia ocidental tropical, com sistema de casas Placidus. "
                "As interpretações não constituem previsão científica nem determinam escolhas, acontecimentos ou capacidades.")),
            ('p', g("intro_p1", "")),
            ('p', g("intro_p2", "")),
            ('callout', "Como aproveitar melhor este relatório", g("intro_como_aproveitar",
                "Observe padrões recorrentes, compare-os com sua experiência e use a previsão anual como um convite à consciência.")),
        ]))

    pages.append(dict(eyebrow="1. DADOS DE NASCIMENTO", title="DADOS DE NASCIMENTO",
        subtitle="Pontos técnicos que organizam a leitura da carta.",
        blocks=[
            ('kv', [
                ("Nome completo", birth.get("full_name", name)),
                ("Data de nascimento", birth.get("date", "")),
                ("Horário informado", birth.get("time", "")),
                ("Local", birth.get("place", "")),
                ("Sistema zodiacal", birth.get("zodiac_system", "Astrologia ocidental tropical")),
                ("Sistema de casas", birth.get("house_system", "Placidus")),
                ("Ascendente calculado", birth.get("ascendant_label", "")),
                ("Meio do Céu calculado", birth.get("midheaven_label", "")),
            ]),
            ('p', g("dados_nascimento_texto", "")),
        ]))

    pages.append(dict(eyebrow="MAPA TÉCNICO", title="RODA NATAL TÉCNICA",
        subtitle="Visualização das casas, signos e principais planetas do nascimento.",
        blocks=[
            ('image', data["_wheel_png_path"], 10.5),
            ('p', g("mapa_tecnico_texto",
                "A roda natal organiza os doze signos no anel externo e as doze casas no anel interno. "
                "Os pontos coloridos indicam a posição simbólica de cada planeta.")),
        ]))

    pages.append(dict(eyebrow="2. SOL, LUA E ASCENDENTE", title="TRÍADE PRINCIPAL",
        subtitle="A estrutura que une identidade, emoção e presença.",
        blocks=[
            ('p', g("triade_sol", "")),
            ('p', g("triade_lua", "")),
            ('p', g("triade_ascendente", "")),
            ('callout', "Síntese da tríade", g("triade_sintese", "")),
        ]))

    planet_rows = [(p["name"], f'{p["deg"]:02d}°{p["min"]:02d} {p["sign"]}', p.get("house", "-"), p.get("key", "")) for p in planets]
    pages.append(dict(eyebrow="3. POSIÇÕES PLANETÁRIAS", title="POSIÇÕES PLANETÁRIAS E CASAS",
        subtitle="Os principais pontos do mapa natal e suas áreas de manifestação.",
        blocks=[
            ('table', ["Ponto", "Posição", "Casa", "Chave simbólica"], planet_rows),
            ('p', g("posicoes_texto", "")),
        ]))

    house_lines = []
    hs = {h["n"]: h for h in houses}
    for a, b in [(1, 7), (2, 8), (3, 9), (4, 10), (5, 11), (6, 12)]:
        ha, hb = hs.get(a), hs.get(b)
        if ha and hb:
            house_lines.append(f'Casa {a}: {ha["deg"]}°{ha["min"]:02d} {ha["sign"]}      Casa {b}: {hb["deg"]}°{hb["min"]:02d} {hb["sign"]}')
    pages.append(dict(eyebrow="CASAS ASTROLÓGICAS", title="CASAS E EIXOS DO MAPA",
        subtitle="As áreas da vida que recebem maior densidade simbólica.",
        blocks=[
            ('p', g("casas_texto", "")),
            ('linelist', house_lines),
            ('callout', g("eixo1_title", "Eixo Ascendente–Descendente"), g("eixo1_text", "")),
            ('callout', g("eixo2_title", "Eixo Fundo do Céu–Meio do Céu"), g("eixo2_text", "")),
        ]))

    pages.append(dict(eyebrow="4. PERSONALIDADE", title="PERSONALIDADE E TRAÇOS CENTRAIS",
        subtitle=g("personalidade_subtitle", "Traços centrais da personalidade."),
        blocks=[('p', t) for t in g("personalidade", [])] +
               [('callout', "Assinatura pessoal", g("personalidade_callout", ""))]))

    pages.append(dict(eyebrow="5. EMOÇÕES E RELAÇÕES", title="MUNDO EMOCIONAL E RELACIONAMENTOS",
        subtitle=g("emocional_subtitle", "O mundo emocional e os vínculos afetivos."),
        blocks=[('p', t) for t in g("emocional", [])] +
               [('callout', "Relações que nutrem", g("emocional_callout", ""))]))

    pages.append(dict(eyebrow="6. COMUNICAÇÃO E MENTE", title="COMUNICAÇÃO, PENSAMENTO E DECISÕES",
        subtitle=g("comunicacao_subtitle", "Como a mente processa, decide e se expressa."),
        blocks=[('p', t) for t in g("comunicacao", [])] +
               [('callout', "Para decidir melhor", g("comunicacao_callout", ""))]))

    pages.append(dict(eyebrow="7. AFETO E VALORES", title="AMOR, AFETO, DESEJO E VALORES",
        subtitle=g("afeto_subtitle", "Amor, afeto, desejo e valores pessoais."),
        blocks=[('p', t) for t in g("afeto", [])] +
               [('callout', "Sua linguagem afetiva", g("afeto_callout", ""))]))

    pages.append(dict(eyebrow="8. CARREIRA E PROPÓSITO", title="CARREIRA, IMAGEM PÚBLICA E PROPÓSITO",
        subtitle=g("carreira_subtitle", "Carreira, imagem pública e propósito de vida."),
        blocks=[('p', t) for t in g("carreira", [])] +
               [('callout', "Vocação simbólica", g("carreira_callout", ""))]))

    pages.append(dict(eyebrow="9. TALENTOS", title="FORÇAS E TALENTOS NATURAIS",
        subtitle="O que ganha potência quando você confia na própria natureza.",
        blocks=[('list', [(t["lead"], t["text"]) for t in g("talentos", [])]),
                ('callout', "Talento central", g("talentos_callout", ""))]))

    pages.append(dict(eyebrow="10. DESENVOLVIMENTO PESSOAL", title="OPORTUNIDADES DE CRESCIMENTO",
        subtitle="Pontos que podem se tornar mais livres e conscientes.",
        blocks=[('numlist', [(i + 1, it["title"], it["text"]) for i, it in enumerate(g("crescimento", []))])]))

    pages.append(dict(eyebrow="11. ASPECTOS PRINCIPAIS - I", title="ASPECTOS ASTROLÓGICOS EM DESTAQUE",
        subtitle="Diálogos simbólicos que se combinam com naturalidade.",
        blocks=[('kv', [(a["aspecto"], f'Orbe aproximado: {a["orbe"]}') for a in g("aspectos_harmonicos", [])]),
                ('p', g("aspectos_harmonicos_texto", ""))]))

    pages.append(dict(eyebrow="11. ASPECTOS PRINCIPAIS - II", title="ASPECTOS ASTROLÓGICOS EM DESTAQUE",
        subtitle="Tensões que pedem consciência, escolha e refinamento.",
        blocks=[('kv', [(a["aspecto"], f'Orbe aproximado: {a["orbe"]}') for a in g("aspectos_tensos", [])]),
                ('p', g("aspectos_tensos_texto", ""))]))

    pages.append(dict(eyebrow="12. CICLOS E TRÂNSITOS", title="PREVISÃO SIMBÓLICA PARA OS PRÓXIMOS 12 MESES",
        subtitle=g("previsao_subtitle", ""),
        blocks=[('p', t) for t in g("previsao_overview", [])]))

    for q in g("trimestres", []):
        pages.append(dict(eyebrow=q["eyebrow"], title="PREVISÃO SIMBÓLICA", subtitle=q["subtitle"], blocks=[
            ('p', q["p1"]), ('p', q["p2"]),
            ('callout', q["opp_title"], q["opp_text"]),
            ('callout', q["dec_title"], q["dec_text"]),
        ]))

    pages.append(dict(eyebrow="13. SÍNTESE FINAL", title="RESUMO FINAL",
        subtitle="A essência que se destaca quando se observa o mapa como conjunto.",
        blocks=[('p', t) for t in g("sintese_final", [])] +
               [('quote', g("sintese_quote", ""))]))

    if houses:
        cusp_rows = [(f'Casa {h["n"]}', f'{h["deg"]}°{h["min"]:02d} {h["sign"]}') for h in sorted(houses, key=lambda x: x["n"])]
        pages.append(dict(eyebrow="DADOS COMPLEMENTARES", title="APÊNDICE TÉCNICO",
            subtitle="Cúspides das casas calculadas para esta leitura.",
            blocks=[('p', g("apendice_texto", "")), ('table', ["Casa", "Cúspide"], cusp_rows)]))

    all_asp = g("aspectos_harmonicos", []) + g("aspectos_tensos", [])
    if all_asp:
        pages.append(dict(eyebrow="DADOS COMPLEMENTARES", title="ASPECTOS SELECIONADOS",
            subtitle="Relações com orbes aproximados.",
            blocks=[('table', ["Aspecto", "Orbe aproximado"], [(a["aspecto"], a["orbe"]) for a in all_asp])]))

    pages.append(dict(eyebrow="NOTA FINAL", title="FECHO",
        subtitle="Um mapa não encerra uma pessoa. Ele abre uma conversa.",
        blocks=[('p', t) for t in g("nota_final", [])] +
               [('quote', g("nota_quote", g("sintese_quote", "")))]))

    return pages

# ------------------------------------------------------------------ document

def build_docx_bytes(data, cover_image_path=None):
    """data: dict seguindo SCHEMA.md. cover_image_path: caminho para JPG/PNG de
    capa (retrato ja gerado, se existir); se None, uma capa navy/dourada simples
    e desenhada automaticamente."""
    tmp_dir = tempfile.mkdtemp(prefix="report_")
    wheel_png = os.path.join(tmp_dir, "wheel.png")
    build_wheel_png(data["planets"], data["houses"], data["ascendant"], wheel_png)
    data = dict(data)
    data["_wheel_png_path"] = wheel_png

    uploaded_photo_path = cover_image_path
    cover_image_path = os.path.join(tmp_dir, "cover.png")
    build_placeholder_cover_png(data["name"], data["birth"], data["planets"],
                                 data["houses"], data["ascendant"], cover_image_path,
                                 photo_path=uploaded_photo_path)

    name = data["name"]
    running = f"MAPA ASTRAL PERSONALIZADO • {name.upper()}"

    doc = Document()
    normal = doc.styles['Normal']
    normal.font.name = FONT; normal.font.size = Pt(9.2); normal.font.color.rgb = BODY_INK

    sec0 = doc.sections[0]
    sec0.page_width = Cm(21.0); sec0.page_height = Cm(29.7)
    sec0.left_margin = sec0.right_margin = sec0.top_margin = sec0.bottom_margin = Cm(0)
    sec0.header_distance = sec0.footer_distance = Cm(0)

    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(0); p.paragraph_format.space_before = Pt(0)
    p.add_run().add_picture(cover_image_path, width=Cm(21.0), height=Cm(29.7))

    sec1 = doc.add_section(WD_SECTION.NEW_PAGE)
    sec1.page_width = Cm(21.0); sec1.page_height = Cm(29.7)
    sec1.left_margin = sec1.right_margin = Cm(2.35)
    sec1.top_margin = sec1.bottom_margin = Cm(1.55)
    sec1.header_distance = Cm(0.5); sec1.footer_distance = Cm(0.6)
    sec1.header.is_linked_to_previous = False
    sec1.footer.is_linked_to_previous = False

    hp = sec1.header.paragraphs[0]; hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hp.paragraph_format.space_after = Pt(6)
    hr = hp.add_run(running); hr.font.name = FONT; hr.font.size = Pt(6.5); hr.font.bold = True; hr.font.color.rgb = GOLD
    para_bottom_border(hp, color=GOLD_LINE, sz=5)

    fp = sec1.footer.paragraphs[0]; fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fr = fp.add_run(running + " • "); fr.font.name = FONT; fr.font.size = Pt(6.3); fr.font.color.rgb = FOOTER_INK
    add_page_field(fp)

    pages = build_pages(data)
    for i, pg in enumerate(pages):
        if i > 0:
            doc.add_page_break()
        render_head(doc, pg['eyebrow'], pg['title'], pg['subtitle'])
        render_blocks(doc, pg['blocks'])

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()
