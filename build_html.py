#!/usr/bin/env python3
"""
Genera la version HTML autocontenida del template "Marketing" (mismas 13 slides
que el .pptx), para inspeccionar en navegador y conectar a herramientas de diseno.
Armin Grotesk embebida en base64, logos SVG inline. Coordenadas = pulgadas x 96px.
"""
import os, base64

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, "assets")
OUT = os.path.join(HERE, "perpetual-marketing-template.html")

# --- fuentes OTF -> @font-face base64 ---
FONTS = [("Normal", 300), ("Regular", 400), ("Semi_Bold", 600), ("Black", 800)]
faces = []
for name, weight in FONTS:
    data = open(os.path.join(ASSETS, "fonts", f"ArminGrotesk_{name}.otf"), "rb").read()
    b64 = base64.b64encode(data).decode()
    faces.append("@font-face{font-family:'Armin Grotesk';font-weight:%d;font-display:swap;"
                 "src:url(data:font/otf;base64,%s) format('opentype');}" % (weight, b64))
FONT_FACES = "\n".join(faces)


def _svg(path):
    return open(os.path.join(ASSETS, "logo", path)).read().split("?>", 1)[-1].strip()
LOGO_COLOR, LOGO_DARK = _svg("perpetual-color.svg"), _svg("perpetual-dark.svg")

# --- tokens ---
ACCENT, ACCENT2, YELLOW = "#1a56db", "#f97316", "#fbb900"
BGD, TEXT, DIM, MUTED = "#0b1220", "#111827", "#374151", "#6b7280"
SURFACE, SURFACE2, BORDER, WHITE, DBE4FF = "#f8f9fc", "#eef1f8", "#dde1ef", "#ffffff", "#dbe4ff"
PXIN = 96


def _p(v):
    return f"{v * PXIN:.1f}px"


def box(x, y, w, h, fill=None, r=0, oval=False, shadow=False, line=None):
    st = f"position:absolute;left:{_p(x)};top:{_p(y)};width:{_p(w)};height:{_p(h)};"
    st += "border-radius:50%;" if oval else (f"border-radius:{r}px;" if r else "")
    if fill: st += f"background:{fill};"
    if line: st += f"border:1px solid {line};"
    if shadow: st += "box-shadow:0 8px 26px rgba(20,40,90,.13);"
    return f'<div style="{st}"></div>'


def txt(x, y, w, h, content, size, color=TEXT, weight=400, align="left",
        valign="top", spacing=None, upper=False, lh=1.1):
    just = {"top": "flex-start", "middle": "center", "bottom": "flex-end"}[valign]
    st = (f"position:absolute;left:{_p(x)};top:{_p(y)};width:{_p(w)};height:{_p(h)};"
          f"display:flex;flex-direction:column;justify-content:{just};overflow:hidden;"
          f"font-size:{size*1.333:.1f}px;color:{color};font-weight:{weight};"
          f"text-align:{align};line-height:{lh};")
    if align == "center": st += "align-items:center;"
    if spacing: st += f"letter-spacing:{spacing}px;"
    if upper: st += "text-transform:uppercase;"
    return f'<div style="{st}">{content}</div>'


def logo(x, y, w, dark=False):
    svg = LOGO_DARK if dark else LOGO_COLOR
    st = f"position:absolute;left:{_p(x)};top:{_p(y)};width:{_p(w)};"
    return f'<div class="lg" style="{st}">{svg}</div>'


def hexagon(x, y, size, fill):
    st = (f"position:absolute;left:{_p(x)};top:{_p(y)};width:{_p(size)};height:{_p(size)};"
          f"background:{fill};clip-path:polygon(25% 0,75% 0,100% 50%,75% 100%,25% 100%,0 50%);")
    return f'<div style="{st}"></div>'


def blob(x, y, d, fill):
    return box(x, y, d, d, fill=fill, oval=True)


def pill(x, y, w, label, fill=ACCENT, fg=WHITE, arrow=True):
    out = [box(x, y, w, 0.62, fill=fill, r=31, shadow=True),
           txt(x + 0.34, y, w - 1.0, 0.62, label, 11.5, fg, 600, "left", "middle",
               spacing=0.8, upper=True)]
    if arrow:
        out.append(box(x + w - 0.74, y + 0.1, 0.42, 0.42, fill=WHITE, oval=True))
        out.append(f'<div style="position:absolute;left:{_p(x + w - 0.74)};top:{_p(y + 0.1)};width:{_p(0.42)};height:{_p(0.42)};display:flex;align-items:center;justify-content:center"><svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="{fill}" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><polyline points="9 6 15 12 9 18"/></svg></div>')
    return "".join(out)


def photo_ph(x, y, w, h, r=12, tint="#E3ECFB"):
    d = min(w, h) * 0.24
    cxp, cyp = x + w / 2, y + h / 2
    return (box(x, y, w, h, fill=tint, r=r)
            + box(cxp - d / 2, cyp - d / 2, d, d, fill=WHITE, oval=True)
            + box(cxp - d * 0.16, cyp - d * 0.16, d * 0.32, d * 0.32, fill=ACCENT, oval=True))


def tool_icon(x, y, d, fill):
    return box(x, y, d, d, fill=fill, oval=True) + hexagon(x + d * 0.3, y + d * 0.3, d * 0.4, WHITE)


def graphic(x, y, w, h, tint="#DBE7FB", variant="abstract", r=12, shadow=False):
    """Grafico de marca (en vez de foto): composicion abstracta on-brand.
    En Perpetual no usamos fotos de personas salvo en la slide de equipo."""
    out = [box(x, y, w, h, fill=tint, r=r, shadow=shadow)]
    cx, cy = x + w / 2, y + h / 2
    if variant == "growth":
        n, bw, gap = 4, w * 0.13, w * 0.06
        total = n * bw + (n - 1) * gap
        bx, base = cx - total / 2, y + h * 0.8
        cols = [ACCENT, ACCENT2, YELLOW, ACCENT]
        for i in range(n):
            bh = h * (0.16 + 0.13 * i)
            out.append(box(bx + i * (bw + gap), base - bh, bw, bh, fill=cols[i], r=4))
        out.append(box(cx - w * 0.3, y + h * 0.16, h * 0.2, h * 0.2, fill=ACCENT, oval=True))
        out.append(hexagon(cx + w * 0.16, y + h * 0.14, h * 0.16, YELLOW))
    elif variant == "quote":
        out.append(txt(x, y + h * 0.06, w, h * 0.45, "&ldquo;", 92, ACCENT, 800, "center"))
        out.append(txt(x, y + h * 0.62, w, h * 0.2,
                       "&#9733; &#9733; &#9733; &#9733; &#9733;", 17, YELLOW, 700, "center"))
    else:  # abstract: circulos + hexagono de marca
        out.append(box(cx - w * 0.28, cy - h * 0.16, h * 0.34, h * 0.34, fill=ACCENT, oval=True))
        out.append(box(cx + w * 0.03, cy - h * 0.02, h * 0.22, h * 0.22, fill=ACCENT2, oval=True))
        out.append(box(cx - w * 0.02, cy + h * 0.16, h * 0.13, h * 0.13, fill=YELLOW, oval=True))
        out.append(hexagon(cx + w * 0.12, cy - h * 0.26, h * 0.17, WHITE))
    return "".join(out)


def title(runs, x=0.7, y=0.7, w=7.5, size=33):
    return logo(0.6, 0.5, 1.15) + txt(x, y + 0.55, w, 1.2, runs, size, TEXT, 800, lh=1.0)


def footer(page):
    return (txt(0.55, 7.0, 8, 0.3, "Confidencial &middot; Perpetual Technologies &copy; 2026",
                8.5, MUTED, 400, "left", "middle")
            + txt(11.7, 7.0, 1.1, 0.3, str(page).zfill(2), 8.5, MUTED, 400, "right", "middle"))


def AC(t):  # helper: envuelve en span de acento
    return f'<span style="color:{ACCENT}">{t}</span>'


# ===========================================================================
# Slides (mismas coordenadas que el .pptx)
# ===========================================================================
def m01():
    return (box(0, 4.55, 13.333, 2.95, fill=ACCENT)
            + blob(9.0, -1.0, 3.2, YELLOW) + blob(11.6, 2.2, 2.0, ACCENT2) + blob(8.0, 1.5, 3.3, ACCENT)
            + hexagon(8.55, 2.35, 2.1, WHITE) + box(9.25, 3.05, 0.7, 0.7, fill=ACCENT, oval=True)
            + blob(11.4, 4.7, 0.8, YELLOW)
            + txt(12.3, 0.55, 0.7, 0.6, "+", 30, ACCENT, 800)
            + txt(7.4, 5.4, 0.6, 0.5, "+", 24, ACCENT2, 800)
            + logo(0.7, 0.7, 1.5)
            + txt(0.65, 1.5, 7.0, 2.0, f"Marketing para<br>{AC('tu marca.')}", 46, TEXT, 800, lh=0.98)
            + txt(0.7, 3.5, 6.2, 0.3, "Estrategia y crecimiento", 13, TEXT, 600)
            + txt(0.7, 3.85, 6.0, 0.8, "Convertimos tu marca en resultados medibles con data y creatividad.",
                  12.5, MUTED, 400, lh=1.3)
            + pill(0.7, 5.45, 3.0, "Empezar", fill=WHITE, fg=ACCENT))


def m02():
    return (title(f"Bienvenido a {AC('Perpetual.')}")
            + blob(5.2, 1.9, 1.7, ACCENT) + graphic(4.7, 1.9, 4.3, 3.0, tint="#DBE7FB", variant="abstract")
            + box(5.0, 4.55, 2.7, 0.34, fill=ACCENT2, r=17)
            + txt(5.0, 4.55, 2.7, 0.34, "Soporte de marketing", 10.5, WHITE, 600, "center", "middle", spacing=1, upper=True)
            + txt(0.7, 2.55, 3.5, 0.7, "2.9M", 40, ACCENT, 800)
            + txt(0.7, 3.4, 3.5, 0.3, "Impresiones gestionadas", 11, MUTED, 600, upper=True, spacing=0.6)
            + txt(0.7, 3.9, 3.6, 2.5, "Acompanamos a tu equipo de principio a fin: estrategia, ejecucion y medicion en un solo lugar.",
                  12.5, MUTED, 400, lh=1.35)
            + graphic(9.4, 2.3, 3.3, 3.6, tint="#FDE9D6", variant="abstract") + footer(2))


def m03():
    out = [title(f"Conoce a nuestro<br>{AC('equipo.')}", w=5.0, size=30)]
    people = [("Ana Rojas", "Marketing Lead"), ("Luis Vega", "Data Analyst"),
              ("Mara Diaz", "Creative Dir."), ("Ivan Soto", "Growth")]
    for i, (n, role) in enumerate(people):
        x = 6.0 + i * 1.75
        out.append(photo_ph(x, 1.6, 1.55, 2.2, tint="#DBE7FB" if i % 2 == 0 else "#FDE9D6"))
        out.append(txt(x, 3.9, 1.65, 0.3, n, 11.5, TEXT, 600))
        out.append(txt(x, 4.2, 1.65, 0.3, role, 9.5, ACCENT, 600, upper=True, spacing=0.4))
    out.append(blob(11.9, 5.2, 1.0, YELLOW))
    out.append(txt(0.7, 4.6, 4.8, 1.2, "Un equipo multidisciplinario enfocado en resultados, no en vanity metrics.",
                   13, MUTED, 400, lh=1.35))
    out.append(footer(3))
    return "".join(out)


def m04():
    return (title(f"Sobre el trabajo<br>{AC('actual.')}", size=30)
            + blob(6.7, 1.3, 3.0, ACCENT) + graphic(7.6, 1.5, 3.4, 2.4, tint="#DBE7FB", variant="abstract")
            + graphic(6.2, 3.4, 3.4, 2.3, tint="#FDE9D6", variant="abstract") + blob(9.7, 4.4, 1.9, ACCENT2)
            + pill(9.5, 4.9, 2.7, "Nueva via del exito", fill=ACCENT2, fg=WHITE, arrow=False)
            + txt(0.7, 2.8, 4.6, 1.8, "Cada cuenta tiene un plan vivo: lo medimos cada semana y ajustamos con base en datos.",
                  13, MUTED, 400, lh=1.35)
            + txt(0.7, 4.7, 4.0, 0.3, "Auditoria y diagnostico", 12, ACCENT, 600)
            + txt(0.7, 5.1, 4.0, 0.3, "Plan de medios y contenido", 12, ACCENT, 600) + footer(4))


def m05():
    out = [title(f"Lo que {AC('ofrecemos.')}"),
           txt(8.2, 1.0, 4.4, 0.9, "Servicios end to end para que tu marca crezca con foco en retorno.",
               12.5, MUTED, 400, lh=1.3)]
    cards = [("SEO y contenido", ACCENT), ("Performance ads", BGD), ("Brand y creatividad", ACCENT2)]
    for i, (t, fill) in enumerate(cards):
        x = 0.7 + i * 4.1
        out += [box(x, 2.6, 3.7, 3.9, fill=fill, r=14, shadow=True),
                box(x + 0.35, 2.95, 0.7, 0.7, fill=WHITE, oval=True),
                hexagon(x + 0.5, 3.1, 0.4, fill),
                txt(x + 0.4, 4.0, 3.0, 0.7, t, 15, WHITE, 600, lh=1.05),
                txt(x + 0.4, 4.9, 3.0, 1.4, "Estrategia, ejecucion y reporte claro de cada iniciativa.",
                    11, WHITE, 400, lh=1.3)]
    out.append(footer(5))
    return "".join(out)


def m06():
    out = [title(f"Tendencias de {AC('mercado.')}"),
           blob(3.0, 2.7, 2.9, ACCENT2), blob(5.2, 3.2, 2.5, YELLOW),
           txt(3.0, 3.7, 2.9, 0.7, "3.8M", 30, WHITE, 800, "center"),
           txt(5.2, 4.0, 2.5, 0.6, "40.5K", 26, TEXT, 800, "center"),
           txt(0.7, 3.4, 2.1, 0.8, "Alcance<br>mensual", 11, MUTED, 600, upper=True, lh=1.1),
           txt(8.1, 3.4, 2.4, 0.8, "Nuevos<br>seguidores", 11, MUTED, 600, upper=True, lh=1.1)]
    # mini bar chart
    vals = [40, 55, 48, 70, 82]
    for i, v in enumerate(vals):
        hh = v / 82 * 2.6
        out.append(box(9.5 + i * 0.62, 6.0 - hh, 0.42, hh, fill=YELLOW, r=4))
    out.append(footer(6))
    return "".join(out)


def m07():
    out = [title(f"Nuestras {AC('soluciones.')}"),
           txt(8.4, 1.0, 4.2, 0.9, "Dos frentes de trabajo que se refuerzan entre si.", 12.5, MUTED, 400, lh=1.3)]
    rows = [("O1", "Adquisicion", "Campanas de performance optimizadas a costo por resultado.", ACCENT),
            ("O2", "Retencion", "Contenido y automatizaciones que sostienen el crecimiento.", BGD)]
    for i, (n, t, d, fill) in enumerate(rows):
        y = 2.7 + i * 1.85
        out += [box(0.7, y, 11.9, 1.55, fill=fill, r=14, shadow=True),
                txt(1.1, y + 0.35, 1.2, 0.8, n, 26, WHITE, 800),
                txt(2.3, y + 0.3, 3.0, 0.4, t, 15, WHITE, 600),
                txt(2.3, y + 0.75, 9.6, 0.6, d, 12, DBE4FF, 400, lh=1.25)]
    out.append(footer(7))
    return "".join(out)


def m08():
    out = [title(f"Herramientas de {AC('marketing.')}")]
    tools = [("Audio social", ACCENT), ("SEO", ACCENT2), ("Video", BGD),
             ("Social", YELLOW), ("Automatizacion", ACCENT)]
    for i, (t, col) in enumerate(tools):
        x = 0.8 + i * 2.45
        out += [tool_icon(x, 2.9, 1.4, col),
                txt(x - 0.25, 4.5, 1.9, 0.6, t, 11, TEXT, 600, "center", lh=1.05),
                txt(x - 0.25, 5.15, 1.9, 0.5, "Stack y procesos", 9.5, MUTED, 400, "center")]
    out.append(footer(8))
    return "".join(out)


def m09():
    out = [title(f"Opiniones de {AC('clientes.')}", w=4.5),
           txt(5.0, 1.05, 3.0, 0.8, "4,890 +", 34, ACCENT, 800),
           graphic(9.7, 1.0, 3.0, 3.4, tint="#DBE7FB", variant="quote"), blob(12.2, 3.6, 1.0, YELLOW)]
    for i, icon in enumerate([ACCENT2, ACCENT]):
        y = 2.7 + i * 1.55
        fill = SURFACE if i == 0 else ACCENT
        fg = DIM if i == 0 else WHITE
        out += [box(0.7, y, 8.5, 1.35, fill=fill, r=16, shadow=True, line=BORDER if i == 0 else None),
                blob(1.0, y + 0.35, 0.65, icon),
                txt(1.95, y + 0.25, 7.0, 0.9, "El equipo entiende el negocio y entrega resultados claros mes a mes.",
                    12, fg, 400, lh=1.25)]
    out.append(footer(9))
    return "".join(out)


def m10():
    out = [title(f"Nueva via del {AC('exito.')}"), graphic(0.7, 2.7, 3.2, 3.6, tint="#DBE7FB", variant="abstract")]
    steps = [("O2", "Diagnostico", ACCENT, 4.4, 2.6), ("O3", "Estrategia", ACCENT2, 8.3, 2.6),
             ("O4", "Ejecucion", YELLOW, 4.4, 4.5), ("O5", "Optimizacion", BGD, 8.3, 4.5)]
    for n, t, col, x, y in steps:
        out += [blob(x, y, 0.85, col), hexagon(x + 0.26, y + 0.26, 0.34, WHITE),
                txt(x + 1.05, y + 0.02, 1.0, 0.3, n, 12, MUTED, 600),
                txt(x + 1.05, y + 0.32, 3.3, 0.35, t, 14, TEXT, 600),
                txt(x + 1.05, y + 0.72, 3.3, 0.4, "Paso del proceso con foco en dato.", 10.5, MUTED, 400)]
    out.append(footer(10))
    return "".join(out)


def m11():
    out = [title(f"Plan de estrategia<br>{AC('de marketing.')}", size=30)]
    blocks = [(ACCENT2, WHITE, 7.0, 1.2, 5.2, "Objetivos y KPIs del trimestre"),
              (BGD, WHITE, 6.0, 2.5, 4.6, "Mensajes y audiencias clave"),
              (ACCENT, WHITE, 7.4, 3.8, 5.0, "Mezcla de canales y presupuesto"),
              (SURFACE, DIM, 6.4, 5.1, 5.4, "Calendario y responsables")]
    for fill, fg, x, y, w, t in blocks:
        out += [box(x, y, w, 1.0, fill=fill, r=14, shadow=True, line=BORDER if fill == SURFACE else None),
                txt(x + 0.35, y, w - 0.7, 1.0, t, 13, fg, 600, "left", "middle")]
    out.append(txt(0.7, 2.9, 4.6, 2.0, "Un plan claro, por fases, con metas medibles y duenos definidos en cada etapa.",
                   13, MUTED, 400, lh=1.35))
    out.append(footer(11))
    return "".join(out)


def m12():
    out = [title(f"Crece con {AC('nosotros.')}"),
           box(6.8, 1.7, 5.9, 4.6, fill=ACCENT, r=16, shadow=True),
           txt(7.2, 2.1, 4.0, 0.8, "+58%", 40, WHITE, 800),
           txt(7.2, 3.0, 5.0, 0.4, "Crecimiento interanual promedio", 12, DBE4FF, 400)]
    for vx, vy in [(7.4, 5.3), (8.6, 4.9), (9.8, 5.0), (11.0, 4.4), (12.0, 4.0)]:
        out.append(blob(vx, vy, 0.16, YELLOW))
    out += [txt(0.7, 2.65, 3.5, 0.7, "100K", 38, ACCENT, 800),
            txt(0.7, 3.45, 3.5, 0.3, "Leads generados", 11, MUTED, 600, upper=True, spacing=0.6),
            txt(0.7, 4.25, 3.5, 0.7, "0.2M", 38, TEXT, 800),
            txt(0.7, 5.05, 3.5, 0.3, "Ventas atribuidas", 11, MUTED, 600, upper=True, spacing=0.6),
            blob(4.9, 1.5, 0.7, YELLOW), footer(12)]
    return "".join(out)


def m13():
    out = [box(0, 0, 13.333, 3.0, fill=ACCENT),
           blob(10.8, -0.8, 2.2, YELLOW), blob(12.2, 1.8, 1.4, ACCENT2),
           logo(0.7, 0.6, 1.5, dark=True),
           pill(0.7, 1.6, 3.0, "Empezar", fill=WHITE, fg=ACCENT),
           txt(4.0, 1.55, 7.0, 0.8, "Hablemos de tu proxima campana.", 15, WHITE, 600, "left", "middle"),
           txt(0.7, 3.7, 6.0, 1.0, f"Contac{AC('tanos.')}", 40, TEXT, 800)]
    cols = [("Correo", "hola@perpetual.pe"), ("Telefono", "+51 999 999 999"), ("Web", "perpetual.pe")]
    for i, (t, v) in enumerate(cols):
        x = 0.7 + i * 4.1
        out += [txt(x, 5.4, 3.7, 0.3, t, 11, ACCENT, 600, upper=True, spacing=0.6),
                txt(x, 5.75, 3.7, 0.4, v, 14, TEXT, 600)]
    out.append(blob(11.8, 5.6, 1.0, ACCENT))
    return "".join(out)


SLIDES = [m01, m02, m03, m04, m05, m06, m07, m08, m09, m10, m11, m12, m13]
stages = "\n".join(f'<div class="slide">{fn()}</div>' for fn in SLIDES)

HTML = f"""<!doctype html><html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Perpetual &middot; Template Marketing</title>
<style>
{FONT_FACES}
*{{margin:0;padding:0;box-sizing:border-box}}
body{{background:#c9ccd6;font-family:'Armin Grotesk',system-ui,sans-serif;padding:30px 0}}
.deck{{width:1280px;margin:0 auto;display:flex;flex-direction:column;gap:24px}}
.slide{{position:relative;width:1280px;height:720px;background:#fff;overflow:hidden;
  border-radius:16px;box-shadow:0 10px 40px rgba(0,0,0,.18)}}
.lg svg{{display:block;width:100%;height:auto}}
</style></head><body>
<div class="deck">
{stages}
</div>
</body></html>"""

with open(OUT, "w", encoding="utf-8") as f:
    f.write(HTML)
print("OK:", OUT, "|", round(len(HTML) / 1024), "KB |", len(SLIDES), "slides")
