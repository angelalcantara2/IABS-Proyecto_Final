"""
armar_bloques.py: genera el documento corregido de bloques y grafo (Word y PDF)
atendiendo los comentarios del profesor.

    python armar_bloques.py

Salida: documentos/SSDC_Bloques_y_Grafo_corregido.docx y .pdf
"""
import os
import math
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from lector import lectura_datos, ruta_archivo
import grafo as gf
from busqueda import construir_heuristica
from reporte import dir_res
from pdf_util import exportar_pdf
import analisis_bloques as ab
import simulaciones as sim

raiz = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.makedirs(os.path.join(raiz, "documentos"), exist_ok=True)
salida = os.path.join(raiz, "documentos", "SSDC_Bloques_y_Grafo_corregido.docx")
AUTORES = "Juan Nicolas Álvarez, Ángel Samuel Alcántara y Mateo Sánchez García"
FUENTE = "Times New Roman"
AZUL, NARANJA, AQUA = "#2a78d6", "#eb6834", "#1baf7a"
GRIS, TINTA = "#52514e", "#0b0b0b"


def n0(x):
    return f"{x:,.0f}".replace(",", ".")


def d(x, k=2):
    return f"{x:.{k}f}".replace(".", ",")


# ================= figuras =================
def caja(ax, x, y, w, h, txt, color="#eef3f8", borde="#2e5578", fs=8, negrita=False):
    ax.add_patch(FancyBboxPatch((x - w / 2, y - h / 2), w, h, boxstyle="round,pad=0.02,rounding_size=0.08",
                                fc=color, ec=borde, lw=1))
    ax.text(x, y, txt, ha="center", va="center", fontsize=fs, color=TINTA,
            fontweight="bold" if negrita else "normal")


def flecha(ax, a, b, txt=None, dy=0.13, rad=0.0, fs=6.5, col=GRIS):
    ax.add_patch(FancyArrowPatch(a, b, arrowstyle="-|>", mutation_scale=9, color=col, lw=0.9,
                                 connectionstyle=f"arc3,rad={rad}"))
    if txt:
        ax.text((a[0] + b[0]) / 2, (a[1] + b[1]) / 2 + dy, txt, ha="center", va="bottom",
                fontsize=fs, color=GRIS)


def fig_bloques(ruta):
    fig, ax = plt.subplots(figsize=(7.2, 3.4), dpi=220)
    ax.set_xlim(0, 10.4); ax.set_ylim(0, 4.6); ax.axis("off")
    caja(ax, 1.0, 3.6, 1.6, 0.8, "B1\nLectura de datos", negrita=True)
    caja(ax, 3.6, 3.6, 1.8, 0.8, "B2\nCosto unificado", negrita=True)
    caja(ax, 6.3, 3.6, 1.7, 0.8, "B3\nConstrucción\ndel grafo", negrita=True)
    # módulo de búsqueda
    ax.add_patch(FancyBboxPatch((7.75, 0.55), 2.5, 3.55, boxstyle="round,pad=0.02,rounding_size=0.1",
                                fc="#fdf1ec", ec=NARANJA, lw=1.2, ls="--"))
    ax.text(9.0, 3.85, "B4  Módulo de búsqueda", ha="center", fontsize=8, fontweight="bold", color=TINTA)
    for i, (t, s) in enumerate([("BFS", "frontera FIFO"), ("UCS", "orden por g(n)"), ("A*", "orden por g(n)+h(n)")]):
        caja(ax, 9.0, 3.1 - i * 0.95, 2.1, 0.7, f"{t}\n{s}", color="white", borde=NARANJA, fs=7.5)
    caja(ax, 5.9, 1.0, 1.7, 0.8, "B5\nMétricas y\ncomparativa", negrita=True)
    caja(ax, 3.2, 1.0, 1.8, 0.8, "B6\nInforme", negrita=True)
    flecha(ax, (1.8, 3.6), (2.7, 3.6), "pacientes,\nmatriz")
    flecha(ax, (4.5, 3.6), (5.45, 3.6), "8 costos\nC(e, p)")
    flecha(ax, (7.15, 3.6), (7.75, 3.6), "G_p")
    flecha(ax, (7.75, 1.0), (6.75, 1.0), "ruta, costo,\nnodos (×3)")
    flecha(ax, (5.05, 1.0), (4.1, 1.0), "tabla de\n36 filas")
    ax.text(1.0, 2.55, "Se repite para\ncada uno de los\n12 pacientes", ha="center", fontsize=7, color=GRIS,
            style="italic")
    fig.tight_layout(); fig.savefig(ruta); plt.close(fig)


POS = {"Diagnostico": (0.6, 0), "Post_Op_Mayor": (3.0, 1.35), "Post_Op_Menor": (3.0, -1.35),
       "Consolidacion": (5.4, 0), "Alta": (7.8, 0)}


def fig_grafo(ruta, m):
    fig, ax = plt.subplots(figsize=(7.2, 3.9), dpi=220)
    ax.set_xlim(-0.4, 8.8); ax.set_ylim(-2.4, 3.0); ax.axis("off")
    r = 0.55
    for n, (x, y) in POS.items():
        col = "#e8f5ef" if n == "Alta" else ("#eaf2fc" if n == "Diagnostico" else "white")
        ax.add_patch(plt.Circle((x, y), r, fc=col, ec="#2e5578", lw=1.1, zorder=3))
        etiqueta = n.replace("_", "_\n", 1) if n.startswith("Post") else n
        ax.text(x, y, etiqueta + ("\n(inicio)" if n == "Diagnostico" else "\n(meta)" if n == "Alta" else ""),
                ha="center", va="center", fontsize=6.8, zorder=4)

    def borde(a, b):
        (x1, y1), (x2, y2) = POS[a], POS[b]
        L = math.hypot(x2 - x1, y2 - y1)
        ux, uy = (x2 - x1) / L, (y2 - y1) / L
        return (x1 + ux * r, y1 + uy * r), (x2 - ux * r, y2 - uy * r)

    aristas = [("e1", "Diagnostico", "Post_Op_Mayor", 0, (1.55, 0.95)),
               ("e2", "Diagnostico", "Post_Op_Menor", 0, (1.55, -0.95)),
               ("e3", "Post_Op_Mayor", "Consolidacion", 0.28, (4.55, 1.2)),
               ("e5", "Post_Op_Mayor", "Consolidacion", -0.28, (3.95, 0.45)),
               ("e4", "Post_Op_Menor", "Consolidacion", -0.28, (4.55, -1.2)),
               ("e6", "Post_Op_Menor", "Consolidacion", 0.28, (3.95, -0.45)),
               ("e7", "Consolidacion", "Alta", 0, (6.6, 0.12)),
               ("e8", "Diagnostico", "Alta", -0.5, (4.2, 2.72))]
    nombres = {"e1": "Cirugia_Abierta", "e2": "Cirugia_Laparos", "e3": "Quimio_Estandar",
               "e4": "Quimio_Estandar", "e5": "Inmunoterapia", "e6": "Inmunoterapia",
               "e7": "Radioterapia", "e8": "Cuidado_Paliativo"}
    for eid, a, b, rad, (lx, ly) in aristas:
        p1, p2 = borde(a, b)
        if eid == "e8":
            p1, p2 = (POS[a][0], POS[a][1] + r), (POS[b][0], POS[b][1] + r)
        flecha(ax, p1, p2, rad=rad, col="#444")
        f = m[m["Tratamiento"] == nombres[eid]].iloc[0]
        ax.text(lx, ly, f"{eid}: {nombres[eid]}\n${n0(f['Costo_Financiero_Base'])} · efect. "
                        f"{int(f['Tasa_Efectividad_Base']*100)} %",
                ha="center", va="center", fontsize=5.8, color=TINTA,
                bbox=dict(boxstyle="round,pad=0.15", fc="white", ec="#d0cfca", lw=0.5), zorder=5)
    ax.text(4.2, -2.3, "Peso de cada arista:  w(e) = 1 en BFS;   w(e) = C(e, p) en UCS y A*  (valores en la Tabla 8)",
            ha="center", fontsize=6.8, color=GRIS)
    fig.tight_layout(); fig.savefig(ruta); plt.close(fig)


def fig_barrido(df, ruta):
    fig, axs = plt.subplots(1, 3, figsize=(7.2, 2.5), dpi=220, sharey=True)
    series = [("Paliativo", AZUL, "o"), ("Curativa con quimio", NARANJA, "s"), ("Curativa con inmuno", AQUA, "^")]
    for ax, par in zip(axs, ["α", "β", "γ"]):
        s = df[df["Parámetro"] == par]
        for nom, col, mk in series:
            ax.plot(s["Factor k"], s[nom], color=col, lw=1.6, marker=mk, ms=3.5, label=nom)
        ax.axvline(1, color="#999", ls="--", lw=0.7)
        ax.set_title(f"{par}′ = k · {par}", fontsize=8, color=TINTA)
        ax.set_xlabel("factor k", fontsize=7, color=GRIS)
        ax.set_ylim(-0.5, 12.8)
        for sp in ["top", "right"]:
            ax.spines[sp].set_visible(False)
        ax.tick_params(colors=GRIS, labelsize=6.5)
        ax.grid(axis="y", color="#e6e5e0", lw=0.5)
    axs[0].set_ylabel("Pacientes", fontsize=7.5, color=GRIS)
    axs[1].legend(fontsize=6.3, loc="upper center", bbox_to_anchor=(0.5, -0.33), ncol=3, frameon=False)
    fig.tight_layout(); fig.savefig(ruta); plt.close(fig)


# ================= documento =================
doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Inches(8.5), Inches(11)
for k in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
    setattr(sec, k, Inches(1))
st = doc.styles["Normal"]
st.font.name, st.font.size = FUENTE, Pt(12)
st.element.rPr.rFonts.set(qn("w:eastAsia"), FUENTE)
st.paragraph_format.space_after = Pt(6)
st.paragraph_format.line_spacing = 1.15
for nom, tam in (("Heading 1", 14), ("Heading 2", 12)):
    h = doc.styles[nom]
    h.font.name, h.font.size, h.font.bold = FUENTE, Pt(tam), True
    h.font.color.rgb = RGBColor(0, 0, 0)
    rf = h.element.rPr.find(qn("w:rFonts"))
    for a in list(rf.attrib):
        del rf.attrib[a]
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf.set(qn(a), FUENTE)
    h.paragraph_format.space_before = Pt(12 if nom == "Heading 1" else 8)
    h.paragraph_format.space_after = Pt(4)
    h.paragraph_format.keep_with_next = True

# número de página
pp = sec.footer.paragraphs[0]
pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = pp.add_run()
for tipo, txt in (("begin", None), (None, "PAGE"), ("end", None)):
    if tipo:
        fc = OxmlElement("w:fldChar"); fc.set(qn("w:fldCharType"), tipo); r._r.append(fc)
    else:
        it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = txt; r._r.append(it)
r.font.size = Pt(10)


def par(texto="", estilo=None, alin=None, tam=None):
    p = doc.add_paragraph(style=estilo)
    for i, trozo in enumerate(texto.split("**")):
        for k, sub in enumerate(trozo.split("__")):
            if sub:
                run = p.add_run(sub)
                run.bold, run.italic = i % 2 == 1, k % 2 == 1
                if tam:
                    run.font.size = Pt(tam)
    if alin == "c":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return p


def vineta(texto):
    p = par(texto, "List Bullet")
    p.paragraph_format.space_after = Pt(3)
    return p


def formula(texto):
    p = par(texto, alin="c")
    p.paragraph_format.space_before = Pt(4)
    return p


def h1(t):
    doc.add_heading(t, level=1)


def h2(t):
    doc.add_heading(t, level=2)


N_TAB, N_FIG = [0], [0]


AZUL_OSC, ZEBRA, LINEA = "1F3B57", "F2F5F8", "BFC9D4"


def borde_tabla(t):
    """Bordes finos grises y márgenes internos amplios."""
    tblPr = t._tbl.tblPr
    bs = OxmlElement("w:tblBorders")
    for lado in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = OxmlElement(f"w:{lado}")
        b.set(qn("w:val"), "single"); b.set(qn("w:sz"), "4"); b.set(qn("w:color"), LINEA)
        bs.append(b)
    tblPr.append(bs)
    mar = OxmlElement("w:tblCellMar")
    for lado, v in (("top", 70), ("bottom", 70), ("left", 110), ("right", 110)):
        x = OxmlElement(f"w:{lado}"); x.set(qn("w:w"), str(v)); x.set(qn("w:type"), "dxa"); mar.append(x)
    tblPr.append(mar)


def sombra(celda, color):
    sh = OxmlElement("w:shd")
    sh.set(qn("w:val"), "clear"); sh.set(qn("w:color"), "auto"); sh.set(qn("w:fill"), color)
    celda._tc.get_or_add_tcPr().append(sh)


def tabla(encab, filas, anchos, titulo, nota=None, tam=10.5):
    N_TAB[0] += 1
    c = par(f"**Tabla {N_TAB[0]}.** __{titulo}__", tam=11)
    c.paragraph_format.keep_with_next = True
    c.paragraph_format.space_before = Pt(6)
    c.paragraph_format.space_after = Pt(4)
    t = doc.add_table(rows=1, cols=len(encab))
    t.autofit = False
    t.alignment = 1
    borde_tabla(t)
    for fila_i, vals in enumerate([encab] + filas):
        cells = t.rows[0].cells if fila_i == 0 else t.add_row().cells
        for j, v in enumerate(vals):
            cells[j].width = Inches(anchos[j])
            cells[j].vertical_alignment = 1
            p = cells[j].paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.1
            for i, trozo in enumerate(str(v).split("**")):
                if trozo:
                    run = p.add_run(trozo)
                    run.font.size = Pt(tam)
                    run.bold = fila_i == 0 or i % 2 == 1
                    if fila_i == 0:
                        run.font.color.rgb = RGBColor(255, 255, 255)
            if fila_i == 0:
                sombra(cells[j], AZUL_OSC)
            elif fila_i % 2 == 0:
                sombra(cells[j], ZEBRA)
    for j, w in enumerate(anchos):
        t.columns[j].width = Inches(w)
    for k, row in enumerate(t.rows):
        row.height = Inches(0.32 if k == 0 else 0.28)
        trPr = row._tr.get_or_add_trPr()
        cs = OxmlElement("w:cantSplit"); trPr.append(cs)
    th = OxmlElement("w:tblHeader"); t.rows[0]._tr.get_or_add_trPr().append(th)
    if nota:  # tabla y nota en la misma página
        for row in t.rows:
            for cel in row.cells:
                for q in cel.paragraphs:
                    q.paragraph_format.keep_with_next = True
    if nota:
        p = par(f"__Nota.__ {nota}", tam=10)
    else:
        p = par("", tam=4)
    p.paragraph_format.space_before = Pt(4)
    return t


def figura(ruta, titulo, ancho=6.5, nota=None):
    N_FIG[0] += 1
    c = par(f"**Figura {N_FIG[0]}.** __{titulo}__", tam=10)
    c.paragraph_format.keep_with_next = True
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(ruta, width=Inches(ancho))
    p.paragraph_format.keep_with_next = bool(nota)
    if nota:
        par(f"__Nota.__ {nota}", tam=9.5)


# ================= datos =================
P, m = lectura_datos(ruta_archivo)
RG = ab.rangos(P, m)
PA = ab.participacion(P, m)
BA = ab.barrido(P, m)
PE = ab.pesos_aristas(P, m)
DE = ab.desglose(P, m)
MD, MR = ab.metricas(P, m, rep=500)
n_enun = ab.pesos_enunciado(P, m)
os.makedirs(dir_res, exist_ok=True)
f_bl, f_gr, f_ba = [os.path.join(dir_res, x) for x in ("fig_bloques.png", "fig_grafo_pesos.png", "fig_barrido_pesos.png")]
fig_bloques(f_bl); fig_grafo(f_gr, m); fig_barrido(BA, f_ba)
a_min, b_min, g_min = gf.PESOS_MINIMIZAR_COSTO
a_max, b_max, g_max = gf.PESOS_MAXIMIZAR_SUPERVIVENCIA

# ================= contenido =================
par(f"**{AUTORES}**", tam=13)
t = par("**Sistema de Soporte a Decisiones Clínicas (SSDC): bloques del proyecto y grafo general**", tam=14)

# ---------- 1. Bloques ----------
h1("1. Bloques del proyecto")
par("El sistema tiene seis bloques. Los bloques B2 a B5 se repiten para cada uno de los 12 pacientes. "
    "BFS, UCS y A* no son tres bloques separados: forman un solo "
    "**módulo de búsqueda (B4)** con un mismo procedimiento, y solo cambia la regla para ordenar la frontera.")
figura(f_bl, "Diagrama de bloques con lo que pasa de un bloque al siguiente", 6.5)
tabla(["Bloque", "Recibe", "Qué hace", "Devuelve"], [
    ["B1. Lectura de datos", "Ruta del archivo datos_clinicos.xlsx",
     "Lee con pandas las pestañas Pacientes y Matriz_Tratamientos. No hace cálculos.",
     "Tabla de pacientes (12 filas) y matriz de tratamientos (8 filas, una por arista)"],
    ["B2. Costo unificado", "Una fila de paciente y la matriz de tratamientos",
     "Elige (α, β, γ) según la preferencia; para cada tratamiento calcula Π (multiplicadores activos), "
     "el castigo −ln(efectividad) y C(e, p). Detalle en 1.1.",
     "Lista de 8 costos C(e, p), en el mismo orden de las filas de la matriz"],
    ["B3. Construcción del grafo", "La matriz y la lista de 8 costos de B2",
     "Crea un multigrafo dirigido (NetworkX): fases como nodos y tratamientos como aristas. "
     "No calcula costos: copia cada C(e, p) en su arista. Detalle en 1.4.",
     "Grafo G_p del paciente: 5 nodos y 8 aristas con atributos tratamiento, costo y efectividad"],
    ["B4. Módulo de búsqueda (BFS, UCS, A*)", "G_p, inicio = Diagnostico, meta = Alta y la severidad del "
     "paciente (para h)", "Ejecuta el mismo ciclo de búsqueda con tres reglas de frontera. Detalle en 1.5.",
     "Por cada algoritmo: ruta, costo unificado y nodos expandidos"],
    ["B5. Métricas y comparativa", "Las 36 salidas de B4 (12 pacientes × 3 algoritmos) y la matriz",
     "Calcula las métricas de la sección 6, mide el tiempo y consolida.",
     "Tabla ID_Paciente | Algoritmo | Ruta | Costo | Nodos (consola, CSV y Excel)"],
    ["B6. Informe", "La tabla de B5 y las simulaciones",
     "Analiza por qué BFS falla, cómo UCS equilibra el riesgo, qué aporta la heurística y la parte ética.",
     "Informe en PDF"],
], [1.15, 1.5, 2.35, 1.5], "Entrada, proceso y salida de cada bloque")

h2("1.1 Bloque B2: cómo se calcula el costo de cada arista")
par("Para un paciente p y un tratamiento (arista) e, B2 calcula:")
formula("C(e, p) = α(p) · Costo_Base(e)  +  β(p) · Π(e, p)  +  γ(p) · [ −ln Efectividad(e) ]")
vineta("**Término económico, α·plata.** Costo_Financiero_Base del tratamiento en dólares.")
vineta("**Término de riesgo, β·Π.** Π(e, p) = ∏ mᵢ(e)^fᵢ(p) con i ∈ {Renal, Cardio, Edad, Inmuno}: se "
       "multiplican solo los multiplicadores de los factores que el paciente tiene (fᵢ = 1). Sin factores, Π = 1.")
vineta("**Término de efectividad, γ·castigo.** −ln(efectividad) vale 0 si el tratamiento siempre funciona y "
       "crece cuando la efectividad baja. El logaritmo convierte el producto de efectividades de una ruta en "
       "una suma, que es lo que acumulan UCS y A*.")
par(f"Los pesos dependen de la preferencia del paciente: minimizar costo → (α, β, γ) = ({d(a_min,1)}; "
    f"{n0(b_min)}; {n0(g_min)}); maximizar supervivencia → ({d(a_max,1)}; {n0(b_max)}; {n0(g_max)}). "
    f"El enunciado sugería α = 1,5 o 0,2 y β = 5.000 sin γ; con esos pesos los **{n_enun} pacientes** "
    f"reciben Cuidado_Paliativo, porque esa arista es la más barata en dólares y no tiene riesgo. Por eso "
    f"agregamos γ.")

h2("1.2 Rango de influencia de cada término")
par("Para saber cuánto pesa cada término, calculamos su valor mínimo y máximo sobre las 8 aristas y los "
    "12 pacientes con los pesos de cada preferencia.")
filas = []
for pref in ["Minimizar costo", "Maximizar supervivencia"]:
    s = RG[RG["Preferencia"] == pref]
    filas.append([pref] + [f"{n0(r['Mínimo'])} – {n0(r['Máximo'])}" for _, r in s.iterrows()])
tabla(["Preferencia", "α · plata", "β · Π", "γ · castigo"], filas, [1.9, 1.5, 1.5, 1.6],
      "Rango de cada término de C(e, p) en los datos del proyecto",
      "Variables de entrada: plata entre $1.500 y $25.000; Π entre 1 (sin factores) y 18 (Quimio_Estandar con "
      "los cuatro factores); castigo entre 0,105 (efectividad 90 %) y 1,609 (efectividad 20 %).")
cur = PA[PA["Ruta"] != "Cuidado_Paliativo"]
g = cur.groupby("Preferencia")[["% plata", "% riesgo", "% castigo"]].agg(["min", "max"])
fil = []
for pref, nom in [("Min. costo", "Minimizar costo"), ("Max. superv.", "Maximizar supervivencia")]:
    fil.append([nom] + [f"{d(g.loc[pref, (c, 'min')], 1)} – {d(g.loc[pref, (c, 'max')], 1)} %"
                        for c in ["% plata", "% riesgo", "% castigo"]])
p9 = PA[PA["ID"] == "P009"].iloc[0]
fil.append(["P009 (paliativo)", f"{d(p9['% plata'],1)} %", f"{d(p9['% riesgo'],1)} %", f"{d(p9['% castigo'],1)} %"])
tabla(["Rutas elegidas por UCS", "α · plata", "β · Π", "γ · castigo"], fil, [1.9, 1.5, 1.5, 1.6],
      "Participación de cada término en el costo de la ruta que elige UCS")
par("Lectura: el término de efectividad es el que más pesa (entre 56 % y 65 % en las rutas curativas), el "
    "riesgo pesa entre 11 % y 30 % según los factores del paciente, y la plata pesa cerca del 10 % al "
    "maximizar supervivencia y cerca del 28 % al minimizar costo. En pacientes con muchos factores el riesgo "
    "de una sola arista puede ser enorme: en P004 la quimioterapia aporta 54.000 por riesgo, más que todo el "
    "costo de la ruta que termina eligiendo (42.118).")

h2("1.3 Exploración de los pesos α, β y γ")
par("Para ver en qué rango la influencia de cada peso cambia las decisiones, multiplicamos un peso a la "
    "vez por un factor k (de 0 a 3) sobre el valor que tiene cada paciente según su preferencia, dejamos "
    "los otros dos fijos y contamos qué tipo de ruta propone UCS. k = 1 es el modelo actual.")
figura(f_ba, "Tipo de ruta que propone UCS al multiplicar un peso por el factor k", 6.5,
       "La línea punteada marca k = 1 (pesos del modelo). Datos completos en resultados/analisis_bloques.xlsx.")


def rango_estable(par_):
    """Intervalo continuo de k alrededor de 1 en que cambia la ruta de a lo sumo un paciente."""
    s_ = BA[BA["Parámetro"] == par_].reset_index(drop=True)
    i1 = s_.index[s_["Factor k"] == 1][0]
    lo = hi = i1
    while lo > 0 and s_.loc[lo - 1, "Cambian vs. base"] <= 1:
        lo -= 1
    while hi < len(s_) - 1 and s_.loc[hi + 1, "Cambian vs. base"] <= 1:
        hi += 1
    return s_.loc[lo, "Factor k"], s_.loc[hi, "Factor k"]


def fmt(v, par_):
    return d(v, 2).rstrip("0").rstrip(",") if par_ == "α" else n0(v)


EXTREMOS = {
    "α": "k ≤ 0,25: la plata casi no pesa y los 12 reciben inmunoterapia, la terapia más efectiva y segura. "
         "k ≥ 1,25: los 4 pacientes que minimizan costo pasan a paliativo; los que maximizan supervivencia no "
         "cambian porque su α es 10 veces menor.",
    "β": "k = 0: el riesgo no cuenta y 4 pacientes reciben quimioterapia aunque sea tóxica para ellos. "
         "k ≥ 2: el riesgo de cualquier cirugía o terapia pesa tanto que 6 a 10 pacientes pasan a paliativo, "
         "la única ruta con riesgo 1.",
    "γ": "k ≤ 0,25: los 12 reciben paliativo (la trampa de la fórmula sin γ). k ≥ 1,5: nadie recibe "
         "paliativo, ni siquiera P009, y la efectividad pasa a dominar la decisión.",
}
fil = []
pesos_txt = {"α": (a_min, a_max), "β": (b_min, b_max), "γ": (g_min, g_max)}
for par_ in ["α", "β", "γ"]:
    lo, hi = rango_estable(par_)
    vmin, vmax = pesos_txt[par_]
    fil.append([par_, f"{d(lo, 2).rstrip('0').rstrip(',')} – {d(hi, 2).rstrip('0').rstrip(',')}",
                f"Min. costo: {fmt(lo * vmin, par_)} – {fmt(hi * vmin, par_)}\nMax. superv.: "
                f"{fmt(lo * vmax, par_)} – {fmt(hi * vmax, par_)}", EXTREMOS[par_]])
tabla(["Peso", "k estable", "Valores equivalentes", "Qué pasa fuera de ese rango"], fil, [0.5, 0.75, 1.75, 3.5],
      "Sensibilidad de las rutas a cada peso",
      "k estable: intervalo alrededor de k = 1 en que cambia la ruta de a lo sumo un paciente respecto al modelo.")
par("Conclusión: el modelo no depende de un valor exacto de cada peso, pero cada uno tolera cambios en una "
    "sola dirección. α puede bajar a la mitad, pero si sube más de un 10 % los pacientes que minimizan costo "
    "pasan a paliativo. β puede bajar un 25 % o subir un 50 %. γ puede subir un 50 %, pero si baja más de un "
    "10 % aparecen rutas paliativas en pacientes que hoy reciben tratamiento curativo.")

h2("1.4 Bloque B3: cómo llegan los costos al grafo")
par("B3 no calcula costos. Recibe de B2 la lista de 8 valores C(e, p), en el mismo orden de las filas de "
    "la matriz, y guarda cada uno en su arista junto con el tratamiento y la efectividad. Por eso todos los "
    "pacientes tienen las mismas aristas pero con pesos distintos. Ejemplo con P004 (severidad 5, los cuatro "
    "factores, maximiza supervivencia: α = 0,1; β = 3.000; γ = 30.000):")
tabla(["Arista", "Tramo", "Π", "α·plata", "β·Π", "γ·castigo", "C(e, P004)"],
      [[r["Tratamiento"], r["Tramo"].replace("Post_Op_", "PO_").replace("Consolidacion", "Consolid.")
        .replace("Diagnostico", "Diagnóst."), d(r["Π"], 3), n0(r["α·plata"]), n0(r["β·Π"]), n0(r["γ·castigo"]),
        f"**{n0(r['C'])}**"] for _, r in DE.iterrows()],
      [1.35, 1.75, 0.6, 0.65, 0.65, 0.75, 0.75], "Costo de cada arista para el paciente P004")

h2("1.5 Bloque B4: un solo módulo de búsqueda")
par("Los tres algoritmos siguen el mismo ciclo; solo cambia cómo se ordena la frontera:")
for i, t_ in enumerate(["Poner el nodo inicial (Diagnostico) en la frontera.",
                        "Sacar el primer nodo de la frontera según la regla del algoritmo y contarlo como expandido.",
                        "Si es Alta, devolver la ruta, su costo y los nodos expandidos.",
                        "Si no, agregar a la frontera sus sucesores con el costo acumulado g(n′) = g(n) + C(e, p).",
                        "Repetir desde el paso 2."], 1):
    p = par(f"{i}. {t_}")
    p.paragraph_format.left_indent = Inches(0.3)
    p.paragraph_format.space_after = Pt(2)
tabla(["Algoritmo", "Frontera", "Se expande primero", "Peso de arista", "Prueba de meta"],
      [["BFS (Nivel 1)", "Cola FIFO", "El de menor profundidad d(n)", "w(e) = 1", "Al extraer"],
       ["UCS (Nivel 2)", "Cola de prioridad (heapq)", "El de menor g(n)", "w(e) = C(e, p)", "Al extraer"],
       ["A* (Nivel 3)", "Cola de prioridad (heapq)", "El de menor f(n) = g(n) + h(n)", "w(e) = C(e, p)",
        "Al extraer"]],
      [1.0, 1.4, 1.75, 1.1, 1.25], "Reglas de cada algoritmo dentro del módulo de búsqueda")

# ---------- 2. Espacio de estados ----------
h1("2. Espacio de estados")
par("Conjunto de estados: S = {Diagnostico, Post_Op_Mayor, Post_Op_Menor, Consolidacion, Alta}, con |S| = 5.")
par("Estado inicial: s₀ = Diagnostico.     Test de meta: Meta(s) = verdadero ⇔ s = Alta.")
par("Acciones (tratamientos): A = {Cirugia_Abierta, Cirugia_Laparos, Quimio_Estandar, Inmunoterapia, "
    "Radioterapia, Cuidado_Paliativo}.")
par("Acciones aplicables: Acciones(s) = { a ∈ A : existe una fila de la matriz con Fase_Origen = s y "
    "Tratamiento = a }.")
tabla(["Estado s", "Acciones(s)"],
      [["Diagnostico", "Cirugia_Abierta, Cirugia_Laparos, Cuidado_Paliativo"],
       ["Post_Op_Mayor", "Quimio_Estandar, Inmunoterapia"], ["Post_Op_Menor", "Quimio_Estandar, Inmunoterapia"],
       ["Consolidacion", "Radioterapia"], ["Alta", "Ninguna (estado terminal)"]],
      [1.8, 4.7], "Acciones aplicables en cada estado")
par("Función de transición: Resultado(s, a) = s′, donde s′ es la Fase_Destino de la fila (s, a). Por "
    "ejemplo, Resultado(Diagnostico, Cirugia_Abierta) = Post_Op_Mayor y Resultado(Diagnostico, "
    "Cuidado_Paliativo) = Alta.")

# ---------- 3. Grafo ----------
h1("3. Grafo general y pesos de las aristas")
par("El grafo se escribe directamente sobre los estados: **G = (S, E)**, con E = { (s, a, s′) : s′ = "
    "Resultado(s, a) } y |E| = 8. Cada nodo del grafo es uno de los estados de S, así que no hace falta "
    "definir un conjunto de vértices aparte.")
par("Es un multigrafo dirigido y acíclico: entre Post_Op_Mayor y Consolidacion (y entre Post_Op_Menor y "
    "Consolidacion) hay dos aristas paralelas, Quimio_Estandar e Inmunoterapia, y cada arista se identifica "
    "por (origen, tratamiento, destino). La Figura 3 muestra las 8 aristas, incluidas la de "
    "Cuidado_Paliativo y las paralelas.")
figura(f_gr, "Grafo general con las 8 aristas y sus datos base", 6.5)
tabla(["Arista", "Tratamiento", "USD", "Efect.", "w BFS", "C P001", "C P004", "C P009"],
      [[f"e{i}", r["Tratamiento"], n0(r["USD"]), f"{int(r['Efectividad']*100)} %", "1",
        n0(r["C P001"]), n0(r["C P004"]), n0(r["C P009"])]
       for i, (_, r) in enumerate(PE.iterrows(), 1)],
      [0.55, 1.5, 0.7, 0.6, 0.6, 0.85, 0.85, 0.85], "Peso de cada arista en BFS y en UCS/A* para tres pacientes",
      "P001: sin factores, maximiza supervivencia. P004: los cuatro factores, maximiza supervivencia. P009: "
      "renal e inmunosupresión, minimiza costo. Orden de aristas: e3/e4 quimio desde PO Mayor/Menor y e5/e6 "
      "inmunoterapia desde PO Mayor/Menor.")
par("Costo de una ruta: g(ruta) = Σₖ C(eₖ, p), con g(s₀) = 0. Lo que BFS minimiza es el número de aristas; "
    "UCS y A* minimizan g.")

# ---------- 4. Heurística ----------
h1("4. Heurística de A*")
par("Primero definimos la heurística y después mostramos que cumple las condiciones de admisibilidad y consistencia.")
h2("4.1 Definición")
formula("h(Alta) = 0          h(n) = ( Severidad_Diagnostico / 5 ) · c_min(n)   para n ≠ Alta")
par("donde c_min(n) es el costo unificado de la arista más barata que sale de n, con los pesos del paciente. "
    "Idea: desde cualquier fase distinta del Alta falta al menos un tratamiento, y el más barato cuesta "
    "c_min(n). La severidad (1 a 5) escala esa cota: con severidad 5 el factor es 1 y A* se enfoca más; con "
    "severidad 1 el factor es 0,2 y A* se comporta casi como UCS.")
p4 = P[P["ID_Paciente"] == "P004"].iloc[0]
G4 = ab.grafo_con(p4, m, gf.obtener_pesos(p4))
h4 = construir_heuristica(G4, 5)
hr = sim.h_real(G4)
tabla(["Fase n", "c_min(n)", "h(n) = 1,0 · c_min(n)", "h*(n) costo real"],
      [[n, n0(h4[n]) if n != "Alta" else "—", n0(h4[n]), n0(hr[n])]
       for n in ["Diagnostico", "Post_Op_Mayor", "Post_Op_Menor", "Consolidacion", "Alta"]],
      [1.8, 1.3, 1.7, 1.7], "Heurística y costo real hasta Alta para P004 (severidad 5)")
h2("4.2 Admisibilidad")
par("Una heurística es admisible si h(n) ≤ h*(n) para todo n, donde h*(n) es el costo real mínimo desde n "
    "hasta Alta. Desde n hay que pagar al menos una arista que sale de n, y ninguna cuesta menos que c_min(n); "
    "además 0,2 ≤ severidad/5 ≤ 1. Por tanto h(n) ≤ c_min(n) ≤ h*(n). En la Tabla 9 se ve en todas las fases.")
h2("4.3 Consistencia")
par("Para toda arista n → n′: como h(n′) ≥ 0 y C(n, a, n′) ≥ c_min(n), se cumple h(n) ≤ c_min(n) ≤ "
    "C(n, a, n′) ≤ C(n, a, n′) + h(n′), y h(Alta) = 0. Con una heurística consistente, el primer nodo que A* "
    "saca de la frontera ya tiene su costo mínimo, así que puede usar una lista cerrada sin perder la ruta "
    "óptima.")

# ---------- 5. Métricas ----------
h1("5. Métodos de medición para la comparativa")
par("Además de los nodos expandidos que pide el enunciado, B5 puede medir la calidad de la ruta y el costo "
    "computacional de la búsqueda:")
tabla(["Métrica", "Cómo se calcula", "Qué compara"], [
    ["Costo unificado", "g(ruta) = Σ C(eₖ, p)", "Qué tan buena es la ruta según el modelo"],
    ["Brecha de optimalidad", "(g − g*) / g*, con g* el costo de UCS", "Cuánto se aleja del óptimo"],
    ["Costo económico", "Σ Costo_Base(aₖ) en USD", "Gasto real del tratamiento"],
    ["Efectividad estimada", "∏ Efectividad(aₖ) (supuesto de independencia)", "Probabilidad de éxito clínico"],
    ["Riesgo máximo", "max Π(eₖ, p) sobre las aristas de la ruta", "El paso más peligroso para el paciente"],
    ["Longitud", "Número de aristas de la ruta", "Pasos de tratamiento"],
    ["Nodos expandidos", "Nodos que salen de la frontera", "Trabajo de la búsqueda"],
    ["Nodos generados", "Nodos que entran a la frontera", "Trabajo incluyendo repetidos"],
    ["Frontera máxima", "Mayor tamaño de la frontera", "Memoria usada"],
    ["Ramificación efectiva b*", "N + 1 = 1 + b* + … + (b*)^d, N = generados, d = longitud", "Qué tan dirigida es la búsqueda"],
    ["Tiempo de cómputo", "Promedio de 500 ejecuciones con time.perf_counter", "Velocidad"],
    ["Concordancia", "Pacientes en que la ruta coincide con la de UCS", "Si el algoritmo llega al óptimo"],
], [1.5, 2.8, 2.2], "Métricas propuestas para comparar los algoritmos")
r_ = MR.set_index("Algoritmo")
cols = [("Brecha %", "Brecha prom.", lambda v: f"{d(v,1)} %"), ("Efectividad", "Efect. prom.", lambda v: f"{d(v*100,1)} %"),
        ("Riesgo máx.", "Riesgo máx. prom.", lambda v: d(v, 2)), ("Expandidos", "Expandidos", lambda v: str(int(v))),
        ("Generados", "Generados", lambda v: str(int(v))), ("Frontera máx.", "Frontera máx. prom.", lambda v: d(v, 2)),
        ("b*", "b* prom.", lambda v: d(v, 2)), ("Tiempo (ms)", "Tiempo (ms)", lambda v: d(v, 3)),
        ("Coincide con UCS", "Coincide con UCS", lambda v: f"{int(v)}/12")]
tabla(["Métrica"] + ["BFS", "UCS", "A*"],
      [[nom] + [f(r_.loc[a, c]) for a in ["BFS", "UCS", "A*"]] for c, nom, f in cols],
      [2.0, 1.5, 1.5, 1.5], "Resultados de las métricas en los 12 pacientes",
      "Expandidos y generados son totales; el resto son promedios por paciente. Tiempos medidos en el "
      "entorno de prueba; en otro equipo cambian los valores absolutos, no la comparación.")
par("Lectura: BFS es el más rápido pero se aleja del óptimo un 21 % en promedio y propone rutas con 20 % de "
    "efectividad. UCS y A* llegan al mismo óptimo en los 12 casos. A* expande menos nodos (57 frente a 60) pero "
    "genera más (103 frente a 79): nuestra implementación no revisa si un nodo ya está en la frontera con "
    "mejor costo, así que mete duplicados. Esto muestra que medir solo nodos expandidos no basta para comparar.")

doc.save(salida)
print("Word guardado:", salida)
motor = exportar_pdf(salida, salida.replace(".docx", ".pdf"))
print(f"PDF guardado ({motor}):", salida.replace(".docx", ".pdf"))
