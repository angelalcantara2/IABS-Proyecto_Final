"""
armar_informe.py: llena la plantilla de los entregables 2 y 3 con los resultados
de reporte.py y simulaciones.py, la guarda en documentos/ (Word) y la exporta a PDF.

    python armar_informe.py

Para el PDF usa Microsoft Word (paquete docx2pdf) y, si no está, LibreOffice.
    pip install docx2pdf
"""
import os
import glob
import shutil
import subprocess
import copy
import datetime
import pandas as pd
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from docx.table import Table

from lector import lectura_datos, ruta_archivo
from reporte import generar_tabla, resumen_algoritmos, dir_res
import simulaciones as sim

raiz = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
plantilla = os.path.join(raiz, "docs", "SSDC_Plantilla_Entregables_2_y_3_actualizada.docx")
carpeta_salida = os.path.join(raiz, "documentos")
os.makedirs(carpeta_salida, exist_ok=True)
salida = os.path.join(carpeta_salida, "SSDC_Entregables_2_y_3.docx")
salida_pdf = salida.replace(".docx", ".pdf")

ASIGNATURA = "Búsqueda y Optimización de Soluciones (IABS)"
UNIVERSIDAD = ""   # opcional: se omite de la portada si está vacío
PROGRAMA = ""
DOCENTE = ""
FECHA = "22 de noviembre de 2026"
TITULO = "Sistema de Soporte a Decisiones Clínicas (SSDC): Reporte Comparativo y Análisis Crítico"
DECL_IA = ("El módulo de reporte (reporte.py), el script de simulaciones (simulaciones.py) y el "
           "borrador de este informe se elaboraron con apoyo de IA (Claude). El equipo ejecutó los "
           "scripts, verificó cada cifra contra la salida del programa y revisó y ajustó la redacción.")
INTEGRANTES = ["Juan Nicolas Alvarez Rodriguez", "Mateo Sanchez Garcia", "Angel Samuel Alcantara Acosta"]


# ---------- formato de números ----------
def n0(x):
    return f"{x:,.0f}".replace(",", ".")


def pct(x):
    return f"{x*100:.1f} %".replace(".", ",")


def dec(x, d=2):
    return f"{x:.{d}f}".replace(".", ",")


def corta(ruta):
    """Abrevia nombres de tratamientos para las tablas."""
    return (ruta.replace("Cirugia_Abierta", "Cir. Abierta").replace("Cirugia_Laparos", "Cir. Laparos")
                .replace("Quimio_Estandar", "Quimio").replace("Inmunoterapia", "Inmuno")
                .replace("Radioterapia", "Radio").replace("Cuidado_Paliativo", "Cuidado Paliativo"))


# ---------- ayudas de documento ----------
doc = Document(plantilla)
pars = lambda: doc.paragraphs


def buscar(inicio):
    for p in doc.paragraphs:
        if p.text.startswith(inicio):
            return p
    raise ValueError(inicio)


def add_runs(p, texto, tam=None):
    """Agrega texto con segmentos **negrita** y __cursiva__."""
    for i, trozo in enumerate(texto.split("**")):
        for k, sub in enumerate(trozo.split("__")):
            if sub:
                r = p.add_run(sub)
                r.bold = i % 2 == 1
                r.italic = k % 2 == 1
                if tam:
                    r.font.size = Pt(tam)


def nuevo_par_despues(ancla, texto="", ref=None):
    """Inserta un párrafo Normal después de 'ancla' (Paragraph o Table)."""
    el = ancla._p if isinstance(ancla, Paragraph) else ancla._tbl
    p = OxmlElement("w:p")
    if ref is not None and ref._p.pPr is not None:
        p.append(copy.deepcopy(ref._p.pPr))
    el.addnext(p)
    par = Paragraph(p, ancla._parent)
    if not isinstance(ancla, Paragraph):
        par.paragraph_format.space_before = Pt(6)
    add_runs(par, texto)
    return par


def anchos(t, ws):
    """Cambia el ancho de las columnas de una tabla (DXA)."""
    for gc, w in zip(t._tbl.find(qn("w:tblGrid")).findall(qn("w:gridCol")), ws):
        gc.set(qn("w:w"), str(w))
    for tr in t._tbl.findall(qn("w:tr")):
        for tc, w in zip(tr.findall(qn("w:tc")), ws):
            tc.find(qn("w:tcPr")).find(qn("w:tcW")).set(qn("w:w"), str(w))


REF = buscar("[COMPLETAR]")  # párrafo Normal de referencia
TIT = {}    # títulos APA de las tablas: elemento tabla -> (título, nota)
FIGS = []   # párrafos de las figuras
T = list(doc.tables)  # tablas originales (antes de insertar clones)


def llenar(inicio, textos):
    """Reemplaza el [COMPLETAR] del párrafo y agrega los párrafos siguientes."""
    p = buscar(inicio)
    run = [r for r in p.runs if "[COMPLETAR]" in r.text][0]
    if p.text.strip() == "[COMPLETAR]":
        run.text = ""
        add_runs(p, textos[0])
    else:
        run.text = run.text.replace("[COMPLETAR]", "")
        add_runs(p, textos[0])
    ult = p
    for t in textos[1:]:
        ult = nuevo_par_despues(ult, t, REF)
    return ult


def imagen_despues(ancla, ruta, ancho, num, titulo, nota=None):
    """Figura en formato APA: número en negrita, título en cursiva, imagen y nota."""
    p1 = nuevo_par_despues(ancla, "", REF)
    p1.add_run(f"Figura {num}").bold = True
    p2 = nuevo_par_despues(p1, "", REF)
    p2.add_run(titulo).italic = True
    p3 = nuevo_par_despues(p2, "", REF)
    p3.add_run().add_picture(ruta, width=Inches(ancho))
    ult = p3
    if nota:
        ult = nuevo_par_despues(p3, "", REF)
        ult.add_run("Nota. ").italic = True
        ult.add_run(nota)
    FIGS.append((p1, p2, p3, ult if nota else None))
    return ult


def celda(c, texto, negrita=False):
    p = c.paragraphs[0]
    for r in list(p.runs):
        r._r.getparent().remove(r._r)
    r = p.add_run(str(texto))
    r.font.name = "Arial"
    r._r.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "Arial")
    r.font.size = Pt(9)
    r.bold = negrita


def llenar_tabla(t, filas, desde=1):
    for i, fila in enumerate(filas):
        for j, v in enumerate(fila):
            celda(t.rows[desde + i].cells[j], v)


def ajustar_filas(t, n):
    """Deja la tabla con encabezado + n filas de datos (clona la primera)."""
    filas = t._tbl.findall(qn("w:tr"))
    proto = filas[1]
    for f in filas[1:]:
        t._tbl.remove(f)
    for _ in range(n):
        t._tbl.append(copy.deepcopy(proto))


def clonar_tabla(base, ancla, encab, filas, anchos):
    """Copia una tabla de la plantilla (mismo nº de columnas) después de 'ancla'."""
    nueva = copy.deepcopy(base._tbl)
    el = ancla._p if isinstance(ancla, Paragraph) else ancla._tbl
    el.addnext(nueva)
    t = Table(nueva, ancla._parent)
    ajustar_filas(t, len(filas))
    for gc, w in zip(nueva.find(qn("w:tblGrid")).findall(qn("w:gridCol")), anchos):
        gc.set(qn("w:w"), str(w))
    for tr in nueva.findall(qn("w:tr")):
        for tc, w in zip(tr.findall(qn("w:tc")), anchos):
            tc.find(qn("w:tcPr")).find(qn("w:tcW")).set(qn("w:w"), str(w))
    for j, h in enumerate(encab):
        celda(t.rows[0].cells[j], h, negrita=True)
    llenar_tabla(t, filas)
    return t


def borrar(p):
    p._p.getparent().remove(p._p)


# ---------- datos (todo sale del Excel; nada fijo) ----------
import math
import grafo as gf
from grafo import construir_grafo

P, m = lectura_datos(ruta_archivo)
NP, NT = len(P), len(m)
tabla = generar_tabla(P, m)
res = resumen_algoritmos(tabla)
CS = sim.casos(P, m)
SIM, COMP, FOCO, PALS_UCS = CS["simple"], CS["complejo"], CS["foco"], CS["paliativo_ucs"]
PAL = sim.paliativos(m)
PALN = PAL[0] if PAL else "la opción paliativa"
S1, S2, S3 = sim.s1(P, m), sim.s2(P, m, [SIM, COMP]), sim.s3(P, m)
S4, S5 = sim.s4(P, m, foco=FOCO), sim.s5(P, m)
PISO, TECHO = 0.30, 3.0
S6 = sim.s6(P, m, piso=PISO, techo=TECHO)
RF = sim.rutas_p(P, m, FOCO)
gF = sim.umbral_gamma(P, m, FOCO)
sim.fig_s4(S4, os.path.join(dir_res, "fig_barrido_gamma.png"))
sim.fig_s5(S5, os.path.join(dir_res, "fig_heuristicas.png"))

G0 = construir_grafo(m, sim.costos_pac(P.iloc[0], m))
FASES = list(G0.nodes())
N_RUTAS = len(sim.todas_rutas(G0))
PARALELAS = sorted({(u, v) for u, v, k in G0.edges(keys=True) if G0.number_of_edges(u, v) > 1})


def pac(pid):
    return P[P["ID_Paciente"] == pid].iloc[0]


def perf(pid):
    return sim.perfil(pac(pid))


def lista(xs):
    xs = list(xs)
    if not xs:
        return "ninguno"
    return xs[0] if len(xs) == 1 else ", ".join(xs[:-1]) + " y " + xs[-1]


def de(n, N):
    """'los 12' si son todos; si no, '7 de 12'."""
    return f"los {N}" if n == N else f"{n} de {N}"


def pref_inf(pid):
    return "minimizar costo" if pac(pid)["Minimizar_Costo_Financiero"] == 1 else "maximizar supervivencia"


NOM_FAC = {"Insuficiencia_Renal": "renal", "Cardiopatia": "cardio", "Edad_Avanzada_Fragil": "edad",
           "Inmunosupresion": "inmuno"}


def pref(pid):
    return "minimiza costo" if pac(pid)["Minimizar_Costo_Financiero"] == 1 else "maximiza supervivencia"


def pesos_txt(t):
    a, b, g = t
    return f"{dec(a, 1)}; {n0(b)}; {n0(g)}"


PESOS_TXT = {str(gf.PESOS_MINIMIZAR_COSTO): pesos_txt(gf.PESOS_MINIMIZAR_COSTO),
             str(gf.PESOS_MAXIMIZAR_SUPERVIVENCIA): pesos_txt(gf.PESOS_MAXIMIZAR_SUPERVIVENCIA)}
fila = lambda pid, alg: tabla[(tabla["ID_Paciente"] == pid) & (tabla["Algoritmo"] == alg)].iloc[0]
s3 = lambda pid, pr: S3[(S3["ID"] == pid) & S3["Preferencia simulada"].str.startswith(pr)].iloc[0]
c2 = lambda pid, t: S2[(S2["ID"] == pid) & (S2["Tratamiento"] == t)].iloc[0]
ucs_tr = {pid: fila(pid, "UCS")["Ruta Propuesta"].split(" → ") for pid in P["ID_Paciente"]}
EF_PAL = m[m["Tratamiento"] == PALN]["Tasa_Efectividad_Base"].iloc[0] if PAL else None
efs_ucs = tabla[(tabla["Algoritmo"] == "UCS") & ~tabla["Ruta Propuesta"].isin(PAL)]["Efectividad Estimada"]
ids_max = [p["ID_Paciente"] for _, p in P.iterrows() if p["Minimizar_Costo_Financiero"] != 1]
pal_max = [pid for pid in ids_max if s3(pid, "Min")["Ruta UCS"] in PAL]
cur_max = [pid for pid in ids_max if pid not in pal_max]
CASOS = list(dict.fromkeys([SIM, COMP]))

# ---------- Parte A ----------
llenar("[COMPLETAR]", [
    f"El módulo **reporte.py** lee el Excel, recorre los {NP} pacientes y, para cada uno, construye su "
    "grafo con el costo unificado y ejecuta BFS, UCS y A* sobre el mismo grafo. Cada ejecución "
    "devuelve la ruta, el costo y los nodos expandidos. La tabla se **imprime en consola** agrupada "
    "por paciente y se **exporta** a resultados/reporte_comparativo.csv y a "
    "resultados/reporte_comparativo.xlsx (hojas Reporte y Resumen). Se ejecuta con "
    "«python reporte.py». La misma tabla está disponible en la interfaz gráfica (botón «Reporte "
    "comparativo» de main.py), que también permite exportarla a CSV."])

t0 = T[0]
TIT[t0._tbl] = ("Definición de las columnas del reporte comparativo", None)
celda(t0.rows[4].cells[1],
      "Ambos: suma de Costo_Financiero_Base de la ruta en dólares y, entre paréntesis, el costo "
      "unificado C con el que deciden UCS y A*.")

t1 = T[1]
filas_a3 = [[f["ID_Paciente"], f["Algoritmo"], corta(f["Ruta Propuesta"]),
             f"${n0(f['Costo Económico'])} (C {n0(f['Costo Unificado'])})", f["Nodos Expandidos"]]
            for _, f in tabla.iterrows()]
ajustar_filas(t1, len(filas_a3))
llenar_tabla(t1, filas_a3)
TIT[t1._tbl] = (f"Reporte comparativo de los {NP} pacientes con los tres algoritmos",
                "C = costo unificado (α·USD + β·riesgo + γ·castigo). Cir. = cirugía; Laparos = "
                "laparoscópica; Quimio = Quimio_Estandar; Inmuno = Inmunoterapia; Radio = Radioterapia.")
anchos(t1, [1300, 1050, 3410, 2300, 1300])
t2 = T[2]
TIT[t2._tbl] = ("Resumen por algoritmo", f"Costos promedio sobre los {NP} pacientes.")
r = res.set_index("Algoritmo")
llenar_tabla(t2, [[a, r.loc[a, "Rutas distintas"],
                   f"${n0(r.loc[a, 'Costo USD promedio'])} (C {n0(r.loc[a, 'Costo unificado promedio'])})",
                   f"{dec(r.loc[a, 'Nodos expandidos (prom.)'])} (total {r.loc[a, 'Nodos expandidos (total)']})"]
                  for a in ["BFS", "UCS", "A*"]])

# ---------- Parte B ----------
llenar("[COMPLETAR]", [
    "Un paciente oncológico pasa por varias fases (diagnóstico, cirugía, consolidación y alta) y en "
    "cada una el equipo médico elige un tratamiento. La mejor elección no es la misma para todos: "
    "una insuficiencia renal, una cardiopatía, la edad avanzada o la inmunosupresión cambian el "
    "riesgo de cada tratamiento, y el paciente puede preferir cuidar su dinero o maximizar su "
    "supervivencia.",
    "El SSDC convierte ese problema en una búsqueda en un grafo: cada paciente tiene su propio grafo, "
    "cuyas aristas pesan según su perfil, y el sistema busca la ruta de Diagnostico a Alta. Comparamos "
    f"tres niveles de búsqueda sobre los {NP} pacientes del Excel (BFS, UCS y A*) con tres criterios: "
    "la ruta propuesta, su costo y los nodos expandidos. Con simulaciones del propio pipeline "
    "analizamos además los riesgos éticos de usar la herramienta en un ambiente clínico real."])

txt_par = (f" Hay aristas paralelas entre {lista([f'{u} y {v}' for u, v in PARALELAS])}, así que usamos un "
           "MultiDiGraph de NetworkX." if PARALELAS else " Usamos un MultiDiGraph de NetworkX.")
llenar("[COMPLETAR]", [
    f"**Estados y grafo.** {len(FASES)} fases ({lista(FASES)}) y {NT} tratamientos como aristas "
    f"dirigidas.{txt_par} Hay {N_RUTAS} rutas posibles de Diagnostico a Alta"
    + (f", entre ellas {PALN}, que va directo de Diagnostico a Alta." if PAL else "."),
    "**Función de costo.** Cada arista pesa C = α·plata + β·riesgo + γ·castigo, donde plata es el "
    "Costo_Financiero_Base, riesgo es el producto de los multiplicadores de los factores que el "
    "paciente tiene (1 si no tiene ninguno) y castigo = −ln(Tasa_Efectividad_Base). Los pesos "
    f"dependen de la preferencia: **minimizar costo (α, β, γ) = ({pesos_txt(gf.PESOS_MINIMIZAR_COSTO)})** y "
    f"**maximizar supervivencia ({pesos_txt(gf.PESOS_MAXIMIZAR_SUPERVIVENCIA)})**.",
    "**Supuestos.** Los multiplicadores son independientes y se multiplican; la efectividad de una "
    "ruta es el producto de las efectividades (por eso usamos −ln, que vuelve suma un producto); el "
    "estado Alta es el mismo sin importar la ruta; y cada factor del paciente se empareja con una "
    "columna de multiplicador (renal, cardio, edad e inmuno). El detalle está en el documento "
    "«Bloques y grafo»."])

# 3. BFS
bfs_t = tabla[tabla["Algoritmo"] == "BFS"]
vc = bfs_t["Ruta Propuesta"].value_counts()
if len(vc) == 1:
    txt_bfs = f"BFS devolvió **{vc.index[0]} en los {NP} pacientes**"
else:
    txt_bfs = "BFS devolvió " + lista([f"**{rt}** en {n} pacientes" for rt, n in vc.items()])
nb = bfs_t["Nodos Expandidos"]
txt_nb = (f"con {nb.iloc[0]} nodos expandidos en cada caso" if nb.nunique() == 1
          else f"con entre {nb.min()} y {nb.max()} nodos expandidos") + f" ({nb.sum()} en total)"
ef_b = bfs_t["Efectividad Estimada"]
txt_ef = (f"La efectividad estimada de esa ruta es del {pct(ef_b.iloc[0])}" if ef_b.nunique() == 1
          else f"La efectividad estimada de sus rutas va de {pct(ef_b.min())} a {pct(ef_b.max())}")
txt_ef += (f", frente a {pct(efs_ucs.min())}–{pct(efs_ucs.max())} de las rutas curativas que propone UCS."
           if len(efs_ucs) else ".")
llenar("3.1 Resultado observado:", [
    f"{txt_bfs}, {txt_nb}. La ruta no cambia con el perfil porque BFS ignora los pesos: el perfil "
    f"clínico no entra en la decisión. {txt_ef}"])

t3 = T[3]
TIT[t3._tbl] = (f"Ruta de BFS frente a ruta de UCS en {len(CASOS)} pacientes",
                "Entre paréntesis, la efectividad estimada de la ruta.")
ajustar_filas(t3, len(CASOS))
llenar_tabla(t3, [[pid, perf(pid),
                   f"{corta(fila(pid, 'BFS')['Ruta Propuesta'])} ({pct(fila(pid, 'BFS')['Efectividad Estimada'])})",
                   f"{corta(fila(pid, 'UCS')['Ruta Propuesta'])} ({pct(fila(pid, 'UCS')['Efectividad Estimada'])})"]
                  for pid in CASOS])

vc1 = S1["BFS sin paliativo"].value_counts()
txt_s1 = (f"Elige {vc1.index[0]} para los {NP} pacientes, solo porque esas aristas aparecen primero en el Excel."
          if len(vc1) == 1 else "Elige " + lista([f"{rt} en {n} pacientes" for rt, n in vc1.items()]) +
          ", según el orden de las aristas en el Excel.")
ult = llenar("3.3 Explicación:", [
    "BFS minimiza el número de pasos, no el resultado clínico. "
    + (f"{PALN} es la ruta más corta, así que siempre gana. " if PAL else "") +
    "El problema de fondo es que el estado Alta es el mismo nodo cuando se llega con un tratamiento "
    "curativo que cuando se llega con cuidado paliativo: para BFS «llegar al alta» es éxito, aunque la "
    "probabilidad de curación sea mínima. En un paciente sano esto ya es inaceptable; en un paciente "
    "complejo es más grave porque sus diferencias son justamente las que BFS no ve.",
    f"**Simulación S1b (BFS sin la opción paliativa).** Quitamos {PALN} para ver qué elige BFS entre "
    f"las rutas curativas. {txt_s1} En los pacientes complejos eso puede ser peligroso:"])
S1c = S1.assign(nf=[sim.n_factores(pac(i)) for i in S1["ID"]])
sel = S1c[S1c["nf"] >= 2] if (S1c["nf"] >= 2).any() else S1c
sel = sel.sort_values(["Riesgo máx. BFS s/p", "nf"], ascending=False).head(6)
t_s1 = clonar_tabla(T[3], ult,
                    ["Paciente", "Perfil", "BFS sin paliativo (riesgo máx.)", "UCS sin paliativo (riesgo máx.)"],
                    [[r_["ID"], r_["Perfil"], f"{corta(r_['BFS sin paliativo'])} ({dec(r_['Riesgo máx. BFS s/p'])})",
                      f"{corta(r_['UCS sin paliativo'])} ({dec(r_['Riesgo máx. UCS s/p'], 3)})"]
                     for _, r_ in sel.iterrows()],
                    [1050, 2750, 2780, 2780])
TIT[t_s1._tbl] = ("Rutas de BFS y UCS sin la opción paliativa en los pacientes con más factores (simulación S1b)", None)


def peor_paso(pid, trats):
    """Tratamiento de mayor riesgo de una ruta y sus multiplicadores activos."""
    p = pac(pid)
    mejor = max(trats, key=lambda t: sim.riesgo(p, m[m["Tratamiento"] == t].iloc[0]))
    f = m[m["Tratamiento"] == mejor].iloc[0]
    partes = [f"{NOM_FAC[k]} {dec(f[c], 1)}" for k, c in sim.FACT.items() if p[k] == 1 and f[c] != 1]
    return mejor, sim.riesgo(p, f), " · ".join(partes)


rc = S1[S1["ID"] == COMP].iloc[0]
pp_t, pp_r, pp_m = peor_paso(COMP, rc["BFS sin paliativo"].split(" → "))
nuevo_par_despues(t_s1,
    f"Riesgo máx. = mayor producto de multiplicadores de un tratamiento de la ruta. En {COMP} "
    f"({perf(COMP)}), BFS propone {pp_t} con riesgo {dec(pp_r)} veces el basal"
    + (f" ({pp_m})" if pp_m else "") +
    f"; UCS propone una ruta cuyo peor paso tiene riesgo {dec(rc['Riesgo máx. UCS s/p'], 3)}. "
    "La decisión de BFS depende del orden de las filas del Excel, no del paciente.", REF)

# 4. UCS
t4 = T[4]
TIT[t4._tbl] = ("Ruta de UCS en el paciente con menos factores de riesgo y en el que tiene más", None)
ajustar_filas(t4, len(CASOS))
llenar_tabla(t4, [[pid, perf(pid), corta(fila(pid, "UCS")["Ruta Propuesta"]),
                  f"${n0(fila(pid, 'UCS')['Costo Económico'])} (C {n0(fila(pid, 'UCS')['Costo Unificado'])})"]
                 for pid in CASOS])
primeros = [t for t in m[m["Fase_Origen"] == "Diagnostico"]["Tratamiento"] if t not in PAL]


def opciones(pid):
    p = pac(pid)
    return ", ".join(f"{t} C = {n0(c2(pid, t)['Costo unificado'])} (riesgo {dec(c2(pid, t)['Riesgo (Π mult.)'], 3)})"
                     for t in primeros)


misma_pref = pref(SIM) == pref(COMP)
d_usd = fila(COMP, "UCS")["Costo Económico"] - fila(SIM, "UCS")["Costo Económico"]
ult = nuevo_par_despues(t4,
    (f"Los dos pacientes prefieren {pref_inf(SIM)}; solo cambian los factores. " if misma_pref else
     f"{SIM} {pref(SIM)} y {COMP} {pref(COMP)}. ") +
    f"Primer paso en {SIM}: {opciones(SIM)}; UCS elige {ucs_tr[SIM][0]}. En {COMP}: {opciones(COMP)}; "
    f"UCS elige {ucs_tr[COMP][0]}"
    + (f", aunque la ruta completa cueste ${n0(d_usd)} más (simulación S2)." if d_usd > 0 else " (simulación S2)."),
    REF)

parrafos_42 = []
for fac, col in sim.FACT.items():
    con = [pid for pid in P["ID_Paciente"] if pac(pid)[fac] == 1]
    peor = m.loc[m[col].idxmax()]
    if not con or peor[col] <= 1:
        continue
    reciben = [pid for pid in con if peor["Tratamiento"] in ucs_tr[pid]]
    nombre = {"Insuficiencia_Renal": "insuficiencia renal", "Cardiopatia": "cardiopatía",
              "Edad_Avanzada_Fragil": "edad avanzada o fragilidad", "Inmunosupresion": "inmunosupresión"}[fac]
    parrafos_42.append(
        f"**{peor['Tratamiento']} en {nombre}.** Es el tratamiento con el multiplicador más alto para ese "
        f"factor ({dec(peor[col], 1)}). De los {len(con)} pacientes con {nombre}, "
        + (f"ninguno lo recibe." if not reciben else
           f"{len(reciben)} lo recibe{'n' if len(reciben) > 1 else ''} ({lista(reciben)}).")
        + ((" En los demás, UCS" if reciben else " UCS") + " lo evita porque su término β·Π crece con cada "
           "factor que se suma." if len(reciben) < len(con) else ""))
riesgosos = []
for _, r_ in S3.iterrows():
    if "(real)" not in r_["Preferencia simulada"] or r_["Ruta UCS"] in PAL:
        continue
    rutas_pid = sim.rutas_p(P, m, r_["ID"])
    mas_segura = rutas_pid[~rutas_pid["Ruta"].isin(PAL)]["Riesgo máx."].min()
    if r_["Riesgo máx."] > mas_segura + 1e-9:
        riesgosos.append((r_["ID"], r_["Ruta UCS"], r_["Riesgo máx."], mas_segura))
if riesgosos:
    parrafos_42.append(
        "**Lo que UCS no evita.** Algunos pacientes reciben una ruta más riesgosa aunque existe una curativa "
        "más segura, porque su preferencia o el costo pesan más que la diferencia de riesgo: "
        + lista([f"{pid} ({corta(rt)}, riesgo máx. {dec(rg)} frente a {dec(ms)} de la más segura, {pref(pid)})"
                 for pid, rt, rg, ms in riesgosos]) + ".")
else:
    parrafos_42.append("**Lo que UCS no evita.** Con los datos actuales todos los pacientes con ruta curativa "
                       "reciben la de menor riesgo máximo disponible.")
llenar("4.2 Tratamientos que UCS evita y por qué:", parrafos_42)

g0 = S4[S4["gamma"] == 0].iloc[0]["Pacientes con paliativo"]
cas_pal = -math.log(EF_PAL) if EF_PAL else None
ef_cur = RF[~RF["Ruta"].isin(PAL)]["Efectividad"]
cas_cur = -math.log(ef_cur.max()) if len(ef_cur) else None
ps = pac(SIM)
r_sim0 = sim.ucs(construir_grafo(m, sim.costos_pac(ps, m, (*gf.obtener_pesos(ps)[:2], 0))))[0][1::2]
segs = []
for _, r_ in S4.iterrows():
    if not segs or segs[-1][2] != r_["Pacientes con paliativo"]:
        segs.append([r_["gamma"], r_["gamma"], r_["Pacientes con paliativo"]])
    else:
        segs[-1][1] = r_["gamma"]
txt_segs = "; ".join(f"{int(n)} con γ = {n0(a)}" if a == b else f"{int(n)} entre γ = {n0(a)} y {n0(b)}"
                     for a, b, n in segs)
ult = llenar("4.3 La trampa de la fórmula y la corrección con γ:", [
    "Sin γ, la fórmula solo suma dinero y riesgo, y la opción paliativa es la más barata en las dos "
    f"cosas. En el barrido de γ (simulación S4), con γ = 0 reciben {PALN} **{de(g0, NP)} pacientes**. "
    "El término γ·(−ln efectividad) castiga las rutas poco efectivas"
    + (f": la opción paliativa paga γ·{dec(cas_pal, 3)} y la curativa más efectiva γ·{dec(cas_cur, 3)}."
       if cas_pal is not None and cas_cur is not None else ".")
    + f" Antes y después en {SIM}: con γ = 0 UCS propone {sim.flecha(r_sim0)}; con los pesos del modelo "
      f"propone {sim.flecha(ucs_tr[SIM])}.",
    f"La figura muestra cómo cambia el número de pacientes con la opción paliativa: {txt_segs}. Con los "
    f"pesos actuales la recibe{'n' if len(PALS_UCS) != 1 else ''} {len(PALS_UCS)}. Esto deja ver que **la recomendación depende de un número que "
    "fijó el equipo**, no de un dato clínico; volvemos a esto en el análisis ético."])
imagen_despues(ult, os.path.join(dir_res, "fig_barrido_gamma.png"), 6.0, 1,
               "Pacientes a los que UCS propone la opción paliativa según el valor de γ",
               "α y β fijos según la preferencia de cada paciente. Las líneas verticales marcan el γ de "
               "cada preferencia en el modelo (simulación S4).")

if PALS_UCS:
    intro44 = (f"Con los pesos actuales {len(PALS_UCS)} paciente{'s' if len(PALS_UCS) > 1 else ''} "
               f"recibe{'n' if len(PALS_UCS) > 1 else ''} {PALN}: {lista(PALS_UCS)}. Tomamos **{FOCO}** "
               f"({perf(FOCO)}). Sus {len(RF)} rutas ordenadas por costo unificado son:")
else:
    intro44 = (f"Con los pesos actuales ningún paciente recibe {PALN}. El caso más cercano es **{FOCO}** "
               f"({perf(FOCO)}), donde la opción paliativa queda más cerca de la ruta elegida. Sus "
               f"{len(RF)} rutas ordenadas por costo unificado son:")
ult = llenar("4.4 Casos donde UCS aún propone paliativo:", [intro44])
t_r9 = clonar_tabla(T[6], ult, ["Ruta", "Costo unificado", "USD", "Efectividad", "Riesgo máx."],
                    [[corta(r_["Ruta"]), n0(r_["Costo unificado"]), f"${n0(r_['USD'])}", pct(r_["Efectividad"]),
                      dec(r_["Riesgo máx."], 1)] for _, r_ in RF.iterrows()],
                    [3760, 1500, 1300, 1400, 1400])
TIT[t_r9._tbl] = (f"Rutas posibles del paciente {FOCO} ordenadas por costo unificado",
                  "Riesgo máx. = mayor producto de multiplicadores de un tratamiento de la ruta.")
cur = RF[~RF["Ruta"].isin(PAL)]
barata, segura = cur.iloc[0], cur.sort_values(["Riesgo máx.", "Costo unificado"]).iloc[0]
opuesta = s3(FOCO, "Max" if pac(FOCO)["Minimizar_Costo_Financiero"] == 1 else "Min")
gap = abs(RF.iloc[1]["Costo unificado"] - RF.iloc[0]["Costo unificado"])
nuevo_par_despues(t_r9,
    f"La primera ruta gana por {n0(gap)} puntos. La curativa más barata es {barata['Ruta']} (riesgo máx. "
    f"{dec(barata['Riesgo máx.'], 1)}) y la más segura es {segura['Ruta']} (${n0(segura['USD'])}, riesgo máx. "
    f"{dec(segura['Riesgo máx.'], 1)}). "
    + (f"El paliativo no sale porque {FOCO} no pueda tratarse, sino de la combinación de su preferencia "
       f"({pref(FOCO)}) con sus factores de riesgo: " if FOCO in PALS_UCS else "")
    + (f"con γ ≥ {n0(gF)}" if gF else "subiendo γ")
    + f" o con la preferencia contraria, UCS propone {opuesta['Ruta UCS']}.", REF)

# 5. A*
llenar("5.1 Definición de h(n):", [
    "h(Alta) = 0 y, para cualquier otra fase, h(n) = (Severidad_Diagnostico / 5) · c_min(n), donde "
    "c_min(n) es el costo unificado de la arista más barata que sale de n (con α, β y γ del "
    "paciente). Desde una fase distinta del Alta todavía falta al menos un tratamiento, y el más "
    "barato cuesta c_min(n). La severidad escala esa cota: con severidad 5 el factor es 1 y la "
    "heurística es más informada; con severidad 1 es 0,2 y A* se parece a UCS. Esa escala es una "
    "decisión de diseño del equipo, no un dato del Excel."])
llenar("5.2 Admisibilidad y consistencia:", [
    "**Admisible.** Para llegar al Alta desde n hay que pagar al menos una arista que sale de n, y "
    "ninguna cuesta menos que c_min(n). Como 0,2 ≤ severidad/5 ≤ 1: h(n) ≤ c_min(n) ≤ h*(n). Como γ ya "
    "está dentro de c_min(n), la cota vale para el costo unificado completo. Con una heurística "
    "admisible, A* garantiza encontrar la ruta de menor costo.",
    "**Consistente.** Para toda arista u → v, como h(v) ≥ 0 y C(u, v) ≥ c_min(u): "
    "h(u) ≤ c_min(u) ≤ C(u, v) ≤ C(u, v) + h(v). Además h(Alta) = 0. Por eso, cuando A* saca un nodo "
    "de la cola ya tiene su menor costo y puede usar la lista cerrada sin perder la ruta óptima. Lo "
    f"verificamos en los {NP} pacientes calculando h*(n) con Dijkstra sobre el grafo invertido."])

redu = [pid for pid in P["ID_Paciente"] if fila(pid, "A*")["Nodos Expandidos"] < fila(pid, "UCS")["Nodos Expandidos"]]
misma = sum(fila(pid, "A*")["Ruta Propuesta"] == fila(pid, "UCS")["Ruta Propuesta"] for pid in P["ID_Paciente"])
t5 = T[5]
TIT[t5._tbl] = ("Nodos expandidos por UCS y A* en los pacientes donde A* reduce la búsqueda", None)
filas5 = []
for pid in redu:
    u, a_ = fila(pid, "UCS")["Nodos Expandidos"], fila(pid, "A*")["Nodos Expandidos"]
    filas5.append([pid, u, a_, f"{u - a_} ({dec(100 * (u - a_) / u, 0)} %)",
                   "Sí" if fila(pid, "UCS")["Ruta Propuesta"] == fila(pid, "A*")["Ruta Propuesta"] else "No"])
if not filas5:
    filas5 = [["—", "—", "—", "0", "—"]]
ajustar_filas(t5, len(filas5))
llenar_tabla(t5, filas5)
tu, ta = r.loc["UCS", "Nodos expandidos (total)"], r.loc["A*", "Nodos expandidos (total)"]
nuevo_par_despues(t5,
    (f"Son los {len(redu)} pacientes donde A* expande menos nodos; en los otros {NP - len(redu)} expande "
     "lo mismo que UCS. " if redu else "En ningún paciente A* expande menos nodos que UCS. ")
    + f"En total: UCS {tu} y A* {ta} nodos ({dec(100 * (tu - ta) / tu, 0)} % menos), con la misma ruta "
      f"en {de(misma, NP)} pacientes.", REF)

hh = {k: S5[k].sum() for k in ["h = 0 (UCS)", "h severidad (equipo)", "h = c_min (factor 1)", "h* (perfecta)"]}
ucs_t = tabla[tabla["Algoritmo"] == "UCS"]
nodos_ruta = (ucs_t["Ruta Propuesta"].str.count("→") + 2).mean()
ult = llenar("5.4 Interpretación:", [
    f"La reducción es pequeña por el tamaño del grafo ({len(FASES)} fases). La ruta óptima ya contiene en "
    f"promedio {dec(nodos_ruta, 1)} nodos y UCS expande en promedio {dec(ucs_t['Nodos Expandidos'].mean(), 1)}, "
    "así que el margen de ahorro es de muy pocos nodos por paciente. A* solo ahorra cuando h hace que la f de "
    "un nodo supere el costo óptimo"
    + (f", y eso ocurre en {lista([f'{pid} (severidad {pac(pid)['Severidad_Diagnostico']})' for pid in redu])}."
       if redu else "."),
    f"**Simulación S5 (cuatro heurísticas).** Con h = 0, A* es UCS ({hh['h = 0 (UCS)']} nodos). Con la "
    f"heurística del equipo expande {hh['h severidad (equipo)']} y con c_min sin escalar "
    f"{hh['h = c_min (factor 1)']}. Con la heurística perfecta h*, A* llega a {hh['h* (perfecta)']}, el mínimo "
    "posible; nuestra heurística queda lejos de ese mínimo porque solo mira un paso hacia adelante.",
    "En un grafo más grande, con más fases, más tratamientos por fase o varias líneas de "
    "tratamiento, UCS expande todo nodo con g(n) menor que el costo óptimo C*, y ese número crece de "
    "forma exponencial con la profundidad. A* solo expande los nodos con g(n) + h(n) < C*, así que "
    "cada punto de información de h recorta ramas completas, y valdría la pena una heurística que mire "
    "más de un paso."])
imagen_despues(ult, os.path.join(dir_res, "fig_heuristicas.png"), 6.0, 2,
               f"Nodos expandidos por A* en los {NP} pacientes según la heurística",
               "En naranja, la heurística del equipo. h* es la heurística perfecta, calculada con "
               "Dijkstra (simulación S5).")

# 6. Ética
llenar("6.1 Responsabilidad:", [
    "Nuestra postura es que la responsabilidad es **compartida pero no se diluye**: cada actor "
    "responde por lo que controla.",
    "**El médico tratante** responde por la decisión final. La herramienta propone; el médico "
    "acepta, ajusta o rechaza y obtiene el consentimiento informado. En Colombia la Ley 23 de 1981 "
    "(ética médica) asigna al médico el deber de cuidado sobre el paciente, y una recomendación "
    "automática no lo reemplaza.",
    "**La institución** responde por haber adoptado la herramienta: validarla con sus propios datos, "
    "capacitar a quien la usa, auditar sus recomendaciones y definir cuándo no debe usarse.",
    "**Los desarrolladores (nosotros)** respondemos por el modelo. "
    + (f"La simulación S4 lo muestra: {FOCO} cambia de recomendación cuando γ cruza {n0(gF)}, y ese valor "
       "lo elegimos nosotros sin evidencia clínica. " if gF else
       "La simulación S4 muestra que las rutas cambian según el valor de γ, que elegimos sin evidencia clínica. ")
    + "Elegir α, β y γ es tomar decisiones clínicas disfrazadas de parámetros, así que tenemos que "
      "documentarlas, justificarlas y advertir sus límites. Si el modelo falla por un error que pudimos "
      "prever (por ejemplo, un multiplicador mal emparejado), la falla es nuestra aunque el médico haya firmado."])

rf_s1 = S1[S1["ID"] == FOCO].iloc[0]
txt_62a = (f"**1. Abandono terapéutico.** Si los {len(ids_max)} pacientes que prefieren supervivencia pasaran a "
           f"«minimizar costo», **{de(len(pal_max), len(ids_max))}** recibirían {PALN}"
           + (f"; solo conserva{'n' if len(cur_max) > 1 else ''} una ruta curativa "
              + lista([f"{pid} ({sim.n_factores(pac(pid))} factores)" if sim.n_factores(pac(pid))
                       else f"{pid} (sin factores de riesgo)" for pid in cur_max]) + "."
              if cur_max else ".")
           + " Cada factor encarece las rutas curativas y la opción paliativa queda como la salida barata: "
             "la herramienta tiende a renunciar al tratamiento justamente con los pacientes más frágiles, "
             "una discriminación por comorbilidad que nadie programó a propósito."
           if ids_max else
           "**1. Abandono terapéutico.** En los datos actuales ningún paciente prefiere maximizar supervivencia, "
           "así que esta comparación no aplica.")
txt_62b = (f"**2. Tratamiento tóxico.** Si se quita la opción paliativa, la ruta más barata para {FOCO} es "
           f"{rf_s1['UCS sin paliativo']}, cuyo peor paso tiene riesgo {dec(rf_s1['Riesgo máx. UCS s/p'])} "
           f"veces el basal (simulación S1b)."
           + (" La preferencia económica empuja a una de dos salidas: no tratar o tratar con lo más tóxico."
              if rf_s1["Riesgo máx. UCS s/p"] >= 1.5 else ""))
llenar("6.2 Priorización económica y opciones letales:", [
    "La simulación S3 corre UCS sobre cada paciente con su preferencia real y con la contraria. "
    "Aparecen dos formas en que el peso del dinero puede llevar a un resultado clínicamente inaceptable:",
    txt_62a, txt_62b])
t6 = T[6]
TIT[t6._tbl] = ("Efecto de la preferencia del paciente sobre la ruta propuesta (simulación S3)",
                "(real) indica la preferencia registrada en el Excel.")
filas6 = []
for pid in dict.fromkeys([FOCO, COMP]):
    for pr in ("Min", "Max"):
        r_ = s3(pid, pr)
        filas6.append([pid, r_["Preferencia simulada"], PESOS_TXT[r_["Pesos (α, β, γ)"]],
                       f"{corta(r_['Ruta UCS'])} (efect. {pct(r_['Efectividad'])})",
                       f"${n0(r_['USD'])} (C {n0(r_['Costo unificado'])})"])
ajustar_filas(t6, len(filas6))
llenar_tabla(t6, filas6)
fm, fx = s3(FOCO, "Min"), s3(FOCO, "Max")
cm, cx = s3(COMP, "Min"), s3(COMP, "Max")
nuevo_par_despues(t6,
    f"Para {FOCO}, pasar de «minimizar costo» a «maximizar supervivencia» cambia la efectividad estimada de "
    f"{pct(fm['Efectividad'])} a {pct(fx['Efectividad'])} con una diferencia de ${n0(abs(fx['USD'] - fm['USD']))}. "
    + (f"Para {COMP}, con «maximizar supervivencia» la ruta es {cx['Ruta UCS']} y con «minimizar costo» "
       f"sería {cm['Ruta UCS']}. " if COMP != FOCO else "")
    + "El cuidado paliativo es una opción legítima cuando el paciente, informado, la elige; lo inaceptable "
      "es que salga de un peso económico sin que nadie se lo explique.", REF)

cambio6 = list(S6[S6["Cambió"] == "Sí"]["ID"])
txt_s6 = (f"Cambió la ruta de {lista([f'{pid} (a {S6[S6.ID == pid].iloc[0]['UCS con guardarraíl']})' for pid in cambio6])}; "
          f"los otros {NP - len(cambio6)} conservaron su ruta."
          if cambio6 else f"Con los datos actuales ninguna ruta cambió: los {NP} pacientes ya cumplen ambos límites.")
llenar("6.3 Recomendaciones de uso en un ambiente clínico real:", [
    "**Apoyo, no decisión.** La salida se presenta como sugerencia y siempre requiere la firma del "
    "médico tratante.",
    f"**Mostrar todas las rutas, no solo la óptima.** En este grafo son {N_RUTAS}: el médico y el paciente "
    f"deben ver cada una con su efectividad, su riesgo y su costo en dólares, como en la tabla de {FOCO}.",
    "**Guardarraíles duros que la preferencia económica no pueda superar.** Probamos esto en la "
    f"simulación S6: descartamos toda ruta con un tratamiento de riesgo mayor a {dec(TECHO, 1)} y toda ruta "
    f"con efectividad menor al {int(PISO * 100)} % si existe una alternativa que cumpla ambos límites. {txt_s6}",
    "**Paliativo con confirmación explícita.** Cuando la ruta propuesta es la opción paliativa y "
    "existe una curativa viable, el sistema debe pedirlo como decisión consciente y registrarla.",
    "**Validar y auditar los pesos.** α, β, γ y la escala de la heurística deben fijarse con "
    "especialistas y datos reales, y revisarse periódicamente comparando las recomendaciones con los "
    "desenlaces.",
    "**Proteger los datos.** Los datos del paciente están anonimizados en este proyecto; en uso real "
    "aplica la Ley 1581 de 2012 de protección de datos personales."])

n_bfs_pal = int(bfs_t["Ruta Propuesta"].isin(PAL).sum())
max_bfs_r = S1["Riesgo máx. BFS s/p"].max()
llenar("[COMPLETAR]", [
    f"**BFS no sirve como apoyo clínico.** Propone la opción paliativa a {de(n_bfs_pal, NP)} pacientes "
    "porque cuenta pasos, y sin esa opción su elección depende del orden del Excel, con riesgos de hasta "
    f"{dec(max_bfs_r, 1)} veces el basal.",
    f"**UCS personaliza el tratamiento.** Con la función de costo unificada encuentra "
    f"{r.loc['UCS', 'Rutas distintas']} rutas distintas según el perfil y evita los tratamientos de mayor "
    "riesgo para cada factor. El término γ es indispensable: sin él UCS repite el error de BFS.",
    f"**A* conserva la ruta óptima y reduce la búsqueda.** Su heurística es admisible y consistente, "
    f"da la misma ruta que UCS en {de(misma, NP)} casos y ahorra {tu - ta} de {tu} nodos. El ahorro es "
    "pequeño por el tamaño del grafo, no por la heurística.",
    "**El riesgo ético está en los pesos.** "
    + (f"La preferencia económica lleva a la opción paliativa a {len(pal_max)} de {len(ids_max)} pacientes "
       "de supervivencia" if ids_max else "La preferencia económica cambia las rutas")
    + (f", y la recomendación de {FOCO} cambia con un valor de γ que eligió el equipo" if gF else "")
    + ". La herramienta solo es aceptable como apoyo, con todas las rutas visibles y con guardarraíles de "
      "riesgo y efectividad."])

llenar("[COMPLETAR]", [
    f"**Tamaño del grafo.** {len(FASES)} fases y {N_RUTAS} rutas no representan un tratamiento oncológico "
    "real (varias líneas de quimioterapia, recaídas, efectos acumulados), y limitan lo que podemos medir "
    "sobre la heurística.",
    "**Pesos sin justificación clínica.** α, β y γ los elegimos para que el modelo se comporte de forma "
    "razonable, no a partir de datos; la simulación S4 muestra cuánto dependen de ellos las rutas.",
    "**Emparejamiento de factores.** Supusimos que cada factor del paciente corresponde a una sola "
    "columna de multiplicador y que los riesgos se multiplican de forma independiente. En la clínica "
    "las comorbilidades interactúan.",
    "**Efectividad y estado Alta.** La efectividad de la ruta es el producto de efectividades "
    "independientes y el Alta no distingue curación de paliación. La severidad solo entra en la "
    "heurística, no en el costo.",
    f"**Datos.** Son {NP} pacientes sintéticos y una matriz de {NT} tratamientos; no hay validación con "
    "desenlaces reales."])

# sin sección de referencias
h = buscar("9. Referencias")
sig = h._p.getnext()
borrar(h)
if sig is not None and Paragraph(sig, h._parent).text.strip() == "[COMPLETAR]":
    sig.getparent().remove(sig)

# Anexo
t7 = T[7]
TIT[t7._tbl] = ("Simulaciones del pipeline usadas en el análisis", None)
ajustar_filas(t7, 7)
llenar_tabla(t7, [
    ["S1", "BFS sobre todos los pacientes (y S1b: BFS y UCS sin la opción paliativa)", "Nivel 1 (BFS)", "Todos"],
    ["S2", "UCS en el paciente con menos factores frente al de más factores", "Nivel 2 (UCS)", lista(CASOS)],
    ["S3", "Mismo paciente cambiando solo la preferencia", "Nivel 2 (UCS) y Análisis ético",
     f"Todos ({lista(list(dict.fromkeys([FOCO, COMP])))} en tabla)"],
    ["S4", f"Barrido de γ de 0 a {n0(S4['gamma'].max())} con α y β fijos; umbral de γ para {FOCO}",
     "Nivel 2 (UCS) y Análisis ético", "Todos"],
    ["S5", "A* con h = 0, h de severidad, h = c_min y h* perfecta", "Nivel 3 (A*)", "Todos"],
    ["S6", f"UCS con guardarraíl: riesgo ≤ {dec(TECHO, 1)} por tratamiento y efectividad ≥ {int(PISO * 100)} %",
     "Análisis ético", "Todos"],
    ["—", "Todas las rutas de un paciente con costo, efectividad y riesgo", "Nivel 2 (UCS)", FOCO],
])
nuevo_par_despues(t7,
    "Todas las simulaciones se reproducen con «python simulaciones.py»; los resultados quedan en "
    "resultados/simulaciones.xlsx (una hoja por simulación) y las figuras en resultados/. Los pacientes "
    "de ejemplo se eligen automáticamente: el de menos factores de riesgo, el de más factores y el primero "
    "al que UCS le propone la opción paliativa.", REF)


# ---------- limpieza: guías e instrucciones ----------
for p in list(doc.paragraphs):
    if p.text.startswith("Guía:") or p.text.startswith("Cómo usar esta plantilla"):
        borrar(p)



def buscar_soffice():
    """Ruta de LibreOffice (PATH o carpetas típicas de Windows)."""
    for nom in ["soffice", "libreoffice"]:
        if shutil.which(nom):
            return shutil.which(nom)
    for patron in [r"C:\Program Files\LibreOffice\program\soffice.exe",
                   r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"]:
        hallados = glob.glob(patron)
        if hallados:
            return hallados[0]
    return None


def exportar_pdf(docx, pdf):
    """Convierte el .docx a PDF: primero con Word (docx2pdf), si no con LibreOffice."""
    try:
        from docx2pdf import convert
        convert(docx, pdf)
        return "Word"
    except Exception as e:
        err_word = e
    soffice = buscar_soffice()
    if soffice:
        subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir",
                        os.path.dirname(pdf), docx], check=True,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return "LibreOffice"
    raise RuntimeError(f"No se pudo crear el PDF. Instala docx2pdf (pip install docx2pdf) "
                       f"con Word, o LibreOffice. Detalle: {err_word}")


from apa import aplicar_apa
aplicar_apa(doc, TIT, FIGS, REF, dict(titulo=TITULO, autores=INTEGRANTES, universidad=UNIVERSIDAD,
                                      programa=PROGRAMA, curso=ASIGNATURA, docente=DOCENTE,
                                      fecha=FECHA, decl_ia=DECL_IA))
doc.save(salida)
print("Word guardado:", salida)
motor = exportar_pdf(salida, salida_pdf)
print(f"PDF guardado ({motor}):", salida_pdf)
