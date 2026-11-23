"""
armar_informe.py: llena la plantilla de los entregables 2 y 3 con los resultados
de reporte.py y simulaciones.py, la guarda en docs/ y la exporta a PDF.

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
salida = os.path.join(raiz, "docs", "SSDC_Entregables_2_y_3.docx")
salida_pdf = salida.replace(".docx", ".pdf")

ASIGNATURA = "Búsqueda y Optimización de Soluciones (IABS)"
UNIVERSIDAD = "[COMPLETAR universidad]"
PROGRAMA = "[COMPLETAR facultad o programa]"
DOCENTE = "[COMPLETAR nombre del docente]"
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


# ---------- datos ----------
P, m = lectura_datos(ruta_archivo)
tabla = generar_tabla(P, m)
res = resumen_algoritmos(tabla)
S1, S2, S3 = sim.s1(P, m), sim.s2(P, m), sim.s3(P, m)
S4, S5, S6 = sim.s4(P, m), sim.s5(P, m), sim.s6(P, m)
R9 = sim.rutas_p(P, m, "P009")
g9 = sim.umbral_gamma(P, m)
sim.fig_s4(S4, os.path.join(dir_res, "fig_barrido_gamma.png"))
sim.fig_s5(S5, os.path.join(dir_res, "fig_heuristicas.png"))

fila = lambda pid, alg: tabla[(tabla["ID_Paciente"] == pid) & (tabla["Algoritmo"] == alg)].iloc[0]
s3 = lambda pid, pref: S3[(S3["ID"] == pid) & S3["Preferencia simulada"].str.startswith(pref)].iloc[0]
pal_max = [r["ID"] for _, r in S3.iterrows()
           if r["Preferencia simulada"] == "Min. costo" and r["Ruta UCS"] == "Cuidado_Paliativo"]

# ---------- portada ----------



# ---------- Parte A ----------
llenar("[COMPLETAR]", [
    "El módulo **reporte.py** lee el Excel, recorre los 12 pacientes y, para cada uno, construye su "
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
filas_a3 = []
for _, f in tabla.iterrows():
    filas_a3.append([f["ID_Paciente"], f["Algoritmo"], corta(f["Ruta Propuesta"]),
                     f"${n0(f['Costo Económico'])} (C {n0(f['Costo Unificado'])})",
                     f["Nodos Expandidos"]])
llenar_tabla(t1, filas_a3)
TIT[t1._tbl] = ("Reporte comparativo de los 12 pacientes con los tres algoritmos",
                "C = costo unificado (α·USD + β·riesgo + γ·castigo). Cir. = cirugía; Laparos = "
                "laparoscópica; Quimio = Quimio_Estandar; Inmuno = Inmunoterapia; Radio = Radioterapia.")
anchos(t1, [1300, 1050, 3410, 2300, 1300])
t2 = T[2]
TIT[t2._tbl] = ("Resumen por algoritmo", "Costos promedio sobre los 12 pacientes.")
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
    "tres niveles de búsqueda sobre los 12 pacientes del Excel (BFS, UCS y A*) con tres criterios: "
    "la ruta propuesta, su costo y los nodos expandidos. Con simulaciones del propio pipeline "
    "analizamos además los riesgos éticos de usar la herramienta en un ambiente clínico real."])

llenar("[COMPLETAR]", [
    "**Estados y grafo.** Cinco fases (Diagnostico, Post_Op_Mayor, Post_Op_Menor, Consolidacion y "
    "Alta) y ocho tratamientos como aristas dirigidas. Como Quimio_Estandar e Inmunoterapia salen de "
    "las dos fases postoperatorias, usamos un MultiDiGraph de NetworkX (Hagberg et al., 2008). Hay cinco rutas posibles: "
    "cuatro curativas de tres pasos y Cuidado_Paliativo, que va directo de Diagnostico a Alta.",
    "**Función de costo.** Cada arista pesa C = α·plata + β·riesgo + γ·castigo, donde plata es el "
    "Costo_Financiero_Base, riesgo es el producto de los multiplicadores de los factores que el "
    "paciente tiene (1 si no tiene ninguno) y castigo = −ln(Tasa_Efectividad_Base). Los pesos "
    "dependen de la preferencia: **minimizar costo (α, β, γ) = (1,0; 2.000; 35.000)** y "
    "**maximizar supervivencia (0,1; 3.000; 30.000)**.",
    "**Supuestos.** Los multiplicadores son independientes y se multiplican; la efectividad de una "
    "ruta es el producto de las efectividades (por eso usamos −ln, que vuelve suma un producto); el "
    "estado Alta es el mismo sin importar la ruta; y cada factor del paciente se empareja con una "
    "columna de multiplicador (renal, cardio, edad e inmuno). El detalle está en el documento "
    "«Bloques y grafo»."])

# 3. BFS
p = llenar("3.1 Resultado observado:", [
    f"BFS devolvió **Cuidado_Paliativo en los 12 pacientes**, con 4 nodos expandidos en cada caso "
    f"(48 en total). La ruta no cambia entre pacientes porque BFS ignora los pesos: el perfil "
    f"clínico no entra en la decisión. La efectividad estimada de esa ruta es del 20 %, frente a "
    f"38–46 % de las rutas curativas que propone UCS."])

t3 = T[3]
TIT[t3._tbl] = ("Ruta de BFS frente a ruta de UCS en dos pacientes", "Entre paréntesis, la efectividad estimada de la ruta.")
perf = lambda pid: sim.perfil(P[P["ID_Paciente"] == pid].iloc[0])
llenar_tabla(t3, [[pid, perf(pid), "Cuidado Paliativo (20 %)",
                   f"{corta(fila(pid, 'UCS')['Ruta Propuesta'])} ({pct(fila(pid, 'UCS')['Efectividad Estimada'])})"]
                  for pid in ["P001", "P004"]])

b4 = S1[S1["ID"] == "P004"].iloc[0]
ult = llenar("3.3 Explicación:", [
    "BFS minimiza el número de pasos, no el resultado clínico (Russell y Norvig, 2021). Cuidado_Paliativo es la única ruta de "
    "un paso, así que siempre gana. El problema de fondo es que el estado Alta es el mismo nodo "
    "cuando se llega con un tratamiento curativo que cuando se llega con cuidado paliativo: para BFS "
    "«llegar al alta» es éxito, aunque la probabilidad de curación sea mínima. En un paciente sano "
    "esto ya es inaceptable; en un paciente complejo es más grave porque sus diferencias son "
    "justamente las que BFS no ve.",
    "**Simulación S1b (BFS sin la opción paliativa).** Quitamos la arista de Cuidado_Paliativo para "
    "ver qué elige BFS entre las rutas curativas. Elige Cirugia_Abierta → Quimio_Estandar → "
    "Radioterapia para los 12 pacientes, solo porque esas aristas aparecen primero en el Excel. En los "
    "pacientes complejos eso es peligroso:"])
t_s1 = clonar_tabla(T[3], ult,
                    ["Paciente", "Perfil", "BFS sin paliativo (riesgo máx.)", "UCS sin paliativo (riesgo máx.)"],
                    [[r["ID"], r["Perfil"], f"{corta(r['BFS sin paliativo'])} ({dec(r['Riesgo máx. BFS s/p'])})",
                      f"{corta(r['UCS sin paliativo'])} ({dec(r['Riesgo máx. UCS s/p'], 3)})"]
                     for _, r in S1[S1["ID"].isin(["P003", "P004", "P008", "P012"])].iterrows()],
                    [1050, 2750, 2780, 2780])
TIT[t_s1._tbl] = ("Rutas de BFS y UCS sin la opción paliativa en pacientes complejos (simulación S1b)", None)
nuevo_par_despues(t_s1,
    f"Riesgo máx. = mayor producto de multiplicadores de un tratamiento de la ruta. En P004 y P012 "
    f"(los cuatro factores), BFS propone una quimioterapia con riesgo {dec(b4['Riesgo máx. BFS s/p'], 0)} "
    f"veces el basal (renal 4,0 · cardio 1,5 · edad 1,5 · inmuno 2,0) y una cirugía abierta de 11,25; "
    f"UCS evita ambas y propone una ruta cuyo peor paso tiene riesgo {dec(b4['Riesgo máx. UCS s/p'], 3)}. "
    f"La decisión de BFS depende del orden de las filas del Excel, no del paciente.", REF)

# 4. UCS
t4 = T[4]
TIT[t4._tbl] = ("Ruta de UCS en un paciente sin factores de riesgo y en uno con los cuatro factores", None)
llenar_tabla(t4, [[pid, perf(pid), corta(fila(pid, "UCS")["Ruta Propuesta"]),
                   f"${n0(fila(pid, 'UCS')['Costo Económico'])} (C {n0(fila(pid, 'UCS')['Costo Unificado'])})"]
                  for pid in ["P001", "P004"]])
c = lambda pid, t, tramo=None: S2[(S2["ID"] == pid) & (S2["Tratamiento"] == t)].iloc[0]
ult = nuevo_par_despues(t4,
    f"Los dos pacientes quieren maximizar supervivencia; solo cambian los factores. En P001 la cirugía "
    f"abierta cuesta C = {n0(c('P001','Cirugia_Abierta')['Costo unificado'])} y la laparoscópica "
    f"{n0(c('P001','Cirugia_Laparos')['Costo unificado'])}, así que UCS elige la abierta. En P004 la "
    f"abierta sube a {n0(c('P004','Cirugia_Abierta')['Costo unificado'])} (riesgo 11,25 por cardio, "
    f"edad e inmuno) y la laparoscópica solo a {n0(c('P004','Cirugia_Laparos')['Costo unificado'])} "
    f"(riesgo 1,716): UCS cambia de cirugía aunque cueste $3.500 más (simulación S2).", REF)

llenar("4.2 Tratamientos que UCS evita y por qué:", [
    f"**Quimio_Estandar en insuficiencia renal.** Su multiplicador renal es 4,0 (el más alto de la "
    f"matriz). En P004 se suma la inmunosupresión (2,0), la cardiopatía y la edad (1,5 cada una): "
    f"riesgo 18 y C = {n0(c('P004','Quimio_Estandar')['Costo unificado'])}, frente a "
    f"{n0(c('P004','Inmunoterapia')['Costo unificado'])} de la inmunoterapia. Ningún paciente renal "
    f"con preferencia de supervivencia (P003, P004, P006, P012) recibe quimioterapia.",
    "**Cirugia_Abierta en cardiopatía o edad avanzada.** Sus multiplicadores son 2,5 y 3,0. Por eso "
    "P002, P003, P004, P008, P010 y P012 pasan a laparoscopía. P006, renal sin cardiopatía, mantiene "
    "la abierta porque su multiplicador renal es 1,0.",
    "**Lo que UCS no evita.** Los pacientes que minimizan costo sin factor renal (P005, P007, P011) "
    "reciben Quimio_Estandar: con α = 1 los $22.000 de diferencia frente a la inmunoterapia pesan "
    "más que su riesgo. P005, inmunosuprimido, acepta así un tratamiento con riesgo 2,0."])

g0 = S4[S4["gamma"] == 0].iloc[0]["Pacientes con paliativo"]
ult = llenar("4.3 La trampa de la fórmula y la corrección con γ:", [
    f"Sin γ, la fórmula solo suma dinero y riesgo, y Cuidado_Paliativo es la ruta más barata en las "
    f"dos cosas ($1.500 y riesgo 1). En el barrido de γ (simulación S4), con γ = 0 los "
    f"**{g0} pacientes** reciben paliativo y UCS coincide con BFS. El término γ·(−ln efectividad) "
    f"castiga las rutas poco efectivas: el paliativo paga γ·1,609 y la mejor ruta curativa γ·0,779. "
    f"Antes y después en P001: con γ = 0 UCS propone paliativo; con γ = 30.000 propone Cirugia_Abierta "
    f"→ Inmunoterapia → Radioterapia.",
    "La figura muestra que el número de pacientes con paliativo cae por escalones: 12 hasta "
    "γ = 10.000, 4 entre 20.000 y 27.500, 1 con los pesos actuales y 0 desde unos 52.500. Esto "
    "deja ver que **la recomendación depende de un número que fijó el equipo**, no de un dato "
    "clínico; volvemos a esto en el análisis ético."])
imagen_despues(ult, os.path.join(dir_res, "fig_barrido_gamma.png"), 6.0, 1,
               "Pacientes a los que UCS propone Cuidado_Paliativo según el valor de γ",
               "α y β fijos según la preferencia de cada paciente. Las líneas verticales marcan el γ de "
               "cada preferencia en el modelo (simulación S4).")

ult = llenar("4.4 Casos donde UCS aún propone paliativo:", [
    f"Con los pesos actuales solo **P009** (severidad 3, insuficiencia renal, inmunosupresión, "
    f"minimiza costo) recibe Cuidado_Paliativo. Sus cinco rutas ordenadas por costo unificado son:"])
t_r9 = clonar_tabla(T[6], ult,
                    ["Ruta", "Costo unificado", "USD", "Efectividad", "Riesgo máx."],
                    [[corta(r["Ruta"]), n0(r["Costo unificado"]), f"${n0(r['USD'])}", pct(r["Efectividad"]),
                      dec(r["Riesgo máx."], 1)] for _, r in R9.iterrows()],
                    [3760, 1500, 1300, 1400, 1400])
TIT[t_r9._tbl] = ("Rutas posibles del paciente P009 ordenadas por costo unificado", "Riesgo máx. = mayor "
                  "producto de multiplicadores de un tratamiento de la ruta.")
nuevo_par_despues(t_r9,
    f"El paliativo gana por {n0(R9.iloc[1]['Costo unificado'] - R9.iloc[0]['Costo unificado'])} "
    f"puntos. Las curativas baratas usan quimioterapia (riesgo 8 por renal e inmuno) y las seguras "
    f"usan inmunoterapia, que cuesta $25.000. Con α = 1, la diferencia en dólares supera el castigo "
    f"por baja efectividad. El paliativo no sale porque P009 no pueda tratarse, sino de la "
    f"combinación de su preferencia económica con el riesgo renal: con γ ≥ {n0(g9)} o con la "
    f"preferencia de supervivencia, UCS propone Cirugia_Abierta → Inmunoterapia → Radioterapia.", REF)

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
    "admisible, A* garantiza encontrar la ruta de menor costo (Hart et al., 1968).",
    "**Consistente.** Para toda arista u → v, como h(v) ≥ 0 y C(u, v) ≥ c_min(u): "
    "h(u) ≤ c_min(u) ≤ C(u, v) ≤ C(u, v) + h(v). Además h(Alta) = 0. Por eso, cuando A* saca un nodo "
    "de la cola ya tiene su menor costo y puede usar la lista cerrada sin perder la ruta óptima. Lo "
    "verificamos en los 12 pacientes calculando h*(n) con Dijkstra sobre el grafo invertido."])

t5 = T[5]
TIT[t5._tbl] = ("Nodos expandidos por UCS y A* en los pacientes donde A* reduce la búsqueda", None)
llenar_tabla(t5, [[pid, fila(pid, "UCS")["Nodos Expandidos"], fila(pid, "A*")["Nodos Expandidos"],
                   f"{fila(pid, 'UCS')['Nodos Expandidos'] - fila(pid, 'A*')['Nodos Expandidos']} (20 %)",
                   "Sí" if fila(pid, "UCS")["Ruta Propuesta"] == fila(pid, "A*")["Ruta Propuesta"] else "No"]
                  for pid in ["P004", "P009", "P012"]])
nuevo_par_despues(t5,
    "Son los tres pacientes donde A* ahorra un nodo; en los otros nueve expande 5, igual que UCS. "
    "En total: UCS 60 y A* 57 nodos (5 % menos), con la misma ruta y el mismo costo en los 12.", REF)

h = {k: S5[k].sum() for k in ["h = 0 (UCS)", "h severidad (equipo)", "h = c_min (factor 1)", "h* (perfecta)"]}
ult = llenar("5.4 Interpretación:", [
    "La reducción es pequeña por el tamaño del grafo. Una ruta curativa obliga a expandir 4 nodos "
    "(Diagnostico, una fase postoperatoria, Consolidacion y Alta), así que UCS, que expande 5, solo "
    "puede ahorrar uno: la otra fase postoperatoria. A* la evita cuando h hace que su f supere el "
    "costo óptimo, y eso solo ocurre en P004 y P012 (severidad 5 con los cuatro factores, donde h es alta) y en P009 (cuya ruta óptima es el paliativo de un paso).",
    f"**Simulación S5 (cuatro heurísticas).** Con h = 0, A* es UCS ({h['h = 0 (UCS)']} nodos). Con la "
    f"heurística del equipo ({h['h severidad (equipo)']}) obtenemos lo mismo que con c_min sin escalar "
    f"({h['h = c_min (factor 1)']}): en este grafo la escala por severidad no cuesta eficiencia. Con la "
    f"heurística perfecta h*, A* llega a {h['h* (perfecta)']}, el mínimo posible; nuestra heurística "
    f"queda lejos de ese mínimo porque solo mira un paso hacia adelante.",
    "En un grafo más grande, con más fases, más tratamientos por fase o varias líneas de "
    "tratamiento, UCS expande todo nodo con g(n) menor que el costo óptimo C*, y ese número crece de "
    "forma exponencial con la profundidad. A* solo expande los nodos con g(n) + h(n) < C*, así que "
    "cada punto de información de h recorta ramas completas (Russell y Norvig, 2021). Ahí la diferencia entre 5 y 4 nodos se "
    "volvería de cientos o miles, y valdría la pena una heurística que mire más de un paso."])
imagen_despues(ult, os.path.join(dir_res, "fig_heuristicas.png"), 6.0, 2,
               "Nodos expandidos por A* en los 12 pacientes según la heurística",
               "En naranja, la heurística del equipo. h* es la heurística perfecta, calculada con "
               "Dijkstra (simulación S5).")

# 6. Ética
llenar("6.1 Responsabilidad:", [
    "Nuestra postura es que la responsabilidad es **compartida pero no se diluye**: cada actor "
    "responde por lo que controla. La OMS insiste en que la IA en salud no debe desplazar la "
    "responsabilidad humana sobre las decisiones clínicas (Organización Mundial de la Salud [OMS], 2021).",
    "**El médico tratante** responde por la decisión final. La herramienta propone; el médico "
    "acepta, ajusta o rechaza y obtiene el consentimiento informado. En Colombia la Ley 23 de 1981 "
    "(Congreso de Colombia, 1981) asigna al médico el deber de cuidado sobre el paciente, y una recomendación "
    "automática no lo reemplaza.",
    "**La institución** responde por haber adoptado la herramienta: validarla con sus propios datos, "
    "capacitar a quien la usa, auditar sus recomendaciones y definir cuándo no debe usarse.",
    f"**Los desarrolladores (nosotros)** respondemos por el modelo. La simulación S4 lo muestra: P009 "
    f"pasa de paliativo a una ruta curativa cuando γ cruza {n0(g9)}, y ese valor lo elegimos nosotros "
    f"sin evidencia clínica. Elegir α, β y γ es tomar decisiones clínicas disfrazadas de parámetros, "
    f"así que tenemos que documentarlas, justificarlas y advertir sus límites. Si el modelo falla "
    f"por un error que pudimos prever (por ejemplo, un multiplicador mal emparejado), la falla es "
    f"nuestra aunque el médico haya firmado."])

PESOS_TXT = {"(1.0, 2000, 35000)": "1,0; 2.000; 35.000", "(0.1, 3000, 30000)": "0,1; 3.000; 30.000"}
p9m, p9s = s3("P009", "Min"), s3("P009", "Max")
p4m, p4s = s3("P004", "Min"), s3("P004", "Max")
llenar("6.2 Priorización económica y opciones letales:", [
    "La simulación S3 corre UCS sobre cada paciente con su preferencia real y con la contraria. "
    "Aparecen dos formas en que el peso del dinero puede llevar a un resultado clínicamente "
    "inaceptable:",
    f"**1. Abandono terapéutico.** Si los 8 pacientes que prefieren supervivencia pasaran a «minimizar "
    f"costo», **{len(pal_max)} de 8** recibirían Cuidado_Paliativo; solo P001, sin factores de riesgo, "
    f"conserva una ruta curativa. El sistema renuncia al tratamiento justamente con los pacientes más "
    f"frágiles: cada factor encarece las rutas curativas y el paliativo, con riesgo 1 en todas las "
    f"columnas, queda como la salida barata. Es una discriminación por comorbilidad que nadie programó "
    f"a propósito, y que va contra los principios de beneficencia y justicia (Beauchamp y "
    f"Childress, 2019).",
    "**2. Tratamiento tóxico.** Si se quita la opción paliativa, la ruta más barata para P009 es "
    "Cirugia_Abierta → Quimio_Estandar → Radioterapia, con una quimioterapia de riesgo 8 en un "
    "paciente renal e inmunosuprimido (simulación S1b). La preferencia económica empuja a una de dos "
    "salidas: no tratar o tratar con lo más tóxico."])
t6 = T[6]
TIT[t6._tbl] = ("Efecto de la preferencia del paciente sobre la ruta propuesta (simulación S3)",
                "(real) indica la preferencia registrada en el Excel.")
ajustar_filas(t6, 4)
llenar_tabla(t6, [[r["ID"], r["Preferencia simulada"], PESOS_TXT[r["Pesos (α, β, γ)"]],
                   f"{corta(r['Ruta UCS'])} (efect. {pct(r['Efectividad'])})",
                   f"${n0(r['USD'])} (C {n0(r['Costo unificado'])})"]
                  for r in [p9m, p9s, p4s, p4m]])
nuevo_par_despues(t6,
    f"Para P009, cambiar solo la preferencia pasa la efectividad estimada de {pct(p9m['Efectividad'])} a "
    f"{pct(p9s['Efectividad'])} a cambio de ${n0(p9s['USD'] - p9m['USD'])}. P004 hace el recorrido "
    f"contrario: con su preferencia real recibe laparoscopía e inmunoterapia y, si declarara «minimizar "
    f"costo», recibiría paliativo. El cuidado paliativo es una opción legítima cuando el paciente, "
    f"informado, la elige; lo inaceptable es que salga de un peso económico sin que nadie se lo "
    f"explique.", REF)

ult = llenar("6.3 Recomendaciones de uso en un ambiente clínico real:", [
    "**Apoyo, no decisión.** La salida se presenta como sugerencia y siempre requiere la firma del "
    "médico tratante.",
    "**Mostrar todas las rutas, no solo la óptima.** En este grafo son cinco: el médico y el paciente "
    "deben ver cada una con su efectividad, su riesgo y su costo en dólares, como en la tabla de P009.",
    "**Guardarraíles duros que la preferencia económica no pueda superar.** Probamos esto en la "
    "simulación S6: descartamos toda ruta con un tratamiento de riesgo mayor a 3,0 y toda ruta con "
    "efectividad menor al 30 % si existe una alternativa que cumpla ambos límites. Solo cambió P009, "
    "que pasó a Cirugia_Abierta → Inmunoterapia → Radioterapia; los otros 11 conservaron su ruta. El "
    "guardarraíl corrige el caso crítico sin alterar el resto.",
    "**Paliativo con confirmación explícita.** Cuando la ruta propuesta es Cuidado_Paliativo y "
    "existe una curativa viable, el sistema debe pedirlo como decisión consciente y registrarla.",
    "**Validar y auditar los pesos.** α, β, γ y la escala de la heurística deben fijarse con "
    "especialistas y datos reales, y revisarse periódicamente comparando las recomendaciones con los "
    "desenlaces.",
    "**Proteger los datos.** Los datos del paciente están anonimizados en este proyecto; en uso real "
    "aplica la Ley 1581 de 2012 de protección de datos personales (Congreso de Colombia, 2012)."])

llenar("[COMPLETAR]", [
    "**BFS no sirve como apoyo clínico.** Propone Cuidado_Paliativo a los 12 pacientes porque cuenta "
    "pasos, y sin esa opción su elección depende del orden del Excel, hasta llegar a riesgos de 18 "
    "veces el basal en los pacientes complejos.",
    "**UCS personaliza el tratamiento.** Con la función de costo unificada encuentra cuatro rutas "
    "distintas según el perfil, evita la quimioterapia en pacientes renales y la cirugía abierta en "
    "cardiópatas y adultos frágiles. El término γ es indispensable: sin él UCS repite el error de BFS.",
    "**A* conserva la ruta óptima y reduce la búsqueda.** Su heurística es admisible y consistente, "
    "da la misma ruta que UCS en los 12 casos y ahorra 3 de 60 nodos. El ahorro es pequeño por el "
    "tamaño del grafo, no por la heurística.",
    "**El riesgo ético está en los pesos.** La preferencia económica lleva a paliativo a 7 de 8 "
    "pacientes de supervivencia, y la recomendación de P009 cambia con un valor de γ que eligió el "
    "equipo. La herramienta solo es aceptable como apoyo, con todas las rutas visibles y con "
    "guardarraíles de riesgo y efectividad."])

llenar("[COMPLETAR]", [
    "**Tamaño del grafo.** Cinco fases y cinco rutas no representan un tratamiento oncológico real "
    "(varias líneas de quimioterapia, recaídas, efectos acumulados), y limitan lo que podemos medir "
    "sobre la heurística.",
    "**Pesos sin justificación clínica.** α, β y γ los elegimos para que el modelo se comporte de forma "
    "razonable, no a partir de datos; la simulación S4 muestra cuánto dependen de ellos las rutas.",
    "**Emparejamiento de factores.** Supusimos que cada factor del paciente corresponde a una sola "
    "columna de multiplicador y que los riesgos se multiplican de forma independiente. En la clínica "
    "las comorbilidades interactúan.",
    "**Efectividad y estado Alta.** La efectividad de la ruta es el producto de efectividades "
    "independientes y el Alta no distingue curación de paliación. La severidad solo entra en la "
    "heurística, no en el costo.",
    "**Datos.** Son 12 pacientes sintéticos y una matriz de 8 tratamientos; no hay validación con "
    "desenlaces reales."])

llenar("[COMPLETAR]", [
    "Beauchamp, T. L. y Childress, J. F. (2019). __Principles of biomedical ethics__ (8.ª ed.). Oxford "
    "University Press.",
    "Congreso de Colombia. (1981). __Ley 23 de 1981, por la cual se dictan normas en materia de ética "
    "médica__.",
    "Congreso de Colombia. (2012). __Ley Estatutaria 1581 de 2012, por la cual se dictan disposiciones "
    "generales para la protección de datos personales__.",
    "Hagberg, A. A., Schult, D. A. y Swart, P. J. (2008). Exploring network structure, dynamics, and "
    "function using NetworkX. En G. Varoquaux, T. Vaught y J. Millman (Eds.), __Proceedings of the 7th "
    "Python in Science Conference__ (pp. 11–15).",
    "Hart, P. E., Nilsson, N. J. y Raphael, B. (1968). A formal basis for the heuristic determination of "
    "minimum cost paths. __IEEE Transactions on Systems Science and Cybernetics, 4__(2), 100–107. "
    "https://doi.org/10.1109/TSSC.1968.300136",
    "Organización Mundial de la Salud. (2021). __Ethics and governance of artificial intelligence for "
    "health: WHO guidance__. https://www.who.int/publications/i/item/9789240029200",
    "Russell, S. y Norvig, P. (2021). __Artificial intelligence: A modern approach__ (4.ª ed.). Pearson."])

# Anexo
t7 = T[7]
TIT[t7._tbl] = ("Simulaciones del pipeline usadas en el análisis", None)
ajustar_filas(t7, 7)
llenar_tabla(t7, [
    ["S1", "BFS sobre los 12 pacientes (y S1b: BFS y UCS sin la arista paliativa)", "Nivel 1 (BFS)", "Todos"],
    ["S2", "UCS en un paciente sin factores frente a uno con varios factores", "Nivel 2 (UCS)", "P001 y P004"],
    ["S3", "Mismo paciente cambiando solo la preferencia", "Nivel 2 (UCS) y Análisis ético", "Todos (P009 y P004 en tabla)"],
    ["S4", "Barrido de γ de 0 a 60.000 con α y β fijos; umbral de γ para P009", "Nivel 2 (UCS) y Análisis ético", "Todos"],
    ["S5", "A* con h = 0, h de severidad, h = c_min y h* perfecta", "Nivel 3 (A*)", "Todos"],
    ["S6", "UCS con guardarraíl: riesgo ≤ 3,0 por tratamiento y efectividad ≥ 30 %", "Análisis ético", "Todos"],
    ["—", "Las cinco rutas de un paciente con costo, efectividad y riesgo", "Nivel 2 (UCS)", "P009 (y P004, P005 en Excel)"],
])
nuevo_par_despues(t7,
    "Todas las simulaciones se reproducen con «python simulaciones.py»; los resultados quedan en "
    "resultados/simulaciones.xlsx (una hoja por simulación) y las figuras en resultados/.", REF)

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
