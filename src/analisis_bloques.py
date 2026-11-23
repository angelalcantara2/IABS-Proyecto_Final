"""
analisis_bloques.py: cálculos de apoyo para el documento de bloques y grafo.

    python analisis_bloques.py

1. Rango de cada término de la función de costo (α·plata, β·riesgo, γ·castigo).
2. Peso de cada término dentro del costo de la ruta elegida por UCS.
3. Barrido de α, β y γ (uno a la vez, multiplicados por un factor k) y efecto sobre las rutas.
4. Pesos de las aristas para pacientes de ejemplo.
5. Métricas de comparación de BFS, UCS y A* (incluye frontera máxima y b*).

Guarda las tablas en resultados/analisis_bloques.xlsx.
"""
import os
import math
import time
from collections import deque
import heapq
import pandas as pd

from lector import lectura_datos, ruta_archivo
import grafo as gf
from grafo import construir_grafo
from busqueda import bfs, ucs, a_estrella, construir_heuristica
from reporte import dir_res
import simulaciones as sim

FACT = sim.FACT


def terminos(f, p, pesos):
    """(α·plata, β·riesgo, γ·castigo) de un tratamiento para un paciente."""
    a, b, g = pesos
    return (a * f["Costo_Financiero_Base"], b * sim.riesgo(p, f),
            g * -math.log(f["Tasa_Efectividad_Base"]))


def grafo_con(p, m, pesos):
    return construir_grafo(m, [sum(terminos(f, p, pesos)) for _, f in m.iterrows()])


def tipo_ruta(trats):
    if trats == ["Cuidado_Paliativo"]:
        return "Paliativo"
    return "Curativa con quimio" if "Quimio_Estandar" in trats else "Curativa con inmuno"


# ---------- 1. rangos de cada término ----------
def rangos(P, m):
    filas = []
    for nom, pesos in [("Minimizar costo", gf.PESOS_MINIMIZAR_COSTO),
                       ("Maximizar supervivencia", gf.PESOS_MAXIMIZAR_SUPERVIVENCIA)]:
        ts = [terminos(f, p, pesos) for _, p in P.iterrows() for _, f in m.iterrows()]
        for i, t in enumerate(["α·plata", "β·riesgo", "γ·castigo"]):
            vals = [x[i] for x in ts]
            filas.append({"Preferencia": nom, "Pesos": str(pesos), "Término": t,
                          "Mínimo": round(min(vals)), "Máximo": round(max(vals))})
    return pd.DataFrame(filas)


# ---------- 2. participación de cada término en la ruta de UCS ----------
def participacion(P, m):
    filas = []
    for _, p in P.iterrows():
        pesos = gf.obtener_pesos(p)
        ruta, c, _ = ucs(grafo_con(p, m, pesos))
        tot = [0, 0, 0]
        for t in ruta[1::2]:
            f = m[m["Tratamiento"] == t].iloc[0]
            tot = [x + y for x, y in zip(tot, terminos(f, p, pesos))]
        s = sum(tot)
        filas.append({"ID": p["ID_Paciente"],
                      "Preferencia": "Min. costo" if p["Minimizar_Costo_Financiero"] == 1 else "Max. superv.",
                      "Ruta": " → ".join(ruta[1::2]), "Costo": round(s),
                      "% plata": round(100 * tot[0] / s, 1), "% riesgo": round(100 * tot[1] / s, 1),
                      "% castigo": round(100 * tot[2] / s, 1)})
    return pd.DataFrame(filas)


# ---------- 3. barrido de α, β y γ ----------
FACTORES = [0, 0.25, 0.5, 0.75, 0.9, 1.0, 1.1, 1.25, 1.5, 2.0, 3.0]  # k: peso' = k · peso del paciente


def barrido(P, m):
    base = {p["ID_Paciente"]: ucs(grafo_con(p, m, gf.obtener_pesos(p)))[0][1::2] for _, p in P.iterrows()}
    filas = []
    for k, par in enumerate(["α", "β", "γ"]):
        for v in FACTORES:
            cuenta = {"Paliativo": 0, "Curativa con quimio": 0, "Curativa con inmuno": 0}
            cambian = 0
            for _, p in P.iterrows():
                pesos = list(gf.obtener_pesos(p))
                pesos[k] *= v
                t = ucs(grafo_con(p, m, tuple(pesos)))[0][1::2]
                cuenta[tipo_ruta(t)] += 1
                cambian += t != base[p["ID_Paciente"]]
            filas.append({"Parámetro": par, "Factor k": v, **cuenta, "Cambian vs. base": cambian})
    return pd.DataFrame(filas)


def pesos_enunciado(P, m):
    """Pesos sugeridos por el enunciado (sin γ): α = 1.5 o 0.2, β = 5000."""
    n = 0
    for _, p in P.iterrows():
        pesos = (1.5, 5000, 0) if p["Minimizar_Costo_Financiero"] == 1 else (0.2, 5000, 0)
        n += ucs(grafo_con(p, m, pesos))[0][1::2] == ["Cuidado_Paliativo"]
    return n


# ---------- 4. pesos de las aristas para pacientes de ejemplo ----------
def pesos_aristas(P, m, ids=("P001", "P004", "P009")):
    filas = []
    for _, f in m.iterrows():
        fila = {"Tratamiento": f["Tratamiento"], "Tramo": f"{f['Fase_Origen']} → {f['Fase_Destino']}",
                "USD": f["Costo_Financiero_Base"], "Efectividad": f["Tasa_Efectividad_Base"], "w BFS": 1}
        for pid in ids:
            p = P[P["ID_Paciente"] == pid].iloc[0]
            fila[f"C {pid}"] = round(sum(terminos(f, p, gf.obtener_pesos(p))))
        filas.append(fila)
    return pd.DataFrame(filas)


def desglose(P, m, pid="P004"):
    p = P[P["ID_Paciente"] == pid].iloc[0]
    pesos = gf.obtener_pesos(p)
    filas = []
    for _, f in m.iterrows():
        a, b, g = terminos(f, p, pesos)
        filas.append({"Tratamiento": f["Tratamiento"], "Tramo": f"{f['Fase_Origen']} → {f['Fase_Destino']}",
                      "Π": round(sim.riesgo(p, f), 3), "α·plata": round(a), "β·Π": round(b),
                      "γ·castigo": round(g), "C": round(a + b + g)})
    return pd.DataFrame(filas)


# ---------- 5. métricas de comparación ----------
def contar(G, alg, h=None, ini="Diagnostico", meta="Alta"):
    """Copia instrumentada de los algoritmos de busqueda.py: nodos generados y frontera máxima."""
    gen, fmax = 1, 1
    if alg == "BFS":
        q, vis = deque([ini]), {ini}
        while q:
            n = q.popleft()
            if n == meta:
                break
            for _, v, _k in G.out_edges(n, keys=True):
                if v not in vis:
                    vis.add(v); q.append(v); gen += 1
            fmax = max(fmax, len(q))
        return gen, fmax
    h = h or {n: 0 for n in G.nodes()}
    fr, best, closed = [(h[ini], 0, ini)], {ini: 0}, set()
    while fr:
        _, g, n = heapq.heappop(fr)
        if alg == "UCS" and g > best[n]:
            continue
        if alg == "A*":
            if n in closed:
                continue
            closed.add(n)
        if n == meta:
            break
        for _, v, d in G.out_edges(n, data=True):
            ng = g + d["costo"]
            if alg == "UCS" and v in best and ng >= best[v]:
                continue
            if alg == "A*" and v in closed:
                continue
            best[v] = ng
            heapq.heappush(fr, (ng + h[v], ng, v)); gen += 1
        fmax = max(fmax, len(fr))
    return gen, fmax


def b_efectivo(N, d):
    """b* tal que N + 1 = 1 + b + b² + … + b^d (Russell y Norvig)."""
    if d == 1:
        return float(N)
    lo, hi = 1.0, float(N)
    for _ in range(100):
        b = (lo + hi) / 2
        s = sum(b ** i for i in range(d + 1))
        lo, hi = (b, hi) if s < N + 1 else (lo, b)
    return b


def metricas(P, m, rep=2000):
    filas = []
    for _, p in P.iterrows():
        G = grafo_con(p, m, gf.obtener_pesos(p))
        h = construir_heuristica(G, p["Severidad_Diagnostico"])
        c_opt = ucs(G)[1]
        for alg, fn in [("BFS", lambda: bfs(G)), ("UCS", lambda: ucs(G)), ("A*", lambda: a_estrella(G, h))]:
            t0 = time.perf_counter()
            for _ in range(rep):
                ruta, c, nodos = fn()
            ms = (time.perf_counter() - t0) / rep * 1000
            trats = ruta[1::2]
            efec = math.prod(m[m["Tratamiento"] == t].iloc[0]["Tasa_Efectividad_Base"] for t in trats)
            gen, fmax = contar(G, alg, h if alg == "A*" else None)
            filas.append({"ID": p["ID_Paciente"], "Algoritmo": alg, "Pasos": len(trats),
                          "Costo unificado": c, "Brecha %": 100 * (c - c_opt) / c_opt,
                          "USD": sum(m[m["Tratamiento"] == t].iloc[0]["Costo_Financiero_Base"] for t in trats),
                          "Efectividad": efec, "Riesgo máx.": sim.riesgo_ruta(p, m, trats),
                          "Expandidos": nodos, "Generados": gen, "Frontera máx.": fmax,
                          "b*": b_efectivo(gen, len(trats)), "Tiempo (ms)": ms,
                          "Coincide con UCS": trats == ucs(G)[0][1::2]})
    df = pd.DataFrame(filas)
    res = df.groupby("Algoritmo", sort=False).agg(
        {"Pasos": "mean", "Costo unificado": "mean", "Brecha %": "mean", "USD": "mean",
         "Efectividad": "mean", "Riesgo máx.": "mean", "Expandidos": "sum", "Generados": "sum",
         "Frontera máx.": "mean", "b*": "mean", "Tiempo (ms)": "mean", "Coincide con UCS": "sum"}).reset_index()
    return df, res


if __name__ == "__main__":
    P, m = lectura_datos(ruta_archivo)
    pd.set_option("display.width", 250)
    out = {"Rangos": rangos(P, m), "Participacion": participacion(P, m), "Barrido": barrido(P, m),
           "Pesos_aristas": pesos_aristas(P, m), "Desglose_P004": desglose(P, m)}
    det, res = metricas(P, m)
    out["Metricas_detalle"], out["Metricas_resumen"] = det, res
    for k, df in out.items():
        print(f"\n===== {k} =====\n{df.to_string(index=False)}")
    print("\nPacientes con paliativo usando los pesos del enunciado (sin γ):", pesos_enunciado(P, m))
    os.makedirs(dir_res, exist_ok=True)
    with pd.ExcelWriter(os.path.join(dir_res, "analisis_bloques.xlsx")) as w:
        for k, df in out.items():
            df.to_excel(w, sheet_name=k, index=False)
