"""
simulaciones.py: simulaciones del pipeline para el análisis crítico (Entregable 3).

    python simulaciones.py

S1  BFS sobre los 12 pacientes, y BFS si se quita Cuidado_Paliativo
S2  UCS: paciente sin factores (P001) frente a paciente con todos (P004)
S3  Mismo paciente cambiando solo la preferencia (los 12)
S4  Barrido de gamma: cuántos pacientes reciben paliativo y umbral de P009
S5  A* con h = 0, h de severidad, h = c_min y h = h* (perfecta)
S6  Guardarraíl ético: piso de efectividad y techo de riesgo por tratamiento

Guarda todo en resultados/simulaciones.xlsx y las figuras en resultados/.
"""
import os
import math
import pandas as pd
import networkx as nx
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

from lector import lectura_datos, ruta_archivo
import grafo as gf
from grafo import construir_grafo, calcular_costo_arista
from busqueda import bfs, ucs, a_estrella, construir_heuristica
from reporte import resumen_ruta, dir_res

FACT = {"Insuficiencia_Renal": "Multiplicador_Riesgo_Renal",
        "Cardiopatia": "Multiplicador_Riesgo_Cardio",
        "Edad_Avanzada_Fragil": "Multiplicador_Riesgo_Edad",
        "Inmunosupresion": "Multiplicador_Inmuno"}
AZUL, NARANJA, AQUA, AMARILLO = "#2a78d6", "#eb6834", "#1baf7a", "#eda100"


# ---------- utilidades ----------
def costos_pac(p, m, pesos=None):
    """Costo unificado de cada arista; pesos=(a,b,g) fuerza otros pesos."""
    if pesos is None:
        return [calcular_costo_arista(f, p) for _, f in m.iterrows()]
    viejo = gf.obtener_pesos
    gf.obtener_pesos = lambda _p: pesos
    try:
        return [calcular_costo_arista(f, p) for _, f in m.iterrows()]
    finally:
        gf.obtener_pesos = viejo


def riesgo(p, f):
    """Producto de multiplicadores activos de un tratamiento para el paciente."""
    r = 1.0
    for fac, col in FACT.items():
        if p[fac] == 1:
            r *= f[col]
    return r


def perfil(p):
    cortos = {"Insuficiencia_Renal": "Renal", "Cardiopatia": "Cardio",
              "Edad_Avanzada_Fragil": "Edad", "Inmunosupresion": "Inmuno"}
    facs = [cortos[k] for k in FACT if p[k] == 1]
    pref = "Min. costo" if p["Minimizar_Costo_Financiero"] == 1 else "Max. superv."
    return f"Sev {p['Severidad_Diagnostico']}; {', '.join(facs) or 'sin factores'}; {pref}"


def todas_rutas(G, ini="Diagnostico", meta="Alta"):
    """Las 5 rutas posibles con su costo unificado (el grafo es pequeño)."""
    rutas = []
    for camino in nx.all_simple_edge_paths(G, ini, meta):
        trats = [k for _, _, k in camino]
        c = sum(G.edges[u, v, k]["costo"] for u, v, k in camino)
        rutas.append((trats, c))
    return sorted(rutas, key=lambda x: x[1])


def riesgo_ruta(p, m, trats):
    """Riesgo máximo de un tratamiento dentro de la ruta (peor paso)."""
    return max(riesgo(p, m[m["Tratamiento"] == t].iloc[0]) for t in trats)


def h_real(G, meta="Alta"):
    """h*(n): costo real mínimo hasta la meta (Dijkstra en el grafo invertido)."""
    R = nx.DiGraph()
    for u, v, d in G.edges(data=True):
        if not R.has_edge(v, u) or d["costo"] < R[v][u]["costo"]:
            R.add_edge(v, u, costo=d["costo"])
    return nx.single_source_dijkstra_path_length(R, meta, weight="costo")


def flecha(trats):
    return " → ".join(trats)


# ---------- simulaciones ----------
def s1(P, m):
    """BFS normal y BFS sin la arista de Cuidado_Paliativo."""
    m2 = m[m["Tratamiento"] != "Cuidado_Paliativo"].reset_index(drop=True)
    filas = []
    for _, p in P.iterrows():
        r1, _, _ = bfs(construir_grafo(m, costos_pac(p, m)))
        G2 = construir_grafo(m2, costos_pac(p, m2))
        r2, c2, _ = bfs(G2)
        ru, cu, _ = ucs(G2)
        t2, tu = r2[1::2], ru[1::2]
        filas.append({"ID": p["ID_Paciente"], "Perfil": perfil(p),
                      "BFS": flecha(r1[1::2]),
                      "Efect. BFS": resumen_ruta(r1, m)[2],
                      "BFS sin paliativo": flecha(t2),
                      "Riesgo máx. BFS s/p": riesgo_ruta(p, m, t2),
                      "UCS sin paliativo": flecha(tu),
                      "Riesgo máx. UCS s/p": riesgo_ruta(p, m, tu)})
    return pd.DataFrame(filas)


def s2(P, m):
    """Costo unificado de cada arista para P001 y P004."""
    filas = []
    for pid in ["P001", "P004"]:
        p = P[P["ID_Paciente"] == pid].iloc[0]
        for (_, f), c in zip(m.iterrows(), costos_pac(p, m)):
            filas.append({"ID": pid, "Tratamiento": f["Tratamiento"],
                          "Tramo": f"{f['Fase_Origen']}→{f['Fase_Destino']}",
                          "Riesgo (Π mult.)": round(riesgo(p, f), 3),
                          "Costo unificado": round(c)})
    return pd.DataFrame(filas)


def s3(P, m):
    """Cada paciente con su preferencia real y con la contraria."""
    filas = []
    for _, p in P.iterrows():
        for nom, pesos in [("Min. costo", gf.PESOS_MINIMIZAR_COSTO),
                           ("Max. superv.", gf.PESOS_MAXIMIZAR_SUPERVIVENCIA)]:
            ru, c, _ = ucs(construir_grafo(m, costos_pac(p, m, pesos)))
            t, usd, ef = resumen_ruta(ru, m)
            real = (p["Minimizar_Costo_Financiero"] == 1) == (nom == "Min. costo")
            filas.append({"ID": p["ID_Paciente"], "Perfil": perfil(p),
                          "Preferencia simulada": nom + (" (real)" if real else ""),
                          "Pesos (α, β, γ)": str(pesos), "Ruta UCS": flecha(t),
                          "USD": usd, "Efectividad": round(ef, 3),
                          "Riesgo máx.": riesgo_ruta(p, m, t), "Costo unificado": round(c)})
    return pd.DataFrame(filas)


def s4(P, m, gammas=range(0, 60001, 2500)):
    """Barrido de gamma (alfa y beta fijos de cada preferencia)."""
    filas = []
    for g in gammas:
        n_pal, rutas_p9 = 0, None
        for _, p in P.iterrows():
            a, b, _ = gf.obtener_pesos(p)
            ru, _, _ = ucs(construir_grafo(m, costos_pac(p, m, (a, b, g))))
            n_pal += ru[1::2] == ["Cuidado_Paliativo"]
            if p["ID_Paciente"] == "P009":
                rutas_p9 = flecha(ru[1::2])
        filas.append({"gamma": g, "Pacientes con paliativo": n_pal, "Ruta P009": rutas_p9})
    return pd.DataFrame(filas)


def umbral_gamma(P, m, pid="P009", lo=0, hi=200000):
    """Menor gamma con el que el paciente deja de recibir paliativo (bisección)."""
    p = P[P["ID_Paciente"] == pid].iloc[0]
    a, b, _ = gf.obtener_pesos(p)
    pal = lambda g: ucs(construir_grafo(m, costos_pac(p, m, (a, b, g))))[0][1::2] == ["Cuidado_Paliativo"]
    if pal(hi):
        return None
    for _ in range(60):
        mid = (lo + hi) / 2
        lo, hi = (mid, hi) if pal(mid) else (lo, mid)
    return hi


def s5(P, m):
    """Nodos expandidos de A* con cuatro heurísticas distintas."""
    filas = []
    for _, p in P.iterrows():
        G = construir_grafo(m, costos_pac(p, m))
        hs = {"h = 0 (UCS)": {n: 0 for n in G.nodes()},
              "h severidad (equipo)": construir_heuristica(G, p["Severidad_Diagnostico"]),
              "h = c_min (factor 1)": construir_heuristica(G, 5),
              "h* (perfecta)": h_real(G)}
        fila = {"ID": p["ID_Paciente"], "Severidad": p["Severidad_Diagnostico"]}
        rutas = set()
        for nom, h in hs.items():
            ru, c, n = a_estrella(G, h)
            fila[nom] = n
            rutas.add(tuple(ru))
        fila["Misma ruta"] = "Sí" if len(rutas) == 1 else "No"
        filas.append(fila)
    return pd.DataFrame(filas)


def s6(P, m, piso=0.30, techo=3.0):
    """UCS con guardarraíl: descarta rutas con efectividad < piso o con algún
    tratamiento de riesgo > techo, salvo que no quede ninguna ruta."""
    filas = []
    for _, p in P.iterrows():
        G = construir_grafo(m, costos_pac(p, m))
        rutas = todas_rutas(G)
        libre = rutas[0][0]
        ok = [(t, c) for t, c in rutas
              if resumen_ruta(["x"] + sum([[x, "x"] for x in t], []), m)[2] >= piso
              and riesgo_ruta(p, m, t) <= techo]
        guard = ok[0][0] if ok else libre
        filas.append({"ID": p["ID_Paciente"], "Perfil": perfil(p),
                      "UCS sin guardarraíl": flecha(libre),
                      "UCS con guardarraíl": flecha(guard),
                      "Cambió": "Sí" if guard != libre else "No",
                      "Rutas válidas": len(ok)})
    return pd.DataFrame(filas)


def rutas_p(P, m, pid):
    """Las 5 rutas de un paciente con costo, efectividad y riesgo."""
    p = P[P["ID_Paciente"] == pid].iloc[0]
    filas = []
    for t, c in todas_rutas(construir_grafo(m, costos_pac(p, m))):
        f = resumen_ruta(["x"] + sum([[x, "x"] for x in t], []), m)
        filas.append({"ID": pid, "Ruta": flecha(t), "Costo unificado": round(c),
                      "USD": f[1], "Efectividad": round(f[2], 3),
                      "Riesgo máx.": riesgo_ruta(p, m, t)})
    return pd.DataFrame(filas)


# ---------- figuras ----------
def estilo(ax):
    for s in ["top", "right"]:
        ax.spines[s].set_visible(False)
    for s in ["left", "bottom"]:
        ax.spines[s].set_color("#b5b4ae")
    ax.tick_params(colors="#52514e", labelsize=8)
    ax.grid(axis="y", color="#e6e5e0", linewidth=0.6)
    ax.set_axisbelow(True)


def fig_s4(df, ruta):
    fig, ax = plt.subplots(figsize=(6.4, 2.8), dpi=200)
    ax.step(df["gamma"], df["Pacientes con paliativo"], where="post", color=AZUL, linewidth=2)
    ax.axvline(35000, color="#52514e", linestyle="--", linewidth=1)
    ax.axvline(30000, color="#52514e", linestyle=":", linewidth=1)
    ax.text(35500, 10.5, "γ min. costo\n(35.000)", fontsize=7, color="#52514e")
    ax.text(29500, 7.2, "γ max. superv.\n(30.000)", fontsize=7, color="#52514e", ha="right")
    ax.set_xlabel("γ (peso de la efectividad)", fontsize=8, color="#52514e")
    ax.set_ylabel("Pacientes con\nCuidado_Paliativo", fontsize=8, color="#52514e")
    ax.set_ylim(0, 12.5)
    ax.xaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda x, _: f"{x/1000:.0f}k"))
    estilo(ax)
    fig.tight_layout()
    fig.savefig(ruta)
    plt.close(fig)


def fig_s5(df, ruta):
    cols = ["h = 0 (UCS)", "h severidad (equipo)", "h = c_min (factor 1)", "h* (perfecta)"]
    colores = ["#86b6ef", NARANJA, "#86b6ef", "#86b6ef"]  # resalta la h del equipo
    tot = [df[c].sum() for c in cols]
    fig, ax = plt.subplots(figsize=(6.4, 2.6), dpi=200)
    barras = ax.bar(range(4), tot, color=colores, width=0.55, edgecolor="#fcfcfb", linewidth=2)
    for b, v in zip(barras, tot):
        ax.text(b.get_x() + b.get_width() / 2, v + 0.8, str(v), ha="center", fontsize=8, color="#0b0b0b")
    ax.set_xticks(range(4), cols, fontsize=8)
    ax.set_ylabel("Nodos expandidos\n(12 pacientes)", fontsize=8, color="#52514e")
    ax.set_ylim(0, max(tot) * 1.2)
    estilo(ax)
    fig.tight_layout()
    fig.savefig(ruta)
    plt.close(fig)


if __name__ == "__main__":
    P, m = lectura_datos(ruta_archivo)
    os.makedirs(dir_res, exist_ok=True)
    pd.set_option("display.width", 250)
    pd.set_option("display.max_colwidth", 60)

    res = {"S1_BFS": s1(P, m), "S2_UCS_P001_P004": s2(P, m), "S3_Preferencia": s3(P, m),
           "S4_Barrido_gamma": s4(P, m), "S5_Heuristicas": s5(P, m), "S6_Guardarrail": s6(P, m),
           "Rutas_P009": rutas_p(P, m, "P009"), "Rutas_P004": rutas_p(P, m, "P004"),
           "Rutas_P005": rutas_p(P, m, "P005")}
    g9 = umbral_gamma(P, m)

    for k, df in res.items():
        print(f"\n===== {k} =====")
        print(df.to_string(index=False))
    print(f"\nUmbral de gamma para que P009 deje el paliativo: {g9:,.0f}")

    with pd.ExcelWriter(os.path.join(dir_res, "simulaciones.xlsx")) as w:
        for k, df in res.items():
            df.to_excel(w, sheet_name=k[:31], index=False)
    fig_s4(res["S4_Barrido_gamma"], os.path.join(dir_res, "fig_barrido_gamma.png"))
    fig_s5(res["S5_Heuristicas"], os.path.join(dir_res, "fig_heuristicas.png"))
    print(f"\nResultados en {dir_res}")
