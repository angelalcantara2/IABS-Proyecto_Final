"""
reporte.py: Módulo de Reporte Comparativo (Entregable 2).

Procesa los 12 pacientes con BFS, UCS y A*, imprime la tabla en consola y la
exporta a CSV y Excel en la carpeta 'resultados/'.

    python reporte.py

Columnas: ID_Paciente | Algoritmo | Ruta Propuesta | Costo Económico | Nodos Expandidos
(+ columnas de apoyo: Costo Unificado y Efectividad Estimada).
"""
import os
import math
import pandas as pd

from lector import lectura_datos, ruta_archivo
from grafo import construir_grafo, calcular_costo_arista
from busqueda import bfs, ucs, a_estrella, construir_heuristica

ALGORITMOS = ["BFS", "UCS", "A*"]
COLS_BASE = ["ID_Paciente", "Algoritmo", "Ruta Propuesta", "Costo Económico", "Nodos Expandidos"]
dir_res = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "resultados")


def grafo_pac(p, matriz):
    """Grafo del paciente (costo unificado en cada arista) y su heurística."""
    costos = [calcular_costo_arista(f, p) for _, f in matriz.iterrows()]
    G = construir_grafo(matriz, costos)
    h = construir_heuristica(G, p["Severidad_Diagnostico"])
    return G, h


def correr(alg, G, h):
    """Devuelve (ruta, costo_unificado, nodos_expandidos)."""
    if alg == "BFS":
        return bfs(G)
    if alg == "UCS":
        return ucs(G)
    return a_estrella(G, h)


def resumen_ruta(ruta, matriz):
    """(tratamientos, costo USD, efectividad acumulada) de una ruta."""
    trats = ruta[1::2]
    usd, efec = 0, 1.0
    for t in trats:
        f = matriz[matriz["Tratamiento"] == t].iloc[0]
        usd += f["Costo_Financiero_Base"]
        efec *= f["Tasa_Efectividad_Base"]
    return trats, usd, efec


def generar_tabla(pacientes, matriz):
    """Tabla de 36 filas: 12 pacientes x 3 algoritmos."""
    filas = []
    for _, p in pacientes.iterrows():
        G, h = grafo_pac(p, matriz)
        for alg in ALGORITMOS:
            ruta, costo, nodos = correr(alg, G, h)
            if ruta is None:
                filas.append({"ID_Paciente": p["ID_Paciente"], "Algoritmo": alg,
                              "Ruta Propuesta": "Sin ruta", "Costo Económico": None,
                              "Nodos Expandidos": nodos})
                continue
            trats, usd, efec = resumen_ruta(ruta, matriz)
            filas.append({"ID_Paciente": p["ID_Paciente"],
                          "Algoritmo": alg,
                          "Ruta Propuesta": " → ".join(trats),
                          "Costo Económico": usd,
                          "Nodos Expandidos": nodos,
                          "Costo Unificado": round(costo, 2),
                          "Efectividad Estimada": round(efec, 3)})
    return pd.DataFrame(filas)


def resumen_algoritmos(tabla):
    """Rutas distintas, costo promedio (USD y unificado) y nodos por algoritmo."""
    g = tabla.groupby("Algoritmo", sort=False)
    return pd.DataFrame({
        "Rutas distintas": g["Ruta Propuesta"].nunique(),
        "Costo USD promedio": g["Costo Económico"].mean().round(0),
        "Costo unificado promedio": g["Costo Unificado"].mean().round(0),
        "Nodos expandidos (total)": g["Nodos Expandidos"].sum(),
        "Nodos expandidos (prom.)": g["Nodos Expandidos"].mean().round(2),
    }).reset_index()


def imprimir(tabla):
    """Imprime la tabla con el formato del enunciado, agrupada por paciente."""
    t = tabla[COLS_BASE].copy()
    t["Costo Económico"] = t["Costo Económico"].map(lambda x: f"${x:,.0f}")
    anchos = [max(len(c), t[c].astype(str).str.len().max()) for c in COLS_BASE]
    linea = "-+-".join("-" * a for a in anchos)
    fmt = lambda vals: " | ".join(str(v).ljust(a) for v, a in zip(vals, anchos))

    print(fmt(COLS_BASE))
    print(linea)
    pac_ant = None
    for _, f in t.iterrows():
        if pac_ant and f["ID_Paciente"] != pac_ant:
            print(linea)
        print(fmt(f.values))
        pac_ant = f["ID_Paciente"]


def exportar(tabla, resumen, carpeta=dir_res):
    """Guarda CSV y Excel (hoja Reporte + hoja Resumen)."""
    os.makedirs(carpeta, exist_ok=True)
    csv = os.path.join(carpeta, "reporte_comparativo.csv")
    xlsx = os.path.join(carpeta, "reporte_comparativo.xlsx")
    tabla.to_csv(csv, index=False, encoding="utf-8-sig")
    with pd.ExcelWriter(xlsx) as w:
        tabla.to_excel(w, sheet_name="Reporte", index=False)
        resumen.to_excel(w, sheet_name="Resumen", index=False)
    return csv, xlsx


if __name__ == "__main__":
    pacientes, matriz = lectura_datos(ruta_archivo)
    tabla = generar_tabla(pacientes, matriz)
    res = resumen_algoritmos(tabla)

    print("\nREPORTE COMPARATIVO - 12 pacientes x 3 algoritmos\n")
    imprimir(tabla)
    print("\nRESUMEN POR ALGORITMO\n")
    print(res.to_string(index=False))

    csv, xlsx = exportar(tabla, res)
    print(f"\nExportado a:\n  {csv}\n  {xlsx}")
