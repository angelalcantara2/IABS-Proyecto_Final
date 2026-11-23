"""
grafo.py: función de costo y construcción del grafo de cada paciente.

    costo = alfa * plata + beta * riesgo + gamma * castigo

    plata   = Costo_Financiero_Base
    riesgo  = producto de los multiplicadores de los factores activos (1 si no hay)
    castigo = -ln(Tasa_Efectividad_Base)

Gamma evita que Cuidado_Paliativo sea siempre la ruta más barata.
"""
import networkx as nx
import matplotlib.pyplot as plt
import math

# (alfa, beta, gamma) según la preferencia del paciente
PESOS_MINIMIZAR_COSTO = (1.0, 2000, 35000)
PESOS_MAXIMIZAR_SUPERVIVENCIA = (0.1, 3000, 30000)


def construir_grafo(matriz, costos, costos_optimistas=None):
    """Arma el MultiDiGraph del paciente (admite aristas paralelas).

    costos_optimistas es opcional; si no se da, A* se comporta como UCS.
    """
    if costos_optimistas is None:
        costos_optimistas = [0] * len(costos)

    G = nx.MultiDiGraph()
    for origen, destino, trat, costo, optimista, efect in zip(matriz["Fase_Origen"],
                                                              matriz["Fase_Destino"],
                                                              matriz["Tratamiento"],
                                                              costos,
                                                              costos_optimistas,
                                                              matriz["Tasa_Efectividad_Base"]):
        G.add_edge(origen, destino, key=trat, tratamiento=trat, costo=costo,
                   costo_optimista=optimista, efectividad=efect)

    return G

def dibujar_grafo(G, titulo="Grafo del paciente", archivo=None, ruta=None):
    """Dibuja el grafo con los costos; si se da la ruta, la pinta en rojo."""
    pos = {"Diagnostico": (0, 0), "Post_Op_Mayor": (1, -0.6), "Post_Op_Menor": (1, -1.8),
           "Consolidacion": (2, -1.2), "Alta": (3, 0)}

    # (origen, destino, tratamiento) de cada paso de la ruta
    aristas_ruta = set()
    if ruta:
        for i in range(0, len(ruta) - 2, 2):
            aristas_ruta.add((ruta[i], ruta[i + 2], ruta[i + 1]))

    colores_aristas = []
    for u, v, k in G.edges(keys=True):
        if (u, v, k) in aristas_ruta:
            colores_aristas.append("red")
        else:
            colores_aristas.append("black")

    nx.draw(G, pos, with_labels=True, node_color="lightblue", node_size=2500, font_size=8,
            edge_color=colores_aristas)

    etiquetas = {}
    for u, v, k, d in G.edges(keys=True, data=True):
        etiquetas[(u, v)] = etiquetas.get((u, v), "") + f"{k}: {d['costo']:,.0f}\n"
    nx.draw_networkx_edge_labels(G, pos, edge_labels=etiquetas, font_size=7)

    plt.title(titulo)
    if archivo:
        plt.savefig(archivo, dpi=300, bbox_inches="tight")
    plt.show()


def obtener_pesos(row_paciente):
    """Devuelve (alfa, beta, gamma) según la preferencia del paciente."""
    if row_paciente['Minimizar_Costo_Financiero'] == 1:
        return PESOS_MINIMIZAR_COSTO
    return PESOS_MAXIMIZAR_SUPERVIVENCIA


def calcular_costo_arista(row_tratamiento, row_paciente):
    """Costo unificado de un tratamiento para un paciente."""
    alpha, beta, gamma = obtener_pesos(row_paciente)

    producto_riesgos = 1.0
    if row_paciente['Insuficiencia_Renal'] == 1:
        producto_riesgos *= row_tratamiento['Multiplicador_Riesgo_Renal']
    if row_paciente['Cardiopatia'] == 1:
        producto_riesgos *= row_tratamiento['Multiplicador_Riesgo_Cardio']
    if row_paciente['Edad_Avanzada_Fragil'] == 1:
        producto_riesgos *= row_tratamiento['Multiplicador_Riesgo_Edad']
    if row_paciente['Inmunosupresion'] == 1:
        producto_riesgos *= row_tratamiento['Multiplicador_Inmuno']

    castigo = -math.log(row_tratamiento['Tasa_Efectividad_Base'])

    return (row_tratamiento['Costo_Financiero_Base'] * alpha
            + beta * producto_riesgos
            + gamma * castigo)


def calcular_costo_optimista(row_tratamiento, row_paciente):
    """Costo con riesgo fijo en 1: nunca supera el real, sirve para la heurística."""
    alpha, beta, gamma = obtener_pesos(row_paciente)
    castigo = -math.log(row_tratamiento['Tasa_Efectividad_Base'])

    return (row_tratamiento['Costo_Financiero_Base'] * alpha
            + beta * 1.0
            + gamma * castigo)