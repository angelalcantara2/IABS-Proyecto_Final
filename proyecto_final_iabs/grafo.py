import networkx as nx
import matplotlib.pyplot as plt
import math

def construir_grafo(matriz, costos):
    G = nx.MultiDiGraph()
    for origen, destino, trat, costo, efect in zip(matriz["Fase_Origen"],
                                                    matriz["Fase_Destino"],
                                                    matriz["Tratamiento"],
                                                    costos,
                                                    matriz["Tasa_Efectividad_Base"]):
        G.add_edge(origen, destino, key=trat, tratamiento=trat, costo=costo, efectividad=efect)

    return G

def dibujar_grafo(G, titulo="Grafo del paciente", archivo=None):
    pos = {"Diagnostico": (0, 0), "Post_Op_Mayor": (1, -0.6), "Post_Op_Menor": (1, -1.8),
           "Consolidacion": (2, -1.2), "Alta": (3, 0)}

    nx.draw(G, pos, with_labels=True, node_color="lightblue", node_size=2500, font_size=8)

    etiquetas = {}
    for u, v, k, d in G.edges(keys=True, data=True):
        etiquetas[(u, v)] = etiquetas.get((u, v), "") + f"{k}: {d['costo']:,.0f}\n"
    nx.draw_networkx_edge_labels(G, pos, edge_labels=etiquetas, font_size=7)

    plt.title(titulo)
    if archivo:
        plt.savefig(archivo, dpi=300, bbox_inches="tight")
    plt.show()

def calcular_costo_arista(row_tratamiento, row_paciente):
    if row_paciente['Minimizar_Costo_Financiero'] == 1:
        alpha, beta, gamma = 1.0, 2000, 35000
    else:
        alpha, beta, gamma = 0.1, 3000, 30000

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