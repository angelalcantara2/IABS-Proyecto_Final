import networkx as nx
import matplotlib.pyplot as plt

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