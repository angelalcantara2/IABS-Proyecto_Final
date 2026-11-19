"""
grafo.py  (Bloques 2 y 3: función de costo y construcción del grafo)
--------------------------------------------------------------------
Convierte la matriz de tratamientos (igual para todos) en un grafo ponderado
PERSONALIZADO para un paciente.

Función de costo unificada de cada arista (tratamiento):

    costo = alfa * plata  +  beta * riesgo  +  gamma * castigo

    plata   = Costo_Financiero_Base del tratamiento
    riesgo  = producto de los multiplicadores de los factores que el paciente TIENE
              (si no tiene ninguno, vale 1)
    castigo = -ln(Tasa_Efectividad_Base)  -> penaliza los tratamientos poco efectivos

    alfa  = cuánto le importa la plata al paciente
    beta  = cuánto le importa el riesgo
    gamma = cuánto le importa que el tratamiento funcione

El término gamma se agregó porque, sin él, Cuidado_Paliativo (barato y sin
riesgo extra) resultaba siempre la ruta de menor costo y UCS repetía a BFS.
"""
import networkx as nx
import matplotlib.pyplot as plt
import math

# Pesos (alfa, beta, gamma) según la preferencia del paciente.
# No varían de forma continua por paciente: son dos juegos fijos que se eligen
# según la preferencia. Lo que sí cambia por paciente es el producto de riesgos.
PESOS_MINIMIZAR_COSTO = (1.0, 2000, 35000)
PESOS_MAXIMIZAR_SUPERVIVENCIA = (0.1, 3000, 30000)


def construir_grafo(matriz, costos, costos_optimistas=None):
    """Arma el grafo dirigido del paciente.

    Es un MultiDiGraph (y no un DiGraph) porque Quimio_Estandar e Inmunoterapia
    unen los mismos dos nodos: un DiGraph se quedaría con una sola arista.

    Parámetros:
        matriz: DataFrame de Matriz_Tratamientos (una fila = una arista).
        costos: lista con el costo unificado de cada fila, en el mismo orden
                que la matriz (ver calcular_costo_arista).
        costos_optimistas: (opcional) lista con el costo de cada fila SUPONIENDO
                que el paciente no tiene factores de riesgo (ver
                calcular_costo_optimista). Lo usa la heurística de A*.
                Si no se entrega, se guarda 0 y A* se comporta igual que UCS.
    Retorna:
        MultiDiGraph. Cada arista tiene los atributos: tratamiento, costo,
        costo_optimista y efectividad.
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
        # key=trat identifica cada arista paralela por el nombre del tratamiento
        G.add_edge(origen, destino, key=trat, tratamiento=trat, costo=costo,
                   costo_optimista=optimista, efectividad=efect)

    return G


def dibujar_grafo(G, titulo="Grafo del paciente", archivo=None, ruta=None):
    """Dibuja el grafo con posiciones fijas y el costo de cada tratamiento.

    Si se entrega la ruta elegida por un algoritmo, sus flechas se pintan de
    ROJO para que se vea de un vistazo el camino propuesto. El resto no cambia.

    Parámetros:
        G: grafo del paciente.
        titulo: título de la figura.
        archivo: (opcional) si se da, también guarda la imagen en esa ruta.
        ruta: (opcional) ruta devuelta por bfs, ucs o a_estrella:
              [nodo, tratamiento, nodo, tratamiento, ..., nodo].
    """
    pos = {"Diagnostico": (0, 0), "Post_Op_Mayor": (1, -0.6), "Post_Op_Menor": (1, -1.8),
           "Consolidacion": (2, -1.2), "Alta": (3, 0)}

    # Aristas de la ruta: (origen, destino, tratamiento). La ruta alterna
    # nodo, tratamiento, nodo, tratamiento, ...
    aristas_ruta = set()
    if ruta:
        for i in range(0, len(ruta) - 2, 2):
            aristas_ruta.add((ruta[i], ruta[i + 2], ruta[i + 1]))

    # Color de cada flecha, en el mismo orden en que networkx las dibuja
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
    """Devuelve (alfa, beta, gamma) según la preferencia del paciente.

    Si Minimizar_Costo_Financiero es 1 usa el juego de 'minimizar costo';
    en cualquier otro caso usa el de 'maximizar supervivencia'.
    """
    if row_paciente['Minimizar_Costo_Financiero'] == 1:
        return PESOS_MINIMIZAR_COSTO
    return PESOS_MAXIMIZAR_SUPERVIVENCIA


def calcular_costo_arista(row_tratamiento, row_paciente):
    """Costo unificado de un tratamiento para un paciente concreto.

    costo = alfa * plata + beta * riesgo + gamma * castigo
    """
    alpha, beta, gamma = obtener_pesos(row_paciente)

    # Producto de los multiplicadores de los factores que el paciente TIENE.
    # Un factor en 0 no multiplica nada (equivale a multiplicar por 1).
    producto_riesgos = 1.0
    if row_paciente['Insuficiencia_Renal'] == 1:
        producto_riesgos *= row_tratamiento['Multiplicador_Riesgo_Renal']
    if row_paciente['Cardiopatia'] == 1:
        producto_riesgos *= row_tratamiento['Multiplicador_Riesgo_Cardio']
    if row_paciente['Edad_Avanzada_Fragil'] == 1:
        producto_riesgos *= row_tratamiento['Multiplicador_Riesgo_Edad']
    if row_paciente['Inmunosupresion'] == 1:
        producto_riesgos *= row_tratamiento['Multiplicador_Inmuno']

    # -ln(efectividad): vale 0 si la efectividad fuera 100% y crece rápido cuando baja
    castigo = -math.log(row_tratamiento['Tasa_Efectividad_Base'])

    return (row_tratamiento['Costo_Financiero_Base'] * alpha
            + beta * producto_riesgos
            + gamma * castigo)


def calcular_costo_optimista(row_tratamiento, row_paciente):
    """Costo del tratamiento SUPONIENDO que el paciente no tiene factores de riesgo.

    Es igual a calcular_costo_arista pero con el producto de riesgos fijo en 1.
    Como todos los multiplicadores de la matriz son >= 1, este costo nunca es
    mayor que el costo real. Por eso sirve para construir una heurística
    admisible (ver construir_heuristica en busqueda.py).
    """
    alpha, beta, gamma = obtener_pesos(row_paciente)
    castigo = -math.log(row_tratamiento['Tasa_Efectividad_Base'])

    return (row_tratamiento['Costo_Financiero_Base'] * alpha
            + beta * 1.0
            + gamma * castigo)
