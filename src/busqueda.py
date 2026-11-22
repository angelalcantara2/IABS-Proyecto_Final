"""
busqueda.py: los tres niveles de búsqueda.

    Nivel 1  BFS : cola FIFO, cuenta pasos.
    Nivel 2  UCS : min-heap por costo acumulado g(n).
    Nivel 3  A*  : min-heap por f(n) = g(n) + h(n).

Los tres devuelven (ruta, costo, nodos_expandidos), donde la ruta alterna
[nodo, tratamiento, nodo, ...]. Nodos expandidos incluye la meta al extraerla.
"""
from collections import deque
import heapq

SEVERIDAD_MAXIMA = 5


def bfs(grafo, inicio='Diagnostico', meta='Alta'):
    """Nivel 1: ruta con menos pasos, sin mirar costos ni riesgos."""
    queue = deque([(inicio, [inicio], 0)])   # (nodo, ruta, costo)
    visitados = set([inicio])
    pasos = 0

    while queue:
        nodo_actual, path, costo = queue.popleft()
        pasos += 1

        if nodo_actual == meta:
            return path, costo, pasos

        for origen, vecino, trat, datos in grafo.out_edges(nodo_actual, keys=True, data=True):
            if vecino not in visitados:
                visitados.add(vecino)
                queue.append((vecino, path + [trat, vecino], costo + datos['costo']))

    return None, None, pasos


def ucs(grafo, inicio='Diagnostico', meta='Alta'):
    """Nivel 2: ruta de menor costo unificado. La meta se evalúa al extraer."""
    frontera = []
    heapq.heappush(frontera, (0, inicio, [inicio]))   # (g, nodo, ruta)
    costos_minimos = {inicio : 0}
    pasos = 0

    while frontera:
        costo_actual, nodo_actual, camino_actual = heapq.heappop(frontera)

        if costo_actual > costos_minimos[nodo_actual]:   # entrada obsoleta
            continue
        pasos += 1

        if nodo_actual == meta:
            return camino_actual, costo_actual, pasos

        for origen, vecino, trat, datos in grafo.out_edges(nodo_actual, keys=True, data=True):
            nuevo_costo = costo_actual + datos['costo']
            if vecino not in costos_minimos or nuevo_costo < costos_minimos[vecino]:
                costos_minimos[vecino] = nuevo_costo
                nueva_ruta = camino_actual + [trat, vecino]
                heapq.heappush(frontera, (nuevo_costo, vecino, nueva_ruta))

    return None, None, pasos


# Pasos mínimos que faltan desde cada fase hasta Alta
PASOS_A_LA_META = {'Diagnostico': 1, 'Post_Op_Mayor': 2, 'Post_Op_Menor': 2,
                   'Consolidacion': 1, 'Alta': 0}


def construir_heuristica(grafo, severidad):
    """h(n) = (severidad/5) * pasos que faltan * costo de la arista más barata.

    Cada paso que falta cuesta al menos la arista más barata, así que h
    nunca sobreestima: es admisible y consistente.
    """
    costo_minimo = min(datos['costo'] for _, _, datos in grafo.edges(data=True))
    factor = severidad / SEVERIDAD_MAXIMA

    heuristica = {}
    for nodo in grafo.nodes():
        heuristica[nodo] = factor * PASOS_A_LA_META[nodo] * costo_minimo
    return heuristica


def a_estrella(grafo, heuristica, inicio='Diagnostico', meta='Alta'):
    """Nivel 3: como UCS, pero la frontera se ordena por f = g + h."""
    frontera = []
    heapq.heappush(frontera, (heuristica[inicio], 0, inicio, [inicio]))   # (f, g, nodo, ruta)
    costos_minimos = {inicio : 0}
    pasos = 0

    while frontera:
        f_actual, costo_actual, nodo_actual, camino_actual = heapq.heappop(frontera)

        if costo_actual > costos_minimos[nodo_actual]:   # entrada obsoleta
            continue
        pasos += 1

        if nodo_actual == meta:
            return camino_actual, costo_actual, pasos

        for origen, vecino, trat, datos in grafo.out_edges(nodo_actual, keys=True, data=True):
            nuevo_costo = costo_actual + datos['costo']
            if vecino not in costos_minimos or nuevo_costo < costos_minimos[vecino]:
                costos_minimos[vecino] = nuevo_costo
                nueva_ruta = camino_actual + [trat, vecino]
                heapq.heappush(frontera, (nuevo_costo + heuristica[vecino], nuevo_costo, vecino, nueva_ruta))

    return None, None, pasos
