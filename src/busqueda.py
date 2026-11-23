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
    queue = deque([(inicio, [inicio], 0)]) # (nodo, ruta, costo)
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
    heapq.heappush(frontera, (0, inicio, [inicio])) # (g, nodo, ruta)
    costos_minimos = {inicio : 0}
    pasos = 0

    while frontera:
        costo_actual, nodo_actual, camino_actual = heapq.heappop(frontera)

        if costo_actual > costos_minimos[nodo_actual]: # entrada obsoleta
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

def construir_heuristica(grafo, meta='Alta'):
    """h(n) = camino más corto hasta la meta, donde cada paso cuesta
    la arista de salida más barata de su nodo de origen.
    Cada paso real cuesta al menos eso, así que h nunca sobreestima.
    """
    minimo_saliente = {}
    for nodo in grafo.nodes():
        costos = [datos['costo'] for _, _, datos in grafo.out_edges(nodo, data=True)]
        minimo_saliente[nodo] = min(costos, default=0)

    h = {nodo: float('inf') for nodo in grafo.nodes()}
    h[meta] = 0
    for _ in range(len(h)): # repite hasta que se estabilice
        for nodo in grafo.nodes():
            for _, vecino in grafo.out_edges(nodo):
                h[nodo] = min(h[nodo], minimo_saliente[nodo] + h[vecino])
    return h

def a_estrella(grafo, heuristica, inicio='Diagnostico', meta='Alta'):
    """Nivel 3: A* con lista cerrada. La frontera se ordena por f = g + h."""
    frontera = []
    heapq.heappush(frontera, (heuristica[inicio], 0, inicio, [inicio])) # (f, g, nodo, ruta)
    closed = set()
    pasos = 0

    while frontera:
        f, g, actual, ruta = heapq.heappop(frontera)

        if actual not in closed:
            closed.add(actual)
            pasos += 1

            if actual == meta:
                return ruta, g, pasos

            for _, vecino, trat, datos in grafo.out_edges(actual, keys=True, data=True):
                if vecino not in closed:
                    nuevo_g = g + datos['costo']
                    heapq.heappush(frontera, (nuevo_g + heuristica[vecino], nuevo_g, vecino, ruta + [trat, vecino]))

    return None, None, pasos







    #xd