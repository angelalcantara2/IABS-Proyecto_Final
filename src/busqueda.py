"""
busqueda.py  (Bloques 4, 5 y 6: los tres niveles de búsqueda)
--------------------------------------------------------------
Los tres algoritmos recorren el mismo grafo del paciente. Lo único que cambia
es CÓMO se ordena la frontera (los nodos descubiertos que aún no se expanden):

    Nivel 1  BFS : por orden de llegada (cola FIFO). Solo cuenta pasos.
    Nivel 2  UCS : por costo acumulado g(n).
    Nivel 3  A*  : por f(n) = g(n) + h(n), con una heurística h(n).

Los tres devuelven (ruta, costo, nodos_expandidos).
La ruta es una lista alternada [nodo, tratamiento, nodo, tratamiento, ..., nodo].
Los algoritmos se programan a mano con deque y heapq; NetworkX solo se usa
para leer las aristas de salida de cada nodo.

Convención de 'nodos expandidos' (igual en los tres): cuenta cada nodo que se
saca de la frontera, incluido el nodo meta cuando se extrae.
"""
from collections import deque
import heapq

# Severidad máxima del Excel (1 = etapa temprana, 5 = etapa crítica)
SEVERIDAD_MAXIMA = 5


def bfs(grafo, inicio='Diagnostico', meta='Alta'):
    """Nivel 1: búsqueda en anchura. Minimiza el NÚMERO DE PASOS.

    Ignora los costos y los riesgos (cada arista cuenta como un paso).
    Por eso siempre encuentra Cuidado_Paliativo (1 paso de Diagnostico a Alta),
    aunque sea clínicamente absurdo para un paciente joven y sano.
    El costo que devuelve es el costo unificado real de la ruta encontrada,
    solo para poder compararla con UCS y A*.
    """
    # Cada elemento de la cola: (nodo, ruta hasta él, costo acumulado)
    queue = deque([(inicio, [inicio], 0)])
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
    """Nivel 2: búsqueda de costo uniforme. Minimiza el COSTO UNIFICADO.

    La frontera es un min-heap ordenado por el costo acumulado g(n).
    La meta se evalúa al EXTRAER el nodo de la cola (no al agregarlo), que es
    lo que garantiza que la ruta encontrada sea la de menor costo.
    """
    # Cada elemento del heap: (costo acumulado, nodo, ruta hasta él)
    frontera = []
    heapq.heappush(frontera, (0, inicio, [inicio]))
    costos_minimos = {inicio : 0}   # mejor costo conocido hasta cada nodo
    pasos = 0

    while frontera:
        costo_actual, nodo_actual, camino_actual = heapq.heappop(frontera)

        # Entrada obsoleta: después de agregarla se encontró un camino más barato
        if costo_actual > costos_minimos[nodo_actual]:
            continue
        pasos += 1

        if nodo_actual == meta:
            return camino_actual, costo_actual, pasos

        for origen, vecino, trat, datos in grafo.out_edges(nodo_actual, keys=True, data=True):
            nuevo_costo = costo_actual + datos['costo']

            # Solo se agrega si es la primera vez o si el camino nuevo es más barato
            if vecino not in costos_minimos or nuevo_costo < costos_minimos[vecino]:
                costos_minimos[vecino] = nuevo_costo
                nueva_ruta = camino_actual + [trat, vecino]
                heapq.heappush(frontera, (nuevo_costo, vecino, nueva_ruta))

    return None, None, pasos


def construir_heuristica(grafo, severidad, meta='Alta'):
    """Calcula h(n) para todos los nodos del grafo de un paciente.

    Idea ("problema relajado"): h(n) es el costo mínimo para llegar a la meta
    SI EL PACIENTE NO TUVIERA FACTORES DE RIESGO (usando 'costo_optimista'),
    multiplicado por severidad / SEVERIDAD_MAXIMA.

        h(n) = (severidad / 5) * costo_optimista_minimo(n -> meta)

    Por qué es admisible (nunca sobreestima el costo real que falta):
        1. costo_optimista <= costo real en cada arista, porque los
           multiplicadores de riesgo valen >= 1.
        2. severidad / 5 <= 1.
    Por qué es consistente: la distancia mínima en el grafo relajado cumple
    h(n) <= costo(n, n') + h(n'), y multiplicar por un número entre 0 y 1
    mantiene esa desigualdad.
    Además h(meta) = 0.

    Papel de la severidad: en un caso grave (5) h usa el estimado completo y
    A* se guía con más fuerza (decisión más rápida); en un caso leve (1) h
    vale una quinta parte y A* explora más, parecido a UCS.

    La distancia mínima se calcula con relajaciones repetidas sobre las
    aristas (estilo Bellman-Ford), recorriendo el grafo "hacia atrás".

    Parámetros:
        grafo: grafo del paciente (aristas con atributo 'costo_optimista').
        severidad: Severidad_Diagnostico del paciente (1 a 5).
        meta: nodo meta.
    Retorna:
        dict {nodo: h(nodo)}
    """
    factor = severidad / SEVERIDAD_MAXIMA

    infinito = float('inf')
    distancia = {}
    for nodo in grafo.nodes():
        distancia[nodo] = infinito
    distancia[meta] = 0

    # Con N nodos, N pasadas bastan para que la distancia mínima se propague
    for i in range(grafo.number_of_nodes()):
        for origen, destino, datos in grafo.edges(data=True):
            if distancia[destino] + datos['costo_optimista'] < distancia[origen]:
                distancia[origen] = distancia[destino] + datos['costo_optimista']

    heuristica = {}
    for nodo in grafo.nodes():
        heuristica[nodo] = factor * distancia[nodo]
    return heuristica


def a_estrella(grafo, heuristica, inicio='Diagnostico', meta='Alta'):
    """Nivel 3: búsqueda A*. Minimiza el costo unificado con ayuda de h(n).

    Es igual a UCS, pero la frontera se ordena por f(n) = g(n) + h(n):
        g(n) = costo real acumulado desde el inicio (el pasado)
        h(n) = estimación de lo que falta hasta la meta (el futuro)
    Si h es admisible y consistente, A* devuelve la misma ruta óptima que UCS
    pero expandiendo igual o menos nodos.

    Parámetros:
        grafo: grafo del paciente.
        heuristica: dict {nodo: h(nodo)}, ver construir_heuristica().
        inicio, meta: nodos de inicio y de meta.
    """
    # Cada elemento del heap: (f, g, nodo, ruta hasta él)
    frontera = []
    heapq.heappush(frontera, (heuristica[inicio], 0, inicio, [inicio]))
    costos_minimos = {inicio : 0}   # mejor g conocido hasta cada nodo
    pasos = 0

    while frontera:
        f_actual, costo_actual, nodo_actual, camino_actual = heapq.heappop(frontera)

        # Entrada obsoleta (se encontró antes un camino más barato a este nodo)
        if costo_actual > costos_minimos[nodo_actual]:
            continue
        pasos += 1

        # Igual que en UCS, la meta se evalúa al extraer el nodo
        if nodo_actual == meta:
            return camino_actual, costo_actual, pasos

        for origen, vecino, trat, datos in grafo.out_edges(nodo_actual, keys=True, data=True):
            nuevo_costo = costo_actual + datos['costo']

            if vecino not in costos_minimos or nuevo_costo < costos_minimos[vecino]:
                costos_minimos[vecino] = nuevo_costo
                nueva_ruta = camino_actual + [trat, vecino]
                # La prioridad es f = g + h
                heapq.heappush(frontera, (nuevo_costo + heuristica[vecino], nuevo_costo, vecino, nueva_ruta))

    return None, None, pasos
