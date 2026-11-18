from collections import deque
import heapq

def bfs(grafo, inicio='Diagnostico', meta='Alta'):
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
    frontera = []
    heapq.heappush(frontera, (0, inicio, [inicio]))
    costos_minimos = {inicio : 0}
    pasos = 0

    while frontera:
        costo_actual, nodo_actual, camino_actual = heapq.heappop(frontera)

        if costo_actual > costos_minimos[nodo_actual]:
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
    