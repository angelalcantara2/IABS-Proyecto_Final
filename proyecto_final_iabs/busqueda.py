from collections import deque

def bfs(grafo, inicio='Diagnostico', meta='Alta'):
    queue = deque([(inicio, [inicio], 0)])
    visitados = set([inicio])
    pasos = 0 
    
    while queue:
        nodo_actual, path, costo = queue.popleft()
        pasos += 1
  
        if nodo_actual == meta:
            return path, costo, pasos

        for origen, vecino, tratamiento, datos in grafo.out_edges(nodo_actual, keys=True, data=True):
            if vecino not in visitados:
                visitados.add(vecino)
                queue.append((vecino, path + [tratamiento, vecino], costo + datos['costo']))
    
    return None, None, pasos

