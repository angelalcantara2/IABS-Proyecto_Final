"""
main.py  (Bloque 7: programa principal)
---------------------------------------
Menú interactivo: se elige un paciente y un algoritmo de búsqueda, y el programa
muestra la ruta propuesta, su costo y los nodos expandidos.

Flujo para un paciente:
    Excel -> costo de cada tratamiento -> grafo del paciente -> BFS / UCS / A*

Se ejecuta con:  python main.py
"""
from lector import lectura_datos, obtener_paciente, ruta_archivo
from grafo import (construir_grafo, dibujar_grafo, calcular_costo_arista,
                   calcular_costo_optimista)
from busqueda import bfs, ucs, a_estrella, construir_heuristica

pacientes, matriz = lectura_datos(ruta_archivo)


def main():
    """Menú principal: consultar un paciente o salir."""
    while True:
        print("="*20, " Sistema de Soporte a Decisiones Clínicas mediante Algoritmos de Búsqueda (SSDC) ", "="*20)

        print("\nOpciones de menú:")
        print("1. Consultar paciente")
        print("2. Salir")

        opc = input("¿Cuál es la opción que eliges?: ").strip()

        if opc == "1":
            consultar_paciente()

        elif opc == "2":
            print("Saliendo del sistema...")
            break

        else:
            print("La opción no es correcta.")


def consultar_paciente():
    """Pide un ID, construye el grafo del paciente y ejecuta el algoritmo elegido."""
    print("\nIngreso a consultar un paciente...")

    paciente_id = input("¿Cuál es el ID del paciente que va a consultar?: ").strip().upper()

    if paciente_exist(paciente_id):

        p = obtener_paciente(pacientes, paciente_id)

        print("\nPaciente encontrado.")
        print("Información:", p)

        # Un costo por cada fila de la matriz (cada tratamiento), personalizado al paciente
        costos = [calcular_costo_arista(fila, p) for _, fila in matriz.iterrows()]
        # Costos "optimistas" (sin factores de riesgo): solo los usa la heurística de A*
        costos_optimistas = [calcular_costo_optimista(fila, p) for _, fila in matriz.iterrows()]
        G = construir_grafo(matriz, costos, costos_optimistas)

        print("\nNúmero de nodos:", G.number_of_nodes())
        print("Número de conexiones:", G.number_of_edges())

        print("\n¿Qué búsqueda quiere usar?")
        print("1. BFS (Nivel 1)")
        print("2. UCS (Nivel 2)")
        print("3. A* (Nivel 3)")

        alg = input("Opción: ").strip()

        if alg == "1":
            ruta, costo, nodos = bfs(G)
            mostrar_resultado("BFS", (ruta, costo, nodos))
            dibujar_grafo(G, f"Paciente {paciente_id} - BFS", ruta=ruta)

        elif alg == "2":
            ruta, costo, nodos = ucs(G)
            mostrar_resultado("UCS", (ruta, costo, nodos))
            dibujar_grafo(G, f"Paciente {paciente_id} - UCS", ruta=ruta)

        elif alg == "3":
            # La heurística depende de la severidad del paciente
            heuristica = construir_heuristica(G, p['Severidad_Diagnostico'])
            ruta, costo, nodos = a_estrella(G, heuristica)
            mostrar_resultado("A*", (ruta, costo, nodos))
            dibujar_grafo(G, f"Paciente {paciente_id} - A*", ruta=ruta)

        else:
            print("La opción no es correcta.")

    else:
        print("\nEl paciente no existe en la base de datos.")

    print("\nOperación finalizada.")


def mostrar_resultado(nombre, resultado):
    """Imprime la ruta, el costo y los nodos expandidos de un algoritmo.

    resultado es la tupla (ruta, costo, nodos) que devuelven bfs, ucs y a_estrella.
    """
    ruta, costo, nodos = resultado

    print(f"\n=== {nombre} ===")
    if ruta:
        print("Ruta:", " → ".join(ruta))
        print("Costo:", round(costo, 2))
        print("Nodos expandidos:", nodos)
    else:
        print("No se encontró ruta.")


def paciente_exist(paciente):
    """True si el ID está en la pestaña Pacientes del Excel."""
    return paciente in pacientes["ID_Paciente"].values


main()
