from lector import lectura_datos, obtener_paciente, ruta_archivo
from grafo import construir_grafo, dibujar_grafo, calcular_costo_arista
from busqueda import bfs, ucs

pacientes, matriz = lectura_datos(ruta_archivo)


def main():
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
    print("\nIngreso a consultar un paciente...")

    paciente_id = input("¿Cuál es el ID del paciente que va a consultar?: ").strip().upper()

    if paciente_exist(paciente_id):

        p = obtener_paciente(pacientes, paciente_id)

        print("\nPaciente encontrado.")
        print("Información:", p)

        costos = [calcular_costo_arista(fila, p) for _, fila in matriz.iterrows()]
        G = construir_grafo(matriz, costos)

        print("\nNúmero de nodos:", G.number_of_nodes())
        print("Número de conexiones:", G.number_of_edges())

        print("\n¿Qué búsqueda quiere usar?")
        print("1. BFS (Nivel 1)")
        print("2. UCS (Nivel 2)")

        alg = input("Opción: ").strip()

        if alg == "1":
            ruta, costo, nodos = bfs(G)
            mostrar_resultado("BFS", (ruta, costo, nodos))
            dibujar_grafo(G, f"Paciente {paciente_id} - BFS")

        elif alg == "2":
            ruta, costo, nodos = ucs(G)
            mostrar_resultado("UCS", (ruta, costo, nodos))
            dibujar_grafo(G, f"Paciente {paciente_id} - UCS")

        else:
            print("La opción no es correcta.")

    else:
        print("\nEl paciente no existe en la base de datos.")

    print("\nOperación finalizada.")


def mostrar_resultado(nombre, resultado):
    ruta, costo, nodos = resultado

    print(f"\n=== {nombre} ===")
    if ruta:
        print("Ruta:", " → ".join(ruta))
        print("Costo:", round(costo, 2))
        print("Nodos expandidos:", nodos)
    else:
        print("No se encontró ruta.")


def paciente_exist(paciente):
    return paciente in pacientes["ID_Paciente"].values


main()