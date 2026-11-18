from lector import lectura_datos, obtener_paciente, ruta_archivo
from grafo import construir_grafo, dibujar_grafo
from busqueda import bfs

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

        costos = matriz["Costo_Financiero_Base"]
        G = construir_grafo(matriz, costos)

        print("\n¿Qué búsqueda quiere usar?")
        print("1. BFS (Nivel 1)")

        alg = input("Opción: ").strip()

        if alg == "1":
            mostrar_resultado("BFS", bfs(G))
        else:
            print("La opción no es correcta.")

        dibujar_grafo(G, f"Paciente {paciente_id}")

    else:
        print("\nEl paciente no existe en la base de datos.")

    print("\nOperación finalizada.")


def mostrar_resultado(nombre, resultado):
    ruta, costo, nodos = resultado

    print(f"\n=== {nombre} ===")
    if ruta:
        print("Ruta:", " → ".join(ruta))
        print("Costo:", costo)
        print("Nodos expandidos:", nodos)
    else:
        print("No se encontró ruta.")


def paciente_exist(paciente):
    return paciente in pacientes["ID_Paciente"].values


main()