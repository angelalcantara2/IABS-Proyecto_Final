"""
lector.py  (Bloque 1: lectura de datos)
---------------------------------------
Abre el Excel con los datos clínicos y entrega los datos como tablas de pandas.
No calcula nada ni toma decisiones: solo lee.

El Excel debe tener dos pestañas con estos nombres exactos:
    - 'Pacientes'            -> 12 pacientes (factores de riesgo, preferencia, severidad)
    - 'Matriz_Tratamientos'  -> 8 tratamientos (las aristas del grafo base)
"""
import os
import pandas as pd


# El Excel está en la carpeta datos/, al lado de src/. La ruta se arma a partir
# de la ubicación de este archivo, así el programa
# funciona sin importar desde dónde se ejecute (consola, VS Code, otro computador).
directorio_actual = os.path.dirname(os.path.abspath(__file__))
carpeta_datos = os.path.join(os.path.dirname(directorio_actual), "datos")
nombre_archivo = "datos_clinicos.xlsx"
ruta_archivo = os.path.join(carpeta_datos, nombre_archivo)


def lectura_datos(path):
    """Lee las dos pestañas del Excel.

    Parámetros:
        path: ruta del archivo .xlsx (puede ser otro Excel con el mismo formato).
    Retorna:
        (pacientes, matriz): dos DataFrames de pandas.
    """
    pacientes = pd.read_excel(path, sheet_name='Pacientes')
    matriz = pd.read_excel(path, sheet_name='Matriz_Tratamientos')
    return pacientes, matriz


def obtener_paciente(pacientes, id_paciente):
    """Devuelve la fila del paciente con ese ID (por ejemplo 'P001').

    Si el ID no existe lanza IndexError; por eso main.py verifica antes
    con paciente_exist().
    """
    return pacientes[pacientes['ID_Paciente'] == id_paciente].iloc[0]


# Estas líneas se ejecutan al importar el módulo: informan de dónde se lee
# el Excel y dejan 'pacientes' y 'matriz' disponibles.
print(f"Leyendo datos desde: {ruta_archivo}")
pacientes, matriz = lectura_datos(ruta_archivo)
