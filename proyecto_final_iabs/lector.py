import os
import pandas as pd


directorio_actual = os.path.dirname(os.path.abspath(__file__))
nombre_archivo = "datos_clinicos.xlsx"
ruta_archivo = os.path.join(directorio_actual, nombre_archivo)

def lectura_datos(path):
    pacientes = pd.read_excel(path, sheet_name='Pacientes')
    matriz = pd.read_excel(path, sheet_name='Matriz_Tratamientos')
    return pacientes, matriz

def obtener_paciente(pacientes, id_paciente):
    return pacientes[pacientes['ID_Paciente'] == id_paciente].iloc[0]

print(f"Leyendo datos desde: {ruta_archivo}")
pacientes, matriz = lectura_datos(ruta_archivo)