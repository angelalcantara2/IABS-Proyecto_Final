"""
lector.py: lee el Excel de datos clínicos.
Pestañas requeridas: 'Pacientes' y 'Matriz_Tratamientos'.
"""
import os
import pandas as pd

# Ruta relativa a este archivo, para que funcione desde cualquier carpeta
directorio_actual = os.path.dirname(os.path.abspath(__file__))
nombre_archivo = "datos_clinicos.xlsx"
ruta_archivo = os.path.join(os.path.dirname(directorio_actual), "datos", nombre_archivo)

def lectura_datos(path):
    """Devuelve (pacientes, matriz) como DataFrames."""
    pacientes = pd.read_excel(path, sheet_name='Pacientes')
    matriz = pd.read_excel(path, sheet_name='Matriz_Tratamientos')
    return pacientes, matriz

def obtener_paciente(pacientes, id_paciente):
    """Devuelve la fila del paciente (lanza IndexError si no existe)."""
    return pacientes[pacientes['ID_Paciente'] == id_paciente].iloc[0]

print(f"Leyendo datos desde: {ruta_archivo}")
pacientes, matriz = lectura_datos(ruta_archivo)