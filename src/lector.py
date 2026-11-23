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
    # filas vacías (por ejemplo, con formato pero sin datos) se descartan
    pacientes = pacientes.dropna(subset=['ID_Paciente']).reset_index(drop=True)
    matriz = matriz.dropna(subset=['Tratamiento']).reset_index(drop=True)
    # las filas vacías hacen que pandas lea las columnas 0/1 como decimales: se vuelven enteros
    for col in pacientes.columns:
        if col not in ('ID_Paciente', 'Nombre_Anonimizado'):
            pacientes[col] = pacientes[col].fillna(0).astype(int)
    return pacientes, matriz

def obtener_paciente(pacientes, id_paciente):
    """Devuelve la fila del paciente (lanza IndexError si no existe)."""
    return pacientes[pacientes['ID_Paciente'] == id_paciente].iloc[0]

print(f"Leyendo datos desde: {ruta_archivo}")
pacientes, matriz = lectura_datos(ruta_archivo)