# IABS-Proyecto_Final
Sistema de Soporte a Decisiones Clínicas (SSDC): calcula la ruta de tratamiento con BFS, UCS y A*.

## Instalar (una sola vez)
Python 3.8 o superior.

    pip install -r requirements.txt

Para el PDF se usa Microsoft Word (paquete docx2pdf) o, si no hay Word, LibreOffice.
Sin ninguno de los dos el informe se genera igual en Word.

## Ejecutar
Desde la carpeta del proyecto:

    python src/main.py

- Botón **Reporte comparativo**: tabla de los pacientes con los tres algoritmos.
- Botón **Exportar informe APA (Word y PDF)**: genera `documentos/SSDC_Entregables_2_y_3.docx` y `.pdf`
  a partir de `datos/datos_clinicos.xlsx`. Si se agregan pacientes o tratamientos al Excel, el informe cambia solo.

## Carpetas
- `src/`: código.
- `datos/`: Excel de pacientes y tratamientos.
- `docs/`: plantilla del informe (no borrar: el programa la necesita).
- `documentos/`: informes generados.
- `resultados/`: tablas y figuras intermedias (se crea sola).
