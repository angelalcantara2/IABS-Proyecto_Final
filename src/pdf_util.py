"""pdf_util.py: convierte un .docx a PDF con Word (docx2pdf) o, si no está, con LibreOffice."""
import os
import glob
import shutil
import subprocess


def buscar_soffice():
    """Ruta de LibreOffice (PATH o carpetas típicas de Windows)."""
    for nom in ["soffice", "libreoffice"]:
        if shutil.which(nom):
            return shutil.which(nom)
    for patron in [r"C:\Program Files\LibreOffice\program\soffice.exe",
                   r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"]:
        hallados = glob.glob(patron)
        if hallados:
            return hallados[0]
    return None


def exportar_pdf(docx, pdf):
    """Primero intenta con Word (docx2pdf); si falla, con LibreOffice."""
    try:
        from docx2pdf import convert
        convert(docx, pdf)
        return "Word"
    except Exception as e:
        err_word = e
    soffice = buscar_soffice()
    if soffice:
        subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir",
                        os.path.dirname(pdf), docx], check=True,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return "LibreOffice"
    raise RuntimeError(f"No se pudo crear el PDF. Instala docx2pdf (pip install docx2pdf) "
                       f"con Word, o LibreOffice. Detalle: {err_word}")
