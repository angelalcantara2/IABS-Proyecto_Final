"""
apa.py: da formato APA 7 (trabajo de estudiante) al informe que arma armar_informe.py.

- Portada: título en negrita, autores, institución, curso, docente y fecha.
- Times New Roman 12, doble espacio, márgenes de 1", sangría de 0,5" y número de página arriba a la derecha.
- Encabezados sin numeración: nivel 1 centrado en negrita, nivel 2 a la izquierda en negrita,
  nivel 3 en negrita cursiva y nivel 4 en la misma línea del párrafo.
- Tablas y figuras con número en negrita, título en cursiva y nota; tablas solo con líneas horizontales.
"""
import re
import copy
from docx.shared import Pt, Inches, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.table import Table

FUENTE = "Times New Roman"
SANGRIA = Inches(0.5)


# ---------- utilidades ----------
def estilo(doc, sid):
    """Busca un estilo por su id (la plantilla usa 'Heading 1' con mayúscula)."""
    return next(st for st in doc.styles if st.style_id == sid)


def fuente(run, tam=12):
    run.font.name = FUENTE
    rpr = run._r.get_or_add_rPr()
    rf = rpr.get_or_add_rFonts()
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf.set(qn(a), FUENTE)
    run.font.size = Pt(tam)
    run.font.color.rgb = RGBColor(0, 0, 0)


def formato(p, sangria=True, centro=False, interlineado=2.0):
    f = p.paragraph_format
    f.line_spacing = interlineado
    f.space_before = Pt(0)
    f.space_after = Pt(0)
    f.left_indent = Inches(0)
    f.first_line_indent = SANGRIA if sangria else Inches(0)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if centro else WD_ALIGN_PARAGRAPH.LEFT
    for r in p.runs:
        fuente(r)


def par_antes(el, texto="", negrita=False, cursiva=False):
    """Inserta un párrafo nuevo antes del elemento 'el' (w:p o w:tbl)."""
    p = OxmlElement("w:p")
    el.addprevious(p)
    par = Paragraph(p, None)
    if texto:
        r = par.add_run(texto)
        r.bold, r.italic = negrita, cursiva
    return par


def par_despues(el, texto="", negrita=False, cursiva=False):
    p = OxmlElement("w:p")
    el.addnext(p)
    par = Paragraph(p, None)
    if texto:
        r = par.add_run(texto)
        r.bold, r.italic = negrita, cursiva
    return par


def salto_antes(p, si=True):
    p.paragraph_format.page_break_before = si


def tiene_imagen(p):
    return bool(p._p.findall(".//" + qn("w:drawing")))


def borde(padre, lado, val="single", sz=8):
    b = padre.find(qn(f"w:{lado}"))
    if b is None:
        b = OxmlElement(f"w:{lado}")
        padre.append(b)
    b.set(qn("w:val"), val)
    if val != "nil":
        b.set(qn("w:sz"), str(sz))
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), "000000")


def tabla_apa(t):
    """Solo líneas horizontales: arriba, bajo el encabezado y al final."""
    tblPr = t._tbl.tblPr
    viejo = tblPr.find(qn("w:tblBorders"))
    if viejo is not None:
        tblPr.remove(viejo)
    bs = OxmlElement("w:tblBorders")
    for lado in ("top", "left", "bottom", "right", "insideH", "insideV"):
        borde(bs, lado, "single" if lado in ("top", "bottom") else "nil")
    tblPr.append(bs)
    for i, fila in enumerate(t.rows):
        for c in fila.cells:
            tcPr = c._tc.get_or_add_tcPr()
            for tag in ("w:tcBorders", "w:shd"):
                x = tcPr.find(qn(tag))
                if x is not None:
                    tcPr.remove(x)
            if i == 0:
                tb = OxmlElement("w:tcBorders")
                borde(tb, "bottom")
                tcPr.append(tb)
            for p in c.paragraphs:
                formato(p, sangria=False, interlineado=1.0)
                for r in p.runs:
                    fuente(r, 10)
                    if i == 0:
                        r.bold = True


def campo_pagina(p):
    r = p.add_run()
    for tipo, txt in (("begin", None), (None, "PAGE"), ("end", None)):
        if tipo:
            fc = OxmlElement("w:fldChar")
            fc.set(qn("w:fldCharType"), tipo)
            r._r.append(fc)
        else:
            it = OxmlElement("w:instrText")
            it.set(qn("xml:space"), "preserve")
            it.text = txt
            r._r.append(it)
    fuente(r)


# ---------- principal ----------
def aplicar_apa(doc, titulos_tablas, figuras, ref, info):
    cuerpo = doc.element.body

    # 1. Página, estilos, encabezado y pie
    for sec in doc.sections:
        sec.page_width, sec.page_height = Emu(7772400), Emu(10058400)  # carta
        for m in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
            setattr(sec, m, Inches(1))
        sec.different_first_page_header_footer = False
        for p in sec.footer.paragraphs:
            for ch in list(p._p):
                if ch.tag != qn("w:pPr"):
                    p._p.remove(ch)
        hp = sec.header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        campo_pagina(hp)
    rpr_def = doc.styles.element.find(qn("w:docDefaults")).find(qn("w:rPrDefault")).find(qn("w:rPr"))
    rf = rpr_def.find(qn("w:rFonts"))
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf.set(qn(a), FUENTE)
    for tag in ("w:sz", "w:szCs"):
        rpr_def.find(qn(tag)).set(qn("w:val"), "24")
    for sid in ("Normal", "Heading1", "Heading2"):
        st = estilo(doc, sid)
        st.font.name, st.font.size = FUENTE, Pt(12)
        st.font.color.rgb = RGBColor(0, 0, 0)
        st.paragraph_format.space_before = Pt(0)
        st.paragraph_format.space_after = Pt(0)
        st.paragraph_format.line_spacing = 2.0

    # 2. Quitar la portada de la plantilla y los párrafos vacíos
    pars = doc.paragraphs
    i_a = next(i for i, p in enumerate(pars) if p.text.startswith("Parte A"))
    for p in pars[:i_a]:
        cuerpo.remove(p._p)
    for p in doc.paragraphs:
        if not p.text.strip() and not tiene_imagen(p):
            p._p.getparent().remove(p._p)

    # 3. Clasificar párrafos y armar encabezados APA
    figs = {q._p for grupo in figuras for q in grupo[:3] if q is not None}
    notas_fig = {g[3]._p for g in figuras if g[3] is not None}
    tipo = {}  # elemento -> clase
    for p in doc.paragraphs:
        est = p.style.name
        if est == "Heading 1":
            if p.text.startswith("Anexo"):
                tit = p.text.split(".", 1)[1].strip()
                p.runs[0].text = "Anexo A"
                for r in p.runs[1:]:
                    r.text = ""
                q = par_despues(p._p, tit, negrita=True)
                tipo[q._p] = "n1"
                salto_antes(p)
            tipo[p._p] = "n1"
        elif est == "Heading 2":
            txt = re.sub(r"^(?:[A-Z]\.\d+|\d+\.)\s*", "", p.text).replace(" (opcional)", "")
            p.runs[0].text = txt
            for r in p.runs[1:]:
                r.text = ""
            if txt == "Referencias":
                p.style = estilo(doc, "Heading1")
                salto_antes(p)
                tipo[p._p] = "n1"
            else:
                tipo[p._p] = "n2"
        elif p._p in figs:
            tipo[p._p] = "figura"
        elif p._p in notas_fig:
            tipo[p._p] = "nota"
        elif p.runs and p.runs[0].bold and (re.match(r"^\d\.\d\s", p.runs[0].text)
                                            or p.runs[0].text.strip() == "Simulación de apoyo"):
            r0 = p.runs[0]
            r0.text = re.sub(r"^\d\.\d\s+", "", r0.text)
            if p.text.strip() == r0.text.strip():  # solo el rótulo -> nivel 3
                r0.italic = True
                tipo[p._p] = "n3"
            else:  # rótulo en línea -> nivel 4
                lab = r0.text.rstrip()
                lab = lab[:-1] + "." if lab.endswith(":") else lab
                r0.text = lab + " "
                tipo[p._p] = "cuerpo"
        else:
            tipo[p._p] = "cuerpo"

    # 4. Declaración de uso de IA antes del anexo
    h_anx = next(p for p in doc.paragraphs if p.text.startswith("Anexo A"))
    h_ia = par_antes(h_anx._p, "Declaración de uso de inteligencia artificial", negrita=True)
    p_ia = par_antes(h_anx._p, info["decl_ia"])
    tipo[h_ia._p], tipo[p_ia._p] = "n2", "cuerpo"

    # 6. Tablas: número, título, nota y bordes
    n, en_anexo = 0, False
    for el in list(cuerpo.iterchildren()):
        if el.tag == qn("w:p") and "Anexo A" in "".join(t.text or "" for t in el.iter(qn("w:t"))):
            en_anexo, n = True, 0
        if el.tag != qn("w:tbl"):
            continue
        n += 1
        t = Table(el, doc._body)
        tit, nota = titulos_tablas.get(el, ("", None))
        etiqueta = f"Tabla A{n}" if en_anexo else f"Tabla {n}"
        a = par_antes(el, etiqueta, negrita=True)
        b = par_antes(el, tit, cursiva=True)
        tipo[a._p] = tipo[b._p] = "caption"
        a.paragraph_format.keep_with_next = b.paragraph_format.keep_with_next = True
        if nota:
            c = par_despues(el, "")
            c.add_run("Nota. ").italic = True
            c.add_run(nota)
            tipo[c._p] = "nota"
        tabla_apa(t)

    # 7. Formato de cada párrafo según su clase
    for p in doc.paragraphs:
        k = tipo.get(p._p, "cuerpo")
        if k == "n1":
            formato(p, sangria=False, centro=True)
            for r in p.runs:
                r.bold = True
            p.paragraph_format.keep_with_next = True
        elif k in ("n2", "n3"):
            formato(p, sangria=False)
            for r in p.runs:
                r.bold = True
                if k == "n3":
                    r.italic = True
            p.paragraph_format.keep_with_next = True
        elif k == "caption":
            formato(p, sangria=False)
        elif k == "figura":
            img = tiene_imagen(p)
            formato(p, sangria=False, centro=img, interlineado=1.0 if img else 2.0)
            p.paragraph_format.keep_with_next = not img
        elif k == "nota":
            formato(p, sangria=False)
        elif k == "ref":
            formato(p, sangria=False)
            p.paragraph_format.left_indent = SANGRIA
            p.paragraph_format.first_line_indent = -SANGRIA
        else:
            formato(p)

    # 8. Portada y título en la primera página del texto
    h_a = next(p for p in doc.paragraphs if p.text.startswith("Parte A"))
    autores = info["autores"]
    txt_aut = ", ".join(autores[:-1]) + " y " + autores[-1] if len(autores) > 1 else autores[0]
    inst = ", ".join(x for x in (info["programa"], info["universidad"]) if x)
    portada = ["", "", "", (info["titulo"], True), "", (txt_aut, False), (inst, False),
               (info["curso"], False), (info["docente"], False), (info["fecha"], False)]
    portada = [x for x in portada if x == "" or x[0]]  # omite datos vacíos
    for item in portada:
        if item == "":
            q = par_antes(h_a._p, "")
        else:
            q = par_antes(h_a._p, item[0], negrita=item[1])
        formato(q, sangria=False, centro=True)
    t1 = par_antes(h_a._p, info["titulo"], negrita=True)
    formato(t1, sangria=False, centro=True)
    salto_antes(t1)
    salto_antes(h_a, False)
