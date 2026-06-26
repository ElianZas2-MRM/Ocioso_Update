# -*- coding: utf-8 -*-
"""Actualiza la sección 5 de FINAL3.docx con el tracking PasoN::campo."""
from docx import Document
from docx.oxml.ns import qn
import copy

doc = Document('Arquitectura_Consolidada_ Osocio Form Automation_FINAL3.docx')
body = doc.element.body

def get_text(elem):
    return ''.join(r.text for r in elem.iter(qn('w:t')) if r.text)

def get_style(elem):
    pPr = elem.find(qn('w:pPr'))
    if pPr is None: return ''
    pStyle = pPr.find(qn('w:pStyle'))
    return pStyle.get(qn('w:val'), '') if pStyle is not None else ''

# Encontrar inicio y fin de sección 5
elems = list(body)
sec5_start = None
sec5_end = None
for i, e in enumerate(elems):
    if e.tag != qn('w:p'): continue
    txt = get_text(e)
    sty = get_style(e)
    if sec5_start is None and '5.' in txt and 'Tracking' in txt and 'Heading2' in sty:
        sec5_start = i
    elif sec5_start is not None and 'Heading2' in sty:
        sec5_end = i
        break

print(f'Sección 5: párrafos [{sec5_start}:{sec5_end}]')

# Eliminar párrafos del cuerpo de la sección (todo menos el heading)
for e in elems[sec5_start + 1 : sec5_end]:
    if e.tag != qn('w:sectPr'):
        body.remove(e)

# Insertar nuevo contenido antes del heading de la sección 6
anchor = list(body)[sec5_start + 1]

def insert_before(anchor_elem, new_elem):
    anchor_elem.addprevious(new_elem)

def make_para(text, style='Normal'):
    p = doc.add_paragraph()
    p.style = doc.styles[style]
    p.text = text
    body.remove(p._element)
    return p._element

def make_bullet(text):
    return make_para('• ' + text)

lines = [
    make_para('Archivo: core/base_form_filler.py — método write_tracked_fields_to_sheet()'),
    make_para('A medida que el motor llena cada campo del formulario, registra internamente el par (campo, valor efectivamente usado) mediante _record_field_value(). Al finalizar cada fila, write_tracked_fields_to_sheet() escribe todos esos valores en el Excel de resultados.'),
    make_para('Tracking por paso del formulario:'),
    make_bullet('Cada campo registrado lleva el prefijo del paso en el que fue completado. El formato de la clave es PasoN::nombre_campo, donde N es el número de paso (empieza en 1) y nombre_campo es el nombre legible del campo según ids_dinamicos.json, o el ID HTML si no tiene nombre configurado.'),
    make_bullet('El número de paso se mantiene en self._current_step, que se inicializa en 1 al comienzo de cada fila (begin_row_tracking()) y se incrementa automáticamente cada vez que fill_form_fields_auto_step() detecta y presiona un botón "Siguiente/Seguinte/Next/Continuar".'),
    make_bullet('Para formularios de un solo paso, todos los campos quedan bajo Paso1::*. Para un formulario de 3 pasos, los campos del paso 1 quedan como Paso1::models, Paso1::year; los del paso 2 como Paso2::firstname, Paso2::email; y los del paso 3 como Paso3::region, Paso3::city, Paso3::dealer.'),
    make_para('Escritura en Excel:'),
    make_bullet('Para los campos que están en el field_mapping del país: se crea una columna con el nombre PasoN::nombre_campo. Si la columna ya existe (de filas anteriores), se reutiliza. Si no existe, se crea dinámicamente en la primera fila del Excel de resultados.'),
    make_bullet('Para los campos que NO estaban en el mapping (campos dinámicos detectados y rellenados automáticamente): el formato de columna es PasoN::nombre_campo, donde nombre_campo se resuelve desde ids_dinamicos.json o se usa el ID HTML directamente.'),
    make_bullet('Esto permite ver exactamente en qué paso del formulario se completó cada campo, lo que es especialmente útil para detectar si un formulario cambió su estructura de pasos entre ejecuciones (por ejemplo, pasó de 2 a 3 pasos).'),
    make_para('Ejemplo de columnas en el Excel de resultados para un formulario de 3 pasos:'),
    make_bullet('Paso1::Modelo, Paso1::Año de adquisición'),
    make_bullet('Paso2::Nombre, Paso2::Apellido, Paso2::Documento, Paso2::Celular, Paso2::Email'),
    make_bullet('Paso3::Región, Paso3::Ciudad, Paso3::Concesionario, Paso3::Fecha estimada'),
]

for line in reversed(lines):
    insert_before(anchor, line)

doc.save('Arquitectura_Consolidada_ Osocio Form Automation_FINAL3b.docx')
print('Guardado como FINAL3b.docx')
