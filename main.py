import streamlit as st
from docx import Document
from docx.shared import Pt
from io import BytesIO
import time
import random
from docx2pdf import convert
import os
import tempfile

st.title("Carta de presentación")

# Ingresar texto
st.subheader("Datos generales", divider=True)
correlativo = st.number_input("Correlativo", step=1) 
fecha = st.date_input("Fecha de emisión")
tipo_practicas = st.radio("Tipo de prácticas", ["Pre-profesionales", "Profesionales"])

st.subheader("Datos del alumno", divider=True)
nombre = st.text_input("Nombre del alumno")
dni_est = st.text_input("Ingrese el DNI del alumno")
genero_est = st.radio("Género del estudiante", ["Masculino", "Femenino"])
semestre_alumno = st.selectbox("Semestre", ("séptimo", "octavo", "noveno", "décimo", "egresado"))
periodo_pract = st.slider("Periodo de prácticas (meses)", 1, 12)

st.subheader("Datos del empleador", divider=True)
nombre_empresa = st.text_input("Nombre de la empresa")
referencia = st.selectbox("Referencia", ("Señor", "Señora", "Señorita", "Estimado", "Estimada"))
nombre_empleador = st.text_input("Nombre del empleador")
cargo_empleador = st.text_input("Cargo del empleador")

# Conversión de fecha
meses = {
    1: "enero", 2: "febrero", 3: "marzo", 4: "abril", 5: "mayo", 6: "junio", 
    7: "julio", 8: "agosto", 9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
}
fecha_larga = f"{fecha.day} de {meses[fecha.month]} del {fecha.year}"

# Conversión género
gen_alumn = "el alumno" if genero_est == "Masculino" else "la alumna"

# Semestre del alumno
reemplazo_semestre = f"del {semestre_alumno}" if semestre_alumno != "egresado" else "egresado"

# Texto meses
meses_texto = {
    1: "un mes", 2: "dos meses", 3: "tres meses", 4: "cuatro meses", 
    5: "cinco meses", 6: "seis meses", 7: "siete meses", 8: "ocho meses", 
    9: "nueve meses", 10: "diez meses", 11: "once meses", 12: "doce meses"
}
periodo_pract_texto = meses_texto[periodo_pract]

def get_random_step():
    steps = [
        "Tomando un café ☕",
        "Visitando la cafetería 🍽️",
        "Lavando los platos 🍽️",
        "Regando las plantas 🌱",
        "Ordenando el escritorio 📚",
        "Alimentando al gato 🐱",
        "Haciendo ejercicio 🏃",
        "Meditando un momento 🧘",
        "Revisando el correo 📧",
        "Estirando los brazos 💪"
    ]
    return random.choice(steps)

def cook_breakfast():
    msg = st.toast('📜 Preparando el documento...')
    time.sleep(1)
    msg.toast(get_random_step())
    time.sleep(1)
    msg.toast('✅ Documento listo para descargar', icon="📄")

# Formateo

def reemplazar_texto(doc, marcador, nuevo_texto):
    for paragraph in doc.paragraphs:
        if marcador in paragraph.text:
            paragraph.text = paragraph.text.replace(marcador, str(nuevo_texto))
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if marcador in paragraph.text:
                        paragraph.text = paragraph.text.replace(marcador, str(nuevo_texto))

def set_font_style(doc):
    times_new_roman_font = 'Times New Roman'
    font_size = Pt(11)
    
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            run.font.name = times_new_roman_font
            run.font.size = font_size
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.name = times_new_roman_font
                        run.font.size = font_size

def generate_filename(correlativo, year, nombre_alumno, nombre_empresa):
    nombre_alumno = nombre_alumno.split()[0]
    nombre_empresa = nombre_empresa.split()[0]
    return f"DIRADM-{correlativo}-{year}-{nombre_alumno}-{nombre_empresa}"

# Reemplazo de valores

if st.button("Generar Documento"):
    cook_breakfast()

    doc = Document("Plantillas/plantilla_adm.docx")

    reemplazar_texto(doc, "{{CORRELATIVO}}", correlativo)
    reemplazar_texto(doc, "{{ANIO}}", fecha.year)
    reemplazar_texto(doc, "{{FECHA_LARGA}}", fecha_larga)
    reemplazar_texto(doc, "{{REFERENCIA}}", referencia)
    reemplazar_texto(doc, "{{JEFE_DIRECTO}}", nombre_empleador)
    reemplazar_texto(doc, "{{CARGO_JEFE}}", cargo_empleador)
    reemplazar_texto(doc, "{{NOMBRE_EMPRESA}}", nombre_empresa)
    reemplazar_texto(doc, "{{GEN_ALUM}}", gen_alumn)
    reemplazar_texto(doc, "{{SEM_ALUM}}", reemplazo_semestre)
    reemplazar_texto(doc, "{{NOMBRE_ALUMNO}}", nombre)
    reemplazar_texto(doc, "{{DNI_ALUMNO}}", dni_est)
    reemplazar_texto(doc, "{{TIPO_PRACTICAS}}", tipo_practicas)
    reemplazar_texto(doc, "{{PERIODO_MESES}}", periodo_pract_texto)
    
    set_font_style(doc)
    
    base_filename = generate_filename(correlativo, fecha.year, nombre, nombre_empresa)
    
    col1, col2 = st.columns(2)
    
    buffer_docx = BytesIO()
    doc.save(buffer_docx)
    buffer_docx.seek(0)

# Botones de descarga
    
    with col1:
        st.download_button(
            label="📄 Descargar DOCX",
            data=buffer_docx,
            file_name=f"{base_filename}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    
    with col2:
        with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp_docx:
            doc.save(tmp_docx.name)
            pdf_path = tmp_docx.name.replace('.docx', '.pdf')
            convert(tmp_docx.name, pdf_path)
            
            with open(pdf_path, 'rb') as pdf_file:
                pdf_data = pdf_file.read()
                st.download_button(
                    label="📑 Descargar PDF",
                    data=pdf_data,
                    file_name=f"{base_filename}.pdf",
                    mime="application/pdf"
                )
            
            os.unlink(tmp_docx.name)
            os.unlink(pdf_path)
