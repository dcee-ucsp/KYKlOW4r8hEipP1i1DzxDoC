import streamlit as st
from docx import Document
from docx.shared import Pt
from io import BytesIO
import time

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

# Función para reemplazar texto preservando formato
def reemplazar_texto(doc, marcador, nuevo_texto):
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            if marcador in run.text:
                run.text = run.text.replace(marcador, str(nuevo_texto))
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        if marcador in run.text:
                            run.text = run.text.replace(marcador, str(nuevo_texto))

# Función para establecer estilo de fuente
def set_font_style(doc):
    times_new_roman_font = 'Times New Roman'
    font_size = Pt(11)  # 11 point size
    
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

# Función visual de carga
def cook_breakfast():
    msg = st.toast('📜 Preparando el documento...')
    time.sleep(1)
    msg.toast('✍🏼 Reemplazando texto...')
    time.sleep(1)
    msg.toast('✅ Documento listo para descargar', icon="📄")

if st.button("Generar Documento"):
    # Simular proceso
    cook_breakfast()

    # Cargar documento base
    doc = Document("Plantillas/plantilla_adm.docx")

    # Reemplazar marcadores en el documento
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
    
    # Establecer estilo de fuente después de reemplazos
    set_font_style(doc)
    
    # Guardar en un buffer en memoria
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    # Descargar el archivo
    st.download_button(
        label="📄 Descargar Documento",
        data=buffer,
        file_name="documento_personalizado.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
