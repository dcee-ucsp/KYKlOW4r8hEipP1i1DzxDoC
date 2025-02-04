import streamlit as st
from docx import Document
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

# Función para reemplazar texto en párrafos y tablas
def reemplazar_texto(doc, marcador, nuevo_texto):
    for p in doc.paragraphs:
        if marcador in p.text:
            p.text = p.text.replace(marcador, str(nuevo_texto))  # Convertir a string por seguridad

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if marcador in cell.text:
                    cell.text = cell.text.replace(marcador, str(nuevo_texto))

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
    reemplazar_texto(doc, "{{TIPO_PRACTICAS}}", dni_est)
    

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
