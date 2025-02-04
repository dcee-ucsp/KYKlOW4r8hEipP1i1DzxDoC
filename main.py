import streamlit as st
from docx import Document
from io import BytesIO

st.title("Carta de presentación")

# Ingresar texto
st.subheader("Datos generales", divider=True)
correlativo = st.number_input("Correlativo", step=1) 
fecha = st.date_input("Fecha de emisión", value="today")

st.subheader("Datos del alumno", divider=True)
nombre = st.text_input("Nombre del alumno")
dni_est = st.text_input("Ingrese el DNI del alumno")
genero_est = st.radio("Genero del estudiante",["Masculino", "Femenino"])
semestre_alumno = st.radio("Semestre",["septimo", "octavo", "noveno", "decimo", "egresado")

referencia = st.radio("Referencia",["Señor", "Señora", "Señorita", "Estimado", "Estimada"])
nombre_empleador = st.text_input("Nombre del empleador")
cargo_empleador = st.text_input("Cargo del empleador")

if st.button("Generar Documento"):
    # Cargar documento base
    doc = Document("plantilla.docx")
    
    # Reemplazar marcadores
    def reemplazar_texto(doc, marcador, nuevo_texto):
        for p in doc.paragraphs:
            if marcador in p.text:
                p.text = p.text.replace(marcador, nuevo_texto)

    reemplazar_texto(doc, "{{NOMBRE}}", nombre)
    reemplazar_texto(doc, "{{EVENTO}}", evento)

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
