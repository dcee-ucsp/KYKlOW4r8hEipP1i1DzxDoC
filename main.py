import streamlit as st
from docx import Document
from io import BytesIO

st.title("Generador de Documentos")

# Entrada de usuario
nombre = st.text_input("Ingrese su nombre:")
evento = st.text_input("Ingrese el nombre del evento:")

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
