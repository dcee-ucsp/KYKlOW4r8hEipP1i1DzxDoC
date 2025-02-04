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
semestre_alumno = st.selectbox("Semestre", ("séptimo", "octavo", "noveno", "décimo", "egresado"))

st.subheader("Datos del empleador", divider=True)
nombre_empresa = st.text_input("Nombre de la empresa")
referencia = st.selectbox("Referencia",("Señor", "Señora", "Señorita", "Estimado", "Estimada"),)
nombre_empleador = st.text_input("Nombre del empleador")
cargo_empleador = st.text_input("Cargo del empleador")

# Conversión de fecha
meses = {1: "enero", 2: "febrero", 3: "marzo", 4: "abril", 5: "mayo", 6: "junio", 7: "julio", 8: "agosto", 9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"}
fecha_larga = f"{fecha.day} de {meses[fecha.month]} del {fecha.year}"

# Conversión genero
gen_alumn = "el alumno" if genero_est == "Masculino" else "la alumna"

# Semestre del alumno
reemplazo_semestre = f"del {semestre_alumno}" if semestre_alumno != "egresado" else "egresado"


semestre_alumno

if st.button("Generar Documento"):
    # Cargar documento base
    doc = Document("plantilla.docx")
    
    # Reemplazar marcadores
    def reemplazar_texto(doc, marcador, nuevo_texto):
        for p in doc.paragraphs:
            if marcador in p.text:
                p.text = p.text.replace(marcador, nuevo_texto)

    def cook_breakfast():
        msg = st.toast('Gathering ingredients...')
        time.sleep(1)
        msg.toast('Cooking...')
        time.sleep(1)
        msg.toast('Ready!', icon = "🥞")

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

    #fecha

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
