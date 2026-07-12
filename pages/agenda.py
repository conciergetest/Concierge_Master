import streamlit as st
import sqlite3
import pandas as pd
import os

# --- CORRECCIÓN DE RUTA ---
# Usamos os.path.dirname(os.path.dirname(...)) para subir de 'pages/' a la raíz
db_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "recepcion_final.db")

CLAVE_SEGURIDAD = "D6msnp8a" 

# --- FUNCIONES ---
def get_connection():
    return sqlite3.connect(db_path)

def cargar_datos():
    conn = get_connection()
    # Verificamos que la tabla exista antes de leer
    try:
        df = pd.read_sql_query("SELECT * FROM agenda_interna", conn)
        conn.close()
        if not df.empty:
            df = df.sort_values(by=['departamento', 'nombre'])
        return df
    except Exception as e:
        conn.close()
        st.error(f"Error al conectar con la base de datos: {e}")
        return pd.DataFrame()

def insertar_contacto(n, p, d, e, em):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("INSERT INTO agenda_interna (nombre, puesto, departamento, extension, email) VALUES (?,?,?,?,?)",
                   (n, p, d, e, em))
    conn.commit()
    conn.close()

# --- ESTADO DE SESIÓN ---
if "abrir_form" not in st.session_state:
    st.session_state.abrir_form = False
if "pedir_clave" not in st.session_state:
    st.session_state.pedir_clave = False

# Botón para volver al Dashboard
st.page_link("app.py", label="⬅️ Volver al Dashboard", icon="🏠")

st.title("🛎️ Concierge Master 4.8 - Agenda")

# Botón para activar el formulario
if st.button("➕ Agregar nuevo contacto"):
    st.session_state.abrir_form = True

# Lógica del formulario
if st.session_state.abrir_form:
    st.divider()
    with st.container(border=True):
        st.subheader("📝 Nuevo Contacto")
        with st.form("form_agregar", clear_on_submit=True):
            n = st.text_input("Nombre")
            p = st.text_input("Puesto")
            d = st.text_input("Departamento")
            e = st.text_input("Extensión")
            em = st.text_input("Email")
            
            c1, c2 = st.columns([1, 10])
            with c1:
                submit = st.form_submit_button("Guardar contacto")
            with c2:
                cancelar = st.form_submit_button("Cancelar")
            
            if submit:
                insertar_contacto(n, p, d, e, em)
                st.session_state.abrir_form = False
                st.rerun()
            if cancelar:
                st.session_state.abrir_form = False
                st.rerun()

# --- TABLA Y PROTECCIÓN DE BORRADO ---
st.subheader("Directorio de Contactos")
df = cargar_datos()

if not df.empty:
    column_config = {"id": st.column_config.NumberColumn(disabled=True)}
    edited_df = st.data_editor(df, use_container_width=True, hide_index=True, column_config=column_config, num_rows="dynamic")

    # Protección de guardado con clave
    if st.button("💾 Guardar cambios"):
        st.session_state.pedir_clave = True

    if st.session_state.pedir_clave:
        clave_ingresada = st.text_input("Ingresa la clave para confirmar:", type="password")
        if clave_ingresada == CLAVE_SEGURIDAD:
            conn = get_connection()
            edited_df.to_sql("agenda_interna", conn, if_exists="replace", index=False)
            conn.close()
            st.session_state.pedir_clave = False
            st.success("Cambios aplicados.")
            st.rerun()
        elif clave_ingresada:
            st.error("❌ Clave incorrecta.")
else:
    st.info("No hay contactos en la agenda.")
