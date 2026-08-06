import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta
from io import BytesIO
from supabase import create_client, Client

# ============================================================
# CONFIGURACIÓN DE PÁGINA Y CONEXIÓN SUPABASE
# ============================================================
st.set_page_config(page_title="Concierge Master", layout="wide", initial_sidebar_state="collapsed")

# ============================================================
# SPLASHSCREEN
# ============================================================
if "splash_shown" not in st.session_state:
    st.session_state["splash_shown"] = True

    splash_html = """
    <style>
        #splash-overlay {
            position: fixed;
            top: 0;
            left: 0;
            width: 100vw;
            height: 100vh;
            background-color: #0a0a0a;
            z-index: 999999;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            animation: splashFade 4.5s ease-in-out forwards;
        }
        #splash-overlay img {
            width: 100%;
            height: 100%;
            object-fit: cover;
            object-position: center;
            position: absolute;
            top: 0;
            left: 0;
        }
        #splash-content {
            position: relative;
            z-index: 2;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: flex-end;
            width: 100%;
            height: 100%;
            padding-bottom: 80px;
        }
        #progress-container {
            width: 320px;
            height: 3px;
            background-color: rgba(212, 175, 55, 0.15);
            border-radius: 3px;
            overflow: hidden;
            margin-top: 20px;
        }
        #progress-bar {
            width: 0%;
            height: 100%;
            background: linear-gradient(90deg, #d4af37, #f0d878, #d4af37);
            border-radius: 3px;
            animation: progressFill 3.5s ease-out forwards;
            box-shadow: 0 0 10px rgba(212, 175, 55, 0.5);
        }
        #loading-text {
            color: #d4af37;
            font-family: 'Segoe UI', sans-serif;
            font-size: 0.7rem;
            letter-spacing: 4px;
            text-transform: uppercase;
            margin-top: 15px;
            opacity: 0.8;
            animation: textPulse 1.5s ease-in-out infinite;
        }
        @keyframes progressFill {
            0% { width: 0%; }
            20% { width: 15%; }
            50% { width: 60%; }
            80% { width: 85%; }
            100% { width: 100%; }
        }
        @keyframes textPulse {
            0%, 100% { opacity: 0.5; }
            50% { opacity: 1; }
        }
        @keyframes splashFade {
            0% { opacity: 1; }
            78% { opacity: 1; }
            100% { opacity: 0; visibility: hidden; }
        }
        header[data-testid="stHeader"] { display: none !important; }
        .block-container { padding-top: 0 !important; }
    </style>
    <div id="splash-overlay">
        <div id="splash-content">
            <div id="progress-container">
                <div id="progress-bar"></div>
            </div>
            <div id="loading-text">Loading Concierge Master</div>
        </div>
    </div>
    """
    st.html(splash_html)

# ============================================================
# INICIALIZACIÓN SEGURA DE SUPABASE
# ============================================================
@st.cache_resource(show_spinner=False)
def init_supabase():
    """Inicializa la conexión a Supabase usando secrets de Streamlit."""
    try:
        url = st.secrets.get("SUPABASE_URL", "")
        key = st.secrets.get("SUPABASE_KEY", "")

        if not url or not key:
            return None, "Faltan SUPABASE_URL o SUPABASE_KEY en los secrets de Streamlit."

        # Normalizar URL
        url = url.strip()
        if not url.startswith("http"):
            url = "https://" + url

        client = create_client(url, key)

        # Test de conectividad ligero
        client.table("huespedes").select("id", count="exact").limit(1).execute()

        return client, None
    except Exception as e:
        return None, f"Error de conexión con Supabase: {str(e)}"

supabase, supabase_error = init_supabase()
TABLE_NAME = "huespedes"

# ============================================================
# PARÁMETROS DE URL
# ============================================================
query_params = st.query_params
mostrar_formulario = query_params.get("action") == "nueva"
mostrar_editar = query_params.get("action") == "editar"
mostrar_importar = query_params.get("action") == "importar"
mostrar_exportar = query_params.get("action") == "exportar"
filtro_checkout = query_params.get("checkout_filtro")
filtro_fecha_date = query_params.get("fecha_date")
fecha_filtro_activo = query_params.get("fecha_activa") == "true"

# ============================================================
# FUNCIONES CRUD CON SUPABASE (CON MANEJO DE ERRORES)
# ============================================================
@st.cache_data(ttl=30)
def cargar_reservaciones():
    """Carga todas las reservaciones desde Supabase."""
    if supabase is None:
        return pd.DataFrame(columns=["id", "eta", "name", "qty", "room", "email",
                                     "check_in", "check_out", "res_number", "phone",
                                     "info", "ird", "hsk", "rate", "trans"])
    try:
        response = supabase.table(TABLE_NAME).select("*").execute()
        df = pd.DataFrame(response.data)
        if df.empty:
            return pd.DataFrame(columns=["id", "eta", "name", "qty", "room", "email",
                                         "check_in", "check_out", "res_number", "phone",
                                         "info", "ird", "hsk", "rate", "trans"])
        # Ordenar por check_in y nombre
        df["check_in_dt"] = df["check_in"].apply(parse_fecha)
        df = df.sort_values(by=["check_in_dt", "name"])
        return df.drop(columns=["check_in_dt"])
    except Exception as e:
        st.error(f"Error al cargar reservaciones: {e}")
        return pd.DataFrame(columns=["id", "eta", "name", "qty", "room", "email",
                                     "check_in", "check_out", "res_number", "phone",
                                     "info", "ird", "hsk", "rate", "trans"])

def insertar_reserva(data: dict):
    """Inserta una nueva reserva en Supabase."""
    if supabase is None:
        st.error("No hay conexión a Supabase.")
        return
    try:
        supabase.table(TABLE_NAME).insert(data).execute()
        st.cache_data.clear()
        st.success("Reserva insertada correctamente.")
    except Exception as e:
        st.error(f"Error al insertar reserva: {e}")

def actualizar_reserva(reserva_id, data: dict):
    """Actualiza una reserva existente en Supabase."""
    if supabase is None:
        st.error("No hay conexión a Supabase.")
        return
    try:
        supabase.table(TABLE_NAME).update(data).eq("id", reserva_id).execute()
        st.cache_data.clear()
        st.success("Reserva actualizada correctamente.")
    except Exception as e:
        st.error(f"Error al actualizar reserva: {e}")

def eliminar_reserva(reserva_id):
    """Elimina una reserva de Supabase."""
    if supabase is None:
        st.error("No hay conexión a Supabase.")
        return
    try:
        supabase.table(TABLE_NAME).delete().eq("id", reserva_id).execute()
        st.cache_data.clear()
        st.success("Reserva eliminada correctamente.")
    except Exception as e:
        st.error(f"Error al eliminar reserva: {e}")

def insertar_batch_reservas(lista_data: list):
    """Inserta múltiples reservas en lote."""
    if supabase is None:
        st.error("No hay conexión a Supabase.")
        return
    try:
        supabase.table(TABLE_NAME).insert(lista_data).execute()
        st.cache_data.clear()
        st.success(f"{len(lista_data)} reservas importadas correctamente.")
    except Exception as e:
        st.error(f"Error al importar reservas: {e}")


def parse_fecha(fecha_str):
    """Parsea fecha en varios formatos para compatibilidad retroactiva.
    Acepta: 'July 14, 2026', 'Jul 14, 2026', 'July 14', 'Jul 14'
    """
    if not fecha_str or str(fecha_str).strip() == "":
        return None
    s = str(fecha_str).strip()
    for fmt in ["%B %d, %Y", "%b %d, %Y", "%B %d", "%b %d"]:
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            pass
    return None

# ============================================================
# HELPERS DE HORA
# ============================================================
horas_eta_12h, horas_eta_24h = [], []
for h in range(24):
    for m in [0, 30]:
        hora_24 = f"{h:02d}:{m:02d}"
        if h == 0: hora_12 = f"12:{m:02d} AM"
        elif h < 12: hora_12 = f"{h}:{m:02d} AM"
        elif h == 12: hora_12 = f"12:{m:02d} PM"
        else: hora_12 = f"{h-12}:{m:02d} PM"
        horas_eta_12h.append(hora_12)
        horas_eta_24h.append(hora_24)

mapa_12a24 = dict(zip(horas_eta_12h, horas_eta_24h))
mapa_24a12 = dict(zip(horas_eta_24h, horas_eta_12h))

def normalizar_hora_24(hora_str):
    if not hora_str or str(hora_str).strip() == "":
        return ""
    h = str(hora_str).strip()
    if ":" in h:
        partes = h.split(":")
        if len(partes) >= 2:
            hh = partes[0].zfill(2)
            mm = partes[1].zfill(2)
            return f"{hh}:{mm}"
    return h

def hora_24_a_12(hora_str):
    h_norm = normalizar_hora_24(hora_str)
    return mapa_24a12.get(h_norm, "")

def hora_actual_12h():
    ahora = datetime.now()
    h, m = ahora.hour, ahora.minute
    if m >= 45:
        m = 0
        h = (h + 1) % 24
    elif m >= 15:
        m = 30
    else:
        m = 0
    hora_24 = f"{h:02d}:{m:02d}"
    return mapa_24a12.get(hora_24, "12:00 AM")

# ============================================================
# EXPORTAR EXCEL
# ============================================================
def exportar_excel_por_categorias(df):
    try:
        from openpyxl import Workbook
        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    except ImportError:
        st.error("Falta la librería openpyxl. Instálala con: pip install openpyxl")
        return None

    fill_data = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    fill_header = PatternFill(start_color="00B0F0", end_color="00B0F0", fill_type="solid")
    fill_col_header = PatternFill(start_color="404040", end_color="404040", fill_type="solid")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_section = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
    font_data = Font(name="Calibri", size=10, color="000000")
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    thin_border = Border(left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"),
                         top=Side(style="thin", color="D9D9D9"), bottom=Side(style="thin", color="D9D9D9"))
    wb = Workbook()
    ws = wb.active
    ws.title = "Arrivals"
    categorias_config = [
        ("CUMPLEAÑOS", ["BIRTHDAY", "CUMPLE", "BDAY"]),
        ("VIP", ["VIP"]),
        ("HONEYMOON", ["HONEYMOON", "LUNA DE MIEL"]),
        ("ANNIVERSARY", ["ANNIVERSARY", "ANIVERSARIO"]),
        ("BABYMOON", ["BABYMOON"]),
        ("TEAM MEMBER", ["TEAM MEMBER", "STAFF", "EMPLOYEE"]),
        ("GENERAL", [])
    ]
    columnas_excel = ["ID", "ETA", "NAME", "QTY", "ROOM", "EMAIL", "CHECK IN", "CHECK OUT",
                      "RESERVATION", "PHONE", "FORMATIO", "IRD", "HSK", "RATE", "TRANSPORTATION"]
    mapeo_cols = {"id": "ID", "eta": "ETA", "name": "NAME", "qty": "QTY", "room": "ROOM",
                  "email": "EMAIL", "check_in": "CHECK IN", "check_out": "CHECK OUT",
                  "res_number": "RESERVATION", "phone": "PHONE", "info": "FORMATIO",
                  "ird": "IRD", "hsk": "HSK", "rate": "RATE", "trans": "TRANSPORTATION"}
    current_row = 1
    filas_por_categoria, filas_general = {}, []
    for _, row in df.iterrows():
        info_str = str(row.get("info", "")).upper()
        asignada = False
        for cat_nombre, keywords in categorias_config[:-1]:
            for kw in keywords:
                if kw in info_str:
                    filas_por_categoria.setdefault(cat_nombre, []).append(row)
                    asignada = True
                    break
            if asignada: break
        if not asignada: filas_general.append(row)
    for cat_nombre, keywords in categorias_config:
        filas = filas_general if cat_nombre == "GENERAL" else filas_por_categoria.get(cat_nombre, [])
        if not filas: continue
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=len(columnas_excel))
        cell = ws.cell(row=current_row, column=1, value=cat_nombre)
        cell.fill, cell.font, cell.alignment = fill_header, font_section, align_center
        current_row += 1
        for col_idx, col_name in enumerate(columnas_excel, 1):
            cell = ws.cell(row=current_row, column=col_idx, value=col_name)
            cell.fill, cell.font, cell.alignment, cell.border = fill_col_header, font_header, align_center, thin_border
        current_row += 1
        for row_data in filas:
            for col_idx, col_db in enumerate(mapeo_cols.keys(), 1):
                valor = row_data.get(col_db, "")
                if pd.isna(valor): valor = ""
                cell = ws.cell(row=current_row, column=col_idx, value=valor)
                cell.fill, cell.font, cell.alignment, cell.border = fill_data, font_data, align_left, thin_border
            current_row += 1
        current_row += 1
    anchos = {"A": 6, "B": 10, "C": 22, "D": 6, "E": 8, "F": 25, "G": 10, "H": 10,
              "I": 14, "J": 18, "K": 22, "L": 18, "M": 18, "N": 8, "O": 18}
    for col_letter, ancho in anchos.items(): ws.column_dimensions[col_letter].width = ancho
    ws.freeze_panes = "A1"
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# ============================================================
# CSS PERSONALIZADO
# ============================================================
st.markdown("""
<style>
header[data-testid="stHeader"] { display: none !important; }
.block-container { padding-top: 0.1rem !important; padding-bottom: 0.1rem !important; }
div[data-testid="stHorizontalBlock"] { gap: 0.2rem !important; }
div[data-testid="stVerticalBlock"] > div { margin-bottom: 0.1rem !important; }
div.stButton > button { width: 100%; border-radius: 6px; font-weight: bold; border: none;
    font-size: 0.65rem; padding: 4px 2px; white-space: nowrap; min-height: 28px; }
#root > div > div > div > div > div > div[data-testid="stHorizontalBlock"]:nth-of-type(1) > div:nth-child(2) button { background: #00E5FF !important; color: black !important; }
button[key^="checkout_btn_"] { background-color: #000000 !important; color: #ffffff !important; font-weight: bold !important; font-size: 0.7rem !important; border: 1px solid #333 !important; border-radius: 6px !important; text-align: center !important; padding: 2px 3px !important; min-height: 22px !important; margin: 0 !important; }
button[key="btn_ver_todas"] { background-color: #000000 !important; color: #ffffff !important; font-weight: bold !important; font-size: 0.7rem !important; border-radius: 6px !important; }
</style>
""", unsafe_allow_html=True)

# ============================================================
# UI PRINCIPAL
# ============================================================
# Encabezado
st.markdown("""
<div style="display:flex; justify-content:space-between; align-items:center; padding: 0.5rem 1rem; background:#0a0a0a;">
    <div style="color:#00E5FF; font-weight:bold; font-size:1.1rem;">Concierge Master v5.1</div>
    <div style="color:#888; font-size:0.8rem;">{}</div>
</div>
""".format(datetime.now().strftime("%B %d, %Y — %I:%M %p")), unsafe_allow_html=True)

# Alerta de conexión
if supabase_error:
    st.error(f"⚠️ {supabase_error}")
    st.info("💡 Verifica que hayas configurado correctamente los secrets SUPABASE_URL y SUPABASE_KEY en Streamlit Cloud (Settings > Secrets).")
    st.stop()

# Cargar datos
df_todas = cargar_reservaciones()

# Tabs de navegación
menu = ["📋 Vista General", "➕ Nueva Reserva", "📥 Importar", "📤 Exportar"]
if mostrar_formulario:
    idx_menu = 1
elif mostrar_importar:
    idx_menu = 2
elif mostrar_exportar:
    idx_menu = 3
else:
    idx_menu = 0

tabs = st.tabs(menu)

with tabs[0]:
    st.subheader("Reservaciones")

    # Filtros
    col_f1, col_f2, col_f3 = st.columns([2, 2, 2])
    with col_f1:
        filtro_nombre = st.text_input("🔍 Buscar por nombre", "")
    with col_f2:
        filtro_fecha = st.date_input("📅 Filtrar por check-in", value=None)
    with col_f3:
        opciones_estado = ["Todas"] + sorted(df_todas["check_out"].dropna().unique().tolist()) if not df_todas.empty else ["Todas"]
        filtro_estado = st.selectbox("🏷️ Estado checkout", opciones_estado)

    # Aplicar filtros
    df_filtrada = df_todas.copy()
    if filtro_nombre:
        df_filtrada = df_filtrada[df_filtrada["name"].astype(str).str.contains(filtro_nombre, case=False, na=False)]
    if filtro_fecha:
        fecha_str = filtro_fecha.strftime("%B %d, %Y")
        df_filtrada = df_filtrada[df_filtrada["check_in"].astype(str).str.contains(fecha_str[:fecha_str.rfind(",")] if "," in fecha_str else fecha_str, case=False, na=False)]
    if filtro_estado != "Todas":
        df_filtrada = df_filtrada[df_filtrada["check_out"].astype(str) == filtro_estado]

    st.write(f"Mostrando {len(df_filtrada)} de {len(df_todas)} reservas")

    if not df_filtrada.empty:
        st.dataframe(df_filtrada, use_container_width=True, hide_index=True)
    else:
        st.info("No se encontraron reservas con los filtros aplicados.")

    # Acciones rápidas
    c1, c2, c3 = st.columns(3)
    with c1:
        if st.button("➕ Nueva Reserva", key="btn_nueva"):
            st.query_params["action"] = "nueva"
            st.rerun()
    with c2:
        if st.button("📥 Importar CSV/Excel", key="btn_importar"):
            st.query_params["action"] = "importar"
            st.rerun()
    with c3:
        if st.button("📤 Exportar Excel", key="btn_exportar"):
            st.query_params["action"] = "exportar"
            st.rerun()

with tabs[1]:
    st.subheader("Nueva Reserva")
    with st.form("form_nueva"):
        col1, col2 = st.columns(2)
        with col1:
            name = st.text_input("Nombre *")
            qty = st.number_input("Cantidad", min_value=1, value=1)
            room = st.text_input("Habitación")
            email = st.text_input("Email")
            phone = st.text_input("Teléfono")
            res_number = st.text_input("Número de Reserva")
        with col2:
            check_in = st.text_input("Check-in (ej: July 14, 2026)")
            check_out = st.text_input("Check-out (ej: July 20, 2026)")
            eta = st.selectbox("ETA", horas_eta_12h, index=horas_eta_12h.index(hora_actual_12h()) if hora_actual_12h() in horas_eta_12h else 0)
            rate = st.text_input("Rate")
            trans = st.text_input("Transporte")
            info = st.text_area("Información / Notas")

        ird = st.text_input("IRD")
        hsk = st.text_input("HSK")

        submitted = st.form_submit_button("Guardar Reserva")
        if submitted:
            if not name:
                st.error("El nombre es obligatorio.")
            else:
                nueva_data = {
                    "name": name,
                    "qty": qty,
                    "room": room,
                    "email": email,
                    "phone": phone,
                    "res_number": res_number,
                    "check_in": check_in,
                    "check_out": check_out,
                    "eta": mapa_12a24.get(eta, ""),
                    "rate": rate,
                    "trans": trans,
                    "info": info,
                    "ird": ird,
                    "hsk": hsk
                }
                insertar_reserva(nueva_data)
                st.query_params.clear()
                st.rerun()

with tabs[2]:
    st.subheader("Importar Reservas")
    archivo = st.file_uploader("Sube un archivo CSV o Excel", type=["csv", "xlsx", "xls"])
    if archivo:
        try:
            if archivo.name.endswith(".csv"):
                df_import = pd.read_csv(archivo)
            else:
                df_import = pd.read_excel(archivo)
            st.write("Vista previa:")
            st.dataframe(df_import.head())

            if st.button("Confirmar Importación"):
                # Mapear columnas comunes
                columnas_requeridas = ["name", "check_in", "check_out"]
                faltantes = [c for c in columnas_requeridas if c not in df_import.columns]
                if faltantes:
                    st.error(f"Faltan columnas requeridas: {faltantes}")
                else:
                    registros = df_import.to_dict(orient="records")
                    # Limpiar NaN
                    for r in registros:
                        for k, v in r.items():
                            if pd.isna(v):
                                r[k] = None
                    insertar_batch_reservas(registros)
                    st.query_params.clear()
                    st.rerun()
        except Exception as e:
            st.error(f"Error al leer archivo: {e}")

with tabs[3]:
    st.subheader("Exportar a Excel")
    if not df_todas.empty:
        buffer = exportar_excel_por_categorias(df_todas)
        if buffer:
            st.download_button(
                label="📥 Descargar Excel categorizado",
                data=buffer,
                file_name=f"Concierge_Arrivals_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("No hay datos para exportar.")

# Footer
st.markdown("---")
st.caption("Concierge Master v5.1 — Sistema de Gestión de Huéspedes")
