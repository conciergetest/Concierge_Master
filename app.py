import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta
from io import BytesIO
from supabase import create_client, Client

# ============================================================
# CONFIGURACIÓN DE PÁGINA Y CONEXIÓN SUPABASE
# ============================================================
st.set_page_config(page_title="Concierge Master v5.1", layout="wide", initial_sidebar_state="collapsed")

@st.cache_resource
def init_supabase():
    """Inicializa la conexión a Supabase usando secrets de Streamlit."""
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)

supabase = init_supabase()
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
# FUNCIONES CRUD CON SUPABASE
# ============================================================
@st.cache_data(ttl=30)
def cargar_reservaciones():
    """Carga todas las reservaciones desde Supabase."""
    response = supabase.table(TABLE_NAME).select("*").execute()
    df = pd.DataFrame(response.data)
    if df.empty:
        return pd.DataFrame(columns=["id", "eta", "name", "qty", "room", "email",
                                     "check_in", "check_out", "res_number", "phone",
                                     "info", "ird", "hsk", "rate", "trans"])
    # Ordenar por check_in y nombre
    df["check_in_dt"] = pd.to_datetime(df["check_in"], format="%b %d", errors="coerce")
    df = df.sort_values(by=["check_in_dt", "name"])
    return df.drop(columns=["check_in_dt"])

def insertar_reserva(data: dict):
    """Inserta una nueva reserva en Supabase."""
    supabase.table(TABLE_NAME).insert(data).execute()
    st.cache_data.clear()

def actualizar_reserva(reserva_id, data: dict):
    """Actualiza una reserva existente en Supabase."""
    supabase.table(TABLE_NAME).update(data).eq("id", reserva_id).execute()
    st.cache_data.clear()

def eliminar_reserva(reserva_id):
    """Elimina una reserva de Supabase."""
    supabase.table(TABLE_NAME).delete().eq("id", reserva_id).execute()
    st.cache_data.clear()

def insertar_batch_reservas(lista_data: list):
    """Inserta múltiples reservas en lote."""
    supabase.table(TABLE_NAME).insert(lista_data).execute()
    st.cache_data.clear()

# ============================================================
# HELPERS DE HORA (sin cambios)
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

def exportar_excel_por_categorias(df):
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
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
button[key="btn_ver_todas"] { background-color: #000000 !important; color: #ffffff !important; font-weight: bold !important; font-size: 0.7rem !important; border: 1px solid #333 !important; border-radius: 6px !important; text-align: center !important; padding: 2px 3px !important; min-height: 22px !important; }
div[data-testid="stHorizontalBlock"] button[key^="checkout_btn_"] { margin-top: 1px !important; margin-bottom: 1px !important; }
div[data-testid="stTextInput"] > div > div > input { background-color: #1a1a2e !important; color: white !important; border: 1px solid #333 !important; border-radius: 8px !important; padding: 6px 10px !important; font-size: 0.85rem !important; }
div[data-testid="stTextInput"] label { margin-bottom: 0 !important; font-size: 0.75rem !important; }
div[data-testid="stDateInput"] > div > div > input { background-color: #1a1a2e !important; color: white !important; border: 1px solid #333 !important; border-radius: 8px !important; padding: 6px 10px !important; font-size: 0.85rem !important; }
div[data-testid="stDateInput"] label { color: #888 !important; font-size: 0.75rem !important; margin-bottom: 2px !important; }
button[key="btn_aplicar_fecha"] { background-color: #00E5FF !important; color: #000000 !important; font-weight: bold !important; font-size: 0.7rem !important; border: none !important; border-radius: 6px !important; padding: 4px 8px !important; min-height: 28px !important; }
button[key="btn_limpiar_fecha"] { background-color: #333333 !important; color: #ffffff !important; font-weight: bold !important; font-size: 0.7rem !important; border: 1px solid #555 !important; border-radius: 6px !important; padding: 4px 8px !important; min-height: 28px !important; }
button[key="btn_nueva_v2"] { background-color: #00E5FF !important; color: #000000 !important; font-weight: bold !important; font-size: 0.65rem !important; border: none !important; border-radius: 6px !important; padding: 4px 2px !important; min-height: 28px !important; }
button[key="btn_editar_v2"] { background-color: #FF9800 !important; color: #000000 !important; font-weight: bold !important; font-size: 0.65rem !important; border: none !important; border-radius: 6px !important; padding: 4px 2px !important; min-height: 28px !important; }
button[key="btn_importar_v2"] { background-color: #4CAF50 !important; color: #ffffff !important; font-weight: bold !important; font-size: 0.65rem !important; border: none !important; border-radius: 6px !important; padding: 4px 2px !important; min-height: 28px !important; }
button[key="btn_exportar_v2"] { background-color: #2196F3 !important; color: #ffffff !important; font-weight: bold !important; font-size: 0.65rem !important; border: none !important; border-radius: 6px !important; padding: 4px 2px !important; min-height: 28px !important; }
button[key="btn_carta_v2"] { background-color: #9C27B0 !important; color: #ffffff !important; font-weight: bold !important; font-size: 0.65rem !important; border: none !important; border-radius: 6px !important; padding: 4px 2px !important; min-height: 28px !important; }
button[key="btn_cancelar_v2"] { background-color: #f44336 !important; color: #ffffff !important; font-weight: bold !important; font-size: 0.65rem !important; border: none !important; border-radius: 6px !important; padding: 4px 2px !important; min-height: 28px !important; }
button[key="btn_reporte_v2"] { background-color: #FFC107 !important; color: #000000 !important; font-weight: bold !important; font-size: 0.65rem !important; border: none !important; border-radius: 6px !important; padding: 4px 2px !important; min-height: 28px !important; }
button[key="btn_agenda_v2"] { background-color: #00BCD4 !important; color: #000000 !important; font-weight: bold !important; font-size: 0.65rem !important; border: none !important; border-radius: 6px !important; padding: 4px 2px !important; min-height: 28px !important; }
button[key="btn_procesar_excel"] { background-color: #00E5FF !important; color: #000000 !important; font-weight: bold !important; font-size: 0.85rem !important; border: none !important; border-radius: 8px !important; padding: 10px 20px !important; min-height: 40px !important; }
.ag-root-wrapper { background-color: #101010 !important; }
.ag-cell { color: white !important; background-color: #101010 !important; }
.ag-row-selected .ag-cell { background-color: #00FFFF !important; color: #000000 !important; font-weight: bold !important; }
</style>
""", unsafe_allow_html=True)

header_col1, header_col2 = st.columns([1.3, 8.7])
with header_col1:
    st.markdown('''<h2 style="color:#00E5FF; margin:0; padding:0; font-size:1.1rem; line-height:1.0;">Concierge<br>Master v5.1</h2>''', unsafe_allow_html=True)
with header_col2:
    import streamlit.components.v1 as components

    clock_html = """
    <!DOCTYPE html>
    <html>
    <head>
        <style>
            body { margin: 0; padding: 0; background: transparent; }
            #hora-local {
                text-align: right;
                color: #00E5FF;
                font-size: 1.1rem;
                font-weight: bold;
                margin-top: 2px;
                text-shadow: 0 0 10px #00E5FF;
                font-family: 'Segoe UI', sans-serif;
                white-space: nowrap;
            }
        </style>
    </head>
    <body>
        <div id="hora-local">Cargando hora...</div>
        <script>
            function actualizarHora() {
                const ahora = new Date();
                const mes = ahora.toLocaleDateString('en-US', { month: 'long' });
                const dia = ahora.getDate();
                const anio = ahora.getFullYear();
                let hora = ahora.getHours();
                const minutos = String(ahora.getMinutes()).padStart(2, '0');
                const segundos = String(ahora.getSeconds()).padStart(2, '0');
                const ampm = hora >= 12 ? 'PM' : 'AM';
                hora = hora % 12;
                hora = hora ? hora : 12;
                const textoFinal = mes + ' ' + dia + ', ' + anio + ' — ' + hora + ':' + minutos + ':' + segundos + ' ' + ampm;
                document.getElementById('hora-local').textContent = textoFinal;
            }
            actualizarHora();
            setInterval(actualizarHora, 1000);
        </script>
    </body>
    </html>
    """
    components.html(clock_html, height=35)

df_todas = cargar_reservaciones()
total_reservas = len(df_todas)

search_col1, search_col2 = st.columns([2.5, 8.5])
with search_col1:
    st.markdown(f"""<div style="background-color: #1a1a2e; border-radius: 8px; padding: 6px 10px; margin-bottom: 4px; text-align: center; border: 1px solid #00E5FF;">
        <span style="color: #00E5FF; font-size: 0.75rem; font-weight: bold;">📊 TOTAL RESERVAS:</span>
        <span style="color: #ffffff; font-size: 1.0rem; font-weight: bold;"> {total_reservas}</span></div>""", unsafe_allow_html=True)
    st.markdown("<p style='color:#888; font-size:0.75rem; margin:0; padding:0;'>🔍 Búsqueda rápida...</p>", unsafe_allow_html=True)
    busqueda = st.text_input("", placeholder="Buscar por nombre, teléfono, reserva, VIP, Relaxury...", label_visibility="collapsed", key="buscador_global")
with search_col2: pass

left_col, right_col = st.columns([5.0, 5.0])
with left_col:
    st.markdown("""<div style="background-color: #1a1a2e; border-radius: 8px; padding: 5px 8px 2px 8px; margin-bottom: 1px;">
        <div style="color: #ffffff; font-size: 0.8rem; font-weight: bold; text-align: center; margin-bottom: 3px;">🏨 Checking Out Rooms</div></div>""", unsafe_allow_html=True)
    hoy = datetime.now()
    fechas_checkout = [hoy + timedelta(days=i) for i in range(8)]
    for i in range(0, len(fechas_checkout), 2):
        cols = st.columns(2)
        for j in range(2):
            if i + j < len(fechas_checkout):
                fecha = fechas_checkout[i + j]
                mes_abr, dia_num, fecha_db = fecha.strftime("%b").upper(), fecha.strftime("%d"), fecha.strftime("%b %d")
                count = len(df_todas[df_todas["check_out"] == fecha_db])
                with cols[j]:
                    if st.button(f"{dia_num}-{mes_abr}: [{count}]", key=f"checkout_btn_{fecha_db}", use_container_width=True):
                        st.query_params["checkout_filtro"] = fecha_db
                        st.query_params.pop("fecha_date", None)
                        st.query_params.pop("fecha_activa", None)
                        st.rerun()
    st.markdown("<div style='height: 1px;'></div>", unsafe_allow_html=True)
    if st.button("🔄 VER TODAS", key="btn_ver_todas", use_container_width=True):
        st.query_params.pop("checkout_filtro", None)
        st.query_params.pop("fecha_date", None)
        st.query_params.pop("fecha_activa", None)
        st.rerun()
    st.markdown("<div style='height: 4px;'></div>", unsafe_allow_html=True)
    fecha_default = datetime.now()
    if filtro_fecha_date:
        try: fecha_default = datetime.strptime(filtro_fecha_date, "%Y-%m-%d")
        except: fecha_default = datetime.now()
    fecha_seleccionada = st.date_input("📅 DATE", value=fecha_default, key="date_filter_picker", label_visibility="visible")
    col_aplicar, col_limpiar = st.columns(2)
    with col_aplicar:
        if st.button("🔍 APLICAR", key="btn_aplicar_fecha", use_container_width=True):
            st.query_params["fecha_date"] = fecha_seleccionada.strftime("%Y-%m-%d")
            st.query_params["fecha_activa"] = "true"
            st.query_params.pop("checkout_filtro", None)
            st.rerun()
    with col_limpiar:
        if st.button("🧹 LIMPIAR", key="btn_limpiar_fecha", use_container_width=True):
            st.query_params.pop("fecha_date", None)
            st.query_params.pop("fecha_activa", None)
            st.query_params.pop("checkout_filtro", None)
            st.rerun()

    # BOTONES DE ACCION
    st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)
    btn_col1, btn_col2, btn_col3, btn_col4, btn_col5, btn_col6, btn_col7, btn_col8 = st.columns(8)
    if btn_col1.button("NUEVA", key="btn_nueva_v2", use_container_width=True): st.query_params["action"] = "nueva"; st.rerun()
    if btn_col2.button("EDITAR", key="btn_editar_v2", use_container_width=True): st.query_params["action"] = "editar"; st.rerun()
    if btn_col3.button("IMPORTAR", key="btn_importar_v2", use_container_width=True): st.query_params["action"] = "importar"; st.rerun()
    if btn_col4.button("EXPORTAR", key="btn_exportar_v2", use_container_width=True): st.query_params["action"] = "exportar"; st.rerun()
    if btn_col5.button("CARTA", key="btn_carta_v2", use_container_width=True): st.query_params["action"] = "carta"; st.rerun()
    if btn_col6.button("BORRAR", key="btn_cancelar_v2", use_container_width=True): st.query_params["action"] = "cancelar"; st.rerun()
    if btn_col7.button("REPORTE", key="btn_reporte_v2", use_container_width=True): st.query_params["action"] = "reporte"; st.rerun()
    if btn_col8.button("AGENDA", key="btn_agenda_v2", use_container_width=True): st.switch_page("pages/agenda.py")

with right_col:
    categorias = {"VIP": "#00E5FF", "ANNIVERSARY": "#4CAF50", "BIRTHDAY": "#FF5252",
                  "HONEYMOON": "#FF9800", "BABYMOON": "#9C27B0", "TEAM MEMBER": "#FFC107", "LEISURE": "#2196F3"}
    conteo_categorias = {}
    for cat in categorias:
        if cat == "LEISURE": continue
        conteo_categorias[cat] = df_todas["info"].astype(str).str.upper().str.contains(cat, na=False).sum()
    total_categorizadas = sum(conteo_categorias.values())
    conteo_categorias["LEISURE"] = max(0, total_reservas - total_categorizadas)
    conteo_ordenado = dict(sorted(conteo_categorias.items(), key=lambda x: x[1], reverse=True))
    max_valor = max(conteo_ordenado.values()) if conteo_ordenado else 1
    html_chart = """<!DOCTYPE html><html><head><style>
    *{margin:0;padding:0;box-sizing:border-box}body{background-color:#0d0d0d;font-family:'Segoe UI',sans-serif}
    .chart-container{background-color:#0d0d0d;border-radius:12px;padding:12px 15px;min-height:220px;color:white}
    .chart-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:10px}
    .chart-title{color:#888;font-size:10px;font-weight:bold;letter-spacing:1px;text-transform:uppercase}
    .total-circle{width:42px;height:42px;border:2px solid #00E5FF;border-radius:50%;display:flex;align-items:center;justify-content:center}
    .total-number{color:#00E5FF;font-size:14px;font-weight:bold}
    .bar-row{display:flex;align-items:center;margin-bottom:5px}
    .bar-label{width:90px;color:#ccc;font-size:9px;text-align:right;padding-right:8px;white-space:nowrap}
    .bar-track{flex:1;background-color:#1a1a1a;border-radius:3px;height:16px;position:relative;overflow:hidden}
    .bar-fill{height:100%;border-radius:3px;transition:width 0.5s ease}
    .bar-value{width:28px;color:#fff;font-size:11px;font-weight:bold;text-align:right;padding-left:8px}
    </style></head><body><div class="chart-container"><div class="chart-header">
    <div class="chart-title">Guest Categories</div><div class="total-circle"><div class="total-number">""" + str(total_reservas) + """</div></div></div>"""
    for cat, valor in conteo_ordenado.items():
        color = categorias.get(cat, "#888")
        porcentaje = (valor / max_valor * 100) if max_valor > 0 else 0
        html_chart += f"""<div class="bar-row"><div class="bar-label">{cat}</div><div class="bar-track">
        <div class="bar-fill" style="width:{porcentaje}%;background-color:{color};"></div></div><div class="bar-value">{valor}</div></div>"""
    html_chart += """</div></body></html>"""
    st.html(html_chart)
    mask_relaxury = df_todas.astype(str).apply(lambda row: row.str.upper().str.contains("RELAXURY", na=False).any(), axis=1)
    total_relaxury = mask_relaxury.sum()
    st.markdown(f"""<div style="background-color: #1a1a2e; border-radius: 8px; padding: 6px 10px; margin-top: 2px; text-align: center; border: 1px solid #E91E63;">
        <span style="color: #E91E63; font-size: 0.75rem; font-weight: bold;">🏖️ RELAXURY:</span>
        <span style="color: #ffffff; font-size: 1.0rem; font-weight: bold;"> {total_relaxury}</span></div>""", unsafe_allow_html=True)

# ============================================================
# FORMULARIO IMPORTAR DESDE EXCEL
# ============================================================
if mostrar_importar:
    with st.container():
        st.subheader("📥 Importar Reservaciones desde Excel")
        col1, col2 = st.columns([1, 1])
        with col1:
            if st.button("↩️ REGRESAR", key="regresar_importar"):
                st.query_params.clear(); st.rerun()
        st.markdown("""<div style="background-color: #1a1a2e; border-radius: 8px; padding: 15px; margin: 10px 0; border: 1px solid #333;">
            <p style="color: #ccc; font-size: 0.85rem; margin: 0;">📋 <b>Instrucciones:</b> Selecciona tu archivo <code>Plantilla_Importar.xlsx</code>. 
            El archivo debe contener las columnas: <b>eta, name, qty, room, email, check_in, check_out, res_number, phone, info, ird, hsk, rate, trans</b></p></div>""", unsafe_allow_html=True)
        archivo_subido = st.file_uploader("📁 Seleccionar archivo Excel", type=["xlsx", "xls"], key="uploader_excel")
        if archivo_subido is not None:
            try:
                df_excel = pd.read_excel(archivo_subido)
                st.markdown(f"""<div style="background-color: #0d1f0d; border-radius: 8px; padding: 10px 15px; margin: 10px 0; border: 1px solid #2e7d32;">
                    <p style="color: #4CAF50; font-size: 0.85rem; margin: 0;">✅ Archivo cargado: <b>{archivo_subido.name}</b> | Filas detectadas: <b>{len(df_excel)}</b></p></div>""", unsafe_allow_html=True)
                columnas_esperadas = ["eta", "name", "qty", "room", "email", "check_in", "check_out",
                                      "res_number", "phone", "info", "ird", "hsk", "rate", "trans"]
                columnas_faltantes = [c for c in columnas_esperadas if c not in df_excel.columns]
                columnas_extra = [c for c in df_excel.columns if c not in columnas_esperadas and c != "id"]
                if columnas_faltantes:
                    st.error(f"❌ Columnas faltantes: {', '.join(columnas_faltantes)}")
                else:
                    if columnas_extra: st.warning(f"⚠️ Columnas extra (serán ignoradas): {', '.join(columnas_extra)}")
                    df_preview = df_excel[columnas_esperadas].copy()
                    for col in ["check_in", "check_out"]:
                        if col in df_preview.columns:
                            try: df_preview[col] = pd.to_datetime(df_preview[col], errors="coerce").dt.strftime("%b %d")
                            except: pass
                    st.dataframe(df_preview, use_container_width=True, height=250)
                    col_proc, _ = st.columns([1, 3])
                    with col_proc:
                        if st.button("📥 IMPORTAR A BASE DE DATOS", key="btn_procesar_excel", use_container_width=True):
                            try:
                                registros_insertados, registros_error, errores_detalle = 0, 0, []
                                registros_batch = []
                                for idx, row in df_excel.iterrows():
                                    try:
                                        eta = str(row.get("eta", "")).strip()
                                        name = str(row.get("name", "")).strip()
                                        qty = int(row.get("qty", 0)) if pd.notna(row.get("qty")) else 0
                                        room = str(row.get("room", "")).strip()
                                        email = str(row.get("email", "")).strip()
                                        check_in_raw, check_out_raw = row.get("check_in", ""), row.get("check_out", "")
                                        if pd.notna(check_in_raw):
                                            try: check_in = pd.to_datetime(check_in_raw).strftime("%b %d")
                                            except: check_in = str(check_in_raw).strip()
                                        else: check_in = ""
                                        if pd.notna(check_out_raw):
                                            try: check_out = pd.to_datetime(check_out_raw).strftime("%b %d")
                                            except: check_out = str(check_out_raw).strip()
                                        else: check_out = ""
                                        res_number = str(row.get("res_number", "")).strip()
                                        phone = str(row.get("phone", "")).strip()
                                        info = str(row.get("info", "")).strip()
                                        ird = str(row.get("ird", "")).strip()
                                        hsk = str(row.get("hsk", "")).strip()
                                        rate = str(row.get("rate", "")).strip()
                                        trans = str(row.get("trans", "")).strip()

                                        registro = {
                                            "eta": eta, "name": name, "qty": qty, "room": room,
                                            "email": email, "check_in": check_in, "check_out": check_out,
                                            "res_number": res_number, "phone": phone, "info": info,
                                            "ird": ird, "hsk": hsk, "rate": rate, "trans": trans
                                        }
                                        registros_batch.append(registro)
                                        registros_insertados += 1
                                    except Exception as e:
                                        registros_error += 1
                                        errores_detalle.append(f"Fila {idx + 2}: {str(e)}")

                                # Insertar en lote a Supabase
                                if registros_batch:
                                    insertar_batch_reservas(registros_batch)

                                st.markdown(f"""<div style="background-color: #0d1f0d; border-radius: 8px; padding: 15px; margin: 15px 0; border: 1px solid #2e7d32;">
                                    <h4 style="color: #4CAF50; margin: 0 0 10px 0;">✅ Importación Completada</h4>
                                    <p style="color: #ccc; font-size: 0.9rem; margin: 0;">📥 Registros insertados: <b style="color: #4CAF50;">{registros_insertados}</b><br>
                                    ❌ Errores: <b>{registros_error}</b></p></div>""", unsafe_allow_html=True)
                                if registros_error > 0 and errores_detalle:
                                    with st.expander("🔍 Ver detalle de errores"):
                                        for error in errores_detalle[:10]: st.error(error)
                                        if len(errores_detalle) > 10: st.warning(f"... y {len(errores_detalle) - 10} errores más.")
                                if registros_insertados > 0:
                                    st.success("🎉 ¡Las reservas se han importado correctamente!")
                                    st.info("🔄 La página se recargará en 3 segundos...")
                                    import time; time.sleep(3)
                                    st.query_params.clear(); st.rerun()
                            except Exception as e: st.error(f"❌ Error al importar: {str(e)}")
            except Exception as e: st.error(f"❌ Error al leer el archivo: {str(e)}")

# ============================================================
# FORMULARIO EXPORTAR A EXCEL
# ============================================================
if mostrar_exportar:
    with st.container():
        st.subheader("📤 Exportar Reservaciones a Excel")
        col1, col2 = st.columns([1, 1])
        with col1:
            if st.button("↩️ REGRESAR", key="regresar_exportar"):
                st.query_params.clear(); st.rerun()
        df_export = cargar_reservaciones()
        filtro_activo = False
        if filtro_checkout:
            df_export = df_export[df_export["check_out"] == filtro_checkout]; filtro_activo = True
        if fecha_filtro_activo and filtro_fecha_date:
            fecha_filtro = datetime.strptime(filtro_fecha_date, "%Y-%m-%d").strftime("%b %d")
            df_export = df_export[df_export["check_in"] == fecha_filtro]; filtro_activo = True
        if busqueda and busqueda.strip():
            busqueda_lower = busqueda.strip().lower()
            mask = df_export.astype(str).apply(lambda row: row.str.lower().str.contains(busqueda_lower, na=False).any(), axis=1)
            df_export = df_export[mask]; filtro_activo = True
        total_a_exportar = len(df_export)
        st.markdown(f"""<div style="background-color: #1a1a2e; border-radius: 8px; padding: 15px; margin: 10px 0; border: 1px solid #333;">
            <p style="color: #ccc; font-size: 0.85rem; margin: 0;">📋 <b>Resumen de exportación:</b><br>
            • Total de reservas a exportar: <b style="color: #00E5FF;">{total_a_exportar}</b><br>
            • Filtro activo: <b style="color: #00E5FF;">{"Sí" if filtro_activo else "No"}</b><br>
            • El archivo se organizará por categorías: <b>CUMPLEAÑOS, VIP, HONEYMOON, ANNIVERSARY, BABYMOON, TEAM MEMBER, GENERAL</b></p></div>""", unsafe_allow_html=True)
        st.markdown("<p style='color:#888; font-size:0.8rem; margin-top:15px;'>👁️ Vista previa del contenido a exportar:</p>", unsafe_allow_html=True)
        st.dataframe(df_export, use_container_width=True, height=300)
        if total_a_exportar > 0:
            try:
                excel_buffer = exportar_excel_por_categorias(df_export)
                fecha_hoy = datetime.now().strftime("%Y%m%d_%H%M")
                st.download_button(label="📥 DESCARGAR EXCEL", data=excel_buffer,
                    file_name=f"Arrivals_{fecha_hoy}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_descargar_excel")
            except Exception as e:
                st.error(f"❌ Error al generar el Excel: {str(e)}")
                st.exception(e)
        else: st.warning("⚠️ No hay datos para exportar con los filtros actuales.")

# ============================================================
# FORMULARIO REPORTE DE OCUPACIÓN DIARIO
# ============================================================
if st.query_params.get("action") == "reporte":
    with st.container():
        st.subheader("📄 Reporte de Ocupación Diario")
        col1, col2 = st.columns([1, 1])
        with col1:
            if st.button("↩️ REGRESAR", key="regresar_reporte"):
                st.query_params.clear(); st.rerun()
        hoy = datetime.now()
        hoy_str = hoy.strftime("%b %d")
        manana = hoy + timedelta(days=1)
        manana_str = manana.strftime("%b %d")
        df_todas = cargar_reservaciones()

        def fecha_es_menor_igual(fecha_str, referencia_str):
            try:
                fecha = datetime.strptime(fecha_str, "%b %d").replace(year=hoy.year)
                referencia = datetime.strptime(referencia_str, "%b %d").replace(year=hoy.year)
                return fecha <= referencia
            except: return False

        def fecha_es_mayor(fecha_str, referencia_str):
            try:
                fecha = datetime.strptime(fecha_str, "%b %d").replace(year=hoy.year)
                referencia = datetime.strptime(referencia_str, "%b %d").replace(year=hoy.year)
                return fecha > referencia
            except: return False

        def obtener_rooms(df):
            rooms = df["room"].dropna().astype(str)
            return rooms[rooms.str.strip() != ""].str.strip().tolist()

        def formatear_rooms(rooms_list):
            if not rooms_list: return "Ninguna"
            if len(rooms_list) <= 5: return ", ".join(rooms_list)
            return ", ".join(rooms_list[:5]) + f" (+{len(rooms_list) - 5} más)"

        mask_en_casa = df_todas.apply(lambda row: fecha_es_menor_igual(row["check_in"], hoy_str) and fecha_es_mayor(row["check_out"], hoy_str), axis=1)
        df_en_casa = df_todas[mask_en_casa]
        total_en_casa = len(df_en_casa)
        rooms_en_casa = obtener_rooms(df_en_casa)
        mask_vip_en_casa = df_en_casa["info"].astype(str).str.upper().str.contains("VIP", na=False)
        vip_en_casa = df_en_casa[mask_vip_en_casa]
        total_vip_en_casa = len(vip_en_casa)
        rooms_vip_en_casa = obtener_rooms(vip_en_casa)

        df_salen_hoy = df_todas[df_todas["check_out"] == hoy_str]
        total_salen_hoy = len(df_salen_hoy)
        rooms_salen_hoy = obtener_rooms(df_salen_hoy)
        mask_vip_salen_hoy = df_salen_hoy["info"].astype(str).str.upper().str.contains("VIP", na=False)
        vip_salen_hoy = df_salen_hoy[mask_vip_salen_hoy]
        total_vip_salen_hoy = len(vip_salen_hoy)
        rooms_vip_salen_hoy = obtener_rooms(vip_salen_hoy)

        df_salen_manana = df_todas[df_todas["check_out"] == manana_str]
        total_salen_manana = len(df_salen_manana)
        rooms_salen_manana = obtener_rooms(df_salen_manana)
        mask_vip_salen_manana = df_salen_manana["info"].astype(str).str.upper().str.contains("VIP", na=False)
        vip_salen_manana = df_salen_manana[mask_vip_salen_manana]
        total_vip_salen_manana = len(vip_salen_manana)
        rooms_vip_salen_manana = obtener_rooms(vip_salen_manana)

        df_llegan_hoy = df_todas[df_todas["check_in"] == hoy_str]
        total_llegan_hoy = len(df_llegan_hoy)
        rooms_llegan_hoy = obtener_rooms(df_llegan_hoy)
        mask_vip_llegan_hoy = df_llegan_hoy["info"].astype(str).str.upper().str.contains("VIP", na=False)
        vip_llegan_hoy = df_llegan_hoy[mask_vip_llegan_hoy]
        total_vip_llegan_hoy = len(vip_llegan_hoy)
        rooms_vip_llegan_hoy = obtener_rooms(vip_llegan_hoy)

        df_llegan_manana = df_todas[df_todas["check_in"] == manana_str]
        total_llegan_manana = len(df_llegan_manana)
        rooms_llegan_manana = obtener_rooms(df_llegan_manana)
        mask_vip_llegan_manana = df_llegan_manana["info"].astype(str).str.upper().str.contains("VIP", na=False)
        vip_llegan_manana = df_llegan_manana[mask_vip_llegan_manana]
        total_vip_llegan_manana = len(vip_llegan_manana)
        rooms_vip_llegan_manana = obtener_rooms(vip_llegan_manana)

        total_reservas = len(df_todas)

        st.markdown(f"""<div style="background-color: #1a1a2e; border-radius: 8px; padding: 15px; margin: 10px 0; border: 1px solid #333;">
            <p style="color: #ccc; font-size: 0.85rem; margin: 0;">📅 <b>Fecha del Reporte:</b> <span style="color: #00E5FF;">{hoy.strftime("%A, %B %d, %Y")}</span><br>
            🕐 <b>Hora de generación:</b> <span style="color: #00E5FF;">{hoy.strftime("%H:%M")}</span></p></div>""", unsafe_allow_html=True)

        col_m1, col_m2, col_m3 = st.columns(3)
        with col_m1: st.markdown(f"""<div style="background-color: #0d1f0d; border-radius: 8px; padding: 12px; text-align: center; border: 1px solid #2e7d32;">
            <div style="color: #4CAF50; font-size: 0.75rem; font-weight: bold;">🏨 EN CASA</div>
            <div style="color: #fff; font-size: 1.8rem; font-weight: bold;">{total_en_casa}</div>
            <div style="color: #888; font-size: 0.65rem; margin-top: 4px;">🚪 {formatear_rooms(rooms_en_casa)}</div></div>""", unsafe_allow_html=True)
        with col_m2: st.markdown(f"""<div style="background-color: #0d1f0d; border-radius: 8px; padding: 12px; text-align: center; border: 1px solid #2e7d32;">
            <div style="color: #4CAF50; font-size: 0.75rem; font-weight: bold;">👑 VIPs EN CASA</div>
            <div style="color: #fff; font-size: 1.8rem; font-weight: bold;">{total_vip_en_casa}</div>
            <div style="color: #888; font-size: 0.65rem; margin-top: 4px;">🚪 {formatear_rooms(rooms_vip_en_casa)}</div></div>""", unsafe_allow_html=True)
        with col_m3: st.markdown(f"""<div style="background-color: #0d1f0d; border-radius: 8px; padding: 12px; text-align: center; border: 1px solid #2e7d32;">
            <div style="color: #4CAF50; font-size: 0.75rem; font-weight: bold;">📊 TOTAL DB</div>
            <div style="color: #fff; font-size: 1.8rem; font-weight: bold;">{total_reservas}</div>
            <div style="color: #888; font-size: 0.65rem; margin-top: 4px;">Reservas registradas</div></div>""", unsafe_allow_html=True)

        st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)
        col_m4, col_m5, col_m6 = st.columns(3)
        with col_m4: st.markdown(f"""<div style="background-color: #1a0d0d; border-radius: 8px; padding: 12px; text-align: center; border: 1px solid #7d2e2e;">
            <div style="color: #f44336; font-size: 0.75rem; font-weight: bold;">🚪 SALEN HOY ({hoy_str})</div>
            <div style="color: #fff; font-size: 1.8rem; font-weight: bold;">{total_salen_hoy}</div>
            <div style="color: #888; font-size: 0.65rem; margin-top: 4px;">👑 VIPs: {total_vip_salen_hoy} | 🚪 {formatear_rooms(rooms_salen_hoy)}</div></div>""", unsafe_allow_html=True)
        with col_m5: st.markdown(f"""<div style="background-color: #1a0d0d; border-radius: 8px; padding: 12px; text-align: center; border: 1px solid #7d2e2e;">
            <div style="color: #f44336; font-size: 0.75rem; font-weight: bold;">🚪 SALEN MAÑANA ({manana_str})</div>
            <div style="color: #fff; font-size: 1.8rem; font-weight: bold;">{total_salen_manana}</div>
            <div style="color: #888; font-size: 0.65rem; margin-top: 4px;">👑 VIPs: {total_vip_salen_manana} | 🚪 {formatear_rooms(rooms_salen_manana)}</div></div>""", unsafe_allow_html=True)
        with col_m6: st.markdown(f"""<div style="background-color: #0d0d1a; border-radius: 8px; padding: 12px; text-align: center; border: 1px solid #2e2e7d;">
            <div style="color: #2196F3; font-size: 0.75rem; font-weight: bold;">📥 LLEGAN HOY ({hoy_str})</div>
            <div style="color: #fff; font-size: 1.8rem; font-weight: bold;">{total_llegan_hoy}</div>
            <div style="color: #888; font-size: 0.65rem; margin-top: 4px;">👑 VIPs: {total_vip_llegan_hoy} | 🚪 {formatear_rooms(rooms_llegan_hoy)}</div></div>""", unsafe_allow_html=True)

        st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)
        col_m7, col_m8, col_m9 = st.columns(3)
        with col_m7: st.markdown(f"""<div style="background-color: #0d0d1a; border-radius: 8px; padding: 12px; text-align: center; border: 1px solid #2e2e7d;">
            <div style="color: #2196F3; font-size: 0.75rem; font-weight: bold;">📥 LLEGAN MAÑANA ({manana_str})</div>
            <div style="color: #fff; font-size: 1.8rem; font-weight: bold;">{total_llegan_manana}</div>
            <div style="color: #888; font-size: 0.65rem; margin-top: 4px;">👑 VIPs: {total_vip_llegan_manana} | 🚪 {formatear_rooms(rooms_llegan_manana)}</div></div>""", unsafe_allow_html=True)
        with col_m8: st.markdown(f"""<div style="background-color: #1a1a0d; border-radius: 8px; padding: 12px; text-align: center; border: 1px solid #7d7d2e;">
            <div style="color: #FFC107; font-size: 0.75rem; font-weight: bold;">🌙 ESTANCIA PROMEDIO</div>
            <div style="color: #fff; font-size: 1.8rem; font-weight: bold;">{round(total_en_casa / max(total_salen_hoy, 1), 1)}</div>
            <div style="color: #888; font-size: 0.65rem; margin-top: 4px;">Ratio en casa / salen hoy</div></div>""", unsafe_allow_html=True)
        with col_m9:
            ocupacion_pct = round((total_en_casa / max(total_reservas, 1)) * 100, 1)
            st.markdown(f"""<div style="background-color: #1a0d1a; border-radius: 8px; padding: 12px; text-align: center; border: 1px solid #7d2e7d;">
                <div style="color: #9C27B0; font-size: 0.75rem; font-weight: bold;">📈 OCUPACIÓN</div>
                <div style="color: #fff; font-size: 1.8rem; font-weight: bold;">{ocupacion_pct}%</div>
                <div style="color: #888; font-size: 0.65rem; margin-top: 4px;">En casa / Total DB</div></div>""", unsafe_allow_html=True)

        st.markdown("<div style='height: 15px;'></div>", unsafe_allow_html=True)
        st.markdown("""<div style="background-color: #1a1a2e; border-radius: 8px; padding: 10px; margin-bottom: 10px; border: 1px solid #00E5FF;">
            <div style="color: #00E5FF; font-size: 0.85rem; font-weight: bold; text-align: center;">📋 DETALLE DE HABITACIONES POR CATEGORÍA</div></div>""", unsafe_allow_html=True)
        resumen_rooms_data = []
        if rooms_en_casa: resumen_rooms_data.append({"Categoría": "🏨 En Casa", "Total": total_en_casa, "Habitaciones": ", ".join(rooms_en_casa)})
        if rooms_vip_en_casa: resumen_rooms_data.append({"Categoría": "👑 VIPs En Casa", "Total": total_vip_en_casa, "Habitaciones": ", ".join(rooms_vip_en_casa)})
        if rooms_salen_hoy: resumen_rooms_data.append({"Categoría": f"🚪 Salen Hoy ({hoy_str})", "Total": total_salen_hoy, "Habitaciones": ", ".join(rooms_salen_hoy)})
        if rooms_vip_salen_hoy: resumen_rooms_data.append({"Categoría": f"👑 VIPs Salen Hoy ({hoy_str})", "Total": total_vip_salen_hoy, "Habitaciones": ", ".join(rooms_vip_salen_hoy)})
        if rooms_salen_manana: resumen_rooms_data.append({"Categoría": f"🚪 Salen Mañana ({manana_str})", "Total": total_salen_manana, "Habitaciones": ", ".join(rooms_salen_manana)})
        if rooms_vip_salen_manana: resumen_rooms_data.append({"Categoría": f"👑 VIPs Salen Mañana ({manana_str})", "Total": total_vip_salen_manana, "Habitaciones": ", ".join(rooms_vip_salen_manana)})
        if rooms_llegan_hoy: resumen_rooms_data.append({"Categoría": f"📥 Llegan Hoy ({hoy_str})", "Total": total_llegan_hoy, "Habitaciones": ", ".join(rooms_llegan_hoy)})
        if rooms_vip_llegan_hoy: resumen_rooms_data.append({"Categoría": f"👑 VIPs Llegan Hoy ({hoy_str})", "Total": total_vip_llegan_hoy, "Habitaciones": ", ".join(rooms_vip_llegan_hoy)})
        if rooms_llegan_manana: resumen_rooms_data.append({"Categoría": f"📥 Llegan Mañana ({manana_str})", "Total": total_llegan_manana, "Habitaciones": ", ".join(rooms_llegan_manana)})
        if rooms_vip_llegan_manana: resumen_rooms_data.append({"Categoría": f"👑 VIPs Llegan Mañana ({manana_str})", "Total": total_vip_llegan_manana, "Habitaciones": ", ".join(rooms_vip_llegan_manana)})
        if resumen_rooms_data:
            df_resumen_rooms = pd.DataFrame(resumen_rooms_data)
            st.dataframe(df_resumen_rooms, use_container_width=True, hide_index=True)
        else: st.info("📭 No hay habitaciones para mostrar en el resumen.")

        def generar_reporte_excel():
            from openpyxl import Workbook
            from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
            wb = Workbook()
            fill_titulo = PatternFill(start_color="00B0F0", end_color="00B0F0", fill_type="solid")
            fill_header = PatternFill(start_color="404040", end_color="404040", fill_type="solid")
            fill_verde = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
            fill_rojo = PatternFill(start_color="C62828", end_color="C62828", fill_type="solid")
            fill_azul = PatternFill(start_color="1565C0", end_color="1565C0", fill_type="solid")
            fill_morado = PatternFill(start_color="6A1B9A", end_color="6A1B9A", fill_type="solid")
            fill_gris = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
            font_titulo = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
            font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            font_metrica = Font(name="Calibri", size=14, bold=True, color="000000")
            font_label = Font(name="Calibri", size=10, color="000000")
            font_fecha = Font(name="Calibri", size=12, bold=True, color="00B0F0")
            font_rooms = Font(name="Calibri", size=9, color="666666")
            align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
            align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
            thin_border = Border(left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"),
                                 top=Side(style="thin", color="D9D9D9"), bottom=Side(style="thin", color="D9D9D9"))
            # HOJA 1: RESUMEN
            ws_resumen = wb.active
            ws_resumen.title = "Resumen Ejecutivo"
            ws_resumen.merge_cells("A1:E1")
            cell = ws_resumen["A1"]
            cell.value = f"REPORTE DE OCUPACIÓN - {hoy.strftime('%A, %B %d, %Y').upper()}"
            cell.fill, cell.font, cell.alignment = fill_titulo, font_titulo, align_center
            ws_resumen.merge_cells("A2:E2")
            cell = ws_resumen["A2"]
            cell.value = f"Generado el {hoy.strftime('%d/%m/%Y')} a las {hoy.strftime('%H:%M')}"
            cell.font, cell.alignment = font_fecha, align_center
            ws_resumen.row_dimensions[1].height = 30
            ws_resumen.row_dimensions[2].height = 20
            headers_metricas = ["MÉTRICA", "TOTAL", "VIPs", "HABITACIONES", "NOTAS"]
            for col_idx, header in enumerate(headers_metricas, 1):
                cell = ws_resumen.cell(row=4, column=col_idx, value=header)
                cell.fill, cell.font, cell.alignment, cell.border = fill_header, font_header, align_center, thin_border
            metricas_data = [
                ("🏨 EN CASA", total_en_casa, total_vip_en_casa, ", ".join(rooms_en_casa) if rooms_en_casa else "N/A", ""),
                ("🚪 SALEN HOY", total_salen_hoy, total_vip_salen_hoy, ", ".join(rooms_salen_hoy) if rooms_salen_hoy else "N/A", hoy_str),
                ("🚪 SALEN MAÑANA", total_salen_manana, total_vip_salen_manana, ", ".join(rooms_salen_manana) if rooms_salen_manana else "N/A", manana_str),
                ("📥 LLEGAN HOY", total_llegan_hoy, total_vip_llegan_hoy, ", ".join(rooms_llegan_hoy) if rooms_llegan_hoy else "N/A", hoy_str),
                ("📥 LLEGAN MAÑANA", total_llegan_manana, total_vip_llegan_manana, ", ".join(rooms_llegan_manana) if rooms_llegan_manana else "N/A", manana_str),
                ("📊 TOTAL DB", total_reservas, "-", "-", "Todas las reservas"),
                ("📈 % OCUPACIÓN", f"{ocupacion_pct}%", "-", "-", f"{total_en_casa} de {total_reservas}"),
            ]
            for i, (label, total, vips, rooms, notas) in enumerate(metricas_data):
                row = 5 + i
                for col_idx, val in enumerate([label, total, vips, rooms, notas], 1):
                    cell = ws_resumen.cell(row=row, column=col_idx, value=val)
                    cell.fill = fill_gris
                    cell.border = thin_border
                    if col_idx == 1: cell.font, cell.alignment = font_label, align_left
                    elif col_idx == 2: cell.font, cell.alignment = font_metrica, align_center
                    elif col_idx == 3: cell.font, cell.alignment = font_label, align_center
                    else: cell.font, cell.alignment = font_rooms, align_left
                ws_resumen.row_dimensions[row].height = 25
            ws_resumen.column_dimensions["A"].width = 28
            ws_resumen.column_dimensions["B"].width = 12
            ws_resumen.column_dimensions["C"].width = 10
            ws_resumen.column_dimensions["D"].width = 45
            ws_resumen.column_dimensions["E"].width = 25

            def crear_hoja(wb, titulo, nombre_hoja, fill_color, df_data):
                ws = wb.create_sheet(nombre_hoja)
                ws.merge_cells("A1:O1")
                cell = ws["A1"]
                cell.value = titulo
                cell.fill, cell.font, cell.alignment = fill_color, font_titulo, align_center
                ws.row_dimensions[1].height = 25
                headers = ["ID", "ETA", "NAME", "QTY", "ROOM", "EMAIL", "CHECK IN", "CHECK OUT",
                           "RESERVATION", "PHONE", "INFO", "IRD", "HSK", "RATE", "TRANSPORTATION"]
                for col_idx, header in enumerate(headers, 1):
                    cell = ws.cell(row=2, column=col_idx, value=header)
                    cell.fill, cell.font, cell.alignment, cell.border = fill_header, font_header, align_center, thin_border
                for row_idx, (_, row_data) in enumerate(df_data.iterrows(), 3):
                    valores = [row_data.get(c, "") for c in ["id", "eta", "name", "qty", "room", "email", "check_in", "check_out",
                                                               "res_number", "phone", "info", "ird", "hsk", "rate", "trans"]]
                    for col_idx, valor in enumerate(valores, 1):
                        cell = ws.cell(row=row_idx, column=col_idx, value=valor if pd.notna(valor) else "")
                        cell.fill, cell.font, cell.alignment, cell.border = fill_gris, font_label, align_left, thin_border
                anchos = {"A": 6, "B": 10, "C": 22, "D": 6, "E": 8, "F": 25, "G": 10, "H": 10,
                          "I": 14, "J": 18, "K": 22, "L": 18, "M": 18, "N": 8, "O": 18}
                for col_letter, ancho in anchos.items(): ws.column_dimensions[col_letter].width = ancho
                return ws

            crear_hoja(wb, f"HABITACIONES EN CASA - {hoy_str}", "En Casa", fill_verde, df_en_casa)
            crear_hoja(wb, f"HABITACIONES QUE SALEN HOY - {hoy_str}", f"Salen Hoy {hoy_str}", fill_rojo, df_salen_hoy)
            crear_hoja(wb, f"HABITACIONES QUE SALEN MAÑANA - {manana_str}", f"Salen Mañana {manana_str}", fill_rojo, df_salen_manana)
            crear_hoja(wb, f"HABITACIONES QUE LLEGAN HOY - {hoy_str}", f"Llegan Hoy {hoy_str}", fill_azul, df_llegan_hoy)
            crear_hoja(wb, f"HABITACIONES QUE LLEGAN MAÑANA - {manana_str}", f"Llegan Mañana {manana_str}", fill_azul, df_llegan_manana)
            if total_vip_en_casa > 0: crear_hoja(wb, f"VIPs EN CASA - {hoy_str}", "VIPs En Casa", fill_verde, vip_en_casa)
            # HOJA RESUMEN ROOMS
            ws_rooms = wb.create_sheet("Resumen Rooms")
            ws_rooms.merge_cells("A1:D1")
            cell = ws_rooms["A1"]
            cell.value = f"RESUMEN DE HABITACIONES - {hoy_str}"
            cell.fill, cell.font, cell.alignment = fill_morado, font_titulo, align_center
            ws_rooms.row_dimensions[1].height = 25
            headers_rooms = ["CATEGORÍA", "TOTAL", "HABITACIONES", "FECHA"]
            for col_idx, header in enumerate(headers_rooms, 1):
                cell = ws_rooms.cell(row=2, column=col_idx, value=header)
                cell.fill, cell.font, cell.alignment, cell.border = fill_header, font_header, align_center, thin_border
            rooms_data_excel = [
                ("En Casa", total_en_casa, ", ".join(rooms_en_casa) if rooms_en_casa else "N/A", hoy_str),
                ("VIPs En Casa", total_vip_en_casa, ", ".join(rooms_vip_en_casa) if rooms_vip_en_casa else "N/A", hoy_str),
                ("Salen Hoy", total_salen_hoy, ", ".join(rooms_salen_hoy) if rooms_salen_hoy else "N/A", hoy_str),
                ("VIPs Salen Hoy", total_vip_salen_hoy, ", ".join(rooms_vip_salen_hoy) if rooms_vip_salen_hoy else "N/A", hoy_str),
                ("Salen Mañana", total_salen_manana, ", ".join(rooms_salen_manana) if rooms_salen_manana else "N/A", manana_str),
                ("VIPs Salen Mañana", total_vip_salen_manana, ", ".join(rooms_vip_salen_manana) if rooms_vip_salen_manana else "N/A", manana_str),
                ("Llegan Hoy", total_llegan_hoy, ", ".join(rooms_llegan_hoy) if rooms_llegan_hoy else "N/A", hoy_str),
                ("VIPs Llegan Hoy", total_vip_llegan_hoy, ", ".join(rooms_vip_llegan_hoy) if rooms_vip_llegan_hoy else "N/A", hoy_str),
                ("Llegan Mañana", total_llegan_manana, ", ".join(rooms_llegan_manana) if rooms_llegan_manana else "N/A", manana_str),
                ("VIPs Llegan Mañana", total_vip_llegan_manana, ", ".join(rooms_vip_llegan_manana) if rooms_vip_llegan_manana else "N/A", manana_str),
            ]
            for row_idx, (cat, tot, rms, fecha) in enumerate(rooms_data_excel, 3):
                ws_rooms.cell(row=row_idx, column=1, value=cat).fill = fill_gris
                ws_rooms.cell(row=row_idx, column=1).font = font_label
                ws_rooms.cell(row=row_idx, column=1).border = thin_border
                ws_rooms.cell(row=row_idx, column=2, value=tot).fill = fill_gris
                ws_rooms.cell(row=row_idx, column=2).font = font_metrica
                ws_rooms.cell(row=row_idx, column=2).alignment = align_center
                ws_rooms.cell(row=row_idx, column=2).border = thin_border
                ws_rooms.cell(row=row_idx, column=3, value=rms).fill = fill_gris
                ws_rooms.cell(row=row_idx, column=3).font = font_rooms
                ws_rooms.cell(row=row_idx, column=3).alignment = align_left
                ws_rooms.cell(row=row_idx, column=3).border = thin_border
                ws_rooms.cell(row=row_idx, column=4, value=fecha).fill = fill_gris
                ws_rooms.cell(row=row_idx, column=4).font = font_label
                ws_rooms.cell(row=row_idx, column=4).alignment = align_center
                ws_rooms.cell(row=row_idx, column=4).border = thin_border
            ws_rooms.column_dimensions["A"].width = 25
            ws_rooms.column_dimensions["B"].width = 10
            ws_rooms.column_dimensions["C"].width = 50
            ws_rooms.column_dimensions["D"].width = 12
            buffer = BytesIO()
            wb.save(buffer)
            buffer.seek(0)
            return buffer

        st.markdown("<div style='height: 15px;'></div>", unsafe_allow_html=True)
        try:
            excel_buffer = generar_reporte_excel()
            fecha_hoy = hoy.strftime("%Y%m%d_%H%M")
            st.download_button(label="📥 DESCARGAR REPORTE EXCEL", data=excel_buffer,
                file_name=f"Reporte_Ocupacion_{fecha_hoy}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_descargar_reporte")
        except Exception as e:
            st.error(f"❌ Error al generar el reporte: {str(e)}")
            st.exception(e)

# ============================================================
# FORMULARIO CARTA DE DESPEDIDA
# ============================================================
if st.query_params.get("action") == "carta":
    with st.container():
        st.subheader("✉️ Carta de Despedida")
        col1, col2 = st.columns([1, 1])
        with col1:
            if st.button("↩️ REGRESAR", key="regresar_carta"):
                st.query_params.clear(); st.rerun()
        fila_guardada = st.session_state.get("fila_seleccionada")
        if not fila_guardada:
            st.error("❌ Por favor, selecciona una fila en la tabla primero.")
            st.info("💡 Ve a la tabla principal, haz clic en la reserva del huésped, y luego presiona ✉️ CARTA.")
            if st.button("↩️ REGRESAR A LA TABLA", key="regresar_carta_error"):
                st.query_params.clear(); st.rerun()
        else:
            nombre_huesped = str(fila_guardada.get("name", "")).strip()
            room_huesped = str(fila_guardada.get("room", "")).strip()
            check_out_huesped = str(fila_guardada.get("check_out", "")).strip()
            if not nombre_huesped:
                st.error("❌ La reserva seleccionada no tiene nombre de huésped.")
            else:
                st.markdown(f"""<div style="background-color: #1a1a2e; border-radius: 8px; padding: 15px; margin: 10px 0; border: 1px solid #00E5FF;">
                    <p style="color: #ccc; font-size: 0.9rem; margin: 0;">👤 <b>Huésped:</b> <span style="color: #00E5FF; font-size: 1.1rem;">{nombre_huesped}</span><br>
                    🚪 <b>Habitación:</b> <span style="color: #00E5FF;">{room_huesped if room_huesped else 'N/A'}</span><br>
                    📅 <b>Check-out:</b> <span style="color: #00E5FF;">{check_out_huesped if check_out_huesped else 'N/A'}</span></p></div>""", unsafe_allow_html=True)
                plantilla_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "plantilla_despedida.docx")
                if not os.path.exists(plantilla_path):
                    st.warning("⚠️ No se encontró la plantilla `plantilla_despedida.docx`")
                    st.info(f"📁 La plantilla debe estar en: `{os.path.dirname(os.path.abspath(__file__))}`")
                    st.markdown("""<div style="background-color: #1a1a2e; border-radius: 8px; padding: 15px; margin: 10px 0; border: 1px solid #333;">
                        <p style="color: #ccc; font-size: 0.85rem; margin: 0;"><b>📝 Instrucciones para crear la plantilla:</b><br><br>
                        1. Crea un archivo Word llamado <code>plantilla_despedida.docx</code><br>
                        2. Escribe el texto de tu carta de despedida<br>
                        3. Donde quieras que aparezca el nombre del huésped, escribe: <code>{{NOMBRE}}</code><br>
                        4. Guarda el archivo en la misma carpeta donde está este script<br><br>
                        <b>Ejemplo de texto en la plantilla:</b><br>
                        <i>"Dear {{NOMBRE}},<br><br>It was an absolute pleasure having you with us..."</i></p></div>""", unsafe_allow_html=True)
                else:
                    st.success(f"✅ Plantilla encontrada: `plantilla_despedida.docx`")
                    try:
                        from docx import Document
                        doc = Document(plantilla_path)
                        def reemplazar_en_parrafo(paragraph, placeholder, replacement):
                            full_text = paragraph.text
                            if placeholder in full_text:
                                new_text = full_text.replace(placeholder, replacement)
                                for run in paragraph.runs: run.text = ""
                                if paragraph.runs: paragraph.runs[0].text = new_text
                                else: paragraph.add_run(new_text)
                                return True
                            return False
                        placeholders_encontrados = 0
                        for paragraph in doc.paragraphs:
                            if reemplazar_en_parrafo(paragraph, "{{NOMBRE}}", nombre_huesped): placeholders_encontrados += 1
                        for table in doc.tables:
                            for row in table.rows:
                                for cell in row.cells:
                                    for paragraph in cell.paragraphs:
                                        if reemplazar_en_parrafo(paragraph, "{{NOMBRE}}", nombre_huesped): placeholders_encontrados += 1
                        buffer = BytesIO()
                        doc.save(buffer)
                        buffer.seek(0)
                        st.markdown("""<div style="background-color: #0d1f0d; border-radius: 8px; padding: 10px; margin-bottom: 10px; border: 1px solid #2e7d32;">
                            <div style="color: #4CAF50; font-size: 0.8rem; font-weight: bold; text-align: center;">✅ CARTA GENERADA CORRECTAMENTE</div></div>""", unsafe_allow_html=True)
                        if placeholders_encontrados > 0: st.info(f"📌 Se reemplazaron {placeholders_encontrados} placeholder(s) `{{{{NOMBRE}}}}` con: **{nombre_huesped}**")
                        else: st.warning("⚠️ No se encontró el placeholder `{{NOMBRE}}` en la plantilla. El documento se descargará sin modificaciones.")
                        nombre_archivo = f"Carta_Despedida_{nombre_huesped.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
                        st.download_button(label="📥 DESCARGAR CARTA DE DESPEDIDA", data=buffer,
                            file_name=nombre_archivo,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key="btn_descargar_carta")
                        st.markdown("""<div style="background-color: #1a1a2e; border-radius: 8px; padding: 10px; margin-top: 10px; border: 1px solid #333;">
                            <p style="color: #888; font-size: 0.75rem; margin: 0; text-align: center;">💡 Descarga el archivo y ábrelo en Word para imprimir.<br>El nombre del huésped ya está insertado en el documento.</p></div>""", unsafe_allow_html=True)
                    except ImportError:
                        st.error("❌ Falta instalar la librería `python-docx`")
                        st.code("pip install python-docx", language="bash")
                        st.info("💡 Ejecuta el comando anterior en tu terminal y reinicia la aplicación.")
                    except Exception as e:
                        st.error(f"❌ Error al procesar la plantilla: {str(e)}")
                        st.exception(e)

# ============================================================
# FORMULARIO NUEVA RESERVA
# ============================================================
if mostrar_formulario:
    with st.container():
        st.subheader("📝 Nueva Reservación")
        if st.button("↩️ REGRESAR", key="regresar_nueva"):
            st.query_params.clear(); st.rerun()
        with st.form("form_reserva"):
            c1, c2, c3, c4 = st.columns(4)
            eta = c1.text_input("ETA", value=hora_actual_12h())
            name = c2.text_input("Name")
            qty = c3.number_input("Qty", min_value=0, value=0)
            room = c4.text_input("Room")
            c5, c6, c7, c8 = st.columns(4)
            email = c5.text_input("Email")
            check_in = c6.date_input("Check-in")
            check_out = c7.date_input("Check-out")
            res_number = c8.text_input("Reservation #")
            c9, c10, c11, c12 = st.columns(4)
            phone = c9.text_input("Phone")
            info = c10.text_input("Information")
            ird = c11.text_input("IRD")
            hsk = c12.text_input("HSK")
            c13, c14 = st.columns(2)
            rate = c13.text_input("Rate")
            trans = c14.text_input("Transportation")
            if st.form_submit_button("Guardar Reservación"):
                data = {
                    "eta": eta, "name": name, "qty": qty, "room": room,
                    "email": email, "check_in": check_in.strftime("%b %d"),
                    "check_out": check_out.strftime("%b %d"), "res_number": res_number,
                    "phone": phone, "info": info, "ird": ird, "hsk": hsk,
                    "rate": rate, "trans": trans
                }
                insertar_reserva(data)
                st.query_params.clear(); st.rerun()

# ============================================================
# FORMULARIO EDITAR RESERVA
# ============================================================
if mostrar_editar:
    fila_guardada = st.session_state.get("fila_seleccionada")
    if fila_guardada:
        with st.container():
            st.subheader("✏️ Editar Reservación")
            if st.button("↩️ REGRESAR", key="regresar_editar"):
                st.query_params.clear(); st.rerun()
            try: check_in_dt = datetime.strptime(fila_guardada.get("check_in", ""), "%b %d").replace(year=datetime.now().year)
            except: check_in_dt = datetime.now()
            try: check_out_dt = datetime.strptime(fila_guardada.get("check_out", ""), "%b %d").replace(year=datetime.now().year)
            except: check_out_dt = datetime.now()
            with st.form("form_editar"):
                c1, c2, c3, c4 = st.columns(4)
                eta = c1.text_input("ETA", value=str(fila_guardada.get("eta", "")))
                name = c2.text_input("Name", value=fila_guardada.get("name", ""))
                qty_raw = fila_guardada.get("qty", 0)
                try:
                    if qty_raw is None or str(qty_raw).strip() == "":
                        qty_val = 0
                    else:
                        qty_val = int(float(qty_raw))
                except (ValueError, TypeError):
                    qty_val = 0
                qty = c3.number_input("Qty", min_value=0, value=qty_val)
                room = c4.text_input("Room", value=fila_guardada.get("room", ""))
                c5, c6, c7, c8 = st.columns(4)
                email = c5.text_input("Email", value=fila_guardada.get("email", ""))
                check_in = c6.date_input("Check-in", value=check_in_dt)
                check_out = c7.date_input("Check-out", value=check_out_dt)
                res_number = c8.text_input("Reservation #", value=fila_guardada.get("res_number", ""))
                c9, c10, c11, c12 = st.columns(4)
                phone = c9.text_input("Phone", value=fila_guardada.get("phone", ""))
                info = c10.text_input("Information", value=fila_guardada.get("info", ""))
                ird = c11.text_input("IRD", value=fila_guardada.get("ird", ""))
                hsk = c12.text_input("HSK", value=fila_guardada.get("hsk", ""))
                c13, c14 = st.columns(2)
                rate = c13.text_input("Rate", value=fila_guardada.get("rate", ""))
                trans = c14.text_input("Transportation", value=fila_guardada.get("trans", ""))
                if st.form_submit_button("💾 Guardar Cambios"):
                    data = {
                        "eta": eta, "name": name, "qty": qty, "room": room,
                        "email": email, "check_in": check_in.strftime("%b %d"),
                        "check_out": check_out.strftime("%b %d"), "res_number": res_number,
                        "phone": phone, "info": info, "ird": ird, "hsk": hsk,
                        "rate": rate, "trans": trans
                    }
                    actualizar_reserva(fila_guardada["id"], data)
                    st.session_state.pop("fila_seleccionada", None)
                    st.query_params.clear()
                    st.success("Reserva actualizada correctamente.")
                    st.rerun()
    else:
        st.error("Por favor, selecciona una fila en la tabla primero.")
        if st.button("↩️ REGRESAR", key="regresar_editar_error"):
            st.query_params.clear(); st.rerun()

# ============================================================
# TABLA ÚNICA (SIEMPRE VISIBLE) CON BÚSQUEDA INTELIGENTE
# ============================================================
st.markdown("<div style='height: 5px;'></div>", unsafe_allow_html=True)
df_reservas = cargar_reservaciones()
if filtro_checkout:
    df_reservas = df_reservas[df_reservas["check_out"] == filtro_checkout]
    st.info(f"📅 Mostrando reservas que salen el: {filtro_checkout}")
if fecha_filtro_activo and filtro_fecha_date:
    fecha_filtro = datetime.strptime(filtro_fecha_date, "%Y-%m-%d").strftime("%b %d")
    df_reservas = df_reservas[df_reservas["check_in"] == fecha_filtro]
    st.info(f"📅 Mostrando reservas con check-in el: {fecha_filtro}")
if busqueda and busqueda.strip():
    busqueda_lower = busqueda.strip().lower()
    mask = df_reservas.astype(str).apply(lambda row: row.str.lower().str.contains(busqueda_lower, na=False).any(), axis=1)
    df_reservas = df_reservas[mask]
    if len(df_reservas) == 0: st.info(f"🔍 No se encontraron resultados para: '{busqueda}'")

# TABLA CON SELECCIÓN NATIVA DE STREAMLIT
st.markdown("<div style='height: 5px;'></div>", unsafe_allow_html=True)

# Configurar columnas para mejor visualización
column_config = {
    "id": st.column_config.NumberColumn("ID", width="small"),
    "eta": st.column_config.TextColumn("ETA", width="small"),
    "name": st.column_config.TextColumn("NAME", width="large"),
    "qty": st.column_config.NumberColumn("QTY", width="small"),
    "room": st.column_config.TextColumn("ROOM", width="small"),
    "email": st.column_config.TextColumn("EMAIL", width="medium"),
    "check_in": st.column_config.TextColumn("CHECK IN", width="small"),
    "check_out": st.column_config.TextColumn("CHECK OUT", width="small"),
    "res_number": st.column_config.TextColumn("RESERVATION", width="medium"),
    "phone": st.column_config.TextColumn("PHONE", width="medium"),
    "info": st.column_config.TextColumn("INFO", width="large"),
    "ird": st.column_config.TextColumn("IRD", width="medium"),
    "hsk": st.column_config.TextColumn("HSK", width="medium"),
    "rate": st.column_config.TextColumn("RATE", width="small"),
    "trans": st.column_config.TextColumn("TRANS", width="medium"),
}

# Mostrar tabla con selección de fila
seleccion = st.dataframe(
    df_reservas,
    column_config=column_config,
    use_container_width=True,
    height=620,
    selection_mode="single-row",
    on_select="rerun",
    key="tabla_principal_concierge"
)

# Guardar fila seleccionada en session_state
if seleccion and seleccion.selection.rows:
    idx = seleccion.selection.rows[0]
    st.session_state["fila_seleccionada"] = df_reservas.iloc[idx].to_dict()
else:
    st.session_state.pop("fila_seleccionada", None)

# ============================================================
# LÓGICA DE BORRADO CON CLAVE
# ============================================================
if st.query_params.get("action") == "cancelar":
    st.subheader("❌ Cancelar Reservación")
    if st.button("↩️ REGRESAR", key="regresar_cancelar"):
        st.query_params.clear(); st.rerun()
    fila_guardada = st.session_state.get("fila_seleccionada")
    if fila_guardada:
        id_a_borrar = fila_guardada["id"]
        st.warning(f"Se eliminará permanentemente la reserva ID: {id_a_borrar}")
        password = st.text_input("Ingrese clave de autorización:", type="password")
        if st.button("CONFIRMAR Y BORRAR"):
            if password == "D6msnp8a":
                eliminar_reserva(id_a_borrar)
                st.session_state.pop("fila_seleccionada", None)
                st.query_params.clear()
                st.success("Registro eliminado correctamente.")
                st.rerun()
            else: st.error("Clave incorrecta.")
    else:
        st.error("Por favor, selecciona una fila en la tabla primero.")
        if st.button("↩️ REGRESAR", key="regresar_cancelar_error"):
            st.query_params.clear(); st.rerun()
