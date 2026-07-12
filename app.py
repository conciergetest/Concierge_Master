import streamlit as st
import streamlit.components.v1 as components
import sqlite3
import pandas as pd
import os
from st_aggrid import AgGrid, GridOptionsBuilder
from datetime import datetime, timedelta
from io import BytesIO

# 1. CONFIGURACIÓN INICIAL
st.set_page_config(page_title="Concierge Master v5.1", layout="wide", initial_sidebar_state="collapsed")
db_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "recepcion_final.db")

query_params = st.query_params
mostrar_formulario = query_params.get("action") == "nueva"
mostrar_editar = query_params.get("action") == "editar"
mostrar_importar = query_params.get("action") == "importar"
mostrar_exportar = query_params.get("action") == "exportar"

# Variable para filtrar por fecha de checkout desde el panel
filtro_checkout = query_params.get("checkout_filtro")
# Variable para filtrar por fecha del date picker - SOLO se activa cuando el usuario hace clic en "APLICAR"
filtro_fecha_date = query_params.get("fecha_date")
# Variable para saber si el filtro de fecha está ACTIVAMENTE activado
fecha_filtro_activo = query_params.get("fecha_activa") == "true"

def cargar_reservaciones():
    conn = sqlite3.connect(db_path)
    df = pd.read_sql_query("SELECT * FROM huespedes", conn)
    conn.close()
    df['check_in_dt'] = pd.to_datetime(df['check_in'], format='%b %d', errors='coerce')
    df = df.sort_values(by=['check_in_dt', 'name'])
    return df.drop(columns=['check_in_dt'])

# ============================================================
# GENERAR LISTA DE HORAS EN FORMATO 12H (AM/PM)
# ============================================================
horas_eta_12h = []
horas_eta_24h = []  # Para guardar en la base de datos
for h in range(24):
    for m in [0, 30]:
        hora_24 = f"{h:02d}:{m:02d}"
        # Formato AM/PM para mostrar
        if h == 0:
            hora_12 = f"12:{m:02d} AM"
        elif h < 12:
            hora_12 = f"{h}:{m:02d} AM"
        elif h == 12:
            hora_12 = f"12:{m:02d} PM"
        else:
            hora_12 = f"{h-12}:{m:02d} PM"
        horas_eta_12h.append(hora_12)
        horas_eta_24h.append(hora_24)

# Diccionario para convertir 12h -> 24h
mapa_12a24 = dict(zip(horas_eta_12h, horas_eta_24h))
mapa_24a12 = dict(zip(horas_eta_24h, horas_eta_12h))

# ============================================================
# FUNCION PARA EXPORTAR A EXCEL CON COLORES POR CATEGORÍA
# ============================================================
def exportar_excel_por_categorias(df):
    """
    Exporta el DataFrame a un Excel organizado por categorías con colores.
    Cada categoría tiene su propia sección con encabezado de color.
    """
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

    # Color de fondo para las filas de datos (gris muy claro)
    fill_data = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    # Color de fondo para el encabezado de la sección (azul cyan)
    fill_header = PatternFill(start_color="00B0F0", end_color="00B0F0", fill_type="solid")
    # Color de fondo para los headers de columnas (gris oscuro)
    fill_col_header = PatternFill(start_color="404040", end_color="404040", fill_type="solid")

    # Fuentes
    font_header = Font(name='Calibri', size=11, bold=True, color="FFFFFF")
    font_section = Font(name='Calibri', size=12, bold=True, color="FFFFFF")
    font_data = Font(name='Calibri', size=10, color="000000")

    # Alineaciones
    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)

    # Bordes
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    wb = Workbook()
    ws = wb.active
    ws.title = "Arrivals"

    # Definir categorías y sus keywords
    categorias_config = [
        ("CUMPLEAÑOS", ["BIRTHDAY", "CUMPLE", "BDAY"]),
        ("VIP", ["VIP"]),
        ("HONEYMOON", ["HONEYMOON", "LUNA DE MIEL"]),
        ("ANNIVERSARY", ["ANNIVERSARY", "ANIVERSARIO"]),
        ("BABYMOON", ["BABYMOON"]),
        ("TEAM MEMBER", ["TEAM MEMBER", "STAFF", "EMPLOYEE"]),
        ("GENERAL", [])  # El resto
    ]

    # Columnas del Excel
    columnas_excel = ['ID', 'ETA', 'NAME', 'QTY', 'ROOM', 'EMAIL', 'CHECK IN', 'CHECK OUT', 
                      'RESERVATION', 'PHONE', 'FORMATIO', 'IRD', 'HSK', 'RATE', 'TRANSPORTATION']

    # Mapeo de columnas de la DB a columnas del Excel
    mapeo_cols = {
        'id': 'ID',
        'eta': 'ETA',
        'name': 'NAME',
        'qty': 'QTY',
        'room': 'ROOM',
        'email': 'EMAIL',
        'check_in': 'CHECK IN',
        'check_out': 'CHECK OUT',
        'res_number': 'RESERVATION',
        'phone': 'PHONE',
        'info': 'FORMATIO',
        'ird': 'IRD',
        'hsk': 'HSK',
        'rate': 'RATE',
        'trans': 'TRANSPORTATION'
    }

    current_row = 1

    # Agrupar filas por categoría
    filas_por_categoria = {}
    filas_general = []

    for idx, row in df.iterrows():
        info_str = str(row.get('info', '')).upper()
        asignada = False

        for cat_nombre, keywords in categorias_config[:-1]:  # Excluir GENERAL
            for kw in keywords:
                if kw in info_str:
                    if cat_nombre not in filas_por_categoria:
                        filas_por_categoria[cat_nombre] = []
                    filas_por_categoria[cat_nombre].append(row)
                    asignada = True
                    break
            if asignada:
                break

        if not asignada:
            filas_general.append(row)

    # Procesar cada categoría
    for cat_nombre, keywords in categorias_config:
        if cat_nombre == "GENERAL":
            filas = filas_general
        else:
            filas = filas_por_categoria.get(cat_nombre, [])

        if not filas:
            continue

        # Fila de título de categoría
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=len(columnas_excel))
        cell = ws.cell(row=current_row, column=1, value=cat_nombre)
        cell.fill = fill_header
        cell.font = font_section
        cell.alignment = align_center
        current_row += 1

        # Fila de headers de columnas
        for col_idx, col_name in enumerate(columnas_excel, 1):
            cell = ws.cell(row=current_row, column=col_idx, value=col_name)
            cell.fill = fill_col_header
            cell.font = font_header
            cell.alignment = align_center
            cell.border = thin_border
        current_row += 1

        # Filas de datos
        for row_data in filas:
            for col_idx, col_db in enumerate(mapeo_cols.keys(), 1):
                valor = row_data.get(col_db, '')
                if pd.isna(valor):
                    valor = ''
                cell = ws.cell(row=current_row, column=col_idx, value=valor)
                cell.fill = fill_data
                cell.font = font_data
                cell.alignment = align_left
                cell.border = thin_border
            current_row += 1

        # Fila en blanco entre categorías
        current_row += 1

    # Ajustar anchos de columna
    anchos = {
        'A': 6,   # ID
        'B': 10,  # ETA
        'C': 22,  # NAME
        'D': 6,   # QTY
        'E': 8,   # ROOM
        'F': 25,  # EMAIL
        'G': 10,  # CHECK IN
        'H': 10,  # CHECK OUT
        'I': 14,  # RESERVATION
        'J': 18,  # PHONE
        'K': 22,  # FORMATIO
        'L': 18,  # IRD
        'M': 18,  # HSK
        'N': 8,   # RATE
        'O': 18,  # TRANSPORTATION
    }
    for col_letter, ancho in anchos.items():
        ws.column_dimensions[col_letter].width = ancho

    # Congelar paneles
    ws.freeze_panes = 'A1'

    # Guardar en buffer
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# CSS GLOBAL
st.markdown("""
<style>
/* Ocultar header por defecto */
header[data-testid="stHeader"] { display: none !important; }

/* Reducir espacio superior */
.block-container { padding-top: 0.1rem !important; padding-bottom: 0.1rem !important; }

/* Reducir gap entre columnas de botones */
div[data-testid="stHorizontalBlock"] { gap: 0.2rem !important; }

/* Reducir márgenes entre elementos */
div[data-testid="stVerticalBlock"] > div { margin-bottom: 0.1rem !important; }

/* Estilo base de botones */
div.stButton > button {
    width: 100%;
    border-radius: 6px;
    font-weight: bold;
    border: none;
    font-size: 0.65rem;
    padding: 4px 2px;
    white-space: nowrap;
    min-height: 28px;
}

/* Colores del HEADER - solo botones dentro del primer horizontal block que están en header_col2 */
#root > div > div > div > div > div > div[data-testid="stHorizontalBlock"]:nth-of-type(1) > div:nth-child(2) button { background: #00E5FF !important; color: black !important; }

/* CHECKOUT BUTTONS - fondo negro con letra blanca, más compactos */
button[key^="checkout_btn_"] {
    background-color: #000000 !important;
    color: #ffffff !important;
    font-weight: bold !important;
    font-size: 0.7rem !important;
    border: 1px solid #333 !important;
    border-radius: 6px !important;
    text-align: center !important;
    padding: 2px 3px !important;
    min-height: 22px !important;
    margin: 0 !important;
}
/* Botón VER TODAS - fondo negro con letra blanca */
button[key="btn_ver_todas"] {
    background-color: #000000 !important;
    color: #ffffff !important;
    font-weight: bold !important;
    font-size: 0.7rem !important;
    border: 1px solid #333 !important;
    border-radius: 6px !important;
    text-align: center !important;
    padding: 2px 3px !important;
    min-height: 22px !important;
}

/* Reducir espacio entre filas de botones de checkout */
div[data-testid="stHorizontalBlock"] button[key^="checkout_btn_"] {
    margin-top: 1px !important;
    margin-bottom: 1px !important;
}

/* Estilo del buscador */
div[data-testid="stTextInput"] > div > div > input {
    background-color: #1a1a2e !important;
    color: white !important;
    border: 1px solid #333 !important;
    border-radius: 8px !important;
    padding: 6px 10px !important;
    font-size: 0.85rem !important;
}

/* Reducir espacio del label del buscador */
div[data-testid="stTextInput"] label { margin-bottom: 0 !important; font-size: 0.75rem !important; }

/* Estilo del date input */
div[data-testid="stDateInput"] > div > div > input {
    background-color: #1a1a2e !important;
    color: white !important;
    border: 1px solid #333 !important;
    border-radius: 8px !important;
    padding: 6px 10px !important;
    font-size: 0.85rem !important;
}
div[data-testid="stDateInput"] label { 
    color: #888 !important; 
    font-size: 0.75rem !important; 
    margin-bottom: 2px !important;
}

/* Botón APLICAR filtro de fecha */
button[key="btn_aplicar_fecha"] {
    background-color: #00E5FF !important;
    color: #000000 !important;
    font-weight: bold !important;
    font-size: 0.7rem !important;
    border: none !important;
    border-radius: 6px !important;
    padding: 4px 8px !important;
    min-height: 28px !important;
}
button[key="btn_limpiar_fecha"] {
    background-color: #333333 !important;
    color: #ffffff !important;
    font-weight: bold !important;
    font-size: 0.7rem !important;
    border: 1px solid #555 !important;
    border-radius: 6px !important;
    padding: 4px 8px !important;
    min-height: 28px !important;
}

/* Botón IMPORTAR EXCEL */
button[key="btn_procesar_excel"] {
    background-color: #00E5FF !important;
    color: #000000 !important;
    font-weight: bold !important;
    font-size: 0.85rem !important;
    border: none !important;
    border-radius: 8px !important;
    padding: 10px 20px !important;
    min-height: 40px !important;
}

/* Tabla */
.ag-root-wrapper { background-color: #101010 !important; }
.ag-cell { color: white !important; background-color: #101010 !important; }
.ag-row-selected .ag-cell { background-color: #00FFFF !important; color: #000000 !important; font-weight: bold !important; }
</style>
""", unsafe_allow_html=True)

# ===== HEADER: TÍTULO + BOTONES =====
header_col1, header_col2 = st.columns([1.3, 8.7])

with header_col1:
    st.markdown("""
    <h2 style="color:#00E5FF; margin:0; padding:0; font-size:1.1rem; line-height:1.0;">
        🛎️ Concierge<br>Master v5.1
    </h2>
    """, unsafe_allow_html=True)

with header_col2:
    c1, c2, c3, c4, c5, c6, c7, c8 = st.columns(8)
    if c1.button("➕ NUEVA", key="btn_nueva"): 
        st.query_params["action"] = "nueva"; st.rerun()
    if c2.button("✏️ EDITAR", key="btn_editar"): 
        st.query_params["action"] = "editar"; st.rerun()
    if c3.button("📥 IMPORTAR", key="btn_importar"): 
        st.query_params["action"] = "importar"; st.rerun()
    if c4.button("📤 EXPORTAR", key="btn_exportar"): 
        st.query_params["action"] = "exportar"; st.rerun()
    if c5.button("✉️ CARTA", key="btn_carta"): 
        st.query_params["action"] = "carta"; st.rerun()
    if c6.button("❌ BORRAR", key="btn_cancelar"): 
        st.query_params["action"] = "cancelar"; st.rerun()
    if c7.button("📄 REPORTE", key="btn_reporte"): 
        st.query_params["action"] = "reporte"; st.rerun()
    if c8.button("📅 AGENDA", key="btn_agenda"): 
        st.switch_page("pages/agenda.py")

# ===== CARGAR DATOS =====
df_todas = cargar_reservaciones()
total_reservas = len(df_todas)

# ===== FILA 1: CONTADOR + BUSCADOR (columna izquierda) =====
search_col1, search_col2 = st.columns([2.5, 8.5])

with search_col1:
    # CONTADOR DE RESERVAS TOTALES
    st.markdown(f"""
    <div style="background-color: #1a1a2e; border-radius: 8px; padding: 6px 10px; margin-bottom: 4px; text-align: center; border: 1px solid #00E5FF;">
        <span style="color: #00E5FF; font-size: 0.75rem; font-weight: bold;">📊 TOTAL RESERVAS:</span>
        <span style="color: #ffffff; font-size: 1.0rem; font-weight: bold;"> {total_reservas}</span>
    </div>
    """, unsafe_allow_html=True)

    # Buscador
    st.markdown("<p style='color:#888; font-size:0.75rem; margin:0; padding:0;'>🔍 Búsqueda rápida...</p>", unsafe_allow_html=True)
    busqueda = st.text_input("", placeholder="Buscar por nombre, teléfono, reserva, VIP, Relaxury...", label_visibility="collapsed", key="buscador_global")

with search_col2:
    pass

# ===== FILA 2: CHECKING OUT ROOMS + GRÁFICO CATEGORÍAS =====
left_col, right_col = st.columns([5.0, 5.0])

with left_col:
    # Panel Checking Out Rooms
    st.markdown("""
    <div style="background-color: #1a1a2e; border-radius: 8px; padding: 5px 8px 2px 8px; margin-bottom: 1px;">
        <div style="color: #ffffff; font-size: 0.8rem; font-weight: bold; text-align: center; margin-bottom: 3px;">
            🏨 Checking Out Rooms
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Calcular fechas desde hoy hasta 7 días después (8 días total)
    hoy = datetime.now()
    fechas_checkout = []
    for i in range(8):
        fecha = hoy + timedelta(days=i)
        fechas_checkout.append(fecha)

    # Crear botones en 2 columnas para cada fecha
    for i in range(0, len(fechas_checkout), 2):
        cols = st.columns(2)
        for j in range(2):
            if i + j < len(fechas_checkout):
                fecha = fechas_checkout[i + j]
                mes_abr = fecha.strftime("%b").upper()
                dia_num = fecha.strftime("%d")
                fecha_db = fecha.strftime("%b %d")
                count = len(df_todas[df_todas['check_out'] == fecha_db])

                with cols[j]:
                    btn_label = f"{dia_num}-{mes_abr}: [{count}]"
                    if st.button(btn_label, key=f"checkout_btn_{fecha_db}", use_container_width=True):
                        # Al hacer clic en un checkout, DESACTIVAMOS el filtro de fecha
                        st.query_params["checkout_filtro"] = fecha_db
                        st.query_params.pop("fecha_date", None)
                        st.query_params.pop("fecha_activa", None)
                        st.rerun()

    # Botón para ver todas las reservas
    st.markdown("<div style='height: 1px;'></div>", unsafe_allow_html=True)
    if st.button("🔄 VER TODAS", key="btn_ver_todas", use_container_width=True):
        st.query_params.pop("checkout_filtro", None)
        st.query_params.pop("fecha_date", None)
        st.query_params.pop("fecha_activa", None)
        st.rerun()

    # ============================================================
    # FILTRO POR FECHA (DATE PICKER) - Solo se activa con botón APLICAR
    # ============================================================
    st.markdown("<div style='height: 4px;'></div>", unsafe_allow_html=True)

    # Obtener fecha actual del query param si existe
    fecha_default = datetime.now()
    if filtro_fecha_date:
        try:
            fecha_default = datetime.strptime(filtro_fecha_date, "%Y-%m-%d")
        except:
            fecha_default = datetime.now()

    fecha_seleccionada = st.date_input(
        "📅 DATE",
        value=fecha_default,
        key="date_filter_picker",
        label_visibility="visible"
    )

    # Botones para APLICAR o LIMPIAR el filtro de fecha
    col_aplicar, col_limpiar = st.columns(2)
    with col_aplicar:
        if st.button("🔍 APLICAR", key="btn_aplicar_fecha", use_container_width=True):
            # Solo al hacer clic en APLICAR se activa el filtro
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

with right_col:
    # ============================================================
    # GRÁFICO DE CATEGORÍAS (INFO COLUMN) - SIN RELAXURY
    # ============================================================
    categorias = {
        "VIP": "#00E5FF",
        "ANNIVERSARY": "#4CAF50",
        "BIRTHDAY": "#FF5252",
        "HONEYMOON": "#FF9800",
        "BABYMOON": "#9C27B0",
        "TEAM MEMBER": "#FFC107",
        "LEISURE": "#2196F3"
    }

    conteo_categorias = {}
    for cat in categorias:
        if cat == "LEISURE":
            continue
        mask = df_todas['info'].astype(str).str.upper().str.contains(cat, na=False)
        conteo_categorias[cat] = mask.sum()

    total_categorizadas = sum(conteo_categorias.values())
    conteo_categorias["LEISURE"] = max(0, total_reservas - total_categorizadas)
    conteo_ordenado = dict(sorted(conteo_categorias.items(), key=lambda x: x[1], reverse=True))
    max_valor = max(conteo_ordenado.values()) if conteo_ordenado else 1

    html_chart = """
    <!DOCTYPE html>
    <html>
    <head>
        <style>
            * { margin: 0; padding: 0; box-sizing: border-box; }
            body { background-color: #0d0d0d; font-family: 'Segoe UI', sans-serif; }
            .chart-container { 
                background-color: #0d0d0d; 
                border-radius: 12px; 
                padding: 12px 15px; 
                min-height: 220px;
                color: white;
            }
            .chart-header {
                display: flex;
                justify-content: space-between;
                align-items: center;
                margin-bottom: 10px;
            }
            .chart-title {
                color: #888;
                font-size: 10px;
                font-weight: bold;
                letter-spacing: 1px;
                text-transform: uppercase;
            }
            .total-circle {
                width: 42px;
                height: 42px;
                border: 2px solid #00E5FF;
                border-radius: 50%;
                display: flex;
                align-items: center;
                justify-content: center;
            }
            .total-number {
                color: #00E5FF;
                font-size: 14px;
                font-weight: bold;
            }
            .bar-row {
                display: flex;
                align-items: center;
                margin-bottom: 5px;
            }
            .bar-label {
                width: 90px;
                color: #ccc;
                font-size: 9px;
                text-align: right;
                padding-right: 8px;
                white-space: nowrap;
            }
            .bar-track {
                flex: 1;
                background-color: #1a1a1a;
                border-radius: 3px;
                height: 16px;
                position: relative;
                overflow: hidden;
            }
            .bar-fill {
                height: 100%;
                border-radius: 3px;
                transition: width 0.5s ease;
            }
            .bar-value {
                width: 28px;
                color: #fff;
                font-size: 11px;
                font-weight: bold;
                text-align: right;
                padding-left: 8px;
            }
        </style>
    </head>
    <body>
        <div class="chart-container">
            <div class="chart-header">
                <div class="chart-title">Guest Categories</div>
                <div class="total-circle">
                    <div class="total-number">""" + str(total_reservas) + """</div>
                </div>
            </div>
    """

    for cat, valor in conteo_ordenado.items():
        color = categorias.get(cat, "#888")
        porcentaje = (valor / max_valor * 100) if max_valor > 0 else 0
        html_chart += f"""
            <div class="bar-row">
                <div class="bar-label">{cat}</div>
                <div class="bar-track">
                    <div class="bar-fill" style="width: {porcentaje}%; background-color: {color};"></div>
                </div>
                <div class="bar-value">{valor}</div>
            </div>
        """

    html_chart += """
        </div>
    </body>
    </html>
    """

    components.html(html_chart, height=250, scrolling=False)

    # CONTADOR DE RELAXURY
    mask_relaxury = df_todas.astype(str).apply(
        lambda row: row.str.upper().str.contains('RELAXURY', na=False).any(), axis=1
    )
    total_relaxury = mask_relaxury.sum()

    st.markdown(f"""
    <div style="background-color: #1a1a2e; border-radius: 8px; padding: 6px 10px; margin-top: 2px; text-align: center; border: 1px solid #E91E63;">
        <span style="color: #E91E63; font-size: 0.75rem; font-weight: bold;">🏖️ RELAXURY:</span>
        <span style="color: #ffffff; font-size: 1.0rem; font-weight: bold;"> {total_relaxury}</span>
    </div>
    """, unsafe_allow_html=True)

# ============================================================
# 2. FORMULARIO IMPORTAR DESDE EXCEL
# ============================================================
if mostrar_importar:
    with st.container():
        st.subheader("📥 Importar Reservaciones desde Excel")

        col1, col2 = st.columns([1, 1])
        with col1:
            if st.button("↩️ REGRESAR", key="regresar_importar"):
                st.query_params.clear()
                st.rerun()

        st.markdown("""
        <div style="background-color: #1a1a2e; border-radius: 8px; padding: 15px; margin: 10px 0; border: 1px solid #333;">
            <p style="color: #ccc; font-size: 0.85rem; margin: 0;">
                📋 <b>Instrucciones:</b> Selecciona tu archivo <code>Plantilla_Importar.xlsx</code>. 
                El archivo debe contener las columnas: <b>eta, name, qty, room, email, check_in, check_out, res_number, phone, info, ird, hsk, rate, trans</b>
            </p>
        </div>
        """, unsafe_allow_html=True)

        # File uploader
        archivo_subido = st.file_uploader(
            "📁 Seleccionar archivo Excel",
            type=["xlsx", "xls"],
            key="uploader_excel"
        )

        if archivo_subido is not None:
            try:
                # Leer el archivo Excel
                df_excel = pd.read_excel(archivo_subido)

                st.markdown(f"""
                <div style="background-color: #0d1f0d; border-radius: 8px; padding: 10px 15px; margin: 10px 0; border: 1px solid #2e7d32;">
                    <p style="color: #4CAF50; font-size: 0.85rem; margin: 0;">
                        ✅ Archivo cargado: <b>{archivo_subido.name}</b> | Filas detectadas: <b>{len(df_excel)}</b>
                    </p>
                </div>
                """, unsafe_allow_html=True)

                # Columnas esperadas en la base de datos
                columnas_esperadas = ['eta', 'name', 'qty', 'room', 'email', 'check_in', 'check_out', 
                                      'res_number', 'phone', 'info', 'ird', 'hsk', 'rate', 'trans']

                # Verificar columnas
                columnas_faltantes = [col for col in columnas_esperadas if col not in df_excel.columns]
                columnas_extra = [col for col in df_excel.columns if col not in columnas_esperadas and col != 'id']

                if columnas_faltantes:
                    st.error(f"❌ Columnas faltantes en el archivo: {', '.join(columnas_faltantes)}")
                else:
                    if columnas_extra:
                        st.warning(f"⚠️ Columnas extra detectadas (serán ignoradas): {', '.join(columnas_extra)}")

                    # Mostrar preview de los datos
                    st.markdown("<p style='color:#888; font-size:0.8rem; margin-top:15px;'>👁️ Vista previa de los datos a importar:</p>", unsafe_allow_html=True)

                    # Crear una copia para mostrar (formatear fechas si es necesario)
                    df_preview = df_excel[columnas_esperadas].copy()

                    # Intentar convertir fechas si vienen como datetime
                    for col in ['check_in', 'check_out']:
                        if col in df_preview.columns:
                            try:
                                df_preview[col] = pd.to_datetime(df_preview[col], errors='coerce').dt.strftime('%b %d')
                            except:
                                pass

                    st.dataframe(df_preview, use_container_width=True, height=250)

                    # Botón para procesar la importación
                    col_proc, col_can = st.columns([1, 3])
                    with col_proc:
                        if st.button("📥 IMPORTAR A BASE DE DATOS", key="btn_procesar_excel", use_container_width=True):
                            try:
                                conn = sqlite3.connect(db_path)
                                cursor = conn.cursor()

                                registros_insertados = 0
                                registros_error = 0
                                errores_detalle = []

                                for idx, row in df_excel.iterrows():
                                    try:
                                        # Preparar valores
                                        eta = str(row.get('eta', '')).strip()
                                        name = str(row.get('name', '')).strip()
                                        qty = int(row.get('qty', 0)) if pd.notna(row.get('qty')) else 0
                                        room = str(row.get('room', '')).strip()
                                        email = str(row.get('email', '')).strip()

                                        # Formatear fechas
                                        check_in_raw = row.get('check_in', '')
                                        check_out_raw = row.get('check_out', '')

                                        if pd.notna(check_in_raw):
                                            try:
                                                check_in = pd.to_datetime(check_in_raw).strftime('%b %d')
                                            except:
                                                check_in = str(check_in_raw).strip()
                                        else:
                                            check_in = ''

                                        if pd.notna(check_out_raw):
                                            try:
                                                check_out = pd.to_datetime(check_out_raw).strftime('%b %d')
                                            except:
                                                check_out = str(check_out_raw).strip()
                                        else:
                                            check_out = ''

                                        res_number = str(row.get('res_number', '')).strip()
                                        phone = str(row.get('phone', '')).strip()
                                        info = str(row.get('info', '')).strip()
                                        ird = str(row.get('ird', '')).strip()
                                        hsk = str(row.get('hsk', '')).strip()
                                        rate = str(row.get('rate', '')).strip()
                                        trans = str(row.get('trans', '')).strip()

                                        cursor.execute("""
                                            INSERT INTO huespedes 
                                            (eta, name, qty, room, email, check_in, check_out, res_number, phone, info, ird, hsk, rate, trans)
                                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                                        """, (eta, name, qty, room, email, check_in, check_out, 
                                              res_number, phone, info, ird, hsk, rate, trans))
                                        registros_insertados += 1

                                    except Exception as e:
                                        registros_error += 1
                                        errores_detalle.append(f"Fila {idx + 2}: {str(e)}")

                                conn.commit()
                                conn.close()

                                # Mostrar resumen
                                st.markdown(f"""
                                <div style="background-color: #0d1f0d; border-radius: 8px; padding: 15px; margin: 15px 0; border: 1px solid #2e7d32;">
                                    <h4 style="color: #4CAF50; margin: 0 0 10px 0;">✅ Importación Completada</h4>
                                    <p style="color: #ccc; font-size: 0.9rem; margin: 0;">
                                        📥 Registros insertados: <b style="color: #4CAF50;">{registros_insertados}</b><br>
                                        ❌ Errores: <b style="color: {'#f44336' if registros_error > 0 else '#4CAF50'};">{registros_error}</b>
                                    </p>
                                </div>
                                """, unsafe_allow_html=True)

                                if registros_error > 0 and errores_detalle:
                                    with st.expander("🔍 Ver detalle de errores"):
                                        for error in errores_detalle[:10]:  # Mostrar primeros 10
                                            st.error(error)
                                        if len(errores_detalle) > 10:
                                            st.warning(f"... y {len(errores_detalle) - 10} errores más.")

                                if registros_insertados > 0:
                                    st.success("🎉 ¡Las reservas se han importado correctamente!")
                                    st.info("🔄 La página se recargará en 3 segundos...")
                                    import time
                                    time.sleep(3)
                                    st.query_params.clear()
                                    st.rerun()

                            except Exception as e:
                                st.error(f"❌ Error al importar: {str(e)}")

            except Exception as e:
                st.error(f"❌ Error al leer el archivo: {str(e)}")

# ============================================================
# 2B. FORMULARIO EXPORTAR A EXCEL - VERSIÓN CORREGIDA
# ============================================================
if mostrar_exportar:
    with st.container():
        st.subheader("📤 Exportar Reservaciones a Excel")

        col1, col2 = st.columns([1, 1])
        with col1:
            if st.button("↩️ REGRESAR", key="regresar_exportar"):
                st.query_params.clear()
                st.rerun()

        # Determinar qué datos se exportarán (filtrados o totales)
        df_export = cargar_reservaciones()

        # Aplicar los mismos filtros que la tabla principal
        filtro_activo = False

        if filtro_checkout:
            df_export = df_export[df_export['check_out'] == filtro_checkout]
            filtro_activo = True

        if fecha_filtro_activo and filtro_fecha_date:
            fecha_filtro = datetime.strptime(filtro_fecha_date, "%Y-%m-%d").strftime("%b %d")
            df_export = df_export[df_export['check_in'] == fecha_filtro]
            filtro_activo = True

        if busqueda and busqueda.strip():
            busqueda_lower = busqueda.strip().lower()
            mask = df_export.astype(str).apply(
                lambda row: row.str.lower().str.contains(busqueda_lower, na=False).any(), axis=1
            )
            df_export = df_export[mask]
            filtro_activo = True

        total_a_exportar = len(df_export)

        st.markdown(f"""
        <div style="background-color: #1a1a2e; border-radius: 8px; padding: 15px; margin: 10px 0; border: 1px solid #333;">
            <p style="color: #ccc; font-size: 0.85rem; margin: 0;">
                📋 <b>Resumen de exportación:</b><br>
                • Total de reservas a exportar: <b style="color: #00E5FF;">{total_a_exportar}</b><br>
                • Filtro activo: <b style="color: #00E5FF;">{"Sí - Solo datos filtrados" if filtro_activo else "No - Todos los datos"}</b><br>
                • El archivo se organizará por categorías: <b>CUMPLEAÑOS, VIP, HONEYMOON, ANNIVERSARY, BABYMOON, TEAM MEMBER, GENERAL</b>
            </p>
        </div>
        """, unsafe_allow_html=True)

        # Mostrar preview de cómo quedará organizado
        st.markdown("<p style='color:#888; font-size:0.8rem; margin-top:15px;'>👁️ Vista previa del contenido a exportar:</p>", unsafe_allow_html=True)
        st.dataframe(df_export, use_container_width=True, height=300)

        # Botón para descargar - CORREGIDO: sin use_container_width en download_button
        if total_a_exportar > 0:
            try:
                excel_buffer = exportar_excel_por_categorias(df_export)
                fecha_hoy = datetime.now().strftime("%Y%m%d_%H%M")

                st.download_button(
                    label="📥 DESCARGAR EXCEL",
                    data=excel_buffer,
                    file_name=f"Arrivals_{fecha_hoy}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_descargar_excel"
                )
            except Exception as e:
                st.error(f"❌ Error al generar el Excel: {str(e)}")
                st.exception(e)  # Muestra el traceback completo para debug
        else:
            st.warning("⚠️ No hay datos para exportar con los filtros actuales.")

# ============================================================
# 2C. FORMULARIO REPORTE DE OCUPACIÓN DIARIO
# ============================================================
if st.query_params.get("action") == "reporte":
    with st.container():
        st.subheader("📄 Reporte de Ocupación Diario")

        col1, col2 = st.columns([1, 1])
        with col1:
            if st.button("↩️ REGRESAR", key="regresar_reporte"):
                st.query_params.clear()
                st.rerun()

        # Obtener fecha actual
        hoy = datetime.now()
        hoy_str = hoy.strftime("%b %d")
        manana = hoy + timedelta(days=1)
        manana_str = manana.strftime("%b %d")

        # Cargar TODAS las reservaciones
        df_todas = cargar_reservaciones()

        # ============================================================
        # FUNCIONES AUXILIARES PARA COMPARAR FECHAS
        # ============================================================
        def fecha_es_menor_igual(fecha_str, referencia_str):
            """Compara fechas en formato %b %d considerando el año"""
            try:
                fecha = datetime.strptime(fecha_str, "%b %d").replace(year=hoy.year)
                referencia = datetime.strptime(referencia_str, "%b %d").replace(year=hoy.year)
                return fecha <= referencia
            except:
                return False

        def fecha_es_mayor(fecha_str, referencia_str):
            """Compara fechas en formato %b %d considerando el año"""
            try:
                fecha = datetime.strptime(fecha_str, "%b %d").replace(year=hoy.year)
                referencia = datetime.strptime(referencia_str, "%b %d").replace(year=hoy.year)
                return fecha > referencia
            except:
                return False

        def obtener_rooms(df):
            """Obtiene lista de rooms no vacíos de un DataFrame"""
            rooms = df['room'].dropna().astype(str)
            rooms = rooms[rooms.str.strip() != ''].str.strip().tolist()
            return rooms

        def formatear_rooms(rooms_list):
            """Formatea la lista de rooms para mostrar"""
            if not rooms_list:
                return "Ninguna"
            if len(rooms_list) <= 5:
                return ", ".join(rooms_list)
            return ", ".join(rooms_list[:5]) + f" (+{len(rooms_list) - 5} más)"

        # ============================================================
        # CALCULAR MÉTRICAS
        # ============================================================

        # 1. HABITACIONES EN CASA
        mask_en_casa = df_todas.apply(lambda row: 
            fecha_es_menor_igual(row['check_in'], hoy_str) and 
            fecha_es_mayor(row['check_out'], hoy_str), axis=1
        )
        df_en_casa = df_todas[mask_en_casa]
        total_en_casa = len(df_en_casa)
        rooms_en_casa = obtener_rooms(df_en_casa)

        # VIPs en casa
        mask_vip_en_casa = df_en_casa['info'].astype(str).str.upper().str.contains('VIP', na=False)
        vip_en_casa = df_en_casa[mask_vip_en_casa]
        total_vip_en_casa = len(vip_en_casa)
        rooms_vip_en_casa = obtener_rooms(vip_en_casa)

        # 2. HABITACIONES QUE SALEN HOY
        df_salen_hoy = df_todas[df_todas['check_out'] == hoy_str]
        total_salen_hoy = len(df_salen_hoy)
        rooms_salen_hoy = obtener_rooms(df_salen_hoy)

        # VIPs que salen hoy
        mask_vip_salen_hoy = df_salen_hoy['info'].astype(str).str.upper().str.contains('VIP', na=False)
        vip_salen_hoy = df_salen_hoy[mask_vip_salen_hoy]
        total_vip_salen_hoy = len(vip_salen_hoy)
        rooms_vip_salen_hoy = obtener_rooms(vip_salen_hoy)

        # 3. HABITACIONES QUE SALEN MAÑANA
        df_salen_manana = df_todas[df_todas['check_out'] == manana_str]
        total_salen_manana = len(df_salen_manana)
        rooms_salen_manana = obtener_rooms(df_salen_manana)

        # VIPs que salen mañana
        mask_vip_salen_manana = df_salen_manana['info'].astype(str).str.upper().str.contains('VIP', na=False)
        vip_salen_manana = df_salen_manana[mask_vip_salen_manana]
        total_vip_salen_manana = len(vip_salen_manana)
        rooms_vip_salen_manana = obtener_rooms(vip_salen_manana)

        # 4. HABITACIONES QUE LLEGAN HOY
        df_llegan_hoy = df_todas[df_todas['check_in'] == hoy_str]
        total_llegan_hoy = len(df_llegan_hoy)
        rooms_llegan_hoy = obtener_rooms(df_llegan_hoy)

        # VIPs que llegan hoy
        mask_vip_llegan_hoy = df_llegan_hoy['info'].astype(str).str.upper().str.contains('VIP', na=False)
        vip_llegan_hoy = df_llegan_hoy[mask_vip_llegan_hoy]
        total_vip_llegan_hoy = len(vip_llegan_hoy)
        rooms_vip_llegan_hoy = obtener_rooms(vip_llegan_hoy)

        # 5. HABITACIONES QUE LLEGAN MAÑANA
        df_llegan_manana = df_todas[df_todas['check_in'] == manana_str]
        total_llegan_manana = len(df_llegan_manana)
        rooms_llegan_manana = obtener_rooms(df_llegan_manana)

        # VIPs que llegan mañana
        mask_vip_llegan_manana = df_llegan_manana['info'].astype(str).str.upper().str.contains('VIP', na=False)
        vip_llegan_manana = df_llegan_manana[mask_vip_llegan_manana]
        total_vip_llegan_manana = len(vip_llegan_manana)
        rooms_vip_llegan_manana = obtener_rooms(vip_llegan_manana)

        # 6. TOTAL DE RESERVAS EN BASE DE DATOS
        total_reservas = len(df_todas)

        # ============================================================
        # MOSTRAR RESUMEN EN PANTALLA
        # ============================================================
        st.markdown(f"""
        <div style="background-color: #1a1a2e; border-radius: 8px; padding: 15px; margin: 10px 0; border: 1px solid #333;">
            <p style="color: #ccc; font-size: 0.85rem; margin: 0;">
                📅 <b>Fecha del Reporte:</b> <span style="color: #00E5FF;">{hoy.strftime("%A, %B %d, %Y")}</span><br>
                🕐 <b>Hora de generación:</b> <span style="color: #00E5FF;">{hoy.strftime("%H:%M")}</span>
            </p>
        </div>
        """, unsafe_allow_html=True)

        # Tarjetas de métricas con ROOMS
        col_m1, col_m2, col_m3 = st.columns(3)
        with col_m1:
            st.markdown(f"""
            <div style="background-color: #0d1f0d; border-radius: 8px; padding: 12px; text-align: center; border: 1px solid #2e7d32;">
                <div style="color: #4CAF50; font-size: 0.75rem; font-weight: bold;">🏨 EN CASA</div>
                <div style="color: #fff; font-size: 1.8rem; font-weight: bold;">{total_en_casa}</div>
                <div style="color: #888; font-size: 0.65rem; margin-top: 4px;">🚪 {formatear_rooms(rooms_en_casa)}</div>
            </div>
            """, unsafe_allow_html=True)
        with col_m2:
            st.markdown(f"""
            <div style="background-color: #0d1f0d; border-radius: 8px; padding: 12px; text-align: center; border: 1px solid #2e7d32;">
                <div style="color: #4CAF50; font-size: 0.75rem; font-weight: bold;">👑 VIPs EN CASA</div>
                <div style="color: #fff; font-size: 1.8rem; font-weight: bold;">{total_vip_en_casa}</div>
                <div style="color: #888; font-size: 0.65rem; margin-top: 4px;">🚪 {formatear_rooms(rooms_vip_en_casa)}</div>
            </div>
            """, unsafe_allow_html=True)
        with col_m3:
            st.markdown(f"""
            <div style="background-color: #0d1f0d; border-radius: 8px; padding: 12px; text-align: center; border: 1px solid #2e7d32;">
                <div style="color: #4CAF50; font-size: 0.75rem; font-weight: bold;">📊 TOTAL DB</div>
                <div style="color: #fff; font-size: 1.8rem; font-weight: bold;">{total_reservas}</div>
                <div style="color: #888; font-size: 0.65rem; margin-top: 4px;">Reservas registradas</div>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)

        col_m4, col_m5, col_m6 = st.columns(3)
        with col_m4:
            st.markdown(f"""
            <div style="background-color: #1a0d0d; border-radius: 8px; padding: 12px; text-align: center; border: 1px solid #7d2e2e;">
                <div style="color: #f44336; font-size: 0.75rem; font-weight: bold;">🚪 SALEN HOY ({hoy_str})</div>
                <div style="color: #fff; font-size: 1.8rem; font-weight: bold;">{total_salen_hoy}</div>
                <div style="color: #888; font-size: 0.65rem; margin-top: 4px;">👑 VIPs: {total_vip_salen_hoy} | 🚪 {formatear_rooms(rooms_salen_hoy)}</div>
            </div>
            """, unsafe_allow_html=True)
        with col_m5:
            st.markdown(f"""
            <div style="background-color: #1a0d0d; border-radius: 8px; padding: 12px; text-align: center; border: 1px solid #7d2e2e;">
                <div style="color: #f44336; font-size: 0.75rem; font-weight: bold;">🚪 SALEN MAÑANA ({manana_str})</div>
                <div style="color: #fff; font-size: 1.8rem; font-weight: bold;">{total_salen_manana}</div>
                <div style="color: #888; font-size: 0.65rem; margin-top: 4px;">👑 VIPs: {total_vip_salen_manana} | 🚪 {formatear_rooms(rooms_salen_manana)}</div>
            </div>
            """, unsafe_allow_html=True)
        with col_m6:
            st.markdown(f"""
            <div style="background-color: #0d0d1a; border-radius: 8px; padding: 12px; text-align: center; border: 1px solid #2e2e7d;">
                <div style="color: #2196F3; font-size: 0.75rem; font-weight: bold;">📥 LLEGAN HOY ({hoy_str})</div>
                <div style="color: #fff; font-size: 1.8rem; font-weight: bold;">{total_llegan_hoy}</div>
                <div style="color: #888; font-size: 0.65rem; margin-top: 4px;">👑 VIPs: {total_vip_llegan_hoy} | 🚪 {formatear_rooms(rooms_llegan_hoy)}</div>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)

        col_m7, col_m8, col_m9 = st.columns(3)
        with col_m7:
            st.markdown(f"""
            <div style="background-color: #0d0d1a; border-radius: 8px; padding: 12px; text-align: center; border: 1px solid #2e2e7d;">
                <div style="color: #2196F3; font-size: 0.75rem; font-weight: bold;">📥 LLEGAN MAÑANA ({manana_str})</div>
                <div style="color: #fff; font-size: 1.8rem; font-weight: bold;">{total_llegan_manana}</div>
                <div style="color: #888; font-size: 0.65rem; margin-top: 4px;">👑 VIPs: {total_vip_llegan_manana} | 🚪 {formatear_rooms(rooms_llegan_manana)}</div>
            </div>
            """, unsafe_allow_html=True)
        with col_m8:
            st.markdown(f"""
            <div style="background-color: #1a1a0d; border-radius: 8px; padding: 12px; text-align: center; border: 1px solid #7d7d2e;">
                <div style="color: #FFC107; font-size: 0.75rem; font-weight: bold;">🌙 ESTANCIA PROMEDIO</div>
                <div style="color: #fff; font-size: 1.8rem; font-weight: bold;">{round(total_en_casa / max(total_salen_hoy, 1), 1)}</div>
                <div style="color: #888; font-size: 0.65rem; margin-top: 4px;">Ratio en casa / salen hoy</div>
            </div>
            """, unsafe_allow_html=True)
        with col_m9:
            ocupacion_pct = round((total_en_casa / max(total_reservas, 1)) * 100, 1)
            st.markdown(f"""
            <div style="background-color: #1a0d1a; border-radius: 8px; padding: 12px; text-align: center; border: 1px solid #7d2e7d;">
                <div style="color: #9C27B0; font-size: 0.75rem; font-weight: bold;">📈 OCUPACIÓN</div>
                <div style="color: #fff; font-size: 1.8rem; font-weight: bold;">{ocupacion_pct}%</div>
                <div style="color: #888; font-size: 0.65rem; margin-top: 4px;">En casa / Total DB</div>
            </div>
            """, unsafe_allow_html=True)

        # ============================================================
        # TABLA DETALLADA DE ROOMS POR CATEGORÍA
        # ============================================================
        st.markdown("<div style='height: 15px;'></div>", unsafe_allow_html=True)
        st.markdown("""
        <div style="background-color: #1a1a2e; border-radius: 8px; padding: 10px; margin-bottom: 10px; border: 1px solid #00E5FF;">
            <div style="color: #00E5FF; font-size: 0.85rem; font-weight: bold; text-align: center;">
                📋 DETALLE DE HABITACIONES POR CATEGORÍA
            </div>
        </div>
        """, unsafe_allow_html=True)

        # Crear DataFrame resumen de rooms
        resumen_rooms_data = []

        if rooms_en_casa:
            resumen_rooms_data.append({
                "Categoría": "🏨 En Casa",
                "Total": total_en_casa,
                "Habitaciones": ", ".join(rooms_en_casa)
            })
        if rooms_vip_en_casa:
            resumen_rooms_data.append({
                "Categoría": "👑 VIPs En Casa",
                "Total": total_vip_en_casa,
                "Habitaciones": ", ".join(rooms_vip_en_casa)
            })
        if rooms_salen_hoy:
            resumen_rooms_data.append({
                "Categoría": f"🚪 Salen Hoy ({hoy_str})",
                "Total": total_salen_hoy,
                "Habitaciones": ", ".join(rooms_salen_hoy)
            })
        if rooms_vip_salen_hoy:
            resumen_rooms_data.append({
                "Categoría": f"👑 VIPs Salen Hoy ({hoy_str})",
                "Total": total_vip_salen_hoy,
                "Habitaciones": ", ".join(rooms_vip_salen_hoy)
            })
        if rooms_salen_manana:
            resumen_rooms_data.append({
                "Categoría": f"🚪 Salen Mañana ({manana_str})",
                "Total": total_salen_manana,
                "Habitaciones": ", ".join(rooms_salen_manana)
            })
        if rooms_vip_salen_manana:
            resumen_rooms_data.append({
                "Categoría": f"👑 VIPs Salen Mañana ({manana_str})",
                "Total": total_vip_salen_manana,
                "Habitaciones": ", ".join(rooms_vip_salen_manana)
            })
        if rooms_llegan_hoy:
            resumen_rooms_data.append({
                "Categoría": f"📥 Llegan Hoy ({hoy_str})",
                "Total": total_llegan_hoy,
                "Habitaciones": ", ".join(rooms_llegan_hoy)
            })
        if rooms_vip_llegan_hoy:
            resumen_rooms_data.append({
                "Categoría": f"👑 VIPs Llegan Hoy ({hoy_str})",
                "Total": total_vip_llegan_hoy,
                "Habitaciones": ", ".join(rooms_vip_llegan_hoy)
            })
        if rooms_llegan_manana:
            resumen_rooms_data.append({
                "Categoría": f"📥 Llegan Mañana ({manana_str})",
                "Total": total_llegan_manana,
                "Habitaciones": ", ".join(rooms_llegan_manana)
            })
        if rooms_vip_llegan_manana:
            resumen_rooms_data.append({
                "Categoría": f"👑 VIPs Llegan Mañana ({manana_str})",
                "Total": total_vip_llegan_manana,
                "Habitaciones": ", ".join(rooms_vip_llegan_manana)
            })

        if resumen_rooms_data:
            df_resumen_rooms = pd.DataFrame(resumen_rooms_data)
            st.dataframe(df_resumen_rooms, use_container_width=True, hide_index=True)
        else:
            st.info("📭 No hay habitaciones para mostrar en el resumen.")

        # ============================================================
        # GENERAR EXCEL DEL REPORTE
        # ============================================================
        def generar_reporte_excel():
            from openpyxl import Workbook
            from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

            wb = Workbook()

            # ========================================
            # HOJA 1: RESUMEN EJECUTIVO
            # ========================================
            ws_resumen = wb.active
            ws_resumen.title = "Resumen Ejecutivo"

            # Estilos
            fill_titulo = PatternFill(start_color="00B0F0", end_color="00B0F0", fill_type="solid")
            fill_header = PatternFill(start_color="404040", end_color="404040", fill_type="solid")
            fill_verde = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
            fill_rojo = PatternFill(start_color="C62828", end_color="C62828", fill_type="solid")
            fill_azul = PatternFill(start_color="1565C0", end_color="1565C0", fill_type="solid")
            fill_morado = PatternFill(start_color="6A1B9A", end_color="6A1B9A", fill_type="solid")
            fill_amarillo = PatternFill(start_color="F9A825", end_color="F9A825", fill_type="solid")
            fill_gris = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

            font_titulo = Font(name='Calibri', size=16, bold=True, color="FFFFFF")
            font_header = Font(name='Calibri', size=11, bold=True, color="FFFFFF")
            font_metrica = Font(name='Calibri', size=14, bold=True, color="000000")
            font_label = Font(name='Calibri', size=10, color="000000")
            font_fecha = Font(name='Calibri', size=12, bold=True, color="00B0F0")
            font_rooms = Font(name='Calibri', size=9, color="666666")

            align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
            align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)

            thin_border = Border(
                left=Side(style='thin', color='D9D9D9'),
                right=Side(style='thin', color='D9D9D9'),
                top=Side(style='thin', color='D9D9D9'),
                bottom=Side(style='thin', color='D9D9D9')
            )

            # Título del reporte
            ws_resumen.merge_cells('A1:E1')
            cell = ws_resumen['A1']
            cell.value = f"REPORTE DE OCUPACIÓN - {hoy.strftime('%A, %B %d, %Y').upper()}"
            cell.fill = fill_titulo
            cell.font = font_titulo
            cell.alignment = align_center

            # Fecha y hora
            ws_resumen.merge_cells('A2:E2')
            cell = ws_resumen['A2']
            cell.value = f"Generado el {hoy.strftime('%d/%m/%Y')} a las {hoy.strftime('%H:%M')}"
            cell.font = font_fecha
            cell.alignment = align_center

            ws_resumen.row_dimensions[1].height = 30
            ws_resumen.row_dimensions[2].height = 20

            # Headers de métricas
            headers_metricas = ['MÉTRICA', 'TOTAL', 'VIPs', 'HABITACIONES', 'NOTAS']
            for col_idx, header in enumerate(headers_metricas, 1):
                cell = ws_resumen.cell(row=4, column=col_idx, value=header)
                cell.fill = fill_header
                cell.font = font_header
                cell.alignment = align_center
                cell.border = thin_border

            # Métricas con rooms
            metricas_data = [
                ("🏨 EN CASA", total_en_casa, total_vip_en_casa, ", ".join(rooms_en_casa) if rooms_en_casa else "N/A", ""),
                ("🚪 SALEN HOY", total_salen_hoy, total_vip_salen_hoy, ", ".join(rooms_salen_hoy) if rooms_salen_hoy else "N/A", hoy_str),
                ("🚪 SALEN MAÑANA", total_salen_manana, total_vip_salen_manana, ", ".join(rooms_salen_manana) if rooms_salen_manana else "N/A", manana_str),
                ("📥 LLEGAN HOY", total_llegan_hoy, total_vip_llegan_hoy, ", ".join(rooms_llegan_hoy) if rooms_llegan_hoy else "N/A", hoy_str),
                ("📥 LLEGAN MAÑANA", total_llegan_manana, total_vip_llegan_manana, ", ".join(rooms_llegan_manana) if rooms_llegan_manana else "N/A", manana_str),
                ("📊 TOTAL DB", total_reservas, "-", "-", "Todas las reservas"),
                ("📈 % OCUPACIÓN", f"{ocupacion_pct}%", "-", "-", f"{total_en_casa} de {total_reservas}"),
            ]

            row_start = 5
            for i, (label, total, vips, rooms, notas) in enumerate(metricas_data):
                row = row_start + i

                cell_label = ws_resumen.cell(row=row, column=1, value=label)
                cell_label.fill = fill_gris
                cell_label.font = font_label
                cell_label.alignment = align_left
                cell_label.border = thin_border

                cell_total = ws_resumen.cell(row=row, column=2, value=total)
                cell_total.fill = fill_gris
                cell_total.font = font_metrica
                cell_total.alignment = align_center
                cell_total.border = thin_border

                cell_vips = ws_resumen.cell(row=row, column=3, value=vips)
                cell_vips.fill = fill_gris
                cell_vips.font = font_label
                cell_vips.alignment = align_center
                cell_vips.border = thin_border

                cell_rooms = ws_resumen.cell(row=row, column=4, value=rooms)
                cell_rooms.fill = fill_gris
                cell_rooms.font = font_rooms
                cell_rooms.alignment = align_left
                cell_rooms.border = thin_border

                cell_notas = ws_resumen.cell(row=row, column=5, value=notas)
                cell_notas.fill = fill_gris
                cell_notas.font = font_rooms
                cell_notas.alignment = align_left
                cell_notas.border = thin_border

                ws_resumen.row_dimensions[row].height = 25

            # Ajustar anchos
            ws_resumen.column_dimensions['A'].width = 28
            ws_resumen.column_dimensions['B'].width = 12
            ws_resumen.column_dimensions['C'].width = 10
            ws_resumen.column_dimensions['D'].width = 45
            ws_resumen.column_dimensions['E'].width = 25

            # ========================================
            # HOJA 2: DETALLE EN CASA
            # ========================================
            ws_casa = wb.create_sheet("En Casa")

            ws_casa.merge_cells('A1:O1')
            cell = ws_casa['A1']
            cell.value = f"HABITACIONES EN CASA - {hoy_str}"
            cell.fill = fill_verde
            cell.font = font_titulo
            cell.alignment = align_center
            ws_casa.row_dimensions[1].height = 25

            headers = ['ID', 'ETA', 'NAME', 'QTY', 'ROOM', 'EMAIL', 'CHECK IN', 'CHECK OUT', 
                       'RESERVATION', 'PHONE', 'INFO', 'IRD', 'HSK', 'RATE', 'TRANSPORTATION']

            for col_idx, header in enumerate(headers, 1):
                cell = ws_casa.cell(row=2, column=col_idx, value=header)
                cell.fill = fill_header
                cell.font = font_header
                cell.alignment = align_center
                cell.border = thin_border

            for row_idx, (_, row_data) in enumerate(df_en_casa.iterrows(), 3):
                valores = [
                    row_data.get('id', ''), row_data.get('eta', ''), row_data.get('name', ''),
                    row_data.get('qty', ''), row_data.get('room', ''), row_data.get('email', ''),
                    row_data.get('check_in', ''), row_data.get('check_out', ''),
                    row_data.get('res_number', ''), row_data.get('phone', ''),
                    row_data.get('info', ''), row_data.get('ird', ''),
                    row_data.get('hsk', ''), row_data.get('rate', ''), row_data.get('trans', '')
                ]
                for col_idx, valor in enumerate(valores, 1):
                    cell = ws_casa.cell(row=row_idx, column=col_idx, value=valor if pd.notna(valor) else '')
                    cell.fill = fill_gris
                    cell.font = font_label
                    cell.alignment = align_left
                    cell.border = thin_border

            anchos = {'A': 6, 'B': 10, 'C': 22, 'D': 6, 'E': 8, 'F': 25, 'G': 10, 'H': 10, 
                      'I': 14, 'J': 18, 'K': 22, 'L': 18, 'M': 18, 'N': 8, 'O': 18}
            for col_letter, ancho in anchos.items():
                ws_casa.column_dimensions[col_letter].width = ancho

            # ========================================
            # HOJA 3: SALEN HOY
            # ========================================
            ws_salen_hoy_sheet = wb.create_sheet(f"Salen Hoy {hoy_str}")

            ws_salen_hoy_sheet.merge_cells('A1:O1')
            cell = ws_salen_hoy_sheet['A1']
            cell.value = f"HABITACIONES QUE SALEN HOY - {hoy_str}"
            cell.fill = fill_rojo
            cell.font = font_titulo
            cell.alignment = align_center
            ws_salen_hoy_sheet.row_dimensions[1].height = 25

            for col_idx, header in enumerate(headers, 1):
                cell = ws_salen_hoy_sheet.cell(row=2, column=col_idx, value=header)
                cell.fill = fill_header
                cell.font = font_header
                cell.alignment = align_center
                cell.border = thin_border

            for row_idx, (_, row_data) in enumerate(df_salen_hoy.iterrows(), 3):
                valores = [
                    row_data.get('id', ''), row_data.get('eta', ''), row_data.get('name', ''),
                    row_data.get('qty', ''), row_data.get('room', ''), row_data.get('email', ''),
                    row_data.get('check_in', ''), row_data.get('check_out', ''),
                    row_data.get('res_number', ''), row_data.get('phone', ''),
                    row_data.get('info', ''), row_data.get('ird', ''),
                    row_data.get('hsk', ''), row_data.get('rate', ''), row_data.get('trans', '')
                ]
                for col_idx, valor in enumerate(valores, 1):
                    cell = ws_salen_hoy_sheet.cell(row=row_idx, column=col_idx, value=valor if pd.notna(valor) else '')
                    cell.fill = fill_gris
                    cell.font = font_label
                    cell.alignment = align_left
                    cell.border = thin_border

            for col_letter, ancho in anchos.items():
                ws_salen_hoy_sheet.column_dimensions[col_letter].width = ancho

            # ========================================
            # HOJA 4: SALEN MAÑANA
            # ========================================
            ws_salen_manana_sheet = wb.create_sheet(f"Salen Mañana {manana_str}")

            ws_salen_manana_sheet.merge_cells('A1:O1')
            cell = ws_salen_manana_sheet['A1']
            cell.value = f"HABITACIONES QUE SALEN MAÑANA - {manana_str}"
            cell.fill = fill_rojo
            cell.font = font_titulo
            cell.alignment = align_center
            ws_salen_manana_sheet.row_dimensions[1].height = 25

            for col_idx, header in enumerate(headers, 1):
                cell = ws_salen_manana_sheet.cell(row=2, column=col_idx, value=header)
                cell.fill = fill_header
                cell.font = font_header
                cell.alignment = align_center
                cell.border = thin_border

            for row_idx, (_, row_data) in enumerate(df_salen_manana.iterrows(), 3):
                valores = [
                    row_data.get('id', ''), row_data.get('eta', ''), row_data.get('name', ''),
                    row_data.get('qty', ''), row_data.get('room', ''), row_data.get('email', ''),
                    row_data.get('check_in', ''), row_data.get('check_out', ''),
                    row_data.get('res_number', ''), row_data.get('phone', ''),
                    row_data.get('info', ''), row_data.get('ird', ''),
                    row_data.get('hsk', ''), row_data.get('rate', ''), row_data.get('trans', '')
                ]
                for col_idx, valor in enumerate(valores, 1):
                    cell = ws_salen_manana_sheet.cell(row=row_idx, column=col_idx, value=valor if pd.notna(valor) else '')
                    cell.fill = fill_gris
                    cell.font = font_label
                    cell.alignment = align_left
                    cell.border = thin_border

            for col_letter, ancho in anchos.items():
                ws_salen_manana_sheet.column_dimensions[col_letter].width = ancho

            # ========================================
            # HOJA 5: LLEGAN HOY
            # ========================================
            ws_llegan_hoy_sheet = wb.create_sheet(f"Llegan Hoy {hoy_str}")

            ws_llegan_hoy_sheet.merge_cells('A1:O1')
            cell = ws_llegan_hoy_sheet['A1']
            cell.value = f"HABITACIONES QUE LLEGAN HOY - {hoy_str}"
            cell.fill = fill_azul
            cell.font = font_titulo
            cell.alignment = align_center
            ws_llegan_hoy_sheet.row_dimensions[1].height = 25

            for col_idx, header in enumerate(headers, 1):
                cell = ws_llegan_hoy_sheet.cell(row=2, column=col_idx, value=header)
                cell.fill = fill_header
                cell.font = font_header
                cell.alignment = align_center
                cell.border = thin_border

            for row_idx, (_, row_data) in enumerate(df_llegan_hoy.iterrows(), 3):
                valores = [
                    row_data.get('id', ''), row_data.get('eta', ''), row_data.get('name', ''),
                    row_data.get('qty', ''), row_data.get('room', ''), row_data.get('email', ''),
                    row_data.get('check_in', ''), row_data.get('check_out', ''),
                    row_data.get('res_number', ''), row_data.get('phone', ''),
                    row_data.get('info', ''), row_data.get('ird', ''),
                    row_data.get('hsk', ''), row_data.get('rate', ''), row_data.get('trans', '')
                ]
                for col_idx, valor in enumerate(valores, 1):
                    cell = ws_llegan_hoy_sheet.cell(row=row_idx, column=col_idx, value=valor if pd.notna(valor) else '')
                    cell.fill = fill_gris
                    cell.font = font_label
                    cell.alignment = align_left
                    cell.border = thin_border

            for col_letter, ancho in anchos.items():
                ws_llegan_hoy_sheet.column_dimensions[col_letter].width = ancho

            # ========================================
            # HOJA 6: LLEGAN MAÑANA
            # ========================================
            ws_llegan_manana_sheet = wb.create_sheet(f"Llegan Mañana {manana_str}")

            ws_llegan_manana_sheet.merge_cells('A1:O1')
            cell = ws_llegan_manana_sheet['A1']
            cell.value = f"HABITACIONES QUE LLEGAN MAÑANA - {manana_str}"
            cell.fill = fill_azul
            cell.font = font_titulo
            cell.alignment = align_center
            ws_llegan_manana_sheet.row_dimensions[1].height = 25

            for col_idx, header in enumerate(headers, 1):
                cell = ws_llegan_manana_sheet.cell(row=2, column=col_idx, value=header)
                cell.fill = fill_header
                cell.font = font_header
                cell.alignment = align_center
                cell.border = thin_border

            for row_idx, (_, row_data) in enumerate(df_llegan_manana.iterrows(), 3):
                valores = [
                    row_data.get('id', ''), row_data.get('eta', ''), row_data.get('name', ''),
                    row_data.get('qty', ''), row_data.get('room', ''), row_data.get('email', ''),
                    row_data.get('check_in', ''), row_data.get('check_out', ''),
                    row_data.get('res_number', ''), row_data.get('phone', ''),
                    row_data.get('info', ''), row_data.get('ird', ''),
                    row_data.get('hsk', ''), row_data.get('rate', ''), row_data.get('trans', '')
                ]
                for col_idx, valor in enumerate(valores, 1):
                    cell = ws_llegan_manana_sheet.cell(row=row_idx, column=col_idx, value=valor if pd.notna(valor) else '')
                    cell.fill = fill_gris
                    cell.font = font_label
                    cell.alignment = align_left
                    cell.border = thin_border

            for col_letter, ancho in anchos.items():
                ws_llegan_manana_sheet.column_dimensions[col_letter].width = ancho

            # ========================================
            # HOJA 7: VIPs EN CASA
            # ========================================
            if total_vip_en_casa > 0:
                ws_vip_casa = wb.create_sheet("VIPs En Casa")

                ws_vip_casa.merge_cells('A1:O1')
                cell = ws_vip_casa['A1']
                cell.value = f"VIPs EN CASA - {hoy_str}"
                cell.fill = fill_verde
                cell.font = font_titulo
                cell.alignment = align_center
                ws_vip_casa.row_dimensions[1].height = 25

                for col_idx, header in enumerate(headers, 1):
                    cell = ws_vip_casa.cell(row=2, column=col_idx, value=header)
                    cell.fill = fill_header
                    cell.font = font_header
                    cell.alignment = align_center
                    cell.border = thin_border

                for row_idx, (_, row_data) in enumerate(vip_en_casa.iterrows(), 3):
                    valores = [
                        row_data.get('id', ''), row_data.get('eta', ''), row_data.get('name', ''),
                        row_data.get('qty', ''), row_data.get('room', ''), row_data.get('email', ''),
                        row_data.get('check_in', ''), row_data.get('check_out', ''),
                        row_data.get('res_number', ''), row_data.get('phone', ''),
                        row_data.get('info', ''), row_data.get('ird', ''),
                        row_data.get('hsk', ''), row_data.get('rate', ''), row_data.get('trans', '')
                    ]
                    for col_idx, valor in enumerate(valores, 1):
                        cell = ws_vip_casa.cell(row=row_idx, column=col_idx, value=valor if pd.notna(valor) else '')
                        cell.fill = fill_gris
                        cell.font = font_label
                        cell.alignment = align_left
                        cell.border = thin_border

                for col_letter, ancho in anchos.items():
                    ws_vip_casa.column_dimensions[col_letter].width = ancho

            # ========================================
            # HOJA 8: RESUMEN DE ROOMS
            # ========================================
            ws_rooms = wb.create_sheet("Resumen Rooms")

            ws_rooms.merge_cells('A1:D1')
            cell = ws_rooms['A1']
            cell.value = f"RESUMEN DE HABITACIONES - {hoy_str}"
            cell.fill = fill_morado
            cell.font = font_titulo
            cell.alignment = align_center
            ws_rooms.row_dimensions[1].height = 25

            headers_rooms = ['CATEGORÍA', 'TOTAL', 'HABITACIONES', 'FECHA']
            for col_idx, header in enumerate(headers_rooms, 1):
                cell = ws_rooms.cell(row=2, column=col_idx, value=header)
                cell.fill = fill_header
                cell.font = font_header
                cell.alignment = align_center
                cell.border = thin_border

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

            ws_rooms.column_dimensions['A'].width = 25
            ws_rooms.column_dimensions['B'].width = 10
            ws_rooms.column_dimensions['C'].width = 50
            ws_rooms.column_dimensions['D'].width = 12

            # Guardar en buffer
            buffer = BytesIO()
            wb.save(buffer)
            buffer.seek(0)
            return buffer

        # Botón para descargar el reporte
        st.markdown("<div style='height: 15px;'></div>", unsafe_allow_html=True)

        try:
            excel_buffer = generar_reporte_excel()
            fecha_hoy = hoy.strftime("%Y%m%d_%H%M")

            st.download_button(
                label="📥 DESCARGAR REPORTE EXCEL",
                data=excel_buffer,
                file_name=f"Reporte_Ocupacion_{fecha_hoy}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_descargar_reporte"
            )
        except Exception as e:
            st.error(f"❌ Error al generar el reporte: {str(e)}")
            st.exception(e)

# ============================================================
# 2D. FORMULARIO CARTA DE DESPEDIDA
# ============================================================
if st.query_params.get("action") == "carta":
    with st.container():
        st.subheader("✉️ Carta de Despedida")

        col1, col2 = st.columns([1, 1])
        with col1:
            if st.button("↩️ REGRESAR", key="regresar_carta"):
                st.query_params.clear()
                st.rerun()

        # Verificar que hay una fila seleccionada
        fila_guardada = st.session_state.get("fila_seleccionada")

        if not fila_guardada:
            st.error("❌ Por favor, selecciona una fila en la tabla primero.")
            st.info("💡 Ve a la tabla principal, haz clic en la reserva del huésped, y luego presiona ✉️ CARTA.")

            if st.button("↩️ REGRESAR A LA TABLA", key="regresar_carta_error"):
                st.query_params.clear()
                st.rerun()
        else:
            # Obtener datos del huésped seleccionado
            nombre_huesped = str(fila_guardada.get("name", "")).strip()
            room_huesped = str(fila_guardada.get("room", "")).strip()
            check_out_huesped = str(fila_guardada.get("check_out", "")).strip()

            if not nombre_huesped:
                st.error("❌ La reserva seleccionada no tiene nombre de huésped.")
            else:
                # Mostrar info del huésped seleccionado
                st.markdown(f"""
                <div style="background-color: #1a1a2e; border-radius: 8px; padding: 15px; margin: 10px 0; border: 1px solid #00E5FF;">
                    <p style="color: #ccc; font-size: 0.9rem; margin: 0;">
                        👤 <b>Huésped:</b> <span style="color: #00E5FF; font-size: 1.1rem;">{nombre_huesped}</span><br>
                        🚪 <b>Habitación:</b> <span style="color: #00E5FF;">{room_huesped if room_huesped else 'N/A'}</span><br>
                        📅 <b>Check-out:</b> <span style="color: #00E5FF;">{check_out_huesped if check_out_huesped else 'N/A'}</span>
                    </p>
                </div>
                """, unsafe_allow_html=True)

                # Buscar la plantilla
                plantilla_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "plantilla_despedida.docx")

                # Verificar si existe la plantilla
                if not os.path.exists(plantilla_path):
                    st.warning("⚠️ No se encontró la plantilla `plantilla_despedida.docx`")
                    st.info(f"📁 La plantilla debe estar en: `{os.path.dirname(os.path.abspath(__file__))}`")

                    # Mostrar instrucciones para crear la plantilla
                    st.markdown("""
                    <div style="background-color: #1a1a2e; border-radius: 8px; padding: 15px; margin: 10px 0; border: 1px solid #333;">
                        <p style="color: #ccc; font-size: 0.85rem; margin: 0;">
                            <b>📝 Instrucciones para crear la plantilla:</b><br><br>
                            1. Crea un archivo Word llamado <code>plantilla_despedida.docx</code><br>
                            2. Escribe el texto de tu carta de despedida<br>
                            3. Donde quieras que aparezca el nombre del huésped, escribe: <code>{{NOMBRE}}</code><br>
                            4. Guarda el archivo en la misma carpeta donde está este script<br><br>
                            <b>Ejemplo de texto en la plantilla:</b><br>
                            <i>"Dear {{NOMBRE}},<br><br>
                            It was an absolute pleasure having you with us..."</i>
                        </p>
                    </div>
                    """, unsafe_allow_html=True)
                else:
                    st.success(f"✅ Plantilla encontrada: `plantilla_despedida.docx`")

                    try:
                        from docx import Document

                        # Leer la plantilla
                        doc = Document(plantilla_path)

                        # ============================================================
                        # FUNCIÓN PARA REEMPLAZAR PLACEHOLDERS EN UN PÁRRAFO
                        # ============================================================
                        def reemplazar_en_parrafo(paragraph, placeholder, replacement):
                            """
                            Reemplaza un placeholder en un párrafo completo, incluso si está
                            dividido en múltiples runs con diferentes formatos.
                            """
                            # Obtener el texto completo del párrafo
                            full_text = paragraph.text
                            
                            if placeholder in full_text:
                                # Reemplazar en el texto completo
                                new_text = full_text.replace(placeholder, replacement)
                                
                                # Limpiar todos los runs existentes
                                for run in paragraph.runs:
                                    run.text = ""
                                
                                # Si el párrafo tiene runs, usar el primero para poner todo el texto
                                if paragraph.runs:
                                    paragraph.runs[0].text = new_text
                                else:
                                    # Si no hay runs, agregar uno nuevo
                                    paragraph.add_run(new_text)
                                
                                return True
                            return False

                        # Contar cuántos placeholders se reemplazaron
                        placeholders_encontrados = 0

                        # Reemplazar en párrafos
                        for paragraph in doc.paragraphs:
                            if reemplazar_en_parrafo(paragraph, "{{NOMBRE}}", nombre_huesped):
                                placeholders_encontrados += 1

                        # Reemplazar en tablas (si las hay)
                        for table in doc.tables:
                            for row in table.rows:
                                for cell in row.cells:
                                    for paragraph in cell.paragraphs:
                                        if reemplazar_en_parrafo(paragraph, "{{NOMBRE}}", nombre_huesped):
                                            placeholders_encontrados += 1

                        # Guardar en buffer
                        buffer = BytesIO()
                        doc.save(buffer)
                        buffer.seek(0)

                        # Mostrar preview del contenido
                        st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
                        st.markdown("""
                        <div style="background-color: #0d1f0d; border-radius: 8px; padding: 10px; margin-bottom: 10px; border: 1px solid #2e7d32;">
                            <div style="color: #4CAF50; font-size: 0.8rem; font-weight: bold; text-align: center;">
                                ✅ CARTA GENERADA CORRECTAMENTE
                            </div>
                        </div>
                        """, unsafe_allow_html=True)

                        if placeholders_encontrados > 0:
                            st.info(f"📌 Se reemplazaron {placeholders_encontrados} placeholder(s) `{{{{NOMBRE}}}}` con: **{nombre_huesped}**")
                        else:
                            st.warning("⚠️ No se encontró el placeholder `{{NOMBRE}}` en la plantilla. El documento se descargará sin modificaciones.")

                        # Botón para descargar
                        nombre_archivo = f"Carta_Despedida_{nombre_huesped.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"

                        st.download_button(
                            label="📥 DESCARGAR CARTA DE DESPEDIDA",
                            data=buffer,
                            file_name=nombre_archivo,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key="btn_descargar_carta"
                        )

                        st.markdown("""
                        <div style="background-color: #1a1a2e; border-radius: 8px; padding: 10px; margin-top: 10px; border: 1px solid #333;">
                            <p style="color: #888; font-size: 0.75rem; margin: 0; text-align: center;">
                                💡 Descarga el archivo y ábrelo en Word para imprimir.<br>
                                El nombre del huésped ya está insertado en el documento.
                            </p>
                        </div>
                        """, unsafe_allow_html=True)

                    except ImportError:
                        st.error("❌ Falta instalar la librería `python-docx`")
                        st.code("pip install python-docx", language="bash")
                        st.info("💡 Ejecuta el comando anterior en tu terminal y reinicia la aplicación.")
                    except Exception as e:
                        st.error(f"❌ Error al procesar la plantilla: {str(e)}")
                        st.exception(e)

# ============================================================
# 3. FORMULARIO NUEVA RESERVA
# ============================================================
if mostrar_formulario:
    with st.container():
        st.subheader("📝 Nueva Reservación")

        if st.button("↩️ REGRESAR", key="regresar_nueva"):
            st.query_params.clear()
            st.rerun()

        with st.form("form_reserva"):
            c1, c2, c3, c4 = st.columns(4)
            eta_12h = c1.selectbox("ETA", options=horas_eta_12h, index=0)
            eta = mapa_12a24[eta_12h]
            name = c2.text_input("Name")
            qty = c3.number_input("Qty", min_value=0)
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
                conn = sqlite3.connect(db_path)
                cursor = conn.cursor()
                cursor.execute("INSERT INTO huespedes (eta, name, qty, room, email, check_in, check_out, res_number, phone, info, ird, hsk, rate, trans) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", 
                               (eta, name, qty, room, email, check_in.strftime("%b %d"), check_out.strftime("%b %d"), res_number, phone, info, ird, hsk, rate, trans))
                conn.commit()
                conn.close()
                st.query_params.clear(); st.rerun()

# ============================================================
# 4. FORMULARIO EDITAR RESERVA
# ============================================================
if mostrar_editar:
    fila_guardada = st.session_state.get("fila_seleccionada")

    if fila_guardada:
        with st.container():
            st.subheader("✏️ Editar Reservación")

            if st.button("↩️ REGRESAR", key="regresar_editar"):
                st.query_params.clear()
                st.rerun()

            try:
                check_in_dt = datetime.strptime(fila_guardada.get("check_in", ""), "%b %d").replace(year=datetime.now().year)
            except:
                check_in_dt = datetime.now()
            try:
                check_out_dt = datetime.strptime(fila_guardada.get("check_out", ""), "%b %d").replace(year=datetime.now().year)
            except:
                check_out_dt = datetime.now()

            with st.form("form_editar"):
                c1, c2, c3, c4 = st.columns(4)
                eta_actual_24h = fila_guardada.get("eta", "")
                eta_actual_12h = mapa_24a12.get(eta_actual_24h, "12:00 AM")
                eta_index = horas_eta_12h.index(eta_actual_12h) if eta_actual_12h in horas_eta_12h else 0
                eta_12h = c1.selectbox("ETA", options=horas_eta_12h, index=eta_index)
                eta = mapa_12a24[eta_12h]
                name = c2.text_input("Name", value=fila_guardada.get("name", ""))
                qty = c3.number_input("Qty", min_value=0, value=int(fila_guardada.get("qty", 0) or 0))
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
                    conn = sqlite3.connect(db_path)
                    cursor = conn.cursor()
                    cursor.execute("""
                        UPDATE huespedes 
                        SET eta=?, name=?, qty=?, room=?, email=?, check_in=?, check_out=?, 
                            res_number=?, phone=?, info=?, ird=?, hsk=?, rate=?, trans=?
                        WHERE id=?                    """, (
                        eta, name, qty, room, email, check_in.strftime("%b %d"), check_out.strftime("%b %d"),
                        res_number, phone, info, ird, hsk, rate, trans, fila_guardada["id"]
                    ))
                    conn.commit()
                    conn.close()
                    st.session_state.pop("fila_seleccionada", None)
                    st.query_params.clear()
                    st.success("Reserva actualizada correctamente.")
                    st.rerun()
    else:
        st.error("Por favor, selecciona una fila en la tabla primero.")
        if st.button("↩️ REGRESAR", key="regresar_editar_error"):
            st.query_params.clear()
            st.rerun()

# ============================================================
# 5. TABLA ÚNICA (SIEMPRE VISIBLE) CON BÚSQUEDA INTELIGENTE
# ============================================================
st.markdown("<div style='height: 5px;'></div>", unsafe_allow_html=True)
df_reservas = cargar_reservaciones()

# Aplicar filtro por fecha de checkout si se hizo click en el panel
if filtro_checkout:
    df_reservas = df_reservas[df_reservas['check_out'] == filtro_checkout]
    st.info(f"📅 Mostrando reservas que salen el: {filtro_checkout}")

# Aplicar filtro por fecha del date picker SOLO si está ACTIVAMENTE activado
if fecha_filtro_activo and filtro_fecha_date:
    fecha_filtro = datetime.strptime(filtro_fecha_date, "%Y-%m-%d").strftime("%b %d")
    df_reservas = df_reservas[df_reservas['check_in'] == fecha_filtro]
    st.info(f"📅 Mostrando reservas con check-in el: {fecha_filtro}")

# Aplicar búsqueda inteligente si hay texto en el buscador
if busqueda and busqueda.strip():
    busqueda_lower = busqueda.strip().lower()
    mask = df_reservas.astype(str).apply(
        lambda row: row.str.lower().str.contains(busqueda_lower, na=False).any(), axis=1
    )
    df_reservas = df_reservas[mask]
    if len(df_reservas) == 0:
        st.info(f"🔍 No se encontraron resultados para: '{busqueda}'")

gb = GridOptionsBuilder.from_dataframe(df_reservas)
gb.configure_selection(selection_mode="single", use_checkbox=False)

grid_return = AgGrid(
    df_reservas,
    gridOptions=gb.build(),
    height=620,
    theme="streamlit",
    key="tabla_principal_concierge",
    custom_css={
        ".ag-root-wrapper": {"background-color": "#101010 !important"},
        ".ag-cell": {"color": "white !important", "background-color": "#101010 !important"},
        ".ag-row-selected .ag-cell": {"background-color": "#00FFFF !important", "color": "#000000 !important", "font-weight": "bold !important"}
    }
)

# GUARDAR SELECCIÓN EN SESSION STATE cada vez que cambia
seleccion = grid_return.get("selected_rows")
if seleccion is not None and len(seleccion) > 0:
    if isinstance(seleccion, pd.DataFrame):
        st.session_state["fila_seleccionada"] = seleccion.iloc[0].to_dict()
    else:
        st.session_state["fila_seleccionada"] = dict(seleccion[0])
elif seleccion is not None and len(seleccion) == 0:
    st.session_state.pop("fila_seleccionada", None)

# ============================================================
# 6. LÓGICA DE BORRADO CON CLAVE
# ============================================================
if st.query_params.get("action") == "cancelar":
    st.subheader("❌ Cancelar Reservación")

    if st.button("↩️ REGRESAR", key="regresar_cancelar"):
        st.query_params.clear()
        st.rerun()

    fila_guardada = st.session_state.get("fila_seleccionada")

    if fila_guardada:
        id_a_borrar = fila_guardada["id"]

        st.warning(f"Se eliminará permanentemente la reserva ID: {id_a_borrar}")
        password = st.text_input("Ingrese clave de autorización:", type="password")

        if st.button("CONFIRMAR Y BORRAR"):
            if password == "D6msnp8a": 
                conn = sqlite3.connect(db_path)
                cursor = conn.cursor()
                cursor.execute("DELETE FROM huespedes WHERE id = ?", (id_a_borrar,))
                conn.commit()
                conn.close()
                st.session_state.pop("fila_seleccionada", None)
                st.query_params.clear()
                st.success("Registro eliminado correctamente.")
                st.rerun()
            else:
                st.error("Clave incorrecta.")
    else:
        st.error("Por favor, selecciona una fila en la tabla primero.")
        if st.button("↩️ REGRESAR", key="regresar_cancelar_error"):
            st.query_params.clear()
            st.rerun()
