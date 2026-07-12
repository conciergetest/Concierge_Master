import customtkinter as ctk
from tkinter import messagebox, ttk, filedialog, simpledialog
import sqlite3
import pandas as pd
from datetime import datetime, date, timedelta
from PIL import Image
import os
import sys
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from tkcalendar import Calendar
from docx import Document


# --- LIBRERÍAS PARA EL PDF ---
from fpdf import FPDF
import subprocess

def obtener_ruta_recurso(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def inicializar_tabla_agenda():
    conexion = sqlite3.connect(obtener_ruta_recurso("recepcion_final.db"))    
    cursor = conexion.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS agenda_interna (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT NOT NULL,
            puesto TEXT,
            departamento TEXT,
            extension TEXT,
            email TEXT
        )
    """)
    cursor.execute("SELECT COUNT(*) FROM agenda_interna")
    if cursor.fetchone()[0] == 0:
        contactos_iniciales = [
            ("Loreana Aiza", "Concierge Manager", "Concierge", "7368", "Loreana.Aiza@waldorfastoria.com"),
            ("Personal Concierge Oficina", "Concierge Agents", "Concierge", "7367", "personalconcierge.costarica@waldorfastoria.com"),
            ("Martin Kessler", "Director of Engineering", "Engineering", "7240", "martin.kessler@waldorfastoria.com")
        ]
        cursor.executemany("INSERT INTO agenda_interna (nombre, puesto, departamento, extension, email) VALUES (?, ?, ?, ?, ?)", contactos_iniciales)
    conexion.commit()
    conexion.close()

# =========================================================
# --- SPLASH SCREEN MODIFICADO ---
# =========================================================
class SplashScreen(ctk.CTkToplevel):
    def __init__(self, callback):
        super().__init__()
        self.callback = callback
        self.title("Cargando Concierge Master...")
        ancho, alto = 850, 580
        x = (self.winfo_screenwidth() // 2) - (ancho // 2)
        y = (self.winfo_screenheight() // 2) - (alto // 2)
        self.geometry(f"{ancho}x{alto}+{x}+{y}")
        self.overrideredirect(True)
        self.configure(fg_color="#0a0a0a")

        self.frame = ctk.CTkFrame(self, fg_color="transparent")
        self.frame.pack(fill="both", expand=True)

        # --- BLOQUE INTEGRADO PARA CARGA DE IMAGEN ---
        try:
            # Usa la función global obtener_ruta_recurso que ya tienes definida en tu main
            img_path = obtener_ruta_recurso("concierge_logo.png") 
            
            # Cargamos la imagen con las dimensiones adecuadas
            img_logo = ctk.CTkImage(light_image=Image.open(img_path), 
                                    dark_image=Image.open(img_path), 
                                    size=(820, 460))
            
            self.lbl_img = ctk.CTkLabel(self.frame, image=img_logo, text="")
            self.lbl_img.pack(pady=(15, 0))
        except Exception as e:
            # Si falla, imprimimos el error en consola para depurar y ponemos texto de emergencia
            print(f"Error al cargar imagen: {e}")
            self.lbl_img = ctk.CTkLabel(self.frame, text="🛎️\nWALDORF ASTORIA", font=("Garamond", 50), text_color="#d4af37")
            self.lbl_img.pack(pady=100)
        # --- FIN DEL BLOQUE ---

        self.progreso = ctk.CTkProgressBar(self.frame, width=700, height=10, progress_color="#d4af37", fg_color="#1e1e1e")
        self.progreso.pack(pady=(25, 10))
        self.progreso.set(0)

        self.lbl_status = ctk.CTkLabel(self.frame, text="INITIALIZING SYSTEM...", font=("Roboto", 12, "bold"), text_color="#d4af37")
        self.lbl_status.pack()

        self.animar_progreso()

    def animar_progreso(self):
        val = self.progreso.get()
        if val < 1.0:
            if val < 0.3: incremento, espera = 0.008, 60
            elif 0.3 <= val < 0.7: incremento, espera = 0.004, 100
            else: incremento, espera = 0.002, 120
            nuevo_val = min(val + incremento, 1.0)
            self.progreso.set(nuevo_val)
            if nuevo_val > 0.2: self.lbl_status.configure(text="LOADING DATABASE ASSETS...")
            if nuevo_val > 0.5: self.lbl_status.configure(text="CONNECTING TO WALDORF ASTORIA SERVER...")
            if nuevo_val > 0.8: self.lbl_status.configure(text="SYNCHRONIZING GUEST DATA...")
            if nuevo_val > 0.95: self.lbl_status.configure(text="WELCOME YA'LL... FRED WAYNE'S SYSTEM READY.")
            self.update_idletasks()
            self.after(espera, self.animar_progreso)
        else:
            self.after(1000, self.finalizar_splash)

    def finalizar_splash(self):
        self.destroy()
        self.callback()

class AlmanaqueDesplegable(ctk.CTkToplevel):
    def __init__(self, parent, boton_widget, **kwargs):
        super().__init__(parent, **kwargs)
        self.overrideredirect(True)
        self.withdraw()
        self.configure(fg_color="#1E1E1E")

        # --- CONTENEDOR PARA SELECTORES ---
        frame_selectores = ctk.CTkFrame(self, fg_color="transparent")
        frame_selectores.pack(fill="x", padx=15, pady=(15, 0))

        # Selector de Mes
        self.mes_var = ctk.StringVar(value=str(datetime.now().month))
        self.combo_mes = ctk.CTkComboBox(frame_selectores, values=[str(i) for i in range(1, 13)], 
                                         width=60, variable=self.mes_var, command=self.actualizar_calendario)
        self.combo_mes.pack(side="left", padx=2)

        # Selector de Año
        self.anio_var = ctk.StringVar(value=str(datetime.now().year))
        self.combo_anio = ctk.CTkComboBox(frame_selectores, values=[str(i) for i in range(2010, 2035)], 
                                          width=80, variable=self.anio_var, command=self.actualizar_calendario)
        self.combo_anio.pack(side="left", padx=2)

        # Calendario
        self.cal = Calendar(self, selectmode='day', font=("Roboto", 14), 
                            background='#1A1A1A', foreground='white', bordercolor='#2A2A2A',
                            headersbackground='#252525', headersforeground='#00CCCC', selectbackground='#007ACC',
                            normalbackground='#222222', normalforeground='white', weekendbackground='#282828',
                            weekendforeground='#FF5555', othermonthbackground='#1A1A1A', othermonthforeground='#555555')
        self.cal.pack(padx=20, pady=20)
        
        self.after(10, self.posicionar_y_mostrar, boton_widget)
        self.bind("<FocusOut>", self.verificar_cierre)

    def actualizar_calendario(self, *args):
        try:
            mes = int(self.mes_var.get())
            anio = int(self.anio_var.get())
            
            # 1. Definimos la fecha
            nueva_fecha = date(anio, mes, 1)
            
            # 2. Saltamos al mes/año específico
            self.cal.selection_set(nueva_fecha) # Selecciona el día 1
            self.cal.see(nueva_fecha)           # <--- ESTA ES LA CLAVE: 'see' obliga a la vista a saltar ahí
            
            self.cal.update()
        except Exception as e:
            print(f"Error al actualizar: {e}")

    def verificar_cierre(self, event):
        self.after(200, self._destruir_si_fuera)

    def _destruir_si_fuera(self):
        widget_con_foco = self.focus_get()
        if widget_con_foco is None or not str(widget_con_foco).startswith(str(self)):
            self.destroy()

    def posicionar_y_mostrar(self, boton):
        x = boton.winfo_rootx()
        y = boton.winfo_rooty() + boton.winfo_height() + 5
        self.geometry(f"+{x}+{y}")
        self.deiconify()
        self.focus_set()

# (CalculadoraPopup se mantiene igual - copia completa del original si necesitas)

class CalculadoraPopup(ctk.CTkToplevel):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.title("Calculadora")
        
        # --- DIMENSIONES MÁS GRANDES Y ELEGANTES ---
        self.geometry("380x560")  # Tamaño expandido premium
        self.resizable(False, False)
        self.configure(fg_color="#1A1A1A")  # Fondo oscuro elegante integrado
        
        # Mantener el foco y encima de la app principal
        self.transient(master)
        self.grab_set()
        self.focus_set()

        self.expression = ""

        # --- PANTALLA DE RESULTADOS (Diseño inspirado en smartphone) ---
        self.screen_frame = ctk.CTkFrame(self, fg_color="#121212", corner_radius=10)
        self.screen_frame.pack(padx=15, pady=15, fill="x")
        
        self.display = ctk.CTkLabel(
            self.screen_frame, 
            text="0", 
            anchor="e", 
            font=("Consolas", 34, "bold"), 
            text_color="#FFFFFF"
        )
        self.display.pack(padx=15, pady=20, fill="x")

        # --- CONTENEDOR DE BOTONES ---
        self.buttons_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.buttons_frame.pack(padx=15, pady=5, fill="both", expand=True)

        # Configuración de Grid simétrico (4 columnas, 5 filas)
        for i in range(4): self.buttons_frame.columnconfigure(i, weight=1, pad=8)
        for i in range(5): self.buttons_frame.rowconfigure(i, weight=1, pad=8)

        # Estilo estético basado en tu referencia corporativa
        self.crear_botones()
        
        # --- ENLACE MÁGICO CON EL TECLADO FÍSICO ---
        self.bind_keyboard()

    def crear_botones(self):
        # Distribución ejecutiva de botones
        botones = [
            ('C', 0, 0, '#E53935'),    ('±', 0, 1, '#333333'),  ('%', 0, 2, '#333333'),  ('÷', 0, 3, '#00B0F0'),
            ('7', 1, 0, '#262626'),    ('8', 1, 1, '#262626'),  ('9', 1, 2, '#262626'),  ('×', 1, 3, '#00B0F0'),
            ('4', 2, 0, '#262626'),    ('5', 2, 1, '#262626'),  ('6', 2, 2, '#262626'),  ('-', 2, 3, '#00B0F0'),
            ('1', 3, 0, '#262626'),    ('2', 3, 1, '#262626'),  ('3', 3, 2, '#262626'),  ('+', 3, 3, '#00B0F0'),
            ('0', 4, 0, '#262626'),    ('.', 4, 1, '#262626'),  ('⌫', 4, 2, '#333333'),  ('=', 4, 3, '#00F0FF')
        ]

        for texto, fila, col, color in botones:
            # Re-mapeo visual de operadores nativos para mantener estética smartphone
            action_text = texto
            if texto == '÷': action_text = '/'
            elif texto == '×': action_text = '*'
            
            btn = ctk.CTkButton(
                self.buttons_frame,
                text=texto,
                font=("Segoe UI", 20, "bold") if texto.isdigit() or texto == '.' else ("Segoe UI", 22, "bold"),
                fg_color=color,
                text_color="#FFFFFF" if color != '#00F0FF' else '#121212', # Texto oscuro en el igual para contraste cian
                hover_color=self.ajustar_hover(color),
                corner_radius=12,
                command=lambda t=action_text: self.procesar_click(t)
            )
            btn.grid(row=fila, column=col, sticky="nsew", padx=4, pady=4)

    def ajustar_hover(self, color):
        # Tonos oscurecidos automáticos para el hover de CustomTkinter
        transiciones = {'#E53935': '#B71C1C', '#00B0F0': '#0086B3', '#00F0FF': '#0099A1', '#262626': '#404040', '#333333': '#4D4D4D'}
        return transiciones.get(color, '#444444')

    def procesar_click(self, char):
        if char == 'C':
            self.expression = ""
        elif char == '⌫':
            self.expression = self.expression[:-1]
        elif char == '=':
            try:
                # Evalúa de manera segura la expresión matemática
                if self.expression:
                    resultado = eval(self.expression)
                    # Formatea enteros o decimales limpios
                    self.expression = str(int(resultado)) if resultado % 1 == 0 else f"{resultado:.4f}".rstrip('0').rstrip('.')
            except Exception:
                self.expression = "Error"
        elif char == '±':
            try:
                if self.expression:
                    if self.expression.startswith('-'): self.expression = self.expression[1:]
                    else: self.expression = '-' + self.expression
            except Exception: pass
        elif char == '%':
            try:
                if self.expression:
                    self.expression = str(eval(self.expression) / 100)
            except Exception: self.expression = "Error"
        else:
            if self.expression == "Error":
                self.expression = ""
            self.expression += str(char)

        self.display.configure(text=self.expression if self.expression != "" else "0")

    # --- ENLACES CLAVE DEL TECLADO NUMÉRICO ---
    def bind_keyboard(self):
        # Números estándar y del teclado numérico
        for i in range(10):
            self.bind(f"{i}", lambda event, num=i: self.procesar_click(num))
            self.bind(f"<KP_{i}>", lambda event, num=i: self.procesar_click(num))
            
        # Operadores del bloque numérico y teclado general
        self.bind("+", lambda event: self.procesar_click('+'))
        self.bind("<KP_Add>", lambda event: self.procesar_click('+'))
        
        self.bind("-", lambda event: self.procesar_click('-'))
        self.bind("<KP_Subtract>", lambda event: self.procesar_click('-'))
        
        self.bind("*", lambda event: self.procesar_click('*'))
        self.bind("<KP_Multiply>", lambda event: self.procesar_click('*'))
        
        self.bind("/", lambda event: self.procesar_click('/'))
        self.bind("<KP_Divide>", lambda event: self.procesar_click('/'))
        
        self.bind(".", lambda event: self.procesar_click('.'))
        self.bind("<KP_Decimal>", lambda event: self.procesar_click('.'))

        # Teclas de acción de ejecución
        self.bind("<Return>", lambda event: self.procesar_click('='))
        self.bind("<KP_Enter>", lambda event: self.procesar_click('='))
        self.bind("<BackSpace>", lambda event: self.procesar_click('⌫'))
        self.bind("<Escape>", lambda event: self.procesar_click('C'))
        self.bind("<Delete>", lambda event: self.procesar_click('C'))   

# =========================================================
# --- BASE DE DATOS ---
# =========================================================
def iniciar_db():
    ruta_db = obtener_ruta_recurso("recepcion_final.db")
    conn = sqlite3.connect(ruta_db)
    cursor = conn.cursor()
    cursor.execute("""CREATE TABLE IF NOT EXISTS huespedes (
        id INTEGER PRIMARY KEY AUTOINCREMENT, eta TEXT, name TEXT, qty TEXT, room TEXT, email TEXT,
        check_in TEXT, check_out TEXT, res_number TEXT, phone TEXT, info TEXT, ird TEXT, hsk TEXT, rate TEXT, trans TEXT
    )""")
    conn.commit()
    conn.close()

def clave_orden_fecha(fila):
    meses_map = {'Jan':1,'Feb':2,'Mar':3,'Apr':4,'May':5,'Jun':6,'Jul':7,'Aug':8,'Sep':9,'Set':9,'Oct':10,'Nov':11,'Dec':12,'Dic':12}
    try:
        texto_fecha = str(fila[6]).split()
        mes_texto = texto_fecha[0].capitalize()
        dia = int(texto_fecha[1])
        return datetime(2026, meses_map.get(mes_texto, 1), dia)
    except:
        return datetime(2099, 12, 31)

ctk.set_appearance_mode("dark")

# =========================================================
# --- VENTANA EDICIÓN (CORREGIDA) ---
# =========================================================
class VentanaEdicion(ctk.CTkToplevel):
    def __init__(self, parent, datos_huesped=None):
        super().__init__(parent)
        self.parent = parent
        self.id_huesped = datos_huesped[0] if datos_huesped else None
        self.title("Editor de Reservación")
        self.attributes("-topmost", True)
        ancho, alto = 500, 750
        x = (self.winfo_screenwidth() // 2) - (ancho // 2)
        y = (self.winfo_screenheight() // 2) - (alto // 2)
        self.geometry(f"{ancho}x{alto}+{x}+{y}")
        self.grab_set()

        self.frame_scroll = ctk.CTkScrollableFrame(self, label_text="DETALLES DE RESERVACIÓN")
        self.frame_scroll.pack(fill="both", expand=True, padx=20, pady=20)

        self.entries = {}
        self.campos = ["ETA", "NAME", "QTY", "ROOM", "EMAIL", "CHECK IN", "CHECK OUT", 
                       "RESERVATION NUMBER", "PHONE", "INFORMATION", "IRD", "HSK", "RATE", "TRANSPORTATION"]

        meses = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
        dias = [str(i).zfill(2) for i in range(1, 32)]
        fecha_hoy_mes = datetime.now().strftime("%b")
        fecha_hoy_dia = datetime.now().strftime("%d")

        for i, campo in enumerate(self.campos):
            ctk.CTkLabel(self.frame_scroll, text=campo, font=("Roboto", 11, "bold"), text_color="#00ced1").pack(pady=(12,0), anchor="w", padx=45)
            
            if campo == "ETA":
                horas = [f"{h:02d}:00 AM" for h in range(1, 12)] + ["12:00 PM"] + [f"{h:02d}:00 PM" for h in range(1, 12)] + ["12:00 AM"]
                widget = ctk.CTkComboBox(self.frame_scroll, values=horas, width=380, height=35)
                if datos_huesped: widget.set(datos_huesped[i+1])
                widget.pack(pady=(0, 5))
                self.entries[campo] = widget

            elif campo in ["CHECK IN", "CHECK OUT"]:
                frame_fecha = ctk.CTkFrame(self.frame_scroll, fg_color="transparent")
                frame_fecha.pack(pady=(0, 5))
                sel_mes = ctk.CTkOptionMenu(frame_fecha, values=meses, width=185, height=35)
                sel_dia = ctk.CTkOptionMenu(frame_fecha, values=dias, width=185, height=35)
                sel_mes.pack(side="left", padx=5)
                sel_dia.pack(side="left", padx=5)
                if datos_huesped:
                    partes = str(datos_huesped[i+1]).split()
                    if len(partes) == 2:
                        sel_mes.set(partes[0])
                        sel_dia.set(partes[1])
                else:
                    sel_mes.set(fecha_hoy_mes)
                    sel_dia.set(fecha_hoy_dia)
                self.entries[campo] = (sel_mes, sel_dia)

            elif campo == "TRANSPORTATION":
                widget = ctk.CTkComboBox(self.frame_scroll, values=["Car Rental", "Own Car", "Relaxury", "None"], width=380, height=35)
                if datos_huesped: widget.set(datos_huesped[i+1])
                widget.pack(pady=(0, 5))
                self.entries[campo] = widget
                
            else:
                widget = ctk.CTkEntry(self.frame_scroll, width=380, height=35)
                if datos_huesped:
                    valor_db = datos_huesped[i+1]
                    if campo == "RATE" and (valor_db is None or str(valor_db).strip() == ""):
                        widget.insert(0, "$0")
                    else:
                        widget.insert(0, str(valor_db))
                else:
                    if campo == "RATE":
                        widget.insert(0, "$0")
                widget.pack(pady=(0, 5))
                self.entries[campo] = widget

        btn_texto = "ACTUALIZAR DATOS" if self.id_huesped else "REGISTRAR ENTRADA"
        self.btn_guardar = ctk.CTkButton(self, text=btn_texto, fg_color="#00ced1", text_color="black", 
                                         height=45, font=("Roboto", 14, "bold"), command=self.guardar)
        self.btn_guardar.pack(pady=25)

    # --- MÉTODO GUARDAR FUERA DEL __init__ ---
    def guardar(self):
        try:
            datos_finales = []
            for campo in self.campos:
                valor = self.entries[campo]
                if isinstance(valor, tuple): 
                    mes, dia = valor[0].get(), valor[1].get()
                    datos_finales.append(f"{mes} {dia}")
                else:
                    datos_finales.append(valor.get())

            if not datos_finales[1]: 
                messagebox.showwarning("Atención", "El nombre es obligatorio")
                return

            conn = sqlite3.connect("recepcion_final.db")
            cursor = conn.cursor()
            
            if self.id_huesped is None:
                cursor.execute("""INSERT INTO huespedes (eta, name, qty, room, email, check_in, check_out, 
                                  res_number, phone, info, ird, hsk, rate, trans) 
                                  VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", tuple(datos_finales))
            else:
                cursor.execute("""UPDATE huespedes SET eta=?, name=?, qty=?, room=?, email=?, check_in=?, check_out=?, 
                                  res_number=?, phone=?, info=?, ird=?, hsk=?, rate=?, trans=? WHERE id=?""", 
                                  tuple(datos_finales + [self.id_huesped]))
            
            conn.commit()
            conn.close()
            
            # --- LÓGICA DE ACTUALIZACIÓN Y ORDENAMIENTO ---
            if hasattr(self.parent, 'actualizar_tabla'):
                fecha_del_huesped = datos_finales[5] 
                
                # Priorizar búsqueda si hay texto, sino usar la fecha
                busqueda = self.parent.search_entry.get() if hasattr(self.parent, 'search_entry') else ""
                
                if busqueda.strip():
                    self.parent.actualizar_tabla(query=busqueda)
                else:
                    self.parent.actualizar_tabla(fecha=fecha_del_huesped)
                
                # --- LA MAGIA: ORDENAR ALFABÉTICAMENTE AL TERMINAR ---
                if hasattr(self.parent, 'ordenar_tabla_alfabeticamente'):
                    self.parent.ordenar_tabla_alfabeticamente()
            
            self.destroy()

        except Exception as e:
            messagebox.showerror("Error de Guardado", f"Detalle del error:\n{str(e)}")

    # =========================================================
# --- APP PRINCIPAL (CORREGIDA Y CON ORDEN ALFABÉTICO) ---
# =========================================================
class AppHuespedes(ctk.CTk):
    def __init__(self):
        super().__init__()
        iniciar_db()
        inicializar_tabla_agenda()
        self.title("Concierge Master v4.8 - Waldorf Astoria")
        self.geometry("1450x850")
        self.search_timer = None

        # --- SISTEMA DE INACTIVIDAD ---
        # Usaremos 600,000 para 10 minutos
        self.inactividad_timeout = 600000 
        self.timer_id = None
        self.reset_timer()
        
        # Vincular eventos a la ventana principal
        self.bind_all("<Any-Button>", self.reset_timer)
        self.bind_all("<Any-KeyPress>", self.reset_timer)

        # --- INICIALIZACIÓN DEL TOOLTIP DE NOCHES ---
        self.tooltip_noches = ctk.CTkToplevel(self)
        self.tooltip_noches.withdraw()  # Esto lo mantiene oculto al inicio
        self.tooltip_noches.overrideredirect(True) # Quita los bordes de la ventana
        self.lbl_tooltip = ctk.CTkLabel(self.tooltip_noches, text="", fg_color="#00CED1", text_color="black", corner_radius=5)
        self.lbl_tooltip.pack(padx=5, pady=2)
        
        # Variable para recordar el filtro de fecha
        self.filtro_fecha_actual = None

        # Barra superior
        self.frame_top = ctk.CTkFrame(self, height=70, fg_color="#1a1a1a")
        self.frame_top.pack(side="top", fill="x")

        self.btn_nuevo = ctk.CTkButton(self.frame_top, text="+ NUEVA RESERVA", fg_color="#00ced1", text_color="black", font=("Roboto", 12, "bold"), command=self.abrir_nuevo)
        self.btn_nuevo.pack(side="left", padx=20)
        self.btn_importar = ctk.CTkButton(self.frame_top, text="📥 IMPORTAR EXCEL", fg_color="#3498db", text_color="white", font=("Roboto", 12, "bold"), command=self.importar_excel)
        self.btn_importar.pack(side="left", padx=10)
        self.btn_exportar = ctk.CTkButton(self.frame_top, text="📊 EXPORTAR EXCEL", fg_color="#2ecc71", text_color="black", font=("Roboto", 12, "bold"), command=self.exportar_excel)
        self.btn_exportar.pack(side="left", padx=10)
        self.btn_carta = ctk.CTkButton(self.frame_top, text="✉️ CARTA DESPEDIDA", fg_color="#FFB300", text_color="black", font=("Roboto", 12, "bold"), command=self.generar_carta_despedida)
        self.btn_carta.pack(side="left", padx=10)
        self.btn_cancelar = ctk.CTkButton(self.frame_top, text="❌ CANCELAR/BORRAR", fg_color="#e74c3c", text_color="white", font=("Roboto", 12, "bold"), command=self.cancelar_reserva)
        self.btn_cancelar.pack(side="left", padx=10)
        self.btn_reporte = ctk.CTkButton(self.frame_top, text="📋 REPORTE DIA", fg_color="#8e44ad", text_color="white", font=("Roboto", 12, "bold"), command=self.generar_reporte_diario)
        self.btn_reporte.pack(side="left", padx=10)
        
        self.btn_almanaque = ctk.CTkButton(self.frame_top, text="📅", font=("Segoe UI", 16), width=50, height=35, fg_color="#333333", hover_color="#00ced1", command=lambda: AlmanaqueDesplegable(self, self.btn_almanaque))
        self.btn_almanaque.pack(side="left", padx=5)

        self.btn_calculadora = ctk.CTkButton(self.frame_top, text="🧮", font=("Segoe UI", 16), width=50, height=35, fg_color="#333333", hover_color="#00ced1", command=self.abrir_calculadora)
        self.btn_calculadora.pack(side="left", padx=5)

        self.lbl_reloj = ctk.CTkLabel(self.frame_top, text="", font=("Roboto", 18, "bold"), text_color="#00ced1")
        self.lbl_reloj.pack(side="right", padx=25)
        self.actualizar_reloj()

        self.btn_cerrar = ctk.CTkButton(self.frame_top, text="⏻", font=("Segoe UI", 20, "bold"), width=50, height=35, fg_color="#c0392b", hover_color="#e74c3c", command=self.cerrar_programa)
        self.btn_cerrar.pack(side="right", padx=20)

        # === SEGUNDA BARRA ===
        self.frame_main_top = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_main_top.pack(fill="x", padx=20, pady=10)

        self.frame_pronostico = ctk.CTkFrame(self.frame_main_top, fg_color="#262626", border_width=1, border_color="#333333")
        self.frame_pronostico.pack(side="left", padx=(0, 20))
        ctk.CTkLabel(self.frame_pronostico, text="Checking Out Rooms", font=("Roboto", 13, "bold"), text_color="white").grid(row=0, column=0, columnspan=2, pady=5)
        self.labels_pronostico = []
        for i in range(8):
            row = (i % 4) + 1
            col = 0 if i < 4 else 1
            lbl = ctk.CTkLabel(self.frame_pronostico, text="--: [0]", font=("Roboto", 11), text_color="#00ced1")
            lbl.grid(row=row, column=col, padx=20, pady=2, sticky="w")
            self.labels_pronostico.append(lbl)

        self.frame_derecho = ctk.CTkFrame(self.frame_main_top, fg_color="transparent")
        self.frame_derecho.pack(side="right", padx=10)
        self.frame_grafico = ctk.CTkFrame(self.frame_derecho, fg_color="transparent")
        self.frame_grafico.pack(side="right", padx=10)

        try:
            ruta_logo = obtener_ruta_recurso("CONCIERGE LOGO.png")
            img_logo_data = Image.open(ruta_logo) 
            img_logo = ctk.CTkImage(light_image=img_logo_data, dark_image=img_logo_data, size=(120, 75))
            self.lbl_logo = ctk.CTkLabel(self.frame_derecho, image=img_logo, text="")
            self.lbl_logo.pack(side="right", padx=10) 
        except Exception as e:
            print(f"Error al cargar el logo: {e}")

        self.frame_busqueda = ctk.CTkFrame(self.frame_main_top, fg_color="transparent")
        self.frame_busqueda.pack(side="left", fill="x", expand=True, padx=10)

        self.btn_calendario = ctk.CTkButton(self.frame_busqueda, text="📅 FILTRAR", width=120, height=45, fg_color="#333333", hover_color="#00ced1", command=self.mostrar_selector_fecha)
        self.btn_calendario.pack(side="left", padx=5)

        self.search_entry = ctk.CTkEntry(self.frame_busqueda, placeholder_text="🔍 Búsqueda rápida...", width=350, height=45, border_color="#00ced1")
        self.search_entry.pack(side="left", padx=5)
        self.search_entry.bind("<KeyRelease>", self.buscar_universal)

        self.lbl_contador = ctk.CTkLabel(self.frame_busqueda, text="0", font=("Roboto", 24, "bold"), text_color="#00ced1")
        self.lbl_contador.pack(side="left", padx=10)

        self.btn_reset = ctk.CTkButton(self.frame_busqueda, text="🔄", width=45, height=45, fg_color="#e74c3c", command=self.resetear_filtros)
        self.btn_reset.pack(side="left", padx=5)

        self.btn_agenda = ctk.CTkButton(self.frame_busqueda, text="📞 AGENDA", width=100, height=45, fg_color="#2980b9", text_color="white", font=("Roboto", 12, "bold"), command=self.abrir_agenda)
        self.btn_agenda.pack(side="left", padx=5)

        # --- TABLA ---
        self.frame_tabla = ctk.CTkFrame(self)
        self.frame_tabla.pack(fill="both", expand=True, padx=20, pady=10)

        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", background="#1a1a1a", foreground="white", fieldbackground="#1a1a1a", rowheight=30, borderwidth=0, highlightthickness=0)
        style.configure("Treeview.Heading", background="#333333", foreground="white", font=("Roboto", 11, "bold"), relief="flat", borderwidth=0)
        style.map("Treeview.Heading", background=[('active', '#333333'), ('pressed', '#222222')])
        style.map("Treeview", background=[('selected', '#00ced1')], foreground=[('selected', 'black')])

        columnas = ("ID", "ETA", "NAME", "QTY", "ROOM", "EMAIL", "CHECK IN", "CHECK OUT", "RESERVATION #", "PHONE", "INFORMATION", "IRD", "HSK", "RATE", "TRANSPORTATION")
        self.tabla = ttk.Treeview(self.frame_tabla, columns=columnas, show="headings", style="Treeview")
        self.tabla.tag_configure('vip_row', foreground="#FFD700")
        self.tabla.tag_configure('cancelado', background='#424242', foreground='#9e9e9e')
        self.tabla.tag_configure('checkout_hecho', foreground='#666666')

        scroll_y = ttk.Scrollbar(self.frame_tabla, orient="vertical", command=self.tabla.yview)
        scroll_x = ttk.Scrollbar(self.frame_tabla, orient="horizontal", command=self.tabla.xview)
        self.tabla.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        for col in columnas:
            self.tabla.heading(col, text=col)
            self.tabla.column(col, width=150, anchor="center")
        self.tabla.column("ID", width=0, stretch=False)

        self.tabla.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")
        self.frame_tabla.grid_columnconfigure(0, weight=1)
        self.frame_tabla.grid_rowconfigure(0, weight=1)

        self.tabla.bind("<Double-1>", self.abrir_edicion)
        self.tabla.bind("<Motion>", self.mostrar_noches_hover)
        self.tabla.bind("<Leave>", self.ocultar_noches_hover)
        self.actualizar_tabla()
        
        # Ordenar al iniciar
        self.ordenar_tabla_alfabeticamente()
        self.after(200, self.forzar_renderizado)

    # --- MÉTODO REEMPLAZADO ---
    def reset_timer(self, event=None):
        # Cancelamos el temporizador usando siempre el mismo nombre
        if self.timer_id is not None:
            self.after_cancel(self.timer_id)
        
        # Iniciamos el temporizador de 10 minutos (600,000ms)
        self.timer_id = self.after(self.inactividad_timeout, self.cerrar_por_inactividad)

    def cerrar_por_inactividad(self):
        try:
            print("Cerrando por inactividad...")
            self.destroy()
        except Exception as e:
            print(f"Error al cerrar: {e}")
            self.destroy()

    def generar_reporte_diario(self):
        import pandas as pd
        import os
        import sqlite3
        from datetime import datetime, timedelta
        from tkinter import messagebox
        
        # 1. Preparar fechas con el AÑO ACTUAL
        anio_actual = datetime.now().year
        hoy = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        manana = hoy + timedelta(days=1)
        
        # 2. Leer datos
        conn = sqlite3.connect("recepcion_final.db")
        df = pd.read_sql_query("SELECT * FROM huespedes", conn)
        conn.close()

        # 3. Conversión de columnas
        df['check_in_dt'] = pd.to_datetime(df['check_in'] + f" {anio_actual}", format='%b %d %Y', errors='coerce')
        df['check_out_dt'] = pd.to_datetime(df['check_out'] + f" {anio_actual}", format='%b %d %Y', errors='coerce')
        
        df['info_upper'] = df['info'].fillna('').str.upper()
        df['room'] = df['room'].astype(str)

        # 4. Filtros
        en_casa = df[(df['check_in_dt'] <= hoy) & (df['check_out_dt'] >= hoy)]
        salen_hoy = df[df['check_out_dt'] == hoy]
        salen_manana = df[df['check_out_dt'] == manana]

        def obtener_rooms(df_filtrado):
            if df_filtrado.empty: return ""
            return " - ".join(df_filtrado['room'].tolist())

        # 5. Crear Resumen
        data_resumen = {
            "Concepto": ["Total en casa", "VIP en casa", "Salen hoy (Total)", "VIP que salen hoy", "Salen mañana (Total)"],
            "Cantidad": [
                len(en_casa),
                len(en_casa[en_casa['info_upper'].str.contains('VIP', na=False)]),
                len(salen_hoy),
                len(salen_hoy[salen_hoy['info_upper'].str.contains('VIP', na=False)]),
                len(salen_manana)
            ],
            "ROOM": [
                obtener_rooms(en_casa),
                obtener_rooms(en_casa[en_casa['info_upper'].str.contains('VIP', na=False)]),
                obtener_rooms(salen_hoy),
                obtener_rooms(salen_hoy[salen_hoy['info_upper'].str.contains('VIP', na=False)]),
                obtener_rooms(salen_manana)
            ]
        }
        df_resumen = pd.DataFrame(data_resumen)
        
        # 6. Lógica dinámica para guardar en CUALQUIER escritorio
        home = os.path.expanduser("~")
        ruta_escritorio = os.path.join(home, "Desktop")
        
        # Seguro extra para OneDrive
        if not os.path.exists(ruta_escritorio):
            ruta_escritorio = os.path.join(home, "OneDrive", "Desktop")
            
        # Generar nombre dinámico con la fecha
        nombre_archivo = f"Reporte_Briefing_{datetime.now().strftime('%d%m%Y')}.xlsx"
        ruta_final = os.path.join(ruta_escritorio, nombre_archivo)
        
        # 7. Guardar y Confirmar
        df_resumen.to_excel(ruta_final, index=False)
        messagebox.showinfo("Reporte Generado", f"Archivo creado correctamente en:\n{ruta_final}") 
    
    def ordenar_tabla_alfabeticamente(self):
        items = [(self.tabla.item(item)['values'], item) for item in self.tabla.get_children()]
        items.sort(key=lambda x: str(x[0][2]).lower())
        for i, (values, item_id) in enumerate(items):
            self.tabla.move(item_id, '', i)
    # HASTA ACÁ TERMINA APPHUESPEDES

    # --- MÉTODO PARA FILTRAR FECHA CON MEMORIA ---
    def aplicar_filtro_fecha(self, fecha):
        self.filtro_fecha_actual = fecha_seleccionada  # <--- Esto es vital
        self.actualizar_tabla(fecha=fecha_seleccionada)
        
    def resetear_filtros(self):
        self.filtro_fecha_actual = None
        self.search_entry.delete(0, 'end')
        self.actualizar_tabla() 

    def cerrar_programa(self):
        from tkinter import messagebox
        if messagebox.askyesno("Salir", "¿Estás seguro de que quieres cerrar Concierge Master?"):
            self.destroy()


    # --- MÉTODO PARA REFRESCAR LA TABLA (OPTIMIZADO) ---
    def actualizar_tabla(self, query=None, fecha=None):
        # SI VIENE UNA FECHA, LA GUARDAMOS PARA QUE NO SE OLVIDE
        if fecha: self.filtro_fecha_actual = fecha
        
        # Limpiar la tabla actual
        for item in self.tabla.get_children():
            self.tabla.delete(item)
            
        conn = sqlite3.connect("recepcion_final.db")
        cursor = conn.cursor()
        
        # Consultas SIN 'ORDER BY' para que nuestra función de ordenamiento tome el mando
        if query:
            cursor.execute("SELECT * FROM huespedes WHERE name LIKE ? OR room LIKE ?", 
                           ('%'+query+'%', '%'+query+'%'))
        elif self.filtro_fecha_actual: 
            cursor.execute("SELECT * FROM huespedes WHERE check_in = ?", (self.filtro_fecha_actual,))
        else:
            cursor.execute("SELECT * FROM huespedes")
            
        for fila in cursor.fetchall():
            self.tabla.insert("", "end", values=fila)
        conn.close()
        
        # --- AQUÍ ESTÁ LA MAGIA: ORDENAR ALFABÉTICAMENTE CADA VEZ QUE LA TABLA SE LLENA ---
        if hasattr(self, 'ordenar_tabla_alfabeticamente'):
            self.ordenar_tabla_alfabeticamente()
        # -------------------------------------
        
        # El resto de tu código que ya funcionaba (orden, formato, tags)
        def orden_seguro(fila):
            try: return clave_orden_fecha(fila)
            except Exception: return (9999, 12, 31)
            
        for fila in sorted(cursor.fetchall(), key=orden_seguro):
            datos_lista = list(fila)
            # ... (todo tu código de formato de moneda y tags aquí) ...
            self.insertar_fila_segura(datos_lista, tags)
            
        conn.close()
        self.actualizar_contador()
        self.actualizar_pronostico()
        self.generar_estadisticas()

    # --- MÉTODO PARA ABRIR LA AGENDA ---
    def abrir_agenda(self):
        VentanaAgenda(self)

    def abrir_calculadora(self):
        if not hasattr(self, "_calc_window") or not self._calc_window.winfo_exists():
            self._calc_window = CalculadoraPopup(self)
        else:
            self._calc_window.focus()

    def forzar_renderizado(self):
        self.update_idletasks()
        self.update()

    def abrir_nuevo(self):
        VentanaEdicion(self)

    def abrir_edicion(self, event=None):
        sel = self.tabla.selection()
        if sel:
            datos_huesped = self.tabla.item(sel[0])['values']
            VentanaEdicion(self, datos_huesped=datos_huesped)
        else:
            messagebox.showwarning("Atención", "Seleccione una reserva de la tabla.")

    # === Resto de métodos (generar_estadisticas, cancelar_reserva, importar_excel, etc.) ===
    # Copia aquí todos los métodos del original que están después de __init__ en tu archivo.

    def generar_estadisticas(self):
        for widget in self.frame_grafico.winfo_children(): widget.destroy()
        
        categorias = {
            'VIP': 0, 
            'BIRTHDAY': 0, 
            'HONEYMOON': 0, 
            'ANNIVERSARY': 0, 
            'BABYMOON': 0, 
            'TEAM MEMBER': 0,
            'LEISURE': 0
        }
        
        total_reservas = 0
        
        for item in self.tabla.get_children():
            valores = self.tabla.item(item)['values']
            texto_busqueda = f"{valores[10]} {valores[11]}".upper()
            total_reservas += 1
            
            es_especial = False
            
            if 'VIP' in texto_busqueda: 
                categorias['VIP'] += 1
                es_especial = True
            if 'BIRT' in texto_busqueda: 
                categorias['BIRTHDAY'] += 1
                es_especial = True
            if 'HONEY' in texto_busqueda: 
                categorias['HONEYMOON'] += 1
                es_especial = True
            if 'ANIV' in texto_busqueda or 'ANNI' in texto_busqueda: 
                categorias['ANNIVERSARY'] += 1
                es_especial = True
            if 'BABY' in texto_busqueda: 
                categorias['BABYMOON'] += 1
                es_especial = True
            if 'TEAM' in texto_busqueda or 'MEMBER' in texto_busqueda: 
                categorias['TEAM MEMBER'] += 1
                es_especial = True
                
            if not es_especial:
                categorias['LEISURE'] += 1
            
        # Devolvemos la figura a un margen superior normal (top=0.92) para aprovechar bien el espacio
        fig, ax = plt.subplots(figsize=(4.8, 2.9), facecolor='#1a1a1a')
        
        nombres, valores = list(categorias.keys()), list(categorias.values())
        colores = ['#00ced1', '#ff4757', '#ffa500', '#2ecc71', '#a29bfe', '#d4af37', '#0d47a1']
        
        bars = ax.barh(nombres, valores, color=colores)
        ax.set_facecolor('#1a1a1a')
        ax.tick_params(axis='both', colors='white', labelsize=9)
        plt.subplots_adjust(left=0.35, top=0.92, bottom=0.08)
        
        for spine in ax.spines.values(): spine.set_visible(False)
        ax.xaxis.set_visible(False) 
        
        for bar in bars:
            width = bar.get_width()
            ax.text(width, bar.get_y() + bar.get_height()/2, f' {int(width)}', va='center', color='white', fontweight='bold', fontsize=9)
            
        # --- UBICACIÓN MÁGICA EN EL MEDIO DERECHO ---
        # Colocamos X en 0.85 (bien a la derecha) y Y en 0.50 (exactamente a la mitad de la altura)
        ax.text(
            0.85, 0.50, f"{total_reservas}", 
            transform=ax.transAxes, color='#00F0FF', fontsize=16, fontweight='bold',
            va='center', ha='center',
            bbox=dict(boxstyle='circle,pad=0.4', facecolor='#112233', edgecolor='#00F0FF', lw=1.5)
        )
        
        canvas = FigureCanvasTkAgg(fig, master=self.frame_grafico)
        canvas.get_tk_widget().pack()
        canvas.draw()
        plt.close(fig)

        
    def cancelar_reserva(self):
        # 1. Obtener selección
        seleccion = self.tabla.selection()
        if not seleccion:
            messagebox.showwarning("Atención", "Seleccione una reserva primero.")
            return
        
        item = self.tabla.item(seleccion)
        valores = item["values"]
        res_id = valores[0]
        nombre_huesped = valores[2]

        # 2. Confirmación simple (Yes/No)
        # Si presiona 'No', la función termina aquí y no hace nada.
        if not messagebox.askyesno("Confirmar", f"¿Está seguro de eliminar la reserva de {nombre_huesped}?"):
            return
            
        # 3. Validación de contraseña
        password_ingresada = simpledialog.askstring(
            "Validación de Seguridad", 
            "Introduzca la contraseña de supervisor:", 
            show="*"
        )
        
        if password_ingresada != "D6msnp8a":
            messagebox.showerror("Error", "Contraseña incorrecta. Acción cancelada.")
            return
            
        # 4. Procesamiento en BD (Eliminación directa)
        try:
            conn = sqlite3.connect("recepcion_final.db")
            cursor = conn.cursor()
            # Ejecutamos el borrado real
            cursor.execute("DELETE FROM huespedes WHERE id = ?", (res_id,))
            conn.commit()
            conn.close()
            
            # Actualizamos la vista
            self.actualizar_tabla()
            messagebox.showinfo("Éxito", "Reserva eliminada correctamente.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo eliminar: {str(e)}")

    def importar_excel(self):
        ruta_archivo = filedialog.askopenfilename(title="Seleccionar Plantilla de Huéspedes", filetypes=[("Archivos de Excel", "*.xlsx *.xls")])
        if not ruta_archivo: return
        try:
            df = pd.read_excel(ruta_archivo)
            df.columns = [c.strip().upper() for c in df.columns] 
            df = df.fillna('')
            conn = sqlite3.connect("recepcion_final.db")
            cursor = conn.cursor()
            def formatear_fecha_excel(valor):
                if hasattr(valor, 'strftime'): return valor.strftime('%b %d')
                return str(valor)
            duplicados, contador_nuevos = [], 0
            for _, fila in df.iterrows():
                reserva = str(fila.get('RESERVATION #', '')).strip()
                nombre_huesped = str(fila.get('NAME', 'SIN NOMBRE'))
                cursor.execute("SELECT 1 FROM huespedes WHERE res_number = ?", (reserva,))
                if cursor.fetchone():
                    duplicados.append(f"Reserva: {reserva} - {nombre_huesped}")
                else:
                    datos = (
                        str(fila.get('ETA', '')), 
                        nombre_huesped, 
                        str(fila.get('QTY', '0')), 
                        str(fila.get('ROOM', '')), 
                        str(fila.get('EMAIL', '')), 
                        formatear_fecha_excel(fila.get('CHECK IN', '')), 
                        formatear_fecha_excel(fila.get('CHECK OUT', '')), 
                        reserva, 
                        str(fila.get('PHONE', '')), 
                        str(fila.get('INFORMATION', '')), 
                        str(fila.get('IRD', '')), 
                        str(fila.get('HSK', '')), 
                        str(fila.get('RATE', '')),    # <--- AQUÍ ESTABA EL ESPACIO VACÍO
                        str(fila.get('TRANSPORTATION', ''))
                    )
                    cursor.execute("INSERT INTO huespedes (eta, name, qty, room, email, check_in, check_out, res_number, phone, info, ird, hsk, rate, trans) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)", datos)
                    contador_nuevos += 1
            conn.commit(); conn.close(); self.actualizar_tabla()
            mensaje = f"Proceso completado.\n\n✅ Nuevos cargados: {contador_nuevos}"
            if duplicados:
                detalle = "\n".join(duplicados[:10]); mensaje += f"\n\n⚠️ Omitidos duplicados:\n{detalle}"
                messagebox.showwarning("Atención", mensaje)
            else: messagebox.showinfo("Éxito", mensaje)
        except Exception as e: messagebox.showerror("Error", f"Detalle: {str(e)}")

    def actualizar_reloj(self):
        self.lbl_reloj.configure(text=datetime.now().strftime("%B %d, %Y | %H:%M:%S").upper())
        self.after(1000, self.actualizar_reloj)

    def actualizar_pronostico(self):
        conn = sqlite3.connect("recepcion_final.db"); cursor = conn.cursor(); hoy = datetime.now()
        for i in range(8):
            ft = hoy + timedelta(days=i); fdb = ft.strftime("%b %d")
            cursor.execute("SELECT COUNT(*) FROM huespedes WHERE check_out LIKE ?", (f'%{fdb}%',))
            cant = cursor.fetchone()[0]
            col = "#ff4757" if cant >= 10 else "#e67e22" if cant >= 5 else "#00ced1"
            self.labels_pronostico[i].configure(text=f"{ft.strftime('%d-%b').upper()}: [{cant}]", text_color=col, cursor="hand2")
            self.labels_pronostico[i].bind("<Button-1>", lambda e, f=fdb: self.filtrar_por_checkout(f))
        conn.close()

    def filtrar_por_checkout(self, fecha):
        for item in self.tabla.get_children(): self.tabla.delete(item)
        conn = sqlite3.connect("recepcion_final.db"); cursor = conn.cursor()
        cursor.execute("SELECT * FROM huespedes WHERE check_out LIKE ?", (f'%{fecha}%',))
        for fila in sorted(cursor.fetchall(), key=clave_orden_fecha):
            tags = ['vip_row'] if "VIP" in str(fila[10]).upper() else []
            if str(fila[13]).upper() == "CANCELADO": tags.append('cancelado')
            self.tabla.insert("", "end", values=fila, tags=tuple(tags))
        conn.close(); self.actualizar_contador(); self.generar_estadisticas()

    def buscar_universal(self, event):
        term = self.search_entry.get().lower()
        for item in self.tabla.get_children(): self.tabla.delete(item)
        
        conn = sqlite3.connect("recepcion_final.db"); cursor = conn.cursor()
        q = """SELECT * FROM huespedes WHERE LOWER(name) LIKE ? OR LOWER(room) LIKE ? OR LOWER(res_number) LIKE ? OR 
               LOWER(email) LIKE ? OR LOWER(info) LIKE ? OR LOWER(ird) LIKE ? OR LOWER(hsk) LIKE ? OR LOWER(trans) LIKE ?"""
        p = f'%{term}%'; cursor.execute(q, (p, p, p, p, p, p, p, p))
        
        # ¡Blindaje total! Función de seguridad para evitar caídas por datos basura ("XXXX")
        def orden_seguro(fila):
            try:
                return clave_orden_fecha(fila)
            except Exception:
                return (9999, 12, 31)
        
        for fila in sorted(cursor.fetchall(), key=orden_seguro):
            datos_lista = list(fila)
            rate_val = str(datos_lista[12]).strip()
            
            if rate_val and rate_val != "None":
                limpio = rate_val.replace('$', '').replace(',', '').strip()
                if limpio.replace('.', '', 1).isdigit():
                    datos_lista[12] = f"${limpio}"
                
            tags = ['vip_row'] if "VIP" in str(datos_lista[10]).upper() else []
            if str(datos_lista[13]).upper() == "CANCELADO": tags.append('cancelado')
            
            # Encargamos la inserción y el icono 🚪 al nuevo motor centralizado
            self.insertar_fila_segura(datos_lista, tags)
            
        conn.close(); self.actualizar_contador()
        if self.search_timer: self.after_cancel(self.search_timer)
        self.search_timer = self.after(300, self.generar_estadisticas)

        # ==========================================
        # ¡ESTA LÍNEA ES LA MAGIA! 
        # Le regresa el cursor al cuadro inmediatamente
        # ==========================================
        self.search_entry.focus_set()

    def mostrar_selector_fecha(self):
        pop = ctk.CTkToplevel(self); pop.title("Fecha"); pop.geometry("300x250"); pop.grab_set()
        pop.attributes("-topmost", True) # Asegura que se mantenga arriba
        
        m = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
        lista_dias = [str(i).zfill(2) for i in range(1, 32)]
        
        # Selector de Meses (Se mantiene igual)
        sm = ctk.CTkOptionMenu(pop, values=m); sm.set(datetime.now().strftime("%b")); sm.pack(pady=10)
        
        # --- NUEVO CONTENEDOR PARA EL DÍA Y SUS FLECHAS ---
        frame_dia_cont = ctk.CTkFrame(pop, fg_color="transparent")
        frame_dia_cont.pack(pady=10)
        
        # Funciones internas para mover los días con las flechas
        def disminuir_dia():
            dia_actual = sd.get()
            if dia_actual in lista_dias:
                idx = lista_dias.index(dia_actual)
                if idx > 0:
                    sd.set(lista_dias[idx - 1])

        def aumentar_dia():
            dia_actual = sd.get()
            if dia_actual in lista_dias:
                idx = lista_dias.index(dia_actual)
                if idx < len(lista_dias) - 1:
                    sd.set(lista_dias[idx + 1])

        # Flecha Izquierda (Mover día hacia atrás)
        btn_izq = ctk.CTkButton(
            frame_dia_cont, text="«", font=("Roboto", 20, "bold"),
            width=35, height=30, fg_color="transparent", text_color="#00F0FF",
            hover_color="#1e1e1e", command=disminuir_dia
        )
        btn_izq.pack(side="left", padx=10)
        
        # Tu Selector de Días Original (Colocado al centro)
        sd = ctk.CTkOptionMenu(pop, values=lista_dias)
        sd.set(datetime.now().strftime("%d"))
        sd.pack(in_=frame_dia_cont, side="left")
        
        # Flecha Derecha (Mover día hacia adelante)
        btn_der = ctk.CTkButton(
            frame_dia_cont, text="»", font=("Roboto", 20, "bold"),
            width=35, height=30, fg_color="transparent", text_color="#00F0FF",
            hover_color="#1e1e1e", command=aumentar_dia
        )
        btn_der.pack(side="left", padx=10)
        
        # --- BOTÓN APLICAR (CORREGIDO Y ORDENADO) ---
        def filtrar():
            f = f"{sm.get()} {sd.get()}"
            
            # Primero filtramos
            self.filtrar_por_checkin(f)
            
            # ¡LA MAGIA! Después de filtrar, forzamos el ordenamiento
            if hasattr(self, 'ordenar_tabla_alfabeticamente'):
                self.ordenar_tabla_alfabeticamente()
                
            pop.destroy()
            
        ctk.CTkButton(pop, text="APLICAR", command=filtrar).pack(pady=15)
        
    def filtrar_por_checkin(self, fecha):
        from datetime import datetime
        for item in self.tabla.get_children(): self.tabla.delete(item)
        conn = sqlite3.connect("recepcion_final.db")
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM huespedes WHERE check_in LIKE ?", (f'%{fecha}%',))
        
        anio_actual = datetime.now().year
        hoy = datetime.now().date()
        
        # ¡Mantenemos tu blindaje intacto aquí también!
        def orden_seguro(fila):
            try:
                return clave_orden_fecha(fila)
            except Exception:
                return (9999, 12, 31) # Lo manda al final del ordenamiento si da error
        
        for fila in sorted(cursor.fetchall(), key=orden_seguro):
            datos_lista = list(fila)
            tags = ['vip_row'] if "VIP" in str(datos_lista[10]).upper() else []
            if str(datos_lista[13]).upper() == "CANCELADO": tags.append('cancelado')
            
            # --- DETECTOR INTELIGENTE DE CHECKOUT (CON ICONO 🚪) ---
            fecha_out_str = str(datos_lista[7]).strip() # Columna CHECK OUT
            if fecha_out_str and "XXXX" not in fecha_out_str.upper() and "NONE" not in fecha_out_str.upper():
                try:
                    fecha_out_dt = datetime.strptime(f"{fecha_out_str} {anio_actual}", "%b %d %Y").date()
                    # Modificamos el texto para inyectar la puertecita en lugar de cambiar colores
                    if fecha_out_dt <= hoy and str(datos_lista[13]).upper() != "CANCELADO":
                        datos_lista[7] = f"{fecha_out_str} 🚪"
                except Exception:
                    pass
                
            self.tabla.insert("", "end", values=datos_lista, tags=tuple(tags))
            
        conn.close()
        self.actualizar_contador()
        self.generar_estadisticas()

    def actualizar_tabla(self, query=None, fecha=None):
        from datetime import datetime
        # Limpiamos la tabla
        for item in self.tabla.get_children(): 
            self.tabla.delete(item)
            
        conn = sqlite3.connect("recepcion_final.db")
        cursor = conn.cursor()

        # Lógica inteligente de filtrado
        if query:
            cursor.execute("SELECT * FROM huespedes WHERE name LIKE ? OR room LIKE ? ORDER BY eta ASC", 
                           ('%'+query+'%', '%'+query+'%'))
        elif fecha:
            cursor.execute("SELECT * FROM huespedes WHERE check_in = ? ORDER BY eta ASC", (fecha,))
        else:
            cursor.execute("SELECT * FROM huespedes")

        # Lógica de ordenamiento y formato que ya tenías
        def orden_seguro(fila):
            try: return clave_orden_fecha(fila)
            except Exception: return (9999, 12, 31)
        
        # Procesamos filas con el orden y formato original
        for fila in sorted(cursor.fetchall(), key=orden_seguro):
            datos_lista = list(fila)
            # Formato de moneda
            rate_val = str(datos_lista[12]).strip()
            if rate_val and rate_val != "None":
                limpio = rate_val.replace('$', '').replace(',', '').strip()
                if limpio.replace('.', '', 1).isdigit():
                    datos_lista[12] = f"${limpio}"
            
            # Tags (VIP/Cancelado)
            tags = ['vip_row'] if "VIP" in str(datos_lista[10]).upper() else []
            if str(datos_lista[13]).upper() == "CANCELADO": tags.append('cancelado')
            
            # Insertamos con formato
            self.insertar_fila_segura(datos_lista, tags)
            
        conn.close()
        
        # Funciones de cierre que querías mantener
        self.actualizar_contador()
        self.actualizar_pronostico()
        self.generar_estadisticas()

    def filtrar_por_checkin(self, fecha):
        for item in self.tabla.get_children(): self.tabla.delete(item)
        conn = sqlite3.connect("recepcion_final.db")
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM huespedes WHERE check_in LIKE ?", (f'%{fecha}%',))
        
        def orden_seguro(fila):
            try: return clave_orden_fecha(fila)
            except Exception: return (9999, 12, 31)
        
        for fila in sorted(cursor.fetchall(), key=orden_seguro):
            datos_lista = list(fila)
            tags = ['vip_row'] if "VIP" in str(datos_lista[10]).upper() else []
            if str(datos_lista[13]).upper() == "CANCELADO": tags.append('cancelado')
            
            # Usamos el nuevo insertador automático
            self.insertar_fila_segura(datos_lista, tags)
            
        conn.close()
        self.actualizar_contador()
        self.generar_estadisticas()

    def actualizar_contador(self): 
        self.lbl_contador.configure(text=str(len(self.tabla.get_children())))


    def insertar_fila_segura(self, datos_lista, tags):
        from datetime import datetime
        anio_actual = datetime.now().year
        hoy = datetime.now().date()
        
        # --- DETECTOR GENERAL DE CHECKOUT (APLICA A TODO EL SISTEMA) ---
        fecha_out_str = str(datos_lista[7]).strip() # Columna CHECK OUT
        if fecha_out_str and "XXXX" not in fecha_out_str.upper() and "NONE" not in fecha_out_str.upper():
            try:
                fecha_out_dt = datetime.strptime(f"{fecha_out_str} {anio_actual}", "%b %d %Y").date()
                if fecha_out_dt <= hoy and "CANCELADO" not in [str(t).upper() for t in tags]:
                    # Si ya pasó la fecha, le inyectamos la puerta antes de mostrarlo
                    if "🚪" not in fecha_out_str:
                        datos_lista[7] = f"{fecha_out_str} 🚪"
            except Exception:
                pass
                
        self.tabla.insert("", "end", values=datos_lista, tags=tuple(tags))


    def exportar_excel(self):
        try:
            data = []
            # Captura de forma automática "RESERVATION #"
            columnas_tabla = [self.tabla.heading(col)["text"] for col in self.tabla["columns"]]
            
            for child in self.tabla.get_children(): 
                valores_fila = list(self.tabla.item(child)["values"])
                
                # --- LIMPIEZA DE EMOJIS PARA EL REPORTE CORPORATIVO ---
                if len(valores_fila) > 7 and "🚪" in str(valores_fila[7]):
                    valores_fila[7] = str(valores_fila[7]).replace(" 🚪", "").strip()
                    
                data.append(valores_fila)
                
            if not data:
                messagebox.showwarning("Atención", "No hay datos para exportar.")
                return
                
            df_filtrado = pd.DataFrame(data, columns=columnas_tabla)
            
            def categorizar(fila):
                ird, info = str(fila.get('IRD', '')).upper(), str(fila.get('INFORMATION', '')).upper()
                
                if any(x in ird or x in info for x in ['TEAM MEMBER', 'TEAMMEMBER', 'EMPLEADO']): return 'TEAM MEMBER'
                elif any(x in ird or x in info for x in ['BIRTHDAY', 'CUMPLEA']): return 'CUMPLEAÑOS'
                elif 'VIP' in ird or 'VIP' in info: return 'VIP'
                elif 'HONEYMOON' in ird or 'LUNA DE MIEL' in ird: return 'HONEYMOON'
                elif 'ANIVERSARY' in ird or 'ANIVERSARIO' in ird: return 'ANIVERSARY'
                elif 'BABYMOON' in ird or 'BABY MOON' in info: return 'BABYMOON'
                else: return 'GENERAL'
                
            df_filtrado['CATEGORIA'] = df_filtrado.apply(categorizar, axis=1)
            home = os.path.expanduser("~")
            posibles = [os.path.join(home, "OneDrive", "Desktop"), os.path.join(home, "Desktop"), os.getcwd()]
            ruta_escritorio = next((r for r in posibles if os.path.exists(r)), os.getcwd())
            ruta_final = os.path.join(ruta_escritorio, f"Reporte_{datetime.now().strftime('%d_%b_%H%M')}.xlsx")
            
            with pd.ExcelWriter(ruta_final, engine='xlsxwriter') as writer:
                workbook = writer.book; worksheet = workbook.add_worksheet('Arrivals')
                fmt_header = workbook.add_format({'bold': True, 'bg_color': '#E1BEE7', 'border': 1, 'align': 'center'})
                fmt_azul = workbook.add_format({'bold': True, 'bg_color': '#00B0F0', 'border': 1, 'align': 'center'})
                for col_num, value in enumerate(columnas_tabla): worksheet.write(0, col_num, value, fmt_header)
                
                fila_actual = 1
                categorias_orden = ['TEAM MEMBER', 'CUMPLEAÑOS', 'VIP', 'HONEYMOON', 'ANIVERSARY', 'BABYMOON', 'GENERAL']
                
                for cat in categorias_orden:
                    df_cat = df_filtrado[df_filtrado['CATEGORIA'] == cat]
                    if not df_cat.empty:
                        worksheet.merge_range(fila_actual, 0, fila_actual, len(columnas_tabla)-1, cat, fmt_azul)
                        fila_actual += 1
                        for row in df_cat.drop(columns=['CATEGORIA']).values:
                            for c_idx, val in enumerate(row): worksheet.write(fila_actual, c_idx, val, workbook.add_format({'border': 1}))
                            fila_actual += 1
                        fila_actual += 1
            os.startfile(ruta_final)
        except Exception as e: messagebox.showerror("Error", str(e))

    def abrir_agenda(self):
        VentanaAgenda(self)

    def generar_carta_despedida(self):
        try:
            sel = self.tabla.selection()
            if not sel:
                messagebox.showwarning("Atención", "Por favor, seleccione un huésped de la lista primero.")
                return
            
            datos_huesped = self.tabla.item(sel[0])['values']
            nombre_gset = str(datos_huesped[2]).strip()
            
            # --- CARGA CORRECTA USANDO LA FUNCIÓN DE RECURSOS ---
            ruta_plantilla = obtener_ruta_recurso("plantilla_despedida.docx")
            
            if not os.path.exists(ruta_plantilla):
                messagebox.showerror("Error", f"No se encontró la plantilla en: {ruta_plantilla}")
                return
            
            doc = Document(ruta_plantilla)
            
            # --- LÓGICA DE REEMPLAZO ---
            for parrafo in doc.paragraphs:
                if "{{NAME}}" in parrafo.text:
                    parrafo.text = parrafo.text.replace("{{NAME}}", nombre_gset)
            
            # Guardar en una ruta temporal del sistema
            ruta_salida = os.path.join(os.environ['TEMP'], "Carta_Despedida_Temporal.docx")
            doc.save(ruta_salida)
            os.startfile(ruta_salida)
            
        except Exception as e:
            messagebox.showerror("Error al generar carta", f"Detalle del error: {str(e)}")

    def abrir_nuevo(self): VentanaEdicion(self)
    def abrir_edicion(self, event):
        sel = self.tabla.selection()
        if sel: VentanaEdicion(self, datos_huesped=self.tabla.item(sel[0])['values'])


# =====================================================================
# --- CLASE INTEGRADA: AGENDA CON IMPORTACIÓN MASIVA DESDE EXCEL ---
# =====================================================================

    def calcular_noches_estadia(self, check_in, check_out):
        from datetime import datetime

        try:
            anio_actual = datetime.now().year
            check_in = str(check_in).strip()
            check_out = str(check_out).replace("🚪", "").strip()

            if not check_in or not check_out:
                return None

            fecha_in = datetime.strptime(f"{check_in} {anio_actual}", "%b %d %Y")
            fecha_out = datetime.strptime(f"{check_out} {anio_actual}", "%b %d %Y")

            noches = (fecha_out - fecha_in).days
            if noches < 0:
                noches += 365

            return noches
        except Exception:
            return None

    def mostrar_noches_hover(self, event):
        # 1. Identificar la fila
        row_id = self.tabla.identify_row(event.y)
        if not row_id:
            self.ocultar_noches_hover()
            return

        # 2. Si es la misma fila, no hacemos nada (mantenemos el tooltip donde está)
        if hasattr(self, '_tooltip_row_actual') and self._tooltip_row_actual == row_id:
            return

        # 3. Calcular noches
        valores = self.tabla.item(row_id)["values"]
        if not valores or len(valores) < 8:
            self.ocultar_noches_hover()
            return

        noches = self.calcular_noches_estadia(valores[6], valores[7])
        if noches is None:
            self.ocultar_noches_hover()
            return

        # 4. Actualizar el contenido y mostrar la ventana fija
        self._tooltip_row_actual = row_id
        self.lbl_tooltip.configure(text=f"🛏️ {noches} {'noche' if noches == 1 else 'noches'}")
        
        # Posicionar la ventana fija
        self.tooltip_noches.geometry(f"+{event.x_root + 15}+{event.y_root + 10}")
        self.tooltip_noches.deiconify() # Esto hace visible la ventana (el opuesto de withdraw)
        self.tooltip_noches.attributes("-topmost", True)

    def ocultar_noches_hover(self, event=None):
        # Solo ocultamos la ventana, no la destruimos
        self.tooltip_noches.withdraw()
        self._tooltip_row_actual = None

    def _destruir_tooltip(self):
        if self.tooltip_noches:
            self.tooltip_noches.destroy()
            self.tooltip_noches = None
        self._tooltip_row_actual = None


class VentanaAgenda(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Directorio Telefónico Interno - Waldorf Astoria")
        self.geometry("900x650")
        self.configure(fg_color="#1a1a1a")
        
        self.after(200, self.lift)
        self.grab_set()
        
        # --- BARRA SUPERIOR DE BÚSQUEDA Y EDICIÓN (5 CAMPOS HORIZONTALES) ---
        self.frame_busqueda = ctk.CTkFrame(self, fg_color="#121212", height=130)
        self.frame_busqueda.pack(fill="x", padx=10, pady=10)
        
        # ============================================================
        # FILA 0: ENTRADAS DE TEXTO PARA EDICIÓN Y CONTROL DE 5 CAMPOS
        # ============================================================
        # 1. Name (Filtra al escribir)
        ctk.CTkLabel(self.frame_busqueda, text="Name:", font=("Roboto", 11, "bold"), text_color="white").grid(row=0, column=0, padx=(10, 2), pady=10, sticky="w")
        self.txt_busca_nombre = ctk.CTkEntry(self.frame_busqueda, width=150, fg_color="#2c3e50", text_color="white")
        self.txt_busca_nombre.grid(row=0, column=1, padx=4, pady=10, sticky="w")
        self.txt_busca_nombre.bind("<KeyRelease>", self.filtrar_datos)
        
        # 2. Position
        ctk.CTkLabel(self.frame_busqueda, text="Position:", font=("Roboto", 11, "bold"), text_color="white").grid(row=0, column=2, padx=(10, 2), pady=10, sticky="w")
        self.txt_busca_puesto = ctk.CTkEntry(self.frame_busqueda, width=140, fg_color="#2c3e50", text_color="white")
        self.txt_busca_puesto.grid(row=0, column=3, padx=4, pady=10, sticky="w")
        
        # 3. Department (Filtra al escribir)
        ctk.CTkLabel(self.frame_busqueda, text="Dept:", font=("Roboto", 11, "bold"), text_color="white").grid(row=0, column=4, padx=(10, 2), pady=10, sticky="w")
        self.txt_busca_depto = ctk.CTkEntry(self.frame_busqueda, width=120, fg_color="#2c3e50", text_color="white")
        self.txt_busca_depto.grid(row=0, column=5, padx=4, pady=10, sticky="w")
        self.txt_busca_depto.bind("<KeyRelease>", self.filtrar_datos)
        
        # 4. Ext
        ctk.CTkLabel(self.frame_busqueda, text="Ext:", font=("Roboto", 11, "bold"), text_color="white").grid(row=0, column=6, padx=(10, 2), pady=10, sticky="w")
        self.txt_busca_ext = ctk.CTkEntry(self.frame_busqueda, width=60, fg_color="#2c3e50", text_color="white")
        self.txt_busca_ext.grid(row=0, column=7, padx=4, pady=10, sticky="w")
        
        # 5. Email
        ctk.CTkLabel(self.frame_busqueda, text="Email:", font=("Roboto", 11, "bold"), text_color="white").grid(row=0, column=8, padx=(10, 2), pady=10, sticky="w")
        self.txt_busca_email = ctk.CTkEntry(self.frame_busqueda, width=160, fg_color="#2c3e50", text_color="white")
        self.txt_busca_email.grid(row=0, column=9, padx=4, pady=10, sticky="w")
        
        # =======================================================
        # --- ETIQUETA DE TOTAL DE REGISTROS (ALINEADO A LA DERECHA) ---
        # =======================================================
        # Label de texto descriptivo
        self.lbl_total_texto = ctk.CTkLabel(
            self.frame_busqueda, 
            text="TOTAL RECORDS:", 
            font=("Roboto", 11, "bold"), 
            text_color="white"
        )
        self.lbl_total_texto.grid(row=0, column=10, padx=(30, 5), pady=10, sticky="w")
        
        # Número dinámico estilo Aqua/Turquesa
        self.lbl_total_num = ctk.CTkLabel(
            self.frame_busqueda, 
            text="0", 
            font=("Roboto", 18, "bold"), 
            text_color="#00F0FF"
        )
        self.lbl_total_num.grid(row=0, column=11, padx=(5, 10), pady=10, sticky="w")

        # ==========================================
        # FILA 1: BOTONERA COMPLETA CON IMPORT EXCEL
        # ==========================================
        frame_botones_agenda = ctk.CTkFrame(self.frame_busqueda, fg_color="transparent")
        frame_botones_agenda.grid(row=1, column=0, columnspan=10, padx=15, pady=(5, 12), sticky="w")
        
        self.btn_add = ctk.CTkButton(frame_botones_agenda, text="ADD", fg_color="red", text_color="black", font=("Roboto", 11, "bold"), width=85, command=self.agregar_contacto)
        self.btn_add.pack(side="left", padx=4)
        
        self.btn_delete = ctk.CTkButton(frame_botones_agenda, text="Delete", fg_color="violet", text_color="black", font=("Roboto", 11, "bold"), width=85, command=self.eliminar_contacto)
        self.btn_delete.pack(side="left", padx=4)
        
        self.btn_update = ctk.CTkButton(frame_botones_agenda, text="UPDATE", fg_color="#f1c40f", text_color="black", font=("Roboto", 11, "bold"), width=85, command=self.actualizar_contacto_existente)
        self.btn_update.pack(side="left", padx=4)
        
        self.btn_reset = ctk.CTkButton(frame_botones_agenda, text="RESET", fg_color="#d35400", text_color="white", font=("Roboto", 11, "bold"), width=85, command=self.reset_busqueda)
        self.btn_reset.pack(side="left", padx=4)
        
        self.btn_pdf = ctk.CTkButton(frame_botones_agenda, text="PDF", fg_color="cyan", text_color="black", font=("Roboto", 11, "bold"), width=85, command=self.exportar_pdf)
        self.btn_pdf.pack(side="left", padx=4)

        # Botón de Importación Masiva
        self.btn_excel = ctk.CTkButton(frame_botones_agenda, text="IMPORT EXCEL", fg_color="#27ae60", text_color="white", font=("Roboto", 11, "bold"), width=110, command=self.importar_excel)
        self.btn_excel.pack(side="left", padx=15)

        # --- BANNER LOGO ---
        self.frame_banner = ctk.CTkFrame(self, fg_color="#0a1128", height=80)
        self.frame_banner.pack(fill="x", padx=10, pady=5)
        
        lbl_logo = ctk.CTkLabel(self.frame_banner, text="WALDORF ASTORIA\nCOSTA RICA • PUNTA CACIQUE", font=("Times New Roman", 18, "bold"), text_color="#f1c40f")
        lbl_logo.pack(pady=10)

        # --- TABLA DE CONTACTOS (TREEVIEW) ---
        style = ttk.Style()
        style.theme_use("default")
        
        style.configure("Agenda.Treeview", background="black", fieldbackground="black", foreground="#00ff00", rowheight=26, font=("Calibri", 11), borderwidth=0, highlightthickness=0)
        style.configure("Agenda.Treeview.Heading", background="#1e272c", foreground="white", font=("Calibri", 11, "bold"), relief="flat", borderwidth=0)
        style.map("Agenda.Treeview.Heading", background=[('active', '#1e272c'), ('pressed', '#11181c')], foreground=[('active', 'white')])
        style.map("Agenda.Treeview", background=[("selected", "#2980b9")], foreground=[("selected", "white")])
        
        self.tabla_agenda = ttk.Treeview(self, columns=("id", "name", "position", "department", "ext", "email"), show="headings", style="Agenda.Treeview")
        
        self.tabla_agenda.heading("id", text="ID")
        self.tabla_agenda.heading("name", text="NAME")
        self.tabla_agenda.heading("position", text="POSITION")
        self.tabla_agenda.heading("department", text="DEPARTMENT")
        self.tabla_agenda.heading("ext", text="EXT.")
        self.tabla_agenda.heading("email", text="EMAIL")
        
        self.tabla_agenda.column("id", width=40, anchor="center")
        self.tabla_agenda.column("name", width=200, anchor="w")
        self.tabla_agenda.column("position", width=180, anchor="w")
        self.tabla_agenda.column("department", width=130, anchor="center")
        self.tabla_agenda.column("ext", width=70, anchor="center")
        self.tabla_agenda.column("email", width=230, anchor="w")
        
        self.tabla_agenda.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.cargar_datos()
        self.tabla_agenda.bind("<Double-1>", self.cargar_contacto_en_campos)
        self.id_contacto_seleccionado = None

    def cargar_datos(self, query_filtro="SELECT * FROM agenda_interna", parametros=()):
        # Limpiamos la tabla de manera eficiente
        for item in self.tabla_agenda.get_children():
            self.tabla_agenda.delete(item)
            
        try:
            conexion = sqlite3.connect("recepcion_final.db")
            cursor = conexion.cursor()
            
            query_limpio = query_filtro.strip()
            if "ORDER BY" not in query_limpio.upper():
                if query_limpio.endswith(";"):
                    query_limpio = query_limpio[:-1]
                query_final = f"{query_limpio} ORDER BY departamento ASC, nombre ASC"
            else:
                query_final = query_limpio

            cursor.execute(query_final, parametros)
            filas = cursor.fetchall()
            
            for fila in filas:
                self.tabla_agenda.insert("", "end", values=fila)
                
            conexion.close()
            
            # === AQUÍ LE DAMOS VIDA AL CONTADOR ===
            # Contamos cuántas filas devolvió la consulta de la base de datos
            total_records = len(filas)
            self.lbl_total_num.configure(text=str(total_records))
            
        except Exception as e:
            messagebox.showerror("Error de Base de Datos", f"No se pudieron cargar los datos:\n{str(e)}")

    def filtrar_datos(self, event=None):
        nombre = self.txt_busca_nombre.get().strip()
        depto = self.txt_busca_depto.get().strip()
        query = "SELECT * FROM agenda_interna WHERE nombre LIKE ? AND departamento LIKE ?"
        self.cargar_datos(query, (f"%{nombre}%", f"%{depto}%"))

    def reset_busqueda(self):
        self.txt_busca_nombre.delete(0, 'end')
        self.txt_busca_puesto.delete(0, 'end')
        self.txt_busca_depto.delete(0, 'end')
        self.txt_busca_ext.delete(0, 'end')
        self.txt_busca_email.delete(0, 'end')
        self.cargar_datos()

    def agregar_contacto(self):
        sub_ventana = ctk.CTkToplevel(self)
        sub_ventana.title("Añadir Nuevo Contacto")
        sub_ventana.geometry("400x350")
        sub_ventana.configure(fg_color="#1a1a1a")
        sub_ventana.grab_set()
        
        campos = ["Nombre:", "Puesto:", "Departamento:", "Extensión:", "Email:"]
        entries = []
        for i, campo in enumerate(campos):
            ctk.CTkLabel(sub_ventana, text=campo, text_color="white").grid(row=i, column=0, padx=10, pady=10, sticky="e")
            entry = ctk.CTkEntry(sub_ventana, width=220, fg_color="#2c3e50")
            entry.grid(row=i, column=1, padx=10, pady=10)
            entries.append(entry)
            
        def guardar():
            datos = [e.get().strip() for e in entries]
            if not datos[0] or not datos[3]:
                messagebox.showwarning("Atención", "Nombre y Extensión son obligatorios.")
                return
            conexion = sqlite3.connect("recepcion_final.db")
            cursor = conexion.cursor()
            cursor.execute("INSERT INTO agenda_interna (nombre, puesto, departamento, extension, email) VALUES (?, ?, ?, ?, ?)", datos)
            conexion.commit()
            conexion.close()
            messagebox.showinfo("Éxito", "Contacto guardado correctamente.")
            sub_ventana.destroy()
            self.cargar_datos()
            
        ctk.CTkButton(sub_ventana, text="GUARDAR", fg_color="#2ecc71", text_color="black", command=guardar).grid(row=5, column=0, columnspan=2, pady=20)

    def eliminar_contacto(self):
        sel = self.tabla_agenda.selection()
        if not sel:
            messagebox.showwarning("Atención", "Seleccione un contacto para eliminar.")
            return
        id_contacto = self.tabla_agenda.item(sel[0])['values'][0]
        nombre_contacto = self.tabla_agenda.item(sel[0])['values'][1]
        if messagebox.askyesno("Confirmar", f"¿Está seguro de que desea eliminar a {nombre_contacto}?"):
            conexion = sqlite3.connect("recepcion_final.db")
            cursor = conexion.cursor()
            cursor.execute("DELETE FROM agenda_interna WHERE id = ?", (id_contacto,))
            conexion.commit()
            conexion.close()
            self.cargar_datos()

    def exportar_pdf(self):
        from fpdf import FPDF
        from tkinter import filedialog, messagebox
        import subprocess
        import os
        import platform

        # 1. Abrir diálogo para elegir dónde guardar
        ruta_guardado = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("Archivos PDF", "*.pdf")],
            initialfile="Directorio_Interno.pdf",
            title="Guardar Directorio Interno"
        )

        # Si el usuario cancela, no hacemos nada
        if not ruta_guardado:
            return

        try:
            # 2. Crear el PDF
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", 'B', 16)
            pdf.cell(200, 10, txt="WALDORF ASTORIA COSTA RICA", ln=True, align='C')
            pdf.set_font("Arial", size=10)
            pdf.ln(10)
            
            for item in self.tabla_agenda.get_children():
                v = self.tabla_agenda.item(item)['values']
                # Aseguramos que los datos sean strings
                linea = f"{str(v[1])} | {str(v[2])} | DEP: {str(v[3])} | EXT: {str(v[4])}"
                pdf.cell(200, 8, txt=linea, ln=True)

            pdf.output(ruta_guardado)
            
            # 3. Bloque para abrir el archivo automáticamente
            try:
                if platform.system() == 'Windows':
                    # Esto es lo más compatible en Windows para abrir cualquier archivo
                    os.startfile(ruta_guardado)
                else:
                    # Por si acaso usas otra cosa (macOS o Linux)
                    subprocess.call(['open', ruta_guardado])
                
                messagebox.showinfo("Éxito", "El PDF se ha generado correctamente.")
                    
            except Exception as e:
                # Si falla la apertura, avisamos que se guardó pero no se pudo abrir
                messagebox.showwarning("Aviso", f"El PDF se guardó correctamente, pero hubo un problema al intentar abrirlo automáticamente:\n{str(e)}")
                
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo generar el PDF:\n{str(e)}")

    
    def cargar_contacto_en_campos(self, event):
        sel = self.tabla_agenda.selection()
        if not sel: 
            return
        valores = self.tabla_agenda.item(sel[0])['values']
        self.id_contacto_seleccionado = valores[0]
        
        # Limpiar e insertar en cada una de las 5 cajas superiores
        self.txt_busca_nombre.delete(0, 'end')
        self.txt_busca_nombre.insert(0, str(valores[1]))
        
        self.txt_busca_puesto.delete(0, 'end')
        self.txt_busca_puesto.insert(0, str(valores[2]))
        
        self.txt_busca_depto.delete(0, 'end')
        self.txt_busca_depto.insert(0, str(valores[3]))
        
        self.txt_busca_ext.delete(0, 'end')
        self.txt_busca_ext.insert(0, str(valores[4]))
        
        self.txt_busca_email.delete(0, 'end')
        self.txt_busca_email.insert(0, str(valores[5]))

    def importar_excel(self):
        ruta_archivo = filedialog.askopenfilename(
            title="Seleccionar Lista de Contactos",
            filetypes=[("Archivos de Excel", "*.xlsx *.xls")]
        )
        if not ruta_archivo:
            return
            
        try:
            df = pd.read_excel(ruta_archivo)
            df.columns = [str(c).strip().lower() for c in df.columns]
            
            col_nombre = None
            col_puesto = None
            col_depto = None
            col_ext = None
            col_email = None
            
            for col in df.columns:
                if col in ['nombre', 'name']:
                    col_nombre = col
                elif col in ['puesto', 'position']:
                    col_puesto = col
                elif col in ['departamento', 'department', 'dept']:
                    col_depto = col
                elif col in ['extension', 'extensio', 'ext', 'ext.']:
                    col_ext = col
                elif col in ['email', 'correo', 'mail']:
                    col_email = col

            if not all([col_nombre, col_puesto, col_depto, col_ext, col_email]):
                messagebox.showerror("Columnas Inválidas", 
                                     "El archivo de Excel debe contener las columnas de:\n"
                                     "Name, Position, Department, Extension, Email")
                return
                
            conexion = sqlite3.connect("recepcion_final.db")
            cursor = conexion.cursor()
            
            contador_actualizados = 0
            contador_nuevos = 0
            
            for _, fila in df.iterrows():
                nombre = str(fila[col_nombre]).strip()
                puesto = str(fila[col_puesto]).strip() if pd.notna(fila[col_puesto]) else ""
                depto = str(fila[col_depto]).strip() if pd.notna(fila[col_depto]) else ""
                
                ext_raw = str(fila[col_ext]).strip() if pd.notna(fila[col_ext]) else ""
                ext = ext_raw.split('.')[0] if '.' in ext_raw else ext_raw
                
                email = str(fila[col_email]).strip() if pd.notna(fila[col_email]) else ""
                
                if not nombre or nombre == "nan" or nombre.replace('-', '').strip() == "":
                    continue
                    
                cursor.execute("SELECT id FROM agenda_interna WHERE nombre = ?", (nombre,))
                existe = cursor.fetchone()
                
                if existe:
                    cursor.execute("""
                        UPDATE agenda_interna 
                        SET puesto = ?, departamento = ?, extension = ?, email = ?
                        WHERE nombre = ?
                    """, (puesto, depto, ext, email, nombre))
                    contador_actualizados += 1
                else:
                    cursor.execute("""
                        INSERT INTO agenda_interna (nombre, puesto, departamento, extension, email) 
                        VALUES (?, ?, ?, ?, ?)
                    """, (nombre, puesto, depto, ext, email))
                    contador_nuevos += 1
            
            conexion.commit()
            conexion.close()
            
            messagebox.showinfo("Importación Exitosa", 
                                f"¡Proceso completado con éxito!\n\n"
                                f"🔄 Contactos actualizados: {contador_actualizados}\n"
                                f"🆕 Contactos nuevos agregados: {contador_nuevos}")
            
            self.cargar_datos()
            
        except Exception as e:
            messagebox.showerror("Error de Importación", f"No se pudo procesar el archivo Excel:\n{str(e)}")

    def actualizar_contacto_existente(self):
        """ Guarda los cambios realizados en las 5 cajas de texto superiores en la base de datos """
        if self.id_contacto_seleccionado is None:
            messagebox.showwarning("Atención", "Por favor, haz doble clic en un contacto de la tabla para actualizar.")
            return
            
        nuevo_nombre = self.txt_busca_nombre.get().strip()
        nuevo_puesto = self.txt_busca_puesto.get().strip()
        nuevo_depto = self.txt_busca_depto.get().strip()
        nuevo_ext = self.txt_busca_ext.get().strip()
        nuevo_email = self.txt_busca_email.get().strip()
        
        if not nuevo_nombre:
            messagebox.showwarning("Atención", "El nombre no puede quedar vacío.")
            return
            
        try:
            conexion = sqlite3.connect("recepcion_final.db")
            cursor = conexion.cursor()
            
            # Actualizamos de forma absoluta los 5 campos usando el ID interno guardado
            cursor.execute("""
                UPDATE agenda_interna 
                SET nombre = ?, puesto = ?, departamento = ?, extension = ?, email = ?
                WHERE id = ?
            """, (nuevo_nombre, nuevo_puesto, nuevo_depto, nuevo_ext, nuevo_email, self.id_contacto_seleccionado))
            
            conexion.commit()
            conexion.close()
            
            messagebox.showinfo("Éxito", "Contacto actualizado correctamente.")
            self.cargar_datos()
            self.reset_busqueda()
            
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo actualizar el contacto:\n{str(e)}")


# =====================================================================
if __name__ == "__main__":
    # 1. Instanciamos la app pero la ocultamos de inmediato
    app = AppHuespedes()
    app.withdraw()

    # 2. Definimos qué pasa cuando el splash termine
    def mostrar_main():
        app.deiconify()
        app.update()
        app.attributes('-topmost', True)
        app.after(100, lambda: app.attributes('-topmost', False))

    # 3. Lanzamos el Splash. 
    # IMPORTANTE: Asegúrate de que dentro de tu clase SplashScreen 
    # estés usando 'obtener_ruta_recurso' para cargar la imagen.
    splash = SplashScreen(mostrar_main)
    
    # 4. El loop principal es el de la app, el splash lo maneja internamente
    app.mainloop()
