"""
Sistema de inicio de sesion + Dashboard de las 20 problematicas
mas influyentes de Mexico, con datos reales (referenciados) y
mapa interactivo de la Republica Mexicana.

Requisitos:
    pip install openpyxl matplotlib

Al iniciar sesion correctamente se abre un Dashboard donde puedes
seleccionar una problematica, ver datos reales y, con un boton,
visualizar en el mapa de la Republica las entidades donde es mas
probable que ocurra.
"""

import tkinter as tk
from tkinter import messagebox
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import hashlib
import os
import json
import gzip
import base64
from datetime import datetime

import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.cm as cm


# ----------------------------------------------------------------------
#  CONFIGURACION GENERAL
# ----------------------------------------------------------------------
EXCEL_FILE = "Prueba registro.xlsx"
SHEET_NAME = "Usuarios"
HEADERS    = ["ID", "Nombre", "Usuario", "Contrasena", "Fecha Registro"]


BG       = "#F7F8FA"
CARD     = "#FFFFFF"
ACCENT   = "#4F46E5"
ACCENT_H = "#4338CA"
TEXT     = "#1E1E2D"
SUBTEXT  = "#6B7280"
BORDER   = "#E5E7EB"
ERROR    = "#EF4444"
SUCCESS  = "#10B981"
FONT_FAM = "Segoe UI"


# ----------------------------------------------------------------------
#  EXCEL: usuarios
# ----------------------------------------------------------------------
def _init_excel():
    """Crea el archivo Excel con encabezados si no existe."""
    if not os.path.exists(EXCEL_FILE):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = SHEET_NAME

        header_fill = PatternFill("solid", fgColor="4F46E5")
        header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
        center      = Alignment(horizontal="center", vertical="center")
        thin        = Side(style="thin", color="D1D5DB")
        border      = Border(left=thin, right=thin, top=thin, bottom=thin)

        col_widths = [8, 22, 20, 42, 22]
        for col_idx, (header, width) in enumerate(zip(HEADERS, col_widths), 1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            cell.font      = header_font
            cell.fill      = header_fill
            cell.alignment = center
            cell.border    = border
            ws.column_dimensions[cell.column_letter].width = width

        ws.row_dimensions[1].height = 28
        wb.save(EXCEL_FILE)


def _hash(password: str) -> str:
    return hashlib.sha256(password.encode()).hexdigest()


def _next_id() -> int:
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb[SHEET_NAME]
    return ws.max_row


def _user_exists(usuario: str) -> bool:
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb[SHEET_NAME]
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[2] and row[2].strip().lower() == usuario.strip().lower():
            return True
    return False


def _register_user(nombre: str, usuario: str, password: str):
    wb   = openpyxl.load_workbook(EXCEL_FILE)
    ws   = wb[SHEET_NAME]
    uid  = _next_id()
    fecha = datetime.now().strftime("%d/%m/%Y %H:%M")

    thin   = Side(style="thin", color="E5E7EB")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")

    row_data = [uid, nombre.strip(), usuario.strip(), _hash(password), fecha]
    new_row  = ws.max_row + 1

    for col_idx, value in enumerate(row_data, 1):
        cell = ws.cell(row=new_row, column=col_idx, value=value)
        cell.border    = border
        cell.alignment = center
        cell.font      = Font(name="Arial", size=10)
        if new_row % 2 == 0:
            cell.fill = PatternFill("solid", fgColor="F3F4F6")

    ws.row_dimensions[new_row].height = 22
    wb.save(EXCEL_FILE)


def _validate_login(usuario: str, password: str):
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb[SHEET_NAME]
    hashed = _hash(password)
    for row in ws.iter_rows(min_row=2, values_only=True):
        if (row[2] and row[2].strip().lower() == usuario.strip().lower()
                and row[3] == hashed):
            return row[1]
    return None


# ----------------------------------------------------------------------
#  WIDGETS REUTILIZABLES
# ----------------------------------------------------------------------
def _entry(parent, placeholder: str, show: str = "") -> tk.Entry:
    frame = tk.Frame(parent, bg=CARD, highlightbackground=BORDER,
                     highlightthickness=1, bd=0)
    frame.pack(fill="x", pady=6)

    e = tk.Entry(frame, font=(FONT_FAM, 11), bg=CARD, fg=SUBTEXT,
                 bd=0, relief="flat", show=show,
                 insertbackground=TEXT, highlightthickness=0)
    e.pack(padx=14, pady=10, fill="x")
    e.insert(0, placeholder)

    def on_focus_in(_):
        if e.get() == placeholder:
            e.delete(0, "end")
            e.config(fg=TEXT, show=show)
        frame.config(highlightbackground=ACCENT)

    def on_focus_out(_):
        if e.get() == "":
            e.insert(0, placeholder)
            e.config(fg=SUBTEXT, show="")
        frame.config(highlightbackground=BORDER)

    e.bind("<FocusIn>",  on_focus_in)
    e.bind("<FocusOut>", on_focus_out)
    e._placeholder = placeholder
    return e


def _get_val(entry: tk.Entry) -> str:
    val = entry.get()
    return "" if val == entry._placeholder else val


def _btn(parent, text: str, command, bg=ACCENT, fg="white") -> tk.Button:
    b = tk.Button(parent, text=text, command=command,
                  font=(FONT_FAM, 11, "bold"),
                  bg=bg, fg=fg, activebackground=ACCENT_H,
                  activeforeground="white", bd=0, relief="flat",
                  cursor="hand2", pady=11)
    b.pack(fill="x", pady=(10, 4))

    def on_enter(_): b.config(bg=ACCENT_H if bg == ACCENT else "#E5E7EB")
    def on_leave(_): b.config(bg=bg)
    b.bind("<Enter>", on_enter)
    b.bind("<Leave>", on_leave)
    return b


def _label(parent, text: str, size=11, color=TEXT, bold=False, pady=0):
    tk.Label(parent, text=text,
             font=(FONT_FAM, size, "bold" if bold else "normal"),
             bg=CARD, fg=color).pack(pady=pady)


def _card(root, width=380) -> tk.Frame:
    wrap = tk.Frame(root, bg=BG)
    wrap.pack(expand=True, fill="both")
    card = tk.Frame(wrap, bg=CARD, bd=0,
                    highlightbackground=BORDER, highlightthickness=1)
    card.place(relx=.5, rely=.5, anchor="center", width=width)
    return card


# ----------------------------------------------------------------------
#  MAPA DE MEXICO (geometria comprimida embebida, division estatal)
# ----------------------------------------------------------------------
MEXICO_GEO_B64 = "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"

def _load_mexico_geo():
    """Descomprime y decodifica el geojson de los estados de Mexico."""
    raw = gzip.decompress(base64.b64decode(MEXICO_GEO_B64))
    data = json.loads(raw)
    geo = {}
    for feat in data["features"]:
        name = feat["properties"]["name"]
        geom = feat["geometry"]
        coords = geom["coordinates"]
        rings = []
        # Nota: la geometria embebida (simplificada) usa siempre una
        # estructura de lista-de-poligonos-de-anillos, tanto para
        # Polygon como para MultiPolygon.
        polys = coords
        for poly in polys:
            for ring in poly:
                rings.append(ring)
        geo[name] = rings
    return geo


# ----------------------------------------------------------------------
#  DATOS: 20 problematicas mas influyentes en Mexico
#  (cifras de referencia segun fuentes publicas: INEGI, ENVIPE, CONEVAL,
#   SESNSP, CONAGUA, ENSANUT, Registro Nacional de Personas Desaparecidas,
#   Transparencia Internacional, Unidad de Politica Migratoria, etc.)
#  Los pesos en "estados" van de 0 a 1 e indican la intensidad relativa
#  del fenomeno en esa entidad para fines de visualizacion en el mapa.
# ----------------------------------------------------------------------
PROBLEMATICAS = [
    {
        "nombre": "Inseguridad publica (percepcion)",
        "categoria": "Seguridad",
        "dato": "75.6% de la poblacion de 18 anos o mas considera insegura su entidad (ENVIPE 2025, INEGI).",
        "descripcion": "La inseguridad es el problema que mas preocupa a los mexicanos segun las encuestas oficiales. "
                       "Morelos, Tabasco y Guanajuato encabezan la percepcion de inseguridad a nivel nacional.",
        "fuente": "INEGI - ENVIPE 2024/2025",
        "estados": {"Morelos": 1.0, "Tabasco": 0.95, "Guanajuato": 0.9, "Zacatecas": 0.85, "México": 0.7},
    },
    {
        "nombre": "Homicidios dolosos",
        "categoria": "Seguridad",
        "dato": "Alrededor de 25,000 homicidios dolosos al ano; Guanajuato concentra de forma recurrente el mayor numero absoluto de carpetas de investigacion (SESNSP).",
        "descripcion": "La violencia homicida se concentra en corredores asociados a la disputa territorial entre "
                       "grupos delictivos: el Bajio, el occidente y la frontera norte.",
        "fuente": "SESNSP - Reportes mensuales de incidencia delictiva",
        "estados": {"Guanajuato": 1.0, "México": 0.9, "Baja California": 0.85, "Jalisco": 0.8, "Chihuahua": 0.75, "Michoacán": 0.7},
    },
    {
        "nombre": "Pobreza",
        "categoria": "Social",
        "dato": "36.3% de la poblacion vive en situacion de pobreza (CONEVAL 2022); Chiapas registra el porcentaje mas alto del pais (alrededor de 67%).",
        "descripcion": "La pobreza se concentra historicamente en el sur-sureste del pais, donde tambien son menores "
                       "los indicadores de acceso a servicios basicos y educacion.",
        "fuente": "CONEVAL - Medicion de la pobreza 2022",
        "estados": {"Chiapas": 1.0, "Guerrero": 0.9, "Oaxaca": 0.85, "Puebla": 0.7, "Veracruz": 0.65},
    },
    {
        "nombre": "Desempleo",
        "categoria": "Economia",
        "dato": "Tasa de desocupacion nacional cercana a 2.7% (INEGI, ENOE); Tabasco y Ciudad de Mexico se ubican entre las entidades con mayores tasas.",
        "descripcion": "Aunque la tasa de desempleo abierto en Mexico es relativamente baja, esconde altos niveles "
                       "de informalidad y subocupacion en varias regiones.",
        "fuente": "INEGI - Encuesta Nacional de Ocupacion y Empleo (ENOE)",
        "estados": {"Tabasco": 1.0, "Ciudad de México": 0.9, "Tamaulipas": 0.8, "Coahuila": 0.6},
    },
    {
        "nombre": "Informalidad laboral",
        "categoria": "Economia",
        "dato": "Cerca de 54.8% de la poblacion ocupada trabaja en la informalidad (INEGI); en Oaxaca y Guerrero esta tasa supera el 80%.",
        "descripcion": "La informalidad implica falta de acceso a seguridad social, prestaciones y proteccion "
                       "laboral, y es mas alta en estados con economias rurales y de comercio informal.",
        "fuente": "INEGI - ENOE, Indicadores de informalidad laboral",
        "estados": {"Oaxaca": 1.0, "Guerrero": 0.95, "Chiapas": 0.9, "Puebla": 0.75, "Hidalgo": 0.65},
    },
    {
        "nombre": "Corrupcion",
        "categoria": "Gobernanza",
        "dato": "Mexico ocupa el lugar 126 de 180 paises en el Indice de Percepcion de la Corrupcion 2023 (Transparencia Internacional), con 31 puntos sobre 100.",
        "descripcion": "La corrupcion afecta tramites, contratacion publica y procuracion de justicia. Las "
                       "entidades mas pobladas concentran el mayor numero absoluto de quejas y denuncias.",
        "fuente": "Transparencia Internacional - IPC 2023 / INEGI - ENCIG",
        "estados": {"México": 1.0, "Ciudad de México": 0.9, "Quintana Roo": 0.8, "Puebla": 0.7, "Veracruz": 0.65},
    },
    {
        "nombre": "Escasez de agua",
        "categoria": "Medio ambiente",
        "dato": "Mas del 60% del territorio nacional presenta algun grado de sequia (CONAGUA, Monitor de Sequia de Mexico 2024).",
        "descripcion": "El estres hidrico es especialmente critico en el norte y noreste del pais, donde la "
                       "demanda agricola, industrial y urbana supera la disponibilidad de agua.",
        "fuente": "CONAGUA - Monitor de Sequia de Mexico",
        "estados": {"Baja California": 1.0, "Sonora": 0.95, "Coahuila": 0.9, "Nuevo León": 0.85, "Chihuahua": 0.8, "Tamaulipas": 0.6},
    },
    {
        "nombre": "Feminicidios y violencia de genero",
        "categoria": "Seguridad",
        "dato": "Alrededor de 850 feminicidios registrados en 2023 (SESNSP); el Estado de Mexico concentra el mayor numero de casos del pais.",
        "descripcion": "La violencia contra las mujeres incluye feminicidios, violencia familiar y desaparicion. "
                       "El Estado de Mexico, Veracruz, Nuevo Leon y Jalisco registran las cifras mas altas.",
        "fuente": "SESNSP - Incidencia delictiva del fuero comun",
        "estados": {"México": 1.0, "Veracruz": 0.85, "Nuevo León": 0.8, "Jalisco": 0.75, "Ciudad de México": 0.65},
    },
    {
        "nombre": "Extorsion",
        "categoria": "Seguridad",
        "dato": "La extorsion es uno de los delitos con mayor crecimiento en denuncias segun ENVIPE; Estado de Mexico, Puebla y Ciudad de Mexico encabezan las carpetas de investigacion (SESNSP).",
        "descripcion": "Afecta principalmente a comerciantes y transportistas mediante cobro de piso y llamadas "
                       "de extorsion telefonica, frecuentemente desde centros penitenciarios.",
        "fuente": "SESNSP - Incidencia delictiva del fuero comun",
        "estados": {"México": 1.0, "Puebla": 0.85, "Ciudad de México": 0.8, "Guanajuato": 0.75, "Veracruz": 0.6},
    },
    {
        "nombre": "Migracion y transito de personas migrantes",
        "categoria": "Social",
        "dato": "Mas de un millon de eventos de personas migrantes en situacion irregular fueron registrados en un ano reciente (Unidad de Politica Migratoria, Registro e Identidad de Personas).",
        "descripcion": "Chiapas y Tabasco son los principales puntos de entrada por la frontera sur, mientras "
                       "Tamaulipas y Sonora concentran el cruce hacia Estados Unidos.",
        "fuente": "Secretaria de Gobernacion - Unidad de Politica Migratoria",
        "estados": {"Chiapas": 1.0, "Tabasco": 0.85, "Tamaulipas": 0.8, "Sonora": 0.75, "Baja California": 0.6},
    },
    {
        "nombre": "Obesidad y sobrepeso",
        "categoria": "Salud",
        "dato": "Cerca de 75% de los adultos mexicanos vive con sobrepeso u obesidad (ENSANUT); Yucatan, Quintana Roo y Campeche presentan algunas de las prevalencias mas altas.",
        "descripcion": "La region sureste y la peninsula de Yucatan presentan tasas elevadas de obesidad, asociadas "
                       "a cambios en los patrones alimentarios y baja actividad fisica.",
        "fuente": "ENSANUT - Encuesta Nacional de Salud y Nutricion",
        "estados": {"Yucatán": 1.0, "Quintana Roo": 0.9, "Campeche": 0.85, "Sonora": 0.8, "Tamaulipas": 0.7},
    },
    {
        "nombre": "Diabetes mellitus",
        "categoria": "Salud",
        "dato": "La diabetes es una de las principales causas de muerte en Mexico (INEGI, estadisticas de mortalidad); Yucatan, Ciudad de Mexico y Nuevo Leon presentan las tasas de mortalidad mas altas.",
        "descripcion": "Esta enfermedad cronica esta fuertemente asociada con la obesidad y el sedentarismo, y "
                       "representa una carga importante para el sistema de salud publica.",
        "fuente": "INEGI - Estadisticas de mortalidad",
        "estados": {"Yucatán": 1.0, "Ciudad de México": 0.9, "Nuevo León": 0.85, "Veracruz": 0.7, "Tamaulipas": 0.65},
    },
    {
        "nombre": "Deforestacion",
        "categoria": "Medio ambiente",
        "dato": "Mexico pierde alrededor de 127 mil hectareas de bosques y selvas al ano (CONAFOR / Global Forest Watch); la Selva Maya en Chiapas, Campeche y Quintana Roo es una de las zonas mas afectadas.",
        "descripcion": "La deforestacion responde principalmente a la expansion de la frontera agropecuaria, "
                       "los incendios forestales y el cambio de uso de suelo.",
        "fuente": "CONAFOR / Global Forest Watch",
        "estados": {"Chiapas": 1.0, "Campeche": 0.9, "Quintana Roo": 0.85, "Veracruz": 0.6, "Oaxaca": 0.55},
    },
    {
        "nombre": "Contaminacion del aire",
        "categoria": "Medio ambiente",
        "dato": "La Zona Metropolitana del Valle de Mexico supera con frecuencia los limites de calidad del aire establecidos por la norma (Sistema de Monitoreo Atmosferico, CDMX/Edomex).",
        "descripcion": "Las grandes zonas metropolitanas, en especial el Valle de Mexico y Mexicali, enfrentan "
                       "episodios recurrentes de mala calidad del aire por trafico vehicular e industria.",
        "fuente": "Sistema de Monitoreo Atmosferico de la CDMX / SEMARNAT",
        "estados": {"Ciudad de México": 1.0, "México": 0.95, "Baja California": 0.7, "Jalisco": 0.55},
    },
    {
        "nombre": "Crimen organizado y narcotrafico",
        "categoria": "Seguridad",
        "dato": "Operativos de seguridad federal documentan presencia y disputa de grupos del crimen organizado en gran parte del territorio (SEDENA / SSPC); Sinaloa, Michoacan, Guanajuato y Jalisco son focos recurrentes de conflicto.",
        "descripcion": "La disputa por rutas de trafico de drogas y el control territorial generan violencia "
                       "armada, desplazamiento forzado y desapariciones en varias regiones.",
        "fuente": "SEDENA / Secretaria de Seguridad y Proteccion Ciudadana",
        "estados": {"Sinaloa": 1.0, "Michoacán": 0.9, "Guanajuato": 0.85, "Jalisco": 0.8, "Tamaulipas": 0.75, "Chihuahua": 0.7},
    },
    {
        "nombre": "Desaparicion de personas",
        "categoria": "Seguridad",
        "dato": "Mas de 100,000 personas se encuentran registradas como desaparecidas o no localizadas (Registro Nacional de Personas Desaparecidas y No Localizadas, CNB); Jalisco encabeza la lista de entidades.",
        "descripcion": "El fenomeno de desaparicion esta fuertemente ligado a la presencia del crimen organizado "
                       "y afecta de manera particular a Jalisco, Tamaulipas, Estado de Mexico y Nuevo Leon.",
        "fuente": "Comision Nacional de Busqueda (CNB) - RNPDNO",
        "estados": {"Jalisco": 1.0, "Tamaulipas": 0.9, "México": 0.85, "Nuevo León": 0.8, "Veracruz": 0.6},
    },
    {
        "nombre": "Robo de vehiculos",
        "categoria": "Seguridad",
        "dato": "El robo de vehiculo con y sin violencia es uno de los delitos de alto impacto mas frecuentes; Estado de Mexico y Guanajuato concentran el mayor numero de carpetas (SESNSP).",
        "descripcion": "Las zonas industriales y los corredores carreteros del centro del pais son los puntos "
                       "donde mas se concentra este delito.",
        "fuente": "SESNSP - Incidencia delictiva del fuero comun",
        "estados": {"México": 1.0, "Guanajuato": 0.9, "Baja California": 0.8, "Puebla": 0.6, "Tamaulipas": 0.55},
    },
    {
        "nombre": "Desercion escolar",
        "categoria": "Educacion",
        "dato": "El abandono escolar es mayor en el nivel medio superior; Chiapas, Oaxaca y Guerrero presentan los rezagos educativos mas altos del pais (SEP / INEE).",
        "descripcion": "La desercion escolar esta vinculada con la pobreza, la dispersion geografica de las "
                       "comunidades y la falta de infraestructura educativa.",
        "fuente": "SEP / Instituto Nacional para la Evaluacion de la Educacion",
        "estados": {"Chiapas": 1.0, "Oaxaca": 0.9, "Guerrero": 0.85, "Michoacán": 0.6, "Veracruz": 0.55},
    },
    {
        "nombre": "Embarazo adolescente",
        "categoria": "Salud",
        "dato": "Mexico tiene una de las tasas de embarazo adolescente mas altas entre los paises de la OCDE; Coahuila, Chihuahua y Guerrero presentan tasas elevadas (ENADID / Secretaria de Salud).",
        "descripcion": "El embarazo en adolescentes esta asociado con desercion escolar, falta de acceso a "
                       "educacion sexual y limitada disponibilidad de metodos anticonceptivos.",
        "fuente": "ENADID - Encuesta Nacional de la Dinamica Demografica",
        "estados": {"Coahuila": 1.0, "Chihuahua": 0.9, "Guerrero": 0.85, "Nayarit": 0.7, "Durango": 0.6},
    },
    {
        "nombre": "Inflacion y aumento de precios",
        "categoria": "Economia",
        "dato": "El aumento de precios fue identificado como el segundo problema mas importante por la poblacion en la ENVIPE 2024, solo despues de la inseguridad (INEGI).",
        "descripcion": "La inflacion afecta de manera mas severa a los hogares de menores ingresos, en especial "
                       "el costo de los alimentos basicos en zonas urbanas y fronterizas.",
        "fuente": "INEGI - ENVIPE 2024 / INPC",
        "estados": {"Ciudad de México": 0.9, "México": 0.9, "Baja California": 0.85, "Quintana Roo": 0.8, "Jalisco": 0.7},
    },
]


# ----------------------------------------------------------------------
#  VENTANA: MAPA DE LA REPUBLICA PARA UNA PROBLEMATICA
# ----------------------------------------------------------------------
class MapWindow:
    def __init__(self, master, problema):
        self.win = tk.Toplevel(master)
        self.win.title(f"Mapa - {problema['nombre']}")
        self.win.geometry("800x720")
        self.win.configure(bg=BG)
        self.win.grab_set()

        tk.Label(self.win, text=problema["nombre"], font=(FONT_FAM, 15, "bold"),
                 bg=BG, fg=TEXT, wraplength=760).pack(pady=(14, 2))
        tk.Label(self.win, text="Intensidad estimada por entidad federativa "
                                 "(zonas en rojo = mayor probabilidad)",
                 font=(FONT_FAM, 10), bg=BG, fg=SUBTEXT).pack()

        fig, ax = plt.subplots(figsize=(7, 5.6))
        geo = _load_mexico_geo()
        estados = problema["estados"]
        cmap = cm.get_cmap("Reds")

        for name, rings in geo.items():
            weight = estados.get(name, 0)
            if weight > 0:
                color = cmap(0.25 + weight * 0.7)
            else:
                color = "#E5E7EB"
            for ring in rings:
                xs = [c[0] for c in ring]
                ys = [c[1] for c in ring]
                ax.fill(xs, ys, facecolor=color, edgecolor="white", linewidth=0.6)

        ax.set_xlim(-119, -86)
        ax.set_ylim(14, 33)
        ax.set_aspect("equal")
        ax.axis("off")
        fig.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=self.win)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True, padx=10, pady=(6, 4))

        ranked = sorted(estados.items(), key=lambda x: -x[1])
        ranking_txt = "Entidades con mayor probabilidad: " + ", ".join(
            f"{n} ({int(w * 100)}%)" for n, w in ranked[:5]
        )
        tk.Label(self.win, text=ranking_txt, font=(FONT_FAM, 9, "bold"),
                 bg=BG, fg=TEXT, wraplength=760, justify="left").pack(pady=(0, 2), padx=14)

        tk.Label(self.win, text="Fuente: " + problema["fuente"], font=(FONT_FAM, 8),
                 bg=BG, fg=SUBTEXT, wraplength=760, justify="left").pack(pady=(0, 8), padx=14)

        tk.Button(self.win, text="Cerrar", command=self.win.destroy,
                  font=(FONT_FAM, 10, "bold"), bg=ACCENT, fg="white", bd=0,
                  relief="flat", cursor="hand2", padx=24, pady=8).pack(pady=(0, 14))


# ----------------------------------------------------------------------
#  VENTANA: DASHBOARD PRINCIPAL
# ----------------------------------------------------------------------
class DashboardWindow:
    def __init__(self, root: tk.Tk, nombre: str):
        self.root = root
        self.nombre = nombre
        self.selected = None

        for w in self.root.winfo_children():
            w.destroy()

        self.root.title("Dashboard - Problematicas de Mexico")
        self.root.resizable(True, True)
        self.root.geometry("1000x640")
        self.root.configure(bg=BG)
        self._build()

    def _build(self):
        # ---- encabezado ----
        header = tk.Frame(self.root, bg=ACCENT, height=64)
        header.pack(fill="x")
        header.pack_propagate(False)

        tk.Label(header, text=f"Dashboard  •  Bienvenido, {self.nombre}",
                 font=(FONT_FAM, 14, "bold"), bg=ACCENT, fg="white").pack(side="left", padx=20)

        tk.Button(header, text="Cerrar sesion", command=self._logout,
                  font=(FONT_FAM, 10, "bold"), bg=ACCENT_H, fg="white",
                  bd=0, relief="flat", cursor="hand2", padx=14, pady=6).pack(side="right", padx=20)

        # ---- cuerpo ----
        body = tk.Frame(self.root, bg=BG)
        body.pack(fill="both", expand=True, padx=16, pady=16)

        # panel izquierdo: lista de problematicas
        left = tk.Frame(body, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        left.pack(side="left", fill="y", padx=(0, 14))

        tk.Label(left, text="Top 20 problematicas en Mexico",
                 font=(FONT_FAM, 12, "bold"), bg=CARD, fg=TEXT).pack(padx=16, pady=(16, 8))

        list_frame = tk.Frame(left, bg=CARD)
        list_frame.pack(padx=16, pady=(0, 16))

        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")

        self.listbox = tk.Listbox(list_frame, font=(FONT_FAM, 10), width=42, height=24,
                                   bd=0, highlightthickness=0, selectbackground=ACCENT,
                                   selectforeground="white", activestyle="none",
                                   yscrollcommand=scrollbar.set)
        for i, p in enumerate(PROBLEMATICAS, 1):
            self.listbox.insert("end", f"{i:>2}. {p['nombre']}")
        self.listbox.pack(side="left", fill="both")
        scrollbar.config(command=self.listbox.yview)
        self.listbox.bind("<<ListboxSelect>>", self._on_select)

        # panel derecho: detalle
        right = tk.Frame(body, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        right.pack(side="left", fill="both", expand=True)

        inner = tk.Frame(right, bg=CARD)
        inner.pack(fill="both", expand=True, padx=28, pady=28)

        self.lbl_titulo = tk.Label(inner, text="Selecciona una problematica de la lista",
                                    font=(FONT_FAM, 17, "bold"), bg=CARD, fg=TEXT,
                                    anchor="w", justify="left", wraplength=520)
        self.lbl_titulo.pack(fill="x")

        self.lbl_categoria = tk.Label(inner, text="", font=(FONT_FAM, 10, "bold"),
                                       bg=CARD, fg=ACCENT, anchor="w")
        self.lbl_categoria.pack(fill="x", pady=(6, 14))

        tk.Label(inner, text="Dato real", font=(FONT_FAM, 9, "bold"),
                 bg=CARD, fg=SUBTEXT, anchor="w").pack(fill="x")
        self.lbl_dato = tk.Label(inner, text="—", font=(FONT_FAM, 12),
                                  bg=CARD, fg=TEXT, anchor="w", justify="left", wraplength=520)
        self.lbl_dato.pack(fill="x", pady=(2, 14))

        tk.Label(inner, text="Contexto", font=(FONT_FAM, 9, "bold"),
                 bg=CARD, fg=SUBTEXT, anchor="w").pack(fill="x")
        self.lbl_desc = tk.Label(inner, text="Da clic en una problematica para ver datos reales y "
                                              "su distribucion geografica en el mapa de la Republica.",
                                  font=(FONT_FAM, 10), bg=CARD, fg=TEXT, anchor="w",
                                  justify="left", wraplength=520)
        self.lbl_desc.pack(fill="x", pady=(2, 14))

        self.lbl_fuente = tk.Label(inner, text="", font=(FONT_FAM, 8),
                                    bg=CARD, fg=SUBTEXT, anchor="w", justify="left", wraplength=520)
        self.lbl_fuente.pack(fill="x", pady=(0, 18))

        self.btn_mapa = tk.Button(inner, text="Ver en el mapa de Mexico", command=self._show_map,
                                   font=(FONT_FAM, 11, "bold"), bg=ACCENT, fg="white", bd=0,
                                   relief="flat", cursor="hand2", padx=22, pady=11,
                                   state="disabled")
        self.btn_mapa.pack(anchor="w")

        def on_enter(_):
            if self.btn_mapa["state"] != "disabled":
                self.btn_mapa.config(bg=ACCENT_H)

        def on_leave(_):
            if self.btn_mapa["state"] != "disabled":
                self.btn_mapa.config(bg=ACCENT)

        self.btn_mapa.bind("<Enter>", on_enter)
        self.btn_mapa.bind("<Leave>", on_leave)

    def _on_select(self, _event):
        sel = self.listbox.curselection()
        if not sel:
            return
        p = PROBLEMATICAS[sel[0]]
        self.selected = p

        self.lbl_titulo.config(text=p["nombre"])
        self.lbl_categoria.config(text="Categoria: " + p["categoria"])
        self.lbl_dato.config(text=p["dato"])
        self.lbl_desc.config(text=p["descripcion"])
        self.lbl_fuente.config(text="Fuente: " + p["fuente"])
        self.btn_mapa.config(state="normal", bg=ACCENT)

    def _show_map(self):
        if self.selected:
            MapWindow(self.root, self.selected)

    def _logout(self):
        self.root.resizable(False, False)
        LoginWindow(self.root)


# ----------------------------------------------------------------------
#  VENTANA: INICIO DE SESION
# ----------------------------------------------------------------------
class LoginWindow:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title("Inicio de sesion")
        root.geometry("460x520")
        root.resizable(False, False)
        root.configure(bg=BG)
        self._build()

    def _build(self):
        for w in self.root.winfo_children():
            w.destroy()

        self.root.title("Inicio de sesion")
        self.root.geometry("460x520")
        self.root.resizable(False, False)

        card = _card(self.root)

        inner = tk.Frame(card, bg=CARD)
        inner.pack(padx=38, pady=40, fill="x")

        tk.Label(inner, text="🔐", font=(FONT_FAM, 32), bg=CARD).pack(pady=(0, 4))
        _label(inner, "Bienvenido de nuevo", size=17, bold=True)
        _label(inner, "Ingresa tus credenciales para continuar",
               size=10, color=SUBTEXT, pady=(2, 18))

        self.e_user = _entry(inner, "Usuario")
        self.e_pass = _entry(inner, "Contrasena", show="•")

        self.msg = tk.Label(inner, text="", font=(FONT_FAM, 9),
                             bg=CARD, fg=ERROR)
        self.msg.pack()

        _btn(inner, "Iniciar sesion", self._login)

        sep = tk.Frame(inner, bg=CARD)
        sep.pack(fill="x", pady=(12, 0))
        tk.Label(sep, text="¿No tienes cuenta?", font=(FONT_FAM, 10),
                 bg=CARD, fg=SUBTEXT).pack(side="left")
        reg_btn = tk.Label(sep, text=" Registrate", font=(FONT_FAM, 10, "bold"),
                            bg=CARD, fg=ACCENT, cursor="hand2")
        reg_btn.pack(side="left")
        reg_btn.bind("<Button-1>", lambda _: RegisterWindow(self.root))

    def _login(self):
        usuario  = _get_val(self.e_user)
        password = _get_val(self.e_pass)

        if not usuario or not password:
            self.msg.config(text="✖  Por favor completa todos los campos.")
            return

        nombre = _validate_login(usuario, password)
        if nombre:
            DashboardWindow(self.root, nombre)
        else:
            self.msg.config(text="✖  Usuario o contrasena incorrectos.")


# ----------------------------------------------------------------------
#  VENTANA: REGISTRO
# ----------------------------------------------------------------------
class RegisterWindow:
    def __init__(self, master: tk.Tk):
        self.win = tk.Toplevel(master)
        self.win.title("Crear cuenta")
        self.win.geometry("460x580")
        self.win.resizable(False, False)
        self.win.configure(bg=BG)
        self.win.grab_set()
        self._build()

    def _build(self):
        card = _card(self.win)
        inner = tk.Frame(card, bg=CARD)
        inner.pack(padx=38, pady=36, fill="x")

        tk.Label(inner, text="👤", font=(FONT_FAM, 32), bg=CARD).pack(pady=(0, 4))
        _label(inner, "Crear cuenta", size=17, bold=True)
        _label(inner, "Completa los datos para registrarte",
               size=10, color=SUBTEXT, pady=(2, 18))

        self.e_nombre  = _entry(inner, "Nombre completo")
        self.e_user    = _entry(inner, "Usuario")
        self.e_pass    = _entry(inner, "Contrasena", show="•")
        self.e_confirm = _entry(inner, "Confirmar contrasena", show="•")

        self.msg = tk.Label(inner, text="", font=(FONT_FAM, 9),
                             bg=CARD, fg=ERROR, wraplength=300)
        self.msg.pack()

        _btn(inner, "Registrarse", self._register)

        lnk = tk.Label(inner, text="← Volver al inicio de sesion",
                        font=(FONT_FAM, 10), bg=CARD, fg=ACCENT, cursor="hand2")
        lnk.pack(pady=(10, 0))
        lnk.bind("<Button-1>", lambda _: self.win.destroy())

    def _register(self):
        nombre   = _get_val(self.e_nombre)
        usuario  = _get_val(self.e_user)
        password = _get_val(self.e_pass)
        confirm  = _get_val(self.e_confirm)

        if not all([nombre, usuario, password, confirm]):
            self.msg.config(text="✖  Por favor completa todos los campos.")
            return
        if len(password) < 6:
            self.msg.config(text="✖  La contrasena debe tener al menos 6 caracteres.")
            return
        if password != confirm:
            self.msg.config(text="✖  Las contrasenas no coinciden.")
            return
        if _user_exists(usuario):
            self.msg.config(text="✖  Ese nombre de usuario ya esta en uso.")
            return

        _register_user(nombre, usuario, password)
        messagebox.showinfo(
            "¡Registro exitoso!",
            f"✅ Cuenta creada para {nombre}.\nYa puedes iniciar sesion.",
            parent=self.win
        )
        self.win.destroy()


# ----------------------------------------------------------------------
#  PUNTO DE ENTRADA
# ----------------------------------------------------------------------
if __name__ == "__main__":
    _init_excel()
    root = tk.Tk()
    root.update_idletasks()
    w, h = 460, 520
    x = (root.winfo_screenwidth()  - w) // 2
    y = (root.winfo_screenheight() - h) // 2
    root.geometry(f"{w}x{h}+{x}+{y}")
    LoginWindow(root)
    root.mainloop()
