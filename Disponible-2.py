import csv
import glob
import io
import math
import numbers
import os
import re
import tkinter as tk
from html.parser import HTMLParser
from tkinter import ttk, messagebox

import pandas as pd

ruta = r"C:\Users\norbertoacbe\Downloads"


# ---------------------------------------------------------------------------
# Lectura de los archivos de PeopleSoft (son tablas HTML con extensión .xls)
# ---------------------------------------------------------------------------
class LectorTablaHTML(HTMLParser):
    """Extrae las filas de la primera tabla de un HTML usando solo la librería estándar."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.filas = []
        self._fila = None
        self._celda = None
        self._profundidad = 0
        self._tablas_vistas = 0

    def _en_primera_tabla(self):
        return self._profundidad == 1 and self._tablas_vistas == 1

    def _cerrar_celda(self):
        if self._celda is not None and self._fila is not None:
            self._fila.append(" ".join("".join(self._celda).split()))
        self._celda = None

    def _cerrar_fila(self):
        self._cerrar_celda()
        if self._fila:
            self.filas.append(self._fila)
        self._fila = None

    def handle_starttag(self, tag, attrs):
        if tag == "table":
            self._profundidad += 1
            if self._profundidad == 1:
                self._tablas_vistas += 1
        elif self._en_primera_tabla():
            if tag == "tr":
                self._cerrar_fila()
                self._fila = []
            elif tag in ("td", "th"):
                self._cerrar_celda()
                if self._fila is None:
                    self._fila = []
                self._celda = []

    def handle_endtag(self, tag):
        if tag == "table":
            if self._en_primera_tabla():
                self._cerrar_fila()
            self._profundidad = max(0, self._profundidad - 1)
        elif self._en_primera_tabla():
            if tag in ("td", "th"):
                self._cerrar_celda()
            elif tag == "tr":
                self._cerrar_fila()

    def handle_data(self, data):
        if self._celda is not None and self._en_primera_tabla():
            self._celda.append(data)


def leer_tabla(archivo):
    """Lee la tabla del archivo. Usa pandas si lxml está instalado; si no, la librería estándar."""
    try:
        return pd.read_html(archivo)[0]
    except ImportError:
        with open(archivo, "rb") as f:
            contenido = f.read()
        try:
            texto = contenido.decode("utf-8")
        except UnicodeDecodeError:
            texto = contenido.decode("cp1252", errors="replace")
        lector = LectorTablaHTML()
        lector.feed(texto)
        lector.close()
        buffer = io.StringIO()
        csv.writer(buffer).writerows(lector.filas)
        buffer.seek(0)
        return pd.read_csv(buffer, thousands=",")


def a_numero(valor):
    """Convierte un monto a número aunque venga como texto.

    Acepta, por ejemplo: 1234567.89 · 1,234,567.89 · 1.234.567,89 · $ 1.234.567
    · (50,000.00) · 50,000.00- · -50.000
    """
    if isinstance(valor, numbers.Number):
        return float(valor)
    texto = str(valor).strip()
    negativo = texto.startswith(("-", "(")) or texto.endswith("-")
    texto = re.sub(r"[^\d,.]", "", texto)  # quita $, espacios, letras, paréntesis y signos
    if not texto:
        return math.nan
    if "," in texto and "." in texto:
        if texto.rfind(",") > texto.rfind("."):   # 1.234.567,89
            texto = texto.replace(".", "").replace(",", ".")
        else:                                     # 1,234,567.89
            texto = texto.replace(",", "")
    elif "," in texto:
        if re.search(r",\d{1,2}$", texto):        # 1234567,89
            texto = texto.replace(",", ".")
        else:                                     # 1,234,567
            texto = texto.replace(",", "")
    elif texto.count(".") > 1 or re.search(r"\.\d{3}$", texto):  # 1.234.567
        texto = texto.replace(".", "")
    try:
        numero = float(texto)
    except ValueError:
        return math.nan
    return -numero if negativo else numero


# La ventana se crea desde ya (oculta) para poder mostrar mensajes
app = tk.Tk()
app.withdraw()

archivos = glob.glob(os.path.join(ruta, "ps*.xls"))

if not archivos:
    messagebox.showerror("Error", "No se encontraron archivos que coincidan con 'ps*.xls'", parent=app)
    app.destroy()
    raise SystemExit

datos = pd.concat([leer_tabla(archivo) for archivo in archivos], ignore_index=True)

# Los montos pueden llegar como texto; se convierten a número antes de sumar
columna_monto = 'Presupuesto Disponible*'
originales = datos[columna_monto]
datos[columna_monto] = originales.apply(a_numero)

no_convertidos = originales[
    datos[columna_monto].isna()
    & originales.notna()
    & (originales.astype(str).str.strip() != "")
]
if not no_convertidos.empty:
    ejemplos = ", ".join(repr(v) for v in no_convertidos.unique()[:5])
    messagebox.showwarning(
        "Valores no reconocidos",
        f"{len(no_convertidos)} valor(es) de '{columna_monto}' no se pudieron leer como número "
        f"y quedaron fuera de la tabla.\n\nEjemplos: {ejemplos}",
        parent=app,
    )

departamentos = {
    247: "PROG DE ADMON MCDEO Y LOG INT",
    332: "PROG DE ECONOMIA Y FINAN INTER",
    434: "PROGRAMA DE GASTRONOMIA",
    5: "PROG DE ADMON DE EMPRESAS",
    7: "PROG DE ADMON DE NEGOC INTERN",
    6: "PROG DE ADMON Y SERVICIO",
    814: "PROG ECONO FINAN INTER VIRTUAL",
    1372: "MAESTRIA EN GERENCIA INVERSION",
    436: "MAE EN GERENCIA DE OPERACIONES",
    542: "MAE EN GERENCIA ESTRATEGICA",
    1438: "MAE EN GERENCIA INTERN",
    519: "MAE EN GERENCIA INNOVACION",
    520: "DOCTORADO EN ADMINISTRACION",
    780: "MAE EN GERENCIA DE INVER VIRTU",
    784: "MAE EN GERENCIA NEGOCI HOSPITA",
    268: "EDUCONTINUA EICEA",
    1441: "INTERNACIONALIZACION EN CASA",
    454: "PROYECTOS CON EMPRESAS",
    738: "ESCUELA DE VERANO",
    476: "PROYECTOS DE COOPERACION",
    665: "VACACIONALES EICEA",
    813: "EVENTOS GASTRONOMIA",
    4: "ADMON CIENCIAS ECON Y ADMTIVAS",
    365: "ADMIN POSTGRADOS EICEA",
    595: "FERIA EMPRESARIAL",
    441: "FEMI",
    768: "CHALLENGE ESCUELA",
    790: "INNOVACION EICEA"
}

deptos = pd.DataFrame(list(departamentos.items()), columns=['DEPARTAMENTO', 'Nombre_Departamento'])
datos = pd.merge(datos, deptos, how='left', left_on='Dpto', right_on='DEPARTAMENTO')

pivot_table = pd.pivot_table(
    datos,
    values='Presupuesto Disponible*',
    index=['Dpto', 'Nombre_Departamento'],
    aggfunc='sum',
    margins=True,
    margins_name='Total:'
)

dinamica = pivot_table.sort_values(by='Presupuesto Disponible*', ascending=False)
dinamica['Presupuesto Disponible*'] = dinamica['Presupuesto Disponible*'].apply(lambda x: f'{x:,.0f}'.replace(',', '.'))


# ---------------------------------------------------------------------------
# Interfaz (solo tkinter/ttk, sin librerías externas)
# ---------------------------------------------------------------------------
COLOR_FONDO = "#2b3e50"
COLOR_PANEL = "#34495e"
COLOR_TEXTO = "#ebebeb"
COLOR_VERDE = "#5cb85c"
COLOR_VERDE_OSCURO = "#449d44"
COLOR_SELECCION = "#4c9be8"

app.title("📊 Tabla Dinámica de Presupuesto")
app.geometry("900x600")
app.configure(bg=COLOR_FONDO)

estilo = ttk.Style(app)
estilo.theme_use("clam")
estilo.configure(".", background=COLOR_FONDO, foreground=COLOR_TEXTO, font=("Segoe UI", 10))
estilo.configure("Titulo.TLabel", font=("Segoe UI", 14, "bold"))
estilo.configure("Treeview", background=COLOR_PANEL, fieldbackground=COLOR_PANEL,
                 foreground=COLOR_TEXTO, rowheight=26, borderwidth=0)
estilo.map("Treeview", background=[("selected", COLOR_SELECCION)],
           foreground=[("selected", "white")])
estilo.configure("Treeview.Heading", background=COLOR_VERDE, foreground="white",
                 font=("Segoe UI", 10, "bold"), relief="flat")
estilo.map("Treeview.Heading", background=[("active", COLOR_VERDE_OSCURO)])
estilo.configure("Exportar.TButton", background=COLOR_FONDO, foreground=COLOR_VERDE,
                 bordercolor=COLOR_VERDE, font=("Segoe UI", 10, "bold"), padding=(14, 6))
estilo.map("Exportar.TButton", background=[("active", COLOR_VERDE)],
           foreground=[("active", "white")])
for orientacion in ("Vertical", "Horizontal"):
    estilo.configure(f"{orientacion}.TScrollbar", troughcolor=COLOR_FONDO,
                     background=COLOR_PANEL, arrowcolor=COLOR_TEXTO)

label_title = ttk.Label(app, text="Resumen de Presupuesto por Departamento", style="Titulo.TLabel")
label_title.pack(pady=15)

frame = ttk.Frame(app)
frame.pack(fill="both", expand=True, padx=20, pady=10)

scroll_y = ttk.Scrollbar(frame, orient="vertical")
scroll_x = ttk.Scrollbar(frame, orient="horizontal")

tabla = ttk.Treeview(
    frame,
    columns=("Departamento", "Nombre", "Presupuesto"),
    show="headings",
    yscrollcommand=scroll_y.set,
    xscrollcommand=scroll_x.set
)

scroll_y.config(command=tabla.yview)
scroll_y.pack(side="right", fill="y")
scroll_x.config(command=tabla.xview)
scroll_x.pack(side="bottom", fill="x")

tabla.heading("Departamento", text="Código Dpto")
tabla.heading("Nombre", text="Nombre del Departamento")
tabla.heading("Presupuesto", text="Presupuesto Disponible")

tabla.column("Departamento", width=120, anchor="center")
tabla.column("Nombre", width=500, anchor="w")
tabla.column("Presupuesto", width=200, anchor="e")

tabla.tag_configure("total", background=COLOR_VERDE_OSCURO, foreground="white",
                    font=("Segoe UI", 10, "bold"))

for (dpto, nombre), row in dinamica.iterrows():
    etiqueta = ("total",) if dpto == "Total:" else ()
    tabla.insert("", "end", values=(dpto, nombre, row['Presupuesto Disponible*']), tags=etiqueta)

tabla.pack(fill="both", expand=True)


def exportar_excel():
    archivo_salida = os.path.join(ruta, "tabla_dinamica_estetica.xlsx")
    try:
        dinamica.to_excel(archivo_salida)
    except ImportError:
        # Sin openpyxl instalado: se guarda como CSV, que Excel abre igual
        archivo_salida = archivo_salida.replace(".xlsx", ".csv")
        dinamica.to_csv(archivo_salida, sep=";", encoding="utf-8-sig")
    except PermissionError:
        messagebox.showerror("No se pudo guardar",
                             f"Cierra el archivo si está abierto en Excel:\n{archivo_salida}")
        return
    messagebox.showinfo("Exportación exitosa", f"Archivo guardado en:\n{archivo_salida}")


boton_exportar = ttk.Button(app, text="💾 Exportar a Excel", style="Exportar.TButton",
                            command=exportar_excel)
boton_exportar.pack(pady=10)

app.deiconify()
app.mainloop()
