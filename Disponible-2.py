import pandas as pd
import glob, os
import tkinter as tk
from tkinter import ttk, messagebox
import ttkbootstrap as tb  

ruta = r"C:\Users\norbertoacbe\Downloads"
archivos = glob.glob(os.path.join(ruta, 'ps*.xls'))

if not archivos:
    messagebox.showerror("Error", "No se encontraron archivos que coincidan con 'ps*.xls'")
    raise SystemExit

datos = pd.concat([pd.read_html(archivo)[0] for archivo in archivos], ignore_index=True)

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
    #columns='Ppto',
    index=['Dpto', 'Nombre_Departamento'],
    aggfunc='sum',
    margins=True,
    margins_name='Total:'
)

dinamica = pivot_table.sort_values(by='Presupuesto Disponible*', ascending=False)
dinamica['Presupuesto Disponible*'] = dinamica['Presupuesto Disponible*'].apply(lambda x: f'{x:,.0f}'.replace(',', '.'))

app = tb.Window(themename="superhero")  
app.title("📊 Tabla Dinámica de Presupuesto")
app.geometry("900x600")

label_title = ttk.Label(app, text="Resumen de Presupuesto por Departamento",
                        font=("Segoe UI", 14, "bold"))
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
    xscrollcommand=scroll_x.set,
    style="success.Treeview"
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

for (dpto, nombre), row in dinamica.iterrows():
    valor = row['Presupuesto Disponible*']
    tabla.insert("", "end", values=(dpto, nombre, valor))

tabla.pack(fill="both", expand=True)

def exportar_excel():
    archivo_salida = os.path.join(ruta, "tabla_dinamica_estetica.xlsx")
    dinamica.to_excel(archivo_salida)
    messagebox.showinfo("Exportación exitosa", f"Archivo guardado en:\n{archivo_salida}")

boton_exportar = tb.Button(app, text="💾 Exportar a Excel", bootstyle="success-outline", command=exportar_excel)
boton_exportar.pack(pady=10)

app.mainloop()