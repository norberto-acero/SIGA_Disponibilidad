#%%
"""
SIGA (PeopleSoft) -> Base de datos analítica (DuckDB + Parquet), sin Excel
==========================================================================

Diseño desde cero, orientado a ser rápido, simple de correr y escalable.

Flujo:
  1. Login + abrir consulta UNA vez (Selenium, robusto incluso con SSO).
  2. Por cada valor: rellenar Desde/Hasta, Fetch, y LEER la grilla HTML
     directamente del DOM (la misma tabla que el antiguo "ps.xls").
     Sin descargas, sin Excel.
  3. Volcar el detalle crudo a DuckDB y agregarlo con SQL (SUM por Dpto),
     uniendo el nombre del departamento.
  4. Persistir en Parquet (columnar, compacto) y dejar la base .duckdb
     consultable. Mostrar el resumen por consola y, opcional, GUI de solo
     lectura.

Salidas (todas reutilizables, ninguna es Excel):
  - siga.duckdb            -> base consultable (tablas: detalle, resumen)
  - resumen.parquet        -> agregado por departamento
  - detalle.parquet        -> detalle crudo completo

Requisitos:
    pip install selenium webdriver-manager pandas lxml duckdb
    # opcional, solo para la ventana de visualización:
    pip install ttkbootstrap

Ejecutar:
    python siga.py
"""

import getpass
import os
import re
import time

import duckdb
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager


# =========================================================
# CONFIGURACIÓN  (lo único que normalmente tocarás)
# =========================================================
URL = "https://sigafinanciero.unisabana.edu.co/psp/ERP92PRO/?&cmd=login&languageCd=ESP&"
SEL_USUARIO = "#userid"
SEL_PWD = "#pwd"

VALORES = [4, 5, 6, 7, 247, 268, 332, 365, 434, 436, 441, 519, 520, 542,
           595, 738, 784, 813, 814, 1372, 1438, 1441]

DIR_SALIDA = r"C:\DescargasSIGA"
DB_PATH = "siga.duckdb"
COLUMNA_VALOR = "Presupuesto Disponible*"   # columna que se suma

HEADLESS = True
MOSTRAR_GUI = True
TIMEOUT = 30

DEPARTAMENTOS = {
    247: "PROG DE ADMON MCDEO Y LOG INT", 332: "PROG DE ECONOMIA Y FINAN INTER",
    434: "PROGRAMA DE GASTRONOMIA", 5: "PROG DE ADMON DE EMPRESAS",
    7: "PROG DE ADMON DE NEGOC INTERN", 6: "PROG DE ADMON Y SERVICIO",
    814: "PROG ECONO FINAN INTER VIRTUAL", 1372: "MAESTRIA EN GERENCIA INVERSION",
    436: "MAE EN GERENCIA DE OPERACIONES", 542: "MAE EN GERENCIA ESTRATEGICA",
    1438: "MAE EN GERENCIA INTERN", 519: "MAE EN GERENCIA INNOVACION",
    520: "DOCTORADO EN ADMINISTRACION", 780: "MAE EN GERENCIA DE INVER VIRTU",
    784: "MAE EN GERENCIA NEGOCI HOSPITA", 268: "EDUCONTINUA EICEA",
    1441: "INTERNACIONALIZACION EN CASA", 454: "PROYECTOS CON EMPRESAS",
    738: "ESCUELA DE VERANO", 476: "PROYECTOS DE COOPERACION",
    665: "VACACIONALES EICEA", 813: "EVENTOS GASTRONOMIA",
    4: "ADMON CIENCIAS ECON Y ADMTIVAS", 365: "ADMIN POSTGRADOS EICEA",
    595: "FERIA EMPRESARIAL", 441: "FEMI", 768: "CHALLENGE ESCUELA",
    790: "INNOVACION EICEA",
}

VALOR_DB = re.sub(r"[^a-z0-9]+", "_", COLUMNA_VALOR.strip().lower()).strip("_")


# =========================================================
# HELPERS DE PEOPLESOFT
# =========================================================
def esperar_procesamiento(driver, timeout=TIMEOUT):
    fin = time.time() + timeout
    while time.time() < fin:
        try:
            listo = driver.execute_script("return document.readyState") == "complete"
            procesando = driver.execute_script("""
                try {
                    var ids = ['WAIT_win0', 'processing', 'WAIT_win0_modal'];
                    for (var i = 0; i < ids.length; i++) {
                        var e = document.getElementById(ids[i]);
                        if (e) {
                            var st = window.getComputedStyle(e);
                            if (st && st.display !== 'none' && st.visibility !== 'hidden')
                                return true;
                        }
                    }
                } catch (e) {}
                return false;
            """)
            if listo and not procesando:
                return True
        except Exception:
            pass
        time.sleep(0.2)
    return False


def cambiar_a_frame_contenido(driver, timeout=TIMEOUT):
    fin = time.time() + timeout
    while time.time() < fin:
        driver.switch_to.default_content()
        for by, val in [(By.ID, "ptifrmtgtframe"), (By.NAME, "TargetContent")]:
            try:
                driver.switch_to.frame(driver.find_element(by, val))
                if driver.execute_script("return typeof submitAction_win0 !== 'undefined';"):
                    return True
                driver.switch_to.default_content()
            except NoSuchElementException:
                pass
            except Exception:
                driver.switch_to.default_content()
        for f in driver.find_elements(By.TAG_NAME, "iframe"):
            try:
                driver.switch_to.frame(f)
                if driver.execute_script("return typeof submitAction_win0 !== 'undefined';"):
                    return True
                driver.switch_to.default_content()
            except Exception:
                driver.switch_to.default_content()
        time.sleep(0.4)
    return False


def click_robusto(driver, wait, by, locator):
    el = wait.until(EC.presence_of_element_located((by, locator)))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    try:
        wait.until(EC.element_to_be_clickable((by, locator))).click()
    except Exception:
        driver.execute_script("arguments[0].click();", el)


def localizar_campos_desde_hasta(driver):
    for id_desde, id_hasta in [
        ("KK_INQ_LD_CF_CHARTFIELD_VALUE$1", "KK_INQ_LD_CF_TO_VALUE$1"),
        ("KK_INQ_LD_CF_CHARTFIELD_VALUE$0", "KK_INQ_LD_CF_TO_VALUE$0"),
    ]:
        try:
            driver.find_element(By.ID, id_desde)
            driver.find_element(By.ID, id_hasta)
            return id_desde, id_hasta
        except NoSuchElementException:
            continue
    ids = driver.execute_script("""
        var out = [];
        document.querySelectorAll('[id]').forEach(function(e){
            if (e.id.indexOf('KK_INQ_LD_CF') !== -1) out.push(e.id);
        });
        return out;
    """)
    raise NoSuchElementException(
        "No encontré los campos Desde/Hasta. IDs con 'KK_INQ_LD_CF': " + str(ids))


def expandir_resultados(driver):
    try:
        return driver.execute_script("""
            var els = document.querySelectorAll('a, span, input');
            for (var i = 0; i < els.length; i++) {
                var el = els[i];
                var txt = (el.textContent || el.value || '').trim().toLowerCase();
                var id = (el.id || '').toLowerCase();
                if (txt === 'ver todo' || txt === 'view all' || txt === 'visualizar todo'
                    || id.indexOf('viewall') !== -1 || id.indexOf('hviewall') !== -1) {
                    el.click(); return true;
                }
            }
            return false;
        """)
    except Exception:
        return False


# IDs estables de cada columna en la grilla de resultados de PeopleSoft.
# El sufijo $N es el índice de fila (0,1,2,...). Los importes viven en
# AMTxL$span$N (AMT5L = Presupuesto Disponible*); el resto en spans directos.
_JS_EXTRAER_GRID = r"""
var rows = [];
function t(id){ var e = document.getElementById(id); return e ? e.textContent.replace(/\s+/g,' ').trim() : null; }
for (var i = 0; ; i++) {
    var dept = document.getElementById('KK_INQ_LD_WS_DEPTID$' + i);
    if (!dept) break;
    rows.push({
        grupo_contable:        t('KK_INQ_LD_WS_LEDGER_GROUP$' + i),
        cuenta:                t('KK_INQ_LD_WS_ACCOUNT$' + i),
        dpto:                  dept.textContent.trim(),
        proy:                  t('KK_INQ_LD_WS_PROJECT_ID$' + i),
        periodo_presupuesto:   t('BUDGET_PERIOD$' + i),
        ppto:                  t('AMT1L$span$' + i),
        gastos:                t('AMT2L$span$' + i),
        gravamen:              t('AMT3L$span$' + i),
        pregravamen:           t('AMT4L$span$' + i),
        presupuesto_disponible:t('AMT5L$span$' + i),
        porcentaje_disponible: t('KK_INQ_BOV_WRK_PERCENTAGE$' + i)
    });
}
return rows;
"""

_COLS_NUMERICAS = ["ppto", "gastos", "gravamen", "pregravamen",
                   "presupuesto_disponible", "porcentaje_disponible"]


def _a_numero(serie):
    # Los importes vienen como "35,829,651.110" (coma de miles, punto decimal).
    return pd.to_numeric(
        serie.astype(str).str.replace(",", "", regex=False).str.strip(),
        errors="coerce")


def extraer_grid(driver, valor):
    """Lee la grilla de resultados por IDs estables y devuelve un DataFrame limpio."""
    if expandir_resultados(driver):
        esperar_procesamiento(driver)
        cambiar_a_frame_contenido(driver)

    filas = driver.execute_script(_JS_EXTRAER_GRID)
    if not filas:
        return pd.DataFrame()

    df = pd.DataFrame(filas)
    df["dpto"] = pd.to_numeric(df["dpto"], errors="coerce").astype("Int64")
    for c in _COLS_NUMERICAS:
        if c in df.columns:
            df[c] = _a_numero(df[c])
    df["valor_iterado"] = valor
    return df


# =========================================================
# SCRAPER
# =========================================================
class Siga:
    def __init__(self, headless=HEADLESS, timeout=TIMEOUT):
        opts = webdriver.ChromeOptions()
        if headless:
            opts.add_argument("--headless=new")
        for a in ("--disable-gpu", "--no-sandbox", "--disable-dev-shm-usage",
                  "--disable-extensions", "--disable-infobars",
                  "--window-size=1920,1080", "--start-maximized"):
            opts.add_argument(a)
        self.driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()), options=opts)
        self.driver.set_window_size(1920, 1080)
        self.wait = WebDriverWait(self.driver, timeout)

    def login(self, usuario, contrasena):
        self.driver.get(URL)
        self.wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, SEL_USUARIO))).send_keys(usuario)
        pwd = self.driver.find_element(By.CSS_SELECTOR, SEL_PWD)
        pwd.send_keys(contrasena)
        pwd.send_keys(Keys.ENTER)

    def abrir_inquiry(self):
        self.wait.until(EC.element_to_be_clickable(
            (By.ID, "pthnavbca_MYFAVORITES"))).click()
        self.wait.until(EC.element_to_be_clickable(
            (By.ID, "crefli_fav_EP_KK_INQ_LEDGER_GBL"))).click()
        if not cambiar_a_frame_contenido(self.driver):
            raise RuntimeError("No se encontró el iframe de contenido de PeopleSoft.")
        self.driver.execute_script("submitAction_win0(document.win0, '#ICSearch');")
        esperar_procesamiento(self.driver)

    def consultar_valor(self, valor):
        cambiar_a_frame_contenido(self.driver)
        id_desde, id_hasta = localizar_campos_desde_hasta(self.driver)
        d = self.driver.find_element(By.ID, id_desde); d.clear(); d.send_keys(str(valor))
        h = self.driver.find_element(By.ID, id_hasta); h.clear(); h.send_keys(str(valor))
        click_robusto(self.driver, self.wait, By.ID, "KK_INQ_WRK_PB_FETCH")
        esperar_procesamiento(self.driver)
        cambiar_a_frame_contenido(self.driver)
        df = extraer_grid(self.driver, valor)
        self._asegurar_criterios()
        return df

    def _asegurar_criterios(self):
        try:
            cambiar_a_frame_contenido(self.driver)
            localizar_campos_desde_hasta(self.driver)
            return
        except NoSuchElementException:
            try:
                click_robusto(self.driver, self.wait, By.ID, "KK_INQ_WRK_PB_RETURN")
                esperar_procesamiento(self.driver)
            except Exception:
                pass

    def cerrar(self):
        try:
            self.driver.quit()
        except Exception:
            pass


def scrapear_todos(usuario, contrasena):
    siga = Siga()
    partes = []
    t0 = time.time()
    try:
        siga.login(usuario, contrasena)
        siga.abrir_inquiry()
        print("✔ Sesión iniciada y consulta abierta.\n")
        for valor in VALORES:
            ti = time.time()
            try:
                df = siga.consultar_valor(valor)
                if df.empty:
                    print(f"⚠ valor {valor}: sin datos.")
                    continue
                partes.append(df)
                print(f"✔ valor {valor}: {len(df):>5} filas  ({time.time() - ti:.1f}s)")
            except (TimeoutException, NoSuchElementException, RuntimeError) as e:
                print(f"✖ valor {valor}: {e}")
                siga._asegurar_criterios()
    finally:
        siga.cerrar()
    print(f"\n✔ Scraping terminado en {time.time() - t0:.1f}s.")
    return pd.concat(partes, ignore_index=True) if partes else pd.DataFrame()


# =========================================================
# ALMACENAMIENTO + AGREGACIÓN (DuckDB / Parquet, sin Excel)
# =========================================================
def guardar_y_agregar(detalle, dir_salida=DIR_SALIDA, db_path=DB_PATH):
    os.makedirs(dir_salida, exist_ok=True)
    db_full = os.path.join(dir_salida, db_path)

    detalle = detalle.copy()
    detalle["dpto"] = pd.to_numeric(detalle["dpto"], errors="coerce").astype("Int64")
    detalle[VALOR_DB] = pd.to_numeric(detalle[VALOR_DB], errors="coerce")

    deptos = pd.DataFrame(list(DEPARTAMENTOS.items()),
                          columns=["dpto", "nombre_departamento"])

    con = duckdb.connect(db_full)
    con.register("detalle_df", detalle)
    con.register("deptos_df", deptos)

    con.execute("CREATE OR REPLACE TABLE detalle AS SELECT * FROM detalle_df")
    con.execute("CREATE OR REPLACE TABLE departamentos AS SELECT * FROM deptos_df")
    con.execute(f"""
        CREATE OR REPLACE TABLE resumen AS
        SELECT
            d.dpto,
            dp.nombre_departamento,
            SUM(d.{VALOR_DB}) AS {VALOR_DB}
        FROM detalle d
        LEFT JOIN departamentos dp ON d.dpto = dp.dpto
        GROUP BY d.dpto, dp.nombre_departamento
        ORDER BY {VALOR_DB} DESC
    """)

    # Persistir en Parquet (columnar, compacto, reutilizable)
    p_resumen = os.path.join(dir_salida, "resumen.parquet")
    p_detalle = os.path.join(dir_salida, "detalle.parquet")
    con.execute(f"COPY resumen  TO '{p_resumen}' (FORMAT PARQUET)")
    con.execute(f"COPY detalle  TO '{p_detalle}' (FORMAT PARQUET)")

    resumen = con.execute("SELECT * FROM resumen").df()
    con.close()

    print(f"✔ Base DuckDB: {db_full}  (tablas: detalle, resumen, departamentos)")
    print(f"✔ Parquet: {p_resumen} | {p_detalle}")
    return resumen


def formatear_para_mostrar(resumen):
    """Añade fila Total y formatea números al estilo colombiano (solo presentación)."""
    total = resumen[VALOR_DB].sum()
    fila_total = pd.DataFrame(
        [{"dpto": pd.NA, "nombre_departamento": "Total:", VALOR_DB: total}])
    out = pd.concat([fila_total, resumen], ignore_index=True)
    out[VALOR_DB] = out[VALOR_DB].apply(lambda x: f"{x:,.0f}".replace(",", "."))
    return out


# =========================================================
# GUI de solo lectura (opcional)
# =========================================================
def mostrar_gui(tabla_fmt):
    try:
        from tkinter import ttk
        import ttkbootstrap as tb
    except ImportError:
        print("ℹ ttkbootstrap/tkinter no disponibles; se omite la GUI.")
        return
    app = tb.Window(themename="superhero")
    app.title("📊 Presupuesto por Departamento")
    app.geometry("900x600")
    ttk.Label(app, text="Resumen de Presupuesto por Departamento",
              font=("Segoe UI", 14, "bold")).pack(pady=15)
    frame = ttk.Frame(app); frame.pack(fill="both", expand=True, padx=20, pady=10)
    sy = ttk.Scrollbar(frame, orient="vertical")
    tabla = ttk.Treeview(frame, columns=("d", "n", "p"), show="headings",
                         yscrollcommand=sy.set, style="success.Treeview")
    sy.config(command=tabla.yview); sy.pack(side="right", fill="y")
    tabla.heading("d", text="Código Dpto"); tabla.column("d", width=120, anchor="center")
    tabla.heading("n", text="Nombre del Departamento"); tabla.column("n", width=500, anchor="w")
    tabla.heading("p", text="Presupuesto Disponible"); tabla.column("p", width=200, anchor="e")
    for _, r in tabla_fmt.iterrows():
        dpto = "" if pd.isna(r["dpto"]) else int(r["dpto"])
        tabla.insert("", "end", values=(dpto, r["nombre_departamento"], r[VALOR_DB]))
    tabla.pack(fill="both", expand=True)
    app.mainloop()


# =========================================================
# MAIN
# =========================================================
def main():
    usuario = input("Usuario: ")
    contrasena = getpass.getpass("Contraseña: ")

    detalle = scrapear_todos(usuario, contrasena)
    if detalle.empty:
        print("✖ No se obtuvieron datos.")
        return

    resumen = guardar_y_agregar(detalle)
    tabla_fmt = formatear_para_mostrar(resumen)

    print("\n=== RESUMEN ===")
    print(tabla_fmt.to_string(index=False))

    if MOSTRAR_GUI:
        mostrar_gui(tabla_fmt)


if __name__ == "__main__":
    main()