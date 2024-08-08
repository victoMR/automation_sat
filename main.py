import calendar
import time
from tkinter import ttk
from tkinter.filedialog import asksaveasfilename

import customtkinter as ctk
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.service import Service as FirefoxService
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.firefox import GeckoDriverManager
import pyperclip

def browser_login():
    driver = webdriver.Firefox(service=FirefoxService(GeckoDriverManager().install()))
    driver.get("https://portalcfdi.facturaelectronica.sat.gob.mx/")
    return driver

def navigate_to_invoices(driver, option):
    if option == "emitidas":
        driver.find_element(By.LINK_TEXT, "Consultar Facturas Emitidas").click()
        enable_date_inputs(driver)
    elif option == "recibidas":
        driver.find_element(By.LINK_TEXT, "Consultar Facturas Recibidas").click()

def enable_date_inputs(driver):
    fecha_inicio = "01/08/2024"  # Formato día/mes/año
    fecha_fin = "31/08/2024"     # Puedes calcularlo dinámicamente

    driver.find_element(By.ID, "ctl00_MainContent_RdoFechas").click()

    # Establecer los valores de las fechas usando JavaScript
    driver.execute_script(
        "document.getElementById('ctl00_MainContent_CldFechaInicial2_Calendario_text').value = arguments[0];",
        fecha_inicio)
    driver.execute_script(
        "document.getElementById('ctl00_MainContent_CldFechaFinal2_Calendario_text').value = arguments[0];", fecha_fin)

    boton_buscar = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "ctl00_MainContent_BtnBusqueda"))
    )
    boton_buscar.click()

    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "ctl00_MainContent_tblResult"))
    )

    # Verificar que los campos de fecha sean accesibles antes de intentar interactuar con ellos
    input_fecha_inicio = driver.find_element(By.ID, "ctl00_MainContent_CldFechaInicial2_Calendario_text")
    driver.execute_script(
        "arguments[0].value = arguments[1];", input_fecha_inicio, fecha_inicio
    )

    input_fecha_fin = driver.find_element(By.ID, "ctl00_MainContent_CldFechaFinal2_Calendario_text")
    driver.execute_script(
        "arguments[0].value = arguments[1];", input_fecha_fin, fecha_fin
    )

    boton_buscar = driver.find_element(By.ID, "ctl00_MainContent_BtnBusqueda")
    boton_buscar.click()

    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "ctl00_MainContent_tblResult"))
    )

    tabla_resultados = driver.find_element(By.ID, "ctl00_MainContent_tblResult")
    filas = tabla_resultados.find_elements(By.TAG_NAME, "tr")[1:]  # Omitir la primera fila de encabezados

    datos = []
    for fila in filas:
        columnas = fila.find_elements(By.TAG_NAME, "td")
        datos_fila = [columna.text.replace(",", "") for columna in columnas]
        datos.append(datos_fila)

    return datos


def search_invoices_received(driver, year, month):
    driver.find_element(By.ID, "ctl00_MainContent_RdoFechas").click()

    driver.find_element(By.ID, "DdlAnio").send_keys(year)
    driver.find_element(By.ID, "ctl00_MainContent_CldFecha_DdlMes").send_keys(month)

    boton_buscar = driver.find_element(By.ID, "ctl00_MainContent_BtnBusqueda")
    boton_buscar.click()

def search_invoices_emitted(driver, year, month, day=None):
    driver.find_element(By.ID, "ctl00_MainContent_RdoFechas").click()

    # Habilitar los campos de fecha
    driver.execute_script(
        "document.getElementById('ctl00_MainContent_CldFechaInicial2_Calendario_text').removeAttribute('disabled');"
    )
    driver.execute_script(
        "document.getElementById('ctl00_MainContent_CldFechaFinal2_Calendario_text').removeAttribute('disabled');"
    )

    # Establecer las fechas usando JavaScript
    fecha_inicio = f"01/{month}/{year}"
    fecha_fin = f"30/{month}/{year}"

    driver.execute_script(
        "document.getElementById('ctl00_MainContent_CldFechaInicial2_Calendario_text').value = arguments[0];",
        fecha_inicio
    )
    driver.execute_script(
        "document.getElementById('ctl00_MainContent_CldFechaFinal2_Calendario_text').value = arguments[0];",
        fecha_fin
    )

    boton_buscar = driver.find_element(By.ID, "ctl00_MainContent_BtnBusqueda")
    boton_buscar.click()


def extract_invoice_data(driver):
    time.sleep(2)  # Esperar a que la página cargue
    invoice_data = []
    page_number = 1

    while True:
        print(f"Extrayendo datos de la página {page_number}")
        try:
            table_rows = driver.find_elements(By.XPATH, "//table[@id='ctl00_MainContent_tblResult']//tr")[1:]

            for row in table_rows:
                cells = row.find_elements(By.TAG_NAME, "td")
                cell_texts = [cell.text.strip() for cell in cells]
                cleaned_texts = [cell.replace(',', '').replace('$', '').strip() for cell in cell_texts if cell.strip()]

                if cleaned_texts:
                    invoice_data.append(cleaned_texts)

            print(f"Datos extraídos de la página {page_number}: {invoice_data}")

            pagination_div = driver.find_element(By.ID, "ctl00_MainContent_pageNavPosition")
            page_links = pagination_div.find_elements(By.TAG_NAME, "a")

            current_page_element = pagination_div.find_element(By.CLASS_NAME, "pg-selected")
            current_page_number = int(current_page_element.text) if current_page_element else 1

            next_page = None
            for link in page_links:
                page_number_text = link.text
                if page_number_text.isdigit() and int(page_number_text) > current_page_number:
                    next_page = link
                    break

            if next_page:
                print(f"Haciendo clic en la página siguiente: {next_page.text}")
                next_page.click()
                time.sleep(2)  # Esperar a que cargue la nueva página
                page_number += 1
            else:
                print("No hay más páginas disponibles o error en la navegación.")
                break

        except Exception as e:
            print(f"No se pudo encontrar el botón de siguiente página o no hay más páginas: {e}")
            break

    print("Datos extraídos:", invoice_data)
    return invoice_data

def save_to_excel(data, excel_file):
    data = [row for row in data if any(cell.strip() for cell in row)]

    if not data:
        print("No hay datos para guardar.")
        return

    new_df = pd.DataFrame(data, columns=[
        "Folio Fiscal", "RFC Emisor", "Nombre o Razón Social del Emisor", "RFC Receptor",
        "Nombre o Razón Social del Receptor", "Fecha de Emisión", "Fecha de Certificación",
        "PAC que Certificó", "Total", "Efecto del Comprobante", "Estatus de cancelación",
        "Estado del Comprobante", "Estatus de Proceso de Cancelación",
        "Fecha de Proceso de Cancelación", "RFC a cuenta de terceros"
    ])

    try:
        existing_data = pd.read_excel(excel_file, sheet_name='GASTOS SAT', engine='openpyxl')
    except Exception as e:
        print("Error al leer el archivo Excel:", e)
        return

    new_records = new_df[~new_df['Folio Fiscal'].isin(existing_data['Folio Fiscal'])]

    if new_records.empty:
        print("No hay nuevos registros para agregar.")
        return

    try:
        with pd.ExcelWriter(excel_file, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
            updated_data = pd.concat([existing_data, new_records], ignore_index=True)
            updated_data.to_excel(writer, sheet_name='GASTOS SAT', index=False)
        print(f"Datos nuevos guardados exitosamente en el archivo Excel. Se añadieron {new_records.shape[0]} registros.")
    except Exception as e:
        print("Error al guardar los datos en Excel:", e)

class App(ctk.CTk):
    def __init__(self, driver):
        super().__init__()
        self.driver = driver
        self.title("SAT Invoice Automation")
        self.geometry("300x280")

        self.rfc_entry = ctk.CTkEntry(self, placeholder_text="RFC")
        self.rfc_entry.pack(pady=10)

        self.password_entry = ctk.CTkEntry(self, placeholder_text="Contraseña", show="*")
        self.password_entry.pack(pady=10)

        self.captcha_entry = ctk.CTkEntry(self, placeholder_text="Captcha")
        self.captcha_entry.pack(pady=10)

        self.submit_button = ctk.CTkButton(self, text="Login", command=self.login)
        self.submit_button.pack(pady=10)

    def login(self):
        rfc = self.rfc_entry.get()
        password = self.password_entry.get()
        captcha = self.captcha_entry.get()

        try:
            self.driver.find_element(By.ID, "rfc").send_keys(rfc)
            self.driver.find_element(By.ID, "password").send_keys(password)
            self.driver.find_element(By.ID, "userCaptcha").send_keys(captcha)
            self.driver.find_element(By.ID, "submit").click()
            self.destroy()
            OptionWindow(self.driver)
        except Exception as e:
            print(f"Error en el login: {e}")

class OptionWindow(ctk.CTkToplevel):
    def __init__(self, driver):
        super().__init__()
        self.search_window = None
        self.title("Seleccionar Opción")
        self.geometry("300x200")
        self.driver = driver

        self.emitted_button = ctk.CTkButton(self, text="Facturas Emitidas", command=self.search_emitted_invoices)
        self.emitted_button.pack(pady=10)

        self.received_button = ctk.CTkButton(self, text="Facturas Recibidas", command=self.search_received_invoices)
        self.received_button.pack(pady=10)

    def search_emitted_invoices(self):
        if self.search_window is None:
            self.search_window = SearchWindow(self.driver, "emitidas")
        self.search_window.grab_set()

    def search_received_invoices(self):
        if self.search_window is None:
            self.search_window = SearchWindow(self.driver, "recibidas")
        self.search_window.grab_set()

class SearchWindow(ctk.CTkToplevel):
    def __init__(self, driver, option):
        super().__init__()
        self.driver = driver
        self.option = option
        self.title("Buscar Facturas")
        self.geometry("400x350")

        self.year_label = ctk.CTkLabel(self, text="Año")
        self.year_label.pack(pady=10)

        self.year_entry = ctk.CTkEntry(self)
        self.year_entry.pack(pady=10)

        self.label_month = ctk.CTkLabel(self, text="Mes")
        self.label_month.pack(pady=5)
        self.entry_month = ctk.CTkEntry(self)
        self.entry_month.pack(pady=5)

        self.day_label = ctk.CTkLabel(self, text="DIA")
        self.day_label.pack(pady=5)
        self.day_label = ctk.CTkEntry(self)
        self.day_label.pack(pady=5)

        self.search_button = ctk.CTkButton(self, text="Buscar", command=self.search_invoices)
        self.search_button.pack(pady=10)

        self.invoice_data = []
        self.canceled_count = 0
        self.total_sum = 0
        self.non_canceled_sum = 0

    def search_invoices(self):
        year = self.year_entry.get()
        month = self.entry_month.get()
        day = self.day_label.get()

        month_number = month

        navigate_to_invoices(self.driver, self.option)

        if self.option == "emitidas":
            search_invoices_emitted(self.driver, year, month_number, day)
        elif self.option == "recibidas":
            search_invoices_received(self.driver, year, month_number)

        self.invoice_data = extract_invoice_data(self.driver)
        self.process_invoice_data()
        self.show_results()

    def process_invoice_data(self):
        self.canceled_count = sum(1 for row in self.invoice_data if "Cancelada" in row)
        self.total_sum = sum(float(row[8].replace('$', '').replace(',', '')) for row in self.invoice_data if row[8].replace('$', '').replace(',', '').isdigit())
        self.non_canceled_sum = sum(float(row[8].replace('$', '').replace(',', '')) for row in self.invoice_data if "Cancelada" not in row and row[8].replace('$', '').replace(',', '').isdigit())

    def show_results(self):
        self.result_window = ctk.CTkToplevel(self)
        self.result_window.title("Resultados de Facturas")
        self.result_window.geometry("800x600")

        # Frame para seleccionar Emitidas o Recibidas
        tab_control = ttk.Notebook(self.result_window)
        tab_emitidas = ttk.Frame(tab_control)
        tab_recibidas = ttk.Frame(tab_control)

        tab_control.add(tab_emitidas, text="Emitidas")
        tab_control.add(tab_recibidas, text="Recibidas")
        tab_control.pack(expand=1, fill="both")

        # Configuración para la tabla de Emitidas
        self.create_invoice_table(tab_emitidas, self.invoice_data)

        # Configuración para la tabla de Recibidas
        self.create_invoice_table(tab_recibidas, self.invoice_data)

        # Botón para guardar
        save_button = ctk.CTkButton(self.result_window, text="Guardar en Excel", command=self.save_data)
        save_button.pack(pady=10)

        # Botón para copiar los datos al portapapeles
        copy_button = ctk.CTkButton(self.result_window, text="Copiar Datos", command=self.copy_data)
        copy_button.pack(pady=10)

    def create_invoice_table(self, parent_frame, invoice_data):
        table_frame = ctk.CTkFrame(parent_frame)
        table_frame.pack(fill="both", expand=True, padx=10, pady=10)

        columns = ["Folio Fiscal", "RFC Emisor", "Nombre o Razón Social del Emisor", "RFC Receptor",
                   "Nombre o Razón Social del Receptor", "Fecha de Emisión", "Fecha de Certificación",
                   "PAC que Certificó", "Total", "Efecto del Comprobante", "Estatus de cancelación",
                   "Estado del Comprobante", "Estatus de Proceso de Cancelación", "Fecha de Proceso de Cancelación",
                   "RFC a cuenta de terceros"]

        self.invoice_table = ttk.Treeview(table_frame, columns=columns, show="headings")
        self.invoice_table.pack(fill="both", expand=True)

        for col in columns:
            self.invoice_table.heading(col, text=col)
            self.invoice_table.column(col, width=100)

        for invoice in invoice_data:
            self.invoice_table.insert("", "end", values=invoice)

    def copy_data(self):
        # Copiar datos de la tabla
        table_data = []
        for item in self.invoice_table.get_children():
            values = self.invoice_table.item(item, 'values')
            table_data.append("\t".join(values))

        table_text = "\n".join(table_data)

        # Copiar datos del resumen
        summary_text = (
            f"Total de Facturas: {len(self.invoice_table.get_children())}\n"
            f"Facturas Canceladas: {self.canceled_count}\n"
            f"Suma Total: ${self.total_sum:,.2f}\n"
            f"Suma sin Canceladas: ${self.non_canceled_sum:,.2f}"
        )

        # Combinar y copiar ambos datos
        combined_text = f"Datos de la Tabla:\n{table_text}\n\nResumen:\n{summary_text}"
        pyperclip.copy(combined_text)

        print("Datos copiados al portapapeles.")

    def save_data(self):
        file_path = asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            save_to_excel(self.invoice_data, file_path)


if __name__ == "__main__":
    driver = browser_login()
    app = App(driver)
    app.mainloop()
