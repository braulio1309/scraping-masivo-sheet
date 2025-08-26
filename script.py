import time 
import gspread
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.keys import Keys
from oauth2client.service_account import ServiceAccountCredentials
import logging
from datetime import datetime

# Configuración de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class InterrapidisimoTracker:
    def __init__(self):
        # Configurar Selenium (usando Chrome)
        chrome_options = webdriver.ChromeOptions()
        #chrome_options.add_argument('--headless')  # Ejecutar en modo sin interfaz
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        # Añadir preferencias para manejar ventanas emergentes y nuevas pestañas
        chrome_options.add_experimental_option("prefs", {
            "profile.default_content_setting_values.notifications": 2,
        })
        self.driver = webdriver.Chrome(options=chrome_options)
        
        # Aquí iría la configuración de Google Sheets si la activas
        # Ejemplo:
        scope = ["https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
        self.client = gspread.authorize(creds)
        
    def get_shipment_status(self, tracking_number):
        """Obtiene el estado del envío desde la web de Interrapidísimo"""
        
        # Guardar la pestaña principal
        main_window = self.driver.current_window_handle
        
        # Navegar a la página de seguimiento
        self.driver.get("https://interrapidisimo.com/sigue-tu-envio/")
        
        # Localizar y completar el campo de número de guía (ID correcto: inputGuide)
        tracking_input = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.ID, "inputGuide"))
        )
        tracking_input.clear()
        tracking_input.send_keys(tracking_number)
        
        # Enviar la búsqueda con la tecla ENTER
        tracking_input.send_keys(Keys.ENTER)
        
        # Esperar a que se abra la nueva pestaña
        WebDriverWait(self.driver, 10).until(EC.number_of_windows_to_be(2))
        
        # Cambiar a la nueva pestaña
        new_window = [window for window in self.driver.window_handles if window != main_window][0]
        self.driver.switch_to.window(new_window)
        
        # Esperar a que cargue la información en la nueva pestaña
        time.sleep(3)
        
        # Intentar extraer el estado del envío
        status = "NO ENCONTRADO"
        
        try:
            estado_element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//h2[contains(text(), ' Tú envío ')]/following-sibling::p | //div[contains(@class, 'estado-actual')] | //p[contains(@class, 'status')]"))
            )
            status = estado_element.text.strip()
        except TimeoutException:
            try:
                paso_por_element = self.driver.find_element(By.XPATH, "//h3[contains(text(), 'Paso por')]/following-sibling::p")
                status = paso_por_element.text.strip()
            except NoSuchElementException:
                possible_status_elements = self.driver.find_elements(By.XPATH, "//p[contains(., 'entregado') or contains(., 'devuelto') or contains(., 'Ya puedes recoger tu envío')]")
                if possible_status_elements:
                    status = possible_status_elements[0].text.strip()
        
        # Cerrar la pestaña actual y volver a la principal
        self.driver.close()
        self.driver.switch_to.window(main_window)
        
        # Normalizar el estado
        status_lower = status.lower()
        if any(word in status_lower for word in ['entregado', 'entregada', 'entregar']):
            return "ENTREGADO"
        elif any(word in status_lower for word in ['En camino', 'Viajando', 'En Centro', 'ruta', 'camino']):
            return "EN TRÁNSITO"
        elif any(word in status_lower for word in ['pendiente', 'recibido', 'origen']):
            return "PENDIENTE"
        elif any(word in status_lower for word in ['devuelto', 'devolución']):
            return "DEVUELTO"
        elif any(word in status_lower for word in ['Ya puedes recoger tu envío']):
            return "EN AGENCIA"
        else:
            return status
            
        
    def update_google_sheet(self):
        """Actualiza la hoja de Google Sheets con los estados actualizados"""
        try:
            sheet = self.client.open("seguimiento").sheet1
            records = sheet.get_all_records()
            
            headers = sheet.row_values(1)
            if "STATUS TRACKING" not in headers:
                sheet.update_cell(1, len(headers) + 1, "STATUS TRACKING")
            if "Alerta" not in headers:
                sheet.update_cell(1, len(headers) + 2, "Alerta")
            
            headers = sheet.row_values(1)
            tracking_col = headers.index("ID TRACKING") + 1
            status_col = headers.index("STATUS DROPI") + 1
            estado_web_col = headers.index("STATUS TRACKING") + 1 if "STATUS TRACKING" in headers else len(headers) + 1
            alerta_col = headers.index("Alerta") + 1 if "Alerta" in headers else len(headers) + 2
            
            for i, record in enumerate(records, start=2):
                tracking_number = str(record.get("ID TRACKING", "")).strip()
                if tracking_number and tracking_number != "ID TRACKING":
                    web_status = self.get_shipment_status(tracking_number)
                    sheet.update_cell(i, estado_web_col, web_status)
                    
                    internal_status = record.get("STATUS DROPI", "")
                    alerta = "FALSE"
                    if web_status.upper() == "ENTREGADO" and internal_status.upper() != "ENTREGADO":
                        alerta = "TRUE"
                    sheet.update_cell(i, alerta_col, alerta)
                    
                    logging.info(f"Actualizado {tracking_number}: {web_status}, Alerta: {alerta}")
                    time.sleep(2)
            
            logging.info("Proceso de actualización completado")
            
        except Exception as e:
            logging.error(f"Error al actualizar Google Sheets: {str(e)}")
    
    def close(self):
        """Cerrar el navegador"""
        self.driver.quit()

def main():
    tracker = InterrapidisimoTracker()
    try:
        # 🔹 Aquí llamamos a la función que ya programaste para actualizar la hoja
        tracker.update_google_sheet()
    finally:
        tracker.close()

if __name__ == "__main__":
    main()

