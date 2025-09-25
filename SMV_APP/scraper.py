from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException
import time
import os
import logging
from pathlib import Path

class SMVFinancialScraper:
    """
    Clase robusta para scraping de información financiera de la SMV
    """
    
    def __init__(self, headless=True, download_path=None, timeout=30):
        self.headless = headless
        self.timeout = timeout
        self.download_path = download_path or os.path.join(os.getcwd(), "descargas_smv")
        self.driver = None
        self.wait = None
        self.setup_logging()
        
    def setup_logging(self):
        """Configurar logging"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        self.logger = logging.getLogger(__name__)
    
    def setup_driver(self):
        """Configurar el WebDriver de Chrome"""
        chrome_options = Options()
        
        if self.headless:
            chrome_options.add_argument("--headless")
        
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        # Configurar descargas
        if not os.path.exists(self.download_path):
            os.makedirs(self.download_path)
        
        prefs = {
            "download.default_directory": self.download_path,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
            "profile.default_content_settings.popups": 0
        }
        chrome_options.add_experimental_option("prefs", prefs)
        
        try:
            self.driver = webdriver.Chrome(options=chrome_options)
            self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            self.wait = WebDriverWait(self.driver, self.timeout)
            self.logger.info("WebDriver configurado exitosamente")
        except Exception as e:
            self.logger.error(f"Error al configurar WebDriver: {e}")
            raise
    
    def wait_for_element(self, by, value, timeout=None):
        """Esperar por un elemento con reintentos"""
        timeout = timeout or self.timeout
        wait = WebDriverWait(self.driver, timeout)
        return wait.until(EC.presence_of_element_located((by, value)))
    
    def wait_for_element_clickable(self, by, value, timeout=None):
        """Esperar por un elemento clickeable"""
        timeout = timeout or self.timeout
        wait = WebDriverWait(self.driver, timeout)
        return wait.until(EC.element_to_be_clickable((by, value)))
    
    def safe_click(self, element):
        """Click seguro con reintentos"""
        attempts = 0
        while attempts < 3:
            try:
                self.driver.execute_script("arguments[0].click();", element)
                return True
            except StaleElementReferenceException:
                attempts += 1
                time.sleep(1)
        return False
    
    def select_empresa(self, empresa_nombre):
        """Seleccionar empresa del combobox"""
        try:
            self.logger.info(f"Seleccionando empresa: {empresa_nombre}")
            
            # Localizar el input de empresa
            empresa_input = self.wait_for_element_clickable(By.ID, "MainContent_TextBox1")
            
            # Limpiar y escribir el nombre de la empresa
            empresa_input.clear()
            time.sleep(1)
            empresa_input.send_keys(empresa_nombre)
            
            # Esperar a que se carguen las opciones (si hay autocomplete)
            time.sleep(3)
            
            self.logger.info("Empresa seleccionada exitosamente")
            return True
            
        except Exception as e:
            self.logger.error(f"Error al seleccionar empresa: {e}")
            return False
    
    def select_periodo_anual(self):
        """Seleccionar período anual"""
        try:
            self.logger.info("Seleccionando período anual")
            
            # Localizar el radio button anual
            radio_anual = self.wait_for_element_clickable(By.ID, "MainContent_cboPeriodo_1")
            self.safe_click(radio_anual)
            
            # Esperar a que se procese la selección
            time.sleep(2)
            
            self.logger.info("Período anual seleccionado exitosamente")
            return True
            
        except Exception as e:
            self.logger.error(f"Error al seleccionar período anual: {e}")
            return False
    
    def select_anio(self, anio):
        """Seleccionar año específico"""
        try:
            self.logger.info(f"Seleccionando año: {anio}")
            
            # Localizar el combobox de año
            select_element = self.wait_for_element(By.ID, "MainContent_cboAnio")
            select_anio = Select(select_element)
            
            # Seleccionar el año
            select_anio.select_by_value(str(anio))
            
            # Esperar a que se procese la selección
            time.sleep(2)
            
            self.logger.info(f"Año {anio} seleccionado exitosamente")
            return True
            
        except Exception as e:
            self.logger.error(f"Error al seleccionar año {anio}: {e}")
            return False
    
    def click_buscar(self):
        """Hacer click en botón Buscar"""
        try:
            self.logger.info("Haciendo click en Buscar")
            
            # Localizar el botón Buscar
            btn_buscar = self.wait_for_element_clickable(By.ID, "MainContent_cbBuscar")
            self.safe_click(btn_buscar)
            
            # Esperar a que se carguen los resultados
            time.sleep(10)
            
            # Verificar que la búsqueda se completó
            self.wait_for_element(By.XPATH, "//table//tr", timeout=15)
            
            self.logger.info("Búsqueda completada exitosamente")
            return True
            
        except Exception as e:
            self.logger.error(f"Error al hacer click en Buscar: {e}")
            return False
    
    def ver_detalle_estados_financieros(self):
        """Hacer click en Ver detalle de estados financieros"""
        try:
            self.logger.info("Accediendo a detalle de estados financieros")
            
            # Hacer scroll para asegurar visibilidad
            self.driver.execute_script("window.scrollTo(0, 500);")
            time.sleep(2)
            
            # Localizar el enlace de detalle (usando XPath más específico)
            enlace_detalle = self.wait_for_element_clickable(
                By.XPATH, "//a[contains(@title, 'Ver detalle de Estados Financieros')]"
            )
            
            # Guardar la ventana actual
            main_window = self.driver.current_window_handle
            
            # Hacer click en el enlace
            self.safe_click(enlace_detalle)
            
            # Esperar a que se abra la nueva ventana
            time.sleep(5)
            
            # Cambiar a la nueva ventana
            if len(self.driver.window_handles) > 1:
                for window in self.driver.window_handles:
                    if window != main_window:
                        self.driver.switch_to.window(window)
                        break
                
                # Esperar a que cargue la nueva página
                self.wait_for_element(By.ID, "cbExcel", timeout=15)
                self.logger.info("Ventana de detalle cargada exitosamente")
                return True, main_window
            else:
                self.logger.error("No se abrió nueva ventana")
                return False, main_window
                
        except Exception as e:
            self.logger.error(f"Error al acceder a detalle: {e}")
            return False, None
    
    def descargar_excel(self):
        """Descargar archivo Excel"""
        try:
            self.logger.info("Iniciando descarga de Excel")
            
            # Localizar el botón de Excel
            btn_excel = self.wait_for_element_clickable(By.ID, "cbExcel")
            self.safe_click(btn_excel)
            
            # Esperar a que se complete la descarga
            time.sleep(10)
            
            self.logger.info("Descarga de Excel iniciada exitosamente")
            return True
            
        except Exception as e:
            self.logger.error(f"Error al descargar Excel: {e}")
            return False
    
    def procesar_anio(self, empresa_nombre, anio):
        """Procesar un año específico"""
        self.logger.info(f"Iniciando procesamiento para {empresa_nombre} - Año {anio}")
        
        try:
            # Paso 1: Seleccionar empresa
            if not self.select_empresa(empresa_nombre):
                return False, "Error al seleccionar empresa"
            
            # Paso 2: Seleccionar período anual
            if not self.select_periodo_anual():
                return False, "Error al seleccionar período anual"
            
            # Paso 3: Seleccionar año
            if not self.select_anio(anio):
                return False, f"Error al seleccionar año {anio}"
            
            # Paso 4: Click en Buscar
            if not self.click_buscar():
                return False, "Error al realizar búsqueda"
            
            # Paso 5: Ver detalle de estados financieros
            success, main_window = self.ver_detalle_estados_financieros()
            if not success:
                return False, "Error al acceder a detalle"
            
            # Paso 6: Descargar Excel
            if not self.descargar_excel():
                # Cerrar ventana de detalle antes de retornar
                self.driver.close()
                if main_window:
                    self.driver.switch_to.window(main_window)
                return False, "Error al descargar Excel"
            
            # Paso 7: Cerrar ventana de detalle y volver a la principal
            self.driver.close()
            if main_window:
                self.driver.switch_to.window(main_window)
            
            # Esperar antes del siguiente año
            time.sleep(3)
            
            self.logger.info(f"Procesamiento completado para año {anio}")
            return True, "Éxito"
            
        except Exception as e:
            self.logger.error(f"Error en procesamiento para año {anio}: {e}")
            # Intentar recuperar el control
            try:
                if len(self.driver.window_handles) > 1:
                    self.driver.close()
                    self.driver.switch_to.window(self.driver.window_handles[0])
            except:
                pass
            return False, str(e)
    
    def scrape_financial_data(self, empresa_nombre, anios=None):
        """
        Método principal para scrapear datos financieros
        
        Args:
            empresa_nombre (str): Nombre de la empresa
            anios (list): Lista de años a descargar (ej: [2024, 2022, 2020])
        """
        if anios is None:
            anios = [2024, 2022, 2020]
        
        resultados = {}
        
        try:
            # Configurar driver
            self.setup_driver()
            
            # Navegar a la página de la SMV
            self.logger.info("Navegando a la página de la SMV")
            self.driver.get("https://www.smv.gob.pe/SIMV/Frm_InformacionFinanciera?data=A70181B60967D74090DCD93C4920AA1D769614EC12")
            
            # Esperar a que cargue la página
            self.wait_for_element(By.ID, "MainContent_TextBox1", timeout=20)
            time.sleep(5)
            
            # Procesar cada año
            for anio in anios:
                success, mensaje = self.procesar_anio(empresa_nombre, anio)
                resultados[anio] = {
                    'success': success,
                    'message': mensaje
                }
                
                if not success:
                    self.logger.warning(f"Falló el procesamiento para año {anio}")
                else:
                    self.logger.info(f"Éxito en procesamiento para año {anio}")
            
            return {
                'status': 'completado',
                'empresa': empresa_nombre,
                'resultados': resultados,
                'download_path': self.download_path
            }
            
        except Exception as e:
            self.logger.error(f"Error general en scraping: {e}")
            return {
                'status': 'error',
                'message': str(e),
                'resultados': resultados
            }
        
        finally:
            if self.driver:
                self.driver.quit()
                self.logger.info("WebDriver cerrado")


def ejecutar_scraping_smv(empresa_nombre, anios=None):
    """
    Función simple para usar en views de Django
    """
    scraper = SMVFinancialScraper(
        headless=True,
        download_path=os.path.join(os.getcwd(), "descargas_smv")
    )
    
    return scraper.scrape_financial_data(empresa_nombre, anios)