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

    
    def __init__(self, headless=True, download_path=None, timeout=15):
        self.headless = headless
        self.timeout = timeout
        self.download_path = download_path or os.path.join(os.getcwd(), "descargas_smv")
        self.driver = None
        self.wait = None
        self.setup_logging()
        
    def setup_logging(self):
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        self.logger = logging.getLogger(__name__)
    
    def setup_driver(self):
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
        timeout = timeout or self.timeout
        wait = WebDriverWait(self.driver, timeout)
        return wait.until(EC.presence_of_element_located((by, value)))
    
    def wait_for_element_clickable(self, by, value, timeout=None):
        timeout = timeout or self.timeout
        wait = WebDriverWait(self.driver, timeout)
        return wait.until(EC.element_to_be_clickable((by, value)))
    
    def safe_click(self, element):
        attempts = 0
        while attempts < 3:
            try:
                self.driver.execute_script("arguments[0].click();", element)
                return True
            except StaleElementReferenceException:
                attempts += 1
                time.sleep(1)
                self.logger.warning(f"Elemento stale detectado, reintentando... intento {attempts}")
        return False
    
    def find_element_with_retry(self, by, value, max_retries=3):
        for attempt in range(max_retries):
            try:
                element = self.driver.find_element(by, value)
                return element
            except StaleElementReferenceException:
                if attempt < max_retries - 1:
                    time.sleep(1)
                    self.logger.warning(f"Elemento stale detectado, reintentando... intento {attempt + 1}")
                else:
                    raise
    
    def setup_empresa_download_folder(self, empresa_nombre):
        caracteres_permitidos = set("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZáéíóúüÁÉÍÓÚÜ0123456789 -_")
    
        empresa_clean = "".join(c for c in empresa_nombre if c in caracteres_permitidos).rstrip()
        empresa_clean = empresa_clean.replace(' ', '_')
        
        empresa_path = os.path.join(self.download_path, empresa_clean)
        if not os.path.exists(empresa_path):
            os.makedirs(empresa_path)
        
        prefs = {
            "download.default_directory": empresa_path,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
            "profile.default_content_settings.popups": 0
        }
        self.driver.execute_cdp_cmd('Page.setDownloadBehavior', {
            'behavior': 'allow',
            'downloadPath': empresa_path
        })
        
        self.current_download_path = empresa_path
        return empresa_path
    

    def rename_downloaded_file(self, anio):
        try:
            time.sleep(3)
            
            files_before = set(os.listdir(self.current_download_path))
            time.sleep(2)
            files_after = set(os.listdir(self.current_download_path))
            
            new_files = files_after - files_before
            if not new_files:
                all_files = [f for f in os.listdir(self.current_download_path) 
                        if f.endswith(('.xls', '.xlsx')) and not f.startswith(str(anio))]
                if all_files:
                    all_files.sort(key=lambda x: os.path.getmtime(os.path.join(self.current_download_path, x)), reverse=True)
                    new_files = {all_files[0]}
            
            for filename in new_files:
                if filename.endswith(('.xls', '.xlsx')) and not filename.startswith('.'):
                    old_path = os.path.join(self.current_download_path, filename)
                    
                    # Determinar si necesitamos convertir de .xls a .xlsx
                    file_extension = os.path.splitext(filename)[1]
                    
                    if file_extension == '.xls':
                        # Convertir .xls a .xlsx usando el método de pd.read_html
                        try:
                            import pandas as pd
                            from openpyxl import load_workbook
                            from SMV_APP.analisis  import (
                                FormatoSituacionFinanciera,
                                FormatoResultados,
                                FormatoPatrimonio,
                                FormatoFlujoEfectivo
                            )
                            
                            # Crear el nuevo nombre con .xlsx
                            new_filename = f"{anio}-{os.path.splitext(filename)[0]}.xlsx"
                            new_path = os.path.join(self.current_download_path, new_filename)
                            
                            # Verificar si el archivo ya existe
                            if os.path.exists(new_path):
                                self.logger.info(f"Archivo ya existe, omitiendo: {new_filename}")
                                os.remove(old_path)
                                self.logger.info(f"Archivo descargado eliminado: {filename}")
                                return
                            
                            # Obtener nombre de empresa de la carpeta
                            dir_path = os.path.dirname(old_path)
                            nombre = os.path.basename(dir_path)
                            nombreEmpresa = nombre.replace('_', ' ')
                            
                            # Leer todas las tablas del archivo .xls
                            tablas = pd.read_html(old_path)
                            
                            # 1. GUARDAR DATAFRAME
                            with pd.ExcelWriter(new_path, engine="openpyxl") as writer:
                                for i, df in enumerate(tablas, start=1):
                                    
                                    if df.shape[1] >= 4 and i != 3:  # TODOS MENOS LA DE PATRIMONIO O FIRMANTES ELIMINAN LA COLUMNA B
                                        df_cleaned = df.drop(df.columns[[1, 3]], axis=1, errors='ignore')
                                    else:
                                        df_cleaned = df
                                    
                                    def clean_and_coerce(series):
                                        s = series.astype(str)
                                        s = s.str.replace('(', '-', regex=False).str.replace(')', '', regex=False)
                                        s = s.str.replace(',', '', regex=False)
                                        return pd.to_numeric(s, errors='coerce')
                                    
                                    if df_cleaned.shape[1] >= 3 and (i != 3 and i != 6):
                                        df_cleaned.iloc[:, [1, 2]] = df_cleaned.iloc[:, [1, 2]].apply(clean_and_coerce)
                                    else:
                                        df_cleaned.iloc[:, 2:df_cleaned.shape[1]] = df_cleaned.iloc[:, 2:df_cleaned.shape[1]].apply(clean_and_coerce)
                                    
                                    df_cleaned.to_excel(
                                        writer, 
                                        sheet_name=f"Hoja{i}", 
                                        index=False,
                                        header=True,
                                        startrow=6,
                                        startcol=2
                                    )
                            
                            # 2. Declaración de hojas de ESTADOS:
                            wb = load_workbook(new_path)
                            situaFinanciera = wb['Hoja1']
                            estaresultados = wb['Hoja2']
                            patrimonio = wb['Hoja3']
                            flujoEfectivo = wb['Hoja4']
                            
                            # 3. ABRIR CON OPENPYXL PARA APLICAR ESTILOS MANUALES
                            FormatoSituacionFinanciera(situaFinanciera, nombreEmpresa)
                            FormatoResultados(estaresultados, nombreEmpresa)
                            FormatoPatrimonio(wb, patrimonio, nombreEmpresa)
                            FormatoFlujoEfectivo(flujoEfectivo, nombreEmpresa)
                            
                            # 4. ELIMINAR HOJAS QUE NO SON NECESARIAS
                            if len(tablas) >= 5:
                                wb.remove(wb['Hoja5'])
                            if len(tablas) >= 6:
                                wb.remove(wb['Hoja6'])
                            
                            # 5. Guardar el archivo con estilos
                            wb.save(new_path)
                            
                            # Eliminar el archivo .xls original
                            os.remove(old_path)
                            self.logger.info(f"Archivo convertido y formateado: {filename} -> {new_filename}")
                            
                        except Exception as e:
                            self.logger.error(f"Error al convertir archivo .xls a .xlsx: {e}")
                            # Si falla la conversión, simplemente renombrar sin convertir
                            new_filename = f"{anio}-{filename}"
                            new_path = os.path.join(self.current_download_path, new_filename)
                            
                            if os.path.exists(new_path):
                                self.logger.info(f"Archivo ya existe, omitiendo: {new_filename}")
                                os.remove(old_path)
                                self.logger.info(f"Archivo descargado eliminado: {filename}")
                                return
                            
                            os.rename(old_path, new_path)
                            self.logger.info(f"Archivo renombrado sin conversión: {filename} -> {new_filename}")
                    
                    else:  # file_extension == '.xlsx'
                        new_filename = f"{anio}-{filename}"
                        new_path = os.path.join(self.current_download_path, new_filename)
                        
                        # Verificar si el archivo ya existe
                        if os.path.exists(new_path):
                            self.logger.info(f"Archivo ya existe, omitiendo: {new_filename}")
                            os.remove(old_path)
                            self.logger.info(f"Archivo descargado eliminado: {filename}")
                            return
                        
                        os.rename(old_path, new_path)
                        self.logger.info(f"Archivo renombrado: {filename} -> {new_filename}")
                    
                    break
                    
        except Exception as e:
            self.logger.error(f"Error al renombrar archivo: {e}")
            
    def select_empresa(self, empresa_nombre):
        try:
            self.logger.info(f"Seleccionando empresa: {empresa_nombre}")
            
            max_attempts = 3
            for attempt in range(max_attempts):
                try:
                    self.driver.execute_script("""
                        var input = document.getElementById('MainContent_TextBox1');
                        if (input) {
                            input._oldOnFocus = input.onfocus;
                            input._oldOnBlur = input.onblur;
                            input._oldOnChange = input.onchange;
                            input._oldOnInput = input.oninput;
                            
                            input.onfocus = null;
                            input.onblur = null;
                            input.onchange = null;
                            input.oninput = null;
                            
                            input.focus();
                            input.value = '';
                            input.value = arguments[0];
                            
                            setTimeout(function() {
                                input.onfocus = input._oldOnFocus;
                                input.onblur = input._oldOnBlur;
                                input.onchange = input._oldOnChange;
                                input.oninput = input._oldOnInput;
                                
                                var inputEvent = new Event('input', { bubbles: true });
                                var changeEvent = new Event('change', { bubbles: true });
                                
                                input.dispatchEvent(inputEvent);
                                input.dispatchEvent(changeEvent);
                            }, 100);
                        }
                    """, empresa_nombre)
                    
                    time.sleep(1)
                    
                    current_value = self.driver.execute_script("""
                        var input = document.getElementById('MainContent_TextBox1');
                        return input ? input.value : '';
                    """)
                    
                    self.logger.info(f"Intento {attempt + 1}: Valor actual: '{current_value}'")
                    
                    if current_value.strip() == empresa_nombre.strip():
                        self.logger.info("Empresa ingresada correctamente")
                        return True
                    else:
                        self.logger.warning(f"Valor no coincide, reintentando...")
                        time.sleep(1)
                        continue
                    
                except Exception as e:
                    self.logger.error(f"Error en intento {attempt + 1}: {e}")
                    if attempt < max_attempts - 1:
                        time.sleep(1)
                        continue
                    else:
                        raise
            
            self.logger.error("No se pudo ingresar el nombre de la empresa después de todos los intentos")
            return False
            
        except Exception as e:
            self.logger.error(f"Error al seleccionar empresa: {e}")
            return False
    
    
    def select_periodo_anual(self):
        try:
            self.logger.info("Seleccionando período anual")
            
            max_attempts = 3
            for attempt in range(max_attempts):
                try:
                    radio_anual = self.wait_for_element_clickable(By.ID, "MainContent_cboPeriodo_1")
                    
                    if not radio_anual.is_selected():
                        self.driver.execute_script("arguments[0].click();", radio_anual)
                        time.sleep(1)
                    
                    break
                    
                except StaleElementReferenceException:
                    if attempt < max_attempts - 1:
                        self.logger.warning(f"Elemento radio anual stale, reintentando... intento {attempt + 1}")
                        time.sleep(1)
                    else:
                        raise
            
            self.logger.info("Período anual seleccionado exitosamente")
            return True
            
        except Exception as e:
            self.logger.error(f"Error al seleccionar período anual: {e}")
            return False
    
    def select_anio(self, anio):
        try:
            self.logger.info(f"Seleccionando año: {anio}")
            
            max_attempts = 3
            for attempt in range(max_attempts):
                try:
                    select_element = self.wait_for_element(By.ID, "MainContent_cboAnio")
                    select_anio = Select(select_element)
                    select_anio.select_by_value(str(anio))
                    time.sleep(1)
                    
                    break
                    
                except StaleElementReferenceException:
                    if attempt < max_attempts - 1:
                        self.logger.warning(f"Elemento select año stale, reintentando... intento {attempt + 1}")
                        time.sleep(1)
                        select_element = self.find_element_with_retry(By.ID, "MainContent_cboAnio")
                    else:
                        raise
            
            self.logger.info(f"Año {anio} seleccionado exitosamente")
            return True
            
        except Exception as e:
            self.logger.error(f"Error al seleccionar año {anio}: {e}")
            return False
    
    def click_buscar(self):
        try:
            self.logger.info("Haciendo click en Buscar")
            
            empresa_input = self.driver.find_element(By.ID, "MainContent_TextBox1")
            current_value = empresa_input.get_attribute('value')
            self.logger.info(f"Valor de empresa antes de buscar: '{current_value}'")
            
            if not current_value.strip():
                self.logger.error("El campo empresa está vacío antes de buscar")
                return False
            
            max_attempts = 3
            for attempt in range(max_attempts):
                try:
                    btn_buscar = self.wait_for_element_clickable(By.ID, "MainContent_cbBuscar")
                    self.driver.execute_script("arguments[0].click();", btn_buscar)
                    
                    self.wait_for_element(By.XPATH, "//table//tr", timeout=10)
                    
                    break
                    
                except StaleElementReferenceException:
                    if attempt < max_attempts - 1:
                        self.logger.warning(f"Elemento buscar stale, reintentando... intento {attempt + 1}")
                        time.sleep(1)
                    else:
                        raise
            
            self.logger.info("Búsqueda completada exitosamente")
            return True
            
        except Exception as e:
            self.logger.error(f"Error al hacer click en Buscar: {e}")
            return False
    
    def check_resultados_disponibles(self):
        try:
            time.sleep(2)
            
            no_data_messages = [
                "//td[contains(text(), 'No se encontraron registros coincidentes con sus criterios de búsqueda.')]",
            ]
            
            for xpath in no_data_messages:
                try:
                    no_data = self.driver.find_elements(By.XPATH, xpath)
                    if no_data:
                        self.logger.info("No se encontraron resultados para este año")
                        return False
                except:
                    pass
            
            try:
                enlaces_detalle = self.driver.find_elements(
                    By.XPATH, "//a[contains(@title, 'Ver detalle de Estados Financieros')]"
                )
                if enlaces_detalle:
                    self.logger.info(f"Se encontraron {len(enlaces_detalle)} resultados disponibles")
                    return True
            except:
                pass
            
            return False
        
        except Exception as e:
            self.logger.error(f"Error al verificar resultados: {e}")
            return False

    def ver_detalle_estados_financieros(self):
        try:
            self.logger.info("Accediendo a detalle de estados financieros")
            
            self.driver.execute_script("window.scrollTo(0, 500);")
            time.sleep(1)
            
            main_window = self.driver.current_window_handle
            
            max_attempts = 3
            for attempt in range(max_attempts):
                try:
                    # Obtener todos los enlaces
                    enlaces_detalle = self.driver.find_elements(
                        By.XPATH, "//a[contains(@title, 'Ver detalle de Estados Financieros')]"
                    )
                    
                    if not enlaces_detalle:
                        self.logger.error("No se encontraron enlaces de detalle")
                        return False, main_window
                    
                    # Seleccionar el último enlace (el más actual)
                    enlace_mas_actual = enlaces_detalle[-1]
                    self.logger.info(f"Seleccionando el enlace más actual ({len(enlaces_detalle)} encontrados)")
                    
                    # Hacer scroll al elemento si es necesario
                    self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", enlace_mas_actual)
                    time.sleep(0.5)
                    
                    self.driver.execute_script("arguments[0].click();", enlace_mas_actual)
                    time.sleep(4)
                    
                    break
                    
                except StaleElementReferenceException:
                    if attempt < max_attempts - 1:
                        self.logger.warning(f"Elemento enlace detalle stale, reintentando... intento {attempt + 1}")
                        time.sleep(1)
                    else:
                        raise
            
            if len(self.driver.window_handles) > 1:
                for window in self.driver.window_handles:
                    if window != main_window:
                        self.driver.switch_to.window(window)
                        break
                
                self.wait_for_element(By.ID, "cbExcel", timeout=10)
                self.logger.info("Ventana de detalle cargada exitosamente")
                return True, main_window
            else:
                self.logger.error("No se abrió nueva ventana")
                return False, main_window
                
        except Exception as e:
            self.logger.error(f"Error al acceder a detalle: {e}")
            return False, None
    
    def descargar_excel(self, anio):
        try:
            self.logger.info("Iniciando descarga de Excel")
            
            max_attempts = 3
            for attempt in range(max_attempts):
                try:
                    btn_excel = self.wait_for_element_clickable(By.ID, "cbExcel")
                    self.driver.execute_script("arguments[0].click();", btn_excel)
                    time.sleep(3)
                    
                    break
                    
                except StaleElementReferenceException:
                    if attempt < max_attempts - 1:
                        self.logger.warning(f"Elemento excel stale, reintentando... intento {attempt + 1}")
                        time.sleep(1)
                    else:
                        raise
            
            self.rename_downloaded_file(anio)
            
            self.logger.info("Descarga de Excel iniciada exitosamente")
            return True
            
        except Exception as e:
            self.logger.error(f"Error al descargar Excel: {e}")
            return False
    
    def reset_to_main_form(self):
        try:
            self.driver.get("https://www.smv.gob.pe/SIMV/Frm_InformacionFinanciera?data=A70181B60967D74090DCD93C4920AA1D769614EC12")
            self.wait_for_element(By.ID, "MainContent_TextBox1", timeout=15)
            time.sleep(2)
            return True
        except Exception as e:
            self.logger.error(f"Error al resetear formulario: {e}")
            return False
    
    def procesar_anio(self, empresa_nombre, anio):
        self.logger.info(f"Iniciando procesamiento para {empresa_nombre} - Año {anio}")
        
        try:
            if not self.select_empresa(empresa_nombre):
                return False, "Error al seleccionar empresa", False
            
            if not self.select_periodo_anual():
                return False, "Error al seleccionar período anual", False
            
            if not self.select_anio(anio):
                return False, f"Error al seleccionar año {anio}", False
            
            if not self.click_buscar():
                return False, "Error al realizar búsqueda", False
            
            if not self.check_resultados_disponibles():
                self.logger.info(f"No hay datos disponibles para el año {anio}")
                self.reset_to_main_form()
                return False, f"Sin datos para año {anio}", True
            
            success, main_window = self.ver_detalle_estados_financieros()
            if not success:
                return False, "Error al acceder a detalle", False
            
            if not self.descargar_excel(anio):
                try:
                    self.driver.close()
                    if main_window:
                        self.driver.switch_to.window(main_window)
                except:
                    pass
                return False, "Error al descargar Excel", False
            
            try:
                self.driver.close()
                if main_window:
                    self.driver.switch_to.window(main_window)
            except:
                pass
            
            if not self.reset_to_main_form():
                return False, "Error al resetear formulario", False
            
            self.logger.info(f"Procesamiento completado para año {anio}")
            return True, "Éxito", False
            
        except Exception as e:
            self.logger.error(f"Error en procesamiento para año {anio}: {e}")
            try:
                if len(self.driver.window_handles) > 1:
                    self.driver.close()
                    self.driver.switch_to.window(self.driver.window_handles[0])
                self.reset_to_main_form()
            except:
                pass
            return False, str(e), False
    
    def determinar_anio_inicial(self, empresa_nombre, anio_base=2024):
        self.logger.info(f"Determinando año inicial desde {anio_base}")
        
        for anio_test in range(anio_base, anio_base - 10, -1):
            try:
                if not self.select_empresa(empresa_nombre):
                    continue
                
                if not self.select_periodo_anual():
                    continue
                
                if not self.select_anio(anio_test):
                    continue
                
                if not self.click_buscar():
                    continue
                
                if self.check_resultados_disponibles():
                    self.logger.info(f"Año inicial determinado: {anio_test}")
                    self.reset_to_main_form()
                    return anio_test
                
                self.reset_to_main_form()
                
            except Exception as e:
                self.logger.error(f"Error al verificar año {anio_test}: {e}")
                self.reset_to_main_form()
                continue
        
        self.logger.warning(f"No se encontraron datos en rango {anio_base} a {anio_base - 9}")
        return None
    
    def scrape_financial_data(self, empresa_nombre, anio_base=2024, rango_anios=5):
        resultados = {}
        
        try:
            self.setup_driver()
            
            empresa_path = self.setup_empresa_download_folder(empresa_nombre)
            
            self.logger.info("Navegando a la página de la SMV")
            self.driver.get("https://www.smv.gob.pe/SIMV/Frm_InformacionFinanciera?data=A70181B60967D74090DCD93C4920AA1D769614EC12")
            
            self.wait_for_element(By.ID, "MainContent_TextBox1", timeout=15)
            time.sleep(2)
            
            anio_inicial = self.determinar_anio_inicial(empresa_nombre, anio_base)
            
            if anio_inicial is None:
                return {
                    'status': 'error',
                    'message': f'No se encontraron datos disponibles desde {anio_base}',
                    'resultados': resultados,
                    'download_path': empresa_path
                }
            
            anios_a_procesar = list(range(anio_inicial, anio_inicial - rango_anios, -1))
            self.logger.info(f"Procesando años: {anios_a_procesar}")
            
            for anio in anios_a_procesar:
                success, mensaje, sin_datos = self.procesar_anio(empresa_nombre, anio)
                resultados[anio] = {
                    'success': success,
                    'message': mensaje
                }
                
                if not success and not sin_datos:
                    self.logger.warning(f"Falló el procesamiento para año {anio}")
                elif not success and sin_datos:
                    self.logger.info(f"Sin datos para año {anio}, continuando...")
                else:
                    self.logger.info(f"Éxito en procesamiento para año {anio}")
            
            return {
                'status': 'completado',
                'empresa': empresa_nombre,
                'anio_inicial': anio_inicial,
                'resultados': resultados,
                'download_path': empresa_path
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


def ejecutar_scraping_smv(empresa_nombre, anio_base=2024, rango_anios=5):
    scraper = SMVFinancialScraper(
        headless=False,
        download_path=os.path.join(os.getcwd(), "descargas_smv")
    )
    
    return scraper.scrape_financial_data(empresa_nombre, anio_base, rango_anios)

if __name__ == "__main__":
    empresa = "ADMINISTRADORA JOCKEY PLAZA SHOPPING CENTER S.A."
    
    resultado = ejecutar_scraping_smv(empresa, anio_base=2024, rango_anios=5)
    print(resultado)