import logging
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
from win32com.client import DispatchEx
import time
from selenium.webdriver.chrome.service import Service
import logging
import requests
import pandas as pd
import os
import sys
import json
import re

def initialize_driver(download_path=None, headless=False):
    try:
        """Inicializa el controlador de Chrome con la ruta de descarga especificada y otras configuraciones."""
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument("--ignore-certificate-errors")
        chrome_options.add_argument("--ignore-ssl-errors")
        
        if headless:
            chrome_options.add_argument("--headless")
            chrome_options.add_argument("window-size=1920,1080")
        
        # Desactivar la pantalla de bienvenida y otras opciones de primer uso
        chrome_options.add_argument("--no-default-browser-check")
        chrome_options.add_argument("--no-first-run")
        chrome_options.add_argument("--disable-default-apps")
        chrome_options.add_argument("--start-maximized")  # Iniciar en pantalla completa
        chrome_options.add_argument("--disable-popup-blocking")  # Desactivar el bloqueo de ventanas emergentes
        chrome_options.add_argument("--disable-notifications")  # Desactivar notificaciones
        chrome_options.add_argument("--incognito")  # Iniciar en modo incógnito (sin datos persistentes)
        chrome_options.add_argument("--disable-search-engine-choice-screen") # Desactiva la pantalla de motor de búsqueda predeterminado

        prefs = {
            "download.default_directory": download_path,  # Define la ruta de descarga
            "download.prompt_for_download": False,  # No preguntar al descargar
            "download.directory_upgrade": True,  # Permitir actualización de directorio
            "safebrowsing.enabled": True,  # Habilitar navegación segura
            "homepage": "https://www.google.com",
            "default_search_provider.enabled": True,
            "default_search_provider.name": "Google",
            "default_search_provider.keyword": "google.com",
            "default_search_provider.search_url": "https://www.google.com/search?q={searchTerms}",
            "default_search_provider.suggest_url": "https://www.google.com/complete/search?output=toolbar&q={searchTerms}",
            "default_search_provider.new_tab_url": "https://www.google.com/",
            "browser.startup.homepage": "https://www.google.com",
            "browser.startup.homepage_override.mstone": "ignore",
            "intl.accept_languages": "en,en_US",  # Asegúrate de que el idioma esté configurado en inglés
            "browser.search.defaultenginename": "Google",
        }
        chrome_options.add_experimental_option("prefs", prefs)
        
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        return driver
    except Exception as e:
        error_msg = f"Error al inicializar Chrome: {e}"
        logging.error(error_msg)

def esperar_y_clicar(driver, selector_tipo, selector, indice=1, mensaje_log="", es_critico=True):
    try:
        # Definir el tipo de selector de acuerdo al parámetro proporcionado
        by_type = {
            'id': By.ID,
            'xpath': By.XPATH,
            'name': By.NAME,
            'css': By.CSS_SELECTOR,
            'class_name': By.CLASS_NAME,
            'tag_name': By.TAG_NAME,
            'link_text': By.LINK_TEXT,
            'partial_link_text': By.PARTIAL_LINK_TEXT
        }

        if selector_tipo not in by_type:
            raise ValueError(f"Tipo de selector no soportado: {selector_tipo}")

        # Esperar a que los elementos estén presentes y clicables
        WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((by_type[selector_tipo], selector)))
        clickable_elements = WebDriverWait(driver, 10).until(EC.visibility_of_all_elements_located((by_type[selector_tipo], selector)))

        if not clickable_elements:
            raise TimeoutException(f"No se encontraron elementos para el selector: {selector}")
        
        if len(clickable_elements) >= indice > 0:
            element = clickable_elements[indice - 1]  # Los índices en Python empiezan en 0, así que restamos 1.
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((by_type[selector_tipo], selector)))
            element.click()
            logging.info(mensaje_log)
        else:
            raise IndexError(f"El índice {indice} está fuera de los límites para los elementos encontrados.")
    except TimeoutException as e:
        error_msg = f"Tiempo de espera excedido para {mensaje_log}: {e}"
        logging.error(error_msg)
        if es_critico:
            raise e
    except Exception as e:
        error_msg = f"Error al intentar clicar: {mensaje_log}: {e}"
        logging.error(error_msg)
        if es_critico:
            raise e

def esperar_y_doble_clicar(driver, locator, mensaje_log, es_critico=True):
    try:
        element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(locator))
        action_chains = ActionChains(driver)
        action_chains.double_click(element).perform()
        logging.info(mensaje_log)
    except TimeoutException as e:
        error_msg = f"Tiempo de espera excedido para {mensaje_log}: {e}"
        logging.error(error_msg)

def esperar_y_clicar_descarga(driver, selector_tipo, selector, indice=1, mensaje_log="", es_critico=True, 
                              esperar_descarga=True, download_path=None, nombre_fichero = ""):
    try:
        # Definir el tipo de selector de acuerdo al parámetro proporcionado
        by_type = {
            'id': By.ID,
            'xpath': By.XPATH,
            'name': By.NAME,
            'css': By.CSS_SELECTOR,
            'class_name': By.CLASS_NAME,
            'tag_name': By.TAG_NAME,
            'link_text': By.LINK_TEXT,
            'partial_link_text': By.PARTIAL_LINK_TEXT
        }

        if selector_tipo not in by_type:
            raise ValueError(f"Tipo de selector no soportado: {selector_tipo}")

        # Esperar a que los elementos estén presentes y clicables
        WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((by_type[selector_tipo], selector)))
        clickable_elements = WebDriverWait(driver, 10).until(EC.visibility_of_all_elements_located((by_type[selector_tipo], selector)))

        if not clickable_elements:
            raise TimeoutException(f"No se encontraron elementos para el selector: {selector}")
        
        if len(clickable_elements) >= indice > 0:
            element = clickable_elements[indice - 1]  # Los índices en Python empiezan en 0, así que restamos 1.
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((by_type[selector_tipo], selector)))
            element.click()
            logging.info(mensaje_log)
        else:
            raise IndexError(f"El índice {indice} está fuera de los límites para los elementos encontrados.")

        # Si se debe esperar la descarga, hacer la espera
        if esperar_descarga:
		    # Espera a que el archivo se descargue completamente en la carpeta de descargas.
            start_time = time.time()
            timeout = 60
            while True:
				# Revisar si hay algún archivo incompleto (con extensión .crdownload en Chrome)
                # Verificar si el archivo está presente en la carpeta de descargas
                if os.path.exists(os.path.join(download_path, nombre_fichero)) and not any([filename.endswith(".crdownload") for filename in os.listdir(download_path)]):
                    break
                elif time.time() - start_time > timeout:
                    raise TimeoutException("El tiempo de espera para la descarga ha excedido el límite.")
                time.sleep(1)  # Esperar 1 segundo antes de volver a revisar

    except TimeoutException as e:
        error_msg = f"Tiempo de espera excedido para {mensaje_log}: {e}"
        logging.error(error_msg)
        if es_critico:
            raise e
    except Exception as e:
        error_msg = f"Error al intentar clicar para descargar: {mensaje_log}: {e}"
        logging.error(error_msg)
        if es_critico:
            raise e

def cambiar_al_iframe(driver, index, mensaje_log, es_critico=True):
    try:
        driver.switch_to.frame(index)
        logging.info(mensaje_log)
    except Exception as e:
        error_msg = f"Error al cambiar al iframe: {mensaje_log}: {e}"
        logging.error(error_msg)

def obtener_html_content(driver, mensaje_log, es_critico=True):
    try:
        html_content = driver.page_source
        logging.info(mensaje_log)
        return html_content
    except Exception as e:
        error_msg = f"Error al obtener el contenido HTML de la página: {mensaje_log}: {e}"
        logging.error(error_msg)
        
def esperar_y_seleccionar_desplegable(driver, selector_tipo, selector, valor, mensaje_log, es_critico=True):
    try:
        # Definir el tipo de selector de acuerdo al parámetro proporcionado
        by_type = {
            'id': By.ID,
            'xpath': By.XPATH,
            'name': By.NAME,
            'css': By.CSS_SELECTOR,
            'class_name': By.CLASS_NAME,
            'tag_name': By.TAG_NAME,
            'link_text': By.LINK_TEXT,
            'partial_link_text': By.PARTIAL_LINK_TEXT
        }

        if selector_tipo not in by_type:
            raise ValueError(f"Tipo de selector no soportado: {selector_tipo}")

        # Esperar hasta que el desplegable esté presente
        try:
            desplegable = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((by_type[selector_tipo], selector))
            )
            logging.info(f"Desplegable encontrado: {selector}")
        except TimeoutException:
            error_msg = f"No se encontró el desplegable para {mensaje_log}"
            logging.error(error_msg)
            if es_critico:
                raise TimeoutException(error_msg)
            return

        # Esperar hasta que todas las opciones estén presentes
        try:
            opciones = WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, f"//select[@{selector_tipo}='{selector}']/option"))
            )
            logging.info(f"Opciones del desplegable cargadas: {selector}")
        except TimeoutException:
            error_msg = f"No se encontraron las opciones del desplegable para {mensaje_log}"
            logging.error(error_msg)
            if es_critico:
                raise TimeoutException(error_msg)
            return
        
        # Encontrar la opción que coincide con el valor
        opcion_encontrada = False
        for opcion in opciones:
            if opcion.get_attribute('value') == valor:
                opcion.click()
                opcion_encontrada = True
                break
        
        if opcion_encontrada:
            logging.info(mensaje_log)
        else:
            raise ValueError(f"No se encontró la opción con el valor {valor}")
            
    except (TimeoutException, ValueError) as e:
        error_msg = f"Error al seleccionar el desplegable para {mensaje_log}: {e}"
        logging.error(error_msg)
        if es_critico:
            raise e

def enviar_texto_a_input(driver, selector_tipo, selector, indice, texto, mensaje_log, es_critico=True):
    try:
        # Definir el tipo de selector de acuerdo al parámetro proporcionado
        by_type = {
            'id': By.ID,
            'xpath': By.XPATH,
            'name': By.NAME,
            'css': By.CSS_SELECTOR,
            'class_name': By.CLASS_NAME,
            'tag_name': By.TAG_NAME,
            'link_text': By.LINK_TEXT,
            'partial_link_text': By.PARTIAL_LINK_TEXT
        }

        if selector_tipo not in by_type:
            raise ValueError(f"Tipo de selector no soportado: {selector_tipo}")

        WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((by_type[selector_tipo], selector)))
        input_elements = driver.find_elements(by_type[selector_tipo], selector)
        
        if len(input_elements) >= indice:
            input_element = input_elements[indice - 1]  # Los índices en Python empiezan en 0, así que restamos 1.
            input_element.clear()
            input_element.send_keys(texto)
            time.sleep(1)
            logging.info(mensaje_log)
        else:
            raise IndexError(f"El índice {indice} está fuera de los límites para los elementos encontrados.")
    except Exception as e:
        error_msg = f"Error al enviar texto al input {indice}: {mensaje_log}: {e}"
        logging.error(error_msg)
        if es_critico:
            raise e

def leer_texto_de_input(driver, selector_tipo, selector, indice, mensaje_log, es_critico=True):
    try:
        # Definir el tipo de selector de acuerdo al parámetro proporcionado
        by_type = {
            'id': By.ID,
            'xpath': By.XPATH,
            'name': By.NAME,
            'css': By.CSS_SELECTOR,
            'class_name': By.CLASS_NAME,
            'tag_name': By.TAG_NAME,
            'link_text': By.LINK_TEXT,
            'partial_link_text': By.PARTIAL_LINK_TEXT
        }

        if selector_tipo not in by_type:
            raise ValueError(f"Tipo de selector no soportado: {selector_tipo}")

        # Esperar a que los elementos estén presentes
        WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((by_type[selector_tipo], selector)))
        
        # Buscar los elementos input usando el tipo de selector especificado
        input_elements = driver.find_elements(by_type[selector_tipo], selector)
        
        if len(input_elements) >= indice:
            input_element = input_elements[indice - 1]  # Los índices en Python empiezan en 0, así que restamos 1.
            
            # Obtener el valor actual del elemento input
            valor_input = input_element.get_attribute('value')
            logging.info(f"{mensaje_log}: {valor_input}")
            return valor_input  # Retorna el valor leído del input
        else:
            raise IndexError(f"El índice {indice} está fuera de los límites para los elementos encontrados.")
    except Exception as e:
        error_msg = f"Error al leer texto del input {indice}: {mensaje_log}: {e}"
        logging.error(error_msg)
        if es_critico:
            raise e

def quitar_decimales_no_significativos(horas_str):
    # Convertir el string a float, reemplazando coma por punto para manejar correctamente los decimales
    horas = float(horas_str.replace(',', '.'))
    
    # Si el número es un entero, se elimina la parte decimal
    if horas.is_integer():
        return str(int(horas))
    else:
        # Formatear para quitar ceros no significativos, y luego reemplazar punto por coma para el decimal
        return format(horas, ".2f").rstrip('0').rstrip('.').replace('.', ',')
    
def crear_estructura_carpetas(cliente, productor, resultado, dcs):
    base_path = os.path.join("Documentos", cliente, productor)
    if resultado == 'OK':
        path = os.path.join(base_path, "DCS Oks")
    else:
        path = os.path.join(base_path, "Errores")
    
    if not os.path.exists(path):
        os.makedirs(path)

    return path

def setup_logging(current_time) -> str:
    """
    Configura el registro (logging) para el script.

    Returns:
    - log_path (str): path al log generado.
    """

    LOG_ENABLED = True
    log_path = ""  # Inicializa la variable para el caso de que LOG_ENABLED sea False
    
    if LOG_ENABLED:
        # Verifica si el script está siendo ejecutado como un ejecutable
        if getattr(sys, 'frozen', False):
            # Si es así, el directorio base es el directorio donde se encuentra el .exe
            directorio_base = os.path.dirname(sys.executable)
        else:
            # Si no, el directorio base es el directorio donde se encuentra el script .py
            directorio_base = os.path.dirname(os.path.abspath(__file__))
        
        log_directory = os.path.join(directorio_base, "logs")
        
        # Obtener la fecha y hora actual para el nombre del archivo de logF
        log_filename = current_time.strftime("%Y%m%d_%H%M%S") + ".log"
        
        if not os.path.exists(log_directory):
            os.makedirs(log_directory)
        
        log_path = os.path.join(log_directory, log_filename)
        logging.basicConfig(filename=log_path, level=logging.INFO, 
                            format='%(asctime)s - %(levelname)s - %(message)s')
    else:
        # Configurar el logging para que muestre los mensajes en la terminal
        logging.basicConfig(level=logging.INFO, 
                            format='%(asctime)s - %(levelname)s - %(message)s')

    return log_path

def obtener_json(driver, url):
    """
    Proceso para obtener los JSON
    """
    try:
        driver.get(url)
        driver.maximize_window()
        
        # Espera a que el contenido de la página se cargue completamente
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.TAG_NAME, 'body'))
        )
        
        # Obtén el contenido de la página
        page_content = driver.find_element(By.TAG_NAME, 'pre').text  # Supone que el contenido JSON está en una etiqueta <pre>
        
        # Parsear el contenido JSON
        data = json.loads(page_content)
        dataframe = pd.DataFrame(data['DOSC'])
    
        logging.info("Documentos de inicio procesados correctamente:")
        logging.info(data)
        return dataframe
    
    except Exception as e:
        logging.error(f"Error al obtener el JSON: {url}", exc_info=True)
        logging.info("Documentos de inicio con errores:")

def is_alphanumeric(value):
    # Verificar si el valor es un string y contiene solo caracteres alfanuméricos
    if isinstance(value, str) and re.match(r'^[a-zA-Z0-9]+$', value):
        return True
    return False

# Función para hacer clic en una coordenada específica de la página
def click_en_coordenada(driver, x, y):
    action = ActionChains(driver)
    action.move_by_offset(x, y).click().perform()
    # Mover el cursor de vuelta al inicio para evitar efectos secundarios no deseados
    action.move_by_offset(-x, -y).perform()

def elemento_visible(driver, selector_tipo, selector):
    try:
        # Esperar hasta que el elemento sea visible
        elemento = WebDriverWait(driver, 1).until(
            EC.visibility_of_element_located((selector_tipo, selector))
        )
        logging.info(f"El elemento con selector {selector} está visible.")
        return True
    except Exception as e:
        logging.info(f"El elemento con selector {selector} no está visible: {e}")
        return False