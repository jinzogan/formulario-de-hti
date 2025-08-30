from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from flask import Flask, request, render_template, redirect, url_for, flash
from werkzeug.utils import secure_filename
import os
import time
import pandas as pd
import tempfile
import threading
import logging

# =======================
# CONFIGURACIÓN DEL DRIVER
# =======================

# Initialize driver and wait as None - will be created when needed
driver = None
wait = None

app = Flask(__name__)
app.secret_key = os.environ.get("SESSION_SECRET") or "dev-secret-key"

# Configura la carpeta donde se guardarán los archivos subidos
app.config['UPLOAD_FOLDER'] = 'uploads/'
app.config['ALLOWED_EXTENSIONS'] = {'xlsx', 'xls'}

# Asegúrate de que la carpeta exista
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)


def allowed_file(filename):
    return '.' in filename and filename.rsplit(
        '.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    # Verifica si el archivo es parte de la solicitud
    if 'file' not in request.files:
        return "No file part"

    file = request.files['file']

    # Si no seleccionaron un archivo
    if file.filename == '':
        return "No selected file"

    # Si el archivo es válido
    if file and file.filename and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)

        # Lee el archivo Excel y inicia procesamiento en segundo plano
        try:
            datos_excel = pd.read_excel(file_path)
            lista_datos = datos_excel.to_dict(orient='records')
            
            # Iniciar el procesamiento en un hilo separado
            def procesar_en_segundo_plano():
                usuarios_procesados = 0
                usuarios_fallidos = 0
                
                try:
                    for datos in lista_datos:
                        print(f"Procesando usuario: {datos.get('usuario', 'desconocido')}")
                        
                        try:
                            # Ejecutar la automatización para este usuario
                            iniciar_sesion(datos)
                            time.sleep(2)
                            
                            navegar_a_enlace()
                            completar_formulario()
                            time.sleep(1)
                            
                            formulario(datos)
                            
                            cerrar_sesion()
                            usuarios_procesados += 1
                            print(f"✓ Usuario {datos.get('usuario', 'desconocido')} procesado exitosamente")
                            
                        except Exception as user_error:
                            usuarios_fallidos += 1
                            print(f"✗ Error procesando usuario {datos.get('usuario', 'desconocido')}: {user_error}")
                            try:
                                cerrar_sesion()  # Intentar cerrar sesión en caso de error
                            except:
                                pass
                            continue
                    
                    print(f"✓ PROCESAMIENTO COMPLETADO: {usuarios_procesados} exitosos, {usuarios_fallidos} fallidos")
                    
                except Exception as e:
                    print(f"✗ Error general en procesamiento: {e}")
                
                finally:
                    # Cerrar el driver al final
                    global driver
                    try:
                        if driver:
                            driver.quit()
                            driver = None
                    except:
                        pass
                    
                    # Limpiar archivos temporales de Firefox
                    try:
                        os.system("rm -rf /tmp/rust_mozprofile* /tmp/tmp* 2>/dev/null")
                        print("✓ Perfiles temporales de Firefox eliminados")
                    except:
                        pass
                    
                    # Eliminar el archivo Excel
                    try:
                        os.remove(file_path)
                        print(f"✓ Archivo Excel eliminado: {filename}")
                    except Exception as cleanup_error:
                        print(f"Advertencia: No se pudo eliminar el archivo: {cleanup_error}")
            
            # Iniciar el procesamiento en segundo plano
            hilo_procesamiento = threading.Thread(target=procesar_en_segundo_plano)
            hilo_procesamiento.daemon = True  # Permite que el programa termine aunque el hilo esté corriendo
            hilo_procesamiento.start()
            
            return f"Archivo Excel cargado exitosamente. Procesando {len(lista_datos)} usuario(s) en segundo plano. Revisa los logs del servidor para ver el progreso."
            
        except Exception as e:
            # En caso de error, intentar eliminar el archivo
            try:
                os.remove(file_path)
            except:
                pass
            return f"Error al procesar el archivo: {e}"

    return "Archivo no permitido. Solo se permiten archivos .xlsx o .xls"


if __name__ == '__main__':
    app.run(debug=True)

# Driver initialization function
def initialize_driver():
    global driver, wait
    if driver is None:
        options = webdriver.FirefoxOptions()
        options.add_argument('--headless')  # 👉 Oculta el navegador
        options.add_argument('--no-sandbox')  # Necesario en entornos de servidor
        options.add_argument('--disable-dev-shm-usage')  # Evita problemas de memoria
        # Firefox no necesita configuración especial de binary_location
        
        # Configuración para Firefox - deshabilitar notificaciones y prompts
        options.set_preference("dom.webnotifications.enabled", False)
        options.set_preference("dom.push.enabled", False)
        options.set_preference("signon.rememberSignons", False)
        options.set_preference("browser.privatebrowsing.autostart", True)
        
        # Firefox maneja los perfiles temporales automáticamente en modo headless
        
        # Usar GeckoDriver del sistema
        try:
            service = Service()  # Usar geckodriver del PATH
            driver = webdriver.Firefox(service=service, options=options)
            print("✓ Firefox iniciado exitosamente")
        except Exception as firefox_error:
            print(f"Error iniciando Firefox: {firefox_error}")
            raise
        wait = WebDriverWait(driver, 30)  # Más tiempo para elementos lentos

# reiniciar


def cerrar_sesion():
    global driver
    if driver is None:
        initialize_driver()
    try:
        # Intenta encontrar y hacer clic en el botón de cerrar sesión
        driver.get('https://personal.migracion.gob.do/Account/Logout')
        print("Sesión cerrada.")
    except:
        print("No se pudo cerrar sesión (quizás no estás logueado).")


def reiniciar_navegador():
    global driver, wait  # Necesario para cambiar el driver global
    try:
        if driver:
            driver.quit()
    except:
        pass

    time.sleep(2)
    driver = None
    wait = None
    initialize_driver()
    print("Navegador reiniciado.")


# Limpiar campos del formulario
def limpiar_campos(lista_selectores):
    global driver
    if driver is None:
        initialize_driver()
    for selector in lista_selectores:
        try:
            campo = driver.find_element(By.CSS_SELECTOR, selector)
            campo.clear()
            campo.send_keys(Keys.CONTROL + "a")
            campo.send_keys(Keys.DELETE)
        except Exception as e:
            print(f"No se pudo limpiar el campo: {selector}. Error: {e}")


def iniciar_sesion(datos):
    global driver, wait
    if driver is None:
        initialize_driver()
    driver.get('https://personal.migracion.gob.do/Account/Login')
    print("Página de login cargada.")

    # Esperar que cargue el formulario de inicio de sesión
    wait.until(EC.presence_of_element_located((By.NAME, 'id')))

    # Llenar usuario y contraseña
    driver.find_element(By.NAME, 'id').send_keys(datos['usuario'])
    driver.find_element(By.NAME, 'password').send_keys(datos['contra'])

    # Clic en botón de inicio de sesión
    driver.find_element(
        By.XPATH,
        '//input[@type="submit" and @value="Iniciar Sesión"]').click()
    print("Intentando iniciar sesión...")

    # Esperar que se cargue la página de bienvenida (ajusta el selector si es diferente)
    wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'welcome')))
    print("Sesión iniciada correctamente.")

    # Cerrar popup si aparece
    try:
        wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, '//button[contains(text(), "Aceptar")]'))).click()
        print("Popup de cambio de contraseña cerrado.")
    except TimeoutException:
        print("No apareció popup de cambio de contraseña.")


def navegar_a_enlace():
    global driver, wait
    if driver is None:
        initialize_driver()
    # Cambia el texto si el link es distinto
    enlace = wait.until(
        EC.element_to_be_clickable((By.LINK_TEXT, 'LISTA DE APLICACIONES')))
    enlace.click()
    print("Navegando a sección LISTA DE APLICACIONES...")

    # Esperar a que cargue tabla o indicador de que la página cargó
    wait.until(EC.presence_of_element_located((By.ID, 'tblLinks')))
    print("Sección del formulario cargada.")


def completar_formulario():
    global driver, wait
    if driver is None:
        initialize_driver()
    xpath = "//td[a[contains(translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyzáéíóúü', 'ABCDEFGHIJKLMNOPQRSTUVWXYZÁÉÍÓÚÜ'), 'SOLICITUD DE RENOVACIÓN CARNET DE TRABAJADORES TEMPOREROS')]]/following-sibling::td/input[@type='image']"

    try:
        input_element = wait.until(
            EC.element_to_be_clickable((By.XPATH, xpath)))

        # Scroll al elemento para asegurarnos que esté visible
        driver.execute_script(
            "arguments[0].scrollIntoView({block: 'center'});", input_element)
        time.sleep(0.5)

        input_element.click()
        print(
            "Click realizado en el formulario de renovación carnet de trabajadores temporeros."
        )

    except (TimeoutException, NoSuchElementException) as e:
        print("No se pudo encontrar o hacer clic en el elemento:", e)
    except Exception as e:
        print("Error al hacer click, intentando con JavaScript:", e)
        try:
            driver.execute_script("arguments[0].click();", input_element)
            print("Click realizado con JavaScript.")
        except Exception as js_e:
            print("También falló el click con JavaScript:", js_e)

    # Esperar que cargue la siguiente página o sección
    try:
        wait.until(
            EC.visibility_of_element_located((By.CLASS_NAME, "rc-wrap-t")))
        print("Formulario cargado correctamente.")

    except TimeoutException:
        print("No se pudo confirmar la carga de la siguiente página.")

    boton_aplicar = wait.until(
        EC.presence_of_element_located((
            By.XPATH,
            "//div[contains(@class, 'f-right')]//button[contains(text(), 'Aplicar')]"
        )))

    # Scroll para asegurarte que esté visible
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});",
                          boton_aplicar)

    # Click con JavaScript
    driver.execute_script("arguments[0].click();", boton_aplicar)

    # Espera a que el formulario cargue
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located(
            (By.ID,
             'nombre'))  # Espera hasta que el campo 'nombre' esté disponible
    )

    # Seleccionar sede con mejor manejo
    try:
        sede = driver.find_element(By.ID, "radio_button_3")
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", sede)
        time.sleep(0.5)
        sede.click()
        print("✓ Sede seleccionada")
    except Exception as e:
        print(f"✗ Error seleccionando sede: {e}")


    # Rellenar el formulario
def formulario(datos):
    global driver, wait
    if driver is None:
        initialize_driver()

    # Nombre, apellido, fecha de nacimiento y pasaporte hecho en el correo
    pasaporte1_input = driver.find_element(
        By.ID, 'p_tipo')  # Cambia el 'name' por el real
    pasaporte1_input.send_keys('p')

    pasaporte2_input = driver.find_element(
        By.ID, 'p_fecha_expedicion')  # Cambia el 'name' por el real
    pasaporte2_input.send_keys(datos['expedicion'])

    pasaporte3_input = driver.find_element(
        By.ID, 'p_fecha_expiracion')  # Cambia el 'name' por el real
    pasaporte3_input.send_keys(datos['expiracion'])

    pasaporte4_input = driver.find_element(
        By.ID, 'p_pais_emisor')  # Cambia el 'name' por el real
    pasaporte4_input.send_keys('hti')

    # Visa4
    visa = driver.find_element(By.ID, 'v_no')  # Cambia el 'name' por el real
    visa.send_keys('11111111111')

    visa1 = driver.find_element(By.ID,
                                'v_tipo')  # Cambia el 'name' por el real
    visa1.send_keys('p')

    visa2 = driver.find_element(
        By.ID, 'v_fecha_expedicion')  # Cambia el 'name' por el real
    visa2.send_keys(datos['expedicion'])

    visa3 = driver.find_element(
        By.ID, 'v_fecha_expiracion')  # Cambia el 'name' por el real
    visa3.send_keys(datos['expiracion'])

    # carnet4
    carnet = driver.find_element(By.ID, 'e_no')  # Cambia el 'name' por el real
    carnet.send_keys(datos['e_no'])

    carnet1 = driver.find_element(By.ID,
                                  'e_tipo')  # Cambia el 'name' por el real
    carnet1.send_keys('tt1')

    carnet2 = driver.find_element(
        By.ID, 'e_fecha_expedicion')  # Cambia el 'name' por el real
    carnet2.send_keys(datos['e_expedicion'])

    carnet3 = driver.find_element(
        By.ID, 'e_fecha_expiracion')  # Cambia el 'name' por el real
    carnet3.send_keys(datos['e_expiracion'])

    # varios - teléfonos con mejor manejo
    numero = '111-111-1111'
    try:
        telefono_element = driver.find_element(By.ID, 'telefonos1')
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", telefono_element)
        time.sleep(0.3)
        driver.execute_script(
            "document.getElementById('telefonos1').value = arguments[0];", numero)
        print("✓ Teléfono 1 rellenado")
    except Exception as e:
        print(f"✗ Error con teléfono 1: {e}")

    driver.execute_script(
        """
    var telefonoInput = document.getElementById('telefonos1');
    telefonoInput.value = arguments[0];
    telefonoInput.dispatchEvent(new Event('input', { bubbles: true }));
    telefonoInput.dispatchEvent(new Event('change', { bubbles: true }));
""", numero)

    # Teléfono celular
    try:
        celular_element = driver.find_element(By.ID, 'celular')
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", celular_element)
        time.sleep(0.3)
        driver.execute_script(
            "document.getElementById('celular').value = arguments[0];", numero)
        print("✓ Celular rellenado")
    except Exception as e:
        print(f"✗ Error con celular: {e}")

    driver.execute_script(
        """
    var telefonoInput = document.getElementById('celular');
    telefonoInput.value = arguments[0];
    telefonoInput.dispatchEvent(new Event('input', { bubbles: true }));
    telefonoInput.dispatchEvent(new Event('change', { bubbles: true }));
""", numero)

    # direccion - con mejor manejo de elementos
    try:
        municipio_element = driver.find_element(By.ID, 'municipio')
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", municipio_element)
        time.sleep(0.5)
        provincia = Select(municipio_element)
        provincia.select_by_value('01')
        print("✓ Municipio seleccionado")
    except Exception as e:
        print(f"✗ Error seleccionando municipio: {e}")

    WebDriverWait(driver,
                  10).until(EC.presence_of_element_located((By.ID, "sector")))

    # Seleccionar sector - con mejor manejo
    try:
        sector_element = driver.find_element(By.ID, "sector")
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", sector_element)
        time.sleep(0.5)
        sector_select = Select(sector_element)
        sector_select.select_by_visible_text("Santo Domingo de Guzmán")
        print("✓ Sector seleccionado")
    except Exception as e:
        print(f"✗ Error seleccionando sector: {e}")

    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "newsector")))

    # Esperar que el <option> deseado esté presente
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located(
            (By.XPATH,
             '//select[@id="newsector"]/option[@value="03100101010106500"]')))

    # Seleccionar newsector - con mejor manejo
    try:
        select_element = driver.find_element(By.ID, "newsector")
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", select_element)
        time.sleep(0.5)
        valor = "03100101010106500"
        driver.execute_script(
            "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('change'));",
            select_element, valor)
        print("✓ Newsector seleccionado")
    except Exception as e:
        print(f"✗ Error seleccionando newsector: {e}")

    # Calle con mejor manejo
    try:
        calle = driver.find_element(By.ID, 'calle')
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", calle)
        time.sleep(0.3)
        calle.clear()
        calle.send_keys('arzobispo portes')
        print("✓ Calle rellenada")
    except Exception as e:
        print(f"✗ Error con calle: {e}")

    campo = driver.find_element(By.ID, 'salario')
    valor = (datos['salario'])  # Asegúrate de que sea string

    driver.execute_script("arguments[0].value = arguments[1];", campo, valor)

    empleador = driver.find_element(
        By.ID, 'empleador_dominicado')  # Cambia el 'name' por el real
    empleador.send_keys(datos['profesion'])

    # Enviar el formulario 1

    boton_enviar = wait.until(
        EC.presence_of_element_located((
            By.XPATH,
            "//div[contains(@class, 'f-right')]//button[contains(text(), 'Siguiente')]"
        )))

    # Scroll para asegurarte que esté visible
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});",
                          boton_enviar)

    # Click con JavaScript
    driver.execute_script("arguments[0].click();", boton_enviar)

    # Esperar a que cargue el segundo formulario (puedes ajustar el ID más representativo si existe)
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, 'sobre_empresa')))

    # Rellenar los campos con datos por defecto
    empresa = driver.find_element(By.ID, 'sobre_empresa')
    empresa.send_keys(datos['empresa'])

    rnc = driver.find_element(By.ID, 'rnc')
    rnc.send_keys(datos['rnc'])

    tsocietario = driver.find_element(By.ID, 'tipo_societario')
    tsocietario.send_keys(datos['societario'])

    personaacargo = driver.find_element(By.ID, 'persona_cargo')
    personaacargo.send_keys(datos['profesion'])

    # Dar click en el botón "Siguiente"
    boton_siguiente = wait.until(
        EC.presence_of_element_located((
            By.XPATH,
            "//div[contains(@class, 'f-right')]//button[contains(text(), 'Siguiente')]"
        )))

    # Scroll para asegurarte que esté visible
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});",
                          boton_siguiente)

    # Click con JavaScript
    driver.execute_script("arguments[0].click();", boton_siguiente)

    # Esperar a que cargue el segundo formulario (puedes ajustar el ID más representativo si existe)

    # Espera un poco para asegurarse de que el formulario se haya enviado
    time.sleep(3)
    print("Formulario enviado correctamente")
    print("completado, no te preocupes por lo que sigue.")


# Paso 5: Ejecutar el flujo completo
def ejecutar():
    global driver, wait
    if driver is None:
        initialize_driver()
    # Leer datos desde el archivo Excel
    datos_excel = pd.read_excel('datos_usuarios.xlsx')
    lista_datos = datos_excel.to_dict(orient='records')

    for datos in lista_datos:
        try:
            iniciar_sesion(datos)
            time.sleep(2)

            # Aquí puedes navegar si hace falta
            # driver.get("https://...")

            navegar_a_enlace()

            try:
                completar_formulario()
            except Exception as e:
                print(
                    f"[ERROR] Fallo llenando formulario para: {datos['usuario']}"
                )
                print(f"→ Error: {e}")
                cerrar_sesion()
                continue  # Saltar al siguiente usuario

            time.sleep(1)

            # Intentar hacer clic en "Siguiente"
            formulario(datos)

            try:

                siguiente_btn = driver.find_element(
                    By.XPATH, '//button[contains(text(), "Siguiente")]')
                siguiente_btn.click()
            except:
                print("Botón 'Siguiente' no encontrado.")

            # Cerrar sesión y reiniciar para el próximo usuario
            cerrar_sesion()

        except Exception as e:
            print(
                f"[ERROR] Fallo con el usuario: {datos.get('usuario', 'desconocido')}"
            )
            print(f"→ Error: {e}")

    try:
        driver.quit()
    except:
        pass


# Note: ejecutar() is called only when needed, not at module import
