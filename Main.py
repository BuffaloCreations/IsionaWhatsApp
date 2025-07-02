import pandas as pd
import pywhatkit
import time
import random
import re
import os

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# Función para limpiar y formatear número de teléfono
def format_phone_number(phone):
    if pd.isna(phone) or str(phone).strip() == "" or phone == "N/A":
        return None
    cleaned = re.sub(r'\D', '', str(phone))
    if len(cleaned) == 9:
        return "+595" + cleaned
    elif len(cleaned) == 8:
        return "+5959" + cleaned
    else:
        return None

# Ruta del archivo Excel
file_path = "/Users/dominicgossen/Library/CloudStorage/OneDrive-Personal/Isiona/Excel/Resultados_Odontologia.xlsx"

# Leer archivo Excel (hoja por defecto)
df = pd.read_excel(file_path)

print("Columnas encontradas en el Excel:", df.columns.tolist())

# Limpiar y formatear teléfonos
phones = df["Unnamed: 1"].apply(format_phone_number).dropna().unique()

# Archivos de log y estado
log_file_success = "mensajes_enviados.txt"
last_index_file = "ultimo_indice.txt"

# Cargar índice inicial
start_index = 0
if os.path.exists(last_index_file):
    with open(last_index_file, "r") as f:
        try:
            start_index = int(f.read())
        except:
            start_index = 0

# Lista de mensajes
messages = [
    "Hola Doctor, espero que estés bien. Soy Dominic, de Isiona.\nEstoy ayudando a clínicas odontológicas a tener una página web profesional y clara.\nUna buena web te permite atraer nuevos pacientes y mostrar tu trabajo con confianza.\n¿Querés que te cuente más detalles?",
    "¡Buen día Doctor! Te saluda Dominic desde Isiona.\nMe dedico a crear sitios web que reflejan el profesionalismo de clínicas como la tuya.\nCon una página clara y profesional, tus pacientes encuentran más fácil a tu clínica.\nSi te interesa, te paso info sin compromiso.",
    "Hola Doctor, me tomo un momento para escribirte. De paso me presento: soy Dominic, de Isiona.\nTrabajo armando páginas web para odontólogos que buscan destacar online.\nHoy más que nunca, tener presencia online bien cuidada es clave para crecer.\n¿Te gustaría saber cómo lo trabajo?",
    "Buenos días Doctor, ¿cómo viene tu semana? Te saluda Dominic, de Isiona.\nAyudo a clínicas a posicionarse mejor en Google con un sitio web a medida.\nUn sitio web es tu carta de presentación digital: conviene que sea impecable.\nEstoy a disposición si querés más info.",
    "Hola Doctor, espero que estés teniendo un buen día. Soy Dominic, de Isiona.\nDesarrollo sitios web pensados especialmente para clínicas odontológicas modernas.\nMuchos pacientes eligen su odontólogo por lo que ven en internet. Estar bien posicionado ayuda mucho.\n¿Querés ver cómo podría aplicarse a tu clínica?",
    "Hola Doctor, espero que estés bien.\nSoy Dominic, Isiona, y junto con nuestro equipo nos dedicamos a crear sitios web que reflejan el profesionalismo de clínicas como la tuya.\nUna buena web te permite atraer nuevos pacientes y mostrar tu trabajo con confianza.\n¿Querés que te cuente más detalles?",
    "¡Buen día Doctor! Soy Dominic saludando desde Isiona.\nTrabajo armando páginas web para odontólogos que buscan destacar online.\nCon una página clara y profesional, tus pacientes encuentran más fácil a tu clínica.\nSi te interesa, te paso info sin compromiso.",
    "Hola Doctor, me tomo un momento para escribirte. Me presento: soy Dominic, de Isiona.\nAyudo a clínicas a posicionarse mejor en Google con un sitio web a medida.\nHoy más que nunca, tener presencia online bien cuidada es clave para crecer.\n¿Te gustaría saber cómo lo trabajo?",
    "Doctor, ¿cómo viene tu semana? Soy Dominic, de Isiona.\nDesarrollo sitios web pensados especialmente para clínicas odontológicas modernas.\nUn sitio web es tu carta de presentación digital: conviene que sea impecable.\nEstoy a disposición si querés más info.",
    "Hola Doctor, espero que estés teniendo un buen día.\nSoy Dominic de Isiona, y estoy ayudando a clínicas odontológicas a tener una página web profesional y clara.\nMuchos pacientes eligen su odontólogo por lo que ven en internet. Estar bien posicionado ayuda mucho.\n¿Querés ver cómo podría aplicarse a tu clínica?"
]

# Filtro de horarios: solo entre las 9 AM y 6 PM
def is_time_valid():
    current_hour = time.localtime().tm_hour
    return 9 <= current_hour < 18

# Inicializar WebDriver para controlar WhatsApp Web
chrome_options = Options()
chrome_options.add_argument("--user-data-dir=./User_Data")  # Carpeta donde guarda sesión de WhatsApp
chrome_options.add_argument("--profile-directory=Default")
chrome_options.add_argument("--start-maximized")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

# Función para hacer clic en el botón de enviar
def click_send_button():
    try:
        send_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//button[@aria-label="Enviar"]'))
        )
        send_button.click()
    except Exception as e:
        print(f"No se pudo hacer clic en el botón de enviar: {e}")

# Inicializar contadores
sent_count = 0
max_daily_messages = 50

print("Comenzando envío de mensajes...\n")

# Abrir archivos de log para escribir
with open(log_file_success, "a") as success_log, open(last_index_file, "w") as index_log:
    for i in range(start_index, len(phones)):
        number = phones[i]

        if sent_count >= max_daily_messages:
            print("Límite diario alcanzado.")
            index_log.write(str(i))
            break

        if not is_time_valid():
            print("Fuera del horario permitido. Deteniendo...")
            index_log.write(str(i))
            break

        try:
            message = random.choice(messages)
            print(f"Enviando mensaje #{sent_count + 1} a {number}...")

            pywhatkit.sendwhatmsg_instantly(number, message, wait_time=10, tab_close=False)
            time.sleep(5)
            click_send_button()

            success_log.write(f"{number}\n")
            sent_count += 1

            delay = random.randint(30, 60)
            print(f"Esperando {delay} segundos antes del siguiente mensaje...\n")
            time.sleep(delay)

        except Exception as e:
            print(f"Error al enviar a {number}: {e}")
            continue

# Mostrar resumen final
print("\nResumen del envío:")
print(f"Mensajes enviados exitosamente: {sent_count}")
print(f"Próximo inicio desde el índice: {start_index + sent_count}")

# Cerrar navegador de Selenium al final (opcional)
driver.quit()