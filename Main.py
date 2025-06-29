import pandas as pd
import pywhatkit
import time
import random
import re
import os


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
file_path = "/Users/macmini/Library/CloudStorage/OneDrive-Personal/Isiona/resultados_odontologia.xlsx"

# Leer archivo Excel (hoja por defecto)
df = pd.read_excel(file_path)

# Limpiar y formatear teléfonos
phones = df["Teléfono"].apply(format_phone_number).dropna().unique()

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

# Tus 10 mensajes personalizados
messages = [
    # Mensaje 1
    """Hola Doc, ¿cómo estás?
Estoy ayudando a clínicas odontológicas a tener una página web profesional y clara.
Una buena web te permite atraer nuevos pacientes y mostrar tu trabajo con confianza.
¿Querés que te cuente más detalles?""",

    # Mensaje 2
    """¡Buen día Doc!
Me dedico a crear sitios web que reflejan el profesionalismo de clínicas como la tuya.
Con una página clara y profesional, tus pacientes encuentran más fácil a tu clínica.
Si te interesa, te paso info sin compromiso.""",

    # Mensaje 3
    """Hola Doc, me tomo un momento para escribirte.
Trabajo armando páginas web para odontólogos que buscan destacar online.
Hoy más que nunca, tener presencia online bien cuidada es clave para crecer.
¿Te gustaría saber cómo lo trabajo?""",

    # Mensaje 4
    """Doc, ¿cómo viene tu semana?
Ayudo a clínicas a posicionarse mejor en Google con un sitio web a medida.
Un sitio web es tu carta de presentación digital: conviene que sea impecable.
Estoy a disposición si querés más info.""",

    # Mensaje 5
    """Hola Doc, espero que estés teniendo un buen día.
Desarrollo sitios web pensados especialmente para clínicas odontológicas modernas.
Muchos pacientes eligen su odontólogo por lo que ven en internet. Estar bien posicionado ayuda mucho.
¿Querés ver cómo podría aplicarse a tu clínica?""",

    # Mensaje 6
    """Hola Doc, ¿cómo estás?
Me dedico a crear sitios web que reflejan el profesionalismo de clínicas como la tuya.
Una buena web te permite atraer nuevos pacientes y mostrar tu trabajo con confianza.
¿Querés que te cuente más detalles?""",

    # Mensaje 7
    """¡Buen día Doc!
Trabajo armando páginas web para odontólogos que buscan destacar online.
Con una página clara y profesional, tus pacientes encuentran más fácil a tu clínica.
Si te interesa, te paso info sin compromiso.""",

    # Mensaje 8
    """Hola Doc, me tomo un momento para escribirte.
Ayudo a clínicas a posicionarse mejor en Google con un sitio web a medida.
Hoy más que nunca, tener presencia online bien cuidada es clave para crecer.
¿Te gustaría saber cómo lo trabajo?""",

    # Mensaje 9
    """Doc, ¿cómo viene tu semana?
Desarrollo sitios web pensados especialmente para clínicas odontológicas modernas.
Un sitio web es tu carta de presentación digital: conviene que sea impecable.
Estoy a disposición si querés más info.""",

    # Mensaje 10
    """Hola Doc, espero que estés teniendo un buen día.
Estoy ayudando a clínicas odontológicas a tener una página web profesional y clara.
Muchos pacientes eligen su odontólogo por lo que ven en internet. Estar bien posicionado ayuda mucho.
¿Querés ver cómo podría aplicarse a tu clínica?"""
]


# Filtro de horarios: solo entre las 9 AM y 6 PM
def is_time_valid():
    current_hour = time.localtime().tm_hour
    return 9 <= current_hour < 18


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
            index_log.write(str(i))  # Guardar posición actual
            break

        if not is_time_valid():
            print("Fuera del horario permitido. Deteniendo...")
            index_log.write(str(i))  # Guardar posición actual
            break

        try:
            message = random.choice(messages)  # Seleccionar mensaje aleatorio
            print(f"Enviando mensaje #{sent_count + 1} a {number}...")
            pywhatkit.sendwhatmsg_instantly(number, message, wait_time=10, tab_close=True)
            success_log.write(f"{number}\n")
            sent_count += 1
            delay = random.randint(30, 60)  # Entre 30 y 60 segundos
            print(f"Esperando {delay} segundos antes del siguiente mensaje...\n")
            time.sleep(delay)
        except Exception as e:
            print(f"Error al enviar a {number}: {e}")
            continue

# Mostrar resumen final
print("\nResumen del envío:")
print(f"Mensajes enviados exitosamente: {sent_count}")
print(f"Próximo inicio desde el índice: {start_index + sent_count}")