#################################################################################################################################
import imaplib
import email
from email.header import decode_header
import os
import zipfile


# Directorio donde se guardarán los archivos adjuntos
download_folder = "archivos_adjuntos"

# Verificar si el directorio existe, si no existe, crearlo
if not os.path.exists(download_folder):
    os.makedirs(download_folder)


# Datos de la cuenta de correo
username = 'reportes@zayro.com'
password = 'Rexe=Aco85'
remite = "kmc_ace_system@khi.co.jp"
#remite = 'programacion@zayro.com'
# Configuración del servidor IMAP
imap_server = 'vmail.globalpc.net'  
port = 993  # El puerto IMAP seguro estándar es 993 para IMAP sobre SSL/TLS

# Directorio donde se guardarán los archivos adjuntos
download_folder = "archivos_adjuntos"
# Conexión al servidor IMAP
print("Conectándose al servidor IMAP...")
imap = imaplib.IMAP4_SSL(imap_server, port)
print("Conexión establecida.")

# Iniciar sesión
print("Iniciando sesión...")
login_status = imap.login(username, password)
print(login_status)
if login_status[0] == 'OK':
    print("Sesión iniciada correctamente.")
    session_active = True
else:
    print("Error al iniciar sesión.")
    session_active = False

if session_active:
    # Seleccionar la bandeja de entrada
    select_status = imap.select("inbox")
    print("Resultado de la selección de la bandeja de entrada:", select_status)
    if select_status[0] == 'OK':
        print("Bandeja de entrada seleccionada.")
    else:
        print("Error al seleccionar la bandeja de entrada.")

    # Buscar los correos electrónicos del remitente específico
    result, data = imap.search(None, f'(FROM "{remite}")')
    if result == 'OK':
        print("Correos electrónicos encontrados.")
        # Obtener los IDs de los correos electrónicos encontrados
        email_ids = data[0].split()
        if email_ids:
            # Iterar sobre todos los IDs de correo electrónico encontrados
            for email_id in email_ids:
                # Obtener los datos del correo electrónico actual
                result, data = imap.fetch(email_id, "(RFC822)")
                if result == 'OK':
                    raw_email = data[0][1]
                    # Parsear el correo electrónico
                    print("Parseando el correo electrónico...")
                    msg = email.message_from_bytes(raw_email)

                    # Obtener el remitente y el asunto del correo electrónico
                    from_ = msg["From"]
                    subject = msg["Subject"]

                    # Decodificar el asunto si es necesario
                    subject = decode_header(subject)[0][0]
                    if isinstance(subject, bytes):
                        subject = subject.decode()

                    print("De:", from_)
                    print("Asunto:", subject)

                    # Imprimir el contenido del correo electrónico
                    print("Contenido del correo electrónico:")
                    for part in msg.walk():
                        if part.get_content_type() == "text/plain":
                            try:
                                body = part.get_payload(decode=True).decode('utf-8')
                            except UnicodeDecodeError:
                                print("Error al decodificar el contenido del correo electrónico.")
                                continue  # Pasar al siguiente correo electrónico
                            print(body)
                        elif part.get_content_maintype() == 'application' and part.get('Content-Disposition') is not None:
                            # Descargar y descomprimir archivos adjuntos
                            filename = part.get_filename()
                            if filename:
                                print("Descargando archivo adjunto:", filename)
                                filepath = os.path.join(download_folder, filename)
                                with open(filepath, 'wb') as f:
                                    f.write(part.get_payload(decode=True))
                                if filename.lower().endswith('.zip'):
                                    print("Descomprimiendo archivo adjunto:", filename)
                                    with zipfile.ZipFile(filepath, 'r') as zip_ref:
                                        zip_ref.extractall(download_folder)
                                    # Eliminar el archivo comprimido después de descomprimirlo
                                    os.remove(filepath)
                                print("Archivo adjunto guardado en:", filepath)

                else:
                    print("Error al obtener datos del correo electrónico con ID:", email_id)

        else:
            print("No se encontraron correos electrónicos del remitente específico.")
    else:
        print("Error al buscar correos electrónicos del remitente específico.")

    # Cerrar la conexión
    print("Cerrando la conexión...")
    imap.close()
    print("Conexión cerrada.")
    print("Sesión finalizada.")
    imap.logout()
else:
    print("No se pudo realizar ninguna acción debido a un error en la sesión.")
#################################################################################################################################