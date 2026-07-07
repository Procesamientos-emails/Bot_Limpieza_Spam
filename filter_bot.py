import os
import json
import imaplib
import logging
import urllib.request
import urllib.parse
from datetime import datetime

# Configuración del log
log_format = "%(asctime)s - %(levelname)s - %(message)s"
logging.basicConfig(
    level=logging.INFO,
    format=log_format,
    handlers=[
        logging.FileHandler(os.path.join(os.path.dirname(__file__), "bot.log"), encoding="utf-8"),
        logging.StreamHandler()
    ]
)

def find_spam_folder(imap, default_folder):
    """
    Intenta descubrir la carpeta de correo no deseado (Spam/Junk) utilizando
    los atributos estándar de IMAP (RFC 6154). Si no se detecta, usa el valor por defecto.
    """
    try:
        status, folder_list = imap.list()
        if status != 'OK':
            return default_folder
            
        for folder_info in folder_list:
            info_str = folder_info.decode('utf-8', errors='ignore')
            parts = info_str.split(' ")" ', 1)
            if len(parts) < 2:
                parts = info_str.split(' "/" ', 1)
            
            if len(parts) == 2:
                attributes = parts[0].lower()
                folder_name = parts[1].strip('"')
                
                # Comprobar si tiene el atributo \spam o \junk
                if '\\spam' in attributes or '\\junk' in attributes:
                    logging.info(f"Detectada carpeta de Spam por atributos IMAP: '{folder_name}'")
                    return folder_name
        
        # Búsqueda manual por nombres comunes en inglés y español
        common_names = ["spam", "junk", "correo no deseado", "correo_no_deseado", "bulk"]
        for folder_info in folder_list:
            info_str = folder_info.decode('utf-8', errors='ignore')
            parts = info_str.split(' "/" ', 1)
            if len(parts) == 2:
                folder_name = parts[1].strip('"')
                if any(name in folder_name.lower() for name in common_names):
                    logging.info(f"Detectada carpeta de Spam por nombre común: '{folder_name}'")
                    return folder_name

    except Exception as e:
        logging.warning(f"Error al listar carpetas IMAP, se usará la carpeta por defecto: {e}")
        
    logging.info(f"Usando carpeta por defecto: '{default_folder}'")
    return default_folder

def get_microsoft_oauth2_token(client_id, refresh_token):
    """
    Usa el Refresh Token para obtener un nuevo Access Token de Microsoft de forma segura.
    """
    token_url = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
    data = {
        "client_id": client_id,
        "scope": "https://outlook.office.com/IMAP.AccessAsUser.All offline_access",
        "refresh_token": refresh_token,
        "grant_type": "refresh_token"
    }
    encoded_data = urllib.parse.urlencode(data).encode("utf-8")
    req = urllib.request.Request(token_url, data=encoded_data, method="POST")
    req.add_header("Content-Type", "application/x-www-form-urlencoded")
    
    with urllib.request.urlopen(req) as response:
        res_body = response.read().decode("utf-8")
        tokens = json.loads(res_body)
        return tokens.get("access_token")

def clean_spam_account(account_config):
    email_addr = account_config.get("email")
    provider = account_config.get("provider", "desconocido").upper()
    password = account_config.get("password")
    client_id = account_config.get("client_id")
    refresh_token = account_config.get("refresh_token")
    imap_server = account_config.get("imap_server")
    imap_port = account_config.get("imap_port", 993)
    default_folder = account_config.get("default_folder", "Spam")

    logging.info(f"Iniciando limpieza para {provider}: {email_addr}")

    use_oauth2 = client_id is not None and refresh_token is not None

    if not email_addr or (not password and not use_oauth2) or not imap_server:
        logging.error(f"Configuración incompleta para la cuenta {email_addr}. Verifica config.json.")
        return

    # Evitar ejecutar con los valores de ejemplo
    if not use_oauth2 and ("tu_contraseña" in password or "tu_correo" in email_addr):
        logging.warning(f"La cuenta {email_addr} aún tiene los datos de ejemplo de la plantilla. Omitiendo...")
        return

    imap = None
    try:
        # Conexión segura SSL
        logging.info(f"Conectando a {imap_server}:{imap_port}...")
        imap = imaplib.IMAP4_SSL(imap_server, imap_port)
        
        # Autenticación
        if use_oauth2:
            logging.info("Solicitando nuevo token de acceso OAuth2 a Microsoft...")
            access_token = get_microsoft_oauth2_token(client_id, refresh_token)
            logging.info("Autenticando con mecanismo XOAUTH2...")
            # Formato de la cadena XOAUTH2: user=<user>\x01auth=Bearer <token>\x01\x01
            auth_string = f"user={email_addr}\x01auth=Bearer {access_token}\x01\x01"
            imap.authenticate('XOAUTH2', lambda x: auth_string.encode('utf-8'))
        else:
            logging.info("Autenticando con contraseña...")
            imap.login(email_addr, password)
        
        # Buscar la carpeta de Spam
        spam_folder = find_spam_folder(imap, default_folder)
        
        # Seleccionar la carpeta en modo lectura/escritura
        logging.info(f"Seleccionando carpeta '{spam_folder}'...")
        status, data = imap.select(f'"{spam_folder}"', readonly=False)
        if status != 'OK':
            status, data = imap.select(spam_folder, readonly=False)
            if status != 'OK':
                raise Exception(f"No se pudo seleccionar la carpeta '{spam_folder}': {data}")

        # Buscar todos los correos
        status, messages = imap.search(None, 'ALL')
        if status != 'OK':
            raise Exception(f"Error al buscar correos: {messages}")

        mail_ids = messages[0].split()
        total_emails = len(mail_ids)
        logging.info(f"Se encontraron {total_emails} correos en la bandeja de Spam.")

        if total_emails > 0:
            logging.info("Marcando correos para eliminación permanente...")
            for mail_id in mail_ids:
                imap.store(mail_id, '+FLAGS', '\\Deleted')
            
            logging.info("Ejecutando eliminación física (EXPUNGE)...")
            status, data = imap.expunge()
            if status == 'OK':
                logging.info(f"ÉXITO: Se eliminaron permanentemente {total_emails} correos de Spam en {email_addr}.")
            else:
                logging.error(f"Error al ejecutar expunge: {data}")
        else:
            logging.info("La carpeta de Spam ya está limpia. No se requiere ninguna acción.")

    except imaplib.IMAP4.error as imap_err:
        logging.error(f"Error de protocolo IMAP para {email_addr}: {imap_err}. Verifica tus credenciales.")
    except Exception as e:
        logging.error(f"Error inesperado procesando {email_addr}: {e}")
    finally:
        if imap:
            try:
                imap.logout()
                logging.info("Conexión cerrada de forma limpia.")
            except:
                pass
        logging.info("-" * 50)

def main():
    config_path = os.path.join(os.path.dirname(__file__), "config.json")
    if not os.path.exists(config_path):
        logging.error(f"No se encontró el archivo de configuración en {config_path}")
        return

    try:
        with open(config_path, "r", encoding="utf-8") as f:
            config = json.load(f)
    except Exception as e:
        logging.error(f"Error al leer config.json: {e}")
        return

    accounts = config.get("accounts", [])
    if not accounts:
        logging.warning("No hay cuentas configuradas en config.json.")
        return

    logging.info("=== BOT DE LIMPIEZA DE SPAM INICIADO ===")
    for account in accounts:
        clean_spam_account(account)
    logging.info("=== PROCESO DE LIMPIEZA COMPLETADO ===")

if __name__ == "__main__":
    main()
