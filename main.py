
import graph
import json
import base64
from datetime import datetime
import os
from pathlib import Path
import logging

# Creando un logger personalizado
logger = logging.getLogger("office365-email-script")
logger.setLevel(logging.INFO)
formatter = logging.Formatter(
    '%(asctime)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s')
c_handler = logging.StreamHandler()
c_handler.setFormatter(formatter)
logger.addHandler(c_handler)

try:
    parent_directory = Path(__file__).parent.resolve()
    config_file_path = Path(f"{parent_directory}/config.json")
    config = json.load(open(config_file_path, "r"))
except FileNotFoundError as e:
    logger.error("Falta el archivo de configuración.")
    exit(-1)

attachmentsDir = config.get("attachmentsDir")
if not attachmentsDir:
    logger.error("Falta definir \"attachmentsDir\" en config.json: " +
                 "Hay que especificar la ubicación para los archivos adjuntos.")
    exit(-1)


def main():
    # Autenticar usuario
    graph.initialize_graph_for_user_auth(config["azure"])

    # Obtenemos, si existe la fecha de la última ejecución
    last_time = get_last_time_executed()

    filterRecivedTime = None
    # Si no es la primera vez que se ejecuta el script, filtramos los correos
    # para obtener sólo aquellos posteriores a la última ejecución
    if last_time != None:
        filterRecivedTime = last_time.strftime(
            '%Y-%m-%dT%H:%M:%SZ')
    else:
        if config.get("filters").get("emailsAfterDate"):
            emailsAfterDate = datetime.strptime(
                config["filters"]["emailsAfterDate"], "%d/%m/%Y %H:%M:%S")
            filterRecivedTime = emailsAfterDate.strftime('%Y-%m-%dT%H:%M:%SZ')

    logger.info(
        f'Recuperando archivos adjuntos desde {filterRecivedTime if filterRecivedTime else "el principio"}.')

    # Obtenemos el id de los correos que tienen archivos adjuntos
    fromAddres = config.get("filters").get("fromAddress")
    emails = graph.get_emails_with_attachments(
        filterRecivedTime, fromAddres)

    if emails.get("error"):
        error_msg = emails.get("error").get("message")
        logger.error(f"Se ha producido un error: {error_msg}")
        return -1
    emails = emails["value"]

    if len(emails) != 0 and config.get("filters").get("receiversAddresses"):
        # Descarta aquellos correos que no contengan los receptores indicados en el config
        emails = get_emails_filtred(
            emails, config["filters"]["receiversAddresses"])

    if len(emails) == 0:
        logger.info("No hay correos con archivos adjuntos nuevos.")
        return

    # Creamos el directorio donde se van a almacenar los archivos adjuntos
    now = datetime.utcnow()
    date_time = now.strftime("%d-%m-%Y_%H-%M")
    directory_path = Path(f"{attachmentsDir}/{date_time}")
    if not os.path.exists(directory_path):
        logger.debug(
            f"Creando directorio para almacenar archivos adjuntos en {directory_path}")
        os.makedirs(directory_path)

    # Almacenamos los archivos adjuntos
    n_attachments_saved = save_attachments(emails, directory_path)
    logger.info(
        f"Almacenados {n_attachments_saved} archivos en {directory_path}.")


def get_last_time_executed():
    ''' Devuelve un datetime con la última vez que se descargaron ficheros adjuntos'''

    try:
        # Recupera todos los subdirectorios donde almacenamos los archivos adjuntos
        all_subdirs = [d for d in os.listdir(
            attachmentsDir) if os.path.isdir(Path(f'{attachmentsDir}/{d}'))]

        if len(all_subdirs) == 0:
            return None

        def str2ms(str):
            '''Función auxiliar que convierte una fecha en str a microsegundos'''
            return datetime.strptime(str, '%d-%m-%Y_%H-%M').timestamp()

        # Recorremos todos los subdirectorios buscando la carpeta más reciente
        newest_dir = max(all_subdirs, key=str2ms)

        # Devolvemos el timestamp del directorio más reciente
        return datetime.fromtimestamp(str2ms(newest_dir))
    except FileNotFoundError:
        return None


def get_emails_filtred(emails, recipent_filter):
    # Descarta aquellos correos que no contengan los receptores indicados en el config
    emails_filtred_by_recipients = []
    for email in emails:
        email_recipients = []
        for recipient_obj in email["toRecipients"]:
            email_recipients.append(
                recipient_obj["emailAddress"]["address"])
        if all(recipient in email_recipients for recipient in recipent_filter):
            emails_filtred_by_recipients.append(email)

    return emails_filtred_by_recipients


def save_attachments(emails, directory_path):
    n_attachments_saved = 0

    for email in emails:
        # Recuperamos los adjuntos del email
        attachments = graph.get_attchments(email["id"])
        if attachments.get("error"):
            error_msg = attachments.get("error").get("message")
            logger.error(f"Se ha producido un error: {error_msg}")
            return -1
        for attachment in attachments["value"]:
            # Si no es un archivo (fileAttachment), lo ignormaos
            if attachment["@odata.type"] == "#microsoft.graph.fileAttachment":
                file_name = f'{attachment["name"]}'
                data_base64 = attachment["contentBytes"]
                f = open(Path(f"{directory_path}/{file_name}"), "wb")
                f.write(base64.b64decode(data_base64))
                f.close()
                n_attachments_saved += 1

    return n_attachments_saved


try:
    logger.info("Inicio ejecución script.")
    main()
except BaseException as e:
    logger.exception(f"Se ha producido un error: {e}")

logger.info("Fin ejecución script.")
