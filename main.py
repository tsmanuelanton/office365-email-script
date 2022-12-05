
import graph
import json
import base64
from datetime import datetime
import os

try:
    config = json.load(open("config.json", "r"))
except FileNotFoundError as e:
    print("Falta el archivo de configuración")
    exit(-1)

attachmentsDir = config["attachmentsDir"]


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
        if config["filters"]["emailsAfterDate"]:
            emailsAfterDate = datetime.strptime(
                config["filters"]["emailsAfterDate"], "%d/%m/%Y %H:%M:%S")
            filterRecivedTime = emailsAfterDate.strftime('%Y-%m-%dT%H:%M:%SZ')

    if filterRecivedTime:
        print(f"Recuperando archivos adjuntos desde {filterRecivedTime}")
    else:
        print(f"Recuperando archivos adjuntos desde el principio")

    # Obtenemos el id de los correos que tengan archivos adjuntos
    emails = graph.get_emails_with_attachments(
        filterRecivedTime, config["filters"]["fromAddress"])

    if len(emails["value"]) != 0 and config["filters"]["receiversAddresses"]:
        # Descarta aquellos correos que no contengan los receptores indicados en el config
        emails_filtred_by_recipients = []
        for email in emails["value"]:
            email_recipients = []
            for recipient_obj in email["toRecipients"]:
                email_recipients.append(
                    recipient_obj["emailAddress"]["address"])
            if all(recipient in email_recipients for recipient in config["filters"]["receiversAddresses"]):
                emails_filtred_by_recipients.append(email)

        emails = emails_filtred_by_recipients

    if len(emails) == 0:
        print("No hay correos con archivos adjuntos nuevos.")
        return

    # Creamos el directorio donde se van a almacenar los archivos adjuntos
    now = datetime.now()
    date_time = now.strftime("%d-%m-%Y_%H-%M")
    directory_path = f"{attachmentsDir}\\{date_time}"
    if not os.path.exists(directory_path):
        os.makedirs(directory_path)

    for email in emails:
        # Recuperamos los adjuntos del email
        attachments = graph.get_attchments(email["id"])
        for attachment in attachments["value"]:
            # Si no es un archivo (fileAttachment), lo ignormaos
            if attachment["@odata.type"] == "#microsoft.graph.fileAttachment":
                file_name = f'{attachment["name"]}'
                f = open(f"{directory_path}\\{file_name}", "wb")
                data_base64 = attachment["contentBytes"]
                f.write(base64.b64decode(data_base64))
                f.close()


def get_last_time_executed():
    ''' Devuelve un datetime con la última vez que se descargaron ficheros adjuntos'''

    try:
        # Recupera todos los subdirectorios donde almacenamos los archivos adjuntos
        all_subdirs = [d for d in os.listdir(
            attachmentsDir) if os.path.isdir(f'{attachmentsDir}\\{d}')]

        if len(all_subdirs) == 0:
            return None

        def str2ms(str):
            '''Función auxiliar que convierte una fecha en str a mirosegundos'''
            return datetime.strptime(str, '%d-%m-%Y_%H-%M').timestamp()

        # Recorremos todos los subdirectorios buscando la carpeta más reciente
        newest_dir = max(all_subdirs, key=str2ms)

        # Devolvemos el timestamp del directorio más reciente
        return datetime.fromtimestamp(str2ms(newest_dir))
    except FileNotFoundError:
        return None


main()
