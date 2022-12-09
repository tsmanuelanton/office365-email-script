import sys
from azure.identity import UsernamePasswordCredential
from msgraph.core import GraphClient

this = sys.modules[__name__]


def initialize_graph_for_user_auth(config):
    client_id = config['clientId']
    username = config['username']
    password = config['password']

    this.device_code_credential = UsernamePasswordCredential(client_id,
                                                             username, password)
    this.user_client = GraphClient(
        credential=this.device_code_credential, scopes=["mail.read"])


def get_emails_with_attachments(filterRecivedTime, fromAddressFilter):
    '''Obtiene los IDs de aquellos correos que tengan archivos adjuntos'''
    endpoint = '/me/messages'

    filter = "hasAttachments eq true"

    if filterRecivedTime:
        filter += f" AND receivedDateTime ge {filterRecivedTime}"
    if fromAddressFilter:
        filter += f" AND (from/emailAddress/address) eq '{fromAddressFilter}'"

    # Obtenemos los campos id y los receptores del correo
    select = 'id, toRecipients'
    request_url = f'{endpoint}?$filter={filter}&$select={select}'
    inbox_response = this.user_client.get(request_url)
    return inbox_response.json()


def get_attchments(id_email):
    '''Obtiene los archivos adjuntos del email con id pasado por parámetros'''
    endpoint = f'/me/messages/{id_email}/attachments'
    request_url = f'{endpoint}'
    response = this.user_client.get(request_url)
    return response.json()
