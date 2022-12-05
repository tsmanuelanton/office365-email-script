import sys
from azure.identity import DeviceCodeCredential
from msgraph.core import GraphClient

this = sys.modules[__name__]


def initialize_graph_for_user_auth(config):
    this.settings = config
    client_id = this.settings['clientId']
    tenant_id = this.settings['authTenant']
    graph_scopes = this.settings['graphUserScopes'].split(' ')

    this.device_code_credential = DeviceCodeCredential(
        client_id, tenant_id=tenant_id)
    this.user_client = GraphClient(
        credential=this.device_code_credential, scopes=graph_scopes)


def get_emails_with_attachments(filterRecivedTime, fromAddressFilter):
    '''Obtiene los IDs de aquellos correos que tengan archivos adjuntos'''
    endpoint = '/me/messages'

    filter = "hasAttachments eq true"

    if filterRecivedTime:
        filter += f"receivedDateTime ge {filterRecivedTime}"
        if fromAddressFilter:
            filter += " AND "
    if fromAddressFilter:
        filter += f"(from/emailAddress/address) eq '{fromAddressFilter}'"

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
