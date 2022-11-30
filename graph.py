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


def get_user_token():
    access_token = this.device_code_credential.get_token(
        this.settings['graphUserScopes'])
    return access_token.token


def get_inbox():
    endpoint = '/me/mailFolders/inbox/messages'
    # Solo obtener estos campos específicos
    select = 'from,isRead,receivedDateTime,subject'
    # Obtener un máximo de 25 resultados
    top = 25
    request_url = f'{endpoint}?$select={select}&$top={top}'
    inbox_response = this.user_client.get(request_url)
    return inbox_response.json()
