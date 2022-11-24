import msal
from office365.graph_client import GraphClient
import json

configFile = open("config.json", "r")
config = json.loads(configFile.read())


def acquire_token():
    """
    Acquire token via MSAL
    """
    authority_url = f'https://login.microsoftonline.com/{config["id_inquilino"]}'
    app = msal.ConfidentialClientApplication(
        authority=authority_url,
        client_id=f'{config["client_id"]}',
        client_credential=f'{config["client_credential"]}'
    )
    token = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"])
    return token


client = GraphClient(acquire_token)
