
import graph
import json
import warnings
warnings.filterwarnings(
    action='ignore', module='.*azure.*')

try:
    config = json.load(open("config.json", "r"))
except FileNotFoundError as e:
    print("Falta el archivo de configuración")
    exit(-1)


def main():
    graph.initialize_graph_for_user_auth(config)

    list_inbox()


def list_inbox():
    message_page = graph.get_inbox()

    # Output each message's details
    for message in message_page['value']:
        print('Message:', message['subject'])
        print('  From:', message['from']['emailAddress']['name'])
        print('  Status:', 'Read' if message['isRead'] else 'Unread')
        print('  Received:', message['receivedDateTime'])

    # If @odata.nextLink is present
    more_available = '@odata.nextLink' in message_page
    print('\nMore messages available?', more_available, '\n')


main()
