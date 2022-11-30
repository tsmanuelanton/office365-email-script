
import graph
import json
import base64

try:
    config = json.load(open("config.json", "r"))
except FileNotFoundError as e:
    print("Falta el archivo de configuración")
    exit(-1)


def main():
    # Autenticar usuario
    graph.initialize_graph_for_user_auth(config)

    emails = graph.get_emails_with_attachments()

    for email in emails["value"]:

        attachments = graph.get_attchments(email["id"])
        for attachment in attachments["value"]:
            f = open(attachment["name"], "wb")
            data_base64 = attachment["contentBytes"]
            f.write(base64.b64decode(data_base64))
            f.close()


main()
