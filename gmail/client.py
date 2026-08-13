import os
import json

import boto3
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build


SCOPES = [
    "https://www.googleapis.com/auth/gmail.modify"
]


def get_gmail_service():

    creds = None

    token_path = "token.json"
    credentials_path = "credentials.json"

    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(
            token_path,
            SCOPES
        )

    if not creds or not creds.valid:

        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())

        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                credentials_path,
                SCOPES
            )

            creds = flow.run_local_server(port=0)

        with open(token_path, "w") as token:
            token.write(creds.to_json())

    return build(
        "gmail",
        "v1",
        credentials=creds
    )


def get_gmail_service_from_secret(secret_name):
    """
    Version para usar dentro de Lambda: no hay navegador disponible, asi
    que el token (generado una vez en local con get_gmail_service) se lee
    desde AWS Secrets Manager. Si el access_token expiro, se refresca con
    el refresh_token y el secreto se actualiza automaticamente.

    El secreto debe contener el mismo JSON que produce creds.to_json()
    en la Fase 1 (el contenido de token.json).
    """

    client = boto3.client("secretsmanager")

    response = client.get_secret_value(SecretId=secret_name)
    token_data = json.loads(response["SecretString"])

    creds = Credentials.from_authorized_user_info(token_data, SCOPES)

    if not creds.valid:

        if creds.expired and creds.refresh_token:
            creds.refresh(Request())

            # Persistimos el access_token renovado para no gastar
            # refresh_token en cada invocacion.
            client.put_secret_value(
                SecretId=secret_name,
                SecretString=creds.to_json()
            )

        else:
            raise RuntimeError(
                "El token de Gmail en Secrets Manager no es valido y no "
                "se pudo refrescar. Genera uno nuevo localmente (Fase 1, "
                "borrando token.json) y vuelve a subirlo al secreto "
                f"'{secret_name}'."
            )

    return build(
        "gmail",
        "v1",
        credentials=creds
    )
