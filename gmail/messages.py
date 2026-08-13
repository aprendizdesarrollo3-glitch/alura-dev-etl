from datetime import datetime, timedelta


def get_yesterday_range():

    today = datetime.now().date()
    yesterday = today - timedelta(days=1)

    start = yesterday.strftime("%Y/%m/%d")
    end = today.strftime("%Y/%m/%d")

    return start, end


def get_today_range():
    """
    SOLO PARA PRUEBAS: rango de hoy (00:00 de hoy -> 00:00 de manana).
    Usar mientras se prueba el pipeline con correos enviados el mismo dia.
    Volver a get_yesterday_range() antes de pasar a produccion.
    """

    today = datetime.now().date()
    tomorrow = today + timedelta(days=1)

    start = today.strftime("%Y/%m/%d")
    end = tomorrow.strftime("%Y/%m/%d")

    return start, end


def search_messages(service, sender):

    # TEMPORAL (pruebas): usando el rango de HOY en vez de ayer.
    # Cuando termines de probar, cambia get_today_range() por
    # get_yesterday_range() para volver al comportamiento de produccion.
    start, end = get_today_range()

    query = (
        f"from:{sender} "
        f"after:{start} "
        f"before:{end} "
        f"-label:PROCESADO"
    )

    messages = []

    response = service.users().messages().list(
        userId="me",
        q=query,
        maxResults=100
    ).execute()

    messages.extend(
        response.get("messages", [])
    )

    while "nextPageToken" in response:

        response = service.users().messages().list(
            userId="me",
            q=query,
            maxResults=100,
            pageToken=response["nextPageToken"]
        ).execute()

        messages.extend(
            response.get("messages", [])
        )

    return messages


def get_or_create_label(service, label_name):

    response = service.users().labels().list(
        userId="me"
    ).execute()

    labels = response.get("labels", [])

    for label in labels:

        if label["name"] == label_name:
            return label["id"]

    label = service.users().labels().create(
        userId="me",
        body={
            "name": label_name
        }
    ).execute()

    return label["id"]


def add_label(service, message_id, label_id):

    service.users().messages().modify(
        userId="me",
        id=message_id,
        body={
            "addLabelIds": [label_id]
        }
    ).execute()