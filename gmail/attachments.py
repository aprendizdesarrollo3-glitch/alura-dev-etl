import base64


def get_message_details(service, message_id):

    message = service.users().messages().get(
        userId="me",
        id=message_id,
        format="full"
    ).execute()

    headers = message["payload"].get("headers", [])

    subject = next(
        (h["value"] for h in headers if h["name"] == "Subject"),
        "(sin asunto)"
    )

    sender = next(
        (h["value"] for h in headers if h["name"] == "From"),
        "(desconocido)"
    )

    return message, subject, sender


def download_attachments(service, message):

    attachments = []

    parts = message["payload"].get("parts", []) or []

    for part in parts:

        filename = part.get("filename")
        body = part.get("body", {})

        if not filename:
            continue

        if "attachmentId" in body:
            attachment = service.users().messages().attachments().get(
                userId="me",
                messageId=message["id"],
                id=body["attachmentId"]
            ).execute()
            raw_data = attachment["data"]

        elif "data" in body:
            raw_data = body["data"]

        else:
            continue

        file_data = base64.urlsafe_b64decode(raw_data.encode("UTF-8"))

        attachments.append({
            "filename": filename,
            "data": file_data,
            "size": len(file_data)
        })

    return attachments
