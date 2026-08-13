import os
from datetime import datetime, timedelta

from gmail.client import get_gmail_service
from gmail.messages import search_messages, get_or_create_label
from gmail.attachments import get_message_details, download_attachments
from processors.router import get_processor
from utils.files import save_bytes, ensure_dir


EMAIL_REMITENTE = os.environ.get("EMAIL_REMITENTE", "sebastianss898@gmail.com")


def get_date_prefix():
    yesterday = datetime.now().date() - timedelta(days=1)
    return yesterday.strftime("%Y/%m/%d")


def main():

    service = get_gmail_service()

    messages = search_messages(service, EMAIL_REMITENTE)
    print(f"Correos encontrados: {len(messages)}")

    # Se crean/verifican ambas etiquetas, pero aun no se aplican a
    # ningun correo real -- eso llega en la Fase 3/4, ya con AWS.
    get_or_create_label(service, "PROCESADO")
    get_or_create_label(service, "ERROR")

    date_prefix = get_date_prefix()
    raw_dir = ensure_dir(os.path.join("local_storage", "raw", date_prefix))
    processed_dir = ensure_dir(os.path.join("local_storage", "processed", date_prefix))

    for msg in messages:

        message, subject, sender = get_message_details(service, msg["id"])

        print(f"\nCorreo: {subject}")
        print(f"De: {sender}")

        attachments = download_attachments(service, message)

        if not attachments:
            print("  Sin adjuntos.")
            continue

        all_ok = True

        for att in attachments:

            filename = att["filename"]
            raw_path = save_bytes(att["data"], os.path.join(raw_dir, filename))
            print(f"  Adjunto guardado: {raw_path} ({att['size']} bytes)")

            processor = get_processor(filename)

            if processor is None:
                print(f"    -> No hay procesador para '{filename}', se omite.")
                all_ok = False
                continue

            try:
                resultados = processor(raw_path, processed_dir)

                for r in resultados:
                    print(f"    -> Generado: {r}")

            except Exception as e:
                print(f"    -> ERROR procesando '{filename}': {e}")
                all_ok = False

        if all_ok:
            print("  Todos los adjuntos OK -> se marcaria PROCESADO")
        else:
            print("  Algun adjunto fallo -> se marcaria ERROR")


if __name__ == "__main__":
    main()