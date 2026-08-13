import os
import tempfile

import config
from gmail.client import get_gmail_service_from_secret
from gmail.messages import search_messages, get_or_create_label, add_label
from gmail.attachments import get_message_details, download_attachments
from processors.router import get_processor
from storage.s3 import upload_file
from utils.files import save_bytes, ensure_dir


def lambda_handler(event, context):

    service = get_gmail_service_from_secret(config.GMAIL_TOKEN_SECRET_NAME)

    messages = search_messages(service, config.EMAIL_REMITENTE)
    print(f"Correos encontrados: {len(messages)}")

    label_procesado_id = get_or_create_label(service, "PROCESADO")
    label_error_id = get_or_create_label(service, "ERROR")

    archivos_generados = []
    correos_con_error = 0

    for msg in messages:

        message, subject, sender = get_message_details(service, msg["id"])
        print(f"Correo: {subject} | De: {sender}")

        attachments = download_attachments(service, message)

        if not attachments:
            print("  Sin adjuntos.")
            continue

        all_ok = True

        # /tmp es el unico directorio con permisos de escritura en Lambda
        with tempfile.TemporaryDirectory(dir="/tmp") as tmp_dir:

            raw_dir = ensure_dir(os.path.join(tmp_dir, "raw"))
            processed_dir = ensure_dir(os.path.join(tmp_dir, "processed"))

            for att in attachments:

                filename = att["filename"]
                raw_path = save_bytes(att["data"], os.path.join(raw_dir, filename))

                processor = get_processor(filename)

                if processor is None:
                    print(f"  No hay procesador para '{filename}', se omite.")
                    all_ok = False
                    continue

                try:
                    generados = processor(raw_path, processed_dir)

                    for path_generado in generados:
                        nombre_salida = os.path.basename(path_generado)
                        key = f"processed/{nombre_salida}"
                        url = upload_file(path_generado, config.S3_BUCKET, key)
                        print(f"  Subido a {url}")
                        archivos_generados.append(url)

                except Exception as e:
                    print(f"  ERROR procesando '{filename}': {e}")
                    all_ok = False

        if all_ok:
            add_label(service, msg["id"], label_procesado_id)
            print("  Correo marcado como PROCESADO")
        else:
            add_label(service, msg["id"], label_error_id)
            correos_con_error += 1
            print("  Correo marcado como ERROR")

    return {
        "correos_encontrados": len(messages),
        "correos_con_error": correos_con_error,
        "archivos_generados": archivos_generados,
    }
