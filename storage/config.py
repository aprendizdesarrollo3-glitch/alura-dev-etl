import os

# Remitente cuyos correos activan el procesamiento
EMAIL_REMITENTE = os.environ.get(
    "EMAIL_REMITENTE", "tatiana.gonzalez@grupobios.co"
)

# Bucket S3 donde quedan los Excel limpios generados
S3_BUCKET = os.environ.get("S3_BUCKET", "gmail-etl-opav-814383264230")

# Nombre del secreto en AWS Secrets Manager que contiene el token.json
# generado en la Fase 1 (credenciales OAuth ya autorizadas)
GMAIL_TOKEN_SECRET_NAME = os.environ.get(
    "GMAIL_TOKEN_SECRET_NAME", "gmail-etl/token"
)