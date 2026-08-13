import boto3


def upload_file(local_path, bucket, key):
    """
    Sube un archivo local a S3 y devuelve su ruta s3://bucket/key.
    """
    s3 = boto3.client("s3")
    s3.upload_file(local_path, bucket, key)

    return f"s3://{bucket}/{key}"