import boto3
import io

try:
    from .local_settings import *
except ImportError:
    pass


def bind_bucket_and_client(bucket_name):
    s3 = boto3.resource('s3')
    bucket = s3.Bucket(bucket_name)
    client = boto3.client('s3')
    return bucket, client


def receive_xls_from_bucket(bucket_name, xls_name):
    bucket, client = bind_bucket_and_client(bucket_name)
    obj = bucket.Object(xls_name)
    xls_on_memory = io.BytesIO()
    obj.download_fileobj(xls_on_memory)
    xls = xls_on_memory.getvalue()
    xls_on_memory.close()
    return xls


def upload_zip_to_bucket(bucket_name, zip_data, key):
    bucket, client = bind_bucket_and_client(bucket_name)
    response = client.put_object(
        Body=zip_data,
        Bucket=bucket_name,
        Key=key
    )
    return response


def delete_obj_from_bucket(bucket_name, key):
    bucket, client = bind_bucket_and_client(bucket_name)
    response = client.delete_object(Bucket=bucket_name, Key=key)
    return response


def return_size_of_obj(bucket_name, key):
    bucket, client = bind_bucket_and_client(bucket_name)
    response = client.head_object(Bucket=bucket_name, Key=key)
    size = response['ContentLength']
    return size
