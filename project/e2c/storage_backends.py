from storages.backends.s3boto3 import S3Boto3Storage


# 同ファイル名での上書きを許さない
class MediaStorage(S3Boto3Storage):
    location = 'media'
    file_overwrite = False
