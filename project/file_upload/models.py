from django.db import models
import os
from django.conf import settings
from django.core.exceptions import ValidationError

# Create your models here.
# Define user directory path


def file_size(value):
    limit = 524000
    if value.size > limit:
        raise ValidationError('File too large. Size should not exceed 500KiB.')


def user_directory_path(instance, filename):
    return os.path.join("", filename)


class File(models.Model):
    file = models.FileField(upload_to='', null=True, validators=[file_size])

