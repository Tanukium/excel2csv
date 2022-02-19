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
    return os.path.join("files", filename)


class File(models.Model):
    file = models.FileField(upload_to='files', null=True, validators=[file_size])

    def abspath_file(self):
        root = settings.MEDIA_ROOT
        path = os.path.dirname(self.file.name)
        file = os.path.basename(self.file.name)
        return os.path.join(root, path, file)
