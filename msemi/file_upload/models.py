from django.db import models
import os
from django.conf import settings

# Create your models here.
# Define user directory path


def user_directory_path(instance, filename):
    return os.path.join("files", filename)


class File(models.Model):
    file = models.FileField(upload_to=user_directory_path, null=True)
    
    def fullfilename(self):
        a = settings.MEDIA_ROOT
        b = os.path.dirname(self.file.name)
        c = os.path.basename(self.file.name)
        return os.path.join(a, b, c)