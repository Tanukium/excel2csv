from django.urls import re_path, path
from . import views

# namespace
app_name = "file_upload"
urlpatterns = [
    # View File List
    path('file_list.html', views.file_list, name='file_list'),
    # Upload Files Using Model Form
    path('', views.model_form_upload, name='model_form_upload'),
]
