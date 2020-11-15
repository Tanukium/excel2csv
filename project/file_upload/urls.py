from django.urls import path
from . import views

# namespace
app_name = "file_upload"
urlpatterns = [
    path('del/<int:id>', views.delete_file, name='delete'),
    # View File List
    path('list/', views.file_list, name='file_list'),
    # Upload Files Using Model Form
    path('', views.model_form_upload, name='model_form_upload'),
]
