from django.shortcuts import render, redirect
from .forms import FileUploadModelForm
from .models import File
from django.template.defaultfilters import filesizeformat
from excel2csv import excel2csv
import os

# Create your views here.

# Upload File with ModelForm
def model_form_upload(request):
    if request.method == "POST":
        form = FileUploadModelForm(request.POST,
                                   request.FILES)
        if form.is_valid():
            f = form.save()
            e2c = excel2csv.Excel2csv(f.fullfilename())
            e2c.output_csv_files()
            return redirect("/upload/file_list.html")
    else:
        form = FileUploadModelForm()
    return render(request, 'file_upload/index.html', {'form': form, 'title': 'Excelファイルをアップロード'})

# Show file list
def file_list(request):
    files = File.objects.all().order_by("-id")
    return render(request, 'file_upload/file_list.html', {'files': files})
