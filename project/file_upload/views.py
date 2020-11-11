from django.shortcuts import render, redirect
from .forms import FileUploadModelForm
from .models import File
from django.template.defaultfilters import filesizeformat
from converter import excel2csv
import os


# Create your views here.

# Upload File with ModelForm
def model_form_upload(request):
    if request.method == "POST":
        form = FileUploadModelForm(request.POST,
                                   request.FILES)
        if form.is_valid():
            f = form.save()
            e2c = excel2csv.Converter(f.abspath_file())
            e2c.output_csv_files()
            e2c.pack_csv_files()
            return redirect("/upload/list/")
    else:
        form = FileUploadModelForm()
    return render(request, 'file_upload/index.html', {'form': form, 'title': 'Excelファイルをアップロード'})


# Show file list
def file_list(request):
    files = File.objects.all().order_by("-id")
    results = []
    for file in files:
        result = os.path.splitext(file.file.url)[0] + '.zip'
        results.append(result)
    lst = zip(files, results)
    return render(request, 'file_upload/list.html', {'files': files, 'lst': lst})
