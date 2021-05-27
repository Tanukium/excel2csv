from django.shortcuts import render, redirect
from .forms import FileUploadModelForm
from .models import File
from converter import excel2csv
import os
import urllib.parse


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
    return render(request, 'file_upload/index.html',
                  {'form': form, 'title': 'Excelファイルをアップロード'})


# Show file list
def file_list(request):
    files = File.objects.all().order_by("-id")
    results, result_sizes, result_paths = [], [], []
    file_names, result_names = [], []
    for file in files:
        result = os.path.splitext(file.file.url)[0] + '.zip'
        results.append(result)

        file_name = urllib.parse.unquote((file.file.url.split('/'))[3])
        file_names.append(file_name)

        result_name = urllib.parse.unquote((result.split('/'))[3])
        result_names.append(result_name)

        path = file.abspath_file()
        result_paths.append(path)
    for path in result_paths:
        path = os.path.splitext(path)[0] + '.zip'
        size = os.path.getsize(path)
        result_sizes.append(size)
    lst = zip(files, results, result_sizes, file_names, result_names)
    return render(request, 'file_upload/list.html',
                  {
                      'files': files,
                      'lst': lst,
                      'title': 'ファイルリスト'
                  })


def delete_file(request, id):
    file_id = id
    delete_files = File.objects.filter(id=file_id)
    for file in delete_files:
        os.remove(file.abspath_file())
        os.remove(os.path.splitext(file.abspath_file())[0] + '.zip')
    File.objects.filter(id=file_id).delete()
    return file_list(request)
