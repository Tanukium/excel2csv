from django.shortcuts import render, redirect
from .forms import FileUploadModelForm
from .models import File
from e2c import s3_method
from converter import excel2csv


# Create your views here.

# Upload File with ModelForm
def model_form_upload(request):
    if request.method == "POST":
        form = FileUploadModelForm(request.POST,
                                   request.FILES)
        if form.is_valid():
            f = form.save()
            key = "media/" + f.file.name
            print(key)
            bucket_name = s3_method.AWS_STORAGE_BUCKET_NAME
            xls_buffer = s3_method.receive_xls_from_bucket(bucket_name, key)
            xls_converter = excel2csv.Converter(xls_buffer, bucket_name)
            zipped_csv_buffer = xls_converter.pack_csv_files()
            key = key.rstrip('.xls') + '.zip'
            zip_response = s3_method.upload_zip_to_bucket(bucket_name, zipped_csv_buffer, key)
            return redirect("/upload/list/")
    else:
        form = FileUploadModelForm()
    return render(request, 'file_upload/index.html',
                  {'form': form, 'title': 'Excelファイルをアップロード'})


# Show file list
def file_list(request):
    files = File.objects.all().order_by("-id")
    zip_names, zip_urls, zip_sizes = [], [], []
    bucket_name = s3_method.AWS_STORAGE_BUCKET_NAME
    for file in files:
        zip_name = file.file.name.rstrip(".xls") + ".zip"
        zip_names.append(zip_name)

        zip_key = "media/" + zip_name
        zip_url = file.file.url.rstrip(".xls") + ".zip"
        zip_urls.append(zip_url)

        zip_size = s3_method.return_size_of_obj(bucket_name, zip_key)
        zip_sizes.append(zip_size)
    lst = zip(zip_names, zip_urls, zip_sizes, files)
    return render(request, 'file_upload/list.html',
                  {
                      'lst': lst,
                      'title': 'ファイルリスト'
                  })


def delete_file(request, id):
    file_id = id
    delete_files = File.objects.filter(id=file_id)
    for file in delete_files:
        xls_key = "media/" + file.file.name
        zip_key = ("media/" + file.file.name).rstrip(".xls") + ".zip"
        bucket_name = s3_method.AWS_STORAGE_BUCKET_NAME
        xls_del_response = s3_method.delete_obj_from_bucket(bucket_name, xls_key)
        zip_del_response = s3_method.delete_obj_from_bucket(bucket_name, zip_key)
    File.objects.filter(id=file_id).delete()
    return file_list(request)
