from django import forms
from .models import File
from captcha.fields import ReCaptchaField
from captcha.widgets import ReCaptchaV3


# Model form
class FileUploadModelForm(forms.ModelForm):
    class Meta:
        model = File
        labels = {
            'file': ''
        }
        fields = ['file']
        widgets = {
            'file': forms.ClearableFileInput(attrs={'class': 'form-control'})
        }

    def clean_file(self):
        file = self.cleaned_data['file']
        ext = file.name.split('.')[-1].lower()
        if ext != "xls":
            raise forms.ValidationError(".xls以外の拡張子ファイルはアップロードいただけません。")
        # return cleaned data is very important.
        return file

    captcha = ReCaptchaField(
        label='',
        widget=ReCaptchaV3(
            attrs={'required_score': 0.7}
        )
    )
