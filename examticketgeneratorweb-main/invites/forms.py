from django import forms

class RegisterForm(forms.Form):
    invite_code = forms.CharField(label='Код приглашения', max_length=64)
    username = forms.CharField(label='Логин', max_length=150)
    password = forms.CharField(label='Пароль', widget=forms.PasswordInput)
