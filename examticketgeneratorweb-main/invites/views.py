from django.contrib.auth.models import User
from .forms import RegisterForm
from django.shortcuts import render
from .models import Invite

def register(request):
    if request.method == "POST":
        form = RegisterForm(request.POST)
        if form.is_valid():
            invite_code = form.cleaned_data['invite_code']
            username = form.cleaned_data['username']
            password = form.cleaned_data['password']

            # Проверка, что инвайт существует и не был использован
            try:
                invite = Invite.objects.get(code=invite_code, used=False)
            except Invite.DoesNotExist:
                form.add_error('invite_code', 'Неверный или использованный код приглашения')
                return render(request, 'invites/register.html', {'form': form})

            # Создание пользователя
            user = User.objects.create_user(username=username, password=password)
            user.save()

            # Отметить инвайт как использованный
            invite.used = True
            invite.save()

            # (Возможно, автоматически залогинить или отправить на страницу логина)
            return render(request, 'invites/register_success.html', {'username': username})

    else:
        form = RegisterForm()
    return render(request, 'invites/register.html', {'form': form})
