from django.shortcuts import render, redirect
from django.contrib.auth.decorators import login_required
from django.contrib.auth import authenticate, login, logout
from django.contrib import messages
from django.contrib.auth.forms import AuthenticationForm
from django.contrib.auth.models import User, Group

def homepage(request):
    if request.user.is_authenticated:
        is_admin = request.user.is_superuser
        is_applicant = request.user.groups.filter(name='申請人員').exists()
        is_material_handler = request.user.groups.filter(name='撥料人員').exists()
        
        context = {
            'is_admin': is_admin,
            'is_applicant': is_applicant,
            'is_material_handler': is_material_handler,
        }
        return render(request, 'requisitions/homepage.html', context)
    else:
        return render(request, 'requisitions/landing.html')

def user_login(request):
    if request.method == 'POST':
        form = AuthenticationForm(request, data=request.POST)
        if form.is_valid():
            username = form.cleaned_data.get('username')
            password = form.cleaned_data.get('password')
            user = authenticate(username=username, password=password)
            if user is not None:
                login(request, user)
                messages.success(request, f"歡迎回來, {username}!")
                return redirect('homepage')
            else:
                messages.error(request, "無效的使用者名稱或密碼。")
        else:
            messages.error(request, "無效的使用者名稱或密碼。")
    else:
        form = AuthenticationForm()
    return render(request, 'requisitions/login.html', {'form': form})

@login_required
def user_logout(request):
    logout(request)
    messages.info(request, "您已成功登出。")
    return redirect('login')