from django.shortcuts import render, redirect
from django.contrib.auth.decorators import login_required
from django.contrib.auth import authenticate, login, logout
from django.contrib import messages
from django.contrib.auth.forms import AuthenticationForm
from django.contrib.auth.models import User, Group

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
                return redirect('core:homepage')
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
    return redirect('requisitions:login')


from django.contrib.auth.forms import UserCreationForm
from django import forms

# 可直接分配的角色
ALLOWED_ROLES = ['撥料人員', '申請人員']
# 需要審核的角色
APPROVAL_REQUIRED_ROLES = ['管理員', '申請人員主管', '撥料人員主管']

from requisitions.models import UserProfile

class UserRegistrationForm(UserCreationForm):
    """自訂註冊表單，包含角色選擇"""
    surname = forms.CharField(label='姓 (Surname)', max_length=30, required=True, widget=forms.TextInput(attrs={'class': 'form-control'}))
    given_name = forms.CharField(label='名 (Given Name)', max_length=30, required=True, widget=forms.TextInput(attrs={'class': 'form-control'}))
    
    role = forms.ChoiceField(
        label='角色',
        choices=[
            ('申請人員', '申請人員'),
            ('撥料人員', '撥料人員'),
            ('申請人員主管', '申請人員主管（需審核）'),
            ('撥料人員主管', '撥料人員主管（需審核）'),
            ('管理員', '管理員（需審核）'),
        ],
        widget=forms.Select(attrs={'class': 'form-control'})
    )
    
    class Meta(UserCreationForm.Meta):
        model = User
        fields = UserCreationForm.Meta.fields + ('surname', 'given_name', 'role',)


def user_register(request):
    """用戶註冊頁面 - 支援角色選擇"""
    if request.user.is_authenticated:
        return redirect('core:homepage')
    
    if request.method == 'POST':
        form = UserRegistrationForm(request.POST)
        if form.is_valid():
            user = form.save(commit=False)
            # 依照用戶要求：First Name 存「姓」，Last Name 存「名」
            user.first_name = form.cleaned_data.get('surname')
            user.last_name = form.cleaned_data.get('given_name')
            user.save()
            
            username = form.cleaned_data.get('username')
            selected_role = form.cleaned_data.get('role')
            
            try:
                group = Group.objects.get(name=selected_role)
                
                if selected_role in ALLOWED_ROLES:
                    # 直接分配角色
                    user.groups.add(group)
                    user.is_active = True
                    user.save()
                    messages.success(request, f"帳號 {username} 已成功建立，角色為「{selected_role}」！請登入。")
                else:
                    # 需要審核：設為非活動狀態
                    user.is_active = False
                    user.save()
                    
                    # 使用 UserProfile 儲存待審核角色
                    UserProfile.objects.create(user=user, requested_role=selected_role)
                    
                    messages.info(request, f"帳號 {username} 已建立，角色「{selected_role}」需管理員審核後才能使用。")
                    
            except Group.DoesNotExist:
                messages.warning(request, f"角色「{selected_role}」不存在，請聯繫管理員。")
            
            return redirect('requisitions:login')
        else:
            messages.error(request, "註冊失敗，請檢查輸入資料。")
    else:
        form = UserRegistrationForm()
    
    return render(request, 'requisitions/register.html', {'form': form})