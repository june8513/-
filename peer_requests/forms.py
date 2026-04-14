from django import forms
from .models import PeerRequest
from django.contrib.auth.models import User

class PeerRequestForm(forms.ModelForm):
    recipient = forms.ModelChoiceField(
        queryset=User.objects.all(),
        label="收件人",
        widget=forms.Select(attrs={'class': 'form-select select2-recipient'})
    )
    
    cc_users = forms.ModelMultipleChoiceField(
        queryset=User.objects.all(),
        label="副本通知 (選填)",
        required=False,
        widget=forms.SelectMultiple(attrs={'class': 'form-select select2-cc'})
    )


    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Customize the labels in the recipient dropdown
        self.fields['recipient'].label_from_instance = lambda obj: f"{obj.last_name}{obj.first_name}" if obj.last_name or obj.first_name else obj.username
        self.fields['cc_users'].label_from_instance = lambda obj: f"{obj.last_name}{obj.first_name}" if obj.last_name or obj.first_name else obj.username

    request_date = forms.DateField(
        label="需求日期",
        widget=forms.DateInput(attrs={'class': 'form-control', 'type': 'date'})
    )
    description = forms.CharField(
        label="需求內容",
        widget=forms.Textarea(attrs={'class': 'form-control', 'rows': 3, 'placeholder': '請輸入需求描述或貼上照片...'})
    )
    request_photo = forms.ImageField(
        label="需求照片",
        required=False,
        widget=forms.FileInput(attrs={'class': 'form-control'})
    )

    class Meta:
        model = PeerRequest
        fields = ['recipient', 'cc_users', 'description', 'request_photo', 'request_date']

from django.utils import timezone

class PeerReplyForm(forms.ModelForm):
    delivery_date = forms.DateField(
        label="撥料日期",
        widget=forms.DateInput(attrs={'class': 'form-control', 'type': 'date'}),
        initial=timezone.localdate
    )
    delivery_photo = forms.ImageField(
        label="撥料照片",
        required=False,
        widget=forms.FileInput(attrs={'class': 'form-control'})
    )
    delivery_reply = forms.CharField(
        label="備註/回覆",
        required=False,
        widget=forms.Textarea(attrs={'class': 'form-control', 'rows': 2})
    )

    class Meta:
        model = PeerRequest
        fields = ['delivery_date', 'delivery_photo', 'delivery_reply']


class PeerAcceptForm(forms.ModelForm):
    expected_delivery_date = forms.DateField(
        label="預計完成日期",
        widget=forms.DateInput(attrs={'class': 'form-control', 'type': 'date'})
    )
    delivery_reply = forms.CharField(
        label="備註/初步回覆 (選填)",
        required=False,
        widget=forms.Textarea(attrs={'class': 'form-control', 'rows': 2, 'placeholder': '可先留言給申請人...'})
    )

    class Meta:
        model = PeerRequest
        fields = ['expected_delivery_date', 'delivery_reply']
