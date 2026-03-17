from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.utils import timezone
from .models import PeerRequest
from .forms import PeerRequestForm, PeerReplyForm
from django.db.models import Q

@login_required
def peer_request_list(request):
    # 我發起的申請 (最新 5 筆)
    sent_requests_qs = PeerRequest.objects.filter(applicant=request.user).order_by('-created_at')
    sent_requests = sent_requests_qs[:5]
    has_more_sent = sent_requests_qs.exists()
    
    # 我收到的需求 (最新 5 筆)
    received_requests_qs = PeerRequest.objects.filter(recipient=request.user).order_by('-created_at')
    received_requests = received_requests_qs[:5]
    has_more_received = received_requests_qs.exists()
    
    form = PeerRequestForm()
    reply_form = PeerReplyForm()
    
    context = {
        'sent_requests': sent_requests,
        'received_requests': received_requests,
        'has_more_sent': has_more_sent,
        'has_more_received': has_more_received,
        'form': form,
        'reply_form': reply_form,
    }
    return render(request, 'peer_requests/list.html', context)

@login_required
def peer_request_history(request):
    # 顯示所有相關紀錄
    sent_requests = PeerRequest.objects.filter(applicant=request.user).order_by('-created_at')
    received_requests = PeerRequest.objects.filter(recipient=request.user).order_by('-created_at')
    
    context = {
        'sent_requests': sent_requests,
        'received_requests': received_requests,
        'reply_form': PeerReplyForm(),
    }
    return render(request, 'peer_requests/history.html', context)

@login_required
def peer_request_create(request):
    if request.method == 'POST':
        form = PeerRequestForm(request.POST, request.FILES)
        if form.is_valid():
            peer_request = form.save(commit=False)
            peer_request.applicant = request.user
            peer_request.save()
            messages.success(request, "申請已成功送出！")
        else:
            messages.error(request, "申請發起失敗，請檢查輸入內容。")
    return redirect('peer_requests:list')

@login_required
def peer_request_reply(request, pk):
    peer_request = get_object_or_404(PeerRequest, pk=pk, recipient=request.user)
    if request.method == 'POST':
        form = PeerReplyForm(request.POST, request.FILES, instance=peer_request)
        if form.is_valid():
            peer_request = form.save(commit=False)
            peer_request.status = 'shipped'
            peer_request.save()
            messages.success(request, "已成功回覆撥料資訊！")
        else:
            messages.error(request, "回覆失敗，請檢查內容。")
    return redirect(request.META.get('HTTP_REFERER', 'peer_requests:list'))

@login_required
def peer_request_close(request, pk):
    peer_request = get_object_or_404(PeerRequest, pk=pk, applicant=request.user)
    if request.method == 'POST':
        peer_request.status = 'closed'
        peer_request.closed_at = timezone.now()
        peer_request.save()
        messages.success(request, "申請已結案，確認收貨成功！")
    return redirect(request.META.get('HTTP_REFERER', 'peer_requests:list'))
