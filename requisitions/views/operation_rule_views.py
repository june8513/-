from django.shortcuts import render, redirect
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.http import JsonResponse
from ..models import OperationProcessRule
from ..constants import PROCESS_CATEGORIES


@login_required
def classify_operations(request):
    """讓用戶為未知作業說明選擇投料點"""
    unknown_operations = request.session.get('unknown_operations', [])
    
    if not unknown_operations:
        messages.info(request, "沒有需要分類的作業說明。")
        return redirect('core:homepage')
    
    process_types = [name for name, _ in PROCESS_CATEGORIES]
    
    if request.method == 'POST':
        saved_count = 0
        for op in unknown_operations:
            selected_type = request.POST.get(f'process_type_{op}')
            if selected_type:
                OperationProcessRule.objects.update_or_create(
                    operation_description=op,
                    defaults={
                        'process_type': selected_type,
                        'updated_by': request.user
                    }
                )
                saved_count += 1
        
        # 清除 session
        if 'unknown_operations' in request.session:
            del request.session['unknown_operations']
        
        messages.success(request, f"已儲存 {saved_count} 筆作業說明投料點規則。")
        return redirect('core:homepage')
    
    context = {
        'unknown_operations': unknown_operations,
        'process_types': process_types,
    }
    return render(request, 'requisitions/classify_operations.html', context)


@login_required
def operation_rules_list(request):
    """查看所有作業說明投料點規則"""
    rules = OperationProcessRule.objects.all()
    process_types = [name for name, _ in PROCESS_CATEGORIES]
    
    if request.method == 'POST':
        # 處理更新或刪除
        action = request.POST.get('action')
        rule_id = request.POST.get('rule_id')
        
        if action == 'delete' and rule_id:
            OperationProcessRule.objects.filter(pk=rule_id).delete()
            messages.success(request, "規則已刪除。")
        elif action == 'update' and rule_id:
            new_type = request.POST.get('process_type')
            if new_type:
                OperationProcessRule.objects.filter(pk=rule_id).update(
                    process_type=new_type,
                    updated_by=request.user
                )
                messages.success(request, "規則已更新。")
        
        return redirect('requisitions:operation_rules_list')
    
    context = {
        'rules': rules,
        'process_types': process_types,
    }
    return render(request, 'requisitions/operation_rules_list.html', context)
