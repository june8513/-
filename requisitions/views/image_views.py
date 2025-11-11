from django.shortcuts import render, redirect, get_object_or_404
from django.http import HttpResponse
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.utils import timezone
from ..forms import RequisitionImageForm, WorkOrderMaterialImageUploadForm
from ..models import Requisition, RequisitionImage, ProcessType

@login_required
def view_requisition_images(request, pk):
    requisition = get_object_or_404(Requisition, pk=pk)

    is_admin = request.user.is_superuser
    is_applicant = request.user == requisition.applicant
    is_material_handler = request.user.groups.filter(name='撥料人員').exists()

    if not (is_admin or is_applicant or is_material_handler):
        messages.error(request, "您沒有權限訪問此頁面。")
        return redirect('homepage')

    images = requisition.images.all()
    return render(request, 'requisitions/view_requisition_images.html', {'requisition': requisition, 'images': images})

@login_required
def upload_requisition_images_page(request, pk):
    requisition = get_object_or_404(Requisition, pk=pk)

    is_admin = request.user.is_superuser
    is_applicant = request.user == requisition.applicant
    is_material_handler = request.user.groups.filter(name='撥料人員').exists()

    if not (is_admin or is_applicant or is_material_handler):
        messages.error(request, "您沒有權限訪問此頁面。")
        return redirect('homepage')

    form = RequisitionImageForm()
    return render(request, 'requisitions/upload_requisition_images.html', {'requisition': requisition, 'form': form})

@login_required
def upload_dispatch_note_image(request, pk):
    requisition = get_object_or_404(Requisition, pk=pk)
    if request.method == 'POST':
        form = WorkOrderMaterialImageUploadForm(request.POST, request.FILES)
        if form.is_valid():
            image_instance = form.save(commit=False)
            image_instance.requisition = requisition
            image_instance.uploaded_by = request.user
            
            # Try to get the process type object
            try:
                process_type_obj = ProcessType.objects.get(name=requisition.process_type)
                image_instance.process_type = process_type_obj
            except ProcessType.DoesNotExist:
                # Handle case where process type name doesn't match any object
                # You might want to log this or handle it differently
                pass

            image_instance.save()
            messages.success(request, "圖片上傳成功。")
        else:
            messages.error(request, "上傳失敗，請確認已選擇檔案。")
    return redirect('generate_dispatch_note', pk=requisition.pk)

@login_required
def upload_work_order_material_images(request, pk):
    requisition = get_object_or_404(Requisition, pk=pk)

    is_admin = request.user.is_superuser
    is_applicant = request.user == requisition.applicant
    is_material_handler = request.user.groups.filter(name='撥料人員').exists()

    if not (is_admin or is_applicant or is_material_handler):
        messages.error(request, "您沒有權限訪問此頁面。")
        return redirect('homepage')

    # Placeholder function
    return HttpResponse("This is a placeholder for upload_work_order_material_images.")

@login_required
def upload_requisition_images(request, pk):
    requisition = get_object_or_404(Requisition, pk=pk)

    is_admin = request.user.is_superuser
    is_applicant = request.user == requisition.applicant
    is_material_handler = request.user.groups.filter(name='撥料人員').exists()

    if not (is_admin or is_applicant or is_material_handler):
        messages.error(request, "您沒有權限訪問此頁面。")
        return redirect('homepage')

    if request.method == 'POST':
        form = RequisitionImageForm(request.POST, request.FILES)
        if form.is_valid():
            image_instance = form.save(commit=False)
            image_instance.requisition = requisition
            image_instance.uploaded_by = request.user
            image_instance.save()
            messages.success(request, "圖片上傳成功。")
        else:
            messages.error(request, "上傳失敗，請確認已選擇檔案。")
    return redirect('view_requisition_images', pk=requisition.pk)

@login_required
def view_work_order_material_images(request, material_code):
    # Placeholder function
    return HttpResponse("This is a placeholder for view_work_order_material_images.")