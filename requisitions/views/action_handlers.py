from django.http import JsonResponse, HttpResponse
from django.db import models
from django.db.models import Sum, Q, F, ExpressionWrapper, DecimalField
from django.db.models.functions import Coalesce
from decimal import Decimal
from django.contrib.auth.decorators import login_required
from django.utils import timezone
import pandas as pd
import io
import re
from datetime import date, datetime

from requisitions.models import Requisition, WorkOrderMaterial # Import WorkOrderMaterial
from inventory.models import MaterialTransaction, Material

# Helper function to apply filters to Django QuerySets
def _apply_filters_to_queryset(queryset, filters):
    q_objects = Q()
    model = queryset.model
    
    for key, value in filters.items():
        # Core feature: is_shortage
        if key == 'is_shortage' and value is True:
            if model == WorkOrderMaterial:
                queryset = queryset.annotate(
                    temp_confirmed=Coalesce(F('confirmed_quantity'), Decimal('0.00'))
                ).filter(required_quantity__gt=F('temp_confirmed'))
                continue
            elif model == Requisition:
                q_objects &= Q(items__dispatch_status='backordered')
                continue

        # Check if we should strip __date from DateFields to avoid Django error
        clean_key = key
        if '__date' in key:
            field_name = key.split('__')[0]
            try:
                field = model._meta.get_field(field_name)
                # If the field is a DateField (not DateTimeField), __date is not supported
                if isinstance(field, models.DateField) and not isinstance(field, models.DateTimeField):
                    clean_key = key.replace('__date', '')
                    print(f"DEBUG: Stripped __date from {key} because {field_name} is DateField")
            except Exception:
                pass

        # Clean key for joined fields
        field_base = clean_key.split('__')[0]
        try:
            model._meta.get_field(field_base)
        except Exception:
            print(f"DEBUG: Skipping invalid field {field_base} (key: {key}) for model {model.__name__}")
            continue

        # Safety net: WorkOrderMaterial does NOT have a 'status' or 'applicant' field
        if model == WorkOrderMaterial and (key in ['status', 'applicant', 'applicant_id', 'user_id'] or key.startswith('status__')):
             print(f"DEBUG: Stripped invalid field {key} from WorkOrderMaterial query")
             continue

        # Clean date values
        field_name_for_val = clean_key.split('__')[0]
        try:
            field_obj = model._meta.get_field(field_name_for_val)
            is_date_field = isinstance(field_obj, (models.DateField, models.DateTimeField))
            if is_date_field:
                if isinstance(value, str):
                    if isinstance(field_obj, models.DateField) and not isinstance(field_obj, models.DateTimeField):
                        if 'T' in value or ' ' in value:
                            value = value[:10]
                    if not re.search(r'\d{4}', value) or re.search(r'[a-zA-Z\u4e00-\u9fa5]', value):
                        print(f"DEBUG: Dropping illegal date/time value '{value}' for {field_name_for_val}")
                        continue
        except Exception:
            pass
 
        if clean_key.endswith('__gte') or clean_key.endswith('__lte'):
            q_objects &= Q(**{clean_key: value})
        elif clean_key.endswith('__year'):
            if clean_key == 'dispatch_date__year':
                 q_objects &= Q(timestamp__year=value)
            else:
                 q_objects &= Q(**{clean_key: value})
        else:
            q_objects &= Q(**{clean_key: value})
    return queryset.filter(q_objects)


def handle_search(request, parameters):
    data_source = parameters.get('data_source')
    filters = parameters.get('filters', {})

    if data_source == "Requisition":
        queryset = Requisition.objects.all()
        results = _apply_filters_to_queryset(queryset, filters).order_by().distinct()
        serialized_results = list(results.values('order_number', 'applicant__username', 'request_date', 'status', 'process_type', 'created_at'))
    elif data_source == "WorkOrderMaterial":
        queryset = WorkOrderMaterial.objects.all()
        results = _apply_filters_to_queryset(queryset, filters)
        serialized_results = list(results.values('order_number', 'material_number', 'item_name', 'required_quantity', 'process_type__name', 'estimated_arrival_date'))
    elif data_source == "MaterialTransaction":
        queryset = MaterialTransaction.objects.all()
        results = _apply_filters_to_queryset(queryset, filters)
        serialized_results = list(results.values('material__material_code', 'user__username', 'transaction_type', 'quantity_change', 'timestamp'))
    else:
        return JsonResponse({'error': f'Unknown data source: {data_source}'}, status=400)

    return JsonResponse({'intent': 'SEARCH', 'results': serialized_results})


def handle_export(request, parameters):
    file_type = parameters.get('file_type')
    data_source = parameters.get('data_source')
    filters = parameters.get('filters', {})
    aggregation = parameters.get('aggregation', {})

    if file_type != "Excel":
        return JsonResponse({'error': 'Only Excel export is supported'}, status=400)

    if data_source == "Requisition":
        queryset = Requisition.objects.all()
    elif data_source == "WorkOrderMaterial":
        queryset = WorkOrderMaterial.objects.all()
    elif data_source == "MaterialTransaction":
        queryset = MaterialTransaction.objects.all()
    else:
        return JsonResponse({'error': f'Unknown data source: {data_source}'}, status=400)

    filtered_queryset = _apply_filters_to_queryset(queryset, filters)

    df_data = []
    if aggregation:
        group_by_field = aggregation.get('group_by')
        agg_function = aggregation.get('function')
        agg_field = aggregation.get('field')

        if not agg_function or not agg_field:
            return JsonResponse({'error': 'Missing aggregation parameters'}, status=400)

        if data_source == "MaterialTransaction":
            queryset = filtered_queryset.filter(transaction_type='ALLOCATION')
            if group_by_field:
                aggregated_results = queryset.values(group_by_field).annotate(Total_Quantity=Sum(agg_field))
                data_dict = {
                    '物料編號': [item[group_by_field] for item in aggregated_results],
                    '總撥出數量': [item['Total_Quantity'] for item in aggregated_results]
                }
            else:
                total = queryset.aggregate(Total_Quantity=Sum(agg_field))['Total_Quantity'] or 0
                data_dict = {'總撥出數量': [total]}
            df_data = data_dict
        elif data_source == "WorkOrderMaterial":
            aggregated_results = filtered_queryset.values(group_by_field).annotate(Total_Quantity=Sum(agg_field))
            data_dict = {
                group_by_field: [item[group_by_field] for item in aggregated_results],
                '總需求數量': [item['Total_Quantity'] for item in aggregated_results]
            }
            df_data = data_dict
        elif data_source == "Requisition":
            aggregated_results = filtered_queryset.values(group_by_field).annotate(Total_Count=models.Count('id'))
            data_dict = {
                group_by_field: [item[group_by_field] for item in aggregated_results],
                '總計': [item['Total_Count'] for item in aggregated_results]
            }
            df_data = data_dict
        else:
            return JsonResponse({'error': 'Aggregation not implemented for this data source'}, status=400)
    else:
        if data_source == "Requisition":
            df_data = list(filtered_queryset.values('order_number', 'applicant__username', 'request_date', 'status', 'created_at'))
        elif data_source == "WorkOrderMaterial":
            annotated_queryset = filtered_queryset.annotate(
                shortage_quantity=ExpressionWrapper(
                    F('required_quantity') - Coalesce('confirmed_quantity', Decimal('0.00')),
                    output_field=DecimalField()
                )
            )
            df_data = [
                {
                    '訂單單號': item.order_number,
                    '物料編號': item.material_number,
                    '品名': item.item_name,
                    '需求數量': item.required_quantity,
                    '已撥數量': item.confirmed_quantity or 0,
                    '欠料數量': item.shortage_quantity,
                }
                for item in annotated_queryset
            ]
        elif data_source == "MaterialTransaction":
            df_data = list(filtered_queryset.values('material__material_code', 'user__username', 'transaction_type', 'quantity_change', 'timestamp'))

    from openpyxl import Workbook
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = 'Report'

    if aggregation and 'data_dict' in locals():
        headers = list(data_dict.keys())
        worksheet.append(headers)
        rows = zip(*data_dict.values())
        for row in rows:
            worksheet.append(row)
    elif df_data:
        headers = list(df_data[0].keys())
        worksheet.append(headers)
        for item in df_data:
            for key, value in item.items():
                if isinstance(value, datetime):
                    item[key] = value.replace(tzinfo=None)
            worksheet.append(list(item.values()))

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)

    response = HttpResponse(output.getvalue(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = 'attachment; filename="report.xlsx"'
    return response
