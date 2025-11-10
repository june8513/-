from django.http import JsonResponse, HttpResponse
from django.db.models import Sum, Q
from django.contrib.auth.decorators import login_required
from django.utils import timezone
import pandas as pd
import io
from datetime import date, datetime

from requisitions.models import Requisition, WorkOrderMaterial # Import WorkOrderMaterial
from inventory.models import MaterialTransaction, Material

# Helper function to apply filters to Django QuerySets
def _apply_filters_to_queryset(queryset, filters):
    q_objects = Q()
    for key, value in filters.items():
        if key == 'status' and value == '缺料':
            # This is a specific rule for '缺料' status, assuming it's a status field
            # In a real scenario, '缺料' might mean stock_quantity < reorder_level
            # For Requisition, we'll assume '缺料' means status is pending or similar
            q_objects &= Q(status='pending') # Placeholder for actual '缺料' logic
        elif key.endswith('__date__gte') or key.endswith('__date__lte'):
            q_objects &= Q(**{key: value})
        elif key.endswith('__date'):
            q_objects &= Q(**{key: value})
        elif key == 'estimated_arrival_date':
            q_objects &= Q(estimated_arrival_date=value)
        elif key == 'dispatch_date__year':
            q_objects &= Q(timestamp__year=value) # For MaterialTransaction
        elif key == 'created_at__year':
            q_objects &= Q(created_at__year=value) # For Requisition
        else:
            # Generic filter for other fields
            q_objects &= Q(**{key: value})
    return queryset.filter(q_objects)


def handle_search(request, parameters):
    data_source = parameters.get('data_source')
    filters = parameters.get('filters', {})

    if data_source == "Requisition":
        queryset = Requisition.objects.all()
        results = _apply_filters_to_queryset(queryset, filters)
        # Serialize results
        serialized_results = list(results.values('order_number', 'applicant__username', 'request_date', 'status', 'created_at'))
    elif data_source == "WorkOrderMaterial":
        queryset = WorkOrderMaterial.objects.all()
        results = _apply_filters_to_queryset(queryset, filters)
        # Serialize results
        serialized_results = list(results.values('order_number', 'material_number', 'item_name', 'required_quantity', 'estimated_arrival_date'))
    elif data_source == "MaterialTransaction":
        queryset = MaterialTransaction.objects.all()
        results = _apply_filters_to_queryset(queryset, filters)
        # Serialize results
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

        if not group_by_field or not agg_function or not agg_field:
            return JsonResponse({'error': 'Missing aggregation parameters'}, status=400)

        # Django ORM aggregation
        if data_source == "MaterialTransaction":
            # For MaterialTransaction, group by material and sum quantity_change
            # Need to filter by transaction_type='ALLOCATION' for '撥出'
            aggregated_results = filtered_queryset.filter(transaction_type='ALLOCATION') \
                                                .values(group_by_field) \
                                                .annotate(Total_Quantity=Sum(agg_field))
            # Convert queryset to a more explicit dictionary for DataFrame creation
            data_dict = {
                '物料編號': [item['material__material_code'] for item in aggregated_results],
                '總撥出數量': [item['Total_Quantity'] for item in aggregated_results]
            }
            df_data = data_dict # For debug printing
        elif data_source == "WorkOrderMaterial":
            # Aggregation for WorkOrderMaterial (e.g., group by material and sum required_quantity)
            aggregated_results = filtered_queryset.values(group_by_field) \
                                                .annotate(Total_Quantity=Sum(agg_field))
            data_dict = {
                group_by_field: [item[group_by_field] for item in aggregated_results],
                '總需求數量': [item['Total_Quantity'] for item in aggregated_results]
            }
            df_data = data_dict # For debug printing
        elif data_source == "Requisition":
            # Example aggregation for Requisition (e.g., count by status)
            aggregated_results = filtered_queryset.values(group_by_field) \
                                                .annotate(Total_Count=models.Count('id'))
            data_dict = {
                group_by_field: [item[group_by_field] for item in aggregated_results],
                '總計': [item['Total_Count'] for item in aggregated_results]
            }
            df_data = data_dict # For debug printing
        else:
            return JsonResponse({'error': 'Aggregation not implemented for this data source'}, status=400)

    else:
        # If no aggregation, just get all fields
        if data_source == "Requisition":
            df_data = list(filtered_queryset.values('order_number', 'applicant__username', 'request_date', 'status', 'created_at'))
        elif data_source == "WorkOrderMaterial":
            df_data = list(filtered_queryset.values('order_number', 'material_number', 'item_name', 'required_quantity', 'estimated_arrival_date'))
        elif data_source == "MaterialTransaction":
            df_data = list(filtered_queryset.values('material__material_code', 'user__username', 'transaction_type', 'quantity_change', 'timestamp'))

    # Create DataFrame from the prepared data
    if 'data_dict' in locals():
        df = pd.DataFrame(data_dict)
    else:
        df = pd.DataFrame(df_data)

    # --- Create DataFrame and export to Excel --- #
    # This part is now replaced with direct openpyxl writing to avoid formatting issues.

    # --- New openpyxl direct writing logic ---
    from openpyxl import Workbook

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = 'Report'

    # Write headers
    if aggregation and 'data_dict' in locals():
        headers = list(data_dict.keys())
        worksheet.append(headers)
        # Write data rows by zipping the value lists
        rows = zip(*data_dict.values())
        for row in rows:
            worksheet.append(row)
    elif df_data: # Handle non-aggregation case
        headers = list(df_data[0].keys())
        worksheet.append(headers)
        for item in df_data:
            # Convert datetime objects to timezone-unaware
            for key, value in item.items():
                if isinstance(value, datetime):
                    item[key] = value.replace(tzinfo=None)
            worksheet.append(list(item.values()))

    # Save the workbook to a memory buffer
    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)

    response = HttpResponse(output.getvalue(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = 'attachment; filename="report.xlsx"'
    return response
