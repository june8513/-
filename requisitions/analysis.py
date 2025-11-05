# This file will contain shared business logic for material demand analysis.

from .models import WorkOrderMaterial, Inventory
from inventory.models import Material
from django.db.models import Q, F, Sum, Max, Value, DecimalField, OuterRef, Subquery, ExpressionWrapper
from django.db.models.functions import Coalesce
from decimal import Decimal
import datetime

def get_material_demand_analysis():
    """
    Analyzes all active work order materials to determine demand, shortage, and running stock.
    This is the shared logic used by both estimated_material_demand and material_completeness views.
    """
    # Base queryset for WorkOrderMaterial
    demand_qs = WorkOrderMaterial.objects.filter(
        is_active=True
    ).exclude(
        material_number='PARENT_SCOPE'
    ).annotate(
        remaining_required_quantity=ExpressionWrapper(
            F('required_quantity') - Coalesce(F('confirmed_quantity'), Decimal('0.00')),
            output_field=DecimalField()
        ),
        current_stock=Subquery(
            Inventory.objects.filter(
                material_number=OuterRef('material_number')
            ).values('stock_quantity')[:1],
            output_field=DecimalField()
        )
    ).filter(
        remaining_required_quantity__gt=0
    )

    # Aggregate in Python
    final_aggregated_data = {}
    for item in demand_qs.order_by('demand_date', 'material_number').values(
        'pk', 'demand_date', 'material_number', 'item_name', 'machine_model__name',
        'process_type__name', 'remaining_required_quantity', 'order_number', 'current_stock', 'estimated_arrival_date'
    ):
        material_key = item['material_number']
        if material_key not in final_aggregated_data:
            final_aggregated_data[material_key] = {
                'pk': item['pk'], # Add pk to the top-level aggregated data
                'material_number': item['material_number'],
                'item_name': item['item_name'],
                'current_stock': item['current_stock'] or Decimal('0.00'),
                'total_required_quantity': Decimal('0.00'),
                'detail_orders': [],
                'estimated_arrival_date': None # Initialize with None
            }
        
        final_aggregated_data[material_key]['total_required_quantity'] += item['remaining_required_quantity']
        final_aggregated_data[material_key]['detail_orders'].append({
            'pk': item['pk'],
            'order_number': item['order_number'],
            'demand_date': item['demand_date'],
            'required_quantity': item['remaining_required_quantity'],
            'machine_model_name': item['machine_model__name'],
            'process_type_name': item['process_type__name'],
        })

        # This part of the logic is now replaced by a more robust aggregation below.
        # # Keep track of the latest estimated_arrival_date for each material
        # if item['estimated_arrival_date']:
        #     if not final_aggregated_data[material_key]['estimated_arrival_date'] or item['estimated_arrival_date'] > final_aggregated_data[material_key]['estimated_arrival_date']:
        #         final_aggregated_data[material_key]['estimated_arrival_date'] = item['estimated_arrival_date']

    # More robustly fetch the latest estimated arrival date for all relevant materials at once
    if final_aggregated_data:
        material_keys = final_aggregated_data.keys()
        aggregated_dates = WorkOrderMaterial.objects.filter(
            is_active=True,
            material_number__in=material_keys
        ).values('material_number').annotate(max_arrival_date=Max('estimated_arrival_date'))

        date_map = {item['material_number']: item['max_arrival_date'] for item in aggregated_dates}

        for material_key, data in final_aggregated_data.items():
            data['estimated_arrival_date'] = date_map.get(material_key)

    # Calculate final_shortage and determine if a material is in shortage
    for material_key, data in final_aggregated_data.items():
        data['final_shortage'] = data['total_required_quantity'] - data['current_stock']
        data['is_shortage'] = data['final_shortage'] > 0

        # Calculate running stock and is_running_shortage
        running_stock = data['current_stock']
        data['detail_orders'].sort(key=lambda x: (x['demand_date'] if x['demand_date'] else datetime.date.max))
        for detail in data['detail_orders']:
            required_qty = detail['required_quantity']
            running_stock -= required_qty
            detail['running_stock'] = str(running_stock)
            detail['is_running_shortage'] = running_stock < 0

    return final_aggregated_data