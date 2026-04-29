from .models import WorkOrderMaterial, WorkOrder
from inventory.models import Material
from django.db.models import Q, F, Sum, Max, Value, DecimalField, OuterRef, Subquery, ExpressionWrapper
from django.db import models # Import models for models.DateField
from django.db.models.functions import Coalesce, Greatest
from decimal import Decimal
import datetime

def get_material_demand_analysis():
    """
    Analyzes all active work order materials to determine demand, shortage, and running stock.
    This is the shared logic used by both estimated_material_demand and material_completeness views.
    """
    # Get order numbers of archived work orders
    archived_orders = WorkOrder.objects.filter(is_archived=True).values_list('order_number', flat=True)

    # Base queryset for WorkOrderMaterial, excluding materials from archived orders
    demand_qs = WorkOrderMaterial.objects.filter(
        is_active=True
    ).exclude(
        order_number__in=list(archived_orders)
    ).exclude(
        material_number='PARENT_SCOPE'
    ).annotate(
        # Add shipping_date annotation
        shipping_date=Subquery(
            WorkOrder.objects.filter(order_number=OuterRef('order_number')).values('shipping_date')[:1],
            output_field=models.DateField() # Assuming shipping_date is a DateField
        ),
        remaining_required_quantity=ExpressionWrapper(
            Greatest(Decimal('0.00'), F('required_quantity') - Coalesce(F('confirmed_quantity'), Decimal('0.00'))),
            output_field=DecimalField()
        ),
        current_stock=Subquery(
            Material.objects.filter(
                material_code=OuterRef('material_number')
            ).values('system_quantity')[:1],
            output_field=DecimalField()
        )
    ).filter(
        remaining_required_quantity__gt=0
    )

    # Aggregate in Python
    final_aggregated_data = {}
    for item in demand_qs.order_by('demand_date', 'material_number').values(
        'pk', 'demand_date', 'material_number', 'item_name', 'machine_model__name',
        'process_type__name', 'remaining_required_quantity', 'order_number', 'current_stock', 
        'estimated_arrival_date', 'shipping_date' # Use 'shipping_date' here
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
            'shipping_date': item['shipping_date'] # Use 'shipping_date' here
        })

    # More robustly fetch the latest estimated arrival date for all relevant materials at once
    if final_aggregated_data:
        material_keys = final_aggregated_data.keys()
        
        # Fetch max estimated_arrival_date from WorkOrderMaterial
        aggregated_arrival_dates = WorkOrderMaterial.objects.filter(
            is_active=True,
            material_number__in=material_keys
        ).values('material_number').annotate(max_arrival_date=Max('estimated_arrival_date'))
        date_map = {item['material_number']: item['max_arrival_date'] for item in aggregated_arrival_dates}

        # Fetch purchaser from Material model
        purchaser_map = {
            mat.material_code: mat.purchaser 
            for mat in Material.objects.filter(material_code__in=material_keys)
        }

        for material_key, data in final_aggregated_data.items():
            data['estimated_arrival_date'] = date_map.get(material_key)
            data['purchaser'] = purchaser_map.get(material_key)

    # Pre-fetch all WorkOrderMaterial data needed for first_shortage_shipping_date calculation
    # This query will fetch all relevant WOM items once, for all materials in final_aggregated_data
    material_numbers_in_demand = list(final_aggregated_data.keys())
    
    all_wom_for_shortage_check = WorkOrderMaterial.objects.filter(
        material_number__in=material_numbers_in_demand,
        is_active=True
    ).exclude(
        order_number__in=list(archived_orders)
    ).annotate(
        shipping_date=Subquery(
            WorkOrder.objects.filter(order_number=OuterRef('order_number')).values('shipping_date')[:1],
            output_field=models.DateField()
        ),
        remaining_required_quantity=ExpressionWrapper(
            Greatest(Decimal('0.00'), F('required_quantity') - Coalesce(F('confirmed_quantity'), Decimal('0.00'))),
            output_field=DecimalField()
        )
    ).filter(
        remaining_required_quantity__gt=0
    ).order_by('material_number', 'demand_date', 'shipping_date', 'pk').values(
        'material_number', 'demand_date', 'shipping_date', 'remaining_required_quantity'
    )
    
    # Group this prefetched data by material_number for easier access
    grouped_wom_data = {}
    for wom_item_data in all_wom_for_shortage_check:
        mn = wom_item_data['material_number']
        if mn not in grouped_wom_data:
            grouped_wom_data[mn] = []
        grouped_wom_data[mn].append(wom_item_data)

    # Calculate final_shortage and determine if a material is in shortage
    # Also find the first shortage shipping date
    for material_key, data in final_aggregated_data.items():
        data['final_shortage'] = data['total_required_quantity'] - data['current_stock']
        data['is_shortage'] = data['final_shortage'] > 0
        data['shortage_date'] = None # Initialize shortage_date (this is the demand date of the first shortage)
        data['first_shortage_shipping_date'] = None # New field for requirement 2

        # Sort detail_orders by demand_date, shipping_date, and pk for robust running stock calculation
        data['detail_orders'].sort(key=lambda x: (
            x['demand_date'] if x['demand_date'] else datetime.date.max,
            x['shipping_date'] if x['shipping_date'] else datetime.date.max,
            x['pk']
        ))

        # Only proceed to find first_shortage_shipping_date if there is an overall shortage
        if data['is_shortage']:
            running_stock_for_shortage_check = data['current_stock']
            
            # Use the pre-fetched and grouped data
            relevant_wom_items_for_material = grouped_wom_data.get(material_key, [])

            # Loop through the ordered relevant WorkOrderMaterial items to find the first shortage
            for wom_item_data in relevant_wom_items_for_material:
                required_qty = wom_item_data['remaining_required_quantity']
                running_stock_for_shortage_check -= required_qty

                if running_stock_for_shortage_check < 0:
                    # This is the WorkOrderMaterial that causes the first shortage
                    if wom_item_data['shipping_date']: # Access directly from pre-fetched data
                        data['first_shortage_shipping_date'] = wom_item_data['shipping_date']
                    break # Found the first shortage, no need to check further

        # Update running stock and is_running_shortage for display in detail_orders
        running_stock_for_display = data['current_stock']
        running_demand = Decimal('0.00')
        for detail in data['detail_orders']:
            required_qty = detail['required_quantity']
            running_stock_for_display -= required_qty
            running_demand += required_qty
            detail['running_stock'] = str(running_stock_for_display)
            detail['cumulative_demand'] = str(running_demand)
            detail['is_running_shortage'] = running_stock_for_display < 0

            # Find the first shortage demand date (this is different from shipping date)
            if data['shortage_date'] is None and detail['is_running_shortage']:
                data['shortage_date'] = detail['demand_date']

    return final_aggregated_data