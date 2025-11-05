from django.shortcuts import render
from django.contrib.auth.decorators import login_required
from requisitions.models import WorkOrderMaterial
from ..analysis import get_material_demand_analysis
from decimal import Decimal

@login_required
def material_completeness(request):
    order_number = request.GET.get('order_number')
    context = {
        'order_number': order_number,
    }

    if order_number:
        # Get the global material analysis data
        all_materials_analysis = get_material_demand_analysis()

        # Filter the analysis data for the current order
        materials_in_order = WorkOrderMaterial.objects.filter(order_number=order_number, is_active=True)

        dispatched_materials = []
        shortage_materials = []
        pending_materials = []

        for material in materials_in_order:
            # 1. Check for dispatched
            if material.confirmed_quantity and material.confirmed_quantity >= material.required_quantity:
                dispatched_materials.append(material)
                continue

            # 2. Check for shortage using the global analysis data
            material_analysis = all_materials_analysis.get(material.material_number)
            if material_analysis and material_analysis['is_shortage']:
                shortage_materials.append(material)
            else:
                pending_materials.append(material)

        # Get the global final arrival date from all shortage materials
        global_final_arrival_date = None
        all_shortage_arrival_dates = []
        for mat_data in all_materials_analysis.values():
            if mat_data['is_shortage'] and mat_data['estimated_arrival_date']:
                all_shortage_arrival_dates.append(mat_data['estimated_arrival_date'])
        
        if all_shortage_arrival_dates:
            global_final_arrival_date = max(all_shortage_arrival_dates)

        context.update({
            'total_count': len(materials_in_order),
            'dispatched_count': len(dispatched_materials),
            'pending_count': len(pending_materials),
            'shortage_count': len(shortage_materials),
            'final_arrival_date': global_final_arrival_date, # Use the global date
            'dispatched_materials': dispatched_materials,
            'pending_materials': pending_materials,
            'shortage_materials': shortage_materials,
        })

    return render(request, 'requisitions/material_completeness.html', context)
