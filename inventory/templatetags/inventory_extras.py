from django import template
from django.utils.safestring import mark_safe

register = template.Library()

@register.simple_tag(takes_context=True)
def sortable_header(context, sort_field, display_name):
    request = context['request']
    current_sort = request.GET.get('sort_by')
    current_order = request.GET.get('order', 'asc')

    # Determine new order
    if current_sort == sort_field and current_order == 'asc':
        new_order = 'desc'
    else:
        new_order = 'asc'

    # Build URL
    url = f"?sort_by={sort_field}&order={new_order}"

    # Determine icon
    icon = ''
    if current_sort == sort_field:
        if current_order == 'asc':
            icon = ' ▲'
        else:
            icon = ' ▼'
    
    html_output = f'<a href="{url}" class="text-white text-decoration-none">{display_name}{icon}</a>'
    return mark_safe(html_output)
