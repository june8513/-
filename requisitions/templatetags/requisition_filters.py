from django import template

register = template.Library()

@register.filter
def sub(value, arg):
    """Subtracts the arg from the value."""
    try:
        return float(value) - float(arg)
    except (ValueError, TypeError):
        return ''

@register.simple_tag(takes_context=True)
def query_transform(context, **kwargs):
    """
    Updates query parameters in the current request's URL.
    """
    query = context['request'].GET.copy()
    for k, v in kwargs.items():
        query[k] = v
    return query.urlencode()

@register.filter
def get_item(dictionary, key):
    return dictionary.get(key)