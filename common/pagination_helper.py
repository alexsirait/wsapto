from django.core.paginator import Paginator

def paginate_data(request, data):
    page_number = request.GET.get('page', None)
    page_size = request.GET.get('page_size', 10)

    if not page_number and not page_size:
        return {
            "current_page": 1,
            "page_size": len(data),
            "total_pages": 1,
            "total_rows": len(data),
            "last_page": 1,
            "next_page_url": None,
            "prev_page_url": None,
            "rows": data,
        }

    page_number = int(page_number) if page_number else 1
    page_size = int(page_size) if page_size else 10

    paginator = Paginator(data, page_size)
    page_obj = paginator.get_page(page_number)

    base_url = request.build_absolute_uri(request.path)
    next_page_url = prev_page_url = None

    if page_obj.has_next():
        next_page_url = f"{base_url}?page={page_obj.next_page_number()}&page_size={page_size}"
    if page_obj.has_previous():
        prev_page_url = f"{base_url}?page={page_obj.previous_page_number()}&page_size={page_size}"

    return {
        "current_page": page_obj.number,
        "page_size": page_size,
        "total_pages": paginator.num_pages,
        "total_rows": paginator.count,
        "last_page": paginator.num_pages,
        "next_page_url": next_page_url,
        "prev_page_url": prev_page_url,
        "rows": list(page_obj),
    }
