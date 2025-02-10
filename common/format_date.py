import arrow

def format_date(date_string, format_type="DD MMMM YYYY HH:mm:ss"):
    """
    Format the date string into various formats based on format_type.
    Supported formats:
    - "DD MMM YYYY HH:mm:ss" -> 09 Oct 2024 10:10:10
    - "DD MMM YYYY HH:mm" -> 09 Oct 2024 10:10
    - "DD MMM YYYY" -> 09 Oct 2024
    - "DD MMMM YYYY" -> 09 October 2024
    - "DD MMMM YYYY HH:mm:ss" -> 09 October 2024 10:10:10
    - "DD MMMM YYYY HH:mm" -> 09 October 2024 10:10
    """

    if date_string is None:
        return 'N/A'

    try:
        # Parse the date string into an Arrow object
        date = arrow.get(date_string)

        # Define format mappings
        format_mappings = {
            "DD MMM YYYY HH:mm:ss": "DD MMM YYYY HH:mm:ss",  # 09 Oct 2024 10:10:10
            "DD MMM YYYY HH:mm": "DD MMM YYYY HH:mm",        # 09 Oct 2024 10:10
            "DD MMM YYYY": "DD MMM YYYY",                    # 09 Oct 2024
            "DD MMMM YYYY": "DD MMMM YYYY",                  # 09 October 2024
            "DD MMMM YYYY HH:mm:ss": "DD MMMM YYYY HH:mm:ss",# 09 October 2024 10:10:10
            "DD MMMM YYYY HH:mm": "DD MMMM YYYY HH:mm"       # 09 October 2024 10:10
        }

        # Get the desired format or fallback to the default format
        format_pattern = format_mappings.get(format_type, "DD MMMM YYYY HH:mm:ss")

        # Return the formatted date string
        return date.format(format_pattern)

    except Exception as e:
        return f"Error formatting date: {e}"