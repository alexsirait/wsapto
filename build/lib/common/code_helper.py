# helpers.py

import json
from mysatnusa.response import Response

def validate_request(data, rules):
    """
    Validate the input data based on the specified rules.

    Parameters:
    - data (dict): The input data to be validated.
    - rules (dict): A dictionary containing validation rules for each field.

    Returns:
    - Response: A Response object with error messages if validation fails, None otherwise.
    """
    errors = {}

    for field, rule in rules.items():
        rules_list = rule.split('|')

        # Check for 'required'
        if 'required' in rules_list and field not in data:
            errors[field] = f"{field} is required."

        # Check for 'numeric'
        if 'numeric' in rules_list and field in data:
            try:
                float(data.get(field))
            except ValueError:
                errors[field] = f"{field} must be a numeric value."

        # Add other rules as needed
        # Example: if 'email' in rules_list, you could add email validation

    if errors:
       return Response.badRequest(request, message=json.dumps(errors), messagetype="E")

    return None

