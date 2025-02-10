# pdf_helper.py
import os
from django.http import HttpResponse
from django.template import Template, Context
from xhtml2pdf import pisa

def render_pdf(template_path, context):
    # Baca template dari file
    with open(template_path, 'r') as template_file:
        template_content = template_file.read()

    # Render template menjadi HTML
    template = Template(template_content)
    html_string = template.render(Context(context))

    # Buat PDF dari HTML
    response = HttpResponse(content_type='application/pdf')
    response['Content-Disposition'] = 'attachment; filename="file.pdf"'

    # Konversi HTML ke PDF
    pisa_status = pisa.CreatePDF(html_string, dest=response)

    if pisa_status.err:
        return HttpResponse('Error generating PDF')

    return response