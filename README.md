# django_table2excel
convert HTML table tag to excel in Django

## How to use?
···python
import html_table_to_excel
def some_view(request):
    table_str = render_to_string('project/templates/yourhtmltable.html')
    return html_table_to_excel.export_to_xls(table_str, True)
···
