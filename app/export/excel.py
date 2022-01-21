#from django.utils.translation import ugettext as trans, ugettext_lazy as _
from io import BytesIO
import xlsxwriter
import gettext
#sfrom django.utils.translation import ugettext
from django.utils.translation import gettext_lazy as _

def WriteToExcel(weather_data, town=None):
    output = BytesIO()
    workbook = xlsxwriter.Workbook(output)

    # define styles to use
    title = workbook.add_format({
            'bold': True,
            'font_size': 14,
            'align': 'center',
            'valign': 'vcenter'
    })
    header = workbook.add_format({
            'bg_color': '#cfe7f5',
            'color': 'black',
            'align': 'center',
            'valign': 'top',
            'border': 1
    })
    cell = workbook.add_format({})
    cell_center = workbook.add_format({'align': 'center'})

    # add a worksheet to work with
    

    # add chart
    worksheet_c = workbook.add_worksheet("Charts")
    worksheet_d = workbook.add_worksheet("Chart data")

    # chart data
    for row_index, data in enumerate(weather_data):
        worksheet_d.write(row_index, 0, data.date.strftime("%Y-%m-%d"))
        worksheet_d.write_number(row_index, 1, data.min)
        worksheet_d.write_number(row_index, 2, data.mean)
        worksheet_d.write_number(row_index, 3, data.max)

    # line chart
   
    workbook.close()
    xlsx_data = output.getvalue()
    # xlsx_data contains the Excel file
    return xlsx_data

