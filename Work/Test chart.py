from datetime import date

from openpyxl import Workbook
from openpyxl.chart import (
    LineChart,
    Reference,
)
from openpyxl.chart.axis import DateAxis

wb = Workbook()
ws = wb.active

rows = [
    ['Date', 'Batch 1', 'Batch 2', 'Batch 3'],
    [date(2015,9, 1), 40, 30, 25],
    [date(2015,9, 2), 40, 25, 30],
    [date(2015,9, 3), 50, 30, 45],
    [date(2015,9, 4), 30, 25, 40],
    [date(2015,9, 5), 25, 35, 30],
    [date(2015,9, 6), 20, 40, 35],
]

rows2 = [
    ['Date', date(2015,9, 1), date(2015,9, 2), date(2015,9, 3), date(2015,9, 4), date(2015,9, 5),date(2015,9, 6)],
    ['Batch 1 VA Sample GOTV Black and Hispanic Voters in the north east corner', 40, 40, 50, 50, 25, 20],
    ['Batch 2 VA Sample GOTV Black and Hispanic Voters in the north east corner', 30, 25, 30, 25, 35, 40],
    ['Batch 3 VA Sample GOTV Black and Hispanic Voters in the north east corner', 25, 30, 45, 40, 30, 35]
]

rows2[0][2] = ''
rows2[0][4] = ''
rows2[0][6] = ''

for row in rows2:
    ws.append(row)

# size is 15 x 7.5 cm (approximately 5 columns by 14 rows). This can be changed by setting the anchor, width and height properties of the chart.
# chart1.x_axis.title = 'x'
# # chart1.y_axis.title = '1/x'
# chart2.x_axis.scaling.min = 0
# chart2.y_axis.scaling.min = 0
# chart2.x_axis.scaling.max = 11
# chart2.y_axis.scaling.max = 1.5

# for r in dataframe_to_rows(df, index=True, header=True):
#     ws.append(r)

c1 = LineChart()
c1.title = "Line Chart"
c1.style = 13
c1.y_axis.title = 'Addresses'
c1.x_axis.title = 'Date'

# chart.y_axis.crossAx = 500
# chart.x_axis = DateAxis(crossAx=100)
# chart.x_axis.number_format ='yyyy/mm/dd'
# chart.x_axis.majorTimeUnit = "days"

data = Reference(ws, min_col=1, min_row=2, max_col=7, max_row=4)
cats = Reference(ws, min_col=2, min_row=1, max_col=7, max_row=1)
c1.add_data(data, from_rows=True, titles_from_data=True)
c1.x_axis.number_format ='mm/dd'

c1.height = 8.5*2 # default is 7.5
c1.width = 11*2 # default is 15
c1.legend.position = 'b'  # 'tr', 'b', 't', 'l', 'r'

c1.set_categories(cats)
# ws.add_chart(c1, "A10")
ws.add_chart(c1, "A1")
wb.save("SampleChart.xlsx")
