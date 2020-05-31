import openpyxl
import datetime
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from datetime import datetime, date, time
from openpyxl.chart import (LineChart, Reference, ScatterChart, Series, )
from openpyxl.chart.axis import DateAxis

graphs_book = load_workbook('C:\\Work\\Python\\Create_graphs\\графики.xlsx')
ws = graphs_book.active

chart = ScatterChart()
chart.title = "Scatter Chart"
chart.style = 13
chart.x_axis.title = 'Уровень воздействия, ед.'
chart.y_axis.title = '2'


class MLine:
    num_x = 2
    num_y = 2
    line_name = ''

    def __init__(self, x_column, y_column, worksheet):
        self.x_column = x_column
        self.y_column = y_column
        while worksheet.cell(row=self.num_x, column=self.x_column).value is not None:
            self.num_x += 1
        while worksheet.cell(row=self.num_y, column=self. y_column).value is not None:
            self.num_y += 1
        self.line_name = str(worksheet.cell(row=1, column=self.x_column).value)

    def create_plot(self):
        x_values = Reference(ws, min_col=self.x_column, min_row=2, max_row=self.num_x)
        y_values = Reference(ws, min_col=self.y_column, min_row=2, max_row=self.num_y)
        graph = Series(y_values, x_values, title=self.line_name)
        graph.marker.symbol = 'triangle'
        graph.marker.size = 7
        return graph


first_line = MLine(x_column=1, y_column=4, worksheet=ws)
second_line = MLine(x_column=10, y_column=13, worksheet=ws)
chart.series.append(first_line.create_plot())
chart.series.append(second_line.create_plot())
ws.add_chart(chart, "A21")


# class MGraph:
#
#     def __init__(self, x_column, y_column, name_column, ws):
#
#
#         self.chart = ScatterChart()
#         self.chart.style = 13
#         self.chart.x_axis.title = 'Уровень воздействия, ед.'
#         self.chart.y_axis.title = str(ws.cell(row=1, column=y_index).value)





graphs_book.save('C:\\Work\\Python\\Create_graphs\\графики.xlsx')
graphs_book.close()