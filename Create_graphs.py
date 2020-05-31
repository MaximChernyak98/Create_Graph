import openpyxl
import datetime
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from datetime import datetime, date, time
from openpyxl.chart import (LineChart, Reference, ScatterChart, Series, )
from openpyxl.chart.axis import DateAxis

graphs_book = load_workbook('C:\\Work\\Python\\Create_graphs\\графики.xlsx')
ws = graphs_book.active


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


def create_graphs(start_index, step, num_of_param, num_samples, chart):
    current_offset = 0
    for sample in range(num_samples):
        line = MLine(x_column=(start_index+current_offset),
                     y_column=(start_index+current_offset+num_of_param),
                     worksheet=ws)
        chart.series.append(line.create_plot())
        current_offset += step


def number_to_letter(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string


num_of_params = 6
step = 9

for params in range(1, (num_of_params+1)):
    chart = ScatterChart()
    chart.title = "Scatter Chart"
    chart.style = 13
    chart.x_axis.title = 'Уровень воздействия, ед.'
    chart.y_axis.title = '2'
    create_graphs(start_index=1, step=step, num_of_param=params, num_samples=10, chart=chart)
    ws.add_chart(chart, f"{number_to_letter((params-1)*step+1)}21")



graphs_book.save('C:\\Work\\Python\\Create_graphs\\графики.xlsx')
graphs_book.close()