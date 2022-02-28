from math import exp,sqrt
from tempfile import tempdir
import matplotlib.pyplot as plt
from numpy import empty_like
import xlwings as xw 

file_name = 'Thrust_Calculator.xlsm'
sheet_name = 'Calculator'

def read_from_database(file_name, sheet_name, column_number):
    wb = xw.Book(file_name)
    wb2 = xw.Book('Thrust_Calculator.xlsm')
    ws = wb.sheets[sheet_name]
    row = 2
    empty_list = []

    ws2= wb2.sheets['Calculator']
    length = int((ws2.range('L18').value - ws2.range('L15').value) / ws2.range('L19').value) + 1
 
    if column_number == 1:
        for i in range(length):
            empty_list.append(ws.range(f"P{row}").value)
            row += 1
        return empty_list

    elif column_number == 2:
        for i in range(length):
            empty_list.append(ws.range(f"Q{row}").value)
            row += 1
        return empty_list

    elif column_number == 3:
        for i in range(length):
            empty_list.append(ws.range(f"R{row}").value)
            row += 1
        return empty_list

    else:
        for i in range(length):
            empty_list.append(ws.range(f"S{row}").value)
            row += 1
        return empty_list

    wb.save(file_name)



def generate_plot(altitude, thrust):
    plt.plot(thrust, altitude)
    plt.xlabel('static_thrust')
    plt.ylabel('altitude')
    plt.show()


altitude = read_from_database(file_name, sheet_name, 1)
temperature = read_from_database(file_name, sheet_name, 2)
pressure = read_from_database(file_name, sheet_name, 3)
static_thrust = read_from_database(file_name, sheet_name, 4)

generate_plot(altitude, static_thrust)