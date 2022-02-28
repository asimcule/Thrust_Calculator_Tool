from math import exp,sqrt
import matplotlib.pyplot as plt
from numpy import power
import xlwings as xw


def extract_from_workbook(file_name, sheet_name):
    wb = xw.Book(file_name)
    ws = wb.sheets[sheet_name]
    data = []
    current_altitude = ws.range('L15').value
    thrust_current_altitude = ws.range('L16').value
    power_current_altitude = ws.range('L17').value
    target_altitude = ws.range('L18').value
    resolution = ws.range('L19').value
    wb.save(file_name)
    data.append(current_altitude)
    data.append(thrust_current_altitude)
    data.append(power_current_altitude)
    data.append(target_altitude)
    data.append(resolution)
    wb.save(file_name)
    return data
    

def write_to_workbook(file_name, sheet_name, power_target_altitude, thrust_factor, static_thrust,length):
    wb = xw.Book(file_name)
    ws = wb.sheets[sheet_name]
    ws.range('L26').value = static_thrust[length - 1]
    ws.range('L27').value = thrust_factor
    ws.range('L28').value = power_target_altitude
    wb.save(file_name)

