from math import exp,sqrt
from os import stat
from turtle import st
import matplotlib.pyplot as plt
import xlwings as xw
from workbook import extract_from_workbook, write_to_workbook



#constants
g = 9.8 #m/s2
M = 0.0289644 #kg/mol
R = 8.31432 #N.m/(mol.K)
P_not = 101.325 #Kpa    
lamda = 6.5/1000  #in degree celcius per meter
sea_level_temperature = 15  #in degree celcius

def thrustfactor_calculator(static_thrust, thrust_current_altitude):
    lis = []
    for thrust in static_thrust:
        lis.append(thrust/thrust_current_altitude)
    return lis


def power_required_at_each_altitude(thrustfactor_list, power):
    lis = []
    for thrust in thrustfactor_list:
        lis.append(power * (1/thrust))
    return lis


def power_target_altitude(thrust_factor, power):
    return (power * (1/thrust_factor))


def altitude_calculator(current_altitude, steps, resolution):
    altitude = []
    for i in range(steps + 1):
        altitude.append(current_altitude + (i*resolution))
    return altitude


def static_thrust_calculator(pressure,temperature, thrust_sea_level):            #return parameter is a list in each case
    thrust = []

    for i in range(len(pressure)):
        thrust.append(thrust_sea_level * (pressure[i]/P_not) * sqrt(((sea_level_temperature*9/5) + 491)/((temperature[i]*9/5)+491)))
    return thrust


def pressure_calculator(temperature, altitude):
    pressure = []

    for i in range(len(temperature)):
        pressure.append(P_not * (exp((-g*M*altitude[i])/(R*(temperature[i] + 273)))))     #temperature in kelvin
    return pressure  


def temperature_calculator(altitude):
    temperature = []
    for alt in altitude:
        temp = sea_level_temperature - (lamda * alt)
        temperature.append(temp)
    return temperature

def thrust_sea_level_calculator(thrust_current_altitude, current_altitude, current_altitude_temperature):
    current_altitude_pressure = (P_not * (exp((-g*M*current_altitude)/(R*(current_altitude_temperature + 273)))))
    thrust_sea_level = thrust_current_altitude * (P_not/current_altitude_pressure) * sqrt(((current_altitude_temperature*9/5)+491)/((sea_level_temperature*9/5)+491))
    return thrust_sea_level


def generate_tables(file_name, sheet_name, altitude, pressure, temperature, static_thrust, power_ateach_altitude):     #writes to the database
    wb = xw.Book(file_name)
    ws = wb.sheets[sheet_name]
    row = 2
    for i in range(len(altitude)):
        ws.range(f"P{row}").value = altitude[i]
        row += 1
    
    row = 2
    for i in range(len(altitude)):
        ws.range(f"Q{row}").value = temperature[i]
        row += 1
    
    row = 2
    for i in range(len(altitude)):
        ws.range(f"R{row}").value = pressure[i]
        row += 1

    
    row = 2
    for i in range(len(altitude)):
        ws.range(f"S{row}").value = static_thrust[i]
        row += 1
    wb.save(file_name)

    row = 2
    for i in range(len(altitude)):
        ws.range(f"T{row}").value = power_ateach_altitude[i]
        row += 1
    wb.save(file_name)


#driver function that handles the program
#handles the workbook
file_name = 'Thrust_Calculator.xlsm'
database = "database.xlsx"
Calculator = 'Calculator'
Tables = 'Tables'
Table = 'Table'
primary_input = extract_from_workbook(file_name, Calculator)         #of type list


#validation for ambient temperature for the current altitude
current_altitude_temperature = sea_level_temperature - (primary_input[0] * lamda)
current_altitude = primary_input[0]
resolution = primary_input[4]
thrust_current_altitude = primary_input[1]  
target_altitude = primary_input[3]
power = primary_input[2]
print(f"The ambient temperature for your current altitude was determined to be {current_altitude_temperature:.2f} *C")
print()
print()
print()
validation = input("Do you want to manually enter the ambient temperature for your current altitude(y/n):")
if validation.lower() == 'y':
    current_altitude_temperature = float(input("Please enter the ambient temperature:"))

thrust_sea_level = thrust_sea_level_calculator(primary_input[1], primary_input[0], current_altitude_temperature)         #constant value for a specific propeller
steps = int((target_altitude-current_altitude)/primary_input[4])



#calculating all the altitudes
altitude = altitude_calculator(current_altitude, steps, resolution)


#calculating all the corresponding temperature values
temperature = temperature_calculator(altitude)

print(f"The ambient temperature for your target altitude was determined to be {temperature[len(temperature) - 1]:.2f} *C")
validation = input("Do you want to manually enter the ambient temperature for your target altitude(y/n):")
if validation.lower() == 'y':
    temperature[len(temperature) - 1] = float(input("Please enter the ambient temperature:"))

#calculating all the corresponding pressure values
pressure = pressure_calculator(temperature, altitude)

#calculating the static thrust values for all the altitudes
static_thrust = static_thrust_calculator(pressure, temperature, thrust_sea_level)

#calculating the thrust factor and power factor 
thrust_factor = static_thrust[len(static_thrust)-1]/static_thrust[0]
power_target_altitude = power_target_altitude(thrust_factor, power)

thrustfactor_list = thrustfactor_calculator(static_thrust, thrust_current_altitude)
power_ateach_altitude = power_required_at_each_altitude(thrustfactor_list, power)

length = len(static_thrust)
write_to_workbook(file_name, Calculator, power_target_altitude, thrust_factor, static_thrust, length)

#write to database
generate_tables(file_name, 'Calculator', altitude, pressure, temperature, static_thrust, power_ateach_altitude)

#integrate this using a new function
""""
#plotting the graph
generate_plot(static_thrust, altitude)

#generating tables
generate_tables(file_name, Tables, altitude, pressure, temperature, static_thrust)

"""