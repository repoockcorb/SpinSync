import odrive
from odrive.enums import *
import time
import openpyxl
from openpyxl import Workbook
from openpyxl.chart import (
    ScatterChart,
    Reference,
    Series
)
from openpyxl.chart.trendline import Trendline
# import matplotlib.pyplot as plt
import subprocess

# Find a connected ODrive (this will block until you connect one)
odrive_serial_number = "3943355F3231"
odrive = odrive.find_any(serial_number=odrive_serial_number)

# Clear Odrive S1 Errors if any
odrive.clear_errors()

## Odrivetool command to backup and restore configuration
# odrivetool backup-config my_config.json
# odrivetool restore-config my_config.json

# Check if the connection is successful
if odrive is not None:
    print(f"Connected to ODrive S1 with serial number {odrive_serial_number}")

# Ask to Calibrate motor and wait for it to finish
while True:
    response = input("Do you want to run calibration Y/n? ")
    if response == "Y" or response == "y":
        print("starting calibration...")
        odrive.axis0.requested_state = AXIS_STATE_FULL_CALIBRATION_SEQUENCE
        while odrive.axis0.current_state != AXIS_STATE_IDLE:
            time.sleep(0.1)
        print("Calibration Complete")
        break
    elif response == "n" or response == "N":
        print("skipping calibration")
        break
    else:
        print("Invalid input. Please enter Y or n")

# Excel File Name
file_name = "Rogue_Power_SpinSync Comparison 70rpm.xlsx"


# RPM Setpoints
rpm_set_points = [70, 0]
# rpm_set_points = [10, 20, 30, 40, 50, 60, 70, 80, 0]

# rpm_set_points = [5.0, 7.5, 10.0, 12.5, 15.0, 17.5, 20.0, 22.5, 25.0, 27.5, 30.0, 32.5, 35.0, 37.5, 40.0, 
#                   42.5, 45.0, 47.5, 50.0, 52.5, 55.0, 57.5, 60.0, 62.5, 65.0, 67.5, 70.0, 72.5, 75.0, 77.5, 80, 0]

# Efficiency Values
gearbox_efficiency = 0.96
chain_efficiency = 0.98

# Configure motor 
odrive.axis0.controller.config.control_mode = CONTROL_MODE_VELOCITY_CONTROL

odrive.axis0.controller.config.vel_ramp_rate = 5
odrive.axis0.controller.config.vel_limit = 96

# odrive.axis0.controller.config.vel_gain = 1.2
# odrive.axis0.controller.config.vel_integrator_gain = 8.0

odrive.axis0.controller.config.vel_gain = 0.167
odrive.axis0.controller.config.vel_integrator_gain = 0.333

odrive.axis0.controller.config.input_mode = InputMode.VEL_RAMP
odrive.axis0.requested_state = AXIS_STATE_CLOSED_LOOP_CONTROL

# Safety Trip Levels
odrive.config.dc_max_positive_current = 38
odrive.config.dc_bus_undervoltage_trip_level = 25

# Make sure the initial velocity is set to 0
odrive.axis0.controller.input_vel = 0 # target velocity in counts/s

# Create New Excel workbook and sheet
wb = openpyxl.Workbook()
sheet = wb.active
sheet.title = "Raw Data"
sheet["A1"] = "Pedal RPM Setpoint"
sheet["B1"] = "Motor Electrical Power (Watts)"
sheet["C1"] = "Motor Mechcanical Power (Watts)"
sheet["D1"] = "Motor Velocity (rev/s)"
sheet["E1"] = "Pedal Velocity (RPM)"
sheet["F1"] = "Motor Torque (Nm)"
sheet["G1"] = "Controller Electrical Power (Watts)"
sheet["H1"] = "Controller Mechanical Power (Watts)"
sheet["I1"] = "Controller Efficiency (%)"
sheet["J1"] = "Pedal Power (Watts)"

avg_sheet = wb.create_sheet("Averages")
avg_sheet["A1"] = "Pedal RPM Setpoint"
avg_sheet["B1"] = "Average Motor Electrical Power (Watts)"
avg_sheet["C1"] = "Average Motor Mechanical Power (Watts)"
avg_sheet["D1"] = "Average Motor Velocity (rev/s)"
avg_sheet["E1"] = "Average Pedal Velocity (RPM)"
avg_sheet["F1"] = "Average Motor Torque (Nm)"
avg_sheet["G1"] = "Average Controller Electrical Power (Watts)"
avg_sheet["H1"] = "Average Controller Mechanical Power (Watts)"
avg_sheet["I1"] = "Average Controller Efficiency (%)"
avg_sheet["J1"] = "Average Pedal Power (Watts)"

input("Press Enter to continue...")

while True:
    readings = input("How many readings 1 second apart do you want to take? ")

    if readings.isdigit():
        readings = int(readings)
        # Continue with your code using the 'readings' variable
        print(f"You want to take {readings} readings.")
    else:
        print("Please enter a valid integer.")

    response = input("Are you ready to start Y/n? ")
    if response == "Y" or response == "y":
        for rpm_setpoint in rpm_set_points:
            print("Starting...")
            target_velocity = (rpm_setpoint / 60) * 10 * 4.6153846  # Convert RPM to velocity
            odrive.axis0.controller.input_vel = target_velocity
            print(f"Motor Setpoint: {abs(rpm_setpoint)} rpm")

            # Wait until the motor reaches the setpoint
            while abs(odrive.axis0.pos_vel_mapper.vel - target_velocity) > 0.1:
                position = odrive.axis0.pos_vel_mapper.pos_rel
                velocity = abs(odrive.axis0.pos_vel_mapper.vel)
                time.sleep(0.1)

            if abs(odrive.axis0.pos_vel_mapper.vel/10/4.6153846*60) > 0.5:
                print("Target Setpoint Reached stabilising speed")
                # time.sleep(5)
                print("Speed Stabilised \nTaking "+ str(readings) +" readings 1 second apart:")

            # user input to continue
            input("Press Enter to continue...")
            
            # Data Variables
            motor_electrical_power = []
            motor_mechanical_power = []
            motor_vel = []
            pedal_rpm = []
            motor_torque = []
            electrical_power = []
            mecanical_power = []
            controller_efficiency = []
            pedal_power = []

            # Take set readings, 1 second apart, and calculate average power
            for _ in range(readings):
                if abs(odrive.axis0.pos_vel_mapper.vel/10/4.6153846*60) < 0.5:
                    break
                time.sleep(1)
                motor_electrical_power.append(odrive.axis0.motor.electrical_power)
                motor_mechanical_power.append(odrive.axis0.motor.mechanical_power)
                motor_vel.append(abs(odrive.axis0.pos_vel_mapper.vel))
                pedal_rpm.append(abs(odrive.axis0.pos_vel_mapper.vel/10/4.6153846*60))
                motor_torque.append(abs(odrive.axis0.controller.effective_torque_setpoint))
                electrical_power.append(odrive.axis0.controller.spinout_electrical_power)
                mecanical_power.append(odrive.axis0.controller.spinout_mechanical_power)
                controller_efficiency.append((odrive.axis0.controller.spinout_mechanical_power/odrive.axis0.controller.spinout_electrical_power)*100)
                pedal_power.append(odrive.axis0.controller.spinout_mechanical_power * chain_efficiency * gearbox_efficiency)

                sheet.append([abs(rpm_setpoint), motor_electrical_power[-1], motor_mechanical_power[-1], motor_vel[-1], pedal_rpm[-1], motor_torque[-1], electrical_power[-1], mecanical_power[-1], controller_efficiency[-1], pedal_power[-1]])
                print(f"Reading: Pedal_RPM = {pedal_rpm[-1]} Pedal Power = {pedal_power[-1]}")
                # print(f"Reading {len(motor_electrical_power)}: Power = {motor_electrical_power[-1]} Motor_vel = {motor_vel[-1]} Pedal_RPM = {pedal_rpm[-1]} Motor Torque = {motor_torque[-1]} Motor Electrical Power = {electrical_power[-1]} Motor Mechanical Power = {mecanical_power[-1]} Motor Efficiency = {controller_efficiency[-1]} Pedal Power = {pedal_power[-1]}")
            
            if abs(odrive.axis0.pos_vel_mapper.vel/10/4.6153846*60) > 0.5:
                average_motor_electrical_power = sum(motor_electrical_power) / len(motor_electrical_power)
                average_motor_mechanical_power = sum(motor_mechanical_power)/len(motor_mechanical_power)
                average_motor_vel = sum(motor_vel) / len(motor_vel)
                average_pedal_rpm = sum(pedal_rpm) / len(pedal_rpm)
                average_torque = sum(motor_torque) / len(motor_torque)
                avg_electrical= sum(electrical_power) / len(electrical_power)
                avg_mechanical = sum(mecanical_power) / len(mecanical_power)
                avg_efficiency = sum(controller_efficiency) / len(controller_efficiency)
                avg_pedal_power = sum(pedal_power)/len(pedal_power)
                
                print(f"Average: \n Pedal RPM = {average_pedal_rpm} Pedal Power = {avg_pedal_power} ")
                # Append RPM setpoint and average power to the sheet
                avg_sheet.append([abs(rpm_setpoint), average_motor_electrical_power, average_motor_mechanical_power, average_motor_vel, average_pedal_rpm, average_torque, avg_electrical, avg_mechanical, avg_efficiency, avg_pedal_power])

        # Set Odrive State to Idle
        odrive.axis0.requested_state = AXIS_STATE_IDLE


        #### PLOTS ####
        # Plot Data in Excel for each row
        chart = ScatterChart()
        chart.title = "Average Pedal Velocity vs. Average Power"
        chart.style = 2
        chart.x_axis.title = "Average Pedal Velocity (RPM)"
        chart.y_axis.title = "Average Power (Watts)"

        xvalues = Reference(avg_sheet, min_col=5, min_row=2, max_row=len(rpm_set_points))
        for i in range(10, 11):
            values = Reference(avg_sheet, min_col=i, min_row=1, max_row=len(rpm_set_points))
            series = Series(values, xvalues, title_from_data=True)
            series.marker.symbol = "circle"  # Set marker to circle
            series.marker.size = 7  # Set marker size
            series.smooth = True  # Use smooth lines
            chart.series.append(series)
        
        # Set the size of the chart
        chart.width = 25  # adjust the width as needed
        chart.height = 12  # adjust the height as needed

        avg_sheet.add_chart(chart, "A"+str(len(rpm_set_points)+2)) # Adjust the position where the chart will be placed

        # Plot scatter chart for raw data
        chart_raw = ScatterChart()
        chart_raw.title = "Pedal Velocity vs. Power"

        # Set scatterStyle to 'marker' to remove lines connecting the dots
        chart_raw.scatterStyle = 'marker'

        chart_raw.x_axis.title = "Pedal Velocity (RPM)"
        chart_raw.y_axis.title = "Power (Watts)"
        
        xvalues = Reference(sheet, min_col=5, min_row=2, max_row=len(rpm_set_points*readings))
        for i in range (10, 11):
            values = Reference(sheet, min_col=i, min_row=1, max_row=len(rpm_set_points*readings))
            series = Series(values, xvalues, title_from_data=True)
            series.marker.symbol = "circle"
            series.graphicalProperties.line.noFill = True
            series.marker.size = 7
            series.smooth = False
            chart_raw.series.append(series)

        # Add a polynomial trendline to the third degree
        trendline = Trendline()#(name=None, trendlineType='poly', order=3, backward=0, forward=0, intercept=0.0, dispEq=True, dispRSqr=False, dispName=False)
        trendline.trendlineType = "poly"
        trendline.order = 3
        trendline.dispEq = True  # Display the equation on the chart
        chart_raw.series[0].trendline = trendline

        chart_raw.type = "scatter"
        chart_raw.width = 25
        chart_raw.height = 12

        sheet.add_chart(chart_raw, "K1")

        # Save Excel workbook
        print("Saving Excel workbook...")
        wb.save(file_name)
        time.sleep(5)
        print("Done!")

        # Specify the sheet name you want to open
        sheet_name = "Averages"
        # Open the saved Excel file with the default application and navigate to the specified sheet
        file_path = "" + file_name  # Replace with your actual file path
        subprocess.run(["start", "excel", file_path, "/e", f"/select,{sheet_name}"], shell=True)

        break

    elif response == "n" or response == "N":
        print("Aborting...")
        # Set Odrive State to Idle
        odrive.axis0.requested_state = AXIS_STATE_IDLE
        break

    else:
        print("Invalid input. Please enter Y or n")
