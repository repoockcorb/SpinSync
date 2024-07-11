from Phidget22.Devices.VoltageRatioInput import VoltageRatioInput
import pandas as pd
import numpy as np
import time
import os

# Initialize an empty DataFrame for data storage
data = pd.DataFrame(columns=['Channel 1 Force (Kilograms)', 'Channel 2 Force (Kilograms)', 'Timestamp'])


def bike_torque(self, force):
    crank_length = 0.170
    torque = force * crank_length 
    print("Torque: " + str(torque) + " Nm")


# Known forces in Newtons and corresponding voltage ratios
known_forces = [0, 78.48, 122.625, 147.15, 220.725, 392.4, 613.125, 784.8, 981]  # List of known forces in Newtons
voltage_ratios = [0.0000315964033291139, 
                    0.00028741835979747,
                    0.00040700983351899,
                    0.00048704213118987,
                    0.00069698462662025,
                    0.00118386609882278,
                    0.00182402247118987,
                    0.00235306166254430,
                    0.00292822179864557]  # List of corresponding voltage ratios

# Perform linear regression to calculate the slope and intercept (offset)
slope, offset = np.polyfit(voltage_ratios, known_forces, 1)

def newtons_to_kilograms(newtons):
    # Conversion factor
    kg_per_newton = 0.10197162129779283
    return newtons * kg_per_newton

i = 1
channel_1_tare = 0
channel_2_tare = 0
def onVoltageRatioChange(self, voltageRatio):
    global i, channel_1_tare, channel_2_tare, tare_values
    # Convert voltage ratio to Newtons using the calibration equation
    # force_newtons = voltageRatio * slope + (offset-2.3544)
    force_newtons = voltageRatio * slope + offset
    force_kilograms = newtons_to_kilograms(force_newtons)
    channel = self.getChannel()
    if channel == 1:
        if i == 1:
            channel_1_tare = force_kilograms
            i = 2
            print("Tare Channel 1")
            time.sleep(5)
            print("Channel 1 Force TARE (Kilograms): {:.2f} kg".format(force_kilograms), flush=True)
            return None
        data.at[0, 'Channel 1 Force (Kilograms)'] = abs(force_kilograms) - abs(channel_1_tare)
        print("Channel [" + str(channel) + "] Force (Kilograms): {:.2f} kg".format(abs(force_kilograms) - abs(channel_1_tare)), flush=True)
    elif channel == 2:
        if i ==2:
            channel_2_tare = force_kilograms
            i = 0
            print("Tare Channel 2")
            time.sleep(5)
            print("Channel 2 Force TARE (Kilograms): {:.2f} kg".format(force_kilograms), flush=True)
            return None
        data.at[0, 'Channel 2 Force (Kilograms)'] = abs(force_kilograms) - abs(channel_2_tare)
        print("Channel [" + str(channel) + "] Force (Kilograms): {:.2f} kg".format(abs(force_kilograms) - abs(channel_2_tare)), flush=True)
    data.at[0, 'Timestamp'] = time.strftime('%Y-%m-%d %H:%M:%S')    
    log_data()

def log_data():
    global data  # Use the globally defined DataFrame

    # Create a new DataFrame with the data
    new_data = pd.DataFrame({
        'Channel 1 Force (Kilograms)': [data.at[0, 'Channel 1 Force (Kilograms)']],
        'Channel 2 Force (Kilograms)': [data.at[0, 'Channel 2 Force (Kilograms)']],
        'Timestamp': [data.at[0, 'Timestamp']]
    })

    # Concatenate the new DataFrame with the existing data
    data = pd.concat([data, new_data], ignore_index=True)
    
    # Save the DataFrame to the Excel file
    data.to_excel('output/force data test 5.xlsx', index=False)


def verify_tare(input1, input2):
    ratio1 = input1.getVoltageRatio()
    ratio2 = input2.getVoltageRatio()
    print("Channel 1 Tare Value: {:.6f}".format(ratio1))
    print("Channel 2 Tare Value: {:.6f}".format(ratio2))


def main():
    # Create your Phidget channels
    voltageRatioInput1 = VoltageRatioInput()  # Use input 1
    voltageRatioInput2 = VoltageRatioInput()  # Use input 2

    # Set addressing parameters to specify which channel to open (in this case, input 1 and input 2)
    voltageRatioInput1.setChannel(1)
    voltageRatioInput2.setChannel(2)

    # Open your Phidget and wait for attachment
    voltageRatioInput1.openWaitForAttachment(5000)
    voltageRatioInput2.openWaitForAttachment(5000)

    input("Press Enter to Start...")

    # Assign the event handler to log and export data
    voltageRatioInput1.setOnVoltageRatioChangeHandler(onVoltageRatioChange)
    voltageRatioInput2.setOnVoltageRatioChangeHandler(onVoltageRatioChange)

    try:
        input("Press Enter to Stop\n")
    except (Exception, KeyboardInterrupt):
        pass

    # Close your Phidget once the program is done
    voltageRatioInput1.close()
    voltageRatioInput2.close()

if __name__ == "__main__":
    main()






# from Phidget22.Phidget import *
# from Phidget22.Devices.VoltageRatioInput import *
# import time

# # Declare any event handlers here. These will be called every time the associated event occurs.
# readings = []

# def onVoltageRatioChange(self, voltageRatio):
#     readings.append(voltageRatio)
#     print("Reading: " + str(voltageRatio))

# def main():
#     while True:
#         # Create your Phidget channels
#         voltageRatioInput1 = VoltageRatioInput()

#         # Set addressing parameters to specify which channel to open (if any)
#         voltageRatioInput1.setChannel(1)

#         # Assign any event handlers you need before calling open so that no events are missed.
#         voltageRatioInput1.setOnVoltageRatioChangeHandler(onVoltageRatioChange)

#         input("Press Enter to start taking readings...")

#         # Open your Phidgets and wait for attachment
#         voltageRatioInput1.openWaitForAttachment(5000)

#         # Collect 20 readings
#         for _ in range(20):
#             time.sleep(1)  # Wait for 1 second for each reading

#         # Close your Phidgets once the readings are done.
#         voltageRatioInput1.close()

#         # Calculate and print the average
#         if len(readings) > 0:
#             average = sum(readings) / len(readings)
#             print("Average VoltageRatio: " + str(average))

#         readings.clear()  # Clear the readings list for the next round

#         again = input("Press Enter to take more readings or 'q' to quit: ")
#         if again.lower() == 'q':
#             break

# if __name__ == "__main__":
#     main()






















# from Phidget22.Phidget import *
# from Phidget22.Devices.VoltageRatioInput import *
# import time

# #Declare any event handlers here. These will be called every time the associated event occurs.

# def onVoltageRatioChange(self, voltageRatio):
# 	print("VoltageRatio: " + str(voltageRatio))

# def main():
# 	#Create your Phidget channels
# 	voltageRatioInput1 = VoltageRatioInput()

# 	#Set addressing parameters to specify which channel to open (if any)
# 	voltageRatioInput1.setChannel(1)

# 	#Assign any event handlers you need before calling open so that no events are missed.
# 	voltageRatioInput1.setOnVoltageRatioChangeHandler(onVoltageRatioChange)

# 	#Open your Phidgets and wait for attachment
# 	voltageRatioInput1.openWaitForAttachment(5000)

# 	#Do stuff with your Phidgets here or in your event handlers.

# 	try:
# 		input("Press Enter to Stop\n")
# 	except (Exception, KeyboardInterrupt):
# 		pass

# 	#Close your Phidgets once the program is done.
# 	voltageRatioInput1.close()

# main()












# from Phidget22.Devices.VoltageRatioInput import VoltageRatioInput
# import numpy as np

# # Known forces in Newtons and corresponding voltage ratios
# known_forces = [117.6798, 833.56525]  # List of known forces in Newtons
# voltage_ratios = [0.00041301269, 0.002570771612]  # List of corresponding voltage ratios

# # Perform linear regression to calculate the slope and intercept (offset)
# slope, offset = np.polyfit(voltage_ratios, known_forces, 1)

# def onVoltageRatioChange(self, voltageRatio):
#     # Convert voltage ratio to Newtons using the calibration equation
#     force = voltageRatio * slope + offset
#     print("Force (Newtons): {:.2f}".format(force))

# def main():
#     # Create your Phidget channels
#     voltageRatioInput1 = VoltageRatioInput()  # Use input 1

#     # Set addressing parameters to specify which channel to open (in this case, input 1)
#     voltageRatioInput1.setChannel(1)

#     # Assign any event handlers you need before calling open so that no events are missed.
#     voltageRatioInput1.setOnVoltageRatioChangeHandler(onVoltageRatioChange)

#     # Open your Phidgets and wait for attachment
#     voltageRatioInput1.openWaitForAttachment(5000)

#     # Do stuff with your Phidgets here or in your event handlers.

#     try:
#         input("Press Enter to Stop\n")
#     except (Exception, KeyboardInterrupt):
#         pass

#     # Close your Phidgets once the program is done.
#     voltageRatioInput1.close()

# if __name__ == "__main__":
#     main()

























# from Phidget22.Devices.VoltageRatioInput import VoltageRatioInput
# import numpy as np


# # Known forces and corresponding voltage ratios
# # known_forces = [12, 85]  # List of known forces
# known_forces = [117.6798, 833.56525]  # List of known forces
# voltage_ratios = [0.00041301269, 0.002570771612]  # List of corresponding voltage ratios

# # Perform linear regression to calculate the slope and intercept (offset)
# slope, offset = np.polyfit(known_forces, voltage_ratios, 1)


# def onVoltageRatioChange(self, voltageRatio):
#     # Convert voltage ratio to Newtons using the calibration equation
#     force = voltageRatio * slope + offset
#     print("Force (Newtons): {:.2f}".format(force))

# def main():
#     # Create your Phidget channels
#     voltageRatioInput1 = VoltageRatioInput()  # Use input 1

#     # Set addressing parameters to specify which channel to open (in this case, input 1)
#     voltageRatioInput1.setChannel(1)

#     # Assign any event handlers you need before calling open so that no events are missed.
#     voltageRatioInput1.setOnVoltageRatioChangeHandler(onVoltageRatioChange)

#     # Open your Phidgets and wait for attachment
#     voltageRatioInput1.openWaitForAttachment(5000)

#     # Do stuff with your Phidgets here or in your event handlers.

#     try:
#         input("Press Enter to Stop\n")
#     except (Exception, KeyboardInterrupt):
#         pass

#     # Close your Phidgets once the program is done.
#     voltageRatioInput1.close()

# if __name__ == "__main__":
#     main()
