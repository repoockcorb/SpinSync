import time
import board
import digitalio

# Initialize variables
revolutions = 0
last_button_state = True  # Set the initial state to True (unpressed)
last_time = time.monotonic()
last_time_print_rate = time.monotonic()
last_time_zero_gauges = time.monotonic()

update_interval = 0.1  # Update interval in seconds (50 ms)

# Define the button pin as an interrupt source
button = digitalio.DigitalInOut(board.GP10)
button.switch_to_input(pull=digitalio.Pull.UP)
fan_gear_ratio = (60 * (1/7))
rev_per_sec = 0
rev_per_min = 0 / fan_gear_ratio
print("Revolutions per minute (rpm):", rev_per_min)    

rpm_history = [0] * 10  # Initialize rpm_history with default values
counter = 0

while True:
    current_time = time.monotonic()
    current_time_print_rate = time.monotonic()      

    if button.value == 1:
        while True:      
            current_time_print_rate = time.monotonic()      
            if current_time - last_time >= update_interval:
                rev_per_sec = revolutions / (current_time - last_time)
                revolutions = 0  # Reset the count
                last_time = current_time

            if button.value == 0:
                revolutions += 1
                break

            if current_time_print_rate - last_time_print_rate >= update_interval:
                print("Revolutions per minute (rpm):", (rev_per_sec/7)*60/10)
                last_time_print_rate = current_time_print_rate
                rpm_history[counter] = ((rev_per_sec/7)*60/10)
                counter += 1
                if counter >= 10:
                    counter = 0
                if all(rpm == rpm_history[0] for rpm in rpm_history[:-1]):
                    rev_per_sec = 0

    if current_time_print_rate - last_time_print_rate >= update_interval:
        print("Revolutions per minute (rpm):", (rev_per_sec/7)*60/10)
        last_time_print_rate = current_time_print_rate     
        rpm_history[counter] = ((rev_per_sec/7)*60/10)
        counter += 1
        if counter >= 10:
            counter = 0
        if all(rpm == rpm_history[0] for rpm in rpm_history[:-1]):
            rev_per_sec = 0



# ############## BACKUP WORKING ####################

# import time
# import board
# import digitalio

# # Initialize variables
# revolutions = 0
# last_button_state = True  # Set the initial state to True (unpressed)
# last_time = time.monotonic()
# #update_interval = 0.0001  # Update interval in seconds (10 ms)
# # update_interval = 0.01
# update_interval = 0.1



# # Define the button pin as an interrupt source
# button = digitalio.DigitalInOut(board.GP10)
# button.switch_to_input(pull=digitalio.Pull.UP)

# while True:
#     current_time = time.monotonic()
#     if button.value == 0:
#         #print("button pressed")
#         while True:
#             if current_time - last_time >= update_interval:
#                 rev_per_sec = revolutions / (current_time - last_time)
#                 revolutions = 0  # Reset the count
#                 print(((rev_per_sec/7)*60)/10)
#                 #print("Revolutions per minute (rpm):", (rev_per_sec/7)*60)
#                 last_time = current_time
#             if button.value == 1:
#                 #print("waiting for release")
#                 revolutions += 1
#                 break