import os

# Import the Timer from the threading module
from threading import Timer

import pywinstyles

import tkinter as tk
import customtkinter as ctk
from CTkMessagebox import CTkMessagebox

from customtkinter import set_default_color_theme
from tkdial import Meter

import serial
import csv
import threading
import time
from serial.tools.list_ports import comports
import ctypes
import webbrowser

import PIL
from PIL import Image


class MovingAverageFilter:
    def __init__(self, window_size):
        self.window_size = window_size
        self.values = []

    def add_value(self, value):
        self.values.append(value)
        if len(self.values) > self.window_size:
            self.values.pop(0)

    def get_smoothed_value(self):
        if not self.values:
            return None
        return sum(self.values) / len(self.values)

class MyInterface:
    def __init__(self, master):
        self.master = master
        self.master.title("SpinSync")

        self.com_ports = [port.device for port in comports()]  # Fetch available COM ports

        self.bike = []

        self.auto_start_delay_flag = False # Flag to indicate if auto start delay is triggered
        self.logging_active = False  # Flag to indicate whether logging is active

        self.live_update_flag = True  # Flag to control live update thread

        set_default_color_theme("dark-blue")
        ctk.set_appearance_mode("dark")

        self.setup_ui()

         # Your existing initialization code here
        self.moving_avg_filter = MovingAverageFilter(window_size=8)  # Create an instance of MovingAverageFilter

    
    def setup_ui(self):
        
        image = PIL.Image.open("images/background_image.png")
        background_image = ctk.CTkImage(image, size=(1242, 786))

        # Create a bg label
        bg_lbl = ctk.CTkLabel(self.master, text="", image=background_image)
        bg_lbl.place(x=0, y=0)

        # Header Frame
        header_frame = ctk.CTkFrame(master=self.master, bg_color="#000001", fg_color="#000001")  # Use CTkFrame
        pywinstyles.set_opacity(header_frame, color="#000001") # just add this line
        header_frame.pack(pady=10, padx=10)

        IMAGE_WIDTH = 255*1.5
        IMAGE_HEIGHT = 68.2*1.5

        image = ctk.CTkImage(light_image=Image.open("images/SpinSync_logo.png"), dark_image=Image.open("images/SpinSync_logo.png"), size=(IMAGE_WIDTH , IMAGE_HEIGHT))
        

        # Create a label to display the image
        image_label = ctk.CTkLabel(header_frame, image=image, text='', corner_radius=40)
        # pywinstyles.set_opacity(image_label, value=1, color="#000001") # just add this line
        
        image_label.grid(row=0, column=0, columnspan=3, pady=10, padx=10)  # Span across all columns
        

        # Create frame for the top section
        com_port_frame = ctk.CTkFrame(self.master, bg_color="#000001", fg_color="#000001")  # Use CTkFrame
        pywinstyles.set_opacity(com_port_frame, color="#000001") # just add this line
        com_port_frame.place(x=45, y=80-5)
        # com_port_frame.pack(pady=10)

        # Create com port selection dropdowns for each anemometer
        self.com_port_dropdowns = []
        self.clear_port_buttons = []  # List to store the clear port buttons


        com_port_label = ctk.CTkLabel(com_port_frame, text=f"Com Port:", width=100, bg_color="#000001", font=("Arial", 16, "bold"), text_color="white")
        com_port_label.grid(row=0, column=0, padx=0, pady=20)

        com_port_var = ctk.StringVar()
        # Add a blank option at the beginning
        com_ports_with_blank = [''] + self.com_ports

        dropdown = ctk.CTkComboBox(com_port_frame, values=com_ports_with_blank, variable=com_port_var, width=110, justify="center")
        dropdown.grid(row=0, column=1, padx=5, pady=5)


        # refresh_image = ctk.CTkImage(Image.open(r"images/refresh-arrow.png"))
        refresh_image = ctk.CTkImage(Image.open(r"images/refresh_arrow_white.png"))

        refresh_button = ctk.CTkButton(com_port_frame, text="", width=5, bg_color="#000001", fg_color="grey40", font=("Arial", 12, "bold"), corner_radius=20, command=lambda: dropdown.configure(values=[''] + [port.device for port in comports()]), image=refresh_image)
        refresh_button.grid(row=0, column=2, padx=5, pady=5)

        self.com_port_dropdowns.append(com_port_var)


        # Create frame for the top section
        participants_ID_frame = ctk.CTkFrame(self.master, bg_color="#000001", fg_color="#000001")  # Use CTkFrame
        pywinstyles.set_opacity(participants_ID_frame, color="#000001") # just add this line
        participants_ID_frame.place(x=45, y=130)

        # Create input fields in a grid form for min and max values
        self.participants_ID = ctk.CTkEntry(participants_ID_frame, width=270, placeholder_text="Participants ID (Output File Prefix)")
        self.participants_ID.grid(row=0, column=0, padx=5, pady=5)

        # Create frame for input ranges
        input_range_frame = ctk.CTkFrame(self.master, bg_color="#000001", fg_color="#000001")  # Use CTkFrame
        pywinstyles.set_opacity(input_range_frame, color="#000001") # just add this line
        input_range_frame.place(x=45, y=150+20)

        # Create input fields in a grid form for min and max values
        self.min_rpm_entry = ctk.CTkEntry(input_range_frame, width=270/2-5, placeholder_text="Min RPM")
        self.min_rpm_entry.grid(row=1, column=0, padx=5, pady=5)

        self.max_rpm_entry = ctk.CTkEntry(input_range_frame, width=270/2-5, placeholder_text="Max RPM")
        self.max_rpm_entry.grid(row=1, column=1, padx=5, pady=5)

        self.min_power_entry = ctk.CTkEntry(input_range_frame, width=270/2-5, placeholder_text="Min Power")
        self.min_power_entry.grid(row=2, column=0, padx=5, pady=5)

        self.max_power_entry = ctk.CTkEntry(input_range_frame, width=270/2-5, placeholder_text="Max Power")
        self.max_power_entry.grid(row=2, column=1, padx=5, pady=5)

        
        # Create input field for runtime
        runtime_frame = ctk.CTkFrame(self.master, bg_color="#000001", fg_color="#000001")
        pywinstyles.set_opacity(runtime_frame, color="#000001") # just add this line
        runtime_frame.place(x=45, y=225+20)

        self.runtime_entry = ctk.CTkEntry(runtime_frame, width=270, placeholder_text="RunTime (seconds)")
        self.runtime_entry.grid(row=3, column=0, padx=5, pady=5)

         # Create frame for auto/manual mode
        mode_frame = ctk.CTkFrame(self.master, bg_color="#000001", fg_color="#000001")
        pywinstyles.set_opacity(mode_frame, color="#000001") # just add this line
        mode_frame.place(x=45, y=262+20)

        # Create radio buttons for auto/manual mode
        auto_start = ctk.StringVar(value="off")
        self.auto_start_switch = ctk.CTkSwitch(mode_frame, text="Auto Start", width=100, bg_color="#000001", fg_color="white", font=("Arial", 12, "bold"), text_color="white",
                                 variable=auto_start, onvalue="on", offvalue="off", progress_color="#28a745")
        self.auto_start_switch.grid(row=0, column=0, padx=5, pady=5)

        self.auto_start_delay = ctk.CTkEntry(mode_frame, width=270-100-10, placeholder_text="Start Delay (seconds)")
        self.auto_start_delay.grid(row=0, column=1, padx=5, pady=5)


        # Create four buttons stacked vertically
        button_names = ["Connect", "Start Logging", "Stop Logging", "Reset"]
        commands = [self.connect_bike, self.start_logging, self.stop_logging, self.reset_display]
        button_colour = ["#28a745", "#007bff", "#dc3545", "#cc8400"]
        start_button_position_x = 50
        # start_button_position_y = 150
        start_button_position_y = 310+20

        button_positions_y = [start_button_position_y+0, start_button_position_y+37, start_button_position_y+74, start_button_position_y+111]
        self.buttons = []
        for name, command, colour, position_y in zip(button_names, commands, button_colour, button_positions_y):
            button = ctk.CTkButton(self.master, text=name, command=command, hover_color="grey", width=270, fg_color=colour, corner_radius=20, bg_color="#000001")
            button.pack(pady=5, padx = 20)  # Use pack with pady for vertical spacing
            self.buttons.append(button)
            self.buttons[-1].place(x=start_button_position_x, y= position_y)
            pywinstyles.set_opacity(button, color="#000001") # just add this line


        # Create frame for the bottom section (terminal)
        terminal_frame = ctk.CTkFrame(master=self.master, bg_color="#000001", fg_color="#000001")  # Use CTkFrame
        pywinstyles.set_opacity(terminal_frame, color="#000001") # just add this line
        terminal_frame.place(x=35, y=452+20)
        # terminal_frame.pack(pady=10, padx=10)

        # Terminal (text output)
        self.terminal = ctk.CTkTextbox(terminal_frame, height=220, width=350, corner_radius=20)
        self.terminal.pack(pady=10, padx=10)


        # Create frame for the footer section with a larger width
        footer_frame = ctk.CTkFrame(master=self.master, width=200, bg_color="#000001", corner_radius=20)  # Set a larger width
        # pywinstyles.set_opacity(footer_frame, value=0.85, color="#000001") # just add this line
        pywinstyles.set_opacity(footer_frame, value=0.85, color="#000001") # just add this line

        # footer_frame.pack(pady=10, padx=10)  # Use fill='x' to make the frame fill the entire width
        footer_frame.place(x=35, y=720)

        # Developer label
        developer_label = ctk.CTkLabel(footer_frame, text="Developed By: ", anchor="w", font=("Arial", 12, "bold"), text_color="white")
        developer_label.grid(row=0, column=0, sticky="w", padx=10, pady=5)  # Adjust padx as needed

        # Developer's name with hyperlink
        developer_name_label = ctk.CTkLabel(footer_frame, text="Brock Cooper", anchor="w", cursor="hand2", text_color="#007bff", font=("Arial", 12, "bold"))
        developer_name_label.grid(row=0, column=0, sticky="w", padx=95, pady=5)  # Adjust padx as needed
        developer_name_label.bind("<Button-1>", lambda event: self.open_website("https://brockcooper.au"))

        # space
        space_label = ctk.CTkLabel(footer_frame, text="", anchor="e")
        space_label.grid(row=0, column=2, sticky="e", padx=454, pady=5)  # Adjust padx as needed

        # Version label
        version_label = ctk.CTkLabel(footer_frame, text="Version 1.0", anchor="e", font=("Arial", 12, "bold"), text_color="white")
        version_label.grid(row=0, column=2, sticky="e", padx=10, pady=5)  # Adjust padx as needed


        # Gauges Frame
        gauges_frame = ctk.CTkFrame(master=self.master, bg_color="#000001", fg_color="#000001")  # Set a larger width
        pywinstyles.set_opacity(gauges_frame, color="#000001") # just add this line
        gauges_frame.place(x=380, y=130)

        self.rpm_gauge = Meter(gauges_frame, radius=400, start=0, end=100, border_width=10, major_divisions=5, minor_divisions= 0.5,
                    fg="white", text_color="black", start_angle=270, end_angle=-270,
                    text_font="DS-Digital 30", scale_color="black", needle_color="orange", border_color="grey30", bg="#000001")
        self.rpm_gauge.set_mark(2*0, 2*500, color="#FFFFFF")
        # self.rpm_gauge.set_mark(2*30, 2*40, color="#28a745") # set red marking from 140 to 160
        self.rpm_gauge.grid(row=0, column=0, pady=5, padx=5)
        pywinstyles.set_opacity(self.rpm_gauge, color="#000001") # just add this line

        self.power_gauge = Meter(gauges_frame, radius=400, start=0, end=900, border_width=10, major_divisions=50, minor_divisions=2,
            fg="white", text_color="black", start_angle=270, end_angle=-270,
            text_font="DS-Digital 30", scale_color="black", needle_color="orange", border_color="grey30", bg="#000001")
        self.power_gauge.set_mark(0, 5000, color="#FFFFFF")
        
        # self.power_gauge.set_mark(round(38.4/2), round(48/2), color="#28a745") # set red marking from 140 to 160
        self.power_gauge.grid(row=0, column=1, pady=5, padx=5)
        pywinstyles.set_opacity(self.power_gauge, color="#000001") # just add this line

        rpm_gauge_label = ctk.CTkLabel(gauges_frame, text=f"Bike Pedal Speed (rpm)", width=100, bg_color="#000001", font=("Arial", 22, "bold"), text_color="white")
        rpm_gauge_label.grid(row=1, column=0, pady=5, padx=5)

        power_gauge_label = ctk.CTkLabel(gauges_frame, text=f"Bike Pedal Power (watts)", width=100, bg_color="#000001", font=("Arial", 22, "bold"), text_color="white")
        power_gauge_label.grid(row=1, column=1, pady=5, padx=5)



        # # Create frame for the bottom section (terminal)
        # terminal_frame = ctk.CTkFrame(master=self.master, bg_color="#000001", fg_color="#000001")  # Use CTkFrame
        # pywinstyles.set_opacity(terminal_frame, value=0.9, color="#000001") # just add this line
        # terminal_frame.place(x=380, y=600)
        # # terminal_frame.pack(pady=10, padx=10)

        # # Terminal (text output)
        # self.terminal = ctk.CTkTextbox(terminal_frame, height=220, width=350, corner_radius=20)
        # self.terminal.pack(pady=10, padx=10)


        # Countdown Frame
        countdown_frame = ctk.CTkFrame(master=self.master, bg_color="#000001", corner_radius=20)  # Set a larger width
        pywinstyles.set_opacity(countdown_frame, value=0.85, color="#000001") # just add this line
        countdown_frame.place(x=720, y=600)

        # countdown label
        self.countdown_label = ctk.CTkLabel(countdown_frame, text="Timer\n0", width=100, font=("Arial", 42, "bold"), text_color="white", corner_radius=20)
        self.countdown_label.grid(row=0, column=0, pady=5, padx=5)


        # Handle window closing event
        self.master.protocol("WM_DELETE_WINDOW", self.on_close)
 

    def open_website(self, url):
            webbrowser.open_new(url)
            
    def change_theme(self, choice):
        ctk.set_appearance_mode(choice)
        
        # # Update the meter text color based on the chosen theme
        # if choice == "dark":
        #     self.rpm_gauge.configure(fg="black", text_color="white", scale_color="white", needle_color="red")
        #     self.rpm_gauge.update()  # Update/redraw the gauge
        # else:
        #     self.rpm_gauge.configure(start=0, end=85, border_width=10, major_divisions=5, minor_divisions= 0.5,
        #             fg="white", text_color="black", start_angle=270, end_angle=-270,
        #             text_font="DS-Digital 30", scale_color="black", needle_color="red", border_color="grey30", bg="#000001")
        #     self.rpm_gauge.update()  # Update/redraw the gauge


    def convert_fields(self):
        try:
            min_rpm = int(self.min_rpm_entry.get()) if self.min_rpm_entry.get() else 0
            max_rpm = int(self.max_rpm_entry.get()) if self.max_rpm_entry.get() else 0
            min_power = int(self.min_power_entry.get()) if self.min_power_entry.get() else 0
            max_power = int(self.max_power_entry.get()) if self.max_power_entry.get() else 0
            start_delay = int(self.auto_start_delay.get()) if self.auto_start_delay.get() else 0
            run_time = int(self.runtime_entry.get()) if self.runtime_entry.get() else 0

        except ValueError:
            self.update_terminal("Error: Please enter valid numerical values\n")
            raise ValueError("Please enter valid numerical values")
            

        if run_time == 0:
            self.countdown_label.configure(text=f"Timer\n♾️")
        elif run_time != 0 and self.min_rpm_entry.get() and self.max_rpm_entry.get() or self.min_power_entry.get() and self.max_power_entry.get():
            self.countdown_label.configure(text=f"Timer\n{run_time}")
        elif run_time != 0:
            self.countdown_label.configure(text=f"Timer\n{run_time}")
        else:
            self.countdown_label.configure(text=f"Timer\n♾️")


            # Check if both RPM and power fields are empty
        if not (self.min_rpm_entry.get() or self.max_rpm_entry.get()) and not (self.min_power_entry.get() or self.max_power_entry.get()):
            self.update_terminal(f"Warning: You have not entered any min or max values.\nPlease ignore this message if you don't require any\nmin or max values.\n")

        if self.min_rpm_entry.get() and self.max_rpm_entry.get() and self.min_power_entry.get() and self.max_power_entry.get():
            self.update_terminal(f"Error: Please enter either\n(RPM min and max) or (POWER min and max) values.\n")
            raise ValueError("Please enter valid numerical values")
        else:
            if self.min_rpm_entry.get() and self.max_rpm_entry.get():
                min_rpm = round(float(self.min_rpm_entry.get()),1)
                max_rpm = round(float(self.max_rpm_entry.get()),1)

                min_power = self.power_conversion(min_rpm)
                max_power = self.power_conversion(max_rpm)

                self.min_power_entry.delete(0,100)
                self.max_power_entry.delete(0,100)
                self.min_power_entry.configure(placeholder_text=min_power)
                self.max_power_entry.configure(placeholder_text=max_power)
                self.rpm_gauge.set_mark(2*0, 2*120, color="#FFFFFF")
                self.power_gauge.set_mark(0, 5000, color="#FFFFFF")
                self.rpm_gauge.set_mark(2*int(min_rpm), 2*int(max_rpm), color="#28c74c")
                self.power_gauge.set_mark(round(min_power/2), round(max_power/2), color="#28c74c") # set red marking from 140 to 160
                self.rpm_gauge.update()
                self.power_gauge.update()

                if start_delay == 0 and run_time != 0:
                    self.update_terminal(f"Paramaters set using rpm input:\nMin RPM: {min_rpm}\nMax RPM: {max_rpm}\nMin Power (watts): {min_power}\nMax Power (watts): {max_power}\nRuntime: {run_time} seconds\n")
                elif run_time == 0 and start_delay != 0:
                    self.update_terminal(f"Paramaters set using rpm input:\nMin RPM: {min_rpm}\nMax RPM: {max_rpm}\nMin Power (watts): {min_power}\nMax Power (watts): {max_power}\nStart Delay: {start_delay} seconds\n")
                elif run_time == 0 and start_delay == 0:
                    self.update_terminal(f"Paramaters set using rpm input:\nMin RPM: {min_rpm}\nMax RPM: {max_rpm}\nMin Power (watts): {min_power}\nMax Power (watts): {max_power}\n") 
                else:
                    self.update_terminal(f"Paramaters set using rpm input:\nMin RPM: {min_rpm}\nMax RPM: {max_rpm}\nMin Power (watts): {min_power}\nMax Power (watts): {max_power}\nRuntime: {run_time} seconds\nStart Delay: {start_delay} seconds\n")

            elif self.min_power_entry.get() and self.max_power_entry.get():

                min_power = float(self.min_power_entry.get())
                max_power = float(self.max_power_entry.get())

                min_rpm = round(self.bisection_method(min_power, 0, 120),1)
                max_rpm = round(self.bisection_method(max_power, 0, 120),1)

                self.min_rpm_entry.delete(0,100)
                self.max_rpm_entry.delete(0,100)
                self.min_rpm_entry.configure(placeholder_text=min_rpm)
                self.max_rpm_entry.configure(placeholder_text=max_rpm)
                self.rpm_gauge.set_mark(2*0, 2*120, color="#FFFFFF")
                self.power_gauge.set_mark(0, 5000, color="#FFFFFF")
                self.rpm_gauge.set_mark(2*int(round(float(min_rpm),1)), 2*int(round(float(max_rpm),1)), color="#28c74c")
                self.power_gauge.set_mark(round(float(self.min_power_entry.get())/2), round(float(self.max_power_entry.get())/2), color="#28c74c") # set red marking from 140 to 160
                self.rpm_gauge.update()
                self.power_gauge.update()

                if start_delay == 0 and run_time != 0:
                    self.update_terminal(f"Paramaters set using power input:\nMin RPM: {min_rpm}\nMax RPM: {max_rpm}\nMin Power (watts): {min_power}\nMax Power (watts): {max_power}\nRuntime: {run_time} seconds\n")
                elif run_time == 0 and start_delay != 0:
                    self.update_terminal(f"Paramaters set using power input:\nMin RPM: {min_rpm}\nMax RPM: {max_rpm}\nMin Power (watts): {min_power}\nMax Power (watts): {max_power}\nStart Delay: {start_delay} seconds\n")
                elif run_time == 0 and start_delay == 0:
                    self.update_terminal(f"Paramaters set using power input:\nMin RPM: {min_rpm}\nMax RPM: {max_rpm}\nMin Power (watts): {min_power}\nMax Power (watts): {max_power}\n")   
                else:
                    self.update_terminal(f"Paramaters set using power input:\nMin RPM: {min_rpm}\nMax RPM: {max_rpm}\nMin Power (watts): {min_power}\nMax Power (watts): {max_power}\nRuntime: {run_time} seconds\nStart Delay: {start_delay} seconds\n")

        self.min_rpm_entry.configure(state="disabled")
        self.max_rpm_entry.configure(state="disabled")
        self.min_power_entry.configure(state="disabled")
        self.max_power_entry.configure(state="disabled")
        self.auto_start_delay.configure(state="disabled")
        self.runtime_entry.configure(state="disabled")
        self.auto_start_switch.configure(state="disabled")
        self.participants_ID.configure(state="disabled")


    def connect_bike(self):   
            self.live_update_flag = False  # Reset live update flag

            self.update_gauges(0)

            # Clear terminal
            self.clear_terminal()
            # Close all serial connections
            for ser in self.bike:
                ser.close()
            self.bike = []

            self.convert_fields()

            # Open serial connections to anemometers
            for i, com_port_var in enumerate(self.com_port_dropdowns):
                com_port = com_port_var.get()
                if com_port:
                    try:
                        ser = serial.Serial(com_port, 9600)
                        self.bike.append(ser)
                        # Read initial data from the serial port
                        # initial_line = ser.readline().decode().strip()
                        initial_line = ser.readline().strip().decode()
                        # initial_line = ser.readline().strip().decode('utf-8')

                        initial_line = initial_line.replace("Revolutions per minute (rpm): ", "")

                        if initial_line:
                            # Split the line and extract the bike speed if data is available
                            self.update_terminal(f"Connected to Rogue Echo Bike on port {com_port}\n")
                            self.live_update_flag = True  # Reset live update flag

                            # Start a thread to continuously update wind speed
                            threading.Thread(target=self.update_bike_speed_live, args=(ser, i)).start()
                        else:
                            self.update_terminal(f"Failed to connect to Rogue Echo Bike on port {com_port}\n")
                            
                    except serial.SerialException:
                        self.update_terminal(f"Failed to connect to Rogue Echo Bike on port {com_port}\n")
                        continue


    def update_bike_speed_live(self, ser, index):
        while self.live_update_flag:
            try:
                # line = ser.readline().strip().decode('utf-8')
                line = ser.readline().strip().decode()
                
    
                bike_rpm = line.replace("Revolutions per minute (rpm): ", "")

                self.update_gauges(bike_rpm)

                # self.rpm_gauge.set(round(float(bike_rpm),1))
                # self.rpm_gauge.update()
                # self.power_gauge.set(self.power_conversion(bike_rpm))
                # self.rpm_gauge.update()       

                if self.min_rpm_entry.get() and self.max_rpm_entry.get() and self.auto_start_switch.get() == "on" and self.auto_start_delay_flag == False:
                    if float(bike_rpm) >= float(self.min_rpm_entry.get()) and float(bike_rpm) <= float(self.max_rpm_entry.get()):
                        self.auto_start_delay_flag = True
                        # Delay for 5 seconds without freezinging the program
                        self.start_logging_with_delay()       

                if self.min_power_entry.get() and self.max_power_entry.get() and self.auto_start_switch.get() == "on" and self.auto_start_delay_flag == False:
                    if float(self.power_conversion(bike_rpm)) >= float(self.min_power_entry.get()) and float(self.power_conversion(bike_rpm)) <= float(self.max_power_entry.get()):
                        self.auto_start_delay_flag = True
                        # Delay for 5 seconds without freezinging the program
                        self.start_logging_with_delay()     


            except (ValueError, UnicodeDecodeError, IndexError):
                pass

    
    def start_logging_with_delay(self):
        auto_start_delay = self.auto_start_delay.get()
        if self.auto_start_delay.get() == "":
            self.start_logging()
        elif auto_start_delay.isdigit():
            Timer(int(auto_start_delay), self.start_logging).start()
        # else:
        #     self.update_terminal("Invalid delay value. Please enter a valid number\n")

    def start_logging(self):
        self.live_update_flag = False
        # Check if at least one serial connection is established
        if not self.bike:
            self.update_terminal("No serial connection established. Please connect Rogue Echo Bike.\n")
            return
        
        if self.logging_active == True:
            self.update_terminal("Logging already active\n")
            return
    
        # Generate file name based on current date and time
        current_datetime = time.strftime("(%d-%m-%Y)_(%H-%M-%S)")
        folder_path = os.path.join(os.getcwd(), "SpinSync Logs")  # Folder path one level deep
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)  # Create folder if it doesn't exist
        file_name = os.path.join(folder_path, f"{self.participants_ID.get()} - SpinSync_Data_{current_datetime}.csv")

        self.live_update_flag = False  # Stop live update thread

        # Start a thread to handle data logging
        self.logging_active = True
        logging_thread = threading.Thread(target=self.log_data, args=(file_name,))
        logging_thread.start()

        # Pass the file name to stop_logging when logging stops
        # self.stop_button.configure(command=lambda: self.stop_logging(file_name))

    def stop_logging(self):
        if self.logging_active == True:
            # Stop logging
            self.logging_active = False
            self.live_update_flag = True
            self.auto_start_delay_flag = False
            self.min_rpm_entry.configure(state= "normal")
            self.max_rpm_entry.configure(state= "normal")
            self.min_power_entry.configure(state= "normal")
            self.max_power_entry.configure(state= "normal")
            self.auto_start_delay.configure(state= "normal")
            self.runtime_entry.configure(state= "normal")
            self.auto_start_switch.configure(state= "normal")
            self.participants_ID.configure(state="normal")
        else:
            self.update_terminal("Logging not active\n")
            

    # def update_stopwatch(self, start_time):
    #     elapsed_time = time.time() - start_time
    #     minutes = int(elapsed_time // 60)
    #     seconds = int(elapsed_time % 60)
    #     milliseconds = int((elapsed_time - seconds) * 1000)
    #     return f'{minutes:02d}:{seconds:02d}:{milliseconds:03d}'









    def log_data(self, file_name):
        # Write column headings to CSV
        with open(file_name, "w", newline="") as csvfile:
            csv_writer = csv.writer(csvfile)
            # Write headings
            csv_writer.writerow(["Timestamp", "Time (sec)", "Pedal Rpm", "Power (watts)", "Filtered Rpm", "Filtered Power (watts)"])
            time.sleep(0.2)

        start_time = time.time()  # Record the start time
        total_runtime = self.get_total_runtime()  # Get total runtime

        time_sec = 0
        next_record_time = 0.5  # Start recording at 0.5 seconds
        last_data_record_time = start_time  # Track when we last recorded data

        countdown_time_update = start_time

        initial_runtime = int(self.runtime_entry.get()) if self.runtime_entry.get() else 0
        countdown = int(self.runtime_entry.get()) if self.runtime_entry.get() else 0

        # Continuously read data from anemometers and log it
        while self.logging_active and self.master.winfo_exists():
            # Calculate elapsed time
            elapsed_time = time.time() - start_time

            # Check if total runtime exceeded (allow for the last 0.5-second interval)
            if elapsed_time >= total_runtime + 0.5:
                self.stop_logging()
                break

            row_data = []  # Initialize row data

            # Read data from serial connection
            try:
                ser = self.bike[0]  # Assuming only one bike
                if ser.is_open:  
                    line = ser.readline().strip().decode()

                    pedal_rpm = round(float(line.replace("Revolutions per minute (rpm): ", "")), 1)

                    # Get smoothed rpm and power value
                    smoothed_rpm_value = round(self.moving_avg_filter.get_smoothed_value(), 1)
                    smoothed_power_value = round(self.power_conversion(smoothed_rpm_value), 1)

                    # Record data at exact 0.5-second intervals
                    if elapsed_time >= next_record_time:
                        time_sec = round(next_record_time, 1)  # Use the exact interval time
                        row_data = [time.strftime("%Y-%m-%d %H:%M:%S"), str(time_sec), str(pedal_rpm), 
                                    str(self.power_conversion(str(pedal_rpm))), str(smoothed_rpm_value), str(smoothed_power_value)]
                        next_record_time += 0.5  # Increment by exactly 0.5 seconds
                        
                    self.update_gauges(str(pedal_rpm))
                    self.update_terminal(f"Rodgue Echo Pedal Rpm: {pedal_rpm} rpm\n")
                else:
                    self.update_terminal(f"Rogue Echo Serial port is not open\n")
                    time.sleep(1)

            except (serial.SerialException, ValueError, UnicodeDecodeError, IndexError):
                self.update_terminal(f"Rogue Echo - Error reading data - Please Check Device\n")
                time.sleep(1)

            # Write row data to CSV
            if row_data:  # Check if row_data is not empty
                with open(file_name, "a", newline="") as csvfile:
                    csv_writer = csv.writer(csvfile)
                    csv_writer.writerow(row_data)
            
            if time_sec >= total_runtime:
                self.stop_logging()
                break

            # Update countdown every 0.1 seconds for smoother display
            if time.time() - countdown_time_update >= 0.1:
                if initial_runtime == 0:
                    countdown = round(elapsed_time)
                else:
                    countdown = max(0, initial_runtime - round(elapsed_time))
                self.countdown_label.configure(text=f"Timer\n{countdown}")
                countdown_time_update = time.time()

        
        # Close all serial connections
        for ser in self.bike:
            ser.close()
        # Clear the list of anemometers
        self.bike.clear()

        self.update_terminal(f"File location: {file_name}\n")
        self.update_terminal("Logging stopped\n")


    def get_total_runtime(self):
        runtime = self.runtime_entry.get()
        if not runtime:
            return float('inf')
        else:
            return int(runtime)



    def update_gauges(self, rpm):
        # Apply the moving average filter to the RPM data using the instance created in __init__
        self.moving_avg_filter.add_value(float(rpm))
        smoothed_value = self.moving_avg_filter.get_smoothed_value()

        self.rpm_gauge.set(round(float(smoothed_value), 1))
        self.rpm_gauge.update()
        self.power_gauge.set(self.power_conversion(smoothed_value))
        self.power_gauge.update()


    def power_conversion(self, rpm):
        # power = 0.0012 * float(rpm) ** 3 - 0.0182 * float(rpm) ** 2 + 1.462 * float(rpm) - 4.9911
        power = 0.0012 * float(rpm) ** 3 - 0.0168 * float(rpm) ** 2 + 1.3771 * float(rpm) - 4.2804

        power = round(power, 1)
        return power
   

# Define the power equation
    def power_equation(self, rpm, P):
        # return 0.0012 * rpm**3 - 0.0182 * rpm**2 + 1.462 * rpm - 4.9911 - P
        return 0.0012 * rpm**3 - 0.0168 * rpm**2 + 1.3771 * rpm - 4.2804 - P
    

    # Bisection method to find the root
    def bisection_method(self, P, a, b, tolerance=1e-6, max_iterations=10000):
        if self.power_equation(a, P) * self.power_equation(b, P) >= 0:
            raise ValueError("The given interval does not contain the root")

        for _ in range(max_iterations):
            c = (a + b) / 2
            if self.power_equation(c, P) == 0 or (b - a) / 2 < tolerance:
                return c
            if self.power_equation(c, P) * self.power_equation(a, P) < 0:
                b = c
            else:
                a = c

        raise ValueError("Bisection method did not converge")
        

    def reset_display(self):
        self.live_update_flag = False  # Reset live update flag
        self.auto_start_delay_flag = False
        # Stop logging
        self.stop_logging()

        # Close all serial connections
        for ser in self.bike:
            ser.close()
        # # Clear the list of anemometers
        # self.bike.clear()

        # # Clear COM port selections
        # for var in self.com_port_dropdowns:
        #     var.set('')
        
        # Clear terminal
        self.clear_terminal()

        self.live_update_flag = True  # Reset live update flag

        self.rpm_gauge.set_mark(2*0, 2*120, color="#FFFFFF")
        self.power_gauge.set_mark(0, 5000, color="#FFFFFF")
        self.min_rpm_entry.delete(0,100)
        self.max_rpm_entry.delete(0,100)
        self.min_power_entry.delete(0,100)
        self.max_power_entry.delete(0,100)
        self.runtime_entry.delete(0,100)
        self.auto_start_delay.delete(0,100)

        self.min_rpm_entry.configure(placeholder_text="Min RPM")
        self.max_rpm_entry.configure(placeholder_text="Max RPM")
        self.min_power_entry.configure(placeholder_text="Min Power")
        self.max_power_entry.configure(placeholder_text="Max Power")
        self.runtime_entry.configure(placeholder_text="RunTime (seconds)")
        self.auto_start_delay.configure(placeholder_text="Start Delay (seconds)")
        self.countdown_label.configure(text=f"Timer\n0")


        self.rpm_gauge.set(0)
        self.rpm_gauge.update()
        self.power_gauge.set(0)
        self.power_gauge.update()

        self.min_rpm_entry.configure(state= "normal")
        self.max_rpm_entry.configure(state= "normal")
        self.min_power_entry.configure(state= "normal")
        self.max_power_entry.configure(state= "normal")
        self.auto_start_delay.configure(state= "normal")
        self.runtime_entry.configure(state= "normal")
        self.auto_start_switch.configure(state= "normal")
        self.participants_ID.configure(state="normal")


    def clear_terminal(self):
        self.terminal.delete(1.0, ctk.END)
        self.terminal.update()


    def update_terminal(self, message):
        self.terminal.insert(ctk.END, message)
        self.terminal.see(ctk.END)  # Scroll to the end of the text

    # def on_close(self):
    #     # Reset program and close window
    #     self.reset_display
    #     # Register the cleanup function to be called when the program exits
    #     self.master.destroy()

    def on_close(self):
        msg = CTkMessagebox(title="Exit?", message="Do you want to close the program?",
                        icon="question", option_1="No", option_2="Yes")
        response = msg.get()
        
        if response=="Yes":
            self.live_update_flag = False
            
            # Close all serial connections
            for ser in self.bike:
                ser.close()
            self.reset_display()
            self.master.destroy()    


def create_about_dialog(root):
    cur_dir = os.getcwd()
    cur_dir = cur_dir.replace("\\", "/")
    # icon_path = cur_dir+"/favicon.ico"
    icon_path = "images/spinsync_icon.ico"

    # Set the icon for the about dialog
    about_dialog = ctk.CTk()
    
    about_dialog.geometry("560x775")  # Adjust dimensions as needed
    about_dialog.title("About")
    about_dialog.attributes("-topmost", True)  # Set the window to be topmost
    about_dialog.iconbitmap(icon_path)  # Set the icon for the about dialog

    # Frame for content
    content_frame = ctk.CTkFrame(master=about_dialog, width=2000, height=200)
    content_frame.pack(padx=25, pady=25)

    # Application name label (customize text and font)
    app_name_label = ctk.CTkLabel(master=content_frame,
                                text="SpinSync Interface",
                                font=("Arial", 18, "bold"))
    app_name_label.pack(pady=10)

    # Version label (customize text and font)
    version_label = ctk.CTkLabel(master=content_frame,
                                text="Version: 1.0",  # Update version number
                                font=("Arial", 12, "bold"))
    version_label.pack()

    # Author label (customize text and font)
    author_label = ctk.CTkLabel(master=content_frame,
                                text="Developed by: Brock Cooper",
                                font=("Arial", 12, "bold"))
    author_label.pack()

    # Usage
    description_label = ctk.CTkLabel(master=content_frame,
                                text="\n\
This program was designed specifically to be used with the Rogue Echo exercise Bike.\n\n\
Select the COM port of the Bike and click Connect. The Bike should display the set\n\
parameters along with a connected message in the output window. If no limits are\n\
specified a warning message will be displayed. The live rpm and power output will\n\
be displayed on the gauges. If the bike doesn't initially connect, please check the\n\
physical connection to the bike or ensure you have the correct com port selected.\n\n\
The program has multiple fields:\n\
Inputs Being:\n\
- Participants ID\n\
- min RPM and Max RPM values.\n\
- min Power and Max Power values.\n\
- Runtime (seconds)\n\
- Start Delay\n\
- Auto Start Switch\n\n\
NOTE: Only one set of values can be used at a time. Either RPM or Power values.\n\n\
The program will convert the RPM values to power output using a polynomial equation.\n\
The program will also convert the power output to RPM values using a bisection method.\n\n\
The program has a live update feature that will update the gauges with the current\n\
rpm and power output. If the auto start switch is enabled, the program will automatically\n\
start logging when the bike reaches the specified target range for rpm or power output.\n\
Otherwise logging needs to be manually started with the Start Logging button. A specific\n\
runtime can be set which will cause the logging to stop after this timer has finished.\n\
Alternatively if no RunTime is specified the program will continuously log data until\n\
the Stop Logging button is pressed. The program will log the data to a CSV file and\n\
display the location where the file is stored after logging is stopped. After logging\n\
is complete the bike will automatically disconnect. If it's the first time logging\n\
has occurred, the program will generate a folder called 'SpinSync Logs' in the same\n\
location where the program is stored. If you need to start logging again, please\n\
reconnect to the bike with the required parameters.",
                                font=("Arial", 12), padx=20)
    description_label.pack()


    # IMAGE_WIDTH = 200
    # IMAGE_HEIGHT = 200

    # bike_image = ctk.CTkImage(light_image=Image.open("images/bike_microcontroller.png"), dark_image=Image.open("images/bike_microcontroller.png"), size=(IMAGE_WIDTH , IMAGE_HEIGHT))
    

    # # Create a label to display the image
    # bike_image_label = ctk.CTkLabel(content_frame, image=bike_image, text='', corner_radius=40)
    # # pywinstyles.set_opacity(image_label, value=1, color="#000001") # just add this line

    # bike_image_label.pack()

    
    # bike_image_label.grid(row=0, column=0, columnspan=3, pady=10, padx=10)  # Span across all columns



    # bike_image = Image.open("images/rbike_microcontroller.png")
    # bike_image = bike_image.resize((300, 200), Image.ANTIALIAS)

    # Copyright label (customize text and font)
    copyright_label = ctk.CTkLabel(master=content_frame,
                                text="Copyright © 2024 Brock Cooper",
                                font=("Arial", 10))
    copyright_label.pack()

    # Close button
    close_button = ctk.CTkButton(master=content_frame,
                                text="Close",
                                command=about_dialog.destroy)
    close_button.pack(pady=20)

    # Apply desired theme (optional)
    # about_dialog.set_theme("dark")  # Example theme, customize as needed

    about_dialog.mainloop()

def main():
    root = ctk.CTk()
    root.geometry("1242x786")  # Set the initial size of the window to 1000x800 pixels
    root.resizable(False, False)


    app = MyInterface(root)
    # Handle window closing event
    root.protocol("WM_DELETE_WINDOW", app.on_close)

    # Set the window icon
    cur_dir = os.getcwd()
    cur_dir = cur_dir.replace("\\", "/")
    # icon_path = cur_dir+"/favicon.ico"
    icon_path = "images/spinsync_icon.ico"
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(icon_path)
    print(icon_path)
    root.iconbitmap(icon_path)  # Set the window icon
    

    # Create a Tkinter menubar
    menubar = tk.Menu(root)

    # Create "File" menu
    file_menu = tk.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="File", menu=file_menu)

    # # Create "Theme" submenu under "File" menu
    # theme_menu = tk.Menu(file_menu, tearoff=0)
    # file_menu.add_cascade(label="Theme", menu=theme_menu)
    # theme_menu.add_command(label="Light Theme", command=lambda: app.change_theme("light"))
    # theme_menu.add_command(label="Dark Theme", command=lambda: app.change_theme("dark"))
    # theme_menu.add_command(label="System Theme", command=lambda: app.change_theme("system"))

    # Add "Help" command directly to "File" menu
    file_menu.add_command(label="Help", command=lambda: create_about_dialog(root))

    root.configure(menu=menubar)
    root.mainloop()

if __name__ == "__main__":
    main()
