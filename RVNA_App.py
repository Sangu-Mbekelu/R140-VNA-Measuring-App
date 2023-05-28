import tkinter as tk
import pyvisa
from datetime import datetime
from tkinter import *
import paramiko
import pandas as pd
import os
import User_Pass_Key


def measurements():
    global calibration_set, take_measurements, counter, working_directory_path
    
    time_between_measurements = Time_Inbetween_Entry.get()
    if time_between_measurements == '':
        time_between_measurements = "10"  # Used as a way to avoid error for empty string
        
    if calibration_set == 1 and take_measurements == 1:

        CMT.write("TRIG:SOUR BUS")  # Set sweep source to BUS for automated measurement
        CMT.query("*OPC?")  # Wait for measurement to complete

        CMT.write("TRIG:SING")  # Trigger a single sweep
        CMT.query("*OPC?")  # Wait for measurement to complete

        current_datetime = datetime.now()
        current_time_hour = current_datetime.strftime("%H")  # Logging the current hour
        current_time_minute = current_datetime.strftime("%M")  # Logging the current minute
        current_time_second = current_datetime.strftime("%S")  # Logging the current second

        # Read frequency data
        freq = CMT.query_ascii_values("SENS1:FREQ:DATA?")

        # Read smith chart impedance data
        CMT.write("CALC1:PAR2:SEL")
        imp = CMT.query_ascii_values("CALC1:DATA:FDAT?")  # Get data as string
        real_imp = imp[::2]
        imag_imp = imp[1::2]

        # Read log mag data
        CMT.write("CALC1:PAR3:SEL")
        log_mag = CMT.query_ascii_values("CALC1:DATA:FDAT?")  # Get data as string
        log_mag = log_mag[::2]

        # Read phase data
        CMT.write("CALC1:PAR1:SEL")
        phase = CMT.query_ascii_values("CALC1:DATA:FDAT?")  # Get data as string
        phase = phase[::2]

        vna_temp = CMT.write("SYST:TEMP:SENS? '1'")

        CMT.write("TRIG:SOUR INT")  # Set sweep source to INT after measurements are done

        CMT.query("*OPC?")  # Wait for measurement to complete

        current_time_hour = [int(current_time_hour)] * len(freq)

        current_time_minute = [int(current_time_minute)] * len(freq)

        current_time_second = [int(current_time_second)] * len(freq)

        vna_temp = [(vna_temp * (9/5))+32] * len(freq)

        output.insert(tk.END, "Measurements Taken\n")

        data_dictionary = {'Current Hour': current_time_hour, 'Current Minute': current_time_minute, 'Current Second': current_time_second, 'Frequency [Hz]': freq, 'S11 [dB]': log_mag, 'S11 Phase [DEG]': phase, 'Zin [RE]': real_imp, 'Zin [IM]': imag_imp, 'VNA Temp [F]': vna_temp}

        data_frame = pd.DataFrame(data_dictionary)

        modified_datetime = current_datetime.strftime("%m-%d-%Y_%H-%M-%S")

        file_name = 'S_parameters_' + str(modified_datetime) + '.txt'

        data_frame.to_csv(file_name, index=False, sep=',')

        s_parameter_file_path = os.path.join(working_directory_path, file_name)

        sftp_session.put(s_parameter_file_path, sftp_session.getcwd() + '/' + file_name)

        os.remove(s_parameter_file_path)

    else:
        pass
    RVNA_App.after(1000*int(time_between_measurements), measurements)


def create_ssh(host, username, password):
    global ssh
    
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())  # Adds host key if missing
    
    try:
        output.insert(tk.END, "Creating a Connection...\n")
        ssh.connect(host, username=username, password=password)  # Establishes SSH connection
        output.insert(tk.END, "Connected\n")
        sftp = ssh.open_sftp()  # Opens SFTP session
        return sftp
    except:
        output.insert(tk.END, "Failed to make Connection\n")


def start_stop_measurements_button(button):
    global take_measurements, ssh, sftp_session

    if button == 1:
        take_measurements = 1
    elif button == 0:
        take_measurements = 0
        output.insert(tk.END, "Measurements Stopped\n")
        sftp_session.close()
        ssh.close()


def connect_and_calibrate():
    global CMT, measurements_folder_name, sftp_session
    
    circle_color = canvas.itemcget(status_circle, "fill")  # Grabs color of status_circle canvas
    
    if circle_color == "red":
        # RVNA Software Connection =====================================
        rm = pyvisa.ResourceManager('@py')  # use pyvisa-py as backend
        try:
            CMT = rm.open_resource('TCPIPO::127.0.0.1::5025::SOCKET')  # Connects to RVNA application using SCPI

            connection_message = "Connected to VNA\n"
            output.insert(tk.END, connection_message + "\n")
        except:
            error_message = "Failed to Connect to VNA\nCheck network settings"
            output.insert(tk.END, error_message + "\n")
            return

        CMT.read_termination = '\n'  # The VNA ends each line with this. Reads will time out without this
        CMT.timeout = 10000  # Set longer timeout period for slower sweeps

        # Server Connection Check =====================================
        try:
            measurements_folder_name = Measurements_Folder.get()  # Equal to specified file name in application
            
            sftp_session = create_ssh(Server_Host, Server_User, Server_Password)  # Calls function that remote connects to specified server and opens an SFTP session
            sftp_session.chdir(User_Pass_Key.remote_path + measurements_folder_name)  # Changes directory to specified file on the server

            canvas.itemconfig(status_circle, fill="green")

            connection_message = "Connected to Specified Folder\n"
            output.insert(tk.END, connection_message + "\n")
        except:
            error_message = "Failed to Connect to Specified Folder\nCheck folder name"
            output.insert(tk.END, error_message + "\n")
            return

        # RVNA Calibration Check ======================================
        CMT.write("MMEM:LOAD:STAT 'C:\\VNA\\RVNA\\State\\CalFile.cfg'")  # Recalls calibration state with specified file
        CMT.write("DISP:WIND:SPL 2")  # Allocate 2 trace windows
        CMT.write("CALC1:PAR:COUN 3")  # 3 Traces
        CMT.write("CALC1:PAR1:DEF S11")  # Choose S11 for trace 1
        CMT.write("CALC1:PAR2:DEF S11")  # Choose S11 for trace 2
        CMT.write("CALC1:PAR3:DEF S11")  # Choose S11 for trace 3

        CMT.write("CALC1:PAR1:SEL")  # Selects Trace 1 and Smith Chart Format
        CMT.write("CALC1:FORM PHAS")

        CMT.write("CALC1:PAR2:SEL")  # Selects Trace 2 and Log Mag Format
        CMT.write("CALC1:FORM SMIT")

        CMT.write("CALC1:PAR3:SEL")  # Selects Trace 3 and Log Mag Format
        CMT.write("CALC1:FORM MLOG")

        CMT.query("*OPC?")  # Wait for measurement to complete

        create_cal_check_window()  # Function that creates calibration check window to varify that calibration state has been loaded.
    else:
        pass


def create_cal_check_window():
    cal_check = Toplevel()
    cal_check_width = 300
    cal_check_height = 20
    cal_check.geometry(f"{cal_check_width}x{cal_check_height}")
    cal_check.title("Calibration Check")
    cal_check_button = tk.Button(cal_check, text="Has Calibration State been Recalled?",
                                 command=lambda: close_cal_check(cal_check))
    cal_check_button.pack()


def close_cal_check(cal_check):
    global calibration_set, CMT

    calibration_message = "VNA is Calibrated\n"
    calibration_set = 1
    cal_check.destroy()  # Destroys calibration check window, varifying that the calibration state is recalled
    output.insert(tk.END, calibration_message + "\n")


# Constants ========================
ssh = paramiko.SSHClient()  # Defines SSH client
Server_Host = User_Pass_Key.hostname
Server_User = User_Pass_Key.user
Server_Password = User_Pass_Key.password
Server_Root_Directory = User_Pass_Key.remote_path
working_directory_path = os.path.dirname(os.path.abspath(__file__))
# Variables ========================
measurements_folder_name = ''  # Initializing measurement folder location name

calibration_set = 0  # Initializing state of VNA calibration

take_measurements = 0  # Initializing state of measuring status

counter = 0

# TKINTER ==========================
RVNA_App = tk.Tk()  # Creating main window
RVNA_App.title("RVNA App")
window_width = 600
window_height = 500
RVNA_App.geometry(f"{window_width}x{window_height}")
RVNA_App.resizable(False, False)

# FOLDER ENTRY ======================================================================================
Measurements_Folder = tk.Entry(RVNA_App, justify="center", width=60)
Measurements_Folder.grid(row=1, column=1)

# CONNECT AND CALIBRATE BUTTON ======================================================================
Connect_Cal_button = tk.Button(RVNA_App, text="Connect & Calibrate", command=connect_and_calibrate)
Connect_Cal_button.grid(row=2, column=1, padx=50)

# CONNECT AND CALIBRATE STATUS CIRCLE ===============================================================
canvas = tk.Canvas(RVNA_App, width=50, height=47)
status_circle = canvas.create_oval(10, 10, 40, 40, fill="red")  # Creates initial status indication of red circle
canvas.grid(row=2, column=0)

# OUTPUT TEXTBOX AND SCROLLBAR ======================================================================
scrollbar = tk.Scrollbar(RVNA_App)
scrollbar.grid(row=3, column=1, padx=50)

output = tk.Text(RVNA_App, yscrollcommand=scrollbar.set, width=50, height=10)
output.grid(row=3, column=1, padx=50, sticky="nsew")
scrollbar.config(command=output.yview)

# START MEASURING BUTTON AND TIME INBETWEEN MEASUREMENTS ==============================================
Time_Inbetween_Label = tk.Label(RVNA_App, text="Time between measurments (s)")
Time_Inbetween_Label.grid(row=4, column=1, pady=5)

Time_Inbetween_Entry = tk.Entry(RVNA_App, justify="center")
Time_Inbetween_Entry.insert(0, "10")  # Set default value of 10 seconds
Time_Inbetween_Entry.grid(row=5, column=1)

Start_Measurements_button = tk.Button(RVNA_App, text="Start Measurements", command=lambda: start_stop_measurements_button(1))
Start_Measurements_button.grid(row=6, column=1, pady=5)

Stop_Measurements_button = tk.Button(RVNA_App, text="Stop Measurements", command=lambda: start_stop_measurements_button(0))
Stop_Measurements_button.grid(row=7, column=1, pady=5)

# ACTUAL MEASURING FUNCTION ==========================================================================

RVNA_App.after(10000, measurements)  # Schedule the initial execution of measurements() after 10 seconds

RVNA_App.mainloop()


