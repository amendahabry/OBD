import obd
import time
import pandas as pd
import os
import datetime
from openpyxl import load_workbook

connection = obd.OBD()  # auto connect


# OR

# connection = obd.OBD("/dev/ttyUSB0") # create connection with USB 0

# OR

# ports = obd.scan_serial()      # return list of valid USB or RF ports

# connection = obd.OBD(ports[0]) # connect to the first port in the list

def check_excel_file_exists(file_path):
    if os.path.isfile(file_path):
        return True
    else:
        return False


# Create an empty list to store DataFrames
data_frames = []

'''
counter = 0
while counter != 6:
    # Create a dictionary to store the data for each OBD command
    obd_data = {
        "RPM": 777,
        "STATUS": 77,
        "FUEL_STATUS": 7,
        # Add more OBD commands here...
    }

    # Create a DataFrame from the obd_data dictionary
    df = pd.DataFrame([obd_data])

    # Append the DataFrame to the list
    data_frames.append(df)

    # Concatenate all DataFrames in the list
    data = pd.concat(data_frames, ignore_index=True)

    file_name = "D:\\OBD\\obd_data.xlsx"
    
    # Write the DataFrame to an Excel file
    data.to_excel(file_name, index=False)

    # Print the data (optional)
    print(data)
    counter += 1
    time.sleep(1)  # Delay for 1 second before the next iteration

'''
file_name = "D:\\OBD\\obd_data.xlsx"
file_name_counter = 0

def change_file(file_name_counter):
    while check_excel_file_exists(file_name):
        file_name_counter += 1
        file_name = f'D:\\OBD\\obd_data{file_name_counter}.xlsx'

# change the writing file if it exists
change_file(file_name_counter)

while connection.is_connected():
    current_time = datetime.datetime.now()
    # Create a dictionary to store the data for each OBD command
    obd_data = {
        "RPM": connection.query(obd.commands.RPM),
        "STATUS": connection.query(obd.commands.STATUS),
        "FUEL_STATUS": connection.query(obd.commands.FUEL_STATUS),
        "ENGINE_LOAD": connection.query(obd.commands.ENGINE_LOAD),
        "COOLANT_TEMP": connection.query(obd.commands.COOLANT_TEMP),
        "SHORT_FUEL_TRIM_1": connection.query(obd.commands.SHORT_FUEL_TRIM_1),
        "LONG_FUEL_TRIM_1": connection.query(obd.commands.LONG_FUEL_TRIM_1),
        "SHORT_FUEL_TRIM_2": connection.query(obd.commands.SHORT_FUEL_TRIM_2),
        "LONG_FUEL_TRIM_2": connection.query(obd.commands.LONG_FUEL_TRIM_2),
        "FUEL_PRESSURE": connection.query(obd.commands.FUEL_PRESSURE),
        "INTAKE_PRESSURE": connection.query(obd.commands.INTAKE_PRESSURE),
        "SPEED": connection.query(obd.commands.SPEED),
        "TIMING_ADVANCE": connection.query(obd.commands.TIMING_ADVANCE),
        "INTAKE_TEMP": connection.query(obd.commands.INTAKE_TEMP),
        "MAF": connection.query(obd.commands.MAF),
        "THROTTLE_POS": connection.query(obd.commands.THROTTLE_POS),
        "AIR_STATUS": connection.query(obd.commands.AIR_STATUS),
        "O2_SENSORS": connection.query(obd.commands.O2_SENSORS),
        "OBD_COMPLIANCE": connection.query(obd.commands.OBD_COMPLIANCE),
        "O2_SENSORS_ALT": connection.query(obd.commands.O2_SENSORS_ALT),
        "AUX_INPUT_STATUS": connection.query(obd.commands.AUX_INPUT_STATUS),
        "RUN_TIME": connection.query(obd.commands.RUN_TIME),
        "PIDS_B": connection.query(obd.commands.PIDS_B),
        "DISTANCE_W_MIL": connection.query(obd.commands.DISTANCE_W_MIL),
        "COMMANDED_EGR": connection.query(obd.commands.COMMANDED_EGR),
        "EGR_ERROR": connection.query(obd.commands.EGR_ERROR),
        "EVAPORATIVE_PURGE": connection.query(obd.commands.EVAPORATIVE_PURGE),
        "FUEL_LEVEL": connection.query(obd.commands.FUEL_LEVEL),
        "WARMUPS_SINCE_DTC_CLEAR": connection.query(obd.commands.WARMUPS_SINCE_DTC_CLEAR),
        "DISTANCE_SINCE_DTC_CLEAR": connection.query(obd.commands.DISTANCE_SINCE_DTC_CLEAR),
        "EVAP_VAPOR_PRESSURE": connection.query(obd.commands.EVAP_VAPOR_PRESSURE),
        "BAROMETRIC_PRESSURE": connection.query(obd.commands.BAROMETRIC_PRESSURE),
        "CATALYST_TEMP_B1S1": connection.query(obd.commands.CATALYST_TEMP_B1S1),
        "CATALYST_TEMP_B2S1": connection.query(obd.commands.CATALYST_TEMP_B2S1),
        "CATALYST_TEMP_B1S2": connection.query(obd.commands.CATALYST_TEMP_B1S2),
        "CATALYST_TEMP_B2S2": connection.query(obd.commands.CATALYST_TEMP_B2S2),
        "PIDS_C": connection.query(obd.commands.PIDS_C),
        "STATUS_DRIVE_CYCLE": connection.query(obd.commands.STATUS_DRIVE_CYCLE),
        "CONTROL_MODULE_VOLTAGE": connection.query(obd.commands.CONTROL_MODULE_VOLTAGE),
        "ABSOLUTE_LOAD": connection.query(obd.commands.ABSOLUTE_LOAD),
        "COMMANDED_EQUIV_RATIO": connection.query(obd.commands.COMMANDED_EQUIV_RATIO),
        "RELATIVE_THROTTLE_POS": connection.query(obd.commands.RELATIVE_THROTTLE_POS),
        "AMBIANT_AIR_TEMP": connection.query(obd.commands.AMBIANT_AIR_TEMP),
        "THROTTLE_POS_B": connection.query(obd.commands.THROTTLE_POS_B),
        "THROTTLE_POS_C": connection.query(obd.commands.THROTTLE_POS_C),
        "ACCELERATOR_POS_D": connection.query(obd.commands.ACCELERATOR_POS_D),
        "ACCELERATOR_POS_E": connection.query(obd.commands.ACCELERATOR_POS_E),
        "ACCELERATOR_POS_F": connection.query(obd.commands.ACCELERATOR_POS_F),
        "THROTTLE_ACTUATOR": connection.query(obd.commands.THROTTLE_ACTUATOR),
        "RUN_TIME_MIL": connection.query(obd.commands.RUN_TIME_MIL),
        "MAX_MAF": connection.query(obd.commands.MAX_MAF),
        "FUEL_TYPE": connection.query(obd.commands.FUEL_TYPE),
        "ETHANOL_PERCENT": connection.query(obd.commands.ETHANOL_PERCENT),
        "EVAP_VAPOR_PRESSURE_ABS": connection.query(obd.commands.EVAP_VAPOR_PRESSURE_ABS),
        "EVAP_VAPOR_PRESSURE_ALT": connection.query(obd.commands.EVAP_VAPOR_PRESSURE_ALT),
        "SHORT_O2_TRIM_B1": connection.query(obd.commands.SHORT_O2_TRIM_B1),
        "LONG_O2_TRIM_B1": connection.query(obd.commands.LONG_O2_TRIM_B1),
        "SHORT_O2_TRIM_B2": connection.query(obd.commands.SHORT_O2_TRIM_B2),
        "LONG_O2_TRIM_B2": connection.query(obd.commands.LONG_O2_TRIM_B2),
        "FUEL_RAIL_PRESSURE_ABS": connection.query(obd.commands.FUEL_RAIL_PRESSURE_ABS),
        "RELATIVE_ACCEL_POS": connection.query(obd.commands.RELATIVE_ACCEL_POS),
        "HYBRID_BATTERY_REMAINING": connection.query(obd.commands.HYBRID_BATTERY_REMAINING),
        "OIL_TEMP": connection.query(obd.commands.OIL_TEMP),
        "FUEL_INJECT_TIMING": connection.query(obd.commands.FUEL_INJECT_TIMING),
        "FUEL_RATE": connection.query(obd.commands.FUEL_RATE),
        "DATETIME": f'{current_time.hour}:{current_time.minute}:{current_time.second}',
        # Add more OBD commands here...
    }

    # Create a DataFrame from the obd_data dictionary
    df = pd.DataFrame([obd_data])

    # Append the DataFrame to the list
    data_frames.append(df)

    # Concatenate all DataFrames in the list
    data = pd.concat(data_frames, ignore_index=True)

    # Write the DataFrame to an Excel file
    data.to_excel(file_name, index=False)

    # change the file when the data length gets to 400 rows
    if len(data) == 400:
        change_file(file_name_counter)

    time.sleep(1)  # Delay for 1 second before the next iteration

'''
while True:
    print("RPM: ", connection.query(obd.commands.RPM))
    print("STATUS: ", connection.query(obd.commands.STATUS))
    print("FUEL_STATUS: ", connection.query(obd.commands.FUEL_STATUS))
    print("ENGINE_LOAD: ", connection.query(obd.commands.ENGINE_LOAD))
    print("COOLANT_TEMP: ", connection.query(obd.commands.COOLANT_TEMP))
    print("SHORT_FUEL_TRIM_1: ", connection.query(obd.commands.SHORT_FUEL_TRIM_1))
    print("LONG_FUEL_TRIM_1: ", connection.query(obd.commands.LONG_FUEL_TRIM_1))
    print("SHORT_FUEL_TRIM_2: ", connection.query(obd.commands.SHORT_FUEL_TRIM_2))
    print("LONG_FUEL_TRIM_2: ", connection.query(obd.commands.LONG_FUEL_TRIM_2))
    print("FUEL_PRESSURE: ", connection.query(obd.commands.FUEL_PRESSURE))
    print("INTAKE_PRESSURE: ", connection.query(obd.commands.INTAKE_PRESSURE))
    print("RPM: ", connection.query(obd.commands.RPM))
    print("SPEED: ", connection.query(obd.commands.SPEED))
    print("TIMING_ADVANCE: ", connection.query(obd.commands.TIMING_ADVANCE))
    print("INTAKE_TEMP: ", connection.query(obd.commands.INTAKE_TEMP))
    print("MAF: ", connection.query(obd.commands.MAF))
    print("THROTTLE_POS: ", connection.query(obd.commands.THROTTLE_POS))
    print("AIR_STATUS: ", connection.query(obd.commands.AIR_STATUS))
    print("O2_SENSORS: ", connection.query(obd.commands.O2_SENSORS))
    print("OBD_COMPLIANCE: ", connection.query(obd.commands.OBD_COMPLIANCE))
    print("O2_SENSORS_ALT: ", connection.query(obd.commands.O2_SENSORS_ALT))
    print("AUX_INPUT_STATUS: ", connection.query(obd.commands.AUX_INPUT_STATUS))
    print("RUN_TIME: ", connection.query(obd.commands.RUN_TIME))
    print("PIDS_B: ", connection.query(obd.commands.PIDS_B))
    print("DISTANCE_W_MIL: ", connection.query(obd.commands.DISTANCE_W_MIL))
    print("COMMANDED_EGR: ", connection.query(obd.commands.COMMANDED_EGR))
    print("EGR_ERROR: ", connection.query(obd.commands.EGR_ERROR))
    print("EVAPORATIVE_PURGE: ", connection.query(obd.commands.EVAPORATIVE_PURGE))
    print("FUEL_LEVEL: ", connection.query(obd.commands.FUEL_LEVEL))
    print("WARMUPS_SINCE_DTC_CLEAR: ", connection.query(obd.commands.WARMUPS_SINCE_DTC_CLEAR))
    print("DISTANCE_SINCE_DTC_CLEAR: ", connection.query(obd.commands.DISTANCE_SINCE_DTC_CLEAR))
    print("EVAP_VAPOR_PRESSURE: ", connection.query(obd.commands.EVAP_VAPOR_PRESSURE))
    print("BAROMETRIC_PRESSURE: ", connection.query(obd.commands.BAROMETRIC_PRESSURE))
    print("CATALYST_TEMP_B1S1: ", connection.query(obd.commands.CATALYST_TEMP_B1S1))
    print("CATALYST_TEMP_B2S1: ", connection.query(obd.commands.CATALYST_TEMP_B2S1))
    print("CATALYST_TEMP_B1S2: ", connection.query(obd.commands.CATALYST_TEMP_B1S2))
    print("CATALYST_TEMP_B2S2: ", connection.query(obd.commands.CATALYST_TEMP_B2S2))
    print("PIDS_C: ", connection.query(obd.commands.PIDS_C))
    print("STATUS_DRIVE_CYCLE: ", connection.query(obd.commands.STATUS_DRIVE_CYCLE))
    print("CONTROL_MODULE_VOLTAGE: ", connection.query(obd.commands.CONTROL_MODULE_VOLTAGE))
    print("ABSOLUTE_LOAD: ", connection.query(obd.commands.ABSOLUTE_LOAD))
    print("COMMANDED_EQUIV_RATIO: ", connection.query(obd.commands.COMMANDED_EQUIV_RATIO))
    print("RELATIVE_THROTTLE_POS: ", connection.query(obd.commands.RELATIVE_THROTTLE_POS))
    print("AMBIANT_AIR_TEMP: ", connection.query(obd.commands.AMBIANT_AIR_TEMP))
    print("THROTTLE_POS_B: ", connection.query(obd.commands.THROTTLE_POS_B))
    print("THROTTLE_POS_C: ", connection.query(obd.commands.THROTTLE_POS_C))
    print("ACCELERATOR_POS_D: ", connection.query(obd.commands.ACCELERATOR_POS_D))
    print("ACCELERATOR_POS_E: ", connection.query(obd.commands.ACCELERATOR_POS_E))
    print("ACCELERATOR_POS_F: ", connection.query(obd.commands.ACCELERATOR_POS_F))
    print("THROTTLE_ACTUATOR: ", connection.query(obd.commands.THROTTLE_ACTUATOR))
    print("RUN_TIME_MIL: ", connection.query(obd.commands.RUN_TIME_MIL))
    print("MAX_MAF: ", connection.query(obd.commands.MAX_MAF))
    print("FUEL_TYPE: ", connection.query(obd.commands.FUEL_TYPE))
    print("ETHANOL_PERCENT: ", connection.query(obd.commands.ETHANOL_PERCENT))
    print("EVAP_VAPOR_PRESSURE_ABS: ", connection.query(obd.commands.EVAP_VAPOR_PRESSURE_ABS))
    print("EVAP_VAPOR_PRESSURE_ALT: ", connection.query(obd.commands.EVAP_VAPOR_PRESSURE_ALT))
    print("SHORT_O2_TRIM_B1: ", connection.query(obd.commands.SHORT_O2_TRIM_B1))
    print("LONG_O2_TRIM_B1: ", connection.query(obd.commands.LONG_O2_TRIM_B1))
    print("SHORT_O2_TRIM_B2: ", connection.query(obd.commands.SHORT_O2_TRIM_B2))
    print("LONG_O2_TRIM_B2: ", connection.query(obd.commands.LONG_O2_TRIM_B2))
    print("FUEL_RAIL_PRESSURE_ABS: ", connection.query(obd.commands.FUEL_RAIL_PRESSURE_ABS))
    print("RELATIVE_ACCEL_POS: ", connection.query(obd.commands.RELATIVE_ACCEL_POS))
    print("HYBRID_BATTERY_REMAINING: ", connection.query(obd.commands.HYBRID_BATTERY_REMAINING))
    print("OIL_TEMP: ", connection.query(obd.commands.OIL_TEMP))
    print("FUEL_INJECT_TIMING: ", connection.query(obd.commands.FUEL_INJECT_TIMING))
    print("FUEL_RATE: ", connection.query(obd.commands.FUEL_RATE))
    time.sleep(1)
'''
