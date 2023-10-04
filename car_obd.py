import obd
import time
import pandas as pd
import os
import datetime
from openpyxl import load_workbook

print("Reading OBD Port...")
connection = obd.OBD()  # auto connect


# print("Connection Success!")


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


def change_file(file_name_starter, file_name_counter):
    while check_excel_file_exists(file_name):
        file_name_counter += 1
        file_name = f'D:\\OBD\\{file_name_starter}\\{file_name_starter}{file_name_counter}.xlsx'
    return file_name


def live_data_stream():
    print("Start Reading Data And Saving To Excel File")
    # Create an empty list to store DataFrames
    data_frames = []

    file_name = "D:\\OBD\\live_data_stream.xlsx"
    file_name_counter = 0
    # change the writing file if it exists
    change_file("live_data_stream", file_name_counter)

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
            file_name = change_file("live_data_stream", file_name_counter)

        time.sleep(1)  # Delay for 1 second before the next iteration
        key = input("q To Exit")
        if key == "q":
            # if 'q' key-pressed break out
            break
    print()
    main()


def display_error_codes():
    if connection.is_connected():
        print("Start Reading Errors And Saving To Excel File")
        # Create an empty list to store DataFrames
        data_frames = []

        file_name = "D:\\OBD\\display_error_codes\\display_error_codes.xlsx"
        file_name_counter = 0
        # change the writing file if it exists
        change_file("display_error_codes", file_name_counter)

        obd_error = {
            "GET_DTC": connection.query(obd.commands.GET_DTC),
            "GET_CURRENT_DTC": connection.query(obd.commands.GET_CURRENT_DTC)
        }
        # Create a DataFrame from the obd_data dictionary
        df = pd.DataFrame([obd_error])

        # Append the DataFrame to the list
        data_frames.append(df)

        # Concatenate all DataFrames in the list
        data = pd.concat(data_frames, ignore_index=True)

        # Write the DataFrame to an Excel file
        data.to_excel(file_name, index=False)


def clear_error_codes():
    print("clear_error_codes")
    if connection.is_connected():
        accept = input("Clearing Codes, Are You Sure? y/n: ")
        if accept == "y":
            connection.query(obd.commands.CLEAR_DTC)
        else:
            print()
            main()


def display_vehicle_information():
    if connection.is_connected():
        print("Start Reading Vehicle Information And Saving To Excel File")
        # Create an empty list to store DataFrames
        data_frames = []

        file_name = "D:\\OBD\\display_vehicle_information\\display_vehicle_information.xlsx"
        file_name_counter = 0
        # change the writing file if it exists
        change_file("display_vehicle_information", file_name_counter)

        obd_error = {
            "VIN_MESSAGE_COUNT": connection.query(obd.commands.VIN_MESSAGE_COUNT),
            "VIN": connection.query(obd.commands.VIN),
            "CALIBRATION_ID_MESSAGE_COUNT": connection.query(obd.commands.CALIBRATION_ID_MESSAGE_COUNT),
            "CALIBRATION_ID": connection.query(obd.commands.CALIBRATION_ID),
            "CVN_MESSAGE_COUNT": connection.query(obd.commands.CVN_MESSAGE_COUNT),
            "CVN": connection.query(obd.commands.CVN),
            "PERF_TRACKING_MESSAGE_COUNT": connection.query(obd.commands.PERF_TRACKING_MESSAGE_COUNT),
            "PERF_TRACKING_SPARK": connection.query(obd.commands.PERF_TRACKING_SPARK),
            "ECU_NAME_MESSAGE_COUNT": connection.query(obd.commands.ECU_NAME_MESSAGE_COUNT),
            "ECU_NAME": connection.query(obd.commands.ECU_NAME),
            "PERF_TRACKING_COMPRESSION": connection.query(obd.commands.PERF_TRACKING_COMPRESSION),
        }
        # Create a DataFrame from the obd_data dictionary
        df = pd.DataFrame([obd_error])

        # Append the DataFrame to the list
        data_frames.append(df)

        # Concatenate all DataFrames in the list
        data = pd.concat(data_frames, ignore_index=True)

        # Write the DataFrame to an Excel file
        data.to_excel(file_name, index=False)


def main():
    print("1- Live data stream")
    print("2- Display Error Codes")
    print("3- Clear Error Codes")
    print("4- Display Vehicle Information")
    option = int(input("Choose an option: "))
    if option == 1:
        live_data_stream()
    elif option == 2:
        display_error_codes()
    elif option == 3:
        clear_error_codes()
    elif option == 4:
        display_vehicle_information()
    else:
        print("Not Valid Option")
        main()


main()
