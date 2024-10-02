# -*- coding: utf-8 -*-
"""
Created on Fri Sep 27 11:19:28 2024

@authors: philip.gotthelf, alex.dering
"""

# main_window.py

import tkinter as tk  # Import the Tkinter module for GUI development.
import tkinter.ttk as ttk  # Import themed widgets from Tkinter for better styling.
from difflib import SequenceMatcher  # Used for comparing sequences and finding similarities.
from tkinter import messagebox, filedialog  # Import specific Tkinter features for message boxes and file dialogs.
import csv  # Module to handle CSV file operations.
import openpyxl as xl  # Used for working with Excel files (.xlsx format).
import os  # OS module for interacting with the operating system (file paths, etc.).
import re  # Regular expression module for pattern matching in strings.
import time  # Module for time-related functions.
import json  # JSON module to parse and manipulate JSON data.
from collections import OrderedDict  # Import ordered dictionary to maintain the order of keys.
from shutil import copy  # Used to copy files or directories.
from openpyxl import load_workbook, Workbook


"""
____________________________ AD _____________________________
"""

def save_as_csv(excel_file_path, csv_file_path):
    workbook = load_workbook(filename=excel_file_path)
    sheet = workbook.active

    with open(csv_file_path, mode='w', newline="") as file:
        writer = csv.writer(file)

        for row in sheet.iter_rows(values_only=True):
            writer.writerow(row)

def write_direction_data_to_files(sheet, matched_results, relevant_columns, headers, output_start_row=4):
    """
    Writes Volume, PHF, and HeavyVehicles data for each intersection and direction-turn 
    from the specified column ranges in relevant_columns, and saves the results to 
    separate files named based on the header in row 1 of each column range.

    Args:
    - sheet: The active sheet from which data is being read.
    - matched_results: A dictionary containing intersections and their corresponding turn data.
    - relevant_columns: A list of starting columns (e.g., [6, 9, 12] for 'F', 'I', 'L') 
      from which Volume, PHF, and HeavyVehicles are read.
    - output_start_row: The row in the output sheet to start writing the data (default is 4).
    """
    for start_column in relevant_columns:
        # Define column positions relative to the starting column
        volume_col = start_column       # Volume is in start_column (e.g., F)
        phf_col = start_column + 1      # PHF is in start_column + 1 (e.g., G)
        heavy_vehicles_col = start_column + 2  # HeavyVehicles is in start_column + 2 (e.g., H)

        # Get the header name from row 1 of the start_column (e.g., F1, I1, etc.)
        file_name_header = sheet.cell(row=1, column=start_column).value
        if not file_name_header:
            print(f"Skipping columns starting at {start_column} as no header was found in row 1.")
            continue
        
        # Create a new workbook for this specific column set
        output_workbook = Workbook()
        output_sheet = output_workbook.active
        output_sheet.title = "Results"
        output_sheet["A1"] = "[Lanes]"
        output_sheet["A2"] = "Lane Group Data"
        
        # Label cells with corresponding headers (A3-P3)
        for col, header in enumerate(headers, start=1):
            output_sheet.cell(row=3, column=col).value = header
        
        # Reset the output start row for each file
        output_start_row = 4
        
        # Iterate over each intersection and its direction-turn results
        for intersection_id, turns in matched_results.items():
            # Write Intersection ID and Labels in the output sheet
            output_sheet.cell(row=output_start_row, column=1).value = "Volume"
            output_sheet.cell(row=output_start_row + 1, column=1).value = "PHF"
            output_sheet.cell(row=output_start_row + 2, column=1).value = "HeavyVehicles"
            output_sheet.cell(row=output_start_row, column=2).value = intersection_id
            output_sheet.cell(row=output_start_row + 1, column=2).value = intersection_id
            output_sheet.cell(row=output_start_row + 2, column=2).value = intersection_id

            # Process each direction-turn within the intersection
            for direction_turn, info in turns.items():
                row_found = info['row']

                # Read data from the specified columns for the current row
                volume = sheet.cell(row=row_found, column=volume_col).value
                phf = sheet.cell(row=row_found, column=phf_col).value
                heavy_vehicles = sheet.cell(row=row_found, column=heavy_vehicles_col).value

                # Write the data into the output sheet under the correct direction-turn column
                header_column = info['header_column']
                output_sheet.cell(row=output_start_row, column=header_column).value = volume
                output_sheet.cell(row=output_start_row + 1, column=header_column).value = phf
                output_sheet.cell(row=output_start_row + 2, column=header_column).value = heavy_vehicles

                # Debugging output
                print(f"Wrote to Results for intersection {intersection_id}, direction {direction_turn}: "
                      f"Volume: {volume}, PHF: {phf}, HeavyVehicles: {heavy_vehicles}")

            # Move to the next output row for the next intersection
            output_start_row += 3  # 3 rows for data + 1 row for separation

        # Save the output workbook to a file named by the header in row 1 of the start column
        output_file_path = f"{file_name_header}.xlsx"
        output_workbook.save(output_file_path)
        save_as_csv(output_file_path, f"{file_name_header}.csv")
        os.remove(f"{file_name_header}.xlsx")
        print(f"Output file saved as {file_name_header}.csv")

    return

def read_input_file(file_path):
    # Load the input workbook and select the active sheet
    workbook = load_workbook(filename=file_path)
    sheet = workbook.active

    # Define headers for the output sheet
    headers = [
        "RECORDNAME", "INTID", "NBL", "NBT", "NBR", 
        "SBL", "SBT", "SBR", "EBL", "EBT", "EBR", 
        "WBL", "WBT", "WBR","NWR", "NWL", "NWT", "NEL", "NET", "NER",
        "SEL", "SER", "SET", "SWL", "SWR", "SWT" ,"PED", "HOLD"
    ]

    consecutive_empty_cells = 0
    intersections = {}

    # First pass: Find all intersection IDs and their corresponding row numbers
    for row in range(1, sheet.max_row + 1):
        cell_value = sheet.cell(row=row, column=1).value
        if cell_value is None:
            consecutive_empty_cells += 1
            if consecutive_empty_cells >= 25:
                break
        else:
            consecutive_empty_cells = 0
            if isinstance(cell_value, int):
                intersections[cell_value] = row

    print(f"Found intersections: {intersections}")

    directions = ["EB", "WB", "NB", "SB", "NW", "NE", "SW", "SE"]
    # output_start_row = 4  # Start writing from row 4

    # Dictionary to store results for each intersection
    intersection_results = {}

    # Second pass: Process each intersection ID and search for directions
    for intersection_id, row_with_int in intersections.items():
        found_directions = {}

        # Search column C for directions starting from the intersection row
        for search_row in range(row_with_int, sheet.max_row + 1):
            direction_value = sheet.cell(search_row, column=3).value
            if direction_value in directions and direction_value not in found_directions:
                found_directions[direction_value] = search_row
                if len(found_directions) == len(directions):
                    break

        # Dictionary to store combined direction-turn keys (e.g., EBL, WBT)
        direction_turn_results = {}

        # For each found direction, search column D for 'L', 'T', 'R'
        for direction, found_row in found_directions.items():
            turn_values = {"L": None, "T": None, "R": None}  # Default is None (not found)
            for search_row in range(found_row, sheet.max_row + 1):
                turn_value = sheet.cell(search_row, column=4).value
                if turn_value in ["L", "T", "R"]:
                    turn_values[turn_value] = search_row  # Store the row number for each turn type found
                # Break when all turn values have been found
                if all(turn_values.values()):
                    break

            # Combine direction and turn type to form keys like "EBL", "NBT", etc.
            for turn_type, row_found in turn_values.items():
                if row_found is not None:  # Only store if the turn was found
                    combined_key = f"{direction}{turn_type}"
                    direction_turn_results[combined_key] = row_found

        # Store the results for the current intersection
        intersection_results[intersection_id] = direction_turn_results

        # Display the results for debugging
        print(f"Direction-turn results for intersection {intersection_id}: {direction_turn_results}")

    # Match direction-turn results with corresponding headers
    header_mapping = {header: idx + 1 for idx, header in enumerate(headers)}

    matched_results = {}

    for intersection_id, turn_results in intersection_results.items():
        matched_results[intersection_id] = {}
        for direction_turn, row in turn_results.items():
            if direction_turn in header_mapping:
                matched_results[intersection_id][direction_turn] = {
                    "row": row,
                    "header_column": header_mapping[direction_turn]
                }
    
    relevant_columns = [6, 9, 12, 15]  # F-H, I-K, L-N
    
    write_direction_data_to_files(sheet, matched_results, relevant_columns, headers=headers, output_start_row=4)


    # Return intersection results if needed elsewhere
    return intersection_results


"""
______________________ HELPER FUNCTIONS ______________________
"""

# Check if the target is empty
def is_empty(target):
    # Check if the target is None
    if target is None:
        return True
    # If the target is a string, check if it is empty or contains only whitespace
    if type(target) == str:
        if target.strip():
            return False  # The string is not empty
        else:
            return True  # The string is empty
    return False  # The target is not empty (not None or empty string)

# Identify the type of control based on the record name
def identify_type(record_name):
    # Map record names to control types
    if record_name == 'Arrive On Green':
        control_type = 'hcm signalized'
    elif record_name == 'Opposing Approach':
        control_type = 'hcm all way stop'
    elif record_name == 'Int Delay, s/veh':
        control_type = 'hcm two way stop'
    elif record_name == 'Conflicting Circle Lanes':
        control_type = 'hcm roundabout'
    elif record_name == 'Right Turn on Red':
        control_type = 'synchro signalized'
    elif record_name == 'Degree Utilization, x':
        control_type = 'synchro all way stop'
    elif record_name == 'cSH':
        control_type = 'synchro two way stop'
    elif record_name == 'Control Type: Roundabout':
        control_type = 'synchro roundabout'
    else:
        control_type = None  # Unrecognized record name

    return control_type

# Get bounds of intersections in the file
def get_bounds(file):
    # Regex pattern to match intersection records
    pattern = re.compile('([0-9]+):\w*')
    bounds = list()  # To store the bounds of intersections
    intersections = list()  # To store intersection IDs
    data = dict()  # To store intersection data

    # Read the file content
    with open(file) as f:
        reader = csv.reader(f, delimiter='\t')
        file_data = list(reader)
    
    # Iterate through the file data to find intersection bounds
    for index, line in enumerate(file_data):
        if line:  # Skip empty lines
            record_name = line[0].strip()  # Get the first element of the line
            header_match = pattern.match(record_name)  # Match the record name with the pattern
            if header_match:  # If there's a match, it's an intersection record
                bounds.append(index)  # Store the index of the bound
                intersection = int(header_match.groups()[0])  # Get the intersection ID
                intersections.append(intersection)  # Store the intersection ID
    
    bounds.append(index)  # Append the last index for bounds

    # Process the intersections to gather data
    for index, inter in enumerate(intersections):
        if inter not in data.keys():
            data[inter] = dict()  # Initialize a dictionary for each intersection
        data[inter]['bounds'] = bounds[index:index + 2]  # Set bounds for the intersection
        start, end = data[inter]['bounds']
        
        # Loop through the lines within the bounds
        for line in file_data[start:end]:
            if line:  # Skip empty lines
                record_name = line[0].strip()  # Get the record name
                record_type = identify_type(record_name)  # Identify the control type
                if record_type:  # If a control type is found
                    data[inter]['type'] = record_type  # Set the control type
                    break  # Exit loop once the type is found
                else:
                    data[inter]['type'] = None  # No control type found
    return data  # Return the constructed data dictionary

# Find a line in data matching the search term
def find_line(data, search, give_index=False):
    # Loop through each line of data
    for index, line in enumerate(data):
        if line:  # Skip empty lines
            record_name = line[0].strip()  # Get the first element of the line
            if record_name == search:  # Check if it matches the search term
                if give_index:
                    return line, index  # Return line and index if requested
                else:
                    return line  # Return the line only
    return None  # Return None if no match is found

# Get overall values like delay and LOS based on control type
def get_overall(data_list, control_type):
    # returns overall values in form: [delay, LOS]

    # Determine the keys based on the control type
    if control_type == 'hcm signalized':
        keys = ['HCM 6th Ctrl Delay', 'HCM 6th LOS']
    elif control_type == 'hcm all way stop':
        keys = ['Intersection Delay, s/veh', 'Intersection LOS']
    elif control_type == 'hcm two way stop':
        keys = ['Int Delay, s/veh']
    elif control_type == 'hcm roundabout':
        keys = ['Intersection Delay, s/veh', 'Intersection LOS']
    elif control_type == 'synchro signalized':
        pass  # To be implemented for synchro signalized
    elif control_type == 'synchro all way stop':
        keys = ['Delay', 'Level of Service']  # Assumes HCM 2000
    elif control_type == 'synchro two way stop':
        keys = ['Average Delay']  # Assumes HCM 2000
    elif control_type == 'synchro roundabout':
        return [None, None]  # To be implemented for synchro roundabouts
    else:
        return [None, None]  # Return None for unrecognized control types

    # Handle data extraction for 'synchro signalized' control type
    if control_type == 'synchro signalized':
        for row in data_list:
            if row:
                if 'Intersection Signal Delay: ' in row[0]:  # Look for specific record
                    delay = row[0][27:].strip()  # Extract delay
                    los = row[5][-1]  # Extract level of service
                    return [delay, los]  # Return extracted values

    # If not 'synchro signalized', extract data using keys
    output = [None, None]
    for index, key in enumerate(keys):
        row = find_line(data_list, key)  # Find the row for each key
        for entry in row[2:]:  # Skip the first two columns
            if entry:  # Get the first non-empty entry
                output[index] = entry
                break

    return output  # Return the overall values

# Standardize the results from the file
def standardize(results_file):
    # Read the content of the results file
    with open(results_file) as f:
        reader = csv.reader(f, delimiter='\t')
        file_content = list(reader)  # Store the file content as a list
    database = dict()  # To store the standardized data
    parsed = get_bounds(results_file)  # Get intersection bounds and types

    # Iterate through parsed intersections to build the database
    for intersection in parsed:
        db = parsed[intersection]  # Get data for the intersection
        start = min(db['bounds'])  # Get the starting index for bounds
        end = max(db['bounds'])  # Get the ending index for bounds
        subset = file_content[start:end]  # Get the relevant data subset
        control_type = db['type']  # Get the control type
        database[intersection] = OrderedDict()  # Initialize an ordered dictionary for intersection
        database[intersection]['overall'] = dict()  # Initialize overall data dictionary
        delay, los = get_overall(subset, control_type)  # Get delay and LOS
        database[intersection]['overall']['delay'] = delay  # Store delay
        database[intersection]['overall']['los'] = los  # Store LOS

        # Initialize storage variables for detailed data
        header_by_int = OrderedDict()  # Movement headers by intersection
        secondary_key = OrderedDict()  # Secondary keys for alternate headers
        second_info = list()  # List to store additional information
        header_by_int_alt = dict()  # Alternate movement headers
        roundabout_lanes = list()  # To store roundabout lane information

        # Declare search parameters based on control type
        if control_type == 'hcm signalized':
            header_key = 'Movement'

            lookup = {'V/C Ratio(X)': 'vc_ratio',
                      'LnGrp Delay(d),s/veh': 'ln_delay',
                      'LnGrp LOS': 'ln_los',
                      'Approach Delay, s/veh': 'app_delay',
                      'Approach LOS': 'app_los'}

        elif control_type == 'hcm all way stop':

            header_key = 'Movement'
            secondary_header_key = 'Lane'

            lookup = {'HCM Control Delay': 'app_delay',
                      'HCM LOS': 'app_los'}

            lookup_2 = {'HCM Lane V/C Ratio': 'vc_ratio',
                        'HCM Control Delay': 'ln_delay',
                        'HCM Lane LOS': 'ln_los'}

        elif control_type == 'hcm two way stop':

            header_key = 'Movement'
            secondary_header_key = 'Minor Lane/Major Mvmt'
            lookup = {'HCM Control Delay, s': 'app_delay',
                      'HCM LOS': 'app_los'}

            lookup_2 = {'HCM Lane V/C Ratio': 'vc_ratio',
                        'HCM Control Delay (s)': 'ln_delay',
                        'HCM Lane LOS': 'ln_los'}

        elif control_type == 'hcm roundabout':
            header_key = 'Approach'
            lookup = {'Approach Delay, s/veh': 'app_delay',
                      'Approach LOS': 'app_los'}

            lookup_2 = {'V/C Ratio': 'vc_ratio',
                        'Control Delay, s/veh': 'ln_delay',
                        'LOS': 'ln_los'}

        if control_type == 'synchro signalized':
            header_key = 'Lane Group'
            lookup = {'v/c Ratio': 'vc_ratio',
                      'Control Delay': 'ln_delay',
                      'LOS': 'ln_los',
                      'Approach Delay': 'app_delay',
                      'Approach LOS': 'app_los'}

        elif control_type == 'synchro all way stop':
            header_key = 'Movement'
            lookup = {'Degree Utilization, x': 'vc_ratio',
                      'Control Delay (s)': 'ln_delay',
                      'LnGrp LOS': 'ln_los',
                      'Approach Delay (s)': 'app_delay',
                      'Approach LOS': 'app_los'}

        elif control_type == 'synchro two way stop':
            header_key = 'Movement'
            lookup = {'Volume to Capacity': 'vc_ratio',
                      'Control Delay (s)': 'ln_delay',
                      'Lane LOS': 'ln_los',
                      'Approach Delay (s)': 'app_delay',
                      'Approach LOS': 'app_los'}

        elif control_type == 'synchro roundabout':
            header_key = 'Movement'
            lookup = {'Volume to Capacity': 'vc_ratio',
                      'Control Delay (s)': 'ln_delay',
                      'Lane LOS': 'ln_los',
                      'Approach Delay (s)': 'app_delay',
                      'Approach LOS': 'app_los'}

        # main data collection
        if control_type == 'synchro roundabout':
            pass

        elif control_type == 'hcm roundabout':

            movement_headers = find_line(subset, header_key)
            for index, content in enumerate(movement_headers[2:]):
                index += 2
                if content:
                    header_by_int[index] = content
                    header_by_int_alt[index - 1] = content

            lanes = find_line(subset, 'Entry Lanes')
            for index, lane in enumerate(lanes[2:]):
                index += 2
                if lane:
                    for num in range(int(lane)):
                        roundabout_lanes.append(header_by_int[index])

            configurations = find_line(subset, 'Designated Moves')
            for index, content in enumerate(configurations[2:]):
                index += 2
                if content:
                    direction = roundabout_lanes[0]
                    roundabout_lanes.pop(0)
                    if len(content) == 1:
                        move = content
                    elif len(content) == 2:
                        if 'T' in content:
                            move = 'T'
                        else:
                            move = 'L'
                    else:
                        move = 'T'

                    database[intersection][direction + move] = dict()
                    config = str()
                    if 'L' in content:
                        config += '<'
                    if 'T' in content:
                        config += '1'
                    if 'R' in content:
                        config += '>'
                    database[intersection][direction + move]['config'] = config

                    for lookup_value, data_tag in lookup_2.items():
                        line = find_line(subset, lookup_value)
                        value = line[index]
                        database[intersection][direction + move][data_tag] = value
            # todo revisit for multilane roundabout support

            for lookup_value, data_tag in lookup.items():
                line = find_line(subset, lookup_value)
                for index, item in enumerate(line[2:]):
                    index += 2
                    if item:
                        direction = header_by_int[index]
                        for record, dictionary in database[intersection].items():
                            if record[:2] == direction:
                                dictionary[data_tag] = item

        elif control_type in ['hcm signalized', 'synchro signalized']:
            movement_headers = find_line(subset, header_key)
            for index, content in enumerate(movement_headers[2:]):
                index += 2
                if content:
                    database[intersection][content] = dict()
                    header_by_int[index] = content

            configurations = find_line(subset, 'Lane Configurations')
            for index, content in enumerate(configurations[2:]):
                index += 2
                if content:
                    key = header_by_int[index]
                    database[intersection][key]['config'] = content

            for line in subset:
                if line:
                    record_name = line[0].strip()
                    for lookup_value, data_tag in lookup.items():
                        if record_name == lookup_value:
                            database_field = data_tag
                            for column_num, value in enumerate(line):
                                if column_num > 1 and column_num in header_by_int.keys():
                                    movement = header_by_int[column_num]
                                    if movement in database[intersection]:
                                        database[intersection][movement][database_field] = value
                            # exit loop since each line can only be one record
                            break

        elif control_type in ['hcm all way stop', 'hcm two way stop']:
            movement_headers = find_line(subset, header_key)
            alternate_header_line, second_index = find_line(subset, secondary_header_key, give_index=True)
            for index, content in enumerate(movement_headers[2:]):
                index += 2
                if content:
                    database[intersection][content] = dict()
                    header_by_int[index] = content

            for index, header in enumerate(alternate_header_line[2:]):
                index += 2
                if header:
                    second_info.append((header[:2], index))
                    secondary_key[header] = index

            configurations = find_line(subset, 'Lane Configurations')
            for index, content in enumerate(configurations[2:]):
                index += 2
                if content:
                    key = header_by_int[index]
                    database[intersection][key]['config'] = content

            for movement in database[intersection]:
                if 'config' in database[intersection][movement].keys():
                    config = database[intersection][movement]['config']
                else:
                    continue
                if config != '0':
                    for index, pair in enumerate(second_info):
                        if movement[:2] == pair[0]:
                            header_by_int_alt[pair[1]] = movement
                            second_info.pop(index)
                            break

            for line in subset[:second_index]:
                if line:
                    record_name = line[0].strip()
                    for lookup_value, data_tag in lookup.items():
                        if record_name == lookup_value:
                            database_field = data_tag
                            for column_num, value in enumerate(line):
                                if column_num > 1 and column_num in header_by_int.keys():
                                    movement = header_by_int[column_num]
                                    if movement in database[intersection]:
                                        database[intersection][movement][database_field] = value
                            # exit loop since each line can only be one record
                            break

            for line in subset[second_index:]:
                if line:
                    record_name = line[0].strip()
                    for lookup_value, data_tag in lookup_2.items():
                        if record_name == lookup_value:
                            database_field = data_tag
                            for column_num, value in enumerate(line):
                                if column_num > 1 and column_num in header_by_int_alt.keys():
                                    movement = header_by_int_alt[column_num]
                                    if movement in database[intersection]:
                                        database[intersection][movement][database_field] = value
                            # exit loop since each line can only be one record
                            break
    return database


def order(txt):
    """Returns a string that contains 'L', 'T', and 'R' based on their presence in the input text."""
    output = str()
    if txt.find('L') != -1:
        output += 'L'
    if txt.find('T') != -1:
        output += 'T'
    if txt.find('R') != -1:
        output += 'R'
    return output  # Return the constructed output string


def label(field, config):
    """Generates a label based on the field and configuration rules, returning a direction or None."""
    output = str()
    if not field:
        return output
    if len(field) == 2:
        return field
    if field.find('Ln') != -1:
        return None
    direction = field[2]  # Extract the direction from the field
    # If no special characters are found in config, return the direction if '0' is not present
    if config.find('<') == -1 and config.find('>') == -1:
        if config.find('0') == -1:
            return direction
    
    # Check if '<' is in the config, if so, add 'L' to output
    if config.find('<') != -1:
        output += 'L'
    
    # Loop through numbers 1 to 8 and check if they are in the config
    for num in range(1, 9):
        if config.find(str(num)) != -1:
            output += direction  # Append direction for each found number
    
    # Check if '>' is in the config, if so, add 'R' to output
    if config.find('>') != -1:
        output += 'R'
    
    return order(output)  # Call order to finalize the output


def similar(str1, str2):
    """Returns the similarity ratio between two strings using SequenceMatcher."""
    return SequenceMatcher(None, str1, str2).ratio()


def load_settings():
    """Loads settings from a JSON file, creating defaults if the file does not exist."""
    
    def set_default():
        """Sets default settings and saves them to a JSON file."""
        defaults = {
            'synchro_exe': 'C:\\Program Files (x86)\\Trafficware\\Version10\\Synchro10.exe',
            'synchro_dir': '',
            'model_path': '',
            'rows': 1000,
            'columns': 30,
            'update_los': 1
        }
        
        # Write the default settings to a JSON file
        with open('settings.json', 'w') as file:
            json.dump(defaults, file)
    
    try:
        with open('settings.json', 'r') as file:
            defaults = json.load(file)

    except FileNotFoundError:
        # If the file doesn't exist, create default settings and load them
        set_default()
        with open('settings.json', 'r') as file:
            defaults = json.load(file)
            
    return defaults  # Return the loaded or default settings


def center_window(x, y, master):
    """Calculates the position to center a window of size (x, y) on the screen."""
    screen_width, screen_height = master.winfo_screenwidth(), master.winfo_screenheight()
    x_coord, y_coord = int((screen_width - x) / 2), int((screen_height - y) / 2)
    
    # Prepare the size string for window geometry
    if x == 0 and y == 0:
        size = str()
    else:
        size = f'{x}x{y}'
    
    # Create the final position string for the window
    position = f'+{x_coord}+{y_coord}'
    return size + position  # Return the geometry string


def replace_slash(string):
    """Replaces forward slashes with backslashes in a given string."""
    return string.replace('/', '\\')


def get_row(worksheet, intersection):
    """Finds the appropriate row in the worksheet for a given intersection value."""
    for row in range(1, worksheet.max_row + 1):
        cell_value = worksheet.cell(row, 1).value  # Get the value in the first column of the row
        
        # If the cell is empty, return the row and method 'direct'
        if cell_value is None:
            return row, 'direct'
        # If the cell value matches the intersection, return the row and method 'direct'
        elif cell_value == intersection:
            return row, 'direct'
        # If the cell value is greater than the intersection, return the row and method 'insert'
        elif cell_value > intersection:
            method = 'insert'
            return row, method
        # If the cell value is less than the intersection, continue searching
        elif cell_value < intersection:
            for i in range(row, worksheet.max_row + 1):
                # If a subsequent cell value is greater than the intersection, return the row and method 'insert'
                if worksheet.cell(i, 1).value > intersection:
                    return i, 'insert'
                # If we reach the last row without finding a greater value, return the last row and method 'append'
                elif i == worksheet.max_row:
                    return i, 'append'


def get_sheet(wb, name):
    """Retrieves a sheet by name from the workbook, creating it if it does not exist."""
    for sheet in wb.sheetnames:
        if sheet == name:
            return wb[sheet]  # Return the existing sheet
        # If not found, create a new sheet with the specified name
        wb.create_sheet(title=name)
    
    return wb[name]  # Return the newly created sheet

"""
______________________ CLASSES ______________________
"""
# Stores details about a traffic scenario such as its name, hour, year, condition, and various data files.
class Scenario:
    def __init__(self, name):
        self.name = name
        self.hour = None
        self.year = None
        self.condition = None
        self.syn_file = None
        self.volumes = None
        self.los_data = None
        self.los_results = None
        self.model_data_column = None
        
    # Processes the scenario name to extract the hour (e.g., AM, PM, SAT).
    def parse(self):
        for hour in ['AM', 'PM', 'SAT']:
            if self.name.find(hour) != -1:
                self.name.replace(hour, '')
                self.hour = hour
                break

# Manages the settings for the application, including loading and saving settings to a JSON file.
# Builds a user interface to allow the user to configure settings such as default paths and boundaries.
class Settings:
    def __init__(self, master=None):
        self.master = master
        defaults = load_settings()
        # build ui
        self.settings_window = tk.Toplevel(master)
        self.main_frame = ttk.Frame(self.settings_window)
        self.notebook_1 = ttk.Notebook(self.main_frame)

        self.model_outer = ttk.Frame(self.notebook_1)
        self.model_outer.columnconfigure(0, weight=1)

        self.search_bounds = ttk.Labelframe(self.model_outer)
        self.search_bounds.columnconfigure(1, weight=1)

        self.row_label = ttk.Label(self.search_bounds)
        self.row_label.config(text='Rows:')
        self.row_label.grid(sticky='w')
        self.row_entry = ttk.Entry(self.search_bounds)
        _text_ = defaults['rows']
        self.row_entry.delete('0', 'end')
        self.row_entry.insert('0', _text_)
        self.row_entry.grid(column='1', row='0', sticky='nsew', padx=10)

        self.col_label = ttk.Label(self.search_bounds)
        self.col_label.config(text='Columns:')
        self.col_label.grid(sticky='w')
        self.col_entry = ttk.Entry(self.search_bounds)
        _text_ = defaults['columns']
        self.col_entry.delete('0', 'end')
        self.col_entry.insert('0', _text_)
        self.col_entry.grid(column='1', row='1', sticky='nsew', padx=10)

        self.search_bounds.config(height='200', text='Boundaries', width='200')
        self.search_bounds.rowconfigure(0, weight=1)
        self.search_bounds.grid(padx='0', pady='0', sticky='we')

        self.model_path_frame = ttk.Labelframe(self.model_outer)
        self.model_path_label = ttk.Label(self.model_path_frame)
        self.model_path_label.config(text='Default Path:')
        self.model_path_label.grid(sticky='w')
        self.model_path_frame.columnconfigure(1, weight=1)

        self.model_path_entry = ttk.Entry(self.model_path_frame)
        self.model_path_entry.config(text='Default Path:')
        _text_ = defaults['model_path']
        self.model_path_entry.delete('0', 'end')
        self.model_path_entry.insert('0', _text_)
        self.model_path_entry.grid(column='1', row='0', sticky='nsew', padx=10)

        self.model_browse = ttk.Button(self.model_path_frame)
        self.model_browse.config(text='Browse', command=self.model_browse_func)
        self.model_browse.grid(column='2', row='0')

        self.model_path_frame.config(height='200', text='Default Model Path', width='200')
        self.model_path_frame.grid(column='0', row='1', sticky='nsew')

        self.model_outer.config(height='200', width='200')
        self.model_outer.grid()
        self.notebook_1.add(self.model_outer, text='Model')

        self.syn_frame = ttk.Labelframe(self.notebook_1)
        self.syn_frame.columnconfigure(1, weight=1)

        self.syn_app_label = ttk.Label(self.syn_frame)
        self.syn_app_label.config(text='Synchro app location:')
        self.syn_app_label.grid()

        self.syn_app_entry = ttk.Entry(self.syn_frame)
        self.syn_app_entry.config(cursor='arrow')
        _text_ = defaults['synchro_exe']
        self.syn_app_entry.delete('0', 'end')
        self.syn_app_entry.insert('0', _text_)
        self.syn_app_entry.grid(column='1', row='0', sticky='nsew', padx=10)

        self.syn_dir_label = ttk.Label(self.syn_frame)
        self.syn_dir_label.config(text='Default Synchro folder:')
        self.syn_dir_label.grid(column='0', row='1')

        self.syn_dir_entry = ttk.Entry(self.syn_frame)
        self.syn_dir_entry.config(cursor='arrow')
        _text_ = defaults['synchro_dir']
        self.syn_dir_entry.delete('0', 'end')
        self.syn_dir_entry.insert('0', _text_)
        self.syn_dir_entry.grid(column='1', row='1', sticky='nsew', padx=10)

        self.syn_browse = ttk.Button(self.syn_frame)
        self.syn_browse.config(text='Browse', command=self.syn_browse_func)
        self.syn_browse.grid(column='2', row='0')

        self.syn_dir_browse = ttk.Button(self.syn_frame)
        self.syn_dir_browse.config(text='Browse', command=self.syn_dir_browse_func)
        self.syn_dir_browse.grid(column='2', row='1')

        self.syn_frame.config(height='200', text='Synchro')
        self.syn_frame.grid()
        self.notebook_1.add(self.syn_frame, text='Synchro')

        self.gen_tab_frame = ttk.Labelframe(self.notebook_1)
        self.gen_tab_frame.config(height='200', text='General', width='200')
        self.gen_tab_frame.pack(side='top')

        self.gen_label = ttk.Label(self.gen_tab_frame)
        self.gen_label.config(text='Update LOS by Default:')
        self.gen_label.grid()

        self.update_los_yes = ttk.Radiobutton(self.gen_tab_frame, text='Yes')
        self.update_los_yes.config(variable=self.master.update_los, value=1)
        self.update_los_yes.grid(row=0, column=1)

        self.update_los_no = ttk.Radiobutton(self.gen_tab_frame, variable=self.master.update_los, text='No')
        self.update_los_no.config(variable=self.master.update_los, value=0)
        self.update_los_no.grid(row=0, column=2)

        self.notebook_1.add(self.gen_tab_frame, text='General')
        self.notebook_1.config(height='200', width='200')
        self.notebook_1.pack(fill='both', side='top')
        self.button_frame = ttk.Frame(self.main_frame)
        self.button_frame.columnconfigure((0, 1), weight=1)

        self.done_button = ttk.Button(self.button_frame)
        self.done_button.config(text='Done', command=self.done)
        self.done_button.grid(sticky='e')

        self.cancel_button = ttk.Button(self.button_frame)
        self.cancel_button.config(text='Cancel', command=self.settings_window.destroy)
        self.cancel_button.grid(column='1', row='0', sticky='w')

        self.button_frame.config()
        self.button_frame.pack(fill='both', side='top')

        self.main_frame.config(height='200', width='200')
        self.main_frame.pack(fill='both', side='top')

        self.settings_window.config(height='200', width='200')
        self.settings_window.geometry('480x320')
        self.settings_window.title('Settings')

        # Main widget
        self.mainwindow = self.settings_window

    def model_browse_func(self):
        path = filedialog.askopenfile()
        if path:
            path = replace_slash(path.name)
            self.model_path_entry.delete('0', 'end')
            self.model_path_entry.insert('0', path)

    def syn_browse_func(self):
        path = filedialog.askopenfile()
        if path:
            path = replace_slash(path.name)
            self.syn_app_entry.delete('0', 'end')
            self.syn_app_entry.insert('0', path)

    def syn_dir_browse_func(self):
        path = replace_slash(filedialog.askdirectory())
        if path:
            self.syn_dir_entry.delete('0', 'end')
            self.syn_dir_entry.insert('0', path)

    def done(self):
        rows = int(self.row_entry.get())
        columns = int(self.col_entry.get())
        model_path = replace_slash(self.model_path_entry.get())
        syn_exe = replace_slash(self.syn_app_entry.get())
        update_los = self.master.update_los
        syn_dir = replace_slash(self.syn_dir_entry.get())
        defaults = {'synchro_exe': syn_exe,
                    'synchro_dir': syn_dir,
                    'model_path': model_path,
                    'rows': rows,
                    'columns': columns,
                    'update_los': update_los}

        self.master.synchro_app_path = defaults['synchro_exe']
        self.master.synchro_dir = defaults['synchro_dir']
        self.master.model_path = defaults['model_path']
        self.master.default_rows = defaults['rows']
        self.master.default_columns = defaults['columns']
        self.master.update_los = defaults['update_los']

        self.master.main_win.model_entry.delete('0', 'end')
        self.master.main_win.model_entry.insert('0', self.master.model_path)
        self.master.main_win.syn_entry.delete('0', 'end')
        self.master.main_win.syn_entry.insert('0', self.master.synchro_dir)

        with open('settings.json', 'w') as file:
            json.dump(defaults, file)

        self.settings_window.destroy()

# Represents the main window of the application.
# Contains methods to set up the UI, create various UI elements (labels, buttons), and handle user interactions.
# Provides functionality to launch other components like settings and file matching tools.
class MainWindow:
    def __init__(self, master=None):
        self.master = master
        self.data = None

        # UI Setup
        self.setup_ui()

    def setup_ui(self):
        self.master.columnconfigure(0, weight=1)
        self.master.rowconfigure(0, weight=1)

        self.frame_1 = ttk.Frame(self.master)
        self.frame_1.columnconfigure(1, weight=1)
        self.frame_1.rowconfigure((0, 1, 2), weight=1)

        # Create label and entry for Model file location
        self.model_entry = self.create_label_entry("Model file location:", self.master.model_path, 0, self.model_browse_func)

        # Create label and entry for Synchro file folder
        self.syn_entry = self.create_label_entry("Synchro file folder:", self.master.synchro_dir, 1, self.syn_browse_func)

        self.los_button = ttk.Checkbutton(self.frame_1, variable=self.master.update_los, text='Update LOS Table')
        self.los_button.grid(column=0, row=2)

        self.utilities = ttk.Labelframe(self.frame_1, text='Other Functions')
        self.utilities.grid(row=3, column=1)

        self.create_button("Copy Files", self.copy, 0, 0, parent=self.utilities, side='left')
        self.create_button("LOS Only", None, 0, 1, parent=self.utilities, side='left')  # Placeholder for LOS Only button

        notes = '''Instructions:\n
                1. Please ensure the Synchro files you wish to update are not open on any computer.\n
                2. Check that the model file is in our standard format.'''
        self.note_label = ttk.Label(self.frame_1, text=notes)
        self.note_label.grid(row=4, columnspan=3)

        self.frame_1.grid(sticky='nsew')

        # Main widget
        self.mainwindow = self.frame_1

    def create_label_entry(self, label_text, default_value, row, browse_command):
        label = ttk.Label(self.frame_1, text=label_text)
        label.grid(column=0, row=row)

        entry = ttk.Entry(self.frame_1)
        entry.insert(0, default_value)
        entry.grid(column=1, row=row, sticky='nsew')

        browse_button = ttk.Button(self.frame_1, text='Browse', command=browse_command)
        browse_button.grid(row=row, column=2)

        return entry  # Return the entry widget to assign it to an instance variable

    def create_button(self, text, command, col, row, parent=None, sticky=None, side=None):
        button = ttk.Button(parent if parent else self.frame_1, text=text, command=command)
        if side:
            button.pack(side=side)
        else:
            button.grid(column=col, row=row, sticky=sticky)

    def model_browse_func(self):
        path = filedialog.askopenfile(filetypes=[('Excel Files', '*.xlsx')])
        if path:
            self.update_entry(path.name, self.model_entry)

    def syn_browse_func(self):
        path = filedialog.askdirectory()
        if path:
            self.update_entry(path, self.syn_entry)

    def update_entry(self, path, entry):
        path = replace_slash(path)
        entry.delete(0, 'end')
        entry.insert(0, path)

    def launch_settings(self):
        Settings(self.master)

    def copy(self):
        Copier(self.master)

    def run(self):
        self.mainwindow.mainloop()
        
# Extends tk.Tk to serve as the base window for the application.
# Contains several methods for interacting with Synchro files and managing the application's 
# main operations, such as matching worksheet names, converting data, and handling errors.
class Base(tk.Tk):
    # Default values for the number of rows and columns in the application
    DEFAULT_ROWS = 1000
    DEFAULT_COLUMNS = 30
    
    # List of valid scenario types
    VALID_SCENARIOS = ['EXISTING', 'NO BUILD', 'BUILD']
    
    # Mapping of scenario conditions to their abbreviations
    SCENARIO_CONDITIONS = {
        'EXISTING': ['EX'],
        'NO BUILD': ['NB'],
        'BUILD': ['B', 'BD'],
        'IMPROVEMENT': ['IMP']
    }
    
    def __init__(self):
        # Initialize the Tkinter window
        super().__init__()
        self.title('Synchronizer')  # Set the title of the window
        self.geometry(center_window(500, 200, self))  # Center and size the window
        self.wm_attributes('-topmost', 0)  # Allow the window to be behind others
        
        # Initialize various attributes for managing the application state
        self.windows = {}  # Dictionary to hold child windows
        self.main_win = None  # Reference to the main window
        self.storage_dir = None  # Directory to store files
        self.model_sheet_name = str()  # Name of the current model sheet
        self.model_data = {}  # Dictionary to hold model data
        self.scenario_list = []  # List to store scenario objects
        self.scenario_data = {}  # Dictionary to hold scenario-related data
        self.selected_scenarios = []  # List of selected scenarios
        self.scenarios = []  # List to hold Scenario objects
        self.ws = None  # Reference to the currently active worksheet
        self.data_columns = []  # List of data columns

        # Load settings for the application
        defaults = {
            'synchro_exe': 'C:\\Program Files (x86)\\Trafficware\\Version10\\Synchro10.exe',
            'synchro_dir': '',
            'model_path': '',
            'rows': self.DEFAULT_ROWS,
            'columns': self.DEFAULT_COLUMNS,
            'update_los': 1
        }
        saved_settings = load_settings()  # Load previously saved settings

        # Set application paths and default settings, using saved values if available
        self.synchro_app_path = saved_settings.get('synchro_exe', defaults['synchro_exe'])
        self.synchro_dir = saved_settings.get('synchro_dir', defaults['synchro_dir'])
        self.model_path = saved_settings.get('model_path', defaults['model_path'])
        self.default_rows = saved_settings.get('rows', defaults['rows'])
        self.default_columns = saved_settings.get('columns', defaults['columns'])
        self.update_los = saved_settings.get('update_los', defaults['update_los'])

    def find_volume_data(self, extra_scenario=None):
        """
        Load volume data from the model workbook based on specified scenarios.

        Args:
            extra_scenario (str, optional): An additional scenario to consider.

        Returns:
            output.keys(): Keys of the scenario data collected from the model.
        """
        valid_scenarios = [extra_scenario] if extra_scenario else self.VALID_SCENARIOS
        output = {}

        wb = xl.load_workbook(filename=self.model_path, data_only=True)  # Load the model workbook
        self.model_sheet_name = max(wb.sheetnames, key=lambda sheet: similar(sheet, 'Model'), default=None)
        self.ws = wb[self.model_sheet_name]  # Set the active worksheet

        # Iterate through rows of the worksheet to find valid scenario data
        for row in range(1, self.ws.max_row):
            if self.ws.cell(row, 1).value == 1:  # Check if the row is valid
                year, scenario = None, None
                # Iterate through columns to extract year, scenario, and hour data
                for column in range(1, self.ws.max_column):
                    year_cell = self.ws.cell(row - 4, column).value
                    scenario_cell = self.ws.cell(row - 3, column).value
                    hour_cell = self.ws.cell(row - 2, column).value
                    
                    if year_cell is not None:
                        year = str(year_cell)  # Convert year to string
                    if scenario_cell is not None:
                        scenario = str(scenario_cell)  # Convert scenario to string
                    if hour_cell in ['AM', 'PM', 'SAT'] and scenario in valid_scenarios:
                        # Create a scenario name and check for duplicates
                        name = f"{year} {scenario} {hour_cell}"
                        if not any(found_scenario.name == name for found_scenario in self.scenarios):
                            sc = Scenario(name)  # Create a new Scenario object
                            sc.hour = hour_cell
                            sc.year = year
                            sc.condition = scenario
                            sc.model_data_column = column  # Store column index for the model data
                            self.match_syn_file(sc, self.synchro_dir)  # Match the corresponding .syn file
                            self.scenarios.append(sc)  # Add the scenario to the list
                        else:
                            messagebox.showwarning('Duplicate', 'One or more scenarios were duplicated and not added.')

        self.scenario_data = output  # Update scenario data
        return output.keys()  # Return the keys of the collected scenario data

    # Convert model volumes to Synchro UTDF
    def convert_utdf(self, scenario='test_write', column=5):
        # Open model to copy data
        # wb = xl.load_workbook(filename=model)
        # active = wb.active
        ws = self.ws  # need to make sure sheet is titled "Model"
        startColumn = 'C'  # Get direction column from user or default
        dataColumns = ['F', 'G', 'H']  # From scenarios to update

        volume_data = dict()
        movement_list = ['RECORDNAME', 'INTID']

        for row in range(15, self.default_rows):
            intersection = None
            direction = None

            cell = ws.cell(row, 1).value
            if type(cell) in [int, float]:
                intersection = int(cell)
                volume_data[intersection] = dict()
            cell = ws.cell(row, 3).value
            if type(cell) == str:
                direction = cell
                volume_data[intersection][direction] = dict()
            turn = ws.cell(row, 4).value
            if intersection and direction and turn:
                volume = ws.cell(row, column).value
                if volume is None:
                    volume = 0
                else:
                    volume = int(volume)
                volume_data[intersection][direction + turn] = volume
                if direction + turn not in movement_list:
                    movement_list.append(direction + turn)

        # Dictionary Format:
        # {intersection:{direction1:{L:0, T:0, R:0
        #                                        },
        #                            direction2:{},
        #                            direction3:{}}}

        # print(volume_data)
        file = self.storage_dir + '\\' + scenario + '.csv'
        with open(file, 'w', newline='') as f:
            f.write('[Lanes]\nLane Group Data\n')
            writer = csv.DictWriter(f, fieldnames=movement_list)
            writer.writeheader()
            for intid in volume_data:
                payload = volume_data[intid]
                payload['RECORDNAME'] = 'Volume'
                payload['INTID'] = intid
                writer.writerow(payload)
        return file


#_______________LOS_______________
    def update_report(self, scenarios, report_table=None):
        # If no report_table is specified, default to 'synchronizer results.xlsx'
        if report_table is None:
            report_table = 'synchronizer results.xlsx'
    
        # Combine the storage directory with the report_table name to get the full file path
        report_table = self.storage_dir + '\\' + report_table
    
        # Create a new Excel workbook and activate the default sheet
        wb = xl.Workbook()
        ws = wb.active
    
        # Rename the default sheet to 'AM'
        ws.title = 'AM'
    
        # Loop through each scenario in the scenarios list
        for scenario in scenarios:
            # Get the LOS (Level of Service) data and hour from the scenario object
            data = scenario.los_data
            hour = scenario.hour
    
            # Retrieve or create the sheet based on the scenario's hour (e.g., 'AM', 'PM')
            sheet = get_sheet(wb, hour)
    
            # Get the traffic condition (e.g., EXISTING, NO-BUILD, BUILD)
            condition = scenario.condition
    
            # Determine the column to store the data based on the condition
            if condition == 'EXISTING':
                column = 5
            elif condition == 'NO-BUILD':
                column = 8
            elif condition == 'BUILD':
                column = 11
            else:
                column = sheet.max_column  # If an unrecognized condition, use the last column
    
            # Loop through each intersection in the LOS data
            for intersection in data:
                # Get the row in the sheet corresponding to this intersection
                row, method = get_row(sheet, intersection)
                ov_los = None
                ov_delay = None
    
                # Loop through each turning movement in the intersection's data
                for turn_move, values in data[intersection].items():
                    # Special handling for 'overall' turning movement (aggregated data)
                    if turn_move == 'overall':
                        ov_delay = values['delay']
                        ov_los = values['los']
                        continue
    
                    # Generate a name for the movement (e.g., 'Left Turn') based on the config
                    movement_name = label(turn_move, values.get('config', ''))
                    if movement_name:
                        # Initialize lists to store various values (e.g., v/c ratio, LOS, delay)
                        vc_ratios = list()
                        los_values = list()
                        delays = list()
                        app_los_values = list()
                        app_delays = list()
                        last_move = turn_move[:2]
    
                        # Process each direction for the movement (e.g., EB, WB)
                        for direction in movement_name:
                            search = turn_move[:2] + direction
    
                            # Ensure the search key exists before retrieving the data
                            if search not in data[intersection].keys():
                                continue
    
                            # Append the corresponding values for v/c ratio, LOS, and delay
                            vc_ratios.append(data[intersection][search].get('vc_ratio', ''))
                            los_values.append(data[intersection][search].get('ln_los', ''))
                            delays.append(data[intersection][search].get('ln_delay', ''))
                            app_los_values.append(data[intersection][search].get('app_los', ''))
                            app_delays.append(data[intersection][search].get('app_delay', ''))
    
                        # Take the maximum values for each metric
                        vc = max(vc_ratios)
                        los = max(los_values)
                        delay = max(delays)
                        app_los = max(app_los_values)
                        app_delay = max(app_delays)
    
                        # If all values are empty, skip the current movement
                        if vc == '' and los == '' and delay == '':
                            continue
    
                        # Write the data into the sheet based on the method (direct, insert, append)
                        if method == 'direct':
                            sheet.cell(row, 1).value = intersection
                            sheet.cell(row, 3).value = turn_move[:2]
                            sheet.cell(row, 4).value = movement_name
                            sheet.cell(row, column).value = vc
                            sheet.cell(row, column + 1).value = los
                            sheet.cell(row, column + 2).value = delay
                            row += 1
    
                        elif method == 'insert':
                            sheet.insert_rows(row)
                            sheet.cell(row, 1).value = intersection
                            sheet.cell(row, 3).value = turn_move[:2]
                            sheet.cell(row, 4).value = movement_name
                            sheet.cell(row, column).value = vc
                            sheet.cell(row, column + 1).value = los
                            sheet.cell(row, column + 2).value = delay
                            row += 1
    
                        elif method == 'append':
                            row += 1
                            sheet.cell(row, 1).value = intersection
                            sheet.cell(row, 3).value = turn_move[:2]
                            sheet.cell(row, 4).value = movement_name
                            sheet.cell(row, column).value = vc
                            sheet.cell(row, column + 1).value = los
                            sheet.cell(row, column + 2).value = delay
    
                        # If there are no approach LOS and delay values, skip to the next move
                        if app_delay == '' and app_los == '':
                            continue
    
                        # Write the approach LOS and delay if the last movement is different
                        if last_move and turn_move != last_move:
                            if method == 'direct':
                                sheet.cell(row, 1).value = intersection
                                sheet.cell(row, 3).value = 'Approach'
                                sheet.cell(row, column + 1).value = app_los
                                sheet.cell(row, column + 2).value = app_delay
                                row += 1
    
                            elif method == 'insert':
                                sheet.insert_rows(row)
                                sheet.cell(row, 1).value = intersection
                                sheet.cell(row, 3).value = 'Approach'
                                sheet.cell(row, column + 1).value = app_los
                                sheet.cell(row, column + 2).value = app_delay
                                row += 1
    
                            elif method == 'append':
                                row += 1
                                sheet.cell(row, 1).value = intersection
                                sheet.cell(row, 3).value = 'Approach'
                                sheet.cell(row, column + 1).value = app_los
                                sheet.cell(row, column + 2).value = app_delay
    
                    # Write the overall LOS and delay if available
                    if ov_los and ov_delay:
                        if method == 'direct':
                            sheet.cell(row, 1).value = intersection
                            sheet.cell(row, 3).value = 'Overall'
                            sheet.cell(row, column + 1).value = ov_los
                            sheet.cell(row, column + 2).value = ov_delay
                            row += 1
    
                        elif method == 'insert':
                            sheet.insert_rows(row)
                            sheet.cell(row, 1).value = intersection
                            sheet.cell(row, 3).value = 'Overall'
                            sheet.cell(row, column + 1).value = ov_los
                            sheet.cell(row, column + 2).value = ov_delay
                            row += 1
    
                        elif method == 'append':
                            row += 1
                            sheet.cell(row, 1).value = intersection
                            sheet.cell(row, 3).value = 'Overall'
                            sheet.cell(row, column + 1).value = ov_los
                            sheet.cell(row, column + 2).value = ov_delay
    
        # Save the workbook to the specified file
        wb.save(report_table)
    
        # Return the path to the report file
        return report_table


class ProgressWindow:
    def __init__(self, master=None):
        self.master = master
        # build ui
        self.progress_window = tk.Toplevel(self.master)
        self.progress_window.geometry(center_window(400, 400, self.master))
        self.progress_window.columnconfigure(0, weight=1)
        self.progress_frame = ttk.Frame(self.progress_window)
        self.progress_frame.columnconfigure(0, weight=1)
        self.progress_frame.columnconfigure(1, weight=0)
        self.status_text_box = tk.Text(self.progress_frame)
        self.status_text_box.config(autoseparators='false')
        self.status_text_box.grid(column='0', row='0', sticky='nsew')
        self.scrollbar_3 = ttk.Scrollbar(self.progress_frame, command=self.status_text_box.yview)
        self.scrollbar_3.grid(column='1', row='0', sticky='nsew')
        self.status_text_box.configure(yscrollcommand=self.scrollbar_3.set)
        # self.button_3 = ttk.Button(self.progress_frame)
        # self.button_3.config(text='button_3')
        # self.button_3.grid(column='0', row='1', sticky='s')
        # self.progress_frame.config(height='200', width='200')
        self.progress_frame.grid(padx=10, pady=10, sticky='nsew')
        # self.progress_window.config(height='200', width='200')
        self.progress_window.title('Program Status')
        self.progress_window.after(6000, self.run)

    def run(self):
        time.sleep(2)
        # success = self.master.startup()
        # if success != 0:
        #     self.status_text_box.insert('end', 'Failed to start Synchro\n')
        #     return
        for scenario_obj in self.master.selected_scenarios:
            scenario = scenario_obj.name
            filename = scenario_obj.syn_file
            column = scenario_obj.model_data_column
            process_update = 'Processing: ' + scenario + '\n'
            self.status_text_box.insert('end', process_update)
            utdf_volumes = self.master.convert_utdf(scenario=scenario, column=column)
            self.status_text_box.insert('end', 'Importing volumes to Synchro... \n')
            self.master.import_to_synchro(filename, utdf_volumes)
            self.status_text_box.insert('end', 'Import complete\n')

            if self.master.update_los:
                self.status_text_box.insert('end', 'Exporting LOS data from Synchro...\n')
                scenario_obj.los_results = self.master.export_from_synchro(scenario)
                time.sleep(5)
                self.status_text_box.insert('end', 'Export complete\n')
                scenario_obj.los_data = standardize(scenario_obj.los_results)

        if self.master.update_los:
            self.status_text_box.insert('end', 'Writing LOS data to excel file\n')
            output_results = self.master.update_report(self.master.selected_scenarios)
            self.status_text_box.insert('end', 'Write complete\nThe program has finished\n')
            self.status_text_box.insert('end', f'LOS results are saved at: {output_results}')


class Copier:
    def __init__(self, master=None):
        self.window = tk.Toplevel(master)
        self.window.columnconfigure(1, weight=1)

        self.old_dir_label = ttk.Label(self.window, text='Copy from:')
        self.old_dir_label.grid(row=0, column=0, sticky='e', padx=10)

        self.old_dir_entry = ttk.Entry(self.window)
        self.old_dir_entry.grid(row=0, column=1, sticky='ew')

        self.old_dir_button = ttk.Button(self.window, text='Browse', command=self.browse)
        self.old_dir_button.bind('<Button 1>', self.browse)
        self.old_dir_button.grid(row=0, column=2)

        self.new_dir_label = ttk.Label(self.window, text='Copy to:')
        self.new_dir_label.grid(row=1, column=0, sticky='e', padx=10)

        self.new_dir_entry = ttk.Entry(self.window)
        self.new_dir_entry.grid(row=1, column=1, sticky='ew')

        self.new_dir_button = ttk.Button(self.window, text='Browse', command=self.browse)
        self.new_dir_button.bind('<Button 1>', self.browse)
        self.new_dir_button.grid(row=1, column=2)

        self.new_date_label = ttk.Label(self.window, text='New date:')
        self.new_date_label.grid(row=2, column=0, sticky='e', padx=10)

        self.new_date_entry = ttk.Entry(self.window)
        self.new_date_entry.grid(row=2, column=1, sticky='ew')

        self.start = ttk.Button(self.window, text='Start', command=self.copy_files)
        self.start.grid(row=3, columnspan=3)

        # self.check_syn = ttk.Checkbutton(self.window, text='Synchro')
        # self.check_syn.grid(row=3)
        #
        # self.check_pdf = ttk.Checkbutton(self.window, text='Synchro')
        # self.check_pdf.grid(row=4)

    def browse(self, event):
        file = filedialog.askdirectory()
        if file is None:
            return
        row = event.widget.grid_info()['row']
        if row == 0:
            entry = self.old_dir_entry
        else:
            entry = self.new_dir_entry

        entry.delete(0, 'end')
        entry.insert('end', file)

    def copy_files(self):
        pattern = re.compile('[0-9]*')
        old_dir = self.old_dir_entry.get()

        for file in os.scandir(old_dir):
            print(file)
            if not file.name.endswith('syn'):
                continue
            new_date = self.new_date_entry.get()
            old_date = re.match(pattern, file.name).group(0)
            if old_date == '':
                new_file_name = new_date + file.name
            else:
                new_file_name = file.name.replace(old_date, new_date)
            new_path = self.new_dir_entry.get() + '\\' + new_file_name
            copy(file.path, new_path)
        self.window.destroy()


if __name__ == "__main__":
    read_input_file("test-input.xlsx")

