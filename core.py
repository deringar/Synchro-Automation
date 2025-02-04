# -*- coding: utf-8 -*-
"""
Created on Fri Sep 27 2024
Last modified on Thurs Oct 3 2024

@authors: philip.gotthelf, alex.dering - Colliers Engineering & Design
"""

# main_window.py

import tkinter as tk  # Import the Tkinter module for GUI development.
# Import themed widgets from Tkinter for better styling.
import tkinter.ttk as ttk
# Used for comparing sequences and finding similarities.
from difflib import SequenceMatcher
# Import specific Tkinter features for message boxes and file dialogs.
from tkinter import messagebox, filedialog
import csv  # Module to handle CSV file operations.
import openpyxl as xl  # Used for working with Excel files (.xlsx format).
# OS module for interacting with the operating system (file paths, etc.).
import os
import re  # Regular expression module for pattern matching in strings.
import time  # Module for time-related functions.
import json  # JSON module to parse and manipulate JSON data.
# Import ordered dictionary to maintain the order of keys.
from collections import OrderedDict
from shutil import copy  # Used to copy files or directories.
from openpyxl import load_workbook, Workbook
import pandas as pd

def write_to_excel(file_path, lane_groups, delay, vc_ratio, los):
    # Helper function to transform each sublist into a dictionary with lane directions
    def separate_characters(sublist):
        result_dict = {}
        for item in sublist:
            # Extract the prefix and 'L', 'T', 'R' characters
            rest_chars = ''.join([char for char in item if char not in 'LTR'])
            separated_chars = [char for char in item if char in 'LTR']

            if rest_chars in result_dict:
                # Append characters if the key exists
                result_dict[rest_chars].extend(separated_chars)
            else:
                # Create new entry for the key
                result_dict[rest_chars] = separated_chars

        return result_dict

    # Function to enumerate the lane groups and create a dictionary
    def enumerate_result_list(result):
        enumerated_dict = {}
        for index, item in enumerate(result, start=1):
            if isinstance(item, list):  # Ensure each item is processed as a list
                enumerated_dict[index] = separate_characters(
                    item)  # Transform each list into a dictionary
            else:
                # Handle already existing dictionaries
                enumerated_dict[index] = item
        return enumerated_dict

    # Helper function to write headers
    def write_headers(ws, start_col='C'):
        headers = ['V/c', 'LOS', 'Delay']
        for idx, header in enumerate(headers):
            col = chr(ord(start_col) + idx)  # Dynamic column calculation
            ws[f'{col}2'] = header  # Write the header to row 2

    # Get the file name without extension
    file_with_ext = os.path.basename(file_path)
    file_name = os.path.splitext(file_with_ext)[0]

    # Enumerate the lane groups
    intersection_data = enumerate_result_list(lane_groups)

    # Create a new Excel workbook and add a sheet
    wb = Workbook()
    ws = wb.active  # Get the active worksheet

    # Write the file name in cell A1
    ws['A1'] = file_name

    # Write the headers for V/C, LOS, and Delay in row 2 (starting at column C)
    write_headers(ws, 'C')
    write_headers(ws, 'F')
    write_headers(ws, 'I')

    # Define the order of keys to process
    key_order = ['EB', 'WB', 'NB', 'SB', 'NE', 'NW', 'SE', 'SW']

    # Keep track of the last used row
    current_row = 2  # Starting at row 2 (A2 will be filled first)

    # Iterate through the enumerated intersection data
    for intersection_id, data in intersection_data.items():
        # Write the intersection ID in column A
        ws[f'A{current_row}'] = intersection_id

        # Move to the next row after writing the intersection ID
        current_row += 1

        # Ensure that 'data' is a dictionary before attempting to access its keys
        if isinstance(data, dict):
            # For each key in the specified order:
            for key in key_order:
                # Write the key in column A
                ws[f'A{current_row}'] = key

                # Check if there's a corresponding value in the lane data
                if key in data and data[key]:  # If values exist
                    # Write the corresponding values in column B, starting from the current row
                    for idx, item in enumerate(data[key]):
                        # Write each value downwards in column B
                        ws[f'B{current_row + idx}'] = item
                    # Update the current_row to the next empty row after writing all values
                    current_row += len(data[key])
                else:
                    # If no values exist, write an empty cell in column B
                    # Explicitly write an empty string
                    ws[f'B{current_row}'] = ''
                    current_row += 1  # Move to the next row for the next key

    # Write V/C, LOS, and Delay data into their respective columns, ensuring empty entries are included
    # Start from row 3 (the row after the headers)

    # Write V/C Ratio
    current_row_vc = 3  # Start writing V/C values
    for idx, vc_list in enumerate(vc_ratio):
        if idx < len(intersection_data):  # Ensure we don't exceed the intersection data length
            for item in vc_list:
                # Write each value downwards in column C
                ws[f'C{current_row_vc}'] = item
                current_row_vc += 1  # Move to the next row
            # Fill empty cells if the length of vc_list is shorter than the max
            while current_row_vc < (3 + len(vc_ratio)):
                ws[f'C{current_row_vc}'] = ''  # Write an empty string
                current_row_vc += 1

    # Write LOS values
    current_row_los = 3  # Start writing LOS values
    for idx, los_list in enumerate(los):
        if idx < len(intersection_data):  # Ensure we don't exceed the intersection data length
            for item in los_list:
                # Write each value downwards in column D
                ws[f'D{current_row_los}'] = item
                current_row_los += 1  # Move to the next row
            # Fill empty cells if the length of los_list is shorter than the max
            while current_row_los < (3 + len(los)):
                ws[f'D{current_row_los}'] = ''  # Write an empty string
                current_row_los += 1

    # Write Delay values
    current_row_delay = 3  # Start writing Delay values
    for idx, delay_list in enumerate(delay):
        if idx < len(intersection_data):  # Ensure we don't exceed the intersection data length
            for item in delay_list:
                # Write each value downwards in column E
                ws[f'E{current_row_delay}'] = item
                current_row_delay += 1  # Move to the next row
            # Fill empty cells if the length of delay_list is shorter than the max
            while current_row_delay < (3 + len(delay)):
                ws[f'E{current_row_delay}'] = ''  # Write an empty string
                current_row_delay += 1

    # Save the workbook
    excel_file_path = f"{file_name}_results.xlsx"
    wb.save(excel_file_path)
    print(f"Intersection data written to {excel_file_path}")


def separate_characters(result):
    # Initialize a list to hold the dictionaries
    transformed_results = []

    # Iterate through each sublist in the result
    for sublist in result:
        # Initialize a dictionary for this sublist
        result_dict = {}

        # Process each string in the sublist
        for item in sublist:
            # Extract the characters that are not 'L', 'T', or 'R' for the prefix
            # Characters other than L, T, R
            rest_chars = ''.join([char for char in item if char not in 'LTR'])
            # Characters that are L, T, or R
            separated_chars = [char for char in item if char in 'LTR']

            # Use the prefix as the key in the dictionary
            if rest_chars in result_dict:
                # Append to the existing entry if the key already exists
                result_dict[rest_chars].extend(separated_chars)
            else:
                # Create a new entry if the key does not exist
                result_dict[rest_chars] = separated_chars

        # Append the dictionary to the list of transformed results
        transformed_results.append(result_dict)

    return transformed_results


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
        # HeavyVehicles is in start_column + 2 (e.g., H)
        heavy_vehicles_col = start_column + 2

        # Get the header name from row 1 of the start_column (e.g., F1, I1, etc.)
        file_name_header = sheet.cell(row=1, column=start_column).value
        if not file_name_header:
            print(
                f"Skipping columns starting at {start_column} as no header was found in row 1.")
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
            output_sheet.cell(row=output_start_row + 2,
                              column=1).value = "HeavyVehicles"
            output_sheet.cell(row=output_start_row,
                              column=2).value = intersection_id
            output_sheet.cell(row=output_start_row + 1,
                              column=2).value = intersection_id
            output_sheet.cell(row=output_start_row + 2,
                              column=2).value = intersection_id

            # Process each direction-turn within the intersection
            for direction_turn, info in turns.items():
                row_found = info['row']

                # Read data from the specified columns for the current row
                volume = sheet.cell(row=row_found, column=volume_col).value
                phf = sheet.cell(row=row_found, column=phf_col).value
                heavy_vehicles = sheet.cell(
                    row=row_found, column=heavy_vehicles_col).value

                # Write the data into the output sheet under the correct direction-turn column
                header_column = info['header_column']
                output_sheet.cell(row=output_start_row,
                                  column=header_column).value = volume
                output_sheet.cell(row=output_start_row + 1,
                                  column=header_column).value = phf
                output_sheet.cell(row=output_start_row + 2,
                                  column=header_column).value = heavy_vehicles

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


""" STEP 1 """


def read_input_file(file_path):
    # Load the input workbook and select the active sheet
    workbook = load_workbook(filename=file_path)
    sheet = workbook.active

    # Define headers for the output sheet
    headers = [
        "RECORDNAME", "INTID", "NBL", "NBT", "NBR",
        "SBL", "SBT", "SBR", "EBL", "EBT", "EBR",
        "WBL", "WBT", "WBR", "NWR", "NWL", "NWT", "NEL", "NET", "NER",
        "SEL", "SER", "SET", "SWL", "SWR", "SWT", "PED", "HOLD"
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
            # Default is None (not found)
            turn_values = {"L": None, "T": None, "R": None}
            for search_row in range(found_row, sheet.max_row + 1):
                turn_value = sheet.cell(search_row, column=4).value
                if turn_value in ["L", "T", "R"]:
                    # Store the row number for each turn type found
                    turn_values[turn_value] = search_row
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
        print(
            f"Direction-turn results for intersection {intersection_id}: {direction_turn_results}")

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

    write_direction_data_to_files(
        sheet, matched_results, relevant_columns, headers=headers, output_start_row=4)

    # Return intersection results if needed elsewhere
    return intersection_results


"""
 ~ Data extraction functions ~

    * parse_minor_lane_mvmt(lines, start_line, end_line)
    * process_directions(twsc_summary_results)
    * parse_overall_data_v2(file_path)
    * parse_twsc_approach(df)
    * extract_data_to_csv(file_path, output_file)
    * parse_lane_configs(int_lane_groups, intersection_ids)
"""


def parse_minor_lane_mvmt(lines, start_line, end_line):
    """
        Parse the "Minor Lane/Major Mvmt" data between the start and end lines.
        This function extracts the delay, V/C ratio, and LOS from lines containing these terms.
        Helper function to the parse_overall_data function.
    """

    search_phrase = "Minor Lane/Major Mvmt"
    search_terms = [r'\bControl Delay\b',
                    r'\bV/C Ratio\b', r'\bLOS\b', r'\bCapacity\b']

    # Initialize lists for each search term
    result = []
    delay_results = []
    vc_ratio_results = []
    los_results = []
    capacity_results = []

    # Search for "Minor Lane/Major Mvmt" in the provided line range
    for line_number in range(start_line, end_line):
        line = lines[line_number]
        if search_phrase in line:
            # Process the "Minor Lane/Major Mvmt" line to get the directions
            after_phrase = line.split(search_phrase)[1].strip()
            cleaned_value = ' '.join(after_phrase.split()).strip()
            result = (cleaned_value.split())
            # print(f"Result: {result}")
            # Now search for Delay, V/C Ratio, and LOS in lines below the current line
            for term in search_terms:
                term_results = []  # Temporary list to hold results for the current term
                for search_line_number in range(line_number + 1, end_line):
                    # Remove whitespace for accurate matching
                    line = lines[search_line_number].strip()
                    if re.search(term, line, re.IGNORECASE):
                        # Extract values after the term
                        # print(f"Found term: {term} in line: {line}")

                        if 'control delay' in term.lower() or 'v/c ratio' in term.lower() or 'capacity' in term.lower():
                            # For control delay or V/C ratio, we extract numbers (floats or '-')
                            numbers = re.findall(r'(\d+\.\d+|\d+|-)', line)
                            term_results.extend(
                                [float(num) if num != '-' else num for num in numbers])

                        elif 'los' in term.lower():
                            # For LOS, we extract single capital letters (A-F) or '-'
                            capital_letters = re.findall(r'\b[A-F]\b|-', line)
                            term_results.extend(capital_letters)

                # Add the term results to the corresponding results list
                if 'control delay' in term.lower():
                    delay_results.append(
                        term_results if term_results else ['-'])
                elif 'v/c ratio' in term.lower():
                    vc_ratio_results.append(
                        term_results if term_results else ['-'])
                elif 'los' in term.lower():
                    los_results.append(term_results if term_results else ['-'])
                elif 'capacity' or r'^cap' in term.lower():
                    capacity_results.append(
                        term_results if term_results else ['-'])

    # Combine the results into tuples for easier reading
    merged_results = []
    for vc_list, los_list, delay_list, capacity_list in zip(vc_ratio_results, los_results, delay_results, capacity_results):
        merged_results = (
            list(zip(vc_list, los_list, delay_list, capacity_list)))

    # Return the parsed results for integration with other parsing logic
    return result, merged_results


def integrate_awsc_data(awsc_data, combined_data):
    """
    Merges the AWSC lane data into the existing data handling structure,
    ensuring each intersection is collected correctly.
    """
    for awsc_entry in awsc_data:
        intersection_id = awsc_entry.get("ID")
        if not intersection_id:
            continue

        formatted_entry = {"ID": intersection_id}
        
        for lane, values in awsc_entry.items():
            if lane == "ID":
                continue
            # Directly unpack the tuple
            v_c_ratio, los, delay, cap = values
            formatted_entry[lane] = (v_c_ratio, los, delay, cap)
        
        combined_data.append(formatted_entry)
    
    return combined_data


def process_directions_awsc(lane_data):
    """
    Processes lane names using process_directions to get actual movement labels.
    """
    processed_data = []
    
    for entry in lane_data:
        processed_entry = {"ID": entry["ID"]}
        
        for lane, values in entry.items():
            if lane == "ID":
                continue
            processed_lane = process_directions(lane)
            processed_entry[processed_lane] = values
        
        processed_data.append(processed_entry)
    
    return processed_data


def parse_awsc_data(df):
    """
    Parses AWSC data from the dataframe and extracts lane group details,
    v/c ratio, LOS, delay, and capacity.
    """
    awsc_data = []
    intersection_id = None

    for index, row in df.iterrows():
        line = str(row[0]).strip()
        print(line)
        if line.isdigit():
            intersection_id = int(line)
            print(f"Found intersection ID: {intersection_id} at row {index}")
            continue

        if line.lower() == "lane":
            print(f"Processing lane data for intersection ID: {intersection_id} at row {index}...{row}")
            lane_data = {"ID": str(intersection_id)}
            direction_columns = {}

            for direction in ["EB", "WB", "NB", "SB", "NE", "NW", "SE", "SW"]:
                for col_index, cell in enumerate(row.values):
                    if re.fullmatch(f'{direction}(Ln\d+)?', str(cell)):
                        direction_columns[direction] = col_index
                        print(f"Found direction '{direction}' in column {col_index}")

            next_row_index = index + 1
            while next_row_index + 1 < len(df):
                next_row = df.iloc[next_row_index]
                print(df.iloc[next_row_index + 1])
                print(f"Line {next_row_index}: {next_row[0]}")

                for direction, col in direction_columns.items():
                    lane_data.setdefault(direction, ("-", "-", "-", "-"))
                    v_c_ratio, los, delay, cap = lane_data[direction]

                    if "V/C Ratio" in str(next_row.iloc[0]):
                        print(f"Found V/C Ratio at row {next_row_index}")
                        v_c_ratio = next_row[col] if pd.notna(next_row[col]) else '-'
                        print(f"  {direction} V/C Ratio: {v_c_ratio}")
                    elif "LOS" in str(next_row.iloc[0]):
                        print(f"Found LOS at row {next_row_index}")
                        los = next_row[col] if pd.notna(next_row[col]) else '-'
                        print(f"  {direction} LOS: {los}")
                    elif "Delay" in str(next_row.iloc[0]):
                        print(f"Found Delay at row {next_row_index}")
                        delay = next_row[col] if pd.notna(next_row[col]) else '-'
                        print(f"  {direction} Delay: {delay}")
                    elif "Cap" in str(next_row.iloc[0]):
                        print(f"Found Capacity at row {next_row_index}")
                        cap = next_row[col] if pd.notna(next_row[col]) else '-'
                        print(f"  {direction} Capacity: {cap}")

                    lane_data[direction] = (v_c_ratio, los, delay, cap)

                next_row_index += 1
                if re.match(r'^\d+:', str(next_row[0])):
                    print(f"Encountered new intersection or empty row at {next_row_index}, stopping data collection for current intersection.")
                    break

            awsc_data.append(lane_data)
            print(f"Completed data collection for intersection ID: {intersection_id} -> {lane_data}")

    print(f"AWSC Data Parsed: {awsc_data}")
    return awsc_data


def process_directions(twsc_summary_results, lane_configs):
    processed_list = []
    original_key_list = []
    combined_mvmt_names = []
    # print("Processing Directions...\n")
    # print(twsc_summary_results)
    # print(f"\n{lane_configs}\n")
    for entry in twsc_summary_results:
        # Start with a dictionary containing just the ID
        processed_dict = {"ID": entry["ID"]}
        original_key_dict = {"ID": entry["ID"]}

        # processed_list_str = ''
        # original_list_str = ''
        # Retrieve the lane configuration for the current intersection
        intersection_id = int(entry["ID"])
        lane_config = next(
            (config for config in lane_configs if config["Intersection ID"]
             == intersection_id), None
        )

        combined_mvmt = []

        # Loop through the dictionary to process directional keys
        for key, value in entry.items():
            if key == "ID":
                continue  # Skip the ID key itself
            # Split the direction from the suffix
            # The first two characters are the direction (EB, WB, NB, SB, NE, NW, SE, SW)
            direction = key[:2]
            suffix = key[2:]  # The remaining part is the suffix (Ln1, T, etc.)

            config_amount = len(lane_config[direction]) if (
                lane_config and direction in lane_config) else 1

            # Add suffixes to the original_key_dict
            if direction in original_key_dict:
                original_key_dict[direction].append(suffix)
            else:
                original_key_dict[direction] = [suffix]  # Initialize as a list

            # Handle Ln suffix by matching with the lane configuration
            if "Ln" in suffix and lane_config and direction in lane_config:
                try:  # Try matching
                    lane_index = int(suffix[2:]) - 1
                    if 0 <= lane_index < len(lane_config[direction]):
                        suffix = lane_config[direction][lane_index]
                except (ValueError, IndexError):
                    pass  # If parsing or index retrieval fails, keep the original suffix

            # Determine storage format based on config_amount
            if direction not in processed_dict:
                # Initialize as list or string
                processed_dict[direction] = [] if config_amount > 1 else ""

            if config_amount > 1:
                processed_dict[direction].append(suffix)
            else:
                processed_dict[direction] += suffix

        for direction, value in processed_dict.items():
            if direction != "ID":
                # Join lists into strings for combined movement names
                combined_mvmt.append(
                    direction + ''.join(value) if isinstance(value, list) else direction + value)

        combined_mvmt_names.append(combined_mvmt)

        # Append the processed dictionaries to their respective lists
        processed_list.append(processed_dict)
        original_key_list.append(original_key_dict)

    # print(f"\nOriginal Keys:\n{original_key_list}")
    # print(f"\nUpdated Keys:\n{processed_list}")
    # print(f"\nMerged Names:\n{combined_mvmt_names}")

    return processed_list, original_key_list, combined_mvmt_names


def parse_overall_data_v2(file_path, df):
    """
        Function to handle the parsing of the summary data
    """

    int_regex = r'^\d+:'  # Regex to match lines that start with an integer followed by a colon

    search_phrases = ["Minor Lane/Major Mvmt",
                      "Intersection Summary"]

    # Lists to store results
    synchro_results = []
    hcm_results = []
    twsc_results = []
    awsc_results = []

    # List to store matching line numbers
    matching_lines = []

    # Open and read the file into memory
    with open(file_path, 'r') as file:
        lines = file.readlines()  # Read all lines

    # Find the line numbers containing the integer followed by a colon
    for line_number, line in enumerate(lines, start=1):
        if re.match(int_regex, line):  # Use re.match to check if the line starts with the regex
            matching_lines.append(line_number)

    intersection_index = 0

    # Now search from each line in matching_lines until the next line matching the regex
    for start_line in matching_lines:
        # Get the ID from the corresponding line in 'lines'
        # Accessing lines using start_line - 1
        id_match = re.match(int_regex, lines[start_line - 1])
        id_value = id_match.group(0).strip(
            ':') if id_match else None  # Get the ID before the colon

        # print(f"\nProcessing ID {id_value} at line {start_line}")

        # Process Synchro Results first
        found_phrase = False
        for line_number in range(start_line, len(lines)):
            line = lines[line_number]

            # Check for a new ID match before processing further
            new_id_match = re.match(int_regex, line)
            if new_id_match:
                new_id_value = new_id_match.group(0).strip(':')
                # print(f"Found new ID {new_id_value} at line {line_number}")
                id_value = new_id_value  # Update the ID to the new one found

            if "Intersection Summary" in line:
                # print(f"Found 'Intersection Summary' at line {line_number}")

                found_phrase = True  # Mark that we found the phrase

                # Set the end line to the next empty line starting from this line
                end_line = line_number + 1
                while end_line < len(lines) and lines[end_line].strip() != '':
                    if "HCM" in lines[end_line]:
                        break
                    end_line += 1  # Continue until we find an empty line

                if "HCM" in lines[end_line]:
                    # print(f"Found 'HCM' at line {end_line}, skipping Synchro block for this ID\n")
                    continue

                # print(f"Synchro block ends at line {end_line}\n")

                # Initialize values to None
                vc_ratio_value = None
                los_value = None
                delay_value = None

                # Now process the following lines for the search terms until the next blank line
                for search_line_number in range(line_number + 1, end_line):
                    line = lines[search_line_number]

                    # print(f"Processing Synchro data at line {search_line_number}: {line.strip()}")

                    # Check for 'v/c ratio' and extract the next float
                    # if re.search(r'v/c ratio', line, re.IGNORECASE):
                    #     float_match = re.search(r'(\d+\.\d+|\d+)', line)
                    #     if float_match:
                    #         vc_ratio_value = float(float_match.group(0))
                    #         # print(f"Extracted v/c ratio: {vc_ratio_value}")

                    # Check for 'delay' and extract the next float
                    if re.search(r'delay', line, re.IGNORECASE):
                        float_match = re.search(r'(\d+\.\d+|\d+)', line)
                        if float_match:
                            delay_value = float(float_match.group(0))
                            # print(f"Extracted delay: {delay_value}")

                    # Check for 'LOS' and extract the next capital letter (A-F)
                    if re.search(r'LOS', line, re.IGNORECASE):
                        capital_match = re.search(r'\b[A-F]\b', line)
                        if capital_match:
                            los_value = capital_match.group(0)
                            # print(f"Extracted LOS: {los_value}")

                # print(f"Final Synchro values for ID {id_value}: v/c ratio={vc_ratio_value}, delay={delay_value}, LOS={los_value}")

                # Store Synchro results only if ID is not already present
                if not (los_value is None and delay_value is None) and found_phrase:
                    if not any(result['ID'] == id_value for result in synchro_results):
                        synchro_results.append({
                            'ID': id_value,
                            'v/c ratio': vc_ratio_value if vc_ratio_value is not None else '-',
                            'los': los_value if los_value is not None else '-',
                            'delay': delay_value if delay_value is not None else '-',
                            'index': intersection_index
                        })
                        intersection_index += 1

                # Stop further processing of Synchro block and move on to HCM
                break  # Exit after processing this block for Synchro

            # Skip lines between the ID and the next search phrase
            if found_phrase:
                break  # Stop looking at this block and continue with HCM

        # Now search for HCM Results
        found_phrase = False  # Reset for HCM block
        for line_number in range(start_line, len(lines)):
            line = lines[line_number]

            # Check if a new ID appears before parsing the HCM block
            new_id_match = re.match(int_regex, lines[line_number - 1])
            if new_id_match:
                id_value = new_id_match.group(0).strip(':')
                # print(f"Updated ID to {id_value} for HCM parsing at line {line_number}")

            # Process only HCM-related lines
            for search_phrase in search_phrases:
                if search_phrase in line:
                    found_phrase = True
                    phrase_found = search_phrase  # Store which search phrase was found

                    # Set the end line to the next empty line starting from this line
                    end_line = line_number + 1
                    while end_line < len(lines) and lines[end_line].strip() != '':
                        end_line += 1  # Continue until we find an empty line

                    # Initialize values to None
                    vc_ratio_value = '-'
                    los_value = None
                    delay_value = None
                    found_hcm = False  # Flag to track if we found HCM lines

                    # Process HCM lines
                    for search_line_number in range(line_number + 1, end_line):
                        line = lines[search_line_number]

                        # Check for 'v/c ratio' and extract the next float
                        if re.search(r'v/c ratio', line, re.IGNORECASE):
                            float_match = re.search(r'(\d+\.\d+|\d+)', line)
                            if float_match:
                                vc_ratio_value = float(float_match.group(0))

                        # Check for 'delay' and extract the next float
                        if re.search(r'delay', line, re.IGNORECASE):
                            # Find the position of 'delay' in the line
                            delay_pos = line.lower().find('delay')
                            if delay_pos != -1:  # If 'delay' is found
                                # Get everything after 'delay'
                                after_delay = line[delay_pos +
                                                   len('delay'):].strip()
                                # Search for float in the remaining substring
                                float_match = re.search(
                                    r'(\d+\.\d+|\d+)', after_delay)
                                if float_match:
                                    delay_value = float(float_match.group(0))

                        # Check for 'LOS' and extract the next capital letter (A-F)
                        if re.search(r'LOS', line, re.IGNORECASE):
                            capital_match = re.search(r'\b[A-F]\b', line)
                            if capital_match:
                                los_value = capital_match.group(0)

                        if re.search(r'capacity', line, re.IGNORECASE):
                            float_match = re.search(r'(\d+\.\d+|\d+)', line)
                            if float_match:
                                capacity_value = float(float_match.group(0))

                        # Mark the line as HCM only if it starts with "HCM"
                        if line.startswith("HCM"):
                            found_hcm = True

                    # Store HCM results only if we processed HCM lines
                    if found_hcm:
                        print(f"Found HCM block at line {line_number}")  # Confirm the HCM block is found
                        print(f"Current phrase_found: {phrase_found}")  # Print phrase_found value to check what is being evaluated
                    
                        # Check which search phrase was found and structure the result accordingly
                        if phrase_found == "Intersection Summary":
                            print(f"Processing 'Intersection Summary' at line {line_number}")  # Debugging output
                            hcm_results.append({
                                'ID': id_value,
                                'v/c ratio': vc_ratio_value if vc_ratio_value is not None else '-',
                                'los': los_value if los_value is not None else '-',
                                'delay': delay_value if delay_value is not None else '-',
                                'capacity': capacity_value if capacity_value is not None else '-',
                                'index': intersection_index
                            })
                            intersection_index += 1
                        elif phrase_found == "Minor Lane/Major Mvmt":
                            print(f"Processing 'Minor Lane/Major Mvmt' at line {line_number}")  # Debugging output
                            movement_results, merged_results = parse_minor_lane_mvmt(
                                lines, line_number, end_line)
                    
                            # Create a dictionary where keys are from the movement results
                            hcm_entry = {'ID': id_value}
                    
                            for i in range(len(movement_results)):
                                # Using the movement results as keys
                                hcm_entry[movement_results[i]] = merged_results[i]
                            twsc_results.append(hcm_entry)
                        
                        # Stop collecting on a blank line
                        if line.strip() == "":
                            print(f"Blank line encountered at line {line_number}")  # Debugging output
                            break
                        break  # Exit after processing the HCM block
                    
                    # Skip lines between the ID and the next search phrase
                    if found_phrase:
                        print(f"Found phrase at line {line_number}: {phrase_found}")  # Debugging output
                        break  # Stop looking at this block and move on to the next intersection

    # awsc_results = parse_lane_data_from_df()
    # Print for debugging
    print("\nSynchro Signalized Summary Results (Intersection Summary):",
          synchro_results)
    print("\nHCM Signalized Summary Results (Intersection Summary):", hcm_results)
    print("\nTWSC Summary Results (Minor Lane/...):", twsc_results)
    print("\nAWSC Summary Results (Lane):", awsc_results, '\n')

    return twsc_results, synchro_results, hcm_results, awsc_results
    

def parse_twsc_approach(df):
    """
        Parses the approach data for each direction of a TWSC intersection
        Returns a list of dictionaries, one for each TWSC intersection found in the dataframe
    """
    approach_data = []  # List to hold all parsed data
    intersection_id = None  # Store the ID of the intersection we are currently looking at

    # Iterate through each row in the DataFrame
    for index, row in df.iterrows():
        line = str(row[0]).strip()  # Consider column 1 as the line to process
        # print(f"\nProcessing line {index}: {line}")

        # Check if the row in column 1 contains an integer (this is the intersection ID)
        if line.isdigit():
            intersection_id = int(line)
            # print(f"Found Intersection ID {intersection_id} at line {index}")
            continue  # Move to the next row

        # Check if the line starts with "approach" and contains at least one direction
        if line.lower() == "approach":
            # print(f"Found 'Approach' at line {index}: {line}")

            # Check if any of the specified directions are present in the line after "approach"
            present_directions = {direction: direction in row.values for direction in [
                "EB", "WB", "NB", "SB", 'NE', 'NW', 'SE', 'SW']}
            # print(f"Present directions: {present_directions}")

            # If no directions are found after "approach", skip this line
            if not any(present_directions.values()):
                print("No valid directions found, skipping line.")
                continue  # Skip the line if no directions are present

            approach_dict = {
                "Intersection ID": intersection_id,
                "EB": {"Approach Delay": None, "Approach LOS": '-'},
                "WB": {"Approach Delay": None, "Approach LOS": '-'},
                "NB": {"Approach Delay": None, "Approach LOS": '-'},
                "SB": {"Approach Delay": None, "Approach LOS": '-'},
                "NE": {"Approach Delay": None, "Approach LOS": '-'},
                "NW": {"Approach Delay": None, "Approach LOS": '-'},
                "SE": {"Approach Delay": None, "Approach LOS": '-'},
                "SW": {"Approach Delay": None, "Approach LOS": '-'},
            }

            # pd.set_option('display.max_columns', 100)  # Show all columns

            # Step 1: Record the positions (columns) of directions
            direction_columns = {}
            for direction in ["EB", "WB", "NB", "SB", 'NE', 'NW', 'SE', 'SW']:
                if present_directions[direction]:
                    # Find the column where the direction was found
                    direction_columns[direction] = row[row ==
                                                       direction].index[0]
                    # print(f"Direction {direction} found") # in column {direction_columns[direction]}.")

            # Now check subsequent rows for "HCM Control Delay" and "HCM LOS"
            next_row_index = index + 1
            while next_row_index < len(df):
                next_row = df.iloc[next_row_index]  # Get the next row

                # Check for "HCM Control Delay"
                if "hcm control delay" in str(next_row.iloc[0]).lower():
                    # print(f"Found 'HCM Control Delay' at row {next_row_index}.")

                    # Assign the delay values from the columns where directions were found
                    for direction, col in direction_columns.items():
                        delay_value = next_row[col]
                        # Check for numeric values
                        if pd.notna(delay_value) and re.match(r'\b\d+\.\d+|\b\d+', str(delay_value)):
                            approach_dict[direction]["Approach Delay"] = delay_value
                            # print(f"Setting {direction} Approach Delay: {delay_value}")
                        else:
                            # Store '-' if no valid value
                            approach_dict[direction]["Approach Delay"] = '-'
                            # print(f"No valid delay value for {direction}, setting to '-'.")

                # Check for "HCM LOS"
                elif "hcm los" in str(next_row.iloc[0]).lower():
                    # print(f"Found 'HCM LOS' at row {next_row_index}.")

                    # Assign the LOS values (A-F) from the columns where directions were found
                    for direction, col in direction_columns.items():
                        los_value = str(next_row[col]).strip().upper()
                        # Check if the value is a valid LOS (A-F)
                        if los_value in 'ABCDEF' and los_value != '':
                            approach_dict[direction]["Approach LOS"] = los_value
                            # print(f"Setting {direction} Approach LOS: {los_value}")
                        else:
                            # Store '-' if no valid LOS value
                            approach_dict[direction]["Approach LOS"] = '-'
                            # print(f"No LOS value for {direction}, setting to '-'.")

                # Move to the next row
                next_row_index += 1

                # Exit the loop if an empty row is found
                if next_row.isna().all():  # Check if the row is entirely empty or NaN
                    break

            # Step 3: Remove directions with no valid data
            approach_dict = {
                k: v for k, v in approach_dict.items() if k == "Intersection ID" or (isinstance(v, dict) and (v["Approach Delay"] is not None or v["Approach LOS"] != '-'))
            }

            # If there's any valid data, add it to approach_data
            if approach_dict:
                approach_data.append(approach_dict)
                # print(f"Added approach data: {approach_dict}")

    # print(f"\nFinal approach data (TWSC Intersections):\n{approach_data}")
    return approach_data


def extract_data_to_csv(file_path, output_file):
    data = []  # To store the final data for CSV
    # skip_lines = 0  # Counter to track skipped lines
    intersection_count = 0
    intersection_ids = []
    collecting = False  # Flag to indicate if we're collecting data
    # Flag to track if we're collecting after "Minor Lane/Major Mvmt"
    collecting_minor_lane = False
    lane_groups = []

    """
        Parse and extract relevant intersection data from the text file
        into a Dataframe and generate a CSV file
    """
    with open(file_path, 'r') as file:
        for line in file:
            stripped_line = line.rstrip('\n')  # Remove the newline character

            # Step 1: Look for a line starting with a digit and a colon
            if re.match(r'^\d+:', stripped_line):
                # Extract the intersection count from the beginning of the line
                intersection_count = int(
                    re.match(r'^(\d+):', stripped_line).group(1))
                data.append([intersection_count])
                intersection_ids.append(intersection_count)

                collecting = True  # Start collecting data
                collecting_minor_lane = False
                continue  # Skip the current line

            # Step 2: Start collecting after finding "Lane Group", "Movement", or "Intersection"
            if collecting or collecting_minor_lane:
                # If the line is empty, stop collecting (except when we're in the middle of collecting after "Minor Lane/Major Mvmt")
                if stripped_line == "":
                    collecting = False
                    collecting_minor_lane = False
                    continue  # Skip the empty line

                # Split the line based on double tabs or multiple spaces
                new_row = re.split(
                    r'\t\t|\s{2}\t|\s\t|\t', stripped_line.strip())
                # Remove empty cells
                new_row = [cell for cell in new_row if cell]
                data.append(new_row)  # Append the new row to data

            # Look for "Lane Group", "Movement", or "Intersection" after finding the intersection line
            if not collecting and not collecting_minor_lane:
                if re.match(r'\s*Lane Group', stripped_line) or re.match(r'\s*Movement', stripped_line):
                    collecting = True  # Start collecting data when either is found
                    # Split the line based on double tabs or multiple spaces
                    new_row = re.split(
                        r'\t\t|\s{2}\t|\s\t|\t', stripped_line.strip())
                    # Remove empty cells after stripping
                    new_row = [cell.strip() for cell in new_row if cell]
                    if new_row:  # Check if new_row is not empty after cleaning
                        data.append(new_row)  # Append the new row to data
                        # Append cleaned values to lane_groups
                        lane_groups.append(new_row[1:])
                        # print(new_row[1:])  # Print the cleaned values
                    # print(f"Collecting data for {intersection_count} under {stripped_line.strip()}")
                # If we find "Intersection" before "Lane Group" or "Movement"
                elif re.match(r'^\s*Intersection', stripped_line):
                    collecting = True  # Start collecting, ignoring blank lines
                    # Split the line based on double tabs or multiple spaces
                    new_row = re.split(
                        r'\t\t|\s{2}\t|\s\t|\t', stripped_line.strip())
                    # Remove empty cells
                    new_row = [cell for cell in new_row if cell]
                    data.append(new_row)  # Append the new row to data
                    # print(f"Collecting data for {intersection_count} under Intersection")
                # If we find "Minor Lane/Major Mvmt", start collecting until the next blank line
                elif re.match(r'^\s*Minor Lane/Major Mvmt', stripped_line) or re.match(r'^Approach', stripped_line) or re.match(r'^Lane\s[2]*', stripped_line):
                    collecting_minor_lane = True  # Start collecting after "Minor Lane/Major Mvmt"
                    # Split the line based on double tabs or multiple spaces
                    new_row = re.split(
                        r'\t\t|\s{2}\t|\s\t|\t', stripped_line.strip())
                    # Remove empty cells
                    new_row = [cell for cell in new_row if cell]
                    data.append(new_row)  # Append the new row to data
                    # print(f"Started collecting for {intersection_count} after Minor Lane/Major Mvmt")

    # print(f"Intersection ID's stored (length = {len(intersection_ids)}): \n{intersection_ids}\n")

    # Step 5: Create a DataFrame and save to CSV
    df = pd.DataFrame(data)
    df.to_csv(output_file, index=False)
    df = df.map(lambda x: x.strip() if isinstance(x, str) else x)
    
    awsc_data = parse_awsc_data(df)
    combined_data = integrate_awsc_data(awsc_data, [])
    formatted_awsc_data = process_directions_awsc(combined_data)
    print(formatted_awsc_data)
    
    # Define the terms to search for
    terms_to_match = [
        "V/c ratio(x)",
        "LnGrp Delay(d), s/veh",
        "LnGrp LOS",
        "V/c ratio",
        "Control delay (s/veh)",
        "LOS",
        "Approach Delay (s/veh)",
        "Approach Delay, s/veh",
        "Approach LOS",
        # "Sign control"
    ]

    # Initialize an empty dictionary to store row indices
    signalized_int_data = []
    all_intersection_configs = []
    row_indices = {}
    group_config_data = {}

    current_id = None
    j = 0

    """
        Storing and grouping intersection data (signalized/unsignalized)
    """
    # Iterate through DataFrame rows
    for index, row in df.iterrows():
        # Check if the first column contains a new intersection ID that is in the intersection_ids list
        first_column_value = str(row[0]).strip()
        if first_column_value.isdigit():
            potential_id = int(first_column_value)
            if potential_id in intersection_ids:
                # If we have an existing intersection ID, save the collected data for it
                if current_id is not None and (group_config_data or row_indices):
                    # Check if row_indices has at least one term from terms_to_match
                    if any(term in row_indices.values() for term in terms_to_match):
                        intersection_data = {
                            "Intersection ID": current_id,
                            "Configurations": group_config_data,
                            **{term: line_num for line_num, term in row_indices.items()}
                        }
                        signalized_int_data.append(intersection_data)

                    # Reset data structures for the next intersection
                    row_indices = {}
                    group_config_data = {}

                # Set the new intersection ID
                current_id = potential_id
                continue  # Move to the next iteration as we've identified a new intersection

        # Detect rows containing "Lane Group" or "Movement"
        if "Lane Configurations" in row.values:
            # Create an empty dictionary to hold the configurations
            config_dict = {}

            # Iterate over lane_groups[j] and row values simultaneously, skipping empty values
            for i, key in enumerate(lane_groups[j]):
                if i + 1 < len(row):  # Ensure there's a corresponding value in the row
                    # Get and clean the value in the row
                    value = str(row[i + 1]).strip()
                    if value != 'None' and value != '':  # Only add the key-value pair if the value is non-empty
                        config_dict[key] = value

            # Append the config_dict to the group_config_data list if it contains data
            if config_dict:
                group_config_data = config_dict
                all_intersection_configs.append(config_dict)

            # Move to the next set of lane groups
            j += 1
        # Check for the presence of "LOS" first, as it is case-sensitive
        if "LOS" in row.values:
            row_indices[index] = "LOS"
            continue
        # For other terms, check in a case-insensitive manner
        for term in terms_to_match:
            if any(str(cell).lower() == term.lower() for cell in row if term.lower() != "los"):
                row_indices[index] = term
                break  # Exit the inner loop if a term is found

    # Final addition for the last intersection data
    if current_id is not None and (group_config_data or row_indices):
        # Ensure the dictionary contains at least one of the terms_to_match
        if any(term in row_indices.values() for term in terms_to_match):
            intersection_data = {
                "Intersection ID": current_id,
                "Configurations": group_config_data,
                **{term: line_num for line_num, term in row_indices.items()}
            }
            signalized_int_data.append(intersection_data)

    # print(f"All intersection configurations: {all_intersection_configs}", f"\nlength = {len(all_intersection_configs)}")
    # print("Signalized Intersection Data:")
    # for entry in signalized_int_data:
    #     print(entry, '\n')

    lane_configurations, raw_lane_configs = parse_lane_configs(
        all_intersection_configs, intersection_ids)
    # print('\nLane Configurations collected...', lane_configurations,
    #       f"\nlength = {len(lane_configurations)}")
    # print("\nRaw Lane Configurations read...", raw_lane_configs, f"\nlength = {len(raw_lane_configs)}")

    # Initialize an empty list to store the combined dictionaries
    combined_list = []

    """
        Iterate through the list of signalized intersections
    """
    for intersection in signalized_int_data:
        intersection_id = intersection.get("Intersection ID")
        combined_dict = {"Intersection ID": intersection_id}

        # Initialize approach data for all possible directions
        approach_data = {
            direction: {"Approach Delay": None, "Approach LOS": None}
            for direction in ['EB', 'WB', 'NB', 'SB', 'NE', 'NW', 'SE', 'SW']
        }

        # Access lane configurations for the current intersection
        lane_config = lane_configurations[int(intersection_id) - 1]
        raw_lane_config = raw_lane_configs[int(intersection_id) - 1]

        # Initialize dictionaries to store v/c ratio, LOS, and delay values
        directions = ['EB', 'WB', 'NB', 'SB', 'NE', 'NW', 'SE', 'SW']
        vc_ratio_values = {direction: [None, None, None]
                           for direction in directions}
        los_values = {direction: [None, None, None]
                      for direction in directions}
        delay_values = {direction: [None, None, None]
                        for direction in directions}

        # Flags to ensure each type of data is added once
        vc_ratio_added = False
        ln_grp_los_added = False
        los_added = False

        """
            Search and process search terms in the current intersection
        """
        for term, row_index in intersection.items():
            if term not in terms_to_match:
                continue  # Skip terms not in the list

            term_lower = term.lower()
            row_data = df.iloc[row_index].replace(
                "", "-").fillna("-").tolist()[1:]  # Exclude first column value

            # Process v/c ratio values
            if "v/c ratio(x)" == term_lower or "v/c ratio" == term_lower:
                for idx, direction in enumerate(directions):
                    start_index = idx * 3  # Adjust based on 3 values per direction
                    if start_index + 2 < len(row_data):
                        vc_ratio_values[direction] = row_data[start_index:start_index + 3]
                    else:
                        vc_ratio_values[direction] = ["-", "-", "-"]
                if not vc_ratio_added:
                    combined_dict[term] = [
                        value for value in row_data if value != '-']
                    vc_ratio_added = True

            # Process LOS values
            if "lngrp los" == term_lower or "los" == term_lower:
                for idx, direction in enumerate(directions):
                    start_index = idx * 3
                    if start_index + 2 < len(row_data):
                        los_values[direction] = row_data[start_index:start_index + 3]
                    else:
                        los_values[direction] = ["-", "-", "-"]
                if "lngrp los" == term_lower and not ln_grp_los_added:
                    combined_dict[term] = [
                        value for value in row_data if value != '-']
                    ln_grp_los_added = True
                elif "los" == term_lower and not los_added and not ln_grp_los_added:
                    combined_dict[term] = [
                        value for value in row_data if value != '-']
                    los_added = True

            # Process delay values
            if "control delay (s/veh)" == term_lower or "lngrp delay(d), s/veh" == term_lower:
                for idx, direction in enumerate(directions):
                    start_index = idx * 3
                    if start_index + 2 < len(row_data):
                        delay_values[direction] = row_data[start_index:start_index + 3]
                    else:
                        delay_values[direction] = ["-", "-", "-"]

            # Process approach delay
            if "approach delay" in term_lower:
                filtered_row_data = [
                    value for value in row_data if value != '-']
                for idx, direction in enumerate(directions):
                    if lane_config.get(direction) == '-':
                        filtered_row_data.insert(idx, '-')
                for idx, direction in enumerate(directions[:len(filtered_row_data)]):
                    approach_data[direction]["Approach Delay"] = filtered_row_data[idx]

            # Process approach LOS
            if term_lower == "approach los":
                filtered_row_data = [
                    value for value in row_data if value != '-']
                for idx, direction in enumerate(directions):
                    if lane_config.get(direction) == '-':
                        filtered_row_data.insert(idx, '-')
                for idx, direction in enumerate(directions[:len(filtered_row_data)]):
                    approach_data[direction]["Approach LOS"] = filtered_row_data[idx]

            # Add other terms to combined_dict
            if term_lower not in ["v/c ratio", "los", "v/c ratio(x)", "lngrp los"]:
                combined_dict[term] = [
                    value for value in row_data if value != '-']

        # Finalize combined_dict
        value_set = [vc_ratio_values, delay_values, los_values]

        """
            Preference algorithm
        """
        # Process each type of value set: vc_ratio, los, and delay
        for values, term in zip(value_set, list(combined_dict.keys())[1:]):
            for direction in directions:
                # Identify non-zero values based on the type of data
                if values == los_values:
                    # For LOS values, consider grades A-F (skipping '-' or None)
                    non_zero_values = [value if value not in [
                        "-", None] else "Z" for value in values[direction]]
                    max_non_zero_value = max(non_zero_values)
                else:
                    # For vc_ratio and delay values, treat them as floats
                    non_zero_values = [float(value) if value not in [
                        "-", "0", None] else float('inf') for value in values[direction]]
                    non_inf_values = [
                        v for v in non_zero_values if v != float('inf')]
                    min_non_zero_value = min(non_zero_values)
                    max_non_zero_value = max(
                        non_inf_values) if non_inf_values else float('-inf')

                # Determine if we need to update based on raw lane configuration
                needs_update = any(
                    value not in [
                        "-", "0", None] and lane_config in ["0", None]
                    for value, lane_config in zip(values[direction], raw_lane_config[direction])
                )

                # Skip updating if all configurations are valid
                if not needs_update:
                    # print(f"Skipping updates for {direction} in {values} as all configurations are valid.")
                    continue

                # Apply the update logic based on the type of data
                for i, (value, lane_config) in enumerate(zip(values[direction], raw_lane_config[direction])):
                    if value in ["-", "0", None] and lane_config not in ["0", None]:
                        if values == los_values:
                            # Higher letter for LOS
                            values[direction][i] = max_non_zero_value
                        else:
                            values[direction][i] = str(max_non_zero_value)

                # Replace the lowest non-zero value with '-'
                if min_non_zero_value != float('inf') and values != los_values:
                    min_index = non_zero_values.index(min_non_zero_value)
                    values[direction][min_index] = '-'

            # Filter out '-' values and add to combined_dict
            filtered_values = {direction: [value for value in values[direction]
                                           if value != '-' and value != 'Z'] for direction in directions}
            combined_dict[term] = [
                value for sublist in filtered_values.values() for value in sublist]

        # Merge approach data into combined_dict and append to final lists
        combined_dict.update(approach_data)
        combined_list.append(combined_dict)

        # print(f"\nCombined data for signalized Intersection {intersection_id}: \n{combined_dict}")

    # Remove empty lane configurations
    for direction in ['EB', 'WB', 'NB', 'SB', 'NE', 'NW', 'SE', 'SW']:
        for config in lane_configurations:
            if config.get(direction) == '-':
                del config[direction]

    twsc_overall, synchro_overall, hcm_overall, awsc_overall = parse_overall_data_v2(
        file_path, df)
    twsc_intersections = parse_twsc_approach(df)
    twsc_intersection_directions, original_twsc_directions, combined_mvmt_names = process_directions(
        twsc_overall, lane_configurations)

    # print(f"\ntwsc_overall:\n{twsc_overall}\n")
    combined_list.extend(twsc_intersections)
    combined_list.extend(formatted_awsc_data)
    
    # Create an empty DataFrame to hold all intersections' data
    final_df = pd.DataFrame()

    general_terms = {
        'v/c': ['V/c ratio', 'V/c ratio(x)', 'LnGrp v/c'],
        'delay': ['Control delay (s/veh)', 'LnGrp Delay(d), s/veh'],
        'los': ['LOS', 'LnGrp LOS']
    }

    # Sort combined_list by Intersection ID for ordered processing
    combined_list_sorted = sorted(
        combined_list, key=lambda x: int(x.get("Intersection ID", "ID")))

    # print(f"Combined list for '{file_path}': {combined_list_sorted}")

    # print(f"Combined list for '{file_path}':")
    # for entry in combined_list_sorted:
    #     print(entry, '\n')

    # Iterate over each item in the sorted combined_list
    # for idx, item in enumerate(combined_list_sorted, 1):
    #     print(f"Intersection #{idx} (ID: {item.get('Intersection ID')}):")

    #     # Iterate over the keys and values of each dictionary
    #     for key, value in item.items():
    #         # If the value is a list, print it in a readable format
    #         if isinstance(value, list):
    #             value_str = ', '.join(map(str, value))
    #             print(f"  {key}: [{value_str}]")
    #         else:
    #             print(f"  {key}: {value}")

    #     print("\n" + "-"*50)  # Add a separator line between intersections

    # Combine both lists and sort by the "index" key
    combined_overall_data = sorted(
        synchro_overall + hcm_overall + twsc_overall + awsc_overall, key=lambda x: x.get('index', 0))
    overall_idx = 0

    # print(f"\nOverall Data = \n{combined_overall_data}", '\n')

    # Process each intersection in the sorted list
    for data_dict in combined_list_sorted:
        intersection_id = data_dict.get("Intersection ID")

        # Control printing of Intersection ID only once per direction set
        intersection_id_printed = False

        # Find the matching lane configuration for this intersection
        lane_config = next((config for config in lane_configurations if config.get(
            "Intersection ID") == intersection_id), None)

        # Check if the intersection has TWSC data
        twsc_summary_result = next(
            (twsc for twsc in twsc_overall if twsc.get("ID") == str(intersection_id)), None)
        twsc_summary_directions = next((twsc_dir for twsc_dir in twsc_intersection_directions if twsc_dir.get(
            "ID") == str(intersection_id)), None)
        # original_twsc = next((og_dir for og_dir in original_twsc_directions if og_dir.get("ID") == str(intersection_id)), None)

        # Prefer TWSC summary if available
        if twsc_summary_result and twsc_summary_directions:
            lane_config = None
        # Skip intersections without lane config or TWSC summary
        if not lane_config and not twsc_summary_result:
            print(
                f"No lane configuration or TWSC summary found for Intersection ID: {intersection_id}")
            continue

        # print(data_dict, '\n')

        # Prepare data for the intersection's DataFrame
        intersection_data = []

        # Separate indexing for v/c, LOS, and Delay values
        j = 0

        # Retrieve all entries from combined_overall_data for this intersection
        overall_data = [
            item for item in combined_overall_data if item['ID'] == str(intersection_id)]
        # Default values for the overall data row (to be updated if data is found)
        overall_los = '-'
        overall_delay = '-'

        # Processing if lane configuration is available
        if lane_config:
            for direction, lanes in lane_config.items():
                # Skip the "Intersection ID" key in lane_config
                if direction == "Intersection ID":
                    continue

                # Retrieve approach delay and LOS for the current direction
                approach_delay = data_dict.get(
                    direction, {}).get("Approach Delay", '-')
                approach_los = data_dict.get(
                    direction, {}).get("Approach LOS", '-')

                if approach_delay is None:
                    approach_delay = '-'
                if approach_los is None:
                    approach_los = '-'

                # Loop through each lane in the direction
                for i, lane in enumerate(lanes):
                    # Print the Intersection ID only once at the start of the set
                    intersection_id_str = str(
                        intersection_id) if not intersection_id_printed else ''
                    intersection_id_printed = True
                    direction_value = direction if i == 0 else ''  # Only print direction once

                    # Get v/c, LOS, and Delay values based on general terms dictionary
                    vc_value = los_value = delay_value = '-'

                    # Check and get v/c value from general terms
                    for term in general_terms['v/c']:
                        if term in data_dict:
                            vc_value = data_dict[term][j] if j < len(
                                data_dict[term]) else '-'
                            break

                    # Check and get LOS value from general terms
                    for term in general_terms['los']:
                        if term in data_dict:
                            los_value = data_dict[term][j] if j < len(
                                data_dict[term]) else '-'
                            break

                    # Check and get Delay value from general terms
                    for term in general_terms['delay']:
                        if term in data_dict:
                            delay_value = data_dict[term][j] if j < len(
                                data_dict[term]) else '-'
                            break

                    if vc_value and los_value and delay_value != '-':
                        # Append the row for this lane
                        intersection_data.append(
                            [intersection_id_str, direction_value, lane, vc_value, los_value, delay_value])

                    # Increment the indices for v/c, LOS, and Delay values
                    j += 1

                if approach_delay != '-':
                    # Add an overall row for this direction
                    intersection_data.append(
                        ['', f"{direction} Overall", '', '-', f'{approach_los}', f'{approach_delay}'])

            if overall_idx == 0:
                overall_los = overall_data[0].get("los")
                overall_delay = overall_data[0].get("delay")
                overall_idx += 1
            else:
                try:
                    # Attempt to access the second element in the overall_data list
                    overall_los = overall_data[1].get("los")
                    overall_delay = overall_data[1].get("delay")
                    overall_idx = 1  # Indicate that the second item was used
                except IndexError:
                    # If the second element does not exist, fall back to the first element
                    if overall_data:
                        overall_los = overall_data[0].get("los")
                        overall_delay = overall_data[0].get("delay")
                        overall_idx = 0  # Indicate that the first item was used
                    else:
                        # If overall_data is empty, set default values
                        overall_los = '-'
                        overall_delay = '-'
                        overall_idx = None  # No data available
            if overall_delay != 0 or '-':
                # Add an overall row for this intersection, including data from Synchro and HCM
                intersection_data.append(
                    ['', "Overall", '', '-', overall_los, overall_delay])

        # Processing if TWSC summary data is available
        if twsc_summary_result:
            # print(f"\nIntersection {intersection_id}")
            # print("\nDirections: ", twsc_summary_directions, '\n')
            # print("Values:", twsc_summary_result)

            # Iterate through TWSC summary directions
            for direction, movement_values in twsc_summary_result.items():
                # Skip the "ID" key in TWSC summary
                if direction == "ID":
                    continue

                # Find the lane configuration in the TWSC summary for this direction
                if direction in twsc_summary_result:
                    if movement_values[3] == '-':
                        continue
                    lane_data = twsc_summary_result[direction]
                else:
                    # Default or placeholder value
                    lane_data = ('-', '-', '-', '-')

                # Unpack v/c, LOS, and Delay values from TWSC data
                vc_value, los_value, delay_value, capacity_value = (
                    lane_data if isinstance(
                        lane_data, tuple) else ('-', '-', '-', '-')
                )

                # Add an entry for the TWSC summary direction
                intersection_id_str = str(
                    intersection_id) if not intersection_id_printed else ''
                direction_value = twsc_summary_directions[direction[:2]]

                # Append the row for this direction (from TWSC summary)
                if isinstance(direction_value, list):
                    for dir_val in direction_value:
                        key = direction[:2] + dir_val
                        lane_data = twsc_summary_result[key]

                        # For each lane, append a row with the lane and corresponding values
                        if lane_data != ('-', '-', '-', '-'):
                            intersection_data.append([
                                # Intersection ID (if not printed)
                                intersection_id_str,
                                # Direction value (e.g., 'SW')
                                direction[:2],
                                # Lane value (e.g., 'L', 'T')
                                dir_val,
                                vc_value,              # V/c value
                                los_value,             # LOS value
                                delay_value            # Delay value
                            ])
                else:
                    # If not a list, process the single direction normally
                    intersection_data.append([
                        # Intersection ID (if not printed)
                        intersection_id_str,
                        direction[:2],         # Direction value (e.g., 'WB')
                        # Lane data (since it's not a list here)
                        direction_value,
                        vc_value,              # V/c value
                        los_value,             # LOS value
                        delay_value            # Delay value
                    ])
                intersection_id_printed = True
                # print(intersection_data)

        if intersection_data != []:
            # Add a blank row to separate intersections
            intersection_data.append([''] * 6)

            # Create a DataFrame for the current intersection's data
            intersection_df = pd.DataFrame(intersection_data, columns=[
                                           'Intersection ID', 'Direction', 'Lane', 'V/c', 'LOS', 'Delay'])

            # Append it to the final DataFrame
            final_df = pd.concat(
                [final_df, intersection_df], ignore_index=True)

    # Write the final DataFrame to a CSV file
    file_name, _ = os.path.splitext(file_path)
    final_df.to_csv(f"{file_name}-filtered.csv", index=False)

    """
        Output finalized data for debugging and testing
    """
    i = 0
    # Initialize the intersection ID from id_combined_list
    # for item in combined_list_sorted:

    #     # Determine the intersection ID and the data dictionary based on whether the item is a tuple or dictionary

    #     intersection_id = item.get("Intersection ID")
    #     data_dict = item

    #     # Print each term and its data in a readable format, excluding direction data (EB, WB, NB, SB)
    #     print(f"Intersection {intersection_id}:")
    #     for term, data in data_dict.items():
    #         if term not in ['EB', 'WB', 'NB', 'SB', 'NE', 'NW', 'SE', 'SW']:  # Only print non-directional data here
    #             # Join data if it's a list, otherwise convert it to a string
    #             if isinstance(data, list):
    #                 data_str = ", ".join(map(str, data))
    #             else:
    #                 data_str = str(data)
    #             print(f"  {term}: {data_str}")
    #     print()
    #     # Print Approach Delay and LOS for each direction (EB, WB, NB, SB)
    #     for direction in ['EB', 'WB', 'NB', 'SB', 'NE', 'NW', 'SE', 'SW']:
    #         # Retrieve approach delay and LOS for the current direction
    #         # print(f"Getting approach data for '{direction}'...")
    #         # print(data_dict.get(direction, {}))
    #         approach_delay = data_dict.get(direction, {}).get("Approach Delay", '-')
    #         approach_los = data_dict.get(direction, {}).get("Approach LOS", '-')
    #         # Only delete if the direction exists in data_dict and conditions are met
    #         print(f"  {direction}: Approach Delay = {approach_delay}, Approach LOS = {approach_los}")
    #     print()
    #     # Find the matching lane configuration for this intersection ID in group_config_data
    #     lane_config = next((config for config in lane_configurations if config.get("Intersection ID") == intersection_id), None)
    #     raw_config = next((raw for raw in raw_lane_configs if raw.get("Intersection ID") == intersection_id), None)

    #     # Print lane configurations for the current intersection if available
    #     if lane_config:
    #         lane_config_str = ", ".join(
    #             f"{key}: {value}" for key, value in lane_config.items() if key != "Intersection ID" and value != '-'
    #         )
    #         print(f"  Lane Configurations: {lane_config_str}")
    #     else:
    #         print(f"  No lane configurations found for Intersection ID: {intersection_id}")

    #     # Print raw direction configurations for the current intersection if available
    #     if raw_config:
    #         raw_config_str = ", ".join(
    #             f"{key}: {value}" for key, value in raw_config.items() if key != "Intersection ID" and value != [None, None, None]
    #         )
    #         print(f"  Raw Direction Configurations: {raw_config_str}")
    #     else:
    #         print(f"  No raw direction configurations found for Intersection ID: {intersection_id}")

    #     i+=1
    #     # Add a blank line for readability between intersections
    #     print("\n" + "_" * 40 + "\n")

    # print(f"Total number of useable datasets found: {len(combined_list_sorted)}")
    # print("_" * 40 + "\n")


def parse_lane_configs(int_lane_groups, intersection_ids):
    parsed_list = []  # This will store the parsed dictionaries for each group
    raw_data_list = []

    for idx, lane_dict in enumerate(int_lane_groups):

        intersection_id = intersection_ids[idx]

        # Skip if the intersection ID is already in parsed_list
        if any(parsed_dict.get("Intersection ID") == intersection_id for parsed_dict in parsed_list):
            continue

        parsed_dict = {
            "Intersection ID": intersection_id,
            # Initialize with three None values for L, T, R
            'EB': [None, None, None],
            'WB': [None, None, None],
            'NB': [None, None, None],
            'SB': [None, None, None],
            'NE': [None, None, None],
            'NW': [None, None, None],
            'SE': [None, None, None],
            'SW': [None, None, None]
        }

        # Initialize the raw data dictionary
        raw_data_dict = {
            "Intersection ID": intersection_id,
            'EB': [None, None, None],
            'WB': [None, None, None],
            'NB': [None, None, None],
            'SB': [None, None, None],
            'NE': [None, None, None],
            'NW': [None, None, None],
            'SE': [None, None, None],
            'SW': [None, None, None]
        }

        for direction, value in lane_dict.items():
            if value is None or value == '':
                value = '-'
                continue

            # Process each direction and suffix (L, T, R)
            suffixes = {
                'L': 0,  # Index 0 for Left
                'T': 1,  # Index 1 for Through
                'R': 2   # Index 2 for Right
            }

            for suffix, idx in suffixes.items():
                # Parse the value for numbers and special characters < and >
                if direction.endswith(suffix):
                    # Store the raw value directly in raw_data_dict in the correct position
                    direction_prefix = direction[:-1]
                    if direction_prefix in raw_data_dict:
                        # Store unparsed raw value
                        raw_data_dict[direction_prefix][idx] = value

                    parsed_value = ''

                    if '<' in value:
                        parsed_value += 'L'  # Leading left if < is present
                    number_part = ''.join(
                        filter(str.isdigit, value))  # Get the number
                    if number_part:
                        # Repeat based on the number
                        parsed_value += suffix * int(number_part)

                    else:
                        parsed_value += suffix

                    if '>' in value:
                        parsed_value += 'R'  # Trailing right if > is present

                    # Store the parsed value in the respective direction and suffix position
                    # Get the prefix like EB, WB, etc.
                    direction_prefix = direction[:-1]
                    if direction_prefix in parsed_dict:
                        parsed_dict[direction_prefix][idx] = parsed_value or None

        # Remove None values from each list in the parsed_dict
        for key in list(parsed_dict.keys()):
            if key != "Intersection ID":  # Don't touch the Intersection ID key
                parsed_dict[key] = [
                    value for value in parsed_dict[key] if value is not None]
                # If the list is empty (no valid values), set it to '-'
                if not parsed_dict[key]:
                    parsed_dict[key] = '-'

        # Clean up the raw_data_dict in the same way
        for key in raw_data_dict:
            if key != "Intersection ID":
                raw_data_dict[key] = [value for value in raw_data_dict[key]]
                if not raw_data_dict[key]:
                    raw_data_dict[key] = '-'

        # Debugging output
        # print(f"\nParsed Lane Config (Intersection #{intersection_id}):\n{parsed_dict} \nRaw Lane Config (Intersection #{intersection_id}):\n{raw_data_dict}")

        # Append the parsed_dict for this lane group to the final list
        parsed_list.append(parsed_dict)
        raw_data_list.append(raw_data_dict)

    return parsed_list, raw_data_list


if __name__ == "__main__":
    # read_input_file("test-input.xlsx")
    test_report_1 = "test/Test Report 1.txt"
    test_report_2 = "test/Test Report 2.txt"
    test_report_3 = "test/Test Report 3.txt"
    test_report_4 = "test/Test Report 4.txt"
    test_twsc = "test/TEST TWSC.txt"
    test_awsc = "test/TEST AWSC.txt"

    test_report_1_csv = "test-report-1.csv"
    test_report_2_csv = "test-report-2.csv"
    test_report_3_csv = "test-report-3.csv"
    test_report_4_csv = "test-report-4.csv"
    test_twsc_csv = "test-twsc.csv"
    test_awsc_csv = "test-awsc.csv"

    # parse_overall_data_v2(file)  # Gets the data for overall

    # Testing with Test Report 1.txt
    print('\n' + "*"*35 + "\n| Results for 'Test Report 1.txt' |\n" + "*"*35 + '\n')
    extract_data_to_csv(test_report_1, test_report_1_csv)

    # Testing with Test Report 2.
    print('\n' + "*"*35 + "\n| Results for 'Test Report 2.txt' |\n" + "*"*35 + '\n')
    extract_data_to_csv(test_report_2, test_report_2_csv)

    # Testing with Test Report 3.txt
    print('\n' + "*"*35 + "\n| Results for 'Test Report 3.txt' |\n" + "*"*35 + '\n')
    extract_data_to_csv(test_report_3, test_report_3_csv)

    # Testing with Test Report 3.txt
    print('\n' + "*"*35 + "\n| Results for 'Test Report 4.txt' |\n" + "*"*35 + '\n')
    extract_data_to_csv(test_report_4, test_report_4_csv)

    # print('\n' + "*"*35 + "\n| Results for 'TEST TWSC.txt' |\n" + "*"*35 +'\n')
    # extract_data_to_csv(test_twsc, test_twsc_csv)

    print('\n' + "*"*35 + "\n| Results for 'TEST AWSC.txt' |\n" + "*"*35 + '\n')
    extract_data_to_csv(test_awsc, test_awsc_csv)

    # lane_groups = separate_characters(movement)
    # print(f"\nLane groups:\n{lane_groups}")
    # write_to_excel(file, movement, delay, vc, los)
