# Calm down user
print("Loading...")

import pandas as pd
import tkinter as tk
from tkinter import filedialog
import yaml
from itertools import zip_longest
import os
import threading

# Put two strings side by side with | as a divider
def _combine_multiline_strings(str1, str2, separator='  |  '):
    lines1 = str1.split('\n')
    lines2 = str2.split('\n')
    max_length = max(len(line) for line in lines1)
    combined_lines = []
    for line1, line2 in zip_longest(lines1, lines2, fillvalue=' '):
        padded_line1 = line1.ljust(max_length)
        combined_lines.append(f"{padded_line1}{separator}{line2}")
    return '\n'.join(combined_lines)

# Only accepts integer input from list "options"
def _get_valid_input(prompt, options=[]):
    while True:
        try:
            user_input = input(prompt)
            if int(user_input) in options:
                return int(user_input)
            print("That's not a valid option. Please try again.")
        except ValueError:
            print("That's not a valid option. Please try again.")

# Converts py dictionaries to yaml string
def _dict_to_string(dict_obj):
    try:
        yaml_string = yaml.dump(data=dict_obj, default_flow_style=False, indent=4)

        # Replaces null values with true whitespace for user clarity
        new_lines = []
        for line in yaml_string.splitlines():
            if line.endswith("null"):
                line = line[:-4]
            new_lines.append(line)
        new_yaml_string = "\n".join(new_lines)
        return new_yaml_string
    except TypeError as e:
        print(f"Error: This object type cannot be converted to YAML: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

# Converts Numbers to Alphabet like Excel Columns
def _number_to_letters(num):
    letters = ''
    while num > 0:
        num, remainder = divmod(num - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters

# Recursively removes empty dictionaries
def _collapse_empty_dict(dictionary):
    keys_to_delete = []
    for key, val in dictionary.items():
        if isinstance(val, dict):
            _collapse_empty_dict(val)
            if not val:
                keys_to_delete.append(key)
    for key in keys_to_delete:
        del dictionary[key]
    return dictionary

# Return just the filename of a path
def _get_base_file_name(filename):
    filename_only = os.path.basename(filename)
    return filename_only

def _save_file(save_path, contents):
    if save_path:
        with open(save_path, 'w') as file:
            file.write(contents)
        print(f"Discrepancies saved to {save_path}")
    else:
        print("Save operation was cancelled.")

# Opens GUI to load a save path
def save_dialog():
    default_ext = ".txt"
    filetypes = [
        ("Text file", "*.txt")
    ]
    default_name = "Discrepancies_Report.txt"
    root = tk.Tk()
    root.attributes('-topmost', True)
    root.update()
    root.withdraw()
    file_path = filedialog.asksaveasfilename(
        title="Select Location to Save Discrepancies Report",
        filetypes=filetypes,
        defaultextension=default_ext,
        initialfile=default_name
    )
    root.destroy()
    if file_path:
        return file_path
    print("No save file chosen...")
    exit()

# Opens GUI to load a file path
def read_dialog(num):
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title=f"Choose file {num} to compare", filetypes=[("Excel files", "*.xlsx")])
    root.destroy()
    if file_path:
        return file_path
    print("No file selected...")
    exit()

# Main function to collect dictionary of discrepancies
def compare_excel_files(file1, file2):
    global file1_name
    global file2_name

    # *Elevator Music*
    print("Working...")

    # Initializing...
    xls1 = pd.ExcelFile(file1)
    xls2 = pd.ExcelFile(file2)
    file1_name = _get_base_file_name(file1)
    file2_name = _get_base_file_name(file2)
    file1_discrepancies = {file1_name: {}}
    file2_discrepancies = {file2_name: {}}

    # Helper function to compare a sheet of both files
    def compare_sheet(sheet, file1, file2, file1_discrepancies, file2_discrepancies):
        df1 = pd.read_excel(file1, sheet_name=sheet, header=None)
        df2 = pd.read_excel(file2, sheet_name=sheet, header=None)

        # Normalize shapes by filling non-existing cells with NaN
        max_rows = max(df1.shape[0], df2.shape[0])
        max_cols = max(df1.shape[1], df2.shape[1])
        df1 = df1.reindex(index=range(max_rows), columns=range(max_cols))
        df2 = df2.reindex(index=range(max_rows), columns=range(max_cols))

        file1_discrepancies[file1_name][sheet] = {}  # Initialize sheet entry
        file2_discrepancies[file2_name][sheet] = {}  # Initialize sheet entry

        # Iterate through every cell
        for row in range(max_rows):
            for col in range(max_cols):
                # Get cell values, treating NaN as a special string for comparison
                value1 = df1.iat[row, col] if not pd.isnull(df1.iat[row, col]) else None
                value2 = df2.iat[row, col] if not pd.isnull(df2.iat[row, col]) else None

                # Compare cell values
                if value1 != value2:
                    file1_discrepancy = {
                        f"Cell {_number_to_letters(col+1)}{row+1}": value1
                    }
                    file2_discrepancy = {
                        f"Cell {_number_to_letters(col+1)}{row+1}": value2
                    }

                    # Merge existing dictionary with new discrepancies
                    file1_discrepancies[file1_name][sheet].update(file1_discrepancy)
                    file2_discrepancies[file2_name][sheet].update(file2_discrepancy)

    # Multithreading!!!
    threads = []
    sheets = []

    # Create list of sheets
    for sheet in xls1.sheet_names:
        if sheet in xls2.sheet_names:
            sheets.append(sheet)

    # Create threads
    for index, sheet in enumerate(sheets):
        t = threading.Thread(target=compare_sheet, args=(sheet, file1, file2, file1_discrepancies, file2_discrepancies))
        print(f"Sheet {index + 1} of {len(sheets)} started")
        t.start()
        threads.append(t)

    # Wait for all threads to finish
    print("Processing...")
    for index, t in enumerate(threads):
        t.join()
        print(f"Sheet {index + 1} of {len(sheets)} finished")

    return file1_discrepancies, file2_discrepancies

# Main Loop
def main():

    # Get file paths
    file1_path = read_dialog(1)
    file2_path = read_dialog(2)

    # Compare the files
    file1_discrepancies, file2_discrepancies = compare_excel_files(file1_path, file2_path)
    print(f"Finalizing...")

    # Collapse entry dicts
    file1_discrepancies_collapsed = _collapse_empty_dict(file1_discrepancies)
    file2_discrepancies_collapsed = _collapse_empty_dict(file2_discrepancies)

    # Convert dicts to strings (yaml ftw)
    file1_discrepancies_string = _dict_to_string(file1_discrepancies_collapsed)
    print(f"File 1 of 2 finished")
    file2_discrepancies_string = _dict_to_string(file2_discrepancies_collapsed)
    print(f"File 2 of 2 finished")

    # Zip strings
    zipped_string = _combine_multiline_strings(file1_discrepancies_string, file2_discrepancies_string)

    # Choose to print or save discrepancies
    print("\n" + "-" * 50 + "\n")
    print("Display or save results?\n\n1. Display\n2. Save\n")
    choice = _get_valid_input("Enter number: ", [1, 2])

    # If choice is to display
    if choice == 1:
        print("\n" + zipped_string)
        print("\n" + "-" * 50 + "\n")
        print("Save file?\n\n1. Yes\n2. No\n")
        choice = _get_valid_input("Enter number: ", [1, 2])

        # Double check to save before ending
        if choice == 1:
            print("Opening file window... (It may be hidden)")
            save_path = save_dialog()
            _save_file(save_path, zipped_string)

    # If choice is to save
    elif choice == 2:
        print("Opening file window... (It may be hidden)")
        save_path = save_dialog()
        _save_file(save_path, zipped_string)

    # Exit script
    input("You may safely close the window...")
    exit()

if __name__ == "__main__":
    main()