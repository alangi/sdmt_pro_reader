import os
import sys
import logging
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter import scrolledtext as st
from xml.etree import ElementTree as ET
import pandas as pd
import numpy as np

from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font

# create the root window
root = tk.Tk()
root.title('Tkinter Open File Dialog')
root.resizable(False, False)
root.geometry('300x300')

# Make the main window appear at the front of the screen when it initializes
root.lift()
root.attributes('-topmost', True)
root.after_idle(root.attributes, '-topmost', False)


def select_file(callback):
    """
    Opens a file dialog and calls the provided callback function
    with the selected filename.
    """
    filetypes = (
        ('SDMT files', '*.sdmt'),
        ('All files', '*.*')
    )

    filename = fd.askopenfilename(
        title='Open a file',
        initialdir='U:\1 - Projects\2024',
        filetypes=filetypes)

    # Call the provided callback function with the filename if selected
    if filename:
        callback(filename)


def process_filename(filename):
    global dataframes, export_directory, transformed_dataframes
    global POINT_df, DMTG_df, DMTT_df, Seismic_Downhole_df, Vs_Data_df
    global Vs_Parameters_Setup_df, Vs_Parameters_Data_df

    directory = os.path.dirname(filename)
    export_directory = directory
    sdmt_report_filename = 'SDMT Report.xlsx'

    # Call function to read .sdmt file and return tests
    sdmt_tests = read_sdmt_xml(filename)

    # Call function to make tuple of test filepaths
    test_filepaths = get_test_filepaths(directory, sdmt_tests)

    # Call function to read TEST files into DataFrame
    test_meta_df = process_test_files(test_filepaths)

    # Call function to make tuple of dmt filepaths
    dmt_filepaths = get_dmt_filepaths(directory, sdmt_tests)

    # Call function to read DMT files into DataFrames
    dmt_setup_df, dati_dmt_df = process_dmt_files(dmt_filepaths)

    # Call function to read SDMT Report .xlsx file and return results
    s_wave_data_df, p_wave_data_df, dmt_data_df = parse_sdmt_report_xlsx(
        directory, sdmt_report_filename)

    # Call function to read VS files into DataFrame
    vs_setup_df = process_vs_files(directory, sdmt_tests)

    # Call function to process .sdm files based on the directory
    vs_sdm_info_df = process_sdm_files(directory, sdmt_tests)

    if vs_sdm_info_df.empty:
        vs_sdm_info_df = pd.DataFrame(columns=['POINT_ID', 'SdmNumber'])

    # Combine all dataframes into a list for transformation
    dataframes = [
        test_meta_df, dmt_setup_df, dati_dmt_df, dmt_data_df, s_wave_data_df,
        p_wave_data_df, vs_setup_df, vs_sdm_info_df
    ]

    # Transform raw dataframes
    transformed_dataframes = transform_raw_data(dataframes)

    # Extract the transformed dataframes back from the list
    (test_meta_df, dmt_setup_df, dati_dmt_df, dmt_data_df, s_wave_data_df,
     p_wave_data_df, vs_setup_df, vs_sdm_info_df) = transformed_dataframes

    dataframes = [
        test_meta_df, dmt_setup_df, dati_dmt_df, dmt_data_df, s_wave_data_df,
        p_wave_data_df, vs_setup_df, vs_sdm_info_df
    ]

    # Create POINT_df from test_meta_df
    POINT_df = create_POINT_df(test_meta_df)

    # Create DMTG_df from test_meta_df
    DMTG_df = create_DMTG_df(test_meta_df)

    # Create DMTT_df from dmt_data_df
    DMTT_df = create_DMTT_df(dmt_data_df, test_meta_df)

    # Create Seismic_Downhole_df from test_meta_df
    Seismic_Downhole_df = create_Seismic_Downhole_df(test_meta_df)

    # Create Vs_Data_df from s_wave_data_df
    Vs_Data_df = create_Vs_Data_df(s_wave_data_df, test_meta_df)

    # Create Vs_Parameters_Setup_df from test_meta_df
    Vs_Parameters_Setup_df = create_Vs_Parameters_Setup_df(test_meta_df)

    # Create Vs_Parameters_Data_df from s_wave_data_df
    Vs_Parameters_Data_df = create_Vs_Parameters_Data_df(
        s_wave_data_df, test_meta_df)


def show_combined_scrollable_info(title, **dataframes):
    """
    Creates a scrollable info box to display multiple dataframes one
    after the other with titles referencing the dataframe names.

    Args:
        title (str): The title of the info box.
        dataframes (dict): The dataframes to be displayed,
        with their names as keys.
    """
    info_window = tk.Toplevel(root)
    info_window.title(title)
    info_window.geometry('600x800')

    text_area = st.ScrolledText(info_window, wrap=tk.WORD, width=80, height=40)

    # Temporarily set pandas display options
    with pd.option_context('display.max_rows', None, 'display.max_columns',
                           None, 'display.max_colwidth', None):
        for df_name, df in dataframes.items():
            text_area.insert(tk.INSERT, f"{df_name}:\n")
            text_area.insert(tk.INSERT, df.to_string())
            text_area.insert(tk.INSERT, "\n\n")

    text_area.configure(state='disabled')
    text_area.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)


def export_dataframes_to_excel(directory, **dataframes):
    """
    Export given dataframes to individual worksheets in an Excel workbook.

    Args:
        directory (str): The directory where the Excel workbook will be saved.
        dataframes (dict): The dataframes to be exported,
        with their names as keys.
    """
    file_path = os.path.join(directory, 'sdmt_database_import.xlsx')

    # Create a workbook and add a worksheet for each dataframe
    wb = Workbook()
    for df_name, df in dataframes.items():
        # Format the sheet name by removing underscores and 'df'
        sheet_name = df_name.replace('_df', '')
        ws = wb.create_sheet(title=sheet_name)
        for r_idx, row in enumerate(dataframe_to_rows(df, index=False,
                                                      header=True)):
            for c_idx, value in enumerate(row, 1):
                cell = ws.cell(row=r_idx + 1, column=c_idx, value=value)
                if r_idx == 0:  # Header row
                    cell.font = Font(bold=True)
                # Auto-adjust column width
                column_letter = ws.cell(row=1, column=c_idx).column_letter
                ws.column_dimensions[column_letter].width = max(
                    len(str(value)) + 2,
                    ws.column_dimensions[column_letter].width)

    # Remove the default created sheet
    del wb['Sheet']

    # Save the workbook
    wb.save(file_path)
    print(f"Data successfully exported to {file_path}")
    sys.stdout.flush()  # Force the print buffer to flush


def database_export_button_callback():
    global export_directory, POINT_df, DMTG_df, DMTT_df, Seismic_Downhole_df
    global Vs_Data_df, Vs_Parameters_Setup_df, Vs_Parameters_Data_df
    export_dataframes_to_excel(
        export_directory,
        POINT_df=POINT_df,
        DMTG_df=DMTG_df,
        DMTT_df=DMTT_df,
        Seismic_Downhole_df=Seismic_Downhole_df,
        Vs_Data_df=Vs_Data_df,
        Vs_Parameters_Setup_df=Vs_Parameters_Setup_df,
        Vs_Parameters_Data_df=Vs_Parameters_Data_df
    )


def export_button_callback():
    global export_directory, transformed_dataframes
    export_to_excel(export_directory, transformed_dataframes)


# open button
open_button = ttk.Button(
    root,
    text='Open a File',
    command=lambda: select_file(process_filename))

open_button.pack(expand=True)

# new database export button
database_export_button = ttk.Button(
    root,
    text='Database Export',
    command=database_export_button_callback)

database_export_button.pack(expand=True)

# export button
export_button = ttk.Button(
    root,
    text='Export',
    command=export_button_callback)

export_button.pack(expand=True)


def show_scrollable_info(title, *dataframes):
    """
    Creates a scrollable info box to display the given dataframes.

    Args:
        title (str): The title of the info box.
        dataframes (tuple): The dataframes to be displayed in the info box.
    """
    info_window = tk.Toplevel(root)
    info_window.title(title)
    info_window.geometry('600x400')

    text_area = st.ScrolledText(info_window, wrap=tk.WORD, width=80,
                                height=20)

    # Temporarily set pandas display options
    with pd.option_context('display.max_rows', None, 'display.max_columns',
                           None, 'display.max_colwidth', None):
        for df in dataframes:
            text_area.insert(tk.INSERT, df.to_string())
            text_area.insert(tk.INSERT, "\n\n")

    text_area.configure(state='disabled')
    text_area.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)


# Function to show dataframes scrollable info box
def show_dataframes_info():
    global POINT_df, DMTG_df, DMTT_df, Seismic_Downhole_df, Vs_Data_df
    global Vs_Parameters_Setup_df, Vs_Parameters_Data_df
    show_combined_scrollable_info(
        'DataFrames',
        POINT_df=POINT_df,
        DMTG_df=DMTG_df,
        DMTT_df=DMTT_df,
        Seismic_Downhole_df=Seismic_Downhole_df,
        Vs_Data_df=Vs_Data_df,
        Vs_Parameters_Setup_df=Vs_Parameters_Setup_df,
        Vs_Parameters_Data_df=Vs_Parameters_Data_df
    )


# Function to show transformed dataframes scrollable info box
def show_transformed_dataframes_info():
    global transformed_dataframes
    (test_meta_df, dmt_setup_df, dati_dmt_df, dmt_data_df, s_wave_data_df,
     p_wave_data_df, vs_setup_df, vs_sdm_info_df) = transformed_dataframes

    show_scrollable_info(
        'Transformed Dataframes', test_meta_df, dmt_setup_df, dati_dmt_df,
        dmt_data_df, s_wave_data_df, p_wave_data_df, vs_setup_df,
        vs_sdm_info_df
    )


def read_sdmt_xml(filename):
    """
    Reads the XML content from a file and returns a tuple of test folder names.

    Args:
        filename (str): The path to the XML file.

    Returns:
        tuple: A tuple containing the test folder names from the XML.
    """
    try:
        with open(filename, 'r') as f:
            xml_string = f.read()  # Read the entire XML content into a string
    except FileNotFoundError:
        print(f"Error: File not found: {filename}")
        return ()  # Return empty tuple on file not found

    return get_tests_from_xml(xml_string)  # Call the original function


def get_tests_from_xml(xml_string):
    """
    Parses the provided XML string and returns a tuple of test folder names.

    Args:
        xml_string (str): The XML content as a string.

    Returns:
        tuple: A tuple containing the test folder names from the XML.
    """
    root = ET.fromstring(xml_string)  # Parse the XML string
    tests_element = root.find('Tests')  # Find the "Tests" element

    if tests_element is None:
        return ()  # Return empty tuple if no "Tests" element found

    # Extract test folder names
    tests = tuple(test.get('Folder') for test in tests_element.findall('Test'))

    return tests


def get_test_filepaths(directory, sdmt_tests):
    """
    Generates a tuple of test file paths and subdirectory names
    based on test names and directory information.

    Args:
        sdmt_tests (tuple): A tuple containing test names.
        directory (str, optional): The directory where test files are located.

    Returns:
        tuple: A tuple of tuples, each containing the subdirectory name
               and file path for each test name.
    """
    test_filepaths = tuple((test_name, f"{directory}/{test_name}/TEST.dat")
                           for test_name in sdmt_tests)
    return test_filepaths


def get_dmt_filepaths(directory, sdmt_tests):
    """
    Generates a tuple of test file paths based on test names
    and directory information.

    Args:
        sdmt_tests (tuple): A tuple containing test names.
        directory (str, optional): The directory where test files are located.

    Returns:
        tuple: A tuple containing tuples with subdirectory names
        and file paths for each test name.
    """
    dmt_filepaths = tuple((test_name, f"{directory}/{test_name}/DMT.dat")
                          for test_name in sdmt_tests)
    return dmt_filepaths


def process_test_files(test_filepaths):
    """
    Loops through a tuple of file paths, parses each XML file,
    and builds a DataFrame of records, including subdirectory names.
    Renames 'Name' column to 'POINT_ID' and includes 'TEST_N' for subdirectory.

    Args:
        test_filepaths (tuple): A tuple containing tuples with subdirectory
                                names and file paths to XML files.

    Returns:
        DataFrame: A DataFrame containing parsed test data.
    """
    test_records = []
    for subdirectory, filename in test_filepaths:
        try:
            with open(filename, 'r') as f:
                xml_string = f.read()
            record = parse_test_xml(xml_string)
            record['TEST_N'] = subdirectory  # Include subdirectory as TEST_N
            test_records.append(record)
        except FileNotFoundError:
            print(f"Error: File not found: {filename}")

    test_records = pd.DataFrame(test_records)
    test_records = test_records.rename(columns={'Name': 'POINT_ID'})
    return test_records


def parse_test_xml(xml_string):
    """
    Parses the provided XML string and returns a dictionary of test data.

    Args:
        xml_string (str): The XML content as a string.

    Returns:
        dict: A dictionary containing the extracted test data.
    """
    root = ET.fromstring(xml_string)
    test_data = {}

    for child in root:
        # Extract basic data
        if child.tag not in ('GPS', 'U', 'Gamma'):
            test_data[child.tag] = child.text

        # Extract GPS data (separate latitude and longitude)
        elif child.tag == 'GPS':
            test_data['Latitude'] = child.get('Latitude')
            test_data['Longitude'] = child.get('Longitude')

        # Handle elements with attributes (U and Gamma)
        elif child.tag in ('U', 'Gamma'):
            # Extract attribute values if they exist
            z_value = child.get('Z')
            unit_value = child.get('U') or child.get('Gamma')
            # Check for both U and Gamma
            if z_value and unit_value:
                test_data[child.tag] = f"{z_value} {unit_value}"
                # Combine Z and unit

    return test_data


def process_dmt_files(test_filepaths):
    """
    Loops through a tuple of file paths, parses each XML file,
    and builds two DataFrames: one for DMT setup data
    and one for DatiDMT data. Ensures 'POINT_ID' and 'TEST_N' are included.

    Args:
        test_filepaths (tuple): A tuple containing tuples with subdirectory
                                names and file paths to XML files.

    Returns:
        tuple: Two DataFrames, one for DMT setup data and one for DatiDMT data.
    """
    dmt_setup_records = []
    dati_dmt_records = []

    for subdirectory, filename in test_filepaths:
        try:
            with open(filename, 'r') as f:
                xml_string = f.read()
            setup_record, dati_dmt_record = parse_dmt_xml(xml_string,
                                                          subdirectory)
            setup_record['TEST_N'] = subdirectory
            for record in dati_dmt_record:
                record['TEST_N'] = subdirectory
            dmt_setup_records.append(setup_record)
            dati_dmt_records.extend(dati_dmt_record)
        except FileNotFoundError:
            print(f"Error: File not found: {filename}")

    dmt_setup_df = pd.DataFrame(dmt_setup_records)
    dati_dmt_df = pd.DataFrame(dati_dmt_records)
    dati_dmt_df = dati_dmt_df.rename(columns={'POINT_ID': 'POINT_ID'})

    return dmt_setup_df, dati_dmt_df


def parse_dmt_xml(xml_string, test_name):
    """
    Parses the provided XML string and returns a dictionary containing
    DMT data with a test name.

    Args:
        xml_string (str): The XML content as a string.
        test_name (str): The test name associated with the data.

    Returns:
        tuple: Two dictionaries, one for DMT setup data
        and one for DatiDMT data.
    """
    if not xml_string:
        return None, None

    root = ET.fromstring(xml_string)
    dmt_data = {}
    dati_dmt_data = []

    # Set POINT_ID
    dmt_data['POINT_ID'] = test_name

    # Extract basic data from top-level elements
    for child in root:
        if child.tag != 'DatiDMT':
            dmt_data[child.tag] = child.text

    # Extract data from DatiDMT element (nested dictionary)
    for element in root.find('DatiDMT'):
        record = {
            'POINT_ID': test_name,
            'Z': element.get('Z'),
            'A': element.get('A'),
            'B': element.get('B')
        }
        dati_dmt_data.append(record)

    return dmt_data, dati_dmt_data


def parse_sdmt_report_xlsx(directory, sdmt_report_filename):
    """
    Parses the provided XLSX file and returns dataframes containing
    S-wave, P-wave and DMT data with a test name.

    Args:
        directory (str): The location of the file associated with the data.
        sdmt_report_filename (str): The name of the file containing the data.

    Returns:
        three dataframes for S-wave, P-wave and DMT data,
        with the POINT_ID included on each row
    """
    # Load the workbook
    file_path = f'{directory}/{sdmt_report_filename}'
    try:
        xls = pd.ExcelFile(file_path)
    except FileNotFoundError:
        print(f"Error: File not found: {file_path}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    # Initialize DataFrames for S-wave, P-wave, and DMT test types
    s_wave_df = pd.DataFrame()
    p_wave_df = pd.DataFrame()
    dmt_df = pd.DataFrame()

    # Iterate over each sheet in the workbook
    for sheet_name in xls.sheet_names:
        # Use rsplit to handle multiple delimiters
        # and get the last item as the test type
        parts = sheet_name.rsplit(' - ', 1)
        if len(parts) == 2:
            test_name, test_type = parts[0], parts[1]

            # Check the test type and proceed accordingly
            if test_type == 'P-wave':
                # Read the sheet with row 4 as header, and skip rows 1, 2, 3
                # Additionally, only read the first four columns
                df = pd.read_excel(
                    xls,
                    sheet_name=sheet_name,
                    header=3,
                    usecols=range(4),
                    skiprows=[4])
                df['POINT_ID'] = test_name  # Add the test name to a new column
                p_wave_df = pd.concat([p_wave_df, df], ignore_index=True)

            elif test_type in ['S-wave', 'DMT']:
                # Read the sheet with row 3 as header, and skip row 4
                df = pd.read_excel(
                    xls,
                    sheet_name=sheet_name,
                    header=2,
                    skiprows=[3])
                df['POINT_ID'] = test_name  # Add the test name to a new column

                # Append the DataFrame to the appropriate test type DataFrame
                if test_type == 'S-wave':
                    s_wave_df = pd.concat([s_wave_df, df], ignore_index=True)
                elif test_type == 'DMT':
                    dmt_df = pd.concat([dmt_df, df], ignore_index=True)

    # Remove 'Zabs' column from all DataFrames if it exists
    if 'Zabs' in s_wave_df.columns:
        s_wave_df = s_wave_df.drop(columns=['Zabs'])

    if 'Zabs' in p_wave_df.columns:
        p_wave_df = p_wave_df.drop(columns=['Zabs'])

    if 'Zabs' in dmt_df.columns:
        dmt_df = dmt_df.drop(columns=['Zabs'])

    return s_wave_df, p_wave_df, dmt_df


def process_vs_files(directory, sdmt_tests):
    """
    Parses the VS XML files and returns a dataframe containing
    the VS setup data with the test name.

    Args:
        directory (str): The base directory where test folders are located.
        sdmt_tests (list): A list containing test folder names.

    Returns:
        DataFrame: A DataFrame containing parsed VS setup data.
    """
    vs_records = []
    for test_folder in sdmt_tests:
        point_id = test_folder  # Assuming point_id is the same as test_folder
        vs_directory = os.path.join(directory, test_folder, 'VS')

        # Recursively find all 'TI.dat' files in the VS directory
        for root_dir, _, files in os.walk(vs_directory):
            for file in files:
                if file == 'TI.dat':
                    file_path = os.path.join(root_dir, file)
                    try:
                        with open(file_path, 'r') as f:
                            xml_string = f.read()
                        vs_data = parse_vs_xml(xml_string, point_id)
                        vs_records.extend(vs_data)
                    except FileNotFoundError:
                        print(f"Error: File not found: {file_path}")

    vs_setup_df = pd.DataFrame(vs_records)
    return vs_setup_df if not vs_setup_df.empty else pd.DataFrame(
        columns=[
            'POINT_ID',
            'Z',
            'Selection',
            'FileNumber',
            'DeltaS',
            'DeltaT',
            'Filter_Code',
            'Filter_P1',
            'Window_Code',
            'Window_P1',
            'Window_P2',
            'Rephase_Code',
            'Rephase_P1'
            ])


def parse_vs_xml(xml_string, point_id):
    """
    Parses the provided VS XML string and returns a list of dictionaries
    containing VS data with a test name.

    Args:
        xml_string (str): The XML content as a string.
        point_id (str): The point ID associated with the data.

    Returns:
        list: A list of dictionaries containing the extracted VS data.
    """
    if not xml_string:
        return []

    root = ET.fromstring(xml_string)
    vs_data = []

    for processing in root.findall('Processing'):
        z_value = processing.get('Z')
        for seismogram in processing.findall('Seismogram'):
            record = {
                'POINT_ID': point_id,
                'Z': z_value,
                'Selection': seismogram.get('Selection'),
                'FileNumber': seismogram.get('FileNumber'),
                'DeltaS': seismogram.get('DeltaS'),
                'DeltaT': seismogram.get('DeltaT')}
            filter_elem = seismogram.find('Filter')
            if filter_elem is not None:
                record['Filter_Code'] = filter_elem.get('Code')
                record['Filter_P1'] = filter_elem.get('P1')

            window_elem = seismogram.find('Window')
            if window_elem is not None:
                record['Window_Code'] = window_elem.get('Code')
                record['Window_P1'] = window_elem.get('P1')
                record['Window_P2'] = window_elem.get('P2')

            rephase_elem = seismogram.find('Rephase')
            if rephase_elem is not None:
                record['Rephase_Code'] = rephase_elem.get('Code')
                record['Rephase_P1'] = rephase_elem.get('P1')

            vs_data.append(record)

    return vs_data


def process_sdm_files(directory, stack=False):
    """
    Recursively reads .sdm files from the 'VS' subdirectory in each test folder,
    processes the data, and either exports non-stacked or stacked signals.

    Args:
        directory (str): The base directory where test folders are located.
        stack (bool): Flag to indicate whether to export stacked data (True) or non-stacked (False).

    Returns:
        DataFrame: A DataFrame containing parsed sdm info data.
    """
    sdm_info_records = []
    files_by_depth = {}

    for test_folder in os.listdir(directory):  # Loop through test folders
        vs_directory = os.path.join(directory, test_folder, 'VS')

        if not os.path.isdir(vs_directory):
            continue  # Skip if VS folder doesn't exist

        # Recursively find all '.sdm' files in the VS directory
        for root_dir, _, files in os.walk(vs_directory):
            for file in files:
                if file.endswith('.sdm'):
                    file_path = os.path.join(root_dir, file)
                    try:
                        # Parse the SDM file directly
                        sdm_number = int(os.path.splitext(file)[0])
                        metadata, seismic_data = parse_seismic_file(file_path)

                        # If stacking is enabled, group by depth
                        if stack:
                            depth = metadata.get('ZRef', 'Unknown')
                            if depth not in files_by_depth:
                                files_by_depth[depth] = []
                            files_by_depth[depth].append(file_path)
                        else:
                            # Process non-stacked data
                            save_seismic_data_corrected(sdm_number, metadata, seismic_data, root_dir, stack=False)

                        # Ensuring POINT_ID and SdmNumber are properly added to sdm_info_records
                        sdm_info_records.append({
                            'POINT_ID': test_folder,  # The test folder is used as POINT_ID
                            'SdmNumber': sdm_number,  # The SDM number extracted from the file name
                            'Filename': file  # Store the file name for reference
                        })

                    except FileNotFoundError:
                        print(f"Error: File not found: {file_path}")

    if sdm_info_records:
        vs_sdm_info_df = pd.DataFrame(sdm_info_records)
        
        # Ensure both 'POINT_ID' and 'SdmNumber' exist before sorting
        if 'POINT_ID' in vs_sdm_info_df.columns and 'SdmNumber' in vs_sdm_info_df.columns:
            vs_sdm_info_df = vs_sdm_info_df.sort_values(by=['POINT_ID', 'SdmNumber'])
        else:
            print("Warning: 'POINT_ID' or 'SdmNumber' columns are missing in vs_sdm_info_df.")
    else:
        # Return an empty DataFrame with the required columns if no records
        vs_sdm_info_df = pd.DataFrame(columns=['POINT_ID', 'Filename', 'SdmNumber'])

    return vs_sdm_info_df


def process_inverted_sdm_files(directory, stack=False):
    """
    Recursively reads .sdm files from the 'VS' subdirectory in each test folder,
    processes the data, and exports non-stacked inverted signals for each file.

    Args:
        directory (str): The base directory where test folders are located.
    """
    for test_folder in os.listdir(directory):  # Loop through test folders
        vs_directory = os.path.join(directory, test_folder, 'VS')

        if not os.path.isdir(vs_directory):
            continue  # Skip if VS folder doesn't exist

        # Recursively find all '.sdm' files in the VS directory
        for root_dir, _, files in os.walk(vs_directory):
            for file in files:
                if file.endswith('.sdm'):
                    file_path = os.path.join(root_dir, file)
                    try:
                        # Parse the SDM file directly
                        sdm_number = int(os.path.splitext(file)[0])
                        metadata, seismic_data = parse_seismic_file(file_path)

                        # Save inverted non-stacked data
                        save_seismic_data_inverted(sdm_number, metadata, seismic_data, root_dir, stack=False)

                    except FileNotFoundError:
                        print(f"Error: File not found: {file_path}")
                        sys.stdout.flush()  # Force the print buffer to flush


def process_vp_files(directory):
    """
    Recursively reads .sdm files from the 'VP' subdirectory in each test folder,
    processes the data, and exports non-stacked signals for VP.

    Args:
        directory (str): The base directory where test folders are located.
    """
    sdm_info_records = []

    for test_folder in os.listdir(directory):  # Loop through test folders
        vp_directory = os.path.join(directory, test_folder, 'VP')

        if not os.path.isdir(vp_directory):
            continue  # Skip if VP folder doesn't exist

        # Recursively find all '.sdm' files in the VP directory
        for root_dir, _, files in os.walk(vp_directory):
            for file in files:
                if file.endswith('.sdm'):
                    file_path = os.path.join(root_dir, file)
                    try:
                        # Parse the SDM file directly
                        sdm_number = int(os.path.splitext(file)[0])
                        metadata, seismic_data = parse_seismic_file(file_path)

                        # Export non-stacked VP data
                        save_vp_data(sdm_number, metadata, seismic_data, root_dir)

                    except FileNotFoundError:
                        print(f"Error: File not found: {file_path}")
                        sys.stdout.flush()  # Force the print buffer to flush


    if sdm_info_records:
        vp_sdm_info_df = pd.DataFrame(sdm_info_records)
        # Order rows in vp_sdm_info_df by POINT_ID, then by SdmNumber
        vp_sdm_info_df = vp_sdm_info_df.sort_values(by=['POINT_ID', 'SdmNumber'])
    else:
        # Return an empty DataFrame with the required columns if no records
        vp_sdm_info_df = pd.DataFrame(columns=['POINT_ID', 'Filename', 'SdmNumber'])

    return vp_sdm_info_df


import xml.etree.ElementTree as ET

# def parse_sdm_xml(xml_string, filename, test_folder, sdm_number):
#     """
#     Parse the provided XML string and return a dictionary containing the metadata.

#     Args:
#         xml_string (str): The XML content as a string.
#         filename (str): The filename being parsed.
#         test_folder (str): The test folder.
#         sdm_number (int): The SDM number.

#     Returns:
#         dict: A dictionary containing metadata, including sensor-specific ZSensor and DelaySample values.
#     """
#     root = ET.fromstring(xml_string)
#     metadata = {'Filename': filename, 'TestFolder': test_folder, 'SdmNumber': sdm_number}

#     # Extract delay from ExternalTrigger element (if exists)
#     external_trigger = root.find('.//ExternalTrigger')
#     if external_trigger is not None:
#         metadata['Delay'] = int(external_trigger.get('Delay', 0))

#     # Extract ZRef (optional)
#     metadata['ZRef'] = root.findtext('.//ZRef', 'Unknown')

#     # Extract SampleTime for time calculation
#     metadata['SampleTime'] = int(root.findtext('.//SampleTime', 0))

#     # Extract Sensor data specifically from SeismicData
#     seismic_data = root.find('.//SeismicData')
#     if seismic_data is not None:
#         for sensor in seismic_data.findall('Sensor'):
#             sensor_id = sensor.get('ID')
#             z_sensor = sensor.get('ZSensor')
#             delay_sample = sensor.get('DelaySample')

#             if z_sensor is not None:
#                 z_sensor = float(z_sensor)  # Convert ZSensor to float if it exists
#             else:
#                 logging.info(f"Warning: ZSensor missing for Sensor ID {sensor_id}")

#             if delay_sample is not None:
#                 delay_sample = int(delay_sample)  # Convert DelaySample to integer if it exists
#             else:
#                 delay_sample = 0  # Default to 0 if DelaySample is missing

#             if sensor_id == "1":
#                 metadata['ZSensor1'] = z_sensor if z_sensor is not None else 'Unknown'
#                 metadata['DelaySample1'] = delay_sample
#             elif sensor_id == "2":
#                 metadata['ZSensor2'] = z_sensor if z_sensor is not None else 'Unknown'
#                 metadata['DelaySample2'] = delay_sample
#     else:
#         logging.info("Error: SeismicData section not found in the XML.")

#     # Debugging: Print extracted metadata to ensure ZSensor values are correct
#     logging.info(f"Extracted metadata for {filename}: {metadata}")

#     return metadata


def transform_raw_data(dataframes):
    # Unpack the input dataframes
    (test_meta_df, dmt_setup_df, dati_dmt_df, dmt_data_df, s_wave_data_df,
     p_wave_data_df, vs_setup_df, vs_sdm_info_df) = dataframes

    # Check if test_meta_df is empty
    if not test_meta_df.empty:
        # Create a dictionary for quick lookup of POINT_ID based on TEST_N
        test_meta_lookup = test_meta_df.set_index(
            'TEST_N')['POINT_ID'].to_dict()
        reverse_test_meta_lookup = test_meta_df.set_index(
            'POINT_ID')['TEST_N'].to_dict()
    else:
        test_meta_lookup = {}
        reverse_test_meta_lookup = {}

    # Update POINT_ID in dmt_setup_df and dati_dmt_df based on test_meta_df
    if not dmt_setup_df.empty:
        dmt_setup_df['POINT_ID'] = dmt_setup_df['TEST_N'].map(test_meta_lookup)
    if not dati_dmt_df.empty:
        dati_dmt_df['POINT_ID'] = dati_dmt_df['TEST_N'].map(test_meta_lookup)

    # Add TEST_N to dmt_data_df, s_wave_data_df, and p_wave_data_df
    if not dmt_data_df.empty:
        dmt_data_df['TEST_N'] = dmt_data_df['POINT_ID'].map(
            reverse_test_meta_lookup)
    if not s_wave_data_df.empty:
        s_wave_data_df['TEST_N'] = s_wave_data_df['POINT_ID'].map(
            reverse_test_meta_lookup)
    if not p_wave_data_df.empty:
        p_wave_data_df['TEST_N'] = p_wave_data_df['POINT_ID'].map(
            reverse_test_meta_lookup)

    # Add TEST_N to vs_setup_df and vs_sdm_info_df as a duplicate of POINT_ID
    if not vs_setup_df.empty:
        vs_setup_df['TEST_N'] = vs_setup_df['POINT_ID']
        vs_setup_df['POINT_ID'] = vs_setup_df['TEST_N'].map(test_meta_lookup)
    if not vs_sdm_info_df.empty:
        vs_sdm_info_df['TEST_N'] = vs_sdm_info_df['POINT_ID']
        vs_sdm_info_df['POINT_ID'] = vs_sdm_info_df['TEST_N'].map(
            test_meta_lookup)

    # Left join dmt_setup_df with test_meta_df on POINT_ID
    if not test_meta_df.empty and not dmt_setup_df.empty:
        test_meta_df = test_meta_df.merge(dmt_setup_df,
                                          on='POINT_ID',
                                          how='left',
                                          suffixes=('', '_dmt_setup'))

    # Left join vs_sdm_info_df,
    # taking only the first row for each unique POINT_ID
    if not vs_sdm_info_df.empty:
        vs_sdm_info_df_sorted = vs_sdm_info_df.sort_values(by=['POINT_ID',
                                                               'SdmNumber'])
        vs_sdm_info_df_first = vs_sdm_info_df_sorted.drop_duplicates(
            subset=['POINT_ID'], keep='first')
        if not test_meta_df.empty:
            test_meta_df = test_meta_df.merge(vs_sdm_info_df_first,
                                              on='POINT_ID',
                                              how='left',
                                              suffixes=('', '_vs_sdm_info'))

    dataframes = (test_meta_df, dmt_setup_df, dati_dmt_df,
                  dmt_data_df, s_wave_data_df, p_wave_data_df,
                  vs_setup_df, vs_sdm_info_df)

    # Return the transformed dataframes
    return dataframes


def export_to_excel(directory, dataframes):
    """
    Exports given dataframes to individual worksheets in an Excel workbook.

    Args:
        directory (str): The directory where the Excel workbook will be saved.
        dataframes (list): List of dataframes to be exported.
    """
    file_path = os.path.join(directory, 'sdmt_import.xlsx')

    # Create a workbook and add a worksheet for each dataframe
    wb = Workbook()
    ws_mapping = {
        'test_meta_df': dataframes[0],
        'dmt_setup_df': dataframes[1],
        'dati_dmt_df': dataframes[2],
        'dmt_data_df': dataframes[3],
        's_wave_data_df': dataframes[4],
        'p_wave_data_df': dataframes[5],
        'vs_setup_df': dataframes[6],
        'vs_sdm_info_df': dataframes[7]
    }

    for sheet_name, df in ws_mapping.items():
        ws = wb.create_sheet(title=sheet_name)
        for r_idx, row in enumerate(dataframe_to_rows(df, index=False,
                                                      header=True)):
            for c_idx, value in enumerate(row, 1):
                cell = ws.cell(row=r_idx + 1, column=c_idx, value=value)
                if r_idx == 0:  # Header row
                    cell.font = Font(bold=True)
                # Auto-adjust column width
                column_letter = ws.cell(row=1, column=c_idx).column_letter
                ws.column_dimensions[column_letter].width = max(
                    len(str(value)) + 2,
                    ws.column_dimensions[column_letter].width)

    # Remove the default created sheet
    del wb['Sheet']

    # Save the workbook
    wb.save(file_path)
    print(f"Data successfully exported to {file_path}")
    sys.stdout.flush()  # Force the print buffer to flush


def export_button_callback():
    global export_directory, transformed_dataframes
    export_to_excel(export_directory, transformed_dataframes)


def create_POINT_df(test_meta_df):
    """
    Create a new dataframe 'POINT_df' from 'test_meta_df' with specified column
    mappings.

    Args:
        test_meta_df (DataFrame):
            The source dataframe containing test metadata.

    Returns:
        DataFrame: A new dataframe 'POINT_df' with selected columns and new
        column names.
    """
    column_mappings = {
        'POINT_ID': 'POINT_ID',
        'Date': 'StartDate',
        'EndDate': 'EndDate',
        'Remark': 'Remarks',
        'Operator': 'Crew',
        'Latitude': 'Latitude',
        'Longitude': 'Longitude'
    }

    # Check for the existence of required columns
    # and add default values if missing
    for col in column_mappings.keys():
        if col != 'POINT_ID' and col not in test_meta_df.columns:
            test_meta_df[col] = None  # or any default value you prefer

    # Ensure POINT_ID exists
    if 'POINT_ID' not in test_meta_df.columns:
        raise KeyError(
            "The required column 'POINT_ID' is missing from test_meta_df")

    # Create POINT_df with existing columns
    available_columns = [
        col for col in column_mappings.keys() if col in test_meta_df.columns]
    POINT_df = test_meta_df[
        available_columns].rename(
            columns={col: column_mappings[col] for col in available_columns})

    # Replace NaN values in Remarks and Crew columns with a blank value
    if 'Remarks' in POINT_df.columns:
        POINT_df['Remarks'] = POINT_df['Remarks'].fillna('')
    if 'Crew' in POINT_df.columns:
        POINT_df['Crew'] = POINT_df['Crew'].fillna('')

    # Convert date columns to datetime format consistent with SQL Server
    if 'StartDate' in POINT_df.columns:
        POINT_df['StartDate'] = pd.to_datetime(
            POINT_df['StartDate'], errors='coerce')
    if 'EndDate' in POINT_df.columns:
        POINT_df['EndDate'] = pd.to_datetime(
            POINT_df['EndDate'], errors='coerce')

    # Convert latitude and longitude to decimal format
    if 'Latitude' in POINT_df.columns:
        POINT_df['Latitude'] = pd.to_numeric(
            POINT_df['Latitude'], errors='coerce')
    if 'Longitude' in POINT_df.columns:
        POINT_df['Longitude'] = pd.to_numeric(
            POINT_df['Longitude'], errors='coerce')

    # Ensure POINT_ID, Remarks, and Crew are string format
    POINT_df['POINT_ID'] = POINT_df['POINT_ID'].astype(str)
    if 'Remarks' in POINT_df.columns:
        POINT_df['Remarks'] = POINT_df['Remarks'].astype(str)
    if 'Crew' in POINT_df.columns:
        POINT_df['Crew'] = POINT_df['Crew'].astype(str)

    return POINT_df


def create_DMTG_df(test_meta_df):
    """
    Create a new dataframe 'DMTG_df' from 'test_meta_df' with specified column
    mappings. Add a column 'ITEM' and drop rows where 'ITEM' is 0. Ensure
    'POINT_ID' and 'ITEM' are the first two columns.

    Args:
        test_meta_df (DataFrame): The source dataframe containing
        test metadata.

    Returns:
        DataFrame: A new dataframe 'DMTG_df' with selected columns and new
        column names.
    """
    column_mappings = {
        'POINT_ID': 'POINT_ID',
        'Date': 'Date',
        'Zw': 'GroundWaterLevel',
        'IsDmtActive': 'ITEM',
        'DeltaA': 'BladeA',
        'DeltaB': 'BladeB'
    }

    # Ensure required columns exist in the DataFrame
    for col in column_mappings.keys():
        if col not in test_meta_df.columns:
            test_meta_df[col] = None  # or any default value you prefer

    DMTG_df = test_meta_df[list(column_mappings.keys())].rename(
        columns=column_mappings)

    # Convert date columns to datetime format consistent with SQL Server
    DMTG_df['Date'] = pd.to_datetime(DMTG_df['Date'], errors='coerce')

    # Convert GroundWaterLevel, BladeA, and BladeB to decimal format
    DMTG_df['GroundWaterLevel'] = pd.to_numeric(
        DMTG_df['GroundWaterLevel'], errors='coerce')
    DMTG_df['BladeA'] = pd.to_numeric(DMTG_df['BladeA'], errors='coerce')
    DMTG_df['BladeB'] = pd.to_numeric(DMTG_df['BladeB'], errors='coerce')

    # Ensure POINT_ID and ITEM are string format
    DMTG_df['POINT_ID'] = DMTG_df['POINT_ID'].astype(str)
    DMTG_df['ITEM'] = DMTG_df['ITEM'].astype(str)

    # Drop rows where 'ITEM' is 0
    DMTG_df = DMTG_df[DMTG_df['ITEM'] != "0"]

    # Ensure 'POINT_ID' and 'ITEM' are the first two columns
    columns_order = ['POINT_ID', 'ITEM'] + [
        col for col in DMTG_df.columns if col not in ['POINT_ID', 'ITEM']]
    DMTG_df = DMTG_df[columns_order]

    return DMTG_df


def create_DMTT_df(dmt_data_df, test_meta_df):
    """
    Create a new dataframe 'DMTT_df' from 'dmt_data_df' with specified column
    mappings. Add a column 'ITEM' and fill every row with the relevant value
    referenced by the POINT_ID column from the test_meta_df column IsDmtActive.
    Add a column 'REFERENCE' with a constant value.

    Args:
        dmt_data_df (DataFrame): The source dataframe containing DMT data.
        test_meta_df (DataFrame): The dataframe containing test metadata.

    Returns:
        DataFrame: A new dataframe 'DMTT_df' with selected columns and new
        column names.
    """
    if dmt_data_df.empty:
        return pd.DataFrame(columns=[
            'POINT_ID', 'ITEM', 'REFERENCE', 'TOP', 'DeltaA', 'DeltaB',
            'APressure', 'BPressure', 'CPressure', 'BulkUnitWeight',
            'EffectiveVerticalStress', 'u0', 'ACorrected', 'BCorrected',
            'CCorrected', 'MaterialIndex', 'Dilatometer', 'HorizontalStress',
            'PoreIndex', 'VerticalDrainedConstrainedModulus',
            'UndrainedShearStrength', 'FrictionAngle', 'LateralEarthPressure',
            'EffectiveHorizontalStress', 'OCR', 'Description'
        ])

    column_mappings = {
        'Z': 'TOP',
        'Δ A': 'DeltaA',
        'Δ B': 'DeltaB',
        'A': 'APressure',
        'B': 'BPressure',
        'C': 'CPressure',
        'γ': 'BulkUnitWeight',
        "σ' v": 'EffectiveVerticalStress',
        'U₀': 'u0',
        'P₀': 'ACorrected',
        'P₁': 'BCorrected',
        'P₂': 'CCorrected',
        'Id': 'MaterialIndex',
        'Ed': 'Dilatometer',
        'Kd': 'HorizontalStress',
        'Ud': 'PoreIndex',
        'M': 'VerticalDrainedConstrainedModulus',
        'Su': 'UndrainedShearStrength',
        'φ': 'FrictionAngle',
        'K₀': 'LateralEarthPressure',
        "σ' h": 'EffectiveHorizontalStress',
        'OCR': 'OCR',
        'Soil description': 'Description',
        'POINT_ID': 'POINT_ID'
    }

    # Ensure required columns exist in the DataFrame
    for col in column_mappings.keys():
        if col not in dmt_data_df.columns:
            dmt_data_df[col] = None  # or any default value you prefer

    # Create the new dataframe with the mapped columns
    DMTT_df = dmt_data_df[list(column_mappings.keys())].rename(
        columns=column_mappings)

    # Lookup to fill ITEM column
    item_lookup = test_meta_df.set_index('POINT_ID')['IsDmtActive'].to_dict()
    DMTT_df['ITEM'] = DMTT_df['POINT_ID'].map(item_lookup)

    # Add the REFERENCE column with a constant value
    DMTT_df['REFERENCE'] = "Standard Marchetti formulas from SDMT Pro"

    # Ensure POINT_ID, REFERENCE, Description, and ITEM are string format
    DMTT_df['POINT_ID'] = DMTT_df['POINT_ID'].astype(str)
    DMTT_df['REFERENCE'] = DMTT_df['REFERENCE'].astype(str)
    DMTT_df['Description'] = DMTT_df['Description'].astype(str)
    DMTT_df['ITEM'] = DMTT_df['ITEM'].astype(str)

    # Convert all other columns to decimal format
    for col in DMTT_df.columns:
        if col not in ['POINT_ID', 'REFERENCE', 'Description', 'ITEM']:
            DMTT_df[col] = pd.to_numeric(DMTT_df[col], errors='coerce')

    return DMTT_df


def create_Seismic_Downhole_df(test_meta_df):
    """
    Create a new dataframe 'Seismic_Downhole_df' from 'test_meta_df' with
    specified column mappings. Add a column 'ITEM' and drop rows where 'ITEM'
    is 0. Ensure 'POINT_ID' and 'ITEM' are the first two columns.

    Args:
        test_meta_df (DataFrame): The source dataframe containing test
        metadata.

    Returns:
        DataFrame: A new dataframe 'Seismic_Downhole_df' with selected columns
        and new column names.
    """
    column_mappings = {
        'POINT_ID': 'POINT_ID',
        'Date': 'Date',
        'IsVsActive': 'ITEM',
        'DistHammer': 'SourceOffset_Vs',
        'Zw': 'GroundWaterLevel'
    }

    # Ensure required columns exist in the DataFrame
    for col in column_mappings.keys():
        if col not in test_meta_df.columns:
            test_meta_df[col] = None  # or any default value you prefer

    Seismic_Downhole_df = test_meta_df[list(column_mappings.keys())].rename(
        columns=column_mappings)

    # Drop rows where 'ITEM' is 0 (converted to string for comparison)
    Seismic_Downhole_df = Seismic_Downhole_df[
        Seismic_Downhole_df['ITEM'] != "0"]

    # Convert columns to appropriate formats
    Seismic_Downhole_df['POINT_ID'] = Seismic_Downhole_df[
        'POINT_ID'].astype(str)
    Seismic_Downhole_df['ITEM'] = Seismic_Downhole_df[
        'ITEM'].astype(str)
    Seismic_Downhole_df['Date'] = pd.to_datetime(Seismic_Downhole_df[
        'Date'], errors='coerce')
    Seismic_Downhole_df['GroundWaterLevel'] = pd.to_numeric(
        Seismic_Downhole_df['GroundWaterLevel'], errors='coerce')
    Seismic_Downhole_df['SourceOffset_Vs'] = pd.to_numeric(
        Seismic_Downhole_df['SourceOffset_Vs'], errors='coerce')

    columns_order = ['POINT_ID', 'ITEM'] + [
        col for col in Seismic_Downhole_df.columns if col not in [
            'POINT_ID', 'ITEM']]
    Seismic_Downhole_df = Seismic_Downhole_df[columns_order]

    return Seismic_Downhole_df


def create_Vs_Data_df(s_wave_data_df, test_meta_df):
    """
    Create a new dataframe 'Vs_Data_df' from 's_wave_data_df' with specified
    column mappings. Add a column 'ITEM' filled with values from 'IsVsActive'
    in 'test_meta_df' based on 'POINT_ID'. Create two new columns 'VsLB' and
    'VsUB' with the min and max values from the 'Repeatability' column, and
    drop rows with NaN in the 'VsAvg' column.

    Args:
        s_wave_data_df (DataFrame): The source dataframe containing S-wave
        data.
        test_meta_df (DataFrame): The dataframe containing test metadata.

    Returns:
        DataFrame: A new dataframe 'Vs_Data_df' with selected columns, new
        column names, and additional columns.
    """
    column_mappings = {
        'POINT_ID': 'POINT_ID',
        'Z': 'TOP',
        'Vs': 'VsAvg',
        'Repeatability': 'Repeatability',
        'Var Coefficient': 'VarCoefficient'
    }

    # Ensure required columns exist in the DataFrame
    for col in column_mappings.keys():
        if col not in s_wave_data_df.columns:
            s_wave_data_df[col] = None  # or any default value you prefer

    Vs_Data_df = s_wave_data_df[list(column_mappings.keys())].rename(
        columns=column_mappings)
    Vs_Data_df = Vs_Data_df.dropna(subset=['VsAvg'])
    item_lookup = test_meta_df.set_index('POINT_ID')['IsVsActive'].to_dict()
    Vs_Data_df['ITEM'] = Vs_Data_df['POINT_ID'].map(item_lookup)

    # Ensure appropriate formats
    Vs_Data_df['POINT_ID'] = Vs_Data_df['POINT_ID'].astype(str)
    Vs_Data_df['Repeatability'] = Vs_Data_df['Repeatability'].astype(str)
    Vs_Data_df['ITEM'] = Vs_Data_df['ITEM'].astype(str)

    # Convert other columns to decimal
    decimal_columns = ['TOP', 'VsAvg', 'VarCoefficient']
    Vs_Data_df[decimal_columns] = Vs_Data_df[decimal_columns].apply(
        pd.to_numeric, errors='coerce')

    # Create VsLB and VsUB columns
    Vs_Data_df['VsLB'] = Vs_Data_df['Repeatability'].apply(
        lambda x: min(map(float, x.split(','))) if pd.notna(x) else None)
    Vs_Data_df['VsUB'] = Vs_Data_df['Repeatability'].apply(
        lambda x: max(map(float, x.split(','))) if pd.notna(x) else None)

    columns_order = ['POINT_ID', 'ITEM'] + [
        col for col in Vs_Data_df.columns if col not in ['POINT_ID', 'ITEM']]
    Vs_Data_df = Vs_Data_df[columns_order]

    return Vs_Data_df


def create_Vs_Parameters_Setup_df(test_meta_df):
    """
    Create a new dataframe 'Vs_Parameters_Setup' from 'test_meta_df' with
    specified column mappings. Add columns 'REFERENCE' and 'Methods' with
    predefined values. Ensure 'POINT_ID' and 'ITEM' are the first two columns.

    Args:
        test_meta_df (DataFrame):
            The source dataframe containing test metadata.

    Returns:
        DataFrame: A new dataframe 'Vs_Parameters_Setup' with selected columns,
        new column names, and additional columns 'REFERENCE' and 'Methods'.
    """
    column_mappings = {
        'POINT_ID': 'POINT_ID',
        'Date': 'Date',
        'IsVsActive': 'ITEM'
    }

    # Create the new dataframe with the mapped columns
    Vs_Parameters_Setup_df = test_meta_df[list(column_mappings.keys())].rename(
        columns=column_mappings)

    # Drop rows where 'ITEM' is 0 (converted to string for comparison)
    Vs_Parameters_Setup_df = Vs_Parameters_Setup_df[
        Vs_Parameters_Setup_df['ITEM'] != "0"]

    # Add the additional columns with predefined values
    Vs_Parameters_Setup_df['REFERENCE'] = "Marchetti TI"
    Vs_Parameters_Setup_df['Methods'] = "Marchetti True Interval"

    # Ensure 'POINT_ID' and 'ITEM' are the first two columns
    columns_order = [
        'POINT_ID', 'ITEM', 'REFERENCE'] + [
            col for col in Vs_Parameters_Setup_df.columns if col not in [
                'POINT_ID', 'ITEM', 'REFERENCE']]
    Vs_Parameters_Setup_df = Vs_Parameters_Setup_df[columns_order]

    return Vs_Parameters_Setup_df


def create_Vs_Parameters_Data_df(s_wave_data_df, test_meta_df):
    """
    Create a new dataframe 'Vs_Parameters_Data_df' from 's_wave_data_df' with
    specified column mappings. Add a column 'ITEM' filled with values from
    'IsVsActive' in 'test_meta_df' based on 'POINT_ID'. Create two new columns
    'ShearModLB' and 'ShearModUB' with the min and max values from the
    'Repeatability' column, squared and multiplied by the values in the
    'BulkDensity' column, and drop rows with NaN in the 'ShearModAvg' column.

    Args:
        s_wave_data_df (DataFrame): The source dataframe containing S-wave
        data.
        test_meta_df (DataFrame): The dataframe containing test metadata.

    Returns:
        DataFrame: A new dataframe 'Vs_Parameters_Data_df'
        with selected columns,
        new column names, and additional columns.
    """
    if s_wave_data_df.empty:
        return pd.DataFrame(columns=[
            'POINT_ID', 'ITEM', 'REFERENCE', 'TOP', 'ShearModAvg',
            'ShearModLB', 'ShearModUB', 'Repeatability', 'BulkDensity'
        ])

    column_mappings = {
        'POINT_ID': 'POINT_ID',
        'Z': 'TOP',
        'ρ': 'BulkDensity',
        'G₀': 'ShearModAvg',
        'Repeatability': 'Repeatability'
    }

    # Ensure required columns exist in the DataFrame
    for col in column_mappings.keys():
        if col not in s_wave_data_df.columns:
            s_wave_data_df[col] = None  # or any default value you prefer

    Vs_Parameters_Data_df = s_wave_data_df[list(
        column_mappings.keys())].rename(columns=column_mappings)
    Vs_Parameters_Data_df = Vs_Parameters_Data_df.dropna(
        subset=['ShearModAvg'])
    item_lookup = test_meta_df.set_index('POINT_ID')['IsVsActive'].to_dict()
    Vs_Parameters_Data_df['ITEM'] = Vs_Parameters_Data_df[
        'POINT_ID'].map(item_lookup)

    def calculate_shear_modulus(repeatability, bulk_density):
        if pd.notna(repeatability) and pd.notna(bulk_density):
            values = list(map(float, repeatability.split(',')))
            if values:
                vs_lb = min(values)
                vs_ub = max(values)
                shear_mod_lb = (vs_lb ** 2) * bulk_density / 1000000
                shear_mod_ub = (vs_ub ** 2) * bulk_density / 1000000
                return shear_mod_lb, shear_mod_ub
        return None, None

    Vs_Parameters_Data_df['ShearModLB'], Vs_Parameters_Data_df[
        'ShearModUB'] = zip(
            *Vs_Parameters_Data_df.apply(
                lambda row: calculate_shear_modulus(
                    row['Repeatability'], row['BulkDensity']), axis=1)
    )

    Vs_Parameters_Data_df = Vs_Parameters_Data_df.drop(
        columns=['Repeatability'])

    # Add the additional columns
    Vs_Parameters_Data_df['REFERENCE'] = "Marchetti TI"

    # Ensure 'POINT_ID' and 'ITEM' are the first two columns
    columns_order = ['POINT_ID', 'ITEM', 'REFERENCE'] + [
        col for col in Vs_Parameters_Data_df.columns if col not in [
            'POINT_ID', 'ITEM', 'REFERENCE']
    ]
    Vs_Parameters_Data_df = Vs_Parameters_Data_df[columns_order]

    return Vs_Parameters_Data_df


def parse_seismic_file(file_path):
    """
    Parse the seismic file XML content and extract metadata and seismic data (S1, S2).

    Args:
        file_path (str): The path to the seismic file.

    Returns:
        tuple: A dictionary of metadata and a DataFrame containing seismic data (S1, S2).
    """
    tree = ET.parse(file_path)
    root = tree.getroot()

    # Extract metadata
    metadata = {}
    seismic_data = []

    # Debug: Output the structure of the XML root
    print(f"Root tag: {root.tag}, attributes: {root.attrib}")
    sys.stdout.flush()  # Force the print buffer to flush

    # Iterate through the children of the root directly, since the root itself is <SeismicData>
    for element in root:
        # print(f"Processing element: {element.tag}, attributes: {element.attrib}")
        # sys.stdout.flush()  # Force the print buffer to flush

        if element.tag == 'Sensor':  # Extract Sensor information
            sensor_id = element.get('ID')
            z_sensor = element.get('ZSensor', 'Unknown')
            if sensor_id == '1':
                metadata['ZSensor1'] = z_sensor
            elif sensor_id == '2':
                metadata['ZSensor2'] = z_sensor
            print(f"Sensor {sensor_id}: ZSensor={z_sensor}")  # Debug: Print ZSensor values
            sys.stdout.flush()  # Force the print buffer to flush

        elif element.tag == 'ExternalTrigger':  # Extract ExternalTrigger Delay
            metadata['Delay'] = int(element.get('Delay', 0))

        elif element.tag == 'SampleTime':  # Extract SampleTime for time calculation
            metadata['SampleTime'] = int(element.text)

        elif element.tag == 'ZRef':  # Extract ZRef (optional)
            metadata['ZRef'] = element.text

        # Extract other metadata elements
        else:
            metadata[element.tag] = element.attrib if element.attrib else element.text

    # Debug: Print metadata after extraction
    print(f"Metadata extracted: {metadata}")
    sys.stdout.flush()  # Force the print buffer to flush

    # Extract seismic data (S1 and S2 values)
    for element in root.findall('S'):
        seismic_data.append({
            'S1': element.attrib['S1'],
            'S2': element.attrib['S2']
        })

    # Convert seismic data into a DataFrame
    seismic_data_df = pd.DataFrame(seismic_data, columns=['S1', 'S2'])

    return metadata, seismic_data_df


def save_seismic_data_corrected(file_number_str, metadata, seismic_data, output_dir, stack=False):
    """
    Save seismic data with corrected values to files in the "Stacked VS" or "Individual VS" directory.

    Args:
        file_number_str (str): The unique file identifier string.
        metadata (dict): Metadata extracted from the file, including ZSensor values for depth adjustment.
        seismic_data (DataFrame): DataFrame containing seismic data (S1 and S2).
        output_dir (str): The directory where the files will be saved.
        stack (bool): Whether the data is stacked (True) or non-stacked (False).
    """
    # Set folder name based on whether data is stacked or individual
    folder_name = "Stacked VS" if stack else "Individual VS"
    export_dir = os.path.join(output_dir, folder_name)
    os.makedirs(export_dir, exist_ok=True)

    # Prepare header row based on metadata
    z_sensor1 = metadata.get('ZSensor1', 'Unknown')
    z_sensor2 = metadata.get('ZSensor2', 'Unknown')

    print(f"Saving file {file_number_str}: ZSensor1={z_sensor1}, ZSensor2={z_sensor2}")
    sys.stdout.flush()  # Force the print buffer to flush

    # Prepare headers using the correct ZSensor values
    header_s1 = f"Tempo(ms); Amplitude(cm/s) Valores provisorios da Sismica; Profundidade[m] ={metadata.get('ZSensor1', 'Unknown')}\n"
    header_s2 = f"Tempo(ms); Amplitude(cm/s) Valores provisorios da Sismica; Profundidade[m] ={metadata.get('ZSensor2', 'Unknown')}\n"
    
    # Use the correct delay from ExternalTrigger and SampleTime
    delay = metadata.get('Delay', 0)
    sample_time = int(metadata.get('SampleTime', 0))
    
    # Function to create the time column
    def create_time_column(row_num):
        return ((sample_time * row_num) / 1000) + delay  # Convert to milliseconds

    # Create file names for S1 and S2 within the appropriate folder
    file_s1 = os.path.join(export_dir, f"{file_number_str}-S1-A.txt")
    file_s2 = os.path.join(export_dir, f"{file_number_str}-S2-A.txt")
    
    # Save data for S1 with ZSensor1 in the header
    with open(file_s1, 'w') as f_s1:
        f_s1.write(header_s1)
        for idx, row in seismic_data.iterrows():
            time_ms = create_time_column(idx)
            f_s1.write(f"{time_ms:.3f}\t{row['S1']}\n")
    
    # Save data for S2 with ZSensor2 in the header
    with open(file_s2, 'w') as f_s2:
        f_s2.write(header_s2)
        for idx, row in seismic_data.iterrows():
            time_ms = create_time_column(idx)
            f_s2.write(f"{time_ms:.3f}\t{row['S2']}\n")
    
    return file_s1, file_s2


def save_vp_data(file_number_str, metadata, seismic_data, output_dir):
    """
    Save VP data to files in the "Individual VP" directory with the letter C in the file name.

    Args:
        file_number_str (str): The unique file identifier string.
        metadata (dict): Metadata extracted from the file, including ZSensor values for depth adjustment.
        seismic_data (DataFrame): DataFrame containing VP data (S1 and S2).
        output_dir (str): The directory where the files will be saved.
    """
    # Create the "Individual VP" directory if it doesn't exist
    vp_dir = os.path.join(output_dir, "Individual VP")
    os.makedirs(vp_dir, exist_ok=True)

    # Prepare header row based on ZSensor values for each sensor
    z_sensor1 = metadata.get('ZSensor1', 'Unknown')  # ZSensor for Sensor 1 (S1)
    z_sensor2 = metadata.get('ZSensor2', 'Unknown')  # ZSensor for Sensor 2 (S2)
    header_s1 = f"Tempo(ms); Amplitude(cm/s) Valores provisorios da Sismica; Profundidade[m] ={metadata.get('ZSensor1', 'Unknown')}\n"
    header_s2 = f"Tempo(ms); Amplitude(cm/s) Valores provisorios da Sismica; Profundidade[m] ={metadata.get('ZSensor2', 'Unknown')}\n"
    
    # Use the correct delay and SampleTime from metadata
    delay = metadata.get('Delay', 0)
    sample_time = int(metadata.get('SampleTime', 0))
    
    # Function to create the time column, starting from the correct delay and sample time
    def create_time_column(row_num):
        return ((sample_time * row_num) / 1000) + delay # Convert to milliseconds

    # Create file names for VP data with the letter "C"
    file_s1 = os.path.join(vp_dir, f"{file_number_str}-S1-C.txt")
    file_s2 = os.path.join(vp_dir, f"{file_number_str}-S2-C.txt")
    
    # Save VP data for S1 with ZSensor1 in the header
    with open(file_s1, 'w') as f_s1:
        f_s1.write(header_s1)
        for idx, row in seismic_data.iterrows():
            time_ms = create_time_column(idx)  # First row starts at the correct delay
            f_s1.write(f"{time_ms:.3f}\t{row['S1']}\n")
    
    # Save VP data for S2 with ZSensor2 in the header
    with open(file_s2, 'w') as f_s2:
        f_s2.write(header_s2)
        for idx, row in seismic_data.iterrows():
            time_ms = create_time_column(idx)  # First row starts at the correct delay
            f_s2.write(f"{time_ms:.3f}\t{row['S2']}\n")
    
    return file_s1, file_s2


def invert_signal(value):
    """
    Inverts the signal value assuming it is transmitted via hex (0 to 255).
    255 becomes 0, 0 becomes 255, and values in between are inverted proportionally.

    Args:
        value (int or float): The signal value to be inverted.

    Returns:
        int: The inverted signal value.
    """
    max_value = 255
    try:
        # Convert to float and ensure value is between 0 and max_value
        value = float(value)
        if value < 0 or value > max_value:
            raise ValueError("Value is out of range for inversion (0 to 255)")
        inverted_value = max_value - value
        return int(inverted_value)
    except (ValueError, TypeError):
        return value  # Return original value if inversion is not possible


def save_seismic_data_inverted(file_number_str, metadata, seismic_data, output_dir, stack=False):
    """
    Save seismic data with inverted values to files in the "Stacked VS" or "Individual VS" directory.

    Args:
        file_number_str (str): The unique file identifier string.
        metadata (dict): Metadata extracted from the file, including ZSensor values for depth adjustment.
        seismic_data (DataFrame): DataFrame containing seismic data (S1 and S2).
        output_dir (str): The directory where the files will be saved.
        stack (bool): Whether the data is stacked (True) or non-stacked (False).
    """
    # Set folder name based on whether data is stacked or individual
    folder_name = "Stacked VS" if stack else "Individual VS"
    export_dir = os.path.join(output_dir, folder_name)
    os.makedirs(export_dir, exist_ok=True)

    # Prepare header row based on metadata
    z_sensor1 = metadata.get('ZSensor1', 'Unknown')
    z_sensor2 = metadata.get('ZSensor2', 'Unknown')
    header_s1 = f"Tempo(ms); Amplitude(cm/s) Valores provisorios da Sismica; Profundidade[m] ={metadata.get('ZSensor1', 'Unknown')}\n"
    header_s2 = f"Tempo(ms); Amplitude(cm/s) Valores provisorios da Sismica; Profundidade[m] ={metadata.get('ZSensor2', 'Unknown')}\n"
    
    # Use the correct delay from ExternalTrigger and SampleTime
    delay = metadata.get('Delay', 0)
    sample_time = int(metadata.get('SampleTime', 0))
    
    # Function to create the time column
    def create_time_column(row_num):
        return ((sample_time * row_num) / 1000) + delay # Convert to milliseconds

    # Create file names for inverted S1 and S2 within the appropriate folder, with the "B" suffix
    file_s1 = os.path.join(export_dir, f"{file_number_str}-S1-B.txt")
    file_s2 = os.path.join(export_dir, f"{file_number_str}-S2-B.txt")
    
    # Save inverted data for S1 with ZSensor1 in the header
    with open(file_s1, 'w') as f_s1:
        f_s1.write(header_s1)
        for idx, row in seismic_data.iterrows():
            time_ms = create_time_column(idx)
            inverted_s1 = invert_signal(row['S1'])  # Apply inversion
            f_s1.write(f"{time_ms:.3f}\t{inverted_s1}\n")
    
    # Save inverted data for S2 with ZSensor2 in the header
    with open(file_s2, 'w') as f_s2:
        f_s2.write(header_s2)
        for idx, row in seismic_data.iterrows():
            time_ms = create_time_column(idx)
            inverted_s2 = invert_signal(row['S2'])  # Apply inversion
            f_s2.write(f"{time_ms:.3f}\t{inverted_s2}\n")
    
    return file_s1, file_s2


def stack_signals_by_depth(files_by_depth):
    """
    Stack the signals from multiple files that are at the same depth.

    Args:
        files_by_depth (dict): A dictionary where keys are depths (ZRef or similar),
                               and values are lists of file paths at that depth.

    Returns:
        dict: A dictionary where keys are depths and values are tuples containing
              metadata and DataFrames for both normal and inverted stacked signals for each depth.
    """
    stacked_signals = {}

    for depth, file_list in files_by_depth.items():
        signal_s1_stack = None
        signal_s2_stack = None
        metadata = None

        for file_path in file_list:
            # Parse the seismic file to get metadata and signals
            metadata, seismic_data = parse_seismic_file(file_path)

            if signal_s1_stack is None:
                # Initialize the stack with the first file's data
                signal_s1_stack = seismic_data['S1'].astype(float).copy()
                signal_s2_stack = seismic_data['S2'].astype(float).copy()
            else:
                # Add the signals from the current file to the stack
                signal_s1_stack += seismic_data['S1'].astype(float)
                signal_s2_stack += seismic_data['S2'].astype(float)

        # Average the stacked signals by dividing by the number of files
        signal_s1_stack /= len(file_list)
        signal_s2_stack /= len(file_list)

        # Invert the stacked signals
        inverted_s1_stack = signal_s1_stack.apply(invert_signal)
        inverted_s2_stack = signal_s2_stack.apply(invert_signal)

        # Create DataFrames for both normal and inverted stacked signals
        stacked_df_normal = pd.DataFrame({
            'S1': signal_s1_stack,
            'S2': signal_s2_stack
        })

        stacked_df_inverted = pd.DataFrame({
            'S1': inverted_s1_stack,
            'S2': inverted_s2_stack
        })

        # Store both normal and inverted stacked signals
        stacked_signals[depth] = (metadata, stacked_df_normal, stacked_df_inverted)

    return stacked_signals


def export_nonstacked_button_callback():
    """
    Callback function to handle exporting non-stacked SDM files with both normal and inverted signals.
    """
    directory = export_directory  # Use the already selected export directory

    # Process non-stacked SDM files directly (no need for sdmt_tests)
    process_sdm_files(directory, stack=False)  # stack=False means export non-stacked files
    
    # Now process non-stacked SDM files for inverted signals
    process_inverted_sdm_files(directory, stack=False)  # Export inverted signals as non-stacked

    
def sdm_stack_export_button_callback():
    """
    Callback function to handle exporting stacked SDM files (both normal and inverted).
    """
    directory = export_directory  # Use the already selected export directory
    # sdmt_tests = read_sdmt_xml(os.path.join(directory, 'SDMT.xml'))  # Assuming 'SDMT.xml' exists

    # Process stacked SDM files
    process_sdm_files(directory, stack=True)#, sdmt_tests, stack=True)  # stack=True means export stacked files
    

def export_vp_button_callback():
    """
    Callback function to handle exporting non-stacked VP files with the letter C.
    """
    directory = export_directory  # Use the already selected export directory
    # sdmt_tests = read_sdmt_xml(os.path.join(directory, 'SDMT.xml'))  # Assuming 'SDMT.xml' exists

    # Ensure the process_vp_files function is correctly called for the VP data
    process_vp_files(directory)#, sdmt_tests)


# Button to show dataframes info
dataframes_info_button = ttk.Button(
    root,
    text='Show DataFrames Info',
    command=show_dataframes_info
)
dataframes_info_button.pack(expand=True)

# Button to show transformed dataframes info
transformed_dataframes_info_button = ttk.Button(
    root,
    text='Show Transformed DataFrames Info',
    command=show_transformed_dataframes_info
)
transformed_dataframes_info_button.pack(expand=True)

# New button for exporting non-stacked SDM data (both normal and inverted)
export_nonstacked_button = ttk.Button(
    root,
    text='Export Non-stacked SDM Data',
    command=export_nonstacked_button_callback
)
export_nonstacked_button.pack(expand=True)

# Button for exporting stacked SDM data (both normal and inverted)
sdm_stack_export_button = ttk.Button(
    root,
    text='Export Stacked SDM Data',
    command=sdm_stack_export_button_callback
)
sdm_stack_export_button.pack(expand=True)

# Button for exporting non-stacked VP data (with the letter C)
export_vp_button = ttk.Button(
    root,
    text='Export Non-stacked VP Data',
    command=export_vp_button_callback
)
export_vp_button.pack(expand=True)

# run the application
root.mainloop()
