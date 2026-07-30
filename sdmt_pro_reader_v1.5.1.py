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
import re
from typing import Optional, Union, Dict, List, Tuple

from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font

# ---------------------------------------------------------------------------
# Config & Logging
# ---------------------------------------------------------------------------
CONFIG = {
    "sdmt_report_filename": "SDMT Report.xlsx",
    "initial_dir": r"U:\1 - Projects\2024",
    "excel_export_transformed": "sdmt_import.xlsx",
    "excel_export_database": "sdmt_database_import.xlsx",
}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)

# ---------------------------------------------------------------------------
# Tk root window
# ---------------------------------------------------------------------------
root = tk.Tk()
root.title('Tkinter Open File Dialog')
root.resizable(False, False)
root.geometry('300x360')
root.lift()
root.attributes('-topmost', True)
root.after_idle(root.attributes, '-topmost', False)

# ---------------------------------------------------------------------------
# Schema / alias utilities
# ---------------------------------------------------------------------------
def normalize_name(name: str) -> str:
    """Normalize column names (lowercase, strip spaces, unify Greek letters/symbols)."""
    if not isinstance(name, str):
        return name
    s = name
    s = s.replace('Δ', 'delta').replace('δ', 'delta')
    s = s.replace('γ', 'gamma')
    s = s.replace("σ' v", "sigma_v_eff").replace("σ′ v", "sigma_v_eff")
    s = s.replace("σ' h", "sigma_h_eff").replace("σ′ h", "sigma_h_eff")
    s = s.translate(str.maketrans('₀₁₂₃₄₅₆₇₈₉', '0123456789'))  # subscript
    s = s.replace('φ', 'phi')
    s = s.replace(' ', '')
    s = s.lower()
    s = re.sub(r'[^a-z0-9_]', '', s)
    return s

def select_and_rename(
    df: pd.DataFrame,
    required: Dict[str, List[str]],
    optional: Optional[Dict[str, List[str]]] = None
) -> pd.DataFrame:
    """
    required/optional: maps canonical names -> list of acceptable source names (post-normalization).
    Missing columns are added as NaN. Returns a dataframe with only canonical columns.
    """
    if df is None or df.empty:
        cols = list(required.keys()) + (list(optional.keys()) if optional else [])
        return pd.DataFrame(columns=cols)

    norm_map = {normalize_name(c): c for c in df.columns}

    def resolve(candidates):
        for cand in candidates:
            n = normalize_name(cand)
            if n in norm_map:
                return norm_map[n]
        return None

    keep = {}
    for canon, cands in required.items():
        src = resolve(cands)
        keep[canon] = df[src] if src else pd.Series([np.nan] * len(df))

    if optional:
        for canon, cands in optional.items():
            src = resolve(cands)
            keep[canon] = df[src] if src else pd.Series([np.nan] * len(df))

    out = pd.DataFrame(keep)
    return out

def coerce_numeric(df, cols):
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce')
    return df

def coerce_datetime(df, cols):
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors='coerce')
    return df

# ---------------------------------------------------------------------------
# File selection
# ---------------------------------------------------------------------------
def select_file(callback):
    filetypes = (('SDMT files', '*.sdmt'), ('All files', '*.*'))
    filename = fd.askopenfilename(
        title='Open a file',
        initialdir=CONFIG["initial_dir"],
        filetypes=filetypes
    )
    if filename:
        callback(filename)

# ---------------------------------------------------------------------------
# Pipeline entry
# ---------------------------------------------------------------------------
def process_filename(filename):
    global dataframes, export_directory, transformed_dataframes
    global POINT_df, DMTG_df, DMTT_df, Seismic_Downhole_df, Vs_Data_df
    global Vs_Parameters_Setup_df, Vs_Parameters_Data_df

    directory = os.path.dirname(filename)
    export_directory = directory

    # Parse SDMT.xml to get test folders
    sdmt_tests = read_sdmt_xml(filename)

    # TEST.dat -> test_meta_df
    test_filepaths = get_test_filepaths(directory, sdmt_tests)
    test_meta_df = process_test_files(test_filepaths)

    # DMT.dat -> dmt_setup_df + dati_dmt_df
    dmt_filepaths = get_dmt_filepaths(directory, sdmt_tests)
    dmt_setup_df, dati_dmt_df = process_dmt_files(dmt_filepaths)

    # Calibrazione.dat -> calib_df (DeltaA/DeltaB)
    calib_filepaths = get_calib_filepaths(directory, sdmt_tests)
    calib_df = process_calib_files(calib_filepaths)

    # SDMT Report.xlsx -> s_wave, p_wave, dmt_data (excel summaries)
    s_wave_data_df, p_wave_data_df, dmt_data_df = parse_sdmt_report_xlsx(
        directory, CONFIG["sdmt_report_filename"]
    )

    # VS -> TI.dat -> vs_setup_df
    vs_setup_df = process_vs_files(directory, sdmt_tests)

    # SDM files info (non-stacked export is done via button)
    vs_sdm_info_df = process_sdm_files(directory, stack=False)
    if vs_sdm_info_df.empty:
        vs_sdm_info_df = pd.DataFrame(columns=['POINT_ID', 'SdmNumber'])

    # Transform & join
    dataframes = [
        test_meta_df, dmt_setup_df, dati_dmt_df, dmt_data_df, s_wave_data_df,
        p_wave_data_df, vs_setup_df, vs_sdm_info_df, calib_df
    ]
    transformed_dataframes = transform_raw_data(dataframes)

    # Shorthand
    tm = transformed_dataframes["test_meta_df"]
    dx = transformed_dataframes["dmt_data_df"]
    cb = transformed_dataframes["calib_df"]
    sw = transformed_dataframes["s_wave_data_df"]

    # Build database tables (with schema utilities)
    POINT_df = create_POINT_df(tm)
    DMTG_df  = create_DMTG_df(tm, cb)
    DMTT_df  = create_DMTT_df(dx, tm)
    Seismic_Downhole_df    = create_Seismic_Downhole_df(tm)
    Vs_Data_df             = create_Vs_Data_df(sw, tm)
    Vs_Parameters_Setup_df = create_Vs_Parameters_Setup_df(tm)
    Vs_Parameters_Data_df  = create_Vs_Parameters_Data_df(sw, tm)

# ---------------------------------------------------------------------------
# Info windows
# ---------------------------------------------------------------------------
def show_combined_scrollable_info(title, **dataframes):
    info_window = tk.Toplevel(root)
    info_window.title(title)
    info_window.geometry('800x900')

    text_area = st.ScrolledText(info_window, wrap=tk.WORD, width=100, height=48)

    with pd.option_context('display.max_rows', None, 'display.max_columns', None, 'display.max_colwidth', None):
        for df_name, df in dataframes.items():
            text_area.insert(tk.INSERT, f"{df_name}:\n")
            text_area.insert(tk.INSERT, df.to_string())
            text_area.insert(tk.INSERT, "\n\n")

    text_area.configure(state='disabled')
    text_area.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

def show_scrollable_info(title, *dataframes):
    info_window = tk.Toplevel(root)
    info_window.title(title)
    info_window.geometry('800x600')

    text_area = st.ScrolledText(info_window, wrap=tk.WORD, width=100, height=32)
    with pd.option_context('display.max_rows', None, 'display.max_columns', None, 'display.max_colwidth', None):
        for df in dataframes:
            text_area.insert(tk.INSERT, df.to_string())
            text_area.insert(tk.INSERT, "\n\n")
    text_area.configure(state='disabled')
    text_area.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

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

def show_transformed_dataframes_info():
    global transformed_dataframes
    if not transformed_dataframes:
        return
    d = transformed_dataframes
    show_scrollable_info(
        'Transformed Dataframes',
        d["test_meta_df"], d["dmt_setup_df"], d["dati_dmt_df"], d["dmt_data_df"],
        d["s_wave_data_df"], d["p_wave_data_df"], d["vs_setup_df"],
        d["vs_sdm_info_df"], d["calib_df"]
    )

# ---------------------------------------------------------------------------
# Excel export
# ---------------------------------------------------------------------------
def export_dataframes_to_excel(directory, **dataframes):
    """Database-shaped export (tables)."""
    file_path = os.path.join(directory, CONFIG["excel_export_database"])
    wb = Workbook()
    for df_name, df in dataframes.items():
        safe_name = df_name.replace('_df', '')[:31].replace(':', '-').replace('/', '-')
        ws = wb.create_sheet(title=safe_name)
        for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True)):
            for c_idx, value in enumerate(row, 1):
                cell = ws.cell(row=r_idx + 1, column=c_idx, value=value)
                if r_idx == 0:
                    cell.font = Font(bold=True)
                column_letter = ws.cell(row=1, column=c_idx).column_letter
                current_width = ws.column_dimensions[column_letter].width or 0
                ws.column_dimensions[column_letter].width = max(current_width, min(len(str(value)) + 2, 60))
    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']
    wb.save(file_path)
    print(f"Data successfully exported to {file_path}")
    sys.stdout.flush()

def export_to_excel(directory, dataframes):
    """Transformed export (debug/inspection). Accepts dict or legacy list."""
    file_path = os.path.join(directory, CONFIG["excel_export_transformed"])
    wb = Workbook()

    if isinstance(dataframes, dict):
        ws_mapping = dataframes
    else:
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
        if len(dataframes) >= 9:
            ws_mapping['calib_df'] = dataframes[8]

    for sheet_name, df in ws_mapping.items():
        safe_name = sheet_name.replace('_df', '')[:31].replace(':', '-').replace('/', '-')
        ws = wb.create_sheet(title=safe_name)
        for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True)):
            for c_idx, value in enumerate(row, 1):
                cell = ws.cell(row=r_idx + 1, column=c_idx, value=value)
                if r_idx == 0:
                    cell.font = Font(bold=True)
                column_letter = ws.cell(row=1, column=c_idx).column_letter
                current_width = ws.column_dimensions[column_letter].width or 0
                ws.column_dimensions[column_letter].width = max(current_width, min(len(str(value)) + 2, 60))

    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']

    wb.save(file_path)
    print(f"Data successfully exported to {file_path}")
    sys.stdout.flush()

# ---------------------------------------------------------------------------
# Buttons / callbacks
# ---------------------------------------------------------------------------
def database_export_button_callback():
    global export_directory, POINT_df, DMTG_df, DMTT_df, Seismic_Downhole_df
    global Vs_Data_df, Vs_Parameters_Setup_df, Vs_Parameters_Data_df
    try:
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
    except Exception:
        logging.exception("Database export failed")

def export_button_callback():
    global export_directory, transformed_dataframes
    try:
        export_to_excel(export_directory, transformed_dataframes)
    except Exception:
        logging.exception("Transformed export failed")

open_button = ttk.Button(root, text='Open a File', command=lambda: select_file(process_filename))
open_button.pack(expand=True, pady=2)

database_export_button = ttk.Button(root, text='Database Export', command=database_export_button_callback)
database_export_button.pack(expand=True, pady=2)

export_button = ttk.Button(root, text='Export Transformed (Debug)', command=export_button_callback)
export_button.pack(expand=True, pady=2)

# ---------------------------------------------------------------------------
# Readers / parsers
# ---------------------------------------------------------------------------
def read_sdmt_xml(filename):
    try:
        with open(filename, 'r', encoding='utf-8') as f:
            xml_string = f.read()
    except FileNotFoundError:
        print(f"Error: File not found: {filename}")
        return ()
    return get_tests_from_xml(xml_string)

def get_tests_from_xml(xml_string):
    root = ET.fromstring(xml_string)
    tests_element = root.find('Tests')
    if tests_element is None:
        return ()
    tests = tuple(test.get('Folder') for test in tests_element.findall('Test'))
    return tests

def get_test_filepaths(directory, sdmt_tests):
    return tuple((test_name, f"{directory}/{test_name}/TEST.dat") for test_name in sdmt_tests)

def get_dmt_filepaths(directory, sdmt_tests):
    return tuple((test_name, f"{directory}/{test_name}/DMT.dat") for test_name in sdmt_tests)

def process_test_files(test_filepaths):
    records = []
    for subdirectory, filename in test_filepaths:
        try:
            with open(filename, 'r', encoding='utf-8') as f:
                xml_string = f.read()
            record = parse_test_xml(xml_string)
            record['TEST_N'] = subdirectory
            records.append(record)
        except FileNotFoundError:
            print(f"Error: File not found: {filename}")
    df = pd.DataFrame(records).rename(columns={'Name': 'POINT_ID'})
    return df

def parse_test_xml(xml_string):
    root = ET.fromstring(xml_string)
    test_data = {}
    for child in root:
        if child.tag not in ('GPS', 'U', 'Gamma'):
            test_data[child.tag] = child.text
        elif child.tag == 'GPS':
            test_data['Latitude'] = child.get('Latitude')
            test_data['Longitude'] = child.get('Longitude')
        elif child.tag in ('U', 'Gamma'):
            z_value = child.get('Z')
            unit_value = child.get('U') or child.get('Gamma')
            if z_value and unit_value:
                test_data[child.tag] = f"{z_value} {unit_value}"
    return test_data

def process_dmt_files(test_filepaths):
    dmt_setup_records = []
    dati_dmt_records = []
    for subdirectory, filename in test_filepaths:
        try:
            with open(filename, 'r', encoding='utf-8') as f:
                xml_string = f.read()
            setup_record, dati_dmt_record = parse_dmt_xml(xml_string, subdirectory)
            setup_record['TEST_N'] = subdirectory
            for record in dati_dmt_record:
                record['TEST_N'] = subdirectory
            dmt_setup_records.append(setup_record)
            dati_dmt_records.extend(dati_dmt_record)
        except FileNotFoundError:
            print(f"Error: File not found: {filename}")
    dmt_setup_df = pd.DataFrame(dmt_setup_records)
    dati_dmt_df = pd.DataFrame(dati_dmt_records)
    return dmt_setup_df, dati_dmt_df

def parse_dmt_xml(xml_string, test_name):
    if not xml_string:
        return None, None
    root = ET.fromstring(xml_string)
    dmt_data = {'POINT_ID': test_name}
    dati_dmt_data = []
    for child in root:
        if child.tag != 'DatiDMT':
            dmt_data[child.tag] = child.text
    for element in root.find('DatiDMT'):
        record = {'POINT_ID': test_name, 'Z': element.get('Z'),
                  'A': element.get('A'), 'B': element.get('B')}
        dati_dmt_data.append(record)
    return dmt_data, dati_dmt_data

def get_calib_filepaths(directory, sdmt_tests):
    return tuple((test_name, os.path.join(directory, test_name, 'Calibrazione.dat')) for test_name in sdmt_tests)

def parse_calib_xml(xml_string, test_name):
    root = ET.fromstring(xml_string)
    items = root.findall('.//Calibrazioni/Item')
    if not items:
        return {'POINT_ID': test_name, 'TEST_N': test_name, 'DeltaA': None, 'DeltaB': None}
    def _dt(x): return x.get('Date') or ''
    latest = max(items, key=_dt)
    return {'POINT_ID': test_name, 'TEST_N': test_name,
            'DeltaA': latest.get('DeltaA'), 'DeltaB': latest.get('DeltaB')}

def process_calib_files(calib_filepaths):
    recs = []
    for test_name, filename in calib_filepaths:
        try:
            with open(filename, 'r', encoding='utf-8') as f:
                xml_string = f.read()
            recs.append(parse_calib_xml(xml_string, test_name))
        except FileNotFoundError:
            print(f"Warning: missing {filename}")
        except ET.ParseError as e:
            print(f"Error parsing {filename}: {e}")
    return pd.DataFrame(recs, columns=['POINT_ID', 'TEST_N', 'DeltaA', 'DeltaB'])

def parse_sdmt_report_xlsx(directory, sdmt_report_filename):
    file_path = f'{directory}/{sdmt_report_filename}'
    try:
        xls = pd.ExcelFile(file_path)
    except FileNotFoundError:
        print(f"Error: File not found: {file_path}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    s_wave_df = pd.DataFrame()
    p_wave_df = pd.DataFrame()
    dmt_df = pd.DataFrame()

    for sheet_name in xls.sheet_names:
        parts = sheet_name.rsplit(' - ', 1)
        if len(parts) == 2:
            test_name, test_type = parts[0], parts[1]
            if test_type == 'P-wave':
                df = pd.read_excel(xls, sheet_name=sheet_name, header=3, usecols=range(4), skiprows=[4])
                df['POINT_ID'] = test_name
                p_wave_df = pd.concat([p_wave_df, df], ignore_index=True)
            elif test_type in ['S-wave', 'DMT']:
                df = pd.read_excel(xls, sheet_name=sheet_name, header=2, skiprows=[3])
                df['POINT_ID'] = test_name
                if test_type == 'S-wave':
                    s_wave_df = pd.concat([s_wave_df, df], ignore_index=True)
                elif test_type == 'DMT':
                    dmt_df = pd.concat([dmt_df, df], ignore_index=True)

    for df in (s_wave_df, p_wave_df, dmt_df):
        if 'Zabs' in df.columns:
            df.drop(columns=['Zabs'], inplace=True)
    return s_wave_df, p_wave_df, dmt_df

def process_vs_files(directory, sdmt_tests):
    vs_records = []
    for test_folder in sdmt_tests:
        point_id = test_folder
        vs_directory = os.path.join(directory, test_folder, 'VS')
        for root_dir, _, files in os.walk(vs_directory):
            for file in files:
                if file == 'TI.dat':
                    file_path = os.path.join(root_dir, file)
                    try:
                        with open(file_path, 'r', encoding='utf-8') as f:
                            xml_string = f.read()
                        vs_records.extend(parse_vs_xml(xml_string, point_id))
                    except FileNotFoundError:
                        print(f"Error: File not found: {file_path}")
    vs_setup_df = pd.DataFrame(vs_records)
    return vs_setup_df if not vs_setup_df.empty else pd.DataFrame(
        columns=['POINT_ID','Z','Selection','FileNumber','DeltaS','DeltaT',
                 'Filter_Code','Filter_P1','Window_Code','Window_P1','Window_P2',
                 'Rephase_Code','Rephase_P1'])

def parse_vs_xml(xml_string, point_id):
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
                'DeltaT': seismogram.get('DeltaT')
            }
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

# ---------------------------------------------------------------------------
# SDM/VP processing & export
# ---------------------------------------------------------------------------
def process_sdm_files(directory, stack=False):
    sdm_info_records = []
    files_by_depth = {}

    for test_folder in os.listdir(directory):
        vs_directory = os.path.join(directory, test_folder, 'VS')
        if not os.path.isdir(vs_directory):
            continue

        for root_dir, _, files in os.walk(vs_directory):
            for file in files:
                if file.endswith('.sdm'):
                    file_path = os.path.join(root_dir, file)
                    try:
                        sdm_number = int(os.path.splitext(file)[0])
                        metadata, seismic_data = parse_seismic_file(file_path)
                        if stack:
                            depth = metadata.get('ZRef', 'Unknown')
                            files_by_depth.setdefault(depth, []).append(file_path)
                        else:
                            save_seismic_data_corrected(sdm_number, metadata, seismic_data, root_dir, stack=False)
                        sdm_info_records.append({'POINT_ID': test_folder, 'SdmNumber': sdm_number, 'Filename': file})
                    except FileNotFoundError:
                        print(f"Error: File not found: {file_path}")

    if sdm_info_records:
        vs_sdm_info_df = pd.DataFrame(sdm_info_records)
        if 'POINT_ID' in vs_sdm_info_df.columns and 'SdmNumber' in vs_sdm_info_df.columns:
            vs_sdm_info_df = vs_sdm_info_df.sort_values(by=['POINT_ID', 'SdmNumber'])
        else:
            print("Warning: 'POINT_ID' or 'SdmNumber' columns are missing in vs_sdm_info_df.")
    else:
        vs_sdm_info_df = pd.DataFrame(columns=['POINT_ID', 'Filename', 'SdmNumber'])

    return vs_sdm_info_df

def process_inverted_sdm_files(directory, stack=False):
    for test_folder in os.listdir(directory):
        vs_directory = os.path.join(directory, test_folder, 'VS')
        if not os.path.isdir(vs_directory):
            continue
        for root_dir, _, files in os.walk(vs_directory):
            for file in files:
                if file.endswith('.sdm'):
                    file_path = os.path.join(root_dir, file)
                    try:
                        sdm_number = int(os.path.splitext(file)[0])
                        metadata, seismic_data = parse_seismic_file(file_path)
                        save_seismic_data_inverted(sdm_number, metadata, seismic_data, root_dir, stack=False)
                    except FileNotFoundError:
                        print(f"Error: File not found: {file_path}")
                        sys.stdout.flush()

def process_vp_files(directory):
    sdm_info_records = []
    for test_folder in os.listdir(directory):
        vp_directory = os.path.join(directory, test_folder, 'VP')
        if not os.path.isdir(vp_directory):
            continue
        for root_dir, _, files in os.walk(vp_directory):
            for file in files:
                if file.endswith('.sdm'):
                    file_path = os.path.join(root_dir, file)
                    try:
                        sdm_number = int(os.path.splitext(file)[0])
                        metadata, seismic_data = parse_seismic_file(file_path)
                        save_vp_data(sdm_number, metadata, seismic_data, root_dir)
                    except FileNotFoundError:
                        print(f"Error: File not found: {file_path}")
                        sys.stdout.flush()

    if sdm_info_records:
        vp_sdm_info_df = pd.DataFrame(sdm_info_records).sort_values(by=['POINT_ID', 'SdmNumber'])
    else:
        vp_sdm_info_df = pd.DataFrame(columns=['POINT_ID', 'Filename', 'SdmNumber'])
    return vp_sdm_info_df

# ---------------------------------------------------------------------------
# Transform (returns dict)
# ---------------------------------------------------------------------------
def transform_raw_data(dfs_in: Union[List[pd.DataFrame], Tuple[pd.DataFrame, ...]]):
    (test_meta_df, dmt_setup_df, dati_dmt_df, dmt_data_df, s_wave_data_df,
     p_wave_data_df, vs_setup_df, vs_sdm_info_df, calib_df) = dfs_in

    if not test_meta_df.empty:
        test_meta_lookup = test_meta_df.set_index('TEST_N')['POINT_ID'].to_dict()
        reverse_test_meta_lookup = test_meta_df.set_index('POINT_ID')['TEST_N'].to_dict()
    else:
        test_meta_lookup = {}
        reverse_test_meta_lookup = {}

    if not dmt_setup_df.empty:
        dmt_setup_df['POINT_ID'] = dmt_setup_df['TEST_N'].map(test_meta_lookup)
    if not dati_dmt_df.empty:
        dati_dmt_df['POINT_ID'] = dati_dmt_df['TEST_N'].map(test_meta_lookup)

    if not dmt_data_df.empty:
        dmt_data_df['TEST_N'] = dmt_data_df['POINT_ID'].map(reverse_test_meta_lookup)
    if not s_wave_data_df.empty:
        s_wave_data_df['TEST_N'] = s_wave_data_df['POINT_ID'].map(reverse_test_meta_lookup)
    if not p_wave_data_df.empty:
        p_wave_data_df['TEST_N'] = p_wave_data_df['POINT_ID'].map(reverse_test_meta_lookup)

    if not vs_setup_df.empty:
        vs_setup_df['TEST_N'] = vs_setup_df['POINT_ID']
        vs_setup_df['POINT_ID'] = vs_setup_df['TEST_N'].map(test_meta_lookup)
    if not vs_sdm_info_df.empty:
        vs_sdm_info_df['TEST_N'] = vs_sdm_info_df['POINT_ID']
        vs_sdm_info_df['POINT_ID'] = vs_sdm_info_df['TEST_N'].map(test_meta_lookup)

    if not test_meta_df.empty and not dmt_setup_df.empty:
        test_meta_df = test_meta_df.merge(
            dmt_setup_df, on='POINT_ID', how='left', suffixes=('', '_dmt_setup')
        )

    if not vs_sdm_info_df.empty and not test_meta_df.empty:
        vs_sorted = vs_sdm_info_df.sort_values(by=['POINT_ID', 'SdmNumber'])
        vs_first = vs_sorted.drop_duplicates(subset=['POINT_ID'], keep='first')
        test_meta_df = test_meta_df.merge(
            vs_first[['POINT_ID', 'SdmNumber']], on='POINT_ID', how='left'
        )

    if calib_df is not None and not calib_df.empty:
        calib_df['POINT_ID'] = calib_df['TEST_N'].map(test_meta_lookup)
        if not test_meta_df.empty:
            test_meta_df = test_meta_df.merge(
                calib_df[['POINT_ID', 'DeltaA', 'DeltaB']], on='POINT_ID', how='left'
            )

    return {
        "test_meta_df": test_meta_df,
        "dmt_setup_df": dmt_setup_df,
        "dati_dmt_df": dati_dmt_df,
        "dmt_data_df": dmt_data_df,
        "s_wave_data_df": s_wave_data_df,
        "p_wave_data_df": p_wave_data_df,
        "vs_setup_df": vs_setup_df,
        "vs_sdm_info_df": vs_sdm_info_df,
        "calib_df": calib_df,
    }

# ---------------------------------------------------------------------------
# Builders (with schema layer)
# ---------------------------------------------------------------------------
def create_POINT_df(test_meta_df):
    req = {
        'POINT_ID': ['POINT_ID'],
        'Date': ['Date','StartDate'],
        'EndDate': ['EndDate'],
        'Remark': ['Remark','Remarks'],
        'Operator': ['Operator','Crew'],
        'Latitude': ['Latitude'],
        'Longitude': ['Longitude'],
    }
    sel = select_and_rename(test_meta_df, req, optional={})
    POINT_df = sel.rename(columns={
        'Date':'StartDate',
        'Remark':'Remarks',
        'Operator':'Crew'
    })
    coerce_datetime(POINT_df, ['StartDate','EndDate'])
    coerce_numeric(POINT_df, ['Latitude','Longitude'])
    POINT_df['POINT_ID'] = POINT_df['POINT_ID'].astype(str)
    if 'Remarks' in POINT_df.columns: POINT_df['Remarks'] = POINT_df['Remarks'].fillna('').astype(str)
    if 'Crew' in POINT_df.columns: POINT_df['Crew'] = POINT_df['Crew'].fillna('').astype(str)
    return POINT_df

def create_DMTG_df(test_meta_df, calib_df):
    base = select_and_rename(
        test_meta_df,
        required={
            'POINT_ID':['POINT_ID'],
            'Date':['Date'],
            'Zw':['Zw','GroundWaterLevel'],
            'IsDmtActive':['IsDmtActive','ITEM']
        },
        optional={}
    ).rename(columns={'Zw':'GroundWaterLevel','IsDmtActive':'ITEM'})

    if calib_df is not None and not calib_df.empty:
        base = base.merge(calib_df[['POINT_ID','DeltaA','DeltaB']], on='POINT_ID', how='left')
    else:
        base['DeltaA'] = np.nan
        base['DeltaB'] = np.nan

    DMTG_df = base.rename(columns={'DeltaA':'BladeA','DeltaB':'BladeB'})
    coerce_datetime(DMTG_df, ['Date'])
    coerce_numeric(DMTG_df, ['GroundWaterLevel','BladeA','BladeB'])
    DMTG_df['POINT_ID'] = DMTG_df['POINT_ID'].astype(str)
    DMTG_df['ITEM'] = DMTG_df['ITEM'].astype(str)
    DMTG_df = DMTG_df[DMTG_df['ITEM'] != "0"]
    cols = ['POINT_ID','ITEM'] + [c for c in DMTG_df.columns if c not in ['POINT_ID','ITEM']]
    return DMTG_df[cols]

def create_DMTT_df(dmt_data_df, test_meta_df):
    if dmt_data_df.empty:
        return pd.DataFrame(columns=[
            'POINT_ID','ITEM','REFERENCE','TOP','DeltaA','DeltaB',
            'APressure','BPressure','CPressure','BulkUnitWeight',
            'EffectiveVerticalStress','u0','ACorrected','BCorrected',
            'CCorrected','MaterialIndex','Dilatometer','HorizontalStress',
            'PoreIndex','VerticalDrainedConstrainedModulus',
            'UndrainedShearStrength','FrictionAngle','LateralEarthPressure',
            'EffectiveHorizontalStress','OCR','Description'
        ])

    required = {
        'POINT_ID': ['POINT_ID'],
        'Z': ['Z','Depth','Top'],
        'DeltaA': ['Δ A','ΔA','DeltaA','Delta A'],
        'DeltaB': ['Δ B','ΔB','DeltaB','Delta B'],
        'A': ['A'], 'B': ['B'], 'C': ['C'],
    }
    optional = {
        'gamma': ['γ','gamma'],
        'sigma_v_eff': ["σ' v","σ′ v","sigma_v_eff"],
        'u0': ['U₀','U0','u0'],
        'P0': ['P₀','P0'],
        'P1': ['P₁','P1'],
        'P2': ['P₂','P2'],
        'Id': ['Id'], 'Ed': ['Ed'], 'Kd': ['Kd'], 'Ud': ['Ud'],
        'M': ['M'], 'Su': ['Su'], 'phi': ['φ','phi'],
        'K0': ['K₀','K0','k0'],
        'sigma_h_eff': ["σ' h","σ′ h","sigma_h_eff"],
        'OCR': ['OCR'],
        'Soil description': ['Soil description','Description'],
    }

    sel = select_and_rename(dmt_data_df, required, optional)
    DMTT_df = sel.rename(columns={
        'Z':'TOP','A':'APressure','B':'BPressure','C':'CPressure',
        'gamma':'BulkUnitWeight','sigma_v_eff':'EffectiveVerticalStress',
        'P0':'ACorrected','P1':'BCorrected','P2':'CCorrected',
        'Id':'MaterialIndex','Ed':'Dilatometer','Kd':'HorizontalStress',
        'Ud':'PoreIndex','M':'VerticalDrainedConstrainedModulus',
        'Su':'UndrainedShearStrength','phi':'FrictionAngle',
        'K0':'LateralEarthPressure','sigma_h_eff':'EffectiveHorizontalStress',
        'Soil description':'Description'
    })

    item_lookup = test_meta_df.set_index('POINT_ID')['IsDmtActive'].to_dict() if 'IsDmtActive' in test_meta_df.columns else {}
    DMTT_df['ITEM'] = DMTT_df['POINT_ID'].map(item_lookup)
    DMTT_df['REFERENCE'] = "Standard Marchetti formulas from SDMT Pro"

    for c in DMTT_df.columns:
        if c in ['POINT_ID','REFERENCE','Description','ITEM']:
            DMTT_df[c] = DMTT_df[c].astype(str)
        else:
            DMTT_df[c] = pd.to_numeric(DMTT_df[c], errors='coerce')
    return DMTT_df

def create_Seismic_Downhole_df(test_meta_df):
    sel = select_and_rename(
        test_meta_df,
        required={
            'POINT_ID':['POINT_ID'],
            'Date':['Date'],
            'IsVsActive':['IsVsActive','ITEM'],
            'DistHammer':['DistHammer','SourceOffset_Vs'],
            'Zw':['Zw','GroundWaterLevel']
        },
        optional={}
    ).rename(columns={'IsVsActive':'ITEM','DistHammer':'SourceOffset_Vs','Zw':'GroundWaterLevel'})

    sel['POINT_ID'] = sel['POINT_ID'].astype(str)
    sel['ITEM'] = sel['ITEM'].astype(str)
    sel = sel[sel['ITEM'] != "0"]
    coerce_datetime(sel, ['Date'])
    coerce_numeric(sel, ['GroundWaterLevel','SourceOffset_Vs'])
    cols = ['POINT_ID','ITEM'] + [c for c in sel.columns if c not in ['POINT_ID','ITEM']]
    return sel[cols]

def create_Vs_Data_df(s_wave_data_df, test_meta_df):
    sel = select_and_rename(
        s_wave_data_df,
        required={
            'POINT_ID':['POINT_ID'],
            'Z':['Z','Depth','Top'],
            'Vs':['Vs','VsAvg'],
            'Repeatability':['Repeatability'],
            'VarCoefficient':['Var Coefficient','VarCoefficient'],
        },
        optional={}
    ).rename(columns={'Z':'TOP','Vs':'VsAvg'})

    sel = sel.dropna(subset=['VsAvg'])
    item_lookup = test_meta_df.set_index('POINT_ID')['IsVsActive'].to_dict() if 'IsVsActive' in test_meta_df.columns else {}
    sel['ITEM'] = sel['POINT_ID'].map(item_lookup)

    sel['POINT_ID'] = sel['POINT_ID'].astype(str)
    sel['Repeatability'] = sel['Repeatability'].astype(str)
    sel['ITEM'] = sel['ITEM'].astype(str)
    coerce_numeric(sel, ['TOP','VsAvg','VarCoefficient'])

    def _minmax(rep):
        try:
            vals = [float(x) for x in str(rep).split(',') if x.strip()!='']
            return (min(vals), max(vals)) if vals else (None, None)
        except Exception:
            return (None, None)
    lbub = sel['Repeatability'].apply(_minmax)
    sel['VsLB'] = lbub.apply(lambda t: t[0])
    sel['VsUB'] = lbub.apply(lambda t: t[1])

    cols = ['POINT_ID','ITEM'] + [c for c in sel.columns if c not in ['POINT_ID','ITEM']]
    return sel[cols]

def create_Vs_Parameters_Setup_df(test_meta_df):
    sel = select_and_rename(
        test_meta_df,
        required={
            'POINT_ID':['POINT_ID'],
            'Date':['Date'],
            'IsVsActive':['IsVsActive','ITEM'],
        },
        optional={}
    ).rename(columns={'IsVsActive':'ITEM'})

    sel = sel[sel['ITEM'].astype(str) != "0"]
    sel['REFERENCE'] = "Marchetti TI"
    sel['Methods'] = "Marchetti True Interval"

    cols = ['POINT_ID','ITEM','REFERENCE'] + [c for c in sel.columns if c not in ['POINT_ID','ITEM','REFERENCE']]
    return sel[cols]

def create_Vs_Parameters_Data_df(s_wave_data_df, test_meta_df):
    if s_wave_data_df.empty:
        return pd.DataFrame(columns=['POINT_ID','ITEM','REFERENCE','TOP','ShearModAvg','ShearModLB','ShearModUB','BulkDensity'])

    sel = select_and_rename(
        s_wave_data_df,
        required={
            'POINT_ID':['POINT_ID'],
            'Z':['Z','Depth','Top'],
            'ρ':['ρ','BulkDensity'],
            'G0':['G₀','G0','ShearModAvg'],
            'Repeatability':['Repeatability'],
        },
        optional={}
    ).rename(columns={'Z':'TOP','ρ':'BulkDensity','G0':'ShearModAvg'})

    sel = sel.dropna(subset=['ShearModAvg'])
    item_lookup = test_meta_df.set_index('POINT_ID')['IsVsActive'].to_dict() if 'IsVsActive' in test_meta_df.columns else {}
    sel['ITEM'] = sel['POINT_ID'].map(item_lookup)

    def calc_shear(rep, rho):
        try:
            vals = [float(x) for x in str(rep).split(',') if x.strip()!='']
            if not vals or pd.isna(rho):
                return (None, None)
            vs_lb, vs_ub = min(vals), max(vals)
            # Units: Pa = kg/(m*s^2); rho in kg/m^3 and Vs in m/s; result in MPa.
            return ((vs_lb**2)*float(rho)/1_000_000, (vs_ub**2)*float(rho)/1_000_000)
        except Exception:
            return (None, None)

    lbub = sel.apply(lambda r: calc_shear(r['Repeatability'], r['BulkDensity']), axis=1)
    sel['ShearModLB'] = lbub.apply(lambda t: t[0])
    sel['ShearModUB'] = lbub.apply(lambda t: t[1])
    sel.drop(columns=['Repeatability'], inplace=True)

    sel['REFERENCE'] = "Marchetti TI"
    cols = ['POINT_ID','ITEM','REFERENCE'] + [c for c in sel.columns if c not in ['POINT_ID','ITEM','REFERENCE']]
    return sel[cols]

# ---------------------------------------------------------------------------
# Seismic parsing & exports
# ---------------------------------------------------------------------------
def parse_seismic_file(file_path):
    tree = ET.parse(file_path)
    root = tree.getroot()

    metadata = {}
    seismic_data = []

    print(f"Root tag: {root.tag}, attributes: {root.attrib}")
    sys.stdout.flush()
    for element in root:
        if element.tag == 'Sensor':
            sensor_id = element.get('ID')
            z_sensor = element.get('ZSensor', 'Unknown')
            if sensor_id == '1':
                metadata['ZSensor1'] = z_sensor
            elif sensor_id == '2':
                metadata['ZSensor2'] = z_sensor
            print(f"Sensor {sensor_id}: ZSensor={z_sensor}")
            sys.stdout.flush()
        elif element.tag == 'ExternalTrigger':
            metadata['Delay'] = int(element.get('Delay', 0))
        elif element.tag == 'SampleTime':
            metadata['SampleTime'] = int(element.text)
        elif element.tag == 'ZRef':
            metadata['ZRef'] = element.text
        else:
            metadata[element.tag] = element.attrib if element.attrib else element.text

    print(f"Metadata extracted: {metadata}")
    sys.stdout.flush()

    for element in root.findall('S'):
        seismic_data.append({'S1': element.attrib['S1'], 'S2': element.attrib['S2']})
    seismic_data_df = pd.DataFrame(seismic_data, columns=['S1','S2'])
    return metadata, seismic_data_df

def save_seismic_data_corrected(file_number_str, metadata, seismic_data, output_dir, stack=False):
    folder_name = "Stacked VS" if stack else "Individual VS"
    export_dir = os.path.join(output_dir, folder_name)
    os.makedirs(export_dir, exist_ok=True)

    header_s1 = f"Tempo(ms); Amplitude(cm/s) Valores provisorios da Sismica; Profundidade[m] ={metadata.get('ZSensor1', 'Unknown')}\n"
    header_s2 = f"Tempo(ms); Amplitude(cm/s) Valores provisorios da Sismica; Profundidade[m] ={metadata.get('ZSensor2', 'Unknown')}\n"
    delay = metadata.get('Delay', 0)
    sample_time = int(metadata.get('SampleTime', 0))

    def create_time_column(row_num):
        return ((sample_time * row_num) / 1000) + delay

    file_s1 = os.path.join(export_dir, f"{file_number_str}-S1-A.txt")
    file_s2 = os.path.join(export_dir, f"{file_number_str}-S2-A.txt")
    with open(file_s1, 'w') as f_s1:
        f_s1.write(header_s1)
        for idx, row in seismic_data.iterrows():
            time_ms = create_time_column(idx)
            f_s1.write(f"{time_ms:.3f}\t{row['S1']}\n")
    with open(file_s2, 'w') as f_s2:
        f_s2.write(header_s2)
        for idx, row in seismic_data.iterrows():
            time_ms = create_time_column(idx)
            f_s2.write(f"{time_ms:.3f}\t{row['S2']}\n")
    return file_s1, file_s2

def invert_signal(value):
    max_value = 255
    try:
        value = float(value)
        if value < 0 or value > max_value:
            raise ValueError("Value is out of range for inversion (0 to 255)")
        return int(max_value - value)
    except (ValueError, TypeError):
        return value

def save_seismic_data_inverted(file_number_str, metadata, seismic_data, output_dir, stack=False):
    folder_name = "Stacked VS" if stack else "Individual VS"
    export_dir = os.path.join(output_dir, folder_name)
    os.makedirs(export_dir, exist_ok=True)

    header_s1 = f"Tempo(ms); Amplitude(cm/s) Valores provisorios da Sismica; Profundidade[m] ={metadata.get('ZSensor1', 'Unknown')}\n"
    header_s2 = f"Tempo(ms); Amplitude(cm/s) Valores provisorios da Sismica; Profundidade[m] ={metadata.get('ZSensor2', 'Unknown')}\n"
    delay = metadata.get('Delay', 0)
    sample_time = int(metadata.get('SampleTime', 0))

    def create_time_column(row_num):
        return ((sample_time * row_num) / 1000) + delay

    file_s1 = os.path.join(export_dir, f"{file_number_str}-S1-B.txt")
    file_s2 = os.path.join(export_dir, f"{file_number_str}-S2-B.txt")
    with open(file_s1, 'w') as f_s1:
        f_s1.write(header_s1)
        for idx, row in seismic_data.iterrows():
            time_ms = create_time_column(idx)
            inverted_s1 = invert_signal(row['S1'])
            f_s1.write(f"{time_ms:.3f}\t{inverted_s1}\n")
    with open(file_s2, 'w') as f_s2:
        f_s2.write(header_s2)
        for idx, row in seismic_data.iterrows():
            time_ms = create_time_column(idx)
            inverted_s2 = invert_signal(row['S2'])
            f_s2.write(f"{time_ms:.3f}\t{inverted_s2}\n")
    return file_s1, file_s2

def save_vp_data(file_number_str, metadata, seismic_data, output_dir):
    vp_dir = os.path.join(output_dir, "Individual VP")
    os.makedirs(vp_dir, exist_ok=True)

    header_s1 = f"Tempo(ms); Amplitude(cm/s) Valores provisorios da Sismica; Profundidade[m] ={metadata.get('ZSensor1', 'Unknown')}\n"
    header_s2 = f"Tempo(ms); Amplitude(cm/s) Valores provisorios da Sismica; Profundidade[m] ={metadata.get('ZSensor2', 'Unknown')}\n"
    delay = metadata.get('Delay', 0)
    sample_time = int(metadata.get('SampleTime', 0))

    def create_time_column(row_num):
        return ((sample_time * row_num) / 1000) + delay

    file_s1 = os.path.join(vp_dir, f"{file_number_str}-S1-C.txt")
    file_s2 = os.path.join(vp_dir, f"{file_number_str}-S2-C.txt")
    with open(file_s1, 'w') as f_s1:
        f_s1.write(header_s1)
        for idx, row in seismic_data.iterrows():
            time_ms = create_time_column(idx)
            f_s1.write(f"{time_ms:.3f}\t{row['S1']}\n")
    with open(file_s2, 'w') as f_s2:
        f_s2.write(header_s2)
        for idx, row in seismic_data.iterrows():
            time_ms = create_time_column(idx)
            f_s2.write(f"{time_ms:.3f}\t{row['S2']}\n")
    return file_s1, file_s2

# ---------------------------------------------------------------------------
# Extra UI buttons (VS/SDM exports)
# ---------------------------------------------------------------------------
def export_nonstacked_button_callback():
    try:
        directory = export_directory
        process_sdm_files(directory, stack=False)
        process_inverted_sdm_files(directory, stack=False)
    except Exception:
        logging.exception("Export non-stacked SDM failed")

def sdm_stack_export_button_callback():
    try:
        directory = export_directory
        process_sdm_files(directory, stack=True)
        # Stacked inverted export could be added if stacking is implemented to DataFrames.
    except Exception:
        logging.exception("Export stacked SDM failed")

def export_vp_button_callback():
    try:
        directory = export_directory
        process_vp_files(directory)
    except Exception:
        logging.exception("Export VP failed")

dataframes_info_button = ttk.Button(root, text='Show DataFrames Info', command=show_dataframes_info)
dataframes_info_button.pack(expand=True, pady=2)

transformed_dataframes_info_button = ttk.Button(root, text='Show Transformed DataFrames Info', command=show_transformed_dataframes_info)
transformed_dataframes_info_button.pack(expand=True, pady=2)

export_nonstacked_button = ttk.Button(root, text='Export Non-stacked SDM Data', command=export_nonstacked_button_callback)
export_nonstacked_button.pack(expand=True, pady=2)

sdm_stack_export_button = ttk.Button(root, text='Export Stacked SDM Data', command=sdm_stack_export_button_callback)
sdm_stack_export_button.pack(expand=True, pady=2)

export_vp_button = ttk.Button(root, text='Export Non-stacked VP Data', command=export_vp_button_callback)
export_vp_button.pack(expand=True, pady=2)

# ---------------------------------------------------------------------------
# Run the app
# ---------------------------------------------------------------------------
root.mainloop()
