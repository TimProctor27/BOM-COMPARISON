""" Created By: Tim Proctor
    Created Date: 6/8/23
    Version: 7.1
    ----------------------------------------------------------------
    File of functions that creates a GUI to select, import, and export
    csv and xls files
"""

import PySimpleGUI as sg
import pathlib
import pandas as pd
import DataClean

# ----Clear Data from GUI----#
def clear_data(window):
    for key, element in window.key_dict.items():
        if isinstance(element, sg.Input):
            element.update(value="")

# ----GUI Build---- #
def gui():
    sg.theme("DarkTeal9")

    layout = [
        [sg.Text("Please select files to compare:")],
        [
            sg.Input(enable_events=True, key="-DRAW_IN-"),
            sg.FileBrowse(
                "Select Drawing BOM",
                file_types=[
                    ("Drawing files", ".csv .xls .xlsx "),
                ],
            ),
        ],
        [
            sg.Input(enable_events=True, key="-XLSX_IN-"),
            sg.FileBrowse(
                " Select D365 BOM ",
                file_types=[
                    ("XLSX Files", "*.xls*"),
                ],
            ),
        ],
        [
            sg.Input(key="-out-"),
            sg.FileSaveAs(
                " Select Output File ",
                file_types=[("XLSX Files", "*.xls*")],
                key="-filename-",
            ),
        ],
        [
            sg.B("Merge and Save"),
            sg.B("Clear Data"),
            sg.Exit(),
        ],
    ]

    window = sg.Window("File Select Form", layout)

# ----Runs GUI Event Loop---- #
    while True:
        event, values = window.read()

        if event in (sg.WINDOW_CLOSED, "Exit"):
            break

        if event == "Clear Data":
            clear_data(window)

        if event == "-DRAW_IN-":
            drawing_file_path = values["-DRAW_IN-"]
            file_extension = pathlib.Path(drawing_file_path).suffix
            # Checks file extension and reads data from file into dataframe
            if file_extension == ".csv":
                df_drawing_bom = pd.read_csv(drawing_file_path, encoding="latin-1")
                df_drawing_bom_out = DataClean.drawing_files(df_drawing_bom)

            elif file_extension == ".xlsx":
                df_drawing_bom = pd.read_excel(drawing_file_path)
                df_drawing_bom_out = DataClean.drawing_files(df_drawing_bom)

            elif file_extension == ".xls":
                df_drawing_bom = pd.read_excel(drawing_file_path)
                df_drawing_bom_out = DataClean.drawing_files(df_drawing_bom)

            else:
                sg.popup("Incorrect\n\nFile Type")

        if event == "-XLSX_IN-":
            df_xl_bom_filepath = values["-XLSX_IN-"]
            df_xl_bom = pd.read_excel(df_xl_bom_filepath)
            df_xl_bom_out = DataClean.xlsx_files(df_xl_bom)

        if event == "Merge and Save":
            try:
                filename = values["-filename-"]
                DataClean.merged_file(filename, df_drawing_bom_out, df_xl_bom_out)
            except FileNotFoundError:
                sg.popup_error("\nPlease enter all \nnecessary files. \n")
            except NameError:
                sg.popup_error("\nPlease select all \nnecessary files. \n")

    window.close()