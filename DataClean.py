""" Created By: Tim Proctor
    Created Date: 6/8/23
    Version: 7.1
    ----------------------------------------------------------------
    File of functions to check, clean, and merge data imported from
    a csv or xls file
"""

# ----Import Modules----#
import pandas as pd
import os
import PySimpleGUI as sg
from styleframe import StyleFrame

# ----Drawing BOM data----#


def drawing_files(df_drawing_bom):
    headers_csv = list(df_drawing_bom.columns.values)
    standard_alt = ["Designator", "Part Number", "Description", "Quantity"]
    standard_sw = ["ITEM NO.", "PART NUMBER", "DESCRIPTION", "QTY."]

    if (
        headers_csv == standard_sw
    ):  # Checks and modifies dataframe if exported from Solidworks
        df_drawing_bom = df_drawing_bom.drop(
            ["ITEM NO.", "DESCRIPTION"], axis=1)
        df_drawing_bom = df_drawing_bom.rename(
            columns={"PART NUMBER": "Part_Number"})
        df_drawing_bom = df_drawing_bom.rename(columns={"QTY.": "Quantity"})
        df_drawing_bom = df_drawing_bom.dropna(axis=0)
        df_drawing_bom["Part_Number"] = df_drawing_bom["Part_Number"].apply(
            str)
        hardware = ["8SCR", "8HDW", "LUG"]
        df_drawing_bom_out = df_drawing_bom[
            df_drawing_bom.Part_Number.str.contains(
                "|".join(hardware)) == False
        ]
        df_drawing_bom_out.sort_values(["Part_Number"], inplace=True)

    elif (
        headers_csv == standard_alt
    ):  # Checks and modifies dataframe if exported from Altium
        df_drawing_bom = df_drawing_bom.drop(
            ["Designator", "Description"], axis=1)
        df_drawing_bom = df_drawing_bom.rename(
            columns={"Part Number": "Part_Number"})
        df_drawing_bom = df_drawing_bom.dropna(axis=0)
        df_drawing_bom["Part_Number"] = df_drawing_bom["Part_Number"].apply(
            str)
        SIT_resistors = [
            "E4700-SIT",
            "E4700-RNC50 SIT",
            "E4700-1206 SIT",
            "E4700-RN60C SIT",
            "E4700-RS-2B SIT",
            "E4700-0805 SIT",
            "E4700-0402 SIT",
            "E4700-SIT-1210",
            "E1500-SIT1206",
            "E2100-160-2043-03-01"
        ]
        df_drawing_bom_out = df_drawing_bom[
            df_drawing_bom.Part_Number.isin(SIT_resistors) == False
        ]
        df_drawing_bom_out.sort_values(["Part_Number"], inplace=True)

    else:
        sg.popup("\nColumn headers \ndo not match \nDrawing Standard")

    return df_drawing_bom_out


# ----D365 BOM data----#
def xlsx_files(df_xl_bom):
    headers_xl = list(df_xl_bom.columns.values)
    standard_xl = [
        "Item number",
        "Configuration",
        "Size",
        "Version",
        "Notes",
        "Warehouse",
        "Resource consumption",
        "Quantity",
        "Per series",
        "Unit",
        "Configuration group",
        "Product name",
        "Drawing number",
        "Drawing revision",
    ]

    """if headers_xl == standard_xl:  # Checks and modifies dataframe if exported from D365"""
    if "Item number" in headers_xl and "Quantity" in headers_xl:  # Checks and modifies dataframe if Item Number and Quantity are in the df from D365
        df_xl_bom = df_xl_bom.rename(columns={"Item number": "Part_Number"})
        df_xl_bom["Part_Number"] = df_xl_bom["Part_Number"].apply(str)
        used_columns = ["Part_Number", "Quantity"]
        df_xl_bom_out = df_xl_bom[df_xl_bom.columns.intersection(used_columns)]
        df_xl_bom_out.sort_values(["Part_Number"], inplace=True)

    else:
        sg.popup("\nColumn headers \ndo not match \nD365 Standard")

    return df_xl_bom_out


# ----Merge, Compare, Save, and Open Output file----#
def merged_file(filename, df_drawing_bom_out, df_xl_bom_out):
    merged_bom = df_drawing_bom_out.merge(
        df_xl_bom_out, indicator=True, how="outer")

    # Uses StyleFrame to best fit columns in Excel spreadsheet output file #
    sf_df_drawing_bom = StyleFrame(df_drawing_bom_out)
    sf_df_xl_bom = StyleFrame(df_xl_bom_out)
    sf_merged_bom = StyleFrame(merged_bom)

    try:
        with StyleFrame.ExcelWriter(filename) as excel_writer:
            sf_df_drawing_bom.to_excel(
                excel_writer=excel_writer,
                sheet_name="Drawing BOM",
                best_fit=["Part_Number", "Quantity"],
            )
            sf_df_xl_bom.to_excel(
                excel_writer=excel_writer,
                sheet_name="D365 BOM",
                best_fit=["Part_Number", "Quantity"],
            )
            sf_merged_bom.to_excel(
                excel_writer=excel_writer,
                sheet_name="Merged BOM",
                best_fit=["Part_Number", "Quantity", "_merge"],
            )

        os.startfile(filename)

    except PermissionError:
        sg.popup(
            "File is open \nin another application.\nPlease close file \nand try again."
        )
