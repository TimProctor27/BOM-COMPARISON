""" Created By: Tim Proctor
    Created Date: 6/8/23
    Version: 7.1
    ----------------------------------------------------------------
    Main file to open BOM Compare GUI and then convert, clean, and 
    ouput data from csv and xls files.
"""

import BomCompareGUI

if __name__ == "__main__":
    BomCompareGUI.gui()

"""Distribute .exe using following code
pyinstaller --onefile --windowed main.py"""
