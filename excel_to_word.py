# File name: excel_to_word.py
# Author: Kyle LaPolice
# Date created: 21 - Jun - 2022
# Date last modified: 30 - Jun - 2022
# Script Version: 2.0

from pathlib import Path

import PySimpleGUI as sg
import pandas as pd # pip install pandas openpyxL
from docxtpl import DocxTemplate # pip install docxtpl
from docx2pdf import convert # pip install docx2pdf

window_name = 'Excel to Word Docs from Template' # Change name of window per project

# Layout of window
layout = [
    [sg.Text("Select an excel file:")],
    [sg.Input(key='EXCEL', enable_events = True), sg.FileBrowse(key = 'eBROWSE')],

    [sg.Text("Select an excel sheet:")],
    [sg.Combo(values=[], key = 'COMBO', readonly = True, expand_x = True)],

    [sg.Text("Select template file:")],
    [sg.Input(key = 'TEMP', enable_events = True), sg.FileBrowse(key = 'wBROWSE')],

    [sg.Checkbox("Convert Final Word docs to PDF", default = False, key = 'PDF')],

    [sg.OK(key = 'GO'), sg.Cancel(key = 'Exit')]
    ]

window = sg.Window(window_name, layout) # Creates window that user interacts with

# Run the Event Loop
while True:
    event, values = window.read()

    # Close window if events
    if event == 'Exit' or event == sg.WIN_CLOSED:
        break

    # updates Combo with Excel sheets
    if event == 'EXCEL':
        e_sheets = pd.read_excel(values.get('EXCEL'), sheet_name=None)
        window['COMBO'].update(values = list(e_sheets.keys()))
        continue

    # executes program
    if event =='GO':

        # sets directories
        base_dir =  Path(__file__).parent
        output_dir = base_dir / "WORD"
        pdf_dir = base_dir / "PDF"

        # Create output folder for the word documents
        output_dir.mkdir(exist_ok=True)

        # gets data from excel sheet chosen from combo
        e_sheets = pd.read_excel(values.get('EXCEL'), sheet_name = values['COMBO'])

        # itterates over template to create word docs from templates
        for record in e_sheets.to_dict(orient="records"):
            doc = DocxTemplate(values['TEMP'])
            doc.render(record)
            output_path = output_dir / f"{record['DOCNUM']}.docx"
            doc.save(output_path)

        # creates pdf docs from word doc folder
        if values['PDF'] == True:
            pdf_dir.mkdir(exist_ok=True)
            convert(output_dir, pdf_dir)
            continue
        continue
window.close()
