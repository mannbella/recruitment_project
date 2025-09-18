import csv
import pandas as pd
import numpy as np
from openpyxl.styles import PatternFill

tabulationSheetOne = pd.read_csv('results_report (4).csv')
tabulationSheetOneFixed = tabulationSheetOne.drop('List', axis=1)

columnsToKeep = ['Council ID', 'First Name', 'Last Name', 'Overall', 'AOII Interest (0 - Standard Round)', 'Ambition (0 - Standard Round)', 'Likability (0 - Standard Round)', 'Alpha Phi 9/18' ]
columnsToRound = ['Overall', 'AOII Interest (0 - Standard Round)', 'Ambition (0 - Standard Round)', 'Likability (0 - Standard Round)', 'Alpha Phi 9/18']

newSheet = tabulationSheetOneFixed[columnsToKeep]
filteredSheet = newSheet[newSheet['Overall'] != 0]
filteredSheet = filteredSheet.sort_values(by='Overall', ascending=False)

for col in columnsToRound:
    filteredSheet[col] = np.sign(filteredSheet[col]) * np.floor(np.abs(filteredSheet[col]) + 0.5)

aPhiRound = filteredSheet[filteredSheet['Alpha Phi 9/18'].notna()].copy()
nonAPhiRound = filteredSheet[filteredSheet['Alpha Phi 9/18'].isna()].copy()

with pd.ExcelWriter('tabulationOutput3_918.xlsx') as writer: # CHANGE THE INDEX EACH TIME YOU RUN THE PROG
    filteredSheet.to_excel(writer, sheet_name="All Girls MasterList", index=False)
    aPhiRound.to_excel(writer, sheet_name='APhi Round', index=False)
    nonAPhiRound.to_excel(writer, sheet_name='Did not go to APhi', index=False)

    workboox = writer.book
    highlightFill = PatternFill(start_color='008000', end_color='008000', fill_type="solid")

    for sheetName in writer.sheets:
        worksheet = writer.sheets[sheetName]
        overallColumnIndex = None
        for i, cell in enumerate(worksheet[1]):
            if cell.value == 'Overall':
                overallColumnIndex = i + 1
                break

        if overallColumnIndex:
            for row in worksheet.iter_rows(min_row = 2):
                overallCell = row[overallColumnIndex - 1]

                if overallCell.value is not None and overallCell.value >= 8:
                    for cell in row:
                        cell.fill = highlightFill
