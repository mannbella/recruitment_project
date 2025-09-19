import pandas as pd
import numpy as np
from openpyxl.styles import PatternFill

# --- Data Loading and Cleaning ---
tabulationSheetOne = pd.read_csv('results_report (5).csv')
tabulationSheetOne.columns = tabulationSheetOne.columns.str.strip()
tabulationSheetOneFixed = tabulationSheetOne.drop('List', axis=1, errors='ignore')

columnsToKeep = ['Council ID', 'First Name', 'Last Name', 'Overall', 'AOII Interest (0 - Standard Round)', 'Ambition (0 - Standard Round)', 'Likability (0 - Standard Round)', 'Alpha Phi 9/18', 'House Tours 9/19']
columnsToRound = ['Overall', 'AOII Interest (0 - Standard Round)', 'Ambition (0 - Standard Round)', 'Likability (0 - Standard Round)', 'Alpha Phi 9/18', 'House Tours 9/19']

existingColumnsToKeep = [col for col in columnsToKeep if col in tabulationSheetOneFixed.columns]
newSheet = tabulationSheetOneFixed[existingColumnsToKeep].copy()

# Convert score columns to numeric type
for col in columnsToRound:
    if col in newSheet.columns:
        newSheet[col] = pd.to_numeric(newSheet[col], errors='coerce')

filteredSheet = newSheet[newSheet['Overall'].notna() & (newSheet['Overall'] != 0)].copy()
filteredSheet = filteredSheet.sort_values(by='Overall', ascending=False)

# --- Data Processing ---
for col in columnsToRound:
    if col in filteredSheet.columns:
        filteredSheet[col] = filteredSheet[col].apply(lambda x: np.sign(x) * np.floor(np.abs(x) + 0.5) if pd.notna(x) else x)

if 'Alpha Phi 9/18' in filteredSheet.columns:
    aPhiRound = filteredSheet[filteredSheet['Alpha Phi 9/18'].notna()].copy()
else:
    aPhiRound = pd.DataFrame()

if 'House Tours 9/19' in filteredSheet.columns:
    houseToursRound = filteredSheet[filteredSheet['House Tours 9/19'].notna()].copy()
else:
    houseToursRound = pd.DataFrame()

# --- Writing to Excel with Formatting ---
with pd.ExcelWriter('tabulationOutput9_919.xlsx', engine='openpyxl') as writer:
    filteredSheet.to_excel(writer, sheet_name="All Girls MasterList", index=False)
    if not aPhiRound.empty:
        aPhiRound.to_excel(writer, sheet_name='APhi Round', index=False)
    if not houseToursRound.empty:
        houseToursRound.to_excel(writer, sheet_name='House Tours Round', index=False)

    # --- Define Fills ---
    workbook = writer.book
    greenFill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type="solid")
    lightGreenFill = PatternFill(start_color='E2F0D9', end_color='E2F0D9', fill_type="solid")
    purpleFill = PatternFill(start_color='C9A0DC', end_color='C9A0DC', fill_type="solid")
    yellowFill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type="solid")
    redFill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type="solid")

    # --- Main Formatting Loop ---
    for sheetName in writer.sheets:
        worksheet = writer.sheets[sheetName]
        
        headers = [cell.value for cell in worksheet[1]]
        overallCol, aPhiCol, houseToursCol = None, None, None
        
        try:
            overallCol = headers.index('Overall')
            if 'Alpha Phi 9/18' in headers: aPhiCol = headers.index('Alpha Phi 9/18')
            if 'House Tours 9/19' in headers: houseToursCol = headers.index('House Tours 9/19')
        except ValueError:
            continue

        if overallCol is not None:
            for row in worksheet.iter_rows(min_row=2):
                overallCell = row[overallCol]
                score = 0
                if overallCell.value is not None:
                    try:
                        score = float(overallCell.value)
                    except (ValueError, TypeError):
                        continue

                # LOGIC FOR THE MASTER LIST
                if sheetName == 'All Girls MasterList':
                    if aPhiCol is not None and houseToursCol is not None:
                        aPhiCell = row[aPhiCol]
                        houseTourCell = row[houseToursCol]
                        isAPhiEmpty = aPhiCell.value is None or aPhiCell.value == ''
                        isHouseTourEmpty = houseTourCell.value is None or houseTourCell.value == ''

                        if isAPhiEmpty and isHouseTourEmpty:
                            for cell in row: cell.fill = redFill
                        elif isAPhiEmpty or isHouseTourEmpty:
                            for cell in row: cell.fill = purpleFill
                        # <-- FIX: This 'else' block applies the score-based colors
                        # to the remaining girls on the MasterList.
                        else:
                            if score >= 8:
                                for cell in row: cell.fill = greenFill
                            elif 6 <= score < 8:
                                for cell in row: cell.fill = lightGreenFill
                            elif score < 6:
                                for cell in row: cell.fill = yellowFill
                
                # LOGIC FOR THE OTHER SHEETS
                else:
                    if score >= 8:
                        for cell in row: cell.fill = greenFill
                    elif 6 <= score < 8:
                        for cell in row: cell.fill = lightGreenFill
                    elif score < 6:
                        for cell in row: cell.fill = yellowFill

print("✅ Script finished. The MasterList should now be fully highlighted.")