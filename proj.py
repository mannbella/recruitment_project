import pandas as pd
import numpy as np
from openpyxl.styles import PatternFill

# --- 1. Data Loading and Cleaning ---
tabulationSheetOne = pd.read_csv('results_report (6).csv')
commentsSheet = pd.read_csv('raw_scores_report (5).csv')
tabulationSheetOne.columns = tabulationSheetOne.columns.str.strip()
tabulationSheetOneFixed = tabulationSheetOne.drop('List', axis=1, errors='ignore')

# --- 2. Comments Processing ---
keyColumns = ['First Name', 'Last Name']
commentColumnName = 'Comment'

if all(col in commentsSheet.columns for col in keyColumns + [commentColumnName]):
    commentsSubset = commentsSheet[keyColumns + [commentColumnName]].copy()
    for col in keyColumns:
        commentsSubset[col] = commentsSubset[col].astype(str).str.strip()
    
    commentsSubset['comment_num'] = commentsSubset.groupby(keyColumns).cumcount() + 1
    commentsPivot = commentsSubset.pivot_table(
        index=keyColumns, columns='comment_num', values=commentColumnName, aggfunc='first'
    ).reset_index()
    commentsPivot.columns = [f'Comment {col}' if isinstance(col, int) else col for col in commentsPivot.columns]
else:
    commentsPivot = pd.DataFrame(columns=keyColumns)

# --- 3. Main Data Processing ---
columnsToKeep = ['Council ID', 'First Name', 'Last Name', 'Overall', 'AOII Interest (0)', 'Ambition (0)', 'Likability (0)', 'Sisterhood Day 1']
existingColumnsToKeep = [col for col in columnsToKeep if col in tabulationSheetOneFixed.columns]
newSheet = tabulationSheetOneFixed[existingColumnsToKeep].copy()

columnsToRound = ['Overall', 'AOII Interest (0)', 'Ambition (0)', 'Likability (0)', 'Sisterhood Day 1']
for col in columnsToRound:
    if col in newSheet.columns:
        newSheet[col] = pd.to_numeric(newSheet[col], errors='coerce')

# <-- FIX: Re-introduce the filter to only keep girls with a valid 'Overall' score.
filteredSheet = newSheet[newSheet['Overall'].notna() & (newSheet['Overall'] != 0)].copy()
filteredSheet = filteredSheet.sort_values(by='Overall', ascending=False)

# Apply rounding to the filtered data
for col in columnsToRound:
    if col in filteredSheet.columns:
        filteredSheet[col] = filteredSheet[col].apply(lambda x: np.sign(x) * np.floor(np.abs(x) + 0.5) if pd.notna(x) else x)

# <-- FIX: Merge comments into the *filtered* sheet, not the complete list.
mergedSheet = pd.merge(filteredSheet, commentsPivot, on=keyColumns, how='left')

# Create the 'Sisterhood' sheet from the already filtered and merged data
if 'Sisterhood Day 1' in mergedSheet.columns:
    sisterhoodRoundOne = mergedSheet[mergedSheet['Sisterhood Day 1'].notna()].copy()
else:
    sisterhoodRoundOne = pd.DataFrame()

# --- 4. Writing to Excel with Formatting ---
with pd.ExcelWriter('sisterhoodDayOne_1.xlsx', engine='openpyxl') as writer:
    # The mergedSheet is now the correct, filtered list
    mergedSheet.to_excel(writer, sheet_name="All Girls MasterList", index=False)
    
    if not sisterhoodRoundOne.empty:
        sisterhoodRoundOne.to_excel(writer, sheet_name='Sisterhood Day 1', index=False)

    # --- (The rest of the formatting code remains the same) ---
    workbook = writer.book
    greenFill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type="solid")
    lightGreenFill = PatternFill(start_color='E2F0D9', end_color='E2F0D9', fill_type="solid")
    yellowFill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type="solid")
    redFill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type="solid")

    for sheetName in writer.sheets:
        worksheet = writer.sheets[sheetName]
        headers = [cell.value for cell in worksheet[1]]
        try:
            overallCol = headers.index('Overall')
            sisterhoodCol1 = headers.index('Sisterhood Day 1') if 'Sisterhood Day 1' in headers else -1
        except ValueError:
            continue

        for row in worksheet.iter_rows(min_row=2):
            overallCell = row[overallCol]
            score = 0
            if overallCell.value is not None:
                try: score = float(overallCell.value)
                except (ValueError, TypeError): score = 0

            if sheetName == 'All Girls MasterList' and sisterhoodCol1 != -1:
                sisterhood1Cell = row[sisterhoodCol1]
                isSisterhoodEmpty = sisterhood1Cell.value is None or sisterhood1Cell.value == ''
                if isSisterhoodEmpty:
                    for cell in row: cell.fill = redFill
                    continue

            if score >= 8:
                for cell in row: cell.fill = greenFill
            elif 6 <= score < 8:
                for cell in row: cell.fill = lightGreenFill
            elif score > 0 and score < 6:
                for cell in row: cell.fill = yellowFill

print("✅ Script finished")