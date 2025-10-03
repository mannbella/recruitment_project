import pandas as pd
import numpy as np
from openpyxl.styles import PatternFill

tabulationSheetOne = pd.read_csv('results_report (6).csv')
commentsSheet = pd.read_csv('raw_scores_report (5).csv')
tabulationSheetOne.columns = tabulationSheetOne.columns.str.strip()
tabulationSheetOneFixed = tabulationSheetOne.drop('List', axis=1, errors='ignore')

keyColumns = ['First Name', 'Last Name']
commentColumnName = 'Comment'

commentsSubset = commentsSheet[keyColumns + [commentColumnName]]
commentsSubset['comment_num'] = commentsSubset.groupby(keyColumns).cumcount() + 1

commentsPivot = commentsSubset.pivot_table(
    index = keyColumns,
    columns = 'comment_num',
    values = commentColumnName,
    aggfunc = 'first'
).reset_index()

commentsPivot.columns = [f'Comment {col}' if isinstance(col, int) else col for col in commentsPivot.columns]

columnsToKeep = ['Council ID', 'First Name', 'Last Name', 'Overall', 'AOII Interest (0)', 'Ambition (0)', 'Likability (0)', 'Sisterhood Day 1']
columnsToRound = ['Overall', 'AOII Interest (0)', 'Ambition (0)', 'Likability (0)', 'Sisterhood Day 1']

existingColumnsToKeep = [col for col in columnsToKeep if col in tabulationSheetOneFixed.columns]
newSheet = tabulationSheetOneFixed[existingColumnsToKeep].copy()

for col in columnsToRound:
    if col in newSheet.columns:
        newSheet[col] = pd.to_numeric(newSheet[col], errors='coerce')

filteredSheet = newSheet[newSheet['Overall'].notna() & (newSheet['Overall'] != 0)].copy()
filteredSheet = filteredSheet.sort_values(by='Overall', ascending=False)

for col in columnsToRound:
    if col in filteredSheet.columns:
        filteredSheet[col] = filteredSheet[col].apply(lambda x: np.sign(x) * np.floor(np.abs(x) + 0.5) if pd.notna(x) else x)

if 'Sisterhood Day 1' in filteredSheet.columns:
    sisterhoodRoundOne = filteredSheet[filteredSheet['Sisterhood Day 1'].notna()].copy()
else:
    sisterhoodRoundOne = pd.DataFrame()

#if 'House Tours 9/19' in filteredSheet.columns:
#    houseToursRound = filteredSheet[filteredSheet['House Tours 9/19'].notna()].copy()
#else:
#    houseToursRound = pd.DataFrame()

mergedSheet = pd.merge(
    filteredSheet,
    commentsPivot,
    on = keyColumns,
    how = 'left'
)

with pd.ExcelWriter('sisterhoodDayOne_HalfPoint.xlsx', engine='openpyxl') as writer: # CHANGE INDEX EVERY TIME PROG RUNS
    mergedSheet.to_excel(writer, sheet_name="All Girls MasterList", index=False)
    #if not houseToursRound.empty:
    #    houseToursRound.to_excel(writer, sheet_name='House Tours Round', index=False)

    workbook = writer.book
    greenFill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type="solid")
    lightGreenFill = PatternFill(start_color='E2F0D9', end_color='E2F0D9', fill_type="solid")
    purpleFill = PatternFill(start_color='C9A0DC', end_color='C9A0DC', fill_type="solid")
    yellowFill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type="solid")
    redFill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type="solid")

    for sheetName in writer.sheets:
        worksheet = writer.sheets[sheetName]
        
        headers = [cell.value for cell in worksheet[1]]
        overallCol, sisterhoodCol1 = None, None
        
        try:
            overallCol = headers.index('Overall')
            if 'Sisterhood Day 1' in headers: sisterhoodCol1 = headers.index('Sisterhood Day 1')
            #if 'House Tours 9/19' in headers: houseToursCol = headers.index('House Tours 9/19')
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

                if sheetName == 'All Girls MasterList':
                    if sisterhoodCol1 is not None:
                        sisterhood1Cell = row[sisterhoodCol1]
                        #houseTourCell = row[houseToursCol]
                        isSisterhoodEmpty = sisterhood1Cell.value is None or sisterhood1Cell.value == ''
                        #isHouseTourEmpty = houseTourCell.value is None or houseTourCell.value == ''

                        if isSisterhoodEmpty:
                            for cell in row: cell.fill = redFill
                        #elif isAPhiEmpty or isHouseTourEmpty:
                        #    for cell in row: cell.fill = purpleFill
                        else:
                            if score >= 8:
                                for cell in row: cell.fill = greenFill
                            elif 6 <= score < 8:
                                for cell in row: cell.fill = lightGreenFill
                            elif score < 6:
                                for cell in row: cell.fill = yellowFill
                
                    else:
                        if score >= 8:
                            for cell in row: cell.fill = greenFill
                        elif 6 <= score < 8:
                            for cell in row: cell.fill = lightGreenFill
                        elif score < 6:
                            for cell in row: cell.fill = yellowFill

print("✅ Script finished.")