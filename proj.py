import pandas as pd
import numpy as np
from openpyxl.styles import PatternFill

tabulationSheetOne = pd.read_csv('results_report (13).csv')
commentsSheet = pd.read_csv('raw_scores_report (13).csv')
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

columnsToKeep = ['Council ID', 'First Name', 'Last Name', 'Overall', 'AOII Interest (0)', 'Ambition (0)', 'Likability (0)', 'Sisterhood', 'Philanthropy', 'Pool']
columnsToRound = ['Overall', 'AOII Interest (0)', 'Ambition (0)', 'Likability (0)', 'Sisterhood', 'Philanthropy']

existingColumnsToKeep = [col for col in columnsToKeep if col in tabulationSheetOneFixed.columns]
newSheet = tabulationSheetOneFixed[existingColumnsToKeep].copy()

for col in columnsToRound:
    if col in newSheet.columns:
        newSheet[col] = newSheet[col].apply(lambda x: np.sign(x) * np.floor(np.abs(x) + 0.5) if pd.notna(x) else x)

for col in columnsToRound:
    if col in newSheet.columns:
        newSheet[col] = pd.to_numeric(newSheet[col], errors='coerce')

filteredSheet = newSheet[newSheet['Overall'].notna() & (newSheet['Overall'] != 0)].copy()
filteredSheet = filteredSheet.sort_values(by='Overall', ascending=False)

if 'Sisterhood' in filteredSheet.columns:
    sisterhoodRoundOne = filteredSheet[filteredSheet['Sisterhood'].notna()].copy()
else:
    sisterhoodRoundOne = pd.DataFrame()

mergedSheet = pd.merge( # merges filtered and commented sheet to get all filtered cols but add comments
    filteredSheet,
    commentsPivot,
    on = keyColumns,
    how = 'left'
)

sisterhood_primary = pd.DataFrame()
sisterhood_secondary = pd.DataFrame()
philanthropyPrimary = pd.DataFrame()
philantrhopySecondary = pd.DataFrame()

if 'Sisterhood' in mergedSheet.columns and 'Pool' in mergedSheet.columns:
    sisterhood_all = mergedSheet[mergedSheet['Sisterhood'].notna()].copy()
    sisterhood_primary = sisterhood_all[sisterhood_all['Pool'] == 'Primary'].copy()
    sisterhood_secondary = sisterhood_all[sisterhood_all['Pool'] == 'Secondary'].copy()

if 'Philanthropy' in mergedSheet.columns and 'Pool' in mergedSheet.columns:
    philanthropyAll = mergedSheet[mergedSheet['Sisterhood'].notna()].copy()
    philanthropyPrimary = philanthropyAll[philanthropyAll['Pool']== 'Primary'].copy()
    philantrhopySecondary = philanthropyAll[philanthropyAll['Pool']== 'Secondary'].copy

with pd.ExcelWriter('PhilanthropyRound.xlsx', engine='openpyxl') as writer: # CHANGE BASED ON ROUND
    newSheet.to_excel(writer, sheet_name="All Girls MasterList", index=False)
    if not sisterhood_primary.empty:
        sisterhood_primary.to_excel(writer, sheet_name='Sisterhood Primary', index=False)
    if not sisterhood_secondary.empty:
        sisterhood_secondary.to_excel(writer, sheet_name='Sisterhood Secondary', index=False)
    if not philanthropyPrimary.empty:
        philanthropyPrimary.to_excel(writer, sheet_name='Philanthropy Primary', index=False)
    if not philantrhopySecondary.empty:
        philantrhopySecondary.to_excel(writer, sheet_name='Secondary', index=False)


    workbook = writer.book
    greenFill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type="solid")
    lightGreenFill = PatternFill(start_color='E2F0D9', end_color='E2F0D9', fill_type="solid")
    purpleFill = PatternFill(start_color='C9A0DC', end_color='C9A0DC', fill_type="solid")
    yellowFill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type="solid")
    redFill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type="solid")
    whiteFill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type="solid")

    for sheetName in writer.sheets:
        worksheet = writer.sheets[sheetName]

        headers = [cell.value for cell in worksheet[1]]
        overallCol, sisterhoodCol, philoCol = None, None, None

        try:
            overallCol = headers.index('Overall')
            if 'Sisterhood' in headers: sisterhoodCol = headers.index('Sisterhood')
            if 'Philanthropy' in headers: philoCol = headers.index('Philanthropy')
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

                if sheetName == 'Sisterhood Primary' or sheetName == 'Sisterhood Secondary': # sisterhood sheet coloring
                    if sisterhoodCol is not None:
                        sisterhoodCell = row[sisterhoodCol]
                        isSisterhoodEmpty = sisterhoodCell.value is None or sisterhoodCell.value == ''

                        if isSisterhoodEmpty: 
                            for cell in row: cell.fill = redFill
                        else:
                            if score >= 8:
                                for cell in row: cell.fill = greenFill
                            elif 6 <= score < 8:
                                for cell in row: cell.fill = lightGreenFill
                            elif score < 6:
                                for cell in row: cell.fill = yellowFill
                    
                if sheetName == 'Philanthropy Primary' or sheetName == 'Philanthropy Secondary': # philo sheet coloring
                    if philoCol is not None:
                        philoCell = row[philoCol]
                        isPhiloEmpty = philoCell.value is None or philoCell.value == ''

                        if isPhiloEmpty: 
                            for cell in row: cell.fill = redFill
                        else:
                            if score >= 8:
                                for cell in row: cell.fill = greenFill
                            elif 6 <= score < 8:
                                for cell in row: cell.fill = lightGreenFill
                            elif score < 6:
                                for cell in row: cell.fill = yellowFill

                if sheetName == 'All Girls MasterList': # masterlist coloring
                    if score >= 8:
                        for cell in row: cell.fill = greenFill
                    elif 6 <= score < 8:
                        for cell in row: cell.fill = lightGreenFill
                    elif 0 < score < 6:
                        for cell in row: cell.fill = yellowFill
                    elif score == 0:
                        for cell in row: cell.fill = whiteFill

print("✅ Script finished")