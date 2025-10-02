import pandas as pd
import numpy as np
from openpyxl.styles import PatternFill

allPNMS = 'Report.csv'

outputFile = 'sortedByRCGroup.xlsx'
groupingColumn = 'RC Group'

try:
    df = pd.read_csv(allPNMS)
except FileNotFoundError:
    print(f"Error: file {allPNMS} wasnt found")
    exit()

individGroups = df[groupingColumn].unique()

with pd.ExcelWriter(outputFile, engine='openpyxl') as writer:
    print(f"creating excel file ")

    for group in sorted(individGroups):
        groupDF = df[df[groupingColumn] == group]

        sheetName = f'RC {group}'

        groupDF.to_excel(writer, sheet_name=sheetName, index=False)
        print(f" - Sheet '{sheetName}' created with {len(groupDF)} rows")

print("\n done")