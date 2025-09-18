import csv
import pandas as pd

tabulationSheetOne = pd.read_csv('results_report (4).csv')
tabulationSheetOneFixed = tabulationSheetOne.drop('List', axis=1)

columnsToKeep = ['Council ID', 'First Name', 'Last Name', 'Overall', 'AOII Interest (0 - Standard Round)', 'Ambition (0 - Standard Round)', 'Likability (0 - Standard Round)' ]

newSheet = tabulationSheetOneFixed[columnsToKeep]
filteredSheet = newSheet[newSheet['Overall'] != 0]
filteredSheet = filteredSheet.sort_values(by='Overall', ascending=False)
filteredSheet.to_csv('tabulationOutput1.csv', index=False) # CHANGE THE INDEX EACH TIME YOU RUN THE PROG

print(filteredSheet.head(10))