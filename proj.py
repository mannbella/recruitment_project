import csv
import pandas as pd
import numpy as np

tabulationSheetOne = pd.read_csv('results_report_(4).csv')
tabulationSheetOneFixed = tabulationSheetOne.drop('List', axis=1)

columnsToKeep = ['Council ID', 'First Name', 'Last Name', 'Overall', 'AOII Interest (0 - Standard Round)', 'Ambition (0 - Standard Round)', 'Likability (0 - Standard Round)' ]
columnsToRound = ['Overall', 'AOII Interest (0 - Standard Round)', 'Ambition (0 - Standard Round)', 'Likability (0 - Standard Round)']

newSheet = tabulationSheetOneFixed[columnsToKeep]
filteredSheet = newSheet[newSheet['Overall'] != 0]
filteredSheet = filteredSheet.sort_values(by='Overall', ascending=False)

for col in columnsToRound:
    filteredSheet[col] = np.sign(filteredSheet[col]) * np.floor(np.abs(filteredSheet[col]) + 0.5)

filteredSheet.to_csv('tabulationOutput2.csv', index=False) # CHANGE THE INDEX EACH TIME YOU RUN THE PROG

print(filteredSheet.head(10))