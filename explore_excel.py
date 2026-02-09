import pandas as pd

file_path = r'c:\Users\Vstor\Desktop\Ludovico\Analisi Dati\Previsioni 2026 Montagna\Analisi Montagna Forecasting 2026 (dal 2024).xlsx'

# Lista fogli
xl = pd.ExcelFile(file_path)
print('Fogli disponibili:', xl.sheet_names)

# Leggi prime righe del foglio VEND
df = pd.read_excel(file_path, sheet_name='VEND', nrows=20)
print('\nColonne:', list(df.columns))
print('\nPrime 10 righe:')
print(df.head(10))
print('\nTipi dati:')
print(df.dtypes)
