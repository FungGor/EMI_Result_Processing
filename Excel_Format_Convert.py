import pandas as pd

new_dataFrame = pd.read_csv('4A-GDC-BF-VERT.csv')
new_excel = pd.ExcelWriter('4A-GDC-BF-VERT.xlsx')
new_dataFrame.to_excel(new_excel, index=False)
new_excel._save()