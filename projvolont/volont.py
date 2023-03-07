import pandas as pd
from docx import Document





ex = pd.read_excel('list.xlsx')
ex.fillna(0, inplace=True)

z = ex[['Unnamed: 5']].values.tolist()

for i in range(len(z)):
    if (z[i] != [0]) and (z[i] != ['sda']):
        document = Document()
        document.add_heading(f"{z[i]}", 0)
        document.save(f"{z[i]}.docx")



