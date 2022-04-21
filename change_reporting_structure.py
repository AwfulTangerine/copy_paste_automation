import pandas as pd
import openpyxl
df = pd.read_excel('C:/Users/Claudia.Ma/Downloads/test Global.xlsx')
df = df.set_index('SKU')
df = df.stack()
df.index = df.index.rename('week', level =1)
df.name = 'Qty'
df = df.reset_index()
df.to_excel("C:/Users/Claudia.Ma/Downloads/Global Result -3.xlsx")