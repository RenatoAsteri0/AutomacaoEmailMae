import pandas as pd

df = pd.read_excel('BaseClienteMarcia.xlsx')
data = df.to_dict('records')

for value in data:
    print(f'Cliente: {value['Cliente prospect']}\nContato: {value['Contato']}\nEmail: {value['e-mail']}\n')