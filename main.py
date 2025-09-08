"""
1 - colocar cada linha em 1 índice de uma lista, pois terei uma lista de todas as linhas
2 - configurar o email smtp
"""

import pandas as pd
# 1
df = pd.read_excel('Planilha Renato.xlsx', sheet_name='Prospects Insac', usecols='A:E')

df = df.map(lambda x: x.strip().replace('\n', '').replace('-', '').replace('(', '').
            replace(')', '') if isinstance(x, str) else x)
capitalizar_colums = ['Cliente prospect', 'Contato', 'Cidade']
df[capitalizar_colums] = df[capitalizar_colums].apply(lambda x: x.str.title())
to_excel = df.to_excel('ClienteMarciaArrumaTelefones.xlsx', index=False)
emails = df[['Cliente prospect', 'Contato', 'e-mail', 'telefone', 'Cidade']].to_dict('records')
print(emails)