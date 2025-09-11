"""
1 - colocar cada linha em 1 índice de uma lista, pois terei uma lista de todas as linhas
2 - configurar o email smtp
"""

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pandas as pd
# 1
df = pd.read_excel('BaseClienteMarcia.xlsx')
data = df.to_dict()
#for key, value in data.items():
#    print(f'{key}: {value[9]}')

# 2 mandar email
test_dict = [
    {
    'Cliente prospect': 'RenatoAsterio',
    'Contato': '19 996234793',
    'e-mail': 'renato.asterio@hotmail.com',
    'Cidade': 'indaiatuba'
}]

# credenciais Titan
smtp_server = 'smtp.titan.email'
smtp_port = 587
smtp_username = 'comercial@marciaasterio.com'
smtp_password = 'Belinha10@'
from_addr = 'comercial@marciaasterio.com'

# loop para cada cliente
for cliente in test_dict:
    msg = MIMEMultipart("alternative")
    msg['From'] = from_addr
    msg['To'] = cliente['e-mail']
    msg['Subject'] = f"Apresentação Comercial - {cliente['Cliente prospect']}"

    # corpo em HTML
    html = f"""
    <html>
        <body style="font-family: Arial, sans-serif; line-height: 1.6;">
            <p>Olá <b>{cliente['Cliente prospect']}</b>,</p>
            <p>
                Espero que esteja bem!<br>
                Segue abaixo a apresentação comercial da nossa empresa.<br>
                Estaremos à disposição para esclarecer dúvidas e avançar com uma proposta personalizada.
            </p>
            <p>
                <b>Cidade:</b> {cliente['Cidade']}
            </p>
            <p>Atenciosamente,</p>
            <p>
                <b>Equipe Comercial</b><br>
                Marcia Asterio Consultoria<br>
                📧 comercial@marciaasterio.com<br>
                📞 (11) 99999-9999
            </p>
        </body>
    </html>
    """

    # anexar html ao e-mail
    msg.attach(MIMEText(html, "html"))

    # envio
    with smtplib.SMTP(smtp_server, smtp_port) as connection:
        connection.starttls()
        connection.login(smtp_username, smtp_password)
        connection.sendmail(from_addr, cliente['e-mail'], msg.as_string())

    print(f"✔ Email enviado para {cliente['Cliente prospect']} ({cliente['e-mail']})")
'''
df = df.map(lambda x: x.strip().replace('\n', '').replace('-', '').replace('(', '').
            replace(')', '') if isinstance(x, str) else x)
capitalizar_colums = ['Cliente prospect', 'Contato', 'Cidade']
df[capitalizar_colums] = df[capitalizar_colums].apply(lambda x: x.str.title())
to_excel = df.to_excel('ClienteMarciaArrumaTelefones.xlsx', index=False)
emails = df[['Cliente prospect', 'Contato', 'e-mail', 'telefone', 'Cidade']].to_dict('records')
print(emails)
'''