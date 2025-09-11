import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pandas as pd
from email.mime.image import MIMEImage

# 1
df = pd.read_excel('BaseClienteMarcia.xlsx')
data = df.to_dict("records")

# 2 mandar email
clientes = [
    {
        'Cliente prospect': 'RenatoAsterio',
        'Contato': '19 996234793',
        'e-mail': 'renato.asterio@hotmail.com',
        'Cidade': 'Indaiatuba'
    },
    # adicione outros clientes aqui
]

# credenciais Titan
smtp_server = 'smtp.titan.email'
smtp_port = 587
smtp_username = 'comercial@marciaasterio.com'
smtp_password = 'Belinha10@'
from_addr = 'comercial@marciaasterio.com'

# lista de imagens da apresentação
imagens = ["APRESENTAÇÃO_INSAC__page-0001.jpg", "APRESENTAÇÃO_INSAC__page-0002.jpg", "APRESENTAÇÃO_INSAC__page-0003.jpg",
           "APRESENTAÇÃO_INSAC__page-0004.jpg", "APRESENTAÇÃO_INSAC__page-0005.jpg"]

for cliente in data:
    msg = MIMEMultipart("related")
    msg['From'] = from_addr
    msg['To'] = cliente['e-mail']
    bcc = "marcia.insac@gmail.com"
    msg['Subject'] = f"Apresentação INSAC Embalagens à {cliente['Cliente prospect']}"

    # HTML com todas as imagens
    html = f"""
    <html>
        <body style="font-family: Arial, sans-serif; line-height: 1.6;">
            <p>Olá <b>{cliente['Contato']}</b>,</p>
            <p>
                Espero que esteja bem!<br>
                Segue abaixo a Apresentação Comercial da nossa empresa.<br>
                Estou à disposição para esclarecer dúvidas e avançar com uma proposta personalizada.
            </p>
    """
    for i, img_name in enumerate(imagens):
        html += f'<p><img src="cid:img{i + 1}" style="max-width:600px"></p>'

    html += """
            <p>Atenciosamente,</p>
            <p>
                <b>Equipe Comercial</b><br>
                Marcia Astério Consultoria<br>
                📧 Email: comercial@marciaasterio.com<br>
                📞 Telefone/WhatsApp: (19) 98167-0086
            </p>
        </body>
    </html>
    """
    msg.attach(MIMEText(html, "html"))

    # anexar imagens
    for i, img_name in enumerate(imagens):
        with open(img_name, "rb") as f:
            img = MIMEImage(f.read())
            img.add_header("Content-ID", f"<img{i + 1}>")
            msg.attach(img)

    # envio
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(smtp_username, smtp_password)
        server.sendmail(from_addr, [cliente['e-mail'], bcc], msg.as_string())

    print(f"✔ Email enviado para {cliente['Cliente prospect']} - {cliente['Contato']} ({cliente['e-mail']})")

'''
df = df.map(lambda x: x.strip().replace('\n', '').replace('-', '').replace('(', '').
            replace(')', '') if isinstance(x, str) else x)
capitalizar_colums = ['Cliente prospect', 'Contato', 'Cidade']
df[capitalizar_colums] = df[capitalizar_colums].apply(lambda x: x.str.title())
to_excel = df.to_excel('ClienteMarciaArrumaTelefones.xlsx', index=False)
emails = df[['Cliente prospect', 'Contato', 'e-mail', 'telefone', 'Cidade']].to_dict('records')
print(emails)
'''