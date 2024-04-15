import win32com.client
from datetime import datetime


destinatario = 'Nome do destinatário'
saudacao = ['Bom dia' if 1 <= datetime.today().hour < 12 else 'Boa tarde' if 12 < datetime.today().hour < 18 else 'Boa noite'][0]

outlook = win32com.client.Dispatch("Outlook.Application")
email = outlook.CreateItem(0)  # 0 significa email
email.To = "exemplo@email.com"
email.Subject = "Assunto do e-mail"
email.HTMLBody = f'''
<body style="font-family: Arial, sans-serif;">
    <p>{saudacao}, {destinatario}. Espero que esteja bem!</p>
    <br>
    <p>Conteúdo do e-mail.</p>
    <p>É possível adicionar:</p>
    <ul>
        <li>Textos em <span style="font-weight: bold;">negrito</span>, <span style="font-style: italic;">textos em itálico</span>, <span style="text-decoration: underline;">textos sublinhados</span>, etc.</li>
        <li><a href="https://www.exemplo.com">Links clicáveis</a></li>
        <li>Entre outros.</li>
    </ul>
    <br>
    <p>Atenciosamente, <br>Seu Nome</p>

</body>
</html>
'''
# email.Attachments.Add(anexo) # o anexo pode ser uma base, imagem, apresentação, ...
email.Send()

print("E-mail enviado com sucesso!")