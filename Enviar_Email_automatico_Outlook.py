
import win32com.client as win32

outlook = win32.Dispatch('outlook.application')  # criar a integração com o outlook do pc

email = outlook.CreateItem(0)  # OBS. você precisa criar um email e ter o email logado no Outlook do pc

email.To = "Digite aqui o Email de recebimento"

email.Subject = "E-mail automático do Python"

email.HTMLBody = f"""
<p>Olá, aqui é sua planilha automatica</p> 
<p>Abs,</p>
<p>Seu Código Python</p>
"""
anexo = "C://Users/Usuario/Downloads/Planilha.xlsx"

email.Attachments.Add(anexo)

email.Send()

print("Email Enviado!")
