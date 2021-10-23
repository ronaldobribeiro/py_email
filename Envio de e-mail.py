import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


# Modulos para manipulação de email
import email.mime.application

# Cabeçalho da mensagem do email
msg = MIMEMultipart()
msg['Subject'] = 'MENSAGEM-AUTOMATICA: Receba sua tabela de dados'
msg['From'] = 'email remetente'
msg['To'] = 'email destino'

# Corpo principal do email (tb um anexo)
body = email.mime.text.MIMEText("""Segue, tabela analítica referente aos dados de AAO.
\n
Atualização recorrente sempre as 9:30h, todos os dias da semana.
\n
ATENCIOSAMENTE 
\n 
Desenvolvimento Spiltag""")
msg.attach(body)

# Anexando o PDF
arqname = 'Dados_AAO.csv'
pdfname = '//serv/Dados_AAO.csv'
fp = open(pdfname, 'rb')
anexo = email.mime.application.MIMEApplication(fp.read(), _subtype="csv")
fp.close()
anexo.add_header('Content-Disposition', 'attachment', filename=arqname)
msg.attach(anexo)

# Enviando via "fake" server
s = smtplib.SMTP('smtp.outlook.office365.com',587)
s.starttls()
s.login('remetente', 'senha')
s.sendmail('email', ['email']
           , msg.as_string())
s.quit()