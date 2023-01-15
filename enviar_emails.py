from openpyxl import load_workbook
from email.message import EmailMessage
import smtplib

EMAIL = 'E-MAIL'
SENHA = "SENHA GERADA PELO GOOGLE"

# Carregando a base de dados
planilha = load_workbook("ARQUIVO-EXCEL")
folha_ativa = planilha.active

# Procedimento obrigatório do google
s = smtplib.SMTP('smtp.gmail.com: 587')
s.starttls()

# logando no e-mail
s.login(EMAIL, SENHA)

for i in range(0, 500):
    # Pegando os e-mails de contato
    contacts = [str(folha_ativa[f"F{i}"].value)]

    # Criando a nova mensagem de texto
    msg = EmailMessage()
    msg['Subject'] = 'ASSUNTO DO E-MAIL'
    msg['From'] = EMAIL
    msg['To'] = ', '.join(contacts)
    msg.set_content('MENSAGEM')

    # Abrindo o PDF e pegando os dados
    with open(f'{str(folha_ativa[f"B{i}"].value)}.pdf', 'rb') as f:
        file_data = f.read()
        file_name = f.name

    # Adicionando o Anexo
    msg.add_attachment(file_data,
                       maintype='application',
                       subtype='octet-stream',
                       filename=file_name
                       )

    # Enviando o e-mail
    s.sendmail(EMAIL, contacts[0], msg.as_string().encode('utf-8'))

    print(f'{i-1} - Mensagem enviada para {str(folha_ativa[f"B{i}"].value)}!')
