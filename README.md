# Automação de envio de e-mails
## Olá! Esta é uma automação criada por mim para enviar muitas declarações por e-mail.

O script lê uma base de dados em excel, e envia as declarações geradas com base no nome de quem irá receber.

### Bibliotecas usadas: <br>
![BADGE](https://img.shields.io/badge/openpyxl-v3.0.10-blue) : Para a leitura da base de dados <br>
![BADGE](https://img.shields.io/badge/EmailMessage-red) : Para a formatação do e-mail que desejamos enviar<br>
![BADGE](https://img.shields.io/badge/smtplib-yellow) : Para termo acesso ao cliente de protocolo SMTP e enviarmos e-mails para outras máquinas da rede.<br>

Para o script funcionar você deve gerar uma senha de app na aba de seguranças da sua conta do google. <br>
É esta senha que você irá usar como password no script, pois o google não permite usar a senha real em outros aplicativos.  <br>
A base de dados contém os dados que serão lidos tanto para a geração das declarações através de mala direta no próprio word ( Para automatizar a
geração usei um script em vba ), quanto para o envio dos e-mails através do meu script. 

