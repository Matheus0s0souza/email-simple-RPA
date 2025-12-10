import win32com. client as win32  #importar a bíblioteca win32
import datetime
import time

hora_envio = "00:28"

while True:
    agora = datetime.datetime.now().strftime("%H:%M")

    if agora == hora_envio:

        outlook = win32.Dispatch('outlook.application')
        email = outlook.CreateItem(0)
        email.To = 'silvasouzamatheus230@gmail.com'
        email.Subject = 'Relatório diário'

        faturamento = 1500
        qtd_produtos = 10
        ticket_medio = faturamento / qtd_produtos
        vendedor = "Matheus"

        email.HTMLBody = f"""
            <p>Olá, {vendedor}!</p>
            <p>Esta é uma mensagem automática</p>
            <br>
            <br>
            <p>Este mês você vendeu {qtd_produtos} produtos resultando no faturamento de {faturamento}</p>
            <p>sendo assim, seu ticket médio foi de {ticket_medio}</p>

            """

        anexo = 'C:/Users/silva/Downloads/Captura de tela 666.png'
        email.Attachments.Add(anexo)

        email.Send()
        print("E-mail enviado!")

        time.sleep(60) 

    time.sleep(10)




