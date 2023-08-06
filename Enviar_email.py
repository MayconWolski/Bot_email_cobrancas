import win32com.client
import pandas as pd
import datetime as dt

#Lendo o arquivo excel

tabela = pd.read_excel("Contas a Receber.xlsx")

#verificar data de hoje

hoje = dt.datetime.now()
print(hoje)

# Coletando os dados de clientes que estão devendo 

tabela_devedores = tabela.loc[tabela['Status'] == 'Em aberto']
print(tabela_devedores)
tabela_devedores = tabela_devedores.loc[tabela_devedores['Data Prevista para pagamento'] < hoje]
print(tabela_devedores)

#como enviar um email via Outlook

outlook = win32com.client.Dispatch("Outlook.Application")
emissor = outlook.session.Accounts['emaildeenvio@'] #Aviso


dados= tabela_devedores[['Valor em aberto','Data Prevista para pagamento','E-mail','NF']].values.tolist()

# Enviando o e-mail para todos os destinatarios

for dado in dados:
    destinatario = dado[2]
    nf=dado[3]
    prazo=dado[1]
    prazo = prazo.strftime("%d/%m/%Y")
    valor = dado[0]
    assunto = 'Atraso de pagamento teste'
    mensagem = outlook.CreateItem(0)
    mensagem.Display()
    mensagem.To = destinatario
    mensagem.Subject = destinatario
    corpo_mensagem = f'''
    Prezado Cliente,

    Verificamos um atraso no pagamento referente a NF {nf} com vencimento em {prazo} e valor total de R${valor:.2f}.
    Gostaríamos de verificar se há algum problema que necessite de auxílio de nossa equipe. 

    Em caso de dúvidas, é só entrar em contato com nosso time atráves do e-mail teste@gmail.com

    
    Att,
    Equipe teste
    '''
    mensagem.Body = corpo_mensagem
    mensagem.Save()
    mensagem.Send()
    
