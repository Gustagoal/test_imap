from imap_tools import MailBox , AND
import pandas as pd



login = ''
senha = ''

meu_email = MailBox('imap.gmail.com').login(login,senha)
tabela = pd.read_excel('produtos.xlsx')

for marca in tabela['Marca_Produto']:
    lista_email = meu_email.fetch(AND(subject={marca}))
    for email in lista_email:
        if marca == email:
            tabela['conferir']= 'OK'
        else:
             tabela['conferir']= 'NÃO CONFERE'

tabela.to_excel('conferido.xlsx')


