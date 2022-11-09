from audioop import reverse
from datetime import date, datetime, timedelta
import win32com.client as win32
import os

ontem = (datetime.now()-timedelta(1)).isoformat()[2:10] 
print("Data de Ontem: ",ontem)

#caminho da pasta 
caminho  = "C://Users/JuniorE/OneDrive - Novelis Inc/Documents/github/TarefaPythonEnvioEMail"
#listar arquivos do diretorio
lista_arquivos = os.listdir(caminho)
#lista 
lista_NomeArquivo = []

#para cada arquivo em lista_arquivos
for arquivo in lista_arquivos:

    #pegar o timestamp dele
    dataTimestamp = os.path.getmtime(f"{caminho}/{arquivo}")
    #converter esse timestamp em uma data
    dataConvertida = datetime.fromtimestamp(dataTimestamp).isoformat()[2:10]
    #adicionar na lista a data e o nome do arquivo
    lista_NomeArquivo.append((dataConvertida, arquivo ))

print("\nArquivos do diretorio: ")
for lista in lista_NomeArquivo:
    print(lista)

i=0
contagem = 0
contagem = len(lista_NomeArquivo)
print("\nTotal de arquivos na pasta: ",contagem)

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = "elias.junior@novelis.com"

mail.Subject = 'Envio de Erros do Dia Anterior'
mail.Body = 'Segue Anexado Arquivos XML de erros'
print("\nArquivos anexados: ")
for data, nome in lista_NomeArquivo:
    if data == ontem:
        print(data, nome)
        mail.Attachments.Add(caminho+"/"+nome)


mail.Send()

print ("\n\nEmail enviado!")
