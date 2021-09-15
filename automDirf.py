import pyautogui as pag #lib para tratar as automatizações desktop
import time #lib para tratar tempos de pausa
import os #lib para manipular interações com o S.O
from decouple import config #lib para leitura de arquivo .env
import csv #lib para manipulação de .csv

#libs para envio de email
import smtplib as smtp
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase

#lib para transformar arquivos em outras bases
from email import encoders

#install pyautogui
#install Pillow
#install python-decouple

#set variaveis de ambiente
PATH_PDF = config('DIR_PDF')
FILE_CPF = config('PATH_OF_CSV')
LOGO_DIRF = config('LOGO_DIRF')
SELEC_CNPJ = config('SELEC_CNPJ')
USER_MAIL = config('LOGIN_MAIL')
PASS_MAIL = config('PASS_MAIL')
DIRF_EXEC = config('DIRF_EXEC')
DIRF_EXEC = "\"'"+DIRF_EXEC+"\"'""
print(DIRF_EXEC)
#constantes uteis
dirf_exec = '"C:/Arquivos de Programas RFB/Dirf2021/pgdDirf.exe"'


def PauseStep():
    pag.PAUSE = 3 # tempo de pausa para cada chamada do pyautogui



def OpenExec(path):
    #abrindo o programa Dirf
    os.system(path)



def ClickCNPJ():
    time.sleep(3)
    #verifica se o programa ainda está carregando
    CheckLoadProg(LOGO_DIRF)
    #acessando declaração
    time.sleep(1)
    pag.click(SELEC_CNPJ)
    time.sleep(3)



def CheckLoadProg(img):
    location_img = pag.locateOnScreen(img)
    while (location_img != None):
        print(location_img)
        location_img = pag.locateOnScreen(img)



def EnterOpImpress():
    time.sleep(4)
    #entrando na opção impressão
    pag.moveTo(163, 60)
    pag.click()
    time.sleep(1)
    pag.moveTo(202, 122)
    time.sleep(0.5)
    pag.moveTo(359, 141)
    pag.click()
    pag.PAUSE = 0.8



def FindBenefic(num_cpf):
    #encontrar beneficiário
    pag.moveTo(367, 178)
    pag.click()
    pag.write(num_cpf)
    pag.moveTo(870, 561)
    pag.click()
    pag.moveTo(465, 255)
    pag.click()
    time.sleep(1)
    


def SendImpress():
    #enviar para impressão
    pag.moveTo(786, 601)
    pag.click()
    time.sleep(1)



def SavePDF(num_cpf):
    pag.PAUSE = 1
    #salvando arquivo PDF em diretório (estático)
    pag.moveTo(169,43)
    pag.click()
    pag.moveTo(558, 486)
    pag.click()
    path_aux = PATH_PDF+'arquivo'+num_cpf+'.pdf'
    pag.write(path_aux)
    pag.moveTo(812, 544)
    pag.click()
    pag.moveTo(1193, 16)
    pag.click()
    return path_aux



'''def SendEmail(dest, anexo):
    #criando integraçao com outlook
    outlook = win32.Dispatch('outlook.application')

    #criando instancia de email
    email = outlook.CreateItem(0)

    #configurando informações
    email.To = dest
    email.Subject = "Teste envio de relatório"
    email.HTMLBody = """
        <p>Olá, segue em anexo o seu comprovante de beneficiário</p>
        <p>Att,</p>
        <p>RPA Python</p>
    """
    email.Attachments.Add(anexo)
    email.Send()
    print('email enviado!')'''



def SendMailSMTP(dest, anexo):
    #configurações de acesso servidor SMTP
    host= 'smtp.office365.com'
    port= '587'
    login= USER_MAIL
    senha = PASS_MAIL

    #config. servidor
    server = smtp.SMTP(host, port)

    #invocando TLS
    server.ehlo()
    server.starttls()
    
    #login no servidor
    server.login(login, senha)

    #montando email
    body = """
        <p>Olá, segue em anexo o seu comprovante de beneficiário</p>
        <p>Att,</p>
        <p>RPA Python</p>
    """
    mail = MIMEMultipart()
    mail['From'] = login
    mail['To'] = 'lucas.lr184@gmail.com'
    mail['Subject'] = 'Teste envio de relatório'
    mail.attach(MIMEText(body,'html'))

    #abrindo arquivo .pdf
    attachment = open(anexo, 'rb')
    
    #definindo nome do arquivo a ser anexado
    file_name = anexo[43:]

    #lendo arquivo e transformando em base64
    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename={file_name}')

    #colocando anexo no email
    mail.attach(part)

    #fechando arquivo
    attachment.close()
    
    #enviando email
    server.sendmail(mail['From'], mail['To'],mail.as_string())
    server.quit()



def InsertCPFOnApp():
    count = 0
    #lendo arquivo .csv
    #abrindo arquivo texto com cpfs
    with open(FILE_CPF) as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=';')
        #loop no arquivo texto com os cpfs para gerar o PDF
        for i in csv_reader:
            if(count == 0):
                count+=1
                
            else:
                EnterOpImpress()
                FindBenefic(i[0])
                SendImpress()
                path_anexo = SavePDF(i[0])
                #funcao para enviar email
                SendMailSMTP(i[1], path_anexo)





def main():
    #iniciando RPA
    OpenExec(dirf_exec)
    ClickCNPJ()
    InsertCPFOnApp()


if __name__ == "__main__":
    main()



