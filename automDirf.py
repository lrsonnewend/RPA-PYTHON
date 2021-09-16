import pyautogui as pag #lib para tratar as automatizações desktop
import time #lib para tratar tempos de pausa
import os #lib para manipular interações com o S.O
from decouple import config #lib para leitura de arquivo .env
import csv #lib para manipulação de .csv
from RPA.Desktop import Desktop #lib para tratar as automatizações desktop
desktop = Desktop() #criando instancia da lib

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
#install rpaframework

#set variaveis de ambiente
PATH_PDF = config('DIR_PDF')
FILE_CPF = config('PATH_OF_CSV')
LOGO_DIRF = config('LOGO_DIRF')
SELEC_CNPJ = config('SELEC_CNPJ')
USER_MAIL = config('LOGIN_MAIL')
PASS_MAIL = config('PASS_MAIL')
DIRF_EXEC = config('DIRF_EXEC')
IMPRESSORA_IMG = config('IMPRESSORA_IMG')
COMP_RENDIMENTO_IMG = config('COMPROVANTE_RENDIMENTO_IMG')
COMP_BENEFIC_IMG = config('COMPROVANTE_BENEFIC_IMG')
LBL_CPF_IMG = config('FIELD_CPF_IMG')
BTN_EXEC_IMG = config('BUTTON_EXEC_IMG')
HEADER_IMG = config('HEADER_IMG')
BTN_SAVE_IMG = config('BUTTON_SAVE_IMG')
BTN_CLOSE_IMG = config('BUTTON_CLOSE_IMG')
ICON_SAVE_IMG = config('ICON_SAVE_IMG')


def PauseStep():
    pag.PAUSE = 3 # tempo de pausa para cada chamada do pyautogui



def OpenExec(path):
    #abrindo o programa Dirf
    try:
        desktop.open_application(path)

    except (OSError, err):
        print(f'Erro ao abrir programa: {err}')



def ClickCNPJ():
    time.sleep(3)
    #verifica se o programa ainda está carregando
    CheckLoadProg(LOGO_DIRF)
    #acessando declaração
    time.sleep(1)
    pag.click(SELEC_CNPJ)
    time.sleep(4)



def CheckLoadProg(img):
    location_img = pag.locateOnScreen(img)
    while (location_img != None):
        print(location_img)
        location_img = pag.locateOnScreen(img)



def EnterOpImpress():
    time.sleep(4)
    #entrando na opção impressão
    #pag.moveTo(163, 60)
    pag.click(IMPRESSORA_IMG)
    time.sleep(1)
    pag.moveTo(pag.locateOnScreen(COMP_RENDIMENTO_IMG))
    time.sleep(0.5)
    pag.click(COMP_BENEFIC_IMG)
    time.sleep(1)



def FindBenefic(num_cpf):
    #encontrar beneficiário
    #pag.moveTo(367, 178)
    pag.click(LBL_CPF_IMG)
    pag.write(num_cpf)
    pag.click(BTN_EXEC_IMG)
    coord_img = pag.locateOnScreen(HEADER_IMG)
    pag.moveTo(coord_img[0],coord_img[1])
    time.sleep(1)
    pag.doubleClick(coord_img[0], coord_img[1]+25)
    time.sleep(4)
    


def SavePDF(num_cpf):
    pag.PAUSE = 2
    #salvando arquivo PDF em diretório (estático)
    pag.click(ICON_SAVE_IMG)
    path_aux = PATH_PDF+'arquivo'+num_cpf+'.pdf'
    pag.write(path_aux)
    pag.click(BTN_SAVE_IMG)
    pag.click(pag.locateCenterOnScreen(BTN_CLOSE_IMG))
    return path_aux



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

    try:
        #montando email
        body = """
            <p>Olá, segue em anexo o seu comprovante de beneficiário</p>
            <p>Att,</p>
            <p>RPA Python</p>
        """
        mail = MIMEMultipart()
        mail['From'] = login
        mail['To'] = dest
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

    except err:
        print(f'Erro ao enviar email: {err}')    



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
                path_anexo = SavePDF(i[0])
                #funcao para enviar email
                SendMailSMTP(i[1], path_anexo)



def main():
    #iniciando RPA
    OpenExec(DIRF_EXEC)
    ClickCNPJ()
    InsertCPFOnApp()


if __name__ == "__main__":
    main()



