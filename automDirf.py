import pyautogui as pag #install lib
import time
import os
#from dotenv import load_dotenv, find_dotenv
from decouple import config

#install Pillow
#install python-dotenv
#install python-decouple

#set variaveis de ambiente
PATH_PDF = config('DIR_PDF')
TXT_FILE_CPF = config('PATH_OF_TXT')
LOGO_DIRF = config('LOGO_DIRF')
SELEC_CNPJ = config('SELEC_CNPJ')
#PATH_EXEC = config("DIRF_EXEC")
#print(PATH_EXEC)

#constantes uteis
dirfExec =      '"C:/Arquivos de Programas RFB/Dirf2021/pgdDirf.exe"'


def PauseStep():
    pag.PAUSE = 3 # tempo de pausa para cada chamada do pyautogui


def ReadFile(path):
    #abrindo arquivo texto com cpfs
    with open(path) as t:
        content = t.readlines()
    return content


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
    locationImg = pag.locateOnScreen(img)
    while (locationImg != None):
        print(locationImg)
        locationImg = pag.locateOnScreen(img)


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


def FindBenefic(i):
    #encontrar beneficiário
    pag.moveTo(367, 178)
    pag.click()
    pag.write(i)
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


def SavePDF(i):
    pag.PAUSE = 1
    #salvando arquivo PDF em diretório (estático)
    pag.moveTo(169,43)
    pag.click()
    pag.moveTo(558, 486)
    pag.click()
    pag.write(PATH_PDF+'arquivo'+i)
    pag.moveTo(812, 544)
    pag.click()
    pag.moveTo(1193, 16)
    pag.click()


def GetCPF(content):
    #loop no arquivo texto com os cpfs para gerar o PDF
    for i in content:
        EnterOpImpress()
        FindBenefic(i)
        SendImpress()
        SavePDF(i)


#chamando funções para execução
outputContent = ReadFile(TXT_FILE_CPF)
OpenExec(dirfExec)
ClickCNPJ()
GetCPF(outputContent)
