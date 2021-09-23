import os
from docx2pdf import convert
from win32com import client
import time
from colorama import init, Fore

def toRaw(string):
    #Função que converte a string em Raw string
    return fr"{string}"

def excelToPdf(excelPath, pdfFolder):
    #Função que converte arquivos xlsx em pdf


    path = excelPath[:-5]
    #retirando a extensão .xlsx
    
    slash = path.count("\\")
    pathlist = path.split("\\")
    excelName = pathlist[slash]
    #Separação do path em uma lista
    #para a obtenção do nome do arquivo
    
    print(excelName)
    
    pdfPath = f"{pdfFolder}\{excelName}" 
    #print do caminho do arquivo que vai ser convertido
    
    app = client.DispatchEx("Excel.Application")
    #Instância do excel

    app.Interactive = False
    app.Visible = False
    #Configuração do excel
    
    Workbook = app.Workbooks.Open(excelPath)
    #Instância da Planilha

    try:
        Workbook.ActiveSheet.ExportAsFixedFormat(0, pdfPath)
        #Conversão em pdf
    except Exception as e:
        print(f"Deu erro nesse {excelName}")
        print(str(e))
        time.sleep(10)
        #Tratativa caso de algum erro
    finally:
        Workbook.Close()
        app.Quit()
        #Encerramento das instâncias


init()
#Colorama init


print(Fore.MAGENTA + "Bem vindo!")
time.sleep(0.5)
print("Você acabou de entrar na Fabrica de PDFs Miralt")
time.sleep(1.5)
dir = input("Cole aqui o diretorio principal: ")
#input do diretorio 

toRaw(dir)
#Validação do input

unsortedSubfolders = [f.path for f in os.scandir(dir) if f.is_dir()]
subfolders = sorted(unsortedSubfolders, key=len)
# f = folder
#Pegando todos os sub diretorios do input

time.sleep(1)
print("Processando...")
for subfolder in subfolders:
    pdfFolder = [f.path for f in os.scandir(subfolder) if f.is_dir()]
    excels = [f.path for f in os.scandir(subfolder) if f.name.endswith("xlsx") or f.name.endswith("xlsm")]
    #identificação e separação dos excels dos outros arquivos
    
    
    if len(pdfFolder) == 0:
        pdfFolder = [subfolder]


    for excel in excels:
        excelToPdf(excel, pdfFolder[0])
        #conversão em pdf
    convert(subfolder, pdfFolder[0])
    #conversão em pdf de todos arquivos word dos subdiretorios

print("\nTodos os arquivos foram convertidos")
time.sleep(1)
print("Bora pra cima!")
time.sleep(5)