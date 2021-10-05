import os
import time
import shutil
from colorama import init, Fore

def toRaw(string):
    #Função que converte a string em Raw string
    return fr"{string}"

init()
#Colorama init


print(Fore.MAGENTA + "Zip Zip!")
time.sleep(0.5)
print("Você acabou de entrar no Zipnator da Miralt")
time.sleep(1.5)
dir = input("Cole aqui o diretorio principal\n")
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
    #identificação e separação dos excels dos outros arquivos
    
    shutil.make_archive(pdfFolder[0], 'zip', pdfFolder[0])
    #Conversão do arquivo de pdf para zip
    print(pdfFolder[0])

print("Todos Zaper Ziper foram Zipados")
time.sleep(1)
print("Bora pra cima!")
time.sleep(5)