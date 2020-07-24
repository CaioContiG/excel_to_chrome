import pandas as pd
import numpy as np
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from os.path import dirname
pd.options.display.float_format = '{:.1f}'.format
np.set_printoptions(formatter={'all': lambda x: str(x)})
#Feito por Caio Conti e Manoel Anizio
#Qualquer problema contato: caiocontig@gmail.com


#acertando os drivers e abrindo o site-------------------------
pasta_certa = str(dirname(dirname(os.getcwd())))
chromedriver = f"{pasta_certa}/chromedriver"
driver = webdriver.Chrome(chromedriver)
driver.get("http://portalacademico.faacz.com.br:8020/Corpore.Net/Login.aspx")

#Variaveis Globais--------------------------------------------
Notas = []  #vetor com as notas na ordem da planilha
Matricula = []  #vetor com as matriculas na ordem da planilha
Matricula_Notas = [[], []]  #Matriz no programa que relaciona matricula com as notas
todas_planilhas = []  #Conterá todos os nomes das planilhas encontradas para o usuário selecionar
N_linha_planilha = -1  #Conterá número de linhas da planilha original

#definindo funcoes----------------------------------------------
def rec_walk(diretorio):
    contents = os.listdir(diretorio)  #read the contents of dir
    for item in contents:       #loop over those contents
            if item.endswith(".xlsx"):
                todas_planilhas.append(item)
#rec_walk encontra as planilhas de um diretorio

def escolher_planilha():
    global Notas, Matricula, Matricula_Notas, todas_planilhas, N_linha_planilha

    #procura as planilhas e mostra ao usuario
    rec_walk(dirname(dirname(os.getcwd())))
    print("Planilhas .xlsx encontradas:")
    for i in range(len(todas_planilhas)):
        print(f"{i + 1} - {todas_planilhas[i]}")

    #Usuario escolhe a planilha desejada
    while True:
        try:
            escolha_da_planilha = int(input("\nDigite o número da planilha desejada: "))
            break
        except ValueError:
            print("O valor digitado não é um número inteiro.")
            continue
        except:
            print("Passou do número total de planilhas.")
            continue

    print(f"A planilha escolhida foi:{todas_planilhas[escolha_da_planilha-1]}")

    #Processo para pegar os dados da planilha escolhida
    nome_da_planilha = str(todas_planilhas[escolha_da_planilha - 1])
    planilha = pd.read_excel(f"{dirname(dirname(os.getcwd()))}/{nome_da_planilha}")
    N_linha_planilha = len(planilha)

    #Teste se a planilha está correta:
    try:
        teste_erro = planilha.columns.get_loc("notas")
    except:
        print("ERRO: A Planilha não tem uma coluna com o título 'notas'. "
              "\nRenomeie para que a coluna contendo as notas esteja com o título: 'notas'"
              " exatamente, no minúsculo e no plural.")
        enter = input("Dê enter para fechar o programa.")
        driver.quit()
        exit()

    try:
        teste_erro = planilha.columns.get_loc("matricula")
    except:
        print("ERRO: A Planilha não tem uma coluna com o título 'matricula'. "
              "\nRenomeie para que a coluna contendo as matrículas esteja com o título: 'matricula'"
              " exatamente, tudo minúsculo e sem acento.")
        enter = input("Dê enter para fechar o programa.")
        driver.quit()
        exit()

    #Com a planilha correta pegamos seus resultados
    for i in range(N_linha_planilha):
        Notas.append(planilha["notas"][i])
        Matricula.append(planilha["matricula"][i])

    #Criando tabela de relacao notas com matricula
    Matricula_Notas = np.array([Matricula, Notas])
    Matricula_Notas = Matricula_Notas.reshape(2, N_linha_planilha)
    Matricula_Notas = np.transpose(Matricula_Notas)

def colocar_notas_chrome():
    global Matricula_Notas, N_linha_planilha

    #Esperando o usuario abrir o site manualmente-----------------------------------------------------------------
    tens = input("\nSe logue no sistema, vá para página que você quer inserir as notas da planilha pro site."
                 "\nDigite 's' e aperte enter se está na página que o programa vai inserir as notas automaticamente.\n")
    if tens != "s":
        print("Site nao foi aberto, aperte enter para fechar o programa")
        driver.quit()
        exit()

    else:
        print("\n")

        #tentando encontrar tabela, caso não encontre o programa fecha
        try:
            simpleTable = driver.find_element_by_id("ctl23_gvNotasFaltasEtapa")  #Pegando toda tabela
        except:
            print("Não estava na página certa para colocar as notas.")
            print("Aperte enter para fechar o programa.")
            driver.quit()
            exit()

        #com a tabela encontrada procura as celulas para poder escrever as notas e então escreve
        linhas = simpleTable.find_elements(By.TAG_NAME, 'tr')  #Pegando todas linhas
        encontrou = -1
        matricula_site = -1

        for i in range(len(linhas)-1):
            if encontrou == 0:
                print(f"A MATRÍCULA {matricula_site} NÃO FOI ENCONTRADA NA PLANILHA.")

            celulas_da_linha = linhas[i+1].find_elements(By.TAG_NAME, 'td')
            matricula_site = celulas_da_linha[1].text
            j = 0

            for j in range(N_linha_planilha):
                if int(matricula_site) == int(Matricula_Notas[j][0]):
                    caixa_da_nota = celulas_da_linha[4].find_elements(By.TAG_NAME, 'input')
                    caixa_da_nota[0].clear()
                    caixa_da_nota[0].send_keys(str(Matricula_Notas[j][1]).replace('.',','))
                    encontrou = 1
                    break
                else:
                    encontrou = 0


def main():
    global Notas, Matricula, Matricula_Notas, todas_planilhas, N_linha_planilha

    continuar = 's'
    while(continuar == 's'):
        #RESETANDO VARIAVEIS GLOBAIS--------------------------------------------------
        Notas = []  # vetor com as notas na ordem da planilha
        Matricula = []  # vetor com as matriculas na ordem da planilha
        Matricula_Notas = [[], []]  # Matriz no programa que relaciona matricula com as notas
        todas_planilhas = []
        N_linha_planilha = -1

        escolher_planilha()
        colocar_notas_chrome()
        print("\n----------------------------------------------------------------------------------------------")
        continuar = input("deseja colocar mais notas em outra turma?(s/n)\n")
    driver.quit()
    exit()


main()

