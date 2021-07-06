
from copy import deepcopy
import pandas as pd 
import math
import random
import numpy
import os
import time

t_inicio = time.ctime()
#def o tamanho da população
tamPop = 200
#parâmetros do problema
dim = 20
w = 0.73
c1 = 1.68#1.68#1.49445#1.8
c2 = 1.68#1.68#1.49445#2.2
limSup = 10
limInf = -10
numGeracoes = 500

##matriz = pd.read_csv("Oliver30.DAT", header=None)
##print(matriz)
###input("Teste")

maxGlobal = []
listaFitness = []

for i in range (dim+1):
    maxGlobal.append(math.inf)
    
maxParticulas = [0]
for i in range(tamPop):
    lista = []
    for j in range(dim+1):
        lista.append(math.inf)
   
    maxParticulas.append(lista)
    #print(maxGlobal)
#print(maxParticulas)
#inicializando pop inicial
posicao = [0]
velocidade = [0]
for i in range(1, tamPop+1):
    listap=[0]
    listav=[0]
    for j in range(1, dim+1):
        listap.append(random.uniform(limInf, limSup))
        listav.append(0)
    posicao.append(listap)
    velocidade.append(listav)
    #print('Posição: {}'.format(posicao))
   # print('velocidade: {}\n'.format(velocidade))

###################################################################################################
# Lendo arquivo de Saída do Recnod (salvando as variáveis para fator de pico e concentração de boro)
def lendo_boro_pico():
    aux = []
    arquivo = open('Jessica_Pico_Boro.dat', 'r')
    content = arquivo.read().split("\n")
    for i in range(len(content) - 1):
        linha2 = content[i].replace(',', '.')
        aux.append(float(linha2))
    if aux[0] == 10:
        pico = aux[0]
    else:
        pico = aux[0]/1000
    boro = aux[1]
    arquivo.close()

    return pico, boro
###################################################################################################

###################################################################################################
# Criando arquivo de entrada do recnod
def entrada(ninho_x):
    ninho_1 = []
    ninho_2 = []
    posicoes_1 = []
    posicoes_2 = []
    posicoes = numpy.zeros((20, dim))
    for i in range(1,int(len(ninho_x))):
        ninho_1_aux = []
        ninho_2_aux = []
        for j in range(int(len(ninho_x[i]) / 2)):
            ninho_1_aux.append(ninho_x[i][j])
            ninho_2_aux.append(ninho_x[i][j + 10])
        ninho_1.append(ninho_1_aux)
        ninho_2.append(ninho_2_aux)

    for i in range(1,len(ninho_x[1])):
        posicoes_1_aux = []
        posicoes_2_aux = []
        for j in range(int(len(ninho_x[i])/2)):
            indice_1 = ninho_1[i].index(max(ninho_1[i]))
            indice_2 = ninho_2[i].index(max(ninho_2[i]))
            posicoes_1_aux.append(indice_1 + 1)
            ninho_1[i][indice_1] = -1e+9
            posicoes_2_aux.append(indice_2 + 11)
            ninho_2[i][indice_2] = -1e+9
        posicoes_1.append(posicoes_1_aux)
        posicoes_2.append(posicoes_2_aux)
        #print('posicoes 1',posicoes_1)
        #print('posicoes 2',posicoes_2)
    for i in range(20):
        for j in range(10):
            posicoes[i][j] = posicoes_1[i][j]
            posicoes[i][j + 10] = posicoes_2[i][j]

    # print('Entrada_1: {}'.format(posicoes_1))
    # print('Entrada 2: {}'.format(posicoes_2))
    # print('Criando Entradas... entradas = {}\n'.format(posicoes))
    return posicoes
###################################################################################################

###################################################################################################
###################################################################################################
# Função objetivo
def f_obj(posicao):
    fit = numpy.ones((tamPop, 1))
    listafit=[]
    entradas = entrada(posicao)
    #print('entradas',entradas)
    for i in range(len(entradas)):
        arq = open('Jessica_Saida.dat', 'w')
        for j in range(len(entradas[i])):
            arq.write('{}\n'.format(str(int(entradas[i][j]))))
        arq.close()
        try:
            os.remove('Jessica_Pico_Boro.dat')
        except:
            pass
        os.startfile('Jessica.exe')
        while True:
            try:
                arq = open('Jessica_Pico_Boro.dat')
                content = arq.read().split("\n")
                arq.close()
                if len(content) >= 2:
                    break
            except:
                time.sleep(1.5)

        pico, boro = lendo_boro_pico()
        if pico <= 1.395:
            fit_aux = 1.0 / boro
        else:
            fit_aux = pico
        listafit.append(fit_aux)
        fit = fit_aux
        #print('fit',fit)
        # print('Entradas: {}'.format(entradas))
        # print('Fitness: {}'.format(fit))
    return min(listafit), entradas
###################################################################################################

###################################################################################################

#print(f_obj(posicao))

 
def avalia():
    posicaoInt = deepcopy(posicao)
              
    fitness, entradas = f_obj(posicaoInt)
    #mini = min(fitness[0])
    #print('fitness', fitness)
    #print('entradas', entradas)
    lista = [math.inf]

    for i in range(1,tamPop+1):
        
    
        if fitness < maxGlobal[0]:
            maxGlobal[0] = fitness
            for k in range(1,dim+1):
                maxGlobal[k]=posicao[i][k]
        
        if fitness < maxParticulas[i][0]:
            maxParticulas[i][0]=fitness
            for k in range(1,dim+1):
                maxParticulas[i][k]=posicao[i][k]
        
        lista.append(fitness)
##    
##    #print('MaxParticula: {}'.format(maxParticulas))
##    print('MaxGlobal: {}\n'.format(maxGlobal[0]))
##    #print(lista, '\n')
##    
    return lista

listaFitness = avalia()


for geracao in range(1,numGeracoes+1):
    #print('Posição: {}'.format(posicao))
    for i in range(1,tamPop+1):
        for j in range(1, dim+1):
            
            velocidade[i][j]= w*velocidade[i][j] + c1*random.random()*(maxParticulas[i][j] - posicao[i][j]) + c2*random.random()*(maxGlobal[j]-posicao[i][j])
            #print (velocidade[i][j])
            #abc = input("Aguardando")
            if velocidade[i][j] > (limSup-limInf)*0.2:
                velocidade[i][j]=(limSup-limInf)*0.01
            
            if velocidade[i][j] < (limInf-limSup)*0.2:
                velocidade[i][j]=(limInf-limSup)*0.01
            
            #print('velocidade: {}\n'.format(velocidade))
            
            posicao[i][j]=posicao[i][j]+velocidade[i][j]
            if posicao[i][j] < limInf:
                a1=posicao[i][j]%limInf
                posicao[i][j]=limInf+a1
            
            if posicao[i][j] > limSup:
                a1=posicao[i][j]%limSup
                posicao[i][j]=limSup-a1
    
    listaFitness = avalia()
    #print ('Geração: {:03} -> maxGlobal: {}'.format(geracao, maxGlobal))
    #print(velocidade[1:5])
    #print(listaRK)
    print(min(listaFitness))
    print('geracao', geracao)

 
