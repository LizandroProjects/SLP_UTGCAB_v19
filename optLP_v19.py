"""
CÓDIGO DE OTIMIZAÇÃO LINEAR SEQUENCIAL DA UCGCAB [optLP.py]

AUTORES: LADES(PEQ-COPPE) / UFF / CENPES 

EQUIPE: Carlos Rodrigues Paiva
        Thamires
        Roymel
        Cesar
        Lizandro Santos
        Dorigo
        Argimiro
        
ÚLTIMA MODIFICAÇÃO: 11 DE JANEIRO DE 2024, 11H 25 MIN.
***********************************************************************************************************************************

[1] DESCRIÇÃO: optLP: Script para Otimização usando a Programação Linear (com estrtégia iterativa)

[2] EXPLICAÇÃO: Essa é a rotina principal do programa para otimização da UTGCAB usando a estrtatégia de programação linear. Essa 
rotina é dividida nas seguintes funções:
    
    Inputdata -> Função de leitura de dados de entrada a partir da planilha excel input na pasta de projeto
    Hysysconect -> função usada para conecção com o HYSYS
    InputObjects -> função usada para criação dos objetos de comunicação
    SpecVar -> função usada para especificação de variáveis
    SpecLP -> função usada para especificações do LP [valores dos limites e cargas] de acordo com sintaxe do PULP
    SLP -> função de chamada da função de otimização Programação Linear
    
   
[3] OBSERVAÇÕES: Ao longo do código adicionaremos algumas anotações para facilitar a compreensão das considerações utilizadas

"""

#====================================================================================================================

'Importação das Bibliotecas'
import matplotlib.pyplot as plt # Biblioteca para Plotagem
import numpy as np  # Bliblioteca matemática numpy
from time import sleep # Usar a função sleep para atrasar o cálculo seguinte, se necessário
from functions_v19 import (plot_manipuladas, # plotar manipuladas
                           Inputdata,  # Leitura de dados
                           Hysysconect,  # COnexão com o Hysys
                           InputObjects,# Criação de objetos
                           SpecVar,  # Especificação de Variáveis
                           SpecLP,  # Montagem das variáveis do LP (PULP)
                           SLP, # Programação Linear Sequencial
                           Sim_rigorosa, # Rodar simulação Rigorosa
                           f_Rel,   # Geração de Relatório da Otimização
                           f_Plot  # Plotando Resultados
                           )

'''
****************************************************************************************************************
NOTA[1]: O descritivo das funções está explicado no início de cada função! CHECAR NO ARQUIVO FUNCTION_V(X).PY, onde (X) é o número da versão;
NOTA[2]: Para todas as funções, organizamos as informações em dicionários, de modo a simplificar a passagem de arumentos entre as funções;
NOTA[3]: Toda função possui uma flag de saída, para indicar códigos de erros possíveis;
****************************************************************************************************************
'''


'ETAPA [1] -> Conexão com o Hysys para rodar a Simulação Rigorosa'
'****************************************************************************************************************'
print('ETAPA [1]: CONEXÃO COM HYSYS PARA RODAR SIMULAÇÃO RIGOROSA')
print('**********************************************************')
filename = 'UTGCAB_original.hsc' # VERSÃO ATUAL DO ARQUIVO HYSYS DA SIMULAÇÃO RIGOROSA
hconect, simCase, hyApp = Hysysconect(filename) # Conexão com hysys
resultado_rigorosa, R_especs = Sim_rigorosa(simCase) # Rodando a simulação rigorosa
print('Fechando simulação rigorosa')
hyApp.Quit() # Fechando a simulação rigorosa
print('')
print('')

'ETAPA [2] -> Leitura de Dados de Entrada'
'****************************************************************************************************************'
code_input, edata = Inputdata() # Função de leitura de dados de entrada a partir da planilha excel input na pasta de projeto
print('ETAPA [2]: LEITURA DE DADOS')
print('***************************')
print('Leitura de dados realizada')
print('')
print('')

'ETAPA [3] -> Conexão com o Hysys para rodar a Simulação Essencial'
'****************************************************************************************************************'
print('ETAPA [3]: CONEXÃO COM HYSYS PARA RODAR SIMULAÇÃO ESSENCIAL')
print('***********************************************************')
sleep(15)  # Aguardar 15 segundos para nova conexão....da simulação Essencial (FAÇO ISSO PARA EVITAR ERRO NA COMUNICAÇÃO)
filename = 'LP_19_vM_(newcomp).hsc'   #VERSÃO ATUAL DO ARQUIVO HYSYS DA SIMULAÇÃO ESSENCIAL'
hconect, simCase, hyApp = Hysysconect(filename) # Conexão com hysys
print('Conexão realizada')
print('')
print('')
'NOTA: A simulação ESSENCIAL não será fechada, como fizemos na simulação rogorosa. Nesse caso, ela será sempre utilizada!'


'ETAPA [4] -> Criação de Objetos organizados em dicionários para a comunicação entre o python e hysys'
'****************************************************************************************************************'
print('ETAPA [4]: CRIAÇÃO DOS OBJETOS PARA COMUNICAÇÃO PYTHON-HYSYS')
print('************************************************************')
out_obj = InputObjects(simCase) # Criação dos objetos para a comunicação
print('Criação dos objetos realizada')
print('')
print('')

'ETAPA [5] -> Especificação de variáveis das unidades'
'****************************************************************************************************************'
print('ETAPA [5]: ESPECIFICAÇÃO DAS VARIÁVEIS DAS UNIDADES')
print('***************************************************')
out_espvar = SpecVar(edata, out_obj, R_especs) # Rotina para especificação das variáveis
print('Variáveis especificadas')
print('')
print('')

'ETAPA [6] -> Definição das variáveis utilizadas no otimizador PULP (de acordo com a sintaxe da toolbox)'
'****************************************************************************************************************'
print('ETAPA [6]: DEFINIÇÃO DOS VALORES DAS VARIÁVEIS UTILIZADAS NA LP (PULP)')
print('************************************************************************')
cod_speclp, R_min, R_max, R_cap, Carga = SpecLP(edata, out_obj) # especificações do LP [valores dos limites e cargas]
print('Variáveis definidas')
print('')
print('')

'ETAPA [7]-> Programação Linear Sucessiva (SLP) usando o PULP'
'****************************************************************************************************************'
print('ETAPA [7]: PROGRAMAÇÃO LINEAR SUCESSIVA usando a bilioteca PULP')
print('***************************************************************')
# Tipo de Otimização: Obj_type = 'Custo' (R=0 C=1); 'Receita' (R=1, C=0); 'Margem' (R=1, C=1) (R e C são variáveis binárias)
Obj_type = 'Margem'
cod_SLP, model, rel_SLP = SLP(simCase, edata, out_obj, R_min, R_max, R_cap, Carga, Obj_type) # chamada da função de otimização Programação Linear Sequencial
print('Otimização realizada')
print('')
print('')


'ETAPA [8] -> Relatório de Resultados [EM andamento]'
'****************************************************************************************************************'
print('ETAPA [8]: RELATÓRIO DE RESULTADOS')
print('***************************************************************')
print('Escrevendo relatório de resultados')
print('')
print('')
# code_rel = f_Rel(model, rel_SLP)   
print('Relatório Escrito no arquivo Output_Data.xlx')  

'ETAPA [9] -> Plotagem de Resultados [Em Andamento]'
'****************************************************************************************************************'
print('ETAPA [9]: PLOTAGEM DE RESULTADOS')
print('***************************************************************')
print('Plotando os resultados')
print('')
print('')
# code_f_Plot = f_Plot(model, rel_SLP)   
print('Gráficos Gerados')  


