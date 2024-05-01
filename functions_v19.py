# -*- coding: utf-8 -*-
"""

FUNÇÃO QUE CONTÉM AS FUNÇÕES AUXILIARES PARA O PROGRAMA

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
[1] DESCRIÇÃO: FUNÇÃO QUE CONTÉM AS FUNÇÕES AUXILIARES PARA O PROGRAMA

[2] FUNÇÕES:
    
    Sim_rigorosa -> Função utilizada para rodar a simualção rigorosa
    InputObjects -> função usada para criação dos objetos de comunicação
    SLP -> função de chamada da função de otimização Programação Linear
    SpecLP -> função usada para especificações do LP [valores dos limites e cargas] de acordo com sintaze do PULP
    SpecVar -> função usada para especificação de variáveis
    Hysysconect -> função usada para conecção com o HYSYS
    Inputdata -> Função de leitura de dados de entrada a partir da planilha excel input na pasta de projeto
    SLP -> função de chamada da função de otimização Programação Linear
    SimulaLP: função para rodar Simulação Essencial usando a Programação Linear
    SimulaLP_closed -> função para rodar Simulação Essencial com reciclo fechado  usando a Programação Linear
    Spec_prods -> função que obtem os valores das restrições de produtos
    plot_derivatives -> função para plotar as derivadas
    plot_manipuladas -> função para plotar as variáveis de decisão
    
"""   


import numpy as np # Bliblioteca matemática numpy
import matplotlib.pyplot as plt # Biblioteca para Plotagem
import pandas as pd # Bliblioteca da banco de dados Pandas
import os # Bliblioteca para mapear caminho de diretório
import win32com.client as win32  # Biblioteca para comunicação com hysys
from numpy import linalg as la # Bibliotecas de álgebra linear
from pulp import pulp, LpMaximize, LpProblem, LpStatus, LpVariable  # Biblioteca Pulp # Referencias: https://coin-or.github.io/pulp/index.html
import timeit # Contador de tempo computacional
import xlsxwriter # escrever em excel
from func_auxiliar import (aloca_cargas,  # funcções auxiliares para rodar a simulação rogorosa (peguei da versão implementada no servidor)
                            ler_config,
                            ler_inputs,
                            simula_detalhada, # função com a simulação detalhada
                            )

def Sim_rigorosa(simCase):
    
    "*** Lendo o arquivo de configurações ***"
    config = ler_config()

    "*** Lendo os dados de entrada ***"
    inputs = ler_inputs()

   # Realizando o procedimento de alocação de carga antes da simulação
    cargas = aloca_cargas(simCase, config, inputs)
    print("Procedimento de alocação de cargas realizado")

    # |# Simulando o caso base
    resultados_simulacao, R_especs = simula_detalhada(simCase, config, inputs, cargas)
    
    return resultados_simulacao, R_especs

def InputObjects(simCase):

    '''
    *************************************************************************************************************************************
    [1] DESCRIÇÃO: InputObjects: Sequential Linear Programming -> Rotina da Programação Linear Sucessiva
    
    [2] EXPLICAÇÃO: Essa é utilizada para realizar a otimização linear sequencial do processo. Um modelo linear, baseado nas equações
    de balanço de massa global, é utilizado e implementado, nessa versão, na toolbox PULP (https://coin-or.github.io/pulp/index.html).
    Uma explicação detalhada do procesimento matemático está documentada do arquivo ProgramaçãoLinearSucessiva.ppx da pasta do projeto.
    A função SLP deve receber algumas variáveis e parâmetros, para possibilitar a comunicação e troca de "informações" entre o Python e
    o Hysys, de modo a possibilitar a otimização. 
    
    [3] DADOS DE ENTRADA: 
        simCase -> Objeto resultante da comunicação entre Python e Hysys (usado para abrir, fechar e ou iniciar a simulação);
        edata   -> Dicionário contendo os valores de parâmetros e variáveis do arquivo de entrada Input.xls;
        obj     -> Dicionário contendo os objetos resultantes das variáveis e spreadsheets do hsysys que serão utilizados
        R_min   -> Dicionário contendo os valores mínimos das restrições das especificações de produtos
        R_max   -> Dicionário contendo os valores máximos das restrições das especificações de produtos
        R_cap   -> Dicionário contendo os valores das restrições de capacidade das unidades
        Carga   -> Dicionário contendo os valores das vazões de carga da Unidade
    
    [4] DADOS DE SAÌDA: 
        cod_SLP   -> Flag para indicar sucesso ou insucesso do cálculo
        rel_SLP   -> Dicionário contendo dados da otimização (iterações, derivadas, receitas, variáveis de decisão etc.)
        model     -> Objeto com todos os dados da otimização LP
        
    [5] OBSERVAÇÕES: Ao longo do código adicionaremos algumas anotações para facilitar a compreensão das considerações utilizadas
    *************************************************************************************************************************************
    '''
    
    '''Criando objetos auxiliares para importação-exportação de variáveis'''
    
    #====================================================================================================================
    Solver                 = simCase.Solver
    MT_main                = simCase.Flowsheet.MaterialStreams                       # correntes de materia do flowsheet principal
    MT_URGN                = simCase.Flowsheet.Flowsheets('TPL25').MaterialStreams   # correntes da URGN
    MT_URLI                = simCase.Flowsheet.Flowsheets('TPL4').Flowsheets('TPL22').MaterialStreams   # correntes da URL-I
    MMT_URLII              = simCase.Flowsheet.Flowsheets('TPL4').Flowsheets('TPL23').MaterialStreams   # correntes da URL-II
    MT_URLIII              = simCase.Flowsheet.Flowsheets('TPL4').Flowsheets('TPL24').MaterialStreams   # correntes da URL-III
    MT_UPGN                = simCase.Flowsheet.Flowsheets('TPL26').MaterialStreams   # correntes da UPGN
    ES_main                = simCase.Flowsheet.EnergyStreams                         # correntes de energia do flowsheet principal
    Operations             = simCase.Flowsheet.Operations                            # operações da flowsheet principal
    SS_UPGN                = simCase.Flowsheet.Flowsheets('TPL26').Operations.Item('Spread_UPGN')       # operações da flowsheet UPGN-II
    SS_Receita             = simCase.Flowsheet.Operations.Item('RECEITA')            # planilha de cálculo da receita
    SS_Rest                = simCase.Flowsheet.Operations.Item('RESTRIÇÕES')            # planilha de cálculo das restrições
    SS_URLI                = simCase.Flowsheet.Flowsheets.Item("TPL4").Flowsheets.Item("TPL22").Operations.Item("Spread_URLI") # Spreadsheet da URL-I
    SS_URLII               = simCase.Flowsheet.Flowsheets.Item("TPL4").Flowsheets.Item("TPL23").Operations.Item("Spread_URLII") # Spreadsheet da URL-I
    SS_URLIII              = simCase.Flowsheet.Flowsheets.Item("TPL4").Flowsheets.Item("TPL24").Operations.Item("Spread_URLIII") # Spreadsheet da URL-I
    SS_UPCGN               = simCase.Flowsheet.Flowsheets('TPL15').Operations.Item('prop_UPCGNs')       # planilha de cálculo de propriedades das UPCGNs
    SS_f_OBJ               = simCase.Flowsheet.Operations.Item('f_OBJ')  #planilha para o cálculo da função objetivo
    SS_STATUS_UNIDADES     = simCase.Flowsheet.Operations.Item('STATUS UNIDADES')  # versão_19 (planilha com STATUS DAS UNIDADES)
    SS_Custo               = simCase.Flowsheet.Operations.Item('CUSTO')            # planilha de cálculo do custo
    
    out_obj={'Solver':Solver,
             'MT_main':MT_main,
             'MT_URGN':MT_URGN,
             'MT_URGN':MT_URGN,
             'MT_URLI':MT_URLI,
             'MT_URLII':MMT_URLII,
             'MT_URLIII':MT_URLIII,
             'MT_UPGN':MT_UPGN,
             'ES_main':ES_main,
             'Operations':Operations,
             'SS_UPGN':SS_UPGN,
             'SS_Receita':SS_Receita,
             'SS_Rest':SS_Rest,
             'SS_URLI':SS_URLI,
             'SS_URLII':SS_URLII,
             'SS_URLIII':SS_URLIII,
             'SS_UPCGN':SS_UPCGN,
             'SS_f_OBJ':SS_f_OBJ,
             'SS_STATUS_UNIDADES':SS_STATUS_UNIDADES,
             'SS_Custo':SS_Custo,
        }
    
    
    return out_obj

def SLP(simCase, edata, obj, R_min, R_max, R_cap, Carga, FObj_type):
    
    '''
    *************************************************************************************************************************************
    [1] DESCRIÇÃO: SLP: Sequential Linear Programming -> Rotina da Programação Linear Sucessiva
    
    [2] EXPLICAÇÃO: Essa é utilizada para realizar a otimização linear sequencial do processo. Um modelo linear, baseado nas equações
    de balanço de massa global, é utilizado e implementado, nessa versão, na toolbox PULP (https://coin-or.github.io/pulp/index.html).
    Uma explicação detalhada do procedimento matemático está documentada do arquivo ProgramaçãoLinearSucessiva.ppx da pasta do projeto.
    A função SLP deve receber algumas variáveis e parâmetros, para possibilitar a comunicação e troca de "informações" entre o Python e
    o Hysys, de modo a possibilitar a otimização. 
    
    [3] DADOS DE ENTRADA: 
        simCase -> Objeto resultante da comunicação entre Python e Hysys (usado para abrir, fechar e ou iniciar a simulação);
        edata   -> Dicionário contendo os valores de parâmetros e variáveis do arquivo de entrada Input.xls;
        obj     -> Dicionário contendo os objetos resultantes das variáveis e spreadsheets do hsysys que serão utilizados
        R_min   -> Dicionário contendo os valores mínimos das restrições das especificações de produtos
        R_max   -> Dicionário contendo os valores máximos das restrições das especificações de produtos
        R_cap   -> Dicionário contendo os valores das restrições de capacidade das unidades
        Carga   -> Dicionário contendo os valores das vazões de carga da Unidade
        FObj_type -> Variável que indica a formulação da função objetivo: Obj_type = 'Custo'; 'Receita' ou 'Margem'
     
    [4] DADOS DE SAÌDA: 
        cod_SLP   -> Flag para indicar sucesso ou insucesso do cálculo
        rel_SLP   -> Dicionário contendo dados da otimização (iterações, derivadas, receitas, variáveis de decisão etc.)
        model     -> Objeto com todos os dados da otimização LP
        
    [5] OBSERVAÇÕES: Ao longo do código adicionaremos algumas anotações para facilitar a compreensão das considerações utilizadas
    
    [6] Modificações:
        
    07 de abril de 2024: Inclusão de flags para gravar restricções violadas
    *************************************************************************************************************************************
    '''
    
    
    '''
    SEÇÃO[1] DECLARAÇÃO DAS VARIÁVEIS
    *************************************************************************************************************************************
    '''
    'Iniciando contador de tempo'
    init_time = timeit.default_timer() # Estamos iniciando a contagem do tempo a partir daqui...
    
    'Descompactando o dicionário obj (somente os necessários para esta função)'
    MT_main = obj['MT_main'] # Objeto para se comunicar com as correntes do flowsheet principal
    F0 = edata['valor_inicial_manipuladas']  # Essa estimativa vem da planilha especificada pelo usuário!
    SS_f_OBJ =  obj['SS_f_OBJ'] 
    precos = edata['precos']
    SS_Receita=obj['SS_Receita']
         
    'Definição das variáveis da otimização LINEAR, de acordo com a notação do PULP'  
    G_295 = Carga['G_295']
    G_299 = Carga['G_299']
    G_302 = Carga['G_302']
    GASDUC_min = R_cap['GASDUC_min']
    GASDUC_max = R_cap['GASDUC_max']
    MIX_min = R_cap['MIX_min']
    MIX_max = R_cap['MIX_max']
    UPGN_min = R_cap['UPGN_min']
    UPGN_max = R_cap['UPGN_max']
    URGN_min = R_cap['URGN_min']
    URGN_max = R_cap['URGN_max']
    URLs_min = R_cap['URLs_min']
    URLs_max = R_cap['URLs_max']

    
    'Declaração das principais variáveis auxiliares e dimensões'
    itmax = 10 # número máximo de iterações [ESTAMOS USANDO ATUALEMNTE V13 ESTE CRITÉRIO DE PARADA]
    nD = len(F0) # Número de variáveis de decisão
    nC = 14 #Numero de restriçoes dos produtos
    R=np.zeros(itmax) # VETOR DE RECEITA PARA PLOTAGEM
    B=np.zeros(itmax) # VETOR DE RECEITA_BASE PARA PLOTAGEM
    D=np.zeros(itmax) # VETOR DE desvio PARA PLOTAGEM
    # C4p=np.zeros(itmax)
    manip = np.zeros([itmax, nD]) #Matriz das var manipuladas para receita
    x =[] # variáveis de decisão
    desvio = 1e-3 # desvio inicial para início do processo iterativo [SE FOR USAD]...
    index=0 # contado de iterações (inicia com valor zero)
    f_OBJ=np.zeros(nD) # Definindindo o vetor da função OBJETIVO
    dR_dF=np.zeros(nD) # Definindo o vetor das derivadas das Margens
    dC_dF=np.zeros([nD, nC]) #Definindo a matriz das derivadas das restriçoes dos produtos
    delta=np.zeros(nD) # Perturbação da derivada
    ZC_min = np.zeros(17) # valor mínimo de frações molares de reciclo da UPCGN para normalização do desvio
    ZC_max = np.ones(17)  # valor máximo de frações molares de reciclo da UPCGN para normalização do desvio 
    ZF_min = 0.6e6  # valor mínimo de vazão de reciclo da UPCGN para normalização do desvio [atribuído heuristicamente]
    ZF_max = 9e6    # valor máximo de vazão de reciclo da UPCGN para normalização do desvio [atribuído heuristicamente]
    ix = np.linspace(0, itmax, itmax) # vetor de iterações para plotagem
    flag_max=np.zeros(14) # flag para gravar restricoes de mínimo já violadas
    flag_min=np.zeros(14) # flag para gravar restricoes de máximo já violadas
    
    'Ativando a simulação Essencial'
    simCase.Solver.CanSolve = True # Isso é necessário, pois nesse momento vamos ler algumas variáveis como reciclo da UPCGN
    
    'Obtendo os valores das vazões e frações molares da corrente de reciclo [DE saída] da UPCGN -> necessário para o processo ITERATIVO'
    G_Rec_UPCGNs_in = MT_main['Gás de Reciclo UPCGNs RCY'].MolarFlow.GetValue('m3/d_(gas)') # valor atual lido da vazão em m3/d de gás do reciclo da UPCGn
    G_Rec_UPCGNs_C_in = MT_main['Gás de Reciclo UPCGNs RCY'].ComponentMolarFractionValue # # valor atual lido da fração molar dos componentes do gás do reciclo da UPCGn
    
    
    'Especificando o tipo de função-objetivo'
    
    mapping = {'Receita': (1.0, 0.0), 'Custo': (0.0, 1.0), 'Margem': (1.0, 1.0)} # (FLAGS) valores binários a serem exportados para a planilha f_OBJ
    Rec, Ct = mapping.get(FObj_type, (1.0, 0.0)) # Especificação dos valores de R:Receita e C:Custo. [OBS: Valor Default: R=1, C=0]
    SS_f_OBJ.Cell('D2').CellValue = Rec  # A célula D2 recebe a flag da receita
    SS_f_OBJ.Cell('D3').CellValue = Ct  # A célula D3 recebe a flag do custo
    
    # Enviando os Preços para a planilha receitas da simulação em Hysys
    #====================================================================================================================
    SS_Receita.Cell('D2').CellValue = precos['LGN [USD/ MM btu]']
    SS_Receita.Cell('D3').CellValue = precos['GV [USD/ MM btu]']
    SS_Receita.Cell('D4').CellValue = precos['GLP [USD/ MM btu]']
    SS_Receita.Cell('D5').CellValue = precos['C5p [USD/ MM btu]']
    SS_Receita.Cell('D6').CellValue = precos['GASDUC [USD/ MM btu]']
    #====================================================================================================================
    
    # X0 = np.array(G_Rec_UPCGNs_C_0) # composição da corrente de reciclo de saída da UPCGN em forma vetorial
    # XO4 = np.array(G_Rec_UPCGNs_C) # composição da corrente de reciclo de entrada da UPCGN em forma vetorial
    # desvioX = (X0-XO4)/(XO4+1e-10) # valor do desvio inicial das composições (é um vetor)
    # desvioX = la.norm(desvioX) # definindo a norma do desvio das composições
    # desvio = abs( (G_Rec_UPCGNs - G_Rec_UPCGNs_0) / G_Rec_UPCGNs_0 ) + desvioX + 1E-3 # (O 1E-3 É PARA FORÇAR RODAR A PRIMEIRA VEZ) [Desvio= Desvio Vazão + Desvio Composição]
    # C4_L=[] # definindo valor inicial do vetor de c4+
    
    'Valores iniciais das variáveis de decisão (direto da planilha Input_Data)'
    x0 = list(F0.values())  # VALORES INICIAIS DAS VARIÁVEIS MANIPULADAS CONVERTENDO VALORES PARA LISTA (FACILITA)

    # 'Obtenção da Receita_Base, Corrente de Reciclo UPCGN e Vazões_Base das Unidades'  
    # Receita_Base, Reciclo_UPCGN, Cargas_Unidades = SimulaLP(x0, G_Rec_UPCGNs_in,G_Rec_UPCGNs_C_in, obj) # Reciclo_UPCGN -> Dicionário com Vazões e Composições do reciclo UPCGN 
                                                                                                        # Cargas_Unidades -> Vazão de carga das unidades
    # 'Obtenção dos valores-base das condições dass correntes de produto (GV e GLP)'
    # y_base, ppm_H2O_GV = Restricoes(x0,0, obj)
    
    # 'Desativando a simulação Essencial'
    # simCase.Solver.CanSolve = False # Desativo (mas não desligo) a simulação pois será ativada mais tarde....   


    '''
    SEÇÃO[2] PROCESSO ITERATIVO
    *************************************************************************************************************************************
    '''
    'Início do Processo Iterativo'
    # while (abs(desvio) > tol): # PROCESSO ITERATIVO... [não estamos usando desvio nessa versão]
    while (abs(index) < itmax ): # PROCESSO ITERATIVO...[enquanto for menos que itemax iterações]

        'Obtenção da Função_Objetivo_Base, Corrente de Reciclo UPCGN e Vazões_Base das Unidades'  
        f_OBJ_Base, Reciclo_UPCGN, Cargas_Unidades, Receita, Custo = SimulaLP(x0, G_Rec_UPCGNs_in,G_Rec_UPCGNs_C_in, obj) # Reciclo_UPCGN -> Dicionário com Vazões e Composições do reciclo UPCGN 
                                                                                                      # Cargas_Unidades -> Vazão de carga das unidades
        'Obtenção dos valores-base das condições dass correntes de produto (GV e GLP)'
        y_base = Spec_prods(x0, 0, obj)
        
        index = index + 1 # Atualização do contador
        
        'Cálculo das Derivadas'
        
        if index<5: # Só atualiza a derivada nas primeiras iterações....[Default = 5]
            # delta=np.zeros(nD)
            for i in range(nD): # Para cada variável de decisão calcular a derivada da Receita em relação à variável de decisão atual
                delta[i]=1e4# Valor da perturbação [Especifciação heurística, em princípio]
                x = x0+delta # Incremento da perturbação na variável base (CÁLCULO VETORIAL)
                f_OBJ[i], Reciclo_UPCGN, Cargas_Unidades, Receita, Custo = SimulaLP(x, G_Rec_UPCGNs_in,G_Rec_UPCGNs_C_in, obj) # Cálculo da Receita para o novo ponto
                y = Spec_prods(x, 0, obj)
                dR_dF[i] = ( f_OBJ[i] - f_OBJ_Base ) / delta[i] # Cálculo da Derivada [usando o ponto_base]
                dC_dF[i,:] = (y - y_base) / delta[i] # Cálculo da Derivada das restriçoes
                delta[i]=0 # Zerando o incremento

        'Atribuindo os valores base das variáveis de decisão às variáveis do modelo LINEAR'
        
        G_295toGASDUC_0 = x0[0]  # A1
        G_295toURGN_0 =   x0[1]  # A2
        G_295toURLs_0 =   x0[2]  # A3
        G_295toUPGN_0 =   x0[3]  # A4
        
        G_299toGASDUC_0 = x0[4]  # B1
        G_299toURGN_0 =   x0[5]  # B2
        G_299toURLs_0 =   x0[6]  # B3
        G_299toUPGN_0 =   x0[7]  # B4
        
        G_302toURLs_0 =   x0[8]  # C3
        G_302toUPGN_0 =   x0[9]  # C4
        G_302toMIX_0 =    x0[10] # C5
        
        'Derivadas parciais da Função Objetivo no ponto base (Delta f_OBJ / Delta Vazão) [$/(m3/d)]'
        
        d_295toGASDUC = dR_dF[0]  # derivada da f_OBJ em relação ao A1
        d_295toURGN   = dR_dF[1]  # derivada da f_OBJ em relação ao A2
        d_295toURLs   = dR_dF[2]  # derivada da f_OBJ em relação ao A3
        d_295toUPGN   = dR_dF[3]  # derivada da f_OBJ em relação ao A4
        
        d_299toGASDUC = dR_dF[4]  # derivada da f_OBJ em relação ao B1
        d_299toURGN   = dR_dF[5]  # derivada da f_OBJ em relação ao B2
        d_299toURLs   = dR_dF[6]  # derivada da f_OBJ em relação ao B3
        d_299toUPGN   = dR_dF[7]  # derivada da f_OBJ em relação ao B4
        
        d_302toURLs   = dR_dF[8]  # derivada da f_OBJ em relação ao C3
        d_302toUPGN   = dR_dF[9]  # derivada da f_OBJ em relação ao C4
        d_302toMIX    = dR_dF[10] # derivada da f_OBJ em relação ao C5
        
        
        'Criação do objeto "model" para construção do modelo no pulp'
        model = LpProblem(name="Essencial_LP", sense=LpMaximize)
        
        'Definição das variaveis de decisão no PULP'
                   
        G_295toGASDUC = LpVariable("Gás de 295 para GASDUC", lowBound=1e-6, upBound=GASDUC_max) # A1
        G_295toURGN = LpVariable("Gás de 295 para URGN", lowBound=1e-6,upBound=URGN_max)        # A2
        G_295toURLs = LpVariable("Gás de 295 para URLs", lowBound=1e-6, upBound=URLs_max)       # A3
        G_295toUPGN = LpVariable("Gás de 295 para UPGN", lowBound=1e-6, upBound=UPGN_max)       # A4
        
        G_299toGASDUC = LpVariable("Gás de 299 para GASDUC", lowBound=1e-6, upBound=GASDUC_max) # B1
        G_299toURGN   = LpVariable("Gás de 299 para URGN", lowBound=1e-6, upBound=URGN_max)     # B2
        G_299toURLs   = LpVariable("Gás de 299 para URLs", lowBound=1e-6, upBound=URLs_max)     # B3
        G_299toUPGN   = LpVariable("Gás de 299 para UPGN", lowBound=0, upBound=UPGN_max)     # B4
        
        G_302toURLs = LpVariable("Gás de 302 para URLs", lowBound=1e-6, upBound=URLs_max)  # C3
        G_302toUPGN = LpVariable("Gás de 302 para UPGN", lowBound=1e-6, upBound=UPGN_max)  # C4
        G_302toMIX  = LpVariable("Gás de 302 para MIX", lowBound=1e-6, upBound=MIX_max) # C5
        
        
        'INCLUSÃO DE RESTRIÇÕES ATIVAS [só vamos incluir no LP as restrições ativas]'
       
        'avaliação das restrições mínimas'
        itens_min = list(R_min.values())
        keys = list(R_min.keys())        
        for k in range(np.size(y_base)):
            if y_base[k]<itens_min[k] or flag_min[k]!=0:
                if flag_min[k] == 0:
                    print('************************************')
                    print('Violação do valor mínimo: ', keys[k], '=', itens_min[k], '>', y_base[k]  )
                    print('Incluindo restrição no LP')
                    print('************************************')
                else:
                    print('************************************')
                    print('Mantendo restrição no LP já violada previamente')
                    print('Violação do valor mínimo anterior: ', keys[k], '<', itens_min[k]  )
                    print('************************************')
                val_rest = (y_base[k] + 
                            dC_dF[0,k]*(G_295toGASDUC-G_295toGASDUC_0) + dC_dF[1,k]*(G_295toURGN-G_295toURGN_0) + dC_dF[2,k]*(G_295toURLs-G_295toURLs_0) + dC_dF[3,k]*(G_295toUPGN-G_295toUPGN_0) + 
                            dC_dF[4,k]*(G_299toGASDUC-G_299toGASDUC_0) + dC_dF[5,k]*(G_295toURGN-G_295toURGN_0) + dC_dF[6,k]*(G_299toURLs-G_299toURLs_0) + dC_dF[7,k]*(G_299toUPGN-G_299toUPGN_0) +
                            dC_dF[8,k]*(G_302toURLs-G_302toURLs_0) + dC_dF[9,k]*(G_302toUPGN-G_302toUPGN_0) + dC_dF[10,k]*(G_302toMIX-G_302toMIX_0)) 
                val_rest = (y_base[k] + dC_dF[10,k]*(G_302toMIX-G_302toMIX_0))
                model += ((val_rest) - float(itens_min[k]) >= 0, keys[k]) # Restrição Incluída
                flag_min[k]=1
        'avaliação das restrições máximas'            
        itens_max = list(R_max.values())
        keys = list(R_max.keys())        
        for k in range(np.size(y_base)):
            if y_base[k]>itens_max[k] or flag_max[k]!=0:
                if flag_max[k] == 0:
                    print('************************************')
                    print('Violação do valor máximo: ', keys[k],'=', itens_max[k], '<', y_base[k])
                    print('Incluindo restrição no LP')
                    print('************************************')
                else:
                    print('************************************')
                    print('Mantendo restrição no LP já violada previamente')
                    print('Violação do valor máximo anterior: ', keys[k], '>', itens_max[k]  )
                    print('************************************')
                val_rest = (y_base[k] + 
                            dC_dF[0,k]*(G_295toGASDUC-G_295toGASDUC_0) + dC_dF[1,k]*(G_295toURGN-G_295toURGN_0) + dC_dF[2,k]*(G_295toURLs-G_295toURLs_0) + dC_dF[3,k]*(G_295toUPGN-G_295toUPGN_0) + 
                            dC_dF[4,k]*(G_299toGASDUC-G_299toGASDUC_0) + dC_dF[5,k]*(G_295toURGN-G_295toURGN_0) + dC_dF[6,k]*(G_299toURLs-G_299toURLs_0) + dC_dF[7,k]*(G_299toUPGN-G_299toUPGN_0) +
                            dC_dF[8,k]*(G_302toURLs-G_302toURLs_0) + dC_dF[9,k]*(G_302toUPGN-G_302toUPGN_0) + dC_dF[10,k]*(G_302toMIX-G_302toMIX_0)) 
                val_rest = (y_base[k] + dC_dF[10,k]*(G_302toMIX-G_302toMIX_0))
                model += ((val_rest - float(itens_max[k])) <= 0, keys[k]) # Restrição Incluída
                flag_max[k] = 1
        'Função Objetivo (Linealizada no ponto base)'
        model += (f_OBJ_Base + 
                  d_299toGASDUC*(G_299toGASDUC-G_299toGASDUC_0) + d_295toGASDUC*(G_295toGASDUC-G_295toGASDUC_0) +
                  d_299toURGN*(G_299toURGN-G_299toURGN_0)       + d_295toURGN*(G_295toURGN-G_295toURGN_0)       + 
                  d_302toURLs*(G_302toURLs-G_302toURLs_0)       + 
                  d_299toURLs*(G_299toURLs-G_299toURLs_0)       + d_295toURLs*(G_295toURLs-G_295toURLs_0)       +
                  d_302toUPGN*(G_302toUPGN-G_302toUPGN_0)       + d_299toUPGN*(G_299toUPGN-G_299toUPGN_0)       + 
                  d_295toUPGN*(G_295toUPGN-G_295toUPGN_0)       + d_302toMIX*(G_302toMIX-G_302toMIX_0)   )
        
        'Restrições do Modelo'
          
        'IGUALDADE: EQUAÇÕES DE CONSERVAÇÃO DE MASSA'
        model += (G_302toURLs + G_302toUPGN + G_302toMIX == G_302, "Alocar todo o gás do 302")
        model += (G_299toGASDUC +  G_299toURGN + G_299toURLs + G_299toUPGN == G_299, "Alocar todo o gás do 299")
        model += (G_295toGASDUC + G_295toURGN +  G_295toURLs + G_295toUPGN == G_295 + G_Rec_UPCGNs_in, "Alocar todo o gás do 295 e do Reciclo")
        'DEIGUALDADE: CAPACIDADE DAS UNIDADES'
        model += (GASDUC_min - (G_299toGASDUC + G_295toGASDUC)              <= 0, "GASDUC mínimo")
        model += (             (G_299toGASDUC + G_295toGASDUC) - GASDUC_max <= 0, "GASDUC máximo")
        
        model += (             MIX_min - G_302toMIX <= 0, "MIX mínimo")
        model += (G_302toMIX - MIX_max              <= 0, "MIX máximo")
        
        model += (UPGN_min - (G_302toUPGN + G_299toUPGN + G_295toUPGN)             <=0 , "UPGN mínimo")
        model += (           (G_302toUPGN + G_299toUPGN + G_295toUPGN) - UPGN_max <= 0, "UPGN máximo")
        
        model += (URGN_min - (G_299toURGN + G_295toURGN)            <= 0, "URGN mínimo")
        model += (           (G_299toURGN + G_295toURGN) - URGN_max <= 0, "URGN máximo")
        
        model += (URLs_min - (G_302toURLs + G_299toURLs + G_295toURLs)            <= 0, "URLs mínimo")
        model += (           (G_302toURLs + G_299toURLs + G_295toURLs) - URLs_max <= 0, "URLs máximo")
        # model += (           (C4_L) - C4_max <= 0, "C4 máximo") # Restrição de C4+ (está em fração) 
        
        'Executa otimização'
        status = model.solve(pulp.PULP_CBC_CMD(msg=False))
        print(status)

        'Redefinindo as variáveis de decisão com valores ótimos, para a próxima iteração'
        'NOTA: Vamos utilizar os valores ótimos das variáveis de decisão e realizar a simulção com esses valores para verificar a função_Objetivo'
        
        x[0]=G_295toGASDUC.varValue    #A1           # Redefinindo as variáveis de decisão com valores ótimos
        x[1]=G_295toURGN.varValue      #A2
        x[2]=G_295toURLs.varValue      #A3
        x[3]=G_295toUPGN.varValue      #A4
        x[4]=G_299toGASDUC.varValue    #B1
        x[5]=G_299toURGN.varValue      #B2
        x[6]=G_299toURLs.varValue      #B3
        x[7]=G_299toUPGN.varValue      #B4
        x[8]=G_302toURLs.varValue      #C3 
        x[9]=G_302toUPGN.varValue      #C4
        x[10]=G_302toMIX.varValue      #C5
        
        'Tratamendo de dados.. NOTA: Evitar vazões negativas'
        for i in range(nD):
            if x[i]<0:  # se alguma vazão for negativa, atribuir valor zero
                x[i]=1e-5
                cod_SLP = 1  # flag da função SLP para indicar que ocorreram vazões negativas
            else:
                cod_SLP = 0  # flag da função SLP para indicar que está ok
        
        'Obtenção da Receita_Simulada, Corrente de Reciclo UPCGN e Vazões_Base das Unidades' 
        f_OBJ_max, Reciclo_UPCGN, Cargas_Unidades, Receita, Custo = SimulaLP(x, G_Rec_UPCGNs_in,G_Rec_UPCGNs_C_in, obj) # Cálculo da Função-Onjetivo para o novo ponto
        x0  = x
        'Cálculo do Desvio' 
        G_Rec_UPCGNs_C_out = Reciclo_UPCGN['G_Rec_UPCGNs_C_out'] # Composição da corrente de reciclo da UPCNGN na saída
        G_Rec_UPCGNs_out   = Reciclo_UPCGN['G_Rec_UPCGN_out']      # Vazão de reciclo da UPCNGN na saída
        
        ZC_in  = ( np.array(G_Rec_UPCGNs_C_in) - ZC_min ) / (ZC_max - ZC_min)
        ZF_in  = ( np.array(G_Rec_UPCGNs_in) - ZF_min ) / (ZF_max - ZF_min)
        ZF_out = ( np.array(G_Rec_UPCGNs_out) - ZF_min ) /  (ZF_max - ZF_min)  # vazão da corrente de reiclo da UPCGN 
        ZC_out = ( np.array(G_Rec_UPCGNs_C_out) -  ZC_min ) / (ZC_max - ZC_min)  # composição da corrente de reiclo da UPCGN
        desvio_comp = (ZC_in-ZC_out)/(ZC_out+1e-10) # o valor 1e-10 é para evitar divisão por zero
        desvio_comp = la.norm(desvio_comp)
        desvio_vazão = abs( (ZF_in - ZF_out) / ZF_out )
        desvio = desvio_vazão + desvio_comp # cálculo do desvio
        G_Rec_UPCGNs_in = G_Rec_UPCGNs_out # atualização da vazão de reciclo
        G_Rec_UPCGNs_C_in = G_Rec_UPCGNs_C_out # atualização da composição do reciclo
        
        ###################################################################################################################
        'Impressão de Resultados no Terminal'
        
        CARGA_G_URGN = Cargas_Unidades['CARGA_G_URGN']
        CARGA_G_URLI = Cargas_Unidades['CARGA_G_URLI']
        CARGA_G_URLII = Cargas_Unidades['CARGA_G_URLII']
        CARGA_G_URLIII = Cargas_Unidades['CARGA_G_URLIII']
        CARGA_G_UPGN = Cargas_Unidades['CARGA_G_UPGN']
        
        R[index-1]=f_OBJ_max # Funçao Objetivo calculada pelo Hysys
        B[index-1]=f_OBJ_Base # Funçao Objetivo base
        D[index-1]=desvio # desvio
        manip[index-1,:] = x # variáveis manipuladas
        
        'Imprimindo os resultados no TERMINAL'
        print('#'*50)
        print(f"Resultados iteração {index-1}:")
        print("Convergência: ", model.status)
        # The status of the solution is printed to the screen
        print("Status:", LpStatus[model.status])
        print(FObj_type, model.objective.value())
        print('Gás de reciclo: ', G_Rec_UPCGNs_in)
        print('Composicao do reciclo: ', 'C1=', G_Rec_UPCGNs_C_in[0], 'C2=', G_Rec_UPCGNs_C_in[1], 'C3=', G_Rec_UPCGNs_C_in[2]
              , 'iC4=', G_Rec_UPCGNs_C_in[3], 'nC4=', G_Rec_UPCGNs_C_in[4], 'iC5=', G_Rec_UPCGNs_C_in[5], 'C6=', G_Rec_UPCGNs_C_in[6]
              , 'C7=', G_Rec_UPCGNs_C_in[7], 'C8=', G_Rec_UPCGNs_C_in[8] , 'C9=', G_Rec_UPCGNs_C_in[9], 'C10=', G_Rec_UPCGNs_C_in[10]
              , 'N2=', G_Rec_UPCGNs_C_in[11], 'Co2=', G_Rec_UPCGNs_C_in[12], 'H2O=', G_Rec_UPCGNs_C_in[13]
              , 'H2S=', G_Rec_UPCGNs_C_in[14], 'EGLYC=', G_Rec_UPCGNs_C_in[15])
        print('*'*50)
        print('Manipuladas:')
        for var in model.variables():
            print(f"{var.name}: {var.value()}")
        print('*'*50)
        print('Restrições:')
        for name, constraint in model.constraints.items():
            print(f"{name}: {constraint.value()}")
        print('*'*50)
        print('Cargas totais para as unidades:')
        print('GASDUC: ', G_299toGASDUC.value() + G_295toGASDUC.value())
        print('MIX: ', G_302toMIX.value())
        print('UPGN: ', G_302toUPGN.value() + G_299toUPGN.value() + G_295toUPGN.value())
        print('URGN: ', G_299toURGN.value() + G_295toURGN.value())
        print('URLs: ', G_302toURLs.value() + G_299toURLs.value() + G_295toURLs.value())
        print('*'*50)
        
        # Gerando o arquivo .lp da otimização
        model.writeLP("Cabiunas_LP.lp")
        
        print( FObj_type, "_Base: ", f_OBJ_Base)
        print( FObj_type, "_max: ", f_OBJ_max)
        # Receita
        print( "Receita", Receita)
        print( "Custo", Custo)
        print( "Margem", Receita-Custo)
        print('Desvio:', D)
        print('*'*50)
        print('CARGA_GAS_URGN:', CARGA_G_URGN, '||', 'URGN:', G_299toURGN.value() + G_295toURGN.value(), 'CONDENSADO:', ((G_299toURGN.value() + G_295toURGN.value())-CARGA_G_URGN ) )
        URLI = (G_302toURLs.value() + G_299toURLs.value() + G_295toURLs.value() )/3
        print('CARGA_GAS_URLI:', CARGA_G_URLI, '||', 'URLI:', URLI, 'CONDENSADO:', (CARGA_G_URLI - URLI))
        print('CARGA_GAS_URLII:', CARGA_G_URLII, '||', 'URLII:', URLI, 'CONDENSADO:', (CARGA_G_URLII - URLI))
        print('CARGA_GAS_URLIII:', CARGA_G_URLIII, '||', 'URLIII:', URLI, 'CONDENSADO:', (CARGA_G_URLIII - URLI))
        UPGN = G_302toUPGN.value() + G_299toUPGN.value() + G_295toUPGN.value()
        print('CARGA_GAS_UPGN:', CARGA_G_UPGN, '||', 'UPGN:', UPGN, 'CONDENSADO:', (UPGN - CARGA_G_UPGN))
        final_time = timeit.default_timer()
        tempo = final_time - init_time 
        print('Tempo de execução (seg) ', tempo)
        
        'Plotando as derivadas'    
        plot_derivatives(dR_dF, index)
        
        'Relatório com histórico de cálculos da SLP' # Ainda Implementando...
        rel_SLP = {'FOBJ_base': R,
                   'FOBJ': B,
                   'Desvio': D,
                   'Iterações': ix,
                   'Manipuladas':x,
                   }
    
    return cod_SLP, model, rel_SLP

def SpecLP(edata, obj):
    
    '''
    *************************************************************************************************************************************
    [1] DESCRIÇÃO: SpecLP: Estepcificação das Variáveis LP do PULP 
    
    [2] EXPLICAÇÃO: Essa rotina é utilizada para especificar as variáveis e especificações que serão utilizadas para montar o MODELO
    LP no PULP. Será necessário especificar as restrições de capacidade, restrições dos produtos, e vazões de carga.  
    
    
    [3] DADOS DE ENTRADA: 
        edata   -> Dicionário resultante da leitura da dedos da planilha Input_Data.xlsx;
        obj     -> Dicionário contendo os objetos resultantes das variáveis e spreadsheets do hsysys que serão utilizados
    
    [4] DADOS DE SAÌDA: 
        cod_speclp   -> Flag para indicar sucesso ou insucesso do cálculo
        R_min        -> Dicionário contendo dados de restrições mínimas de produtos
        R_max        -> Dicionário contendo dados de restrições máximas de produtos
        R_cap        -> Dicionário contendo dados de restrições de capacidade das UNIDADES 
        Carga        -> Dicionário contendo dados de  vazões de carga
        
    [5] OBSERVAÇÕES: Ao longo do código adicionaremos algumas anotações para facilitar a compreensão das considerações utilizadas
    *************************************************************************************************************************************
    '''
    
    
    
    restricoes_capacidade_LC = edata['restricoes_capacidade_LC']
    restricoes_capacidade_UC = edata['restricoes_capacidade_UC']
    restricoes_produtos_LC = edata['restricoes_produtos_LC']
    restricoes_produtos_UC = edata['restricoes_produtos_UC']
    
    MT_main = obj['MT_main']
    
    # Restriçoes de capacidade das plantas
    #====================================================================================================================
    GASDUC_min = restricoes_capacidade_LC['GASDUC']
    GASDUC_max = restricoes_capacidade_UC['GASDUC']

    MIX_min = restricoes_capacidade_LC['MIX']
    MIX_max = restricoes_capacidade_UC['MIX']

    UPGN_min = restricoes_capacidade_LC['UPGN']
    UPGN_max = restricoes_capacidade_UC['UPGN']

    URGN_min = restricoes_capacidade_LC['URGN']
    URGN_max = restricoes_capacidade_UC['URGN']

    URLs_min = restricoes_capacidade_LC['URLs']
    URLs_max = restricoes_capacidade_UC['URLs']
    #====================================================================================================================

    # Restrições de produtos
    #====================================================================================================================
    
    # Restrições Mínimas
    R_min = {
            "GV_C1": restricoes_produtos_LC['GV_C1 [% mol]'],
            "GV_C2": restricoes_produtos_LC['GV_C2 [% mol]'],
            "GV_C3": restricoes_produtos_LC['GV_C3 [% mol]'],
            "GV_C4p": restricoes_produtos_LC['GV_C4p [% mol]'],
            "GV_CO2": restricoes_produtos_LC['GV_CO2 [% mol]'],
            "GV_inertes": restricoes_produtos_LC['GV_inertes [% mol]'],
            "GV_PCS": restricoes_produtos_LC['GV_PCS [MJ/m3]'],
            "GV_POA": restricoes_produtos_LC['GV_POA [oC]'],
            "GV_POH": restricoes_produtos_LC['GV_POH [oC]'],
            "GV_Wobbe": restricoes_produtos_LC['GV_Wobbe [MJ/m3]'],
            "GV_No_metano": restricoes_produtos_LC['GV_No_metano [-]'],
            "GLP_C2": restricoes_produtos_LC['GLP_C2 [% vol]'],
            "GLP_C5p": restricoes_produtos_LC['GLP_C5p [% vol]'],
            "GLP_rho": restricoes_produtos_LC['GLP_rho [kg/m3]'],
            
            }

    # Restrições Máximas
    R_max = {
            "GV_C1": restricoes_produtos_UC['GV_C1 [% mol]'],
            "GV_C2": restricoes_produtos_UC['GV_C2 [% mol]'],
            "GV_C3": restricoes_produtos_UC['GV_C3 [% mol]'],
            "GV_C4p": restricoes_produtos_UC['GV_C4p [% mol]'],
            "GV_CO2": restricoes_produtos_UC['GV_CO2 [% mol]'],
            "GV_inertes": restricoes_produtos_UC['GV_inertes [% mol]'],
            "GV_PCS": restricoes_produtos_UC['GV_PCS [MJ/m3]'],
            "GV_POA": restricoes_produtos_UC['GV_POA [oC]'],
            "GV_POH": restricoes_produtos_UC['GV_POH [oC]'],
            "GV_Wobbe": restricoes_produtos_UC['GV_Wobbe [MJ/m3]'],
            "GV_No_metano": restricoes_produtos_UC['GV_No_metano [-]'],
            "GLP_C2": restricoes_produtos_UC['GLP_C2 [% vol]'],
            "GLP_C5p": restricoes_produtos_UC['GLP_C5p [% vol]'],
            "GLP_rho": restricoes_produtos_UC['GLP_rho [kg/m3]'],
            }
    
    R_cap = {
        "GASDUC_min": GASDUC_min,
        "GASDUC_max": GASDUC_max,
        "MIX_min": MIX_min,
        "MIX_max": MIX_max,
        "UPGN_min": UPGN_min,
        "UPGN_max": UPGN_max,
        "URGN_min": URGN_min,
        "URGN_max": URGN_max,
        "URLs_min": URLs_min,
        "URLs_max": URLs_max,
        }
    
    G_295 = MT_main['295Gout2_BP'].MolarFlow.GetValue('m3/d_(gas)')
    G_299 = MT_main['299Gout2'].MolarFlow.GetValue('m3/d_(gas)')
    G_302 = MT_main['302Gout2'].MolarFlow.GetValue('m3/d_(gas)')

    G_Rec_UPCGNs = MT_main['Gás de Reciclo UPCGNs RCY'].MolarFlow.GetValue('m3/d_(gas)')
    G_Rec_UPCGNs_C  = MT_main['Gás de Reciclo UPCGNs RCY'].ComponentMolarFractionValue # COMPOSICOES DO RECICLO
    
    Carga = {
        "G_295": G_295,
        "G_299": G_299,
        "G_302": G_302
        }
    
    cod_speclp=1
    
    return cod_speclp, R_min, R_max, R_cap, Carga

def SpecVar(edata, obj, R_especs):
    
    '''
    *************************************************************************************************************************************
    [1] DESCRIÇÃO: SpecVar: Especificação das Variáveis -> Rotina utilizada para especificar algumas variáveis das UNIDADES 
    
    [2] EXPLICAÇÃO: Essa rotina é utilizada para especificar algumas variáveis das UNIDADES URLS, URGN, UPGN e UPCGN. Além
    disso importaremos os valores dos preços e condições dos coletores. A maior parte
    das variáveis é especificada no arquivo Input_Data.xlsx. Outras variáveis (frações molares da T02 das URLs) são calculadas na 
    simulação rogorosa (utilizada no modo offline) e importadas para a simulação Essencial.
    
    [3] DADOS DE ENTRADA: 
        edata   -> Dicionário resultante da leitura da dedos da planilha Input_Data.xlsx;
        obj     -> Dicionário contendo os objetos resultantes das variáveis e spreadsheets do hsysys que serão utilizados
        R_especs-> Dicionário contendo as especificações importadas da simulação rigorosa
    
    [4] DADOS DE SAÌDA: 
        cod_SpecVar  -> Flag para indicar sucesso ou insucesso do cálculo
        
    [5] OBSERVAÇÕES: Ao longo do código adicionaremos algumas anotações para facilitar a compreensão das considerações utilizadas
    
    [6] Última Modificação: Importação dos valores das especificações de fração molar de CO2 nas URLs
    *************************************************************************************************************************************
    '''
    
    'Descompactando o dicionário obj'
    
    MT_main = obj['MT_main']
    MT_URGN = obj['MT_URGN']
    # MT_URLI = obj['MT_URLI'] 
    # MT_URLII = obj['MT_URLII'] 
    # MT_URLIII = obj['MT_URLIII'] 
    MT_UPGN  = obj['MT_UPGN']
    # ES_main = obj['ES_main']
    SS_UPGN = obj['SS_UPGN'] # Spreadsheed da UPGN-II
    SS_URLI = obj['SS_URLI']
    SS_URLII = obj['SS_URLII'] 
    SS_URLIII = obj['SS_URLIII'] 
    SS_UPCGN = obj['SS_UPCGN']
    SS_Receita = obj['SS_Receita']
    SS_STATUS_UNIDADES = obj['SS_STATUS_UNIDADES']  # Spreadsheet com status das unidades...
    SS_Custo = obj['SS_Custo'] # Spreadsheet com os custos [Versão 18]
    
    'Descompactando o dicionário edata'
    
    col_295_condicoes = edata['col_295_condicoes']
    col_299_condicoes = edata['col_299_condicoes']
    col_302_condicoes = edata['col_302_condicoes'] 
    col_295_comp = edata['col_295_comp'] 
    col_299_comp = edata['col_299_comp'] 
    col_302_comp  = edata['col_302_comp']
    espec = edata['espec_plantas']
    valor_inicial_manipuladas = edata['valor_inicial_manipuladas']  
    precos = edata['precos']    
    custos = edata['custos'] # versão 19
    status_unit = edata['status_unit'] # status das unidades
    
    'Descompactando as especificações que vêm da Simulação Rigorosa'
    
    C2_URLI = R_especs['C2_URLI']
    C3_URLI = R_especs['C3_URLI']
    C1_URLI = R_especs['C1_URLI']
    C2_URLII = R_especs['C2_URLII']
    C3_URLII = R_especs['C3_URLII']
    C1_URLII = R_especs['C1_URLII']
    C2_URLIII = R_especs['C2_URLIII']
    C3_URLIII = R_especs['C3_URLIII']
    C1_URLIII = R_especs['C1_URLIII']
    # Temp_V03 = R_especs['Temp_V03']
    T_P24_UPGN = R_especs['T_P24_UPGN']
    
    # Valores de CO2 das URLS
    
    CO2_URLI   = R_especs['CO2_URLI']
    CO2_URLII  = R_especs['CO2_URLII']
    CO2_URLIII = R_especs['CO2_URLIII']    
    
    
  
    'Enviando as condiçoes dos coletores para a simulação em HYSYS'
    
    MT_main['GásRico295'].MolarFlow.SetValue(col_295_condicoes['Vazão_entrada [m3/d_(gas)]'],'m3/d_(gas)') # Carga do coletor 295
    MT_main['GásRico295'].Temperature.SetValue(col_295_condicoes['Temperatura [oC]'],'C') # Temperatura do coletor 295
    MT_main['GásRico295'].Pressure.SetValue(col_295_condicoes['Pressao [kg/cm2_g]'],'kg/cm2_g') # Pressão do coletor 295
    MT_main['GásRico295'].ComponentMolarFractionValue = list(col_295_comp.values()) # Fração molar coletor 295
    

    MT_main['GásRico299'].MolarFlow.SetValue(col_299_condicoes['Vazão_entrada [m3/d_(gas)]'],'m3/d_(gas)') # Carga do coletor 299
    MT_main['GásRico299'].Temperature.SetValue(col_299_condicoes['Temperatura [oC]'],'C') # Temperatura do coletor 299
    MT_main['GásRico299'].Pressure.SetValue(col_299_condicoes['Pressao [kg/cm2_g]'],'kg/cm2_g') # Pressão do coletor 299
    MT_main['GásRico299'].ComponentMolarFractionValue = list(col_299_comp.values()) # Fração molar coletor 299

    MT_main['GásRico302'].MolarFlow.SetValue(col_302_condicoes['Vazão_entrada [m3/d_(gas)]'],'m3/d_(gas)') # Carga do coletor 302
    MT_main['GásRico302'].Temperature.SetValue(col_302_condicoes['Temperatura [oC]'],'C') # Temperatura do coletor 302
    MT_main['GásRico302'].Pressure.SetValue(col_302_condicoes['Pressao [kg/cm2_g]'],'kg/cm2_g') # Pressão do coletor 302
    MT_main['GásRico302'].ComponentMolarFractionValue = list(col_302_comp.values()) # Fração molar coletor 302
    #====================================================================================================================

    'Enviando os valores iniciais das manipuladas para a simulação em Hysys'
    
    MT_main['A1'].MolarFlow.SetValue(valor_inicial_manipuladas['F_A1'],'m3/d_(gas)')   #M1
    MT_main['A2'].MolarFlow.SetValue(valor_inicial_manipuladas['F_A2'],'m3/d_(gas)')   #M2
    MT_main['A3'].MolarFlow.SetValue(valor_inicial_manipuladas['F_A3'],'m3/d_(gas)')   #M3
    MT_main['A4'].MolarFlow.SetValue(valor_inicial_manipuladas['F_A4'],'m3/d_(gas)')   #M4

    MT_main['B1'].MolarFlow.SetValue(valor_inicial_manipuladas['F_B1'],'m3/d_(gas)')   #M5
    MT_main['B2'].MolarFlow.SetValue(valor_inicial_manipuladas['F_B2'],'m3/d_(gas)')   #M6
    MT_main['B3'].MolarFlow.SetValue(valor_inicial_manipuladas['F_B3'],'m3/d_(gas)')   #M7
    MT_main['B4'].MolarFlow.SetValue(valor_inicial_manipuladas['F_B4'],'m3/d_(gas)')   #M8 

    MT_main['C3'].MolarFlow.SetValue(valor_inicial_manipuladas['F_C3'],'m3/d_(gas)')   #M9 
    MT_main['C4'].MolarFlow.SetValue(valor_inicial_manipuladas['F_C4'],'m3/d_(gas)')   #M10 
    MT_main['C5'].MolarFlow.SetValue(valor_inicial_manipuladas['F_C5'],'m3/d_(gas)')   #M11 
    #========================================================================================

    'Enviando os Preços para a planilha receitas da simulação em Hysys'
    
    SS_Receita.Cell('D2').CellValue = precos['LGN [USD/ MM btu]']*precos['Cambio [R$/USD]']
    SS_Receita.Cell('D3').CellValue = precos['GV [USD/ MM btu]']*precos['Cambio [R$/USD]']
    SS_Receita.Cell('D4').CellValue = precos['GLP [USD/ MM btu]']*precos['Cambio [R$/USD]']
    SS_Receita.Cell('D5').CellValue = precos['C5p [USD/ MM btu]']*precos['Cambio [R$/USD]']
    SS_Receita.Cell('D6').CellValue = precos['GASDUC [USD/ MM btu]']*precos['Cambio [R$/USD]']
    SS_Receita.Cell('B8').CellValue = precos['Cambio [R$/USD]']
    #====================================================================================================================
    
    'Enviando os Custos para a planilha receitas da simulação em Hysys'
    
    SS_Custo.Cell('G31').CellValue = custos['custoEnergiaEletrica']
    SS_Custo.Cell('H31').CellValue = custos['custoGasCombustivel']
    # PRODUTOS QUÍMICOS
    SS_Custo.Cell('F31').CellValue = custos['propanoURGN']
    SS_Custo.Cell('F32').CellValue = custos['aminaAtivadaURCO2']
    SS_Custo.Cell('F33').CellValue = custos['propanoUPGNII']
    SS_Custo.Cell('F34').CellValue = custos['peneiraMolecularURLs'] 
    SS_Custo.Cell('F35').CellValue = custos['carvaoAtivadoURCO2']    
    
    
    

    'Descompactando as especificações das Unidades Armazenadas no dicionário espec'
    '******************************************************************************************************************************'

    e1 = espec['URGN_vaso1_pressao [kg/cm2_g]'] # Pressão do vaso V20501 da URGN
    e2 = espec['URGN_vaso2_temperatura [oC]']   # Temperatura do vaso V20502 da URGN
    e3 = espec['UPGNII_P04_temperatura [oC]'] # Temperatura da corrente P08
    e4 = espec['UPGNII_T01_Temp_topo [oC]'] # Temperatura do topo da T01 da UPGN-II
    e5 = C1_URLI # FRAÇÃO MOLAR de C1 no fundo da T01 DA URL-1 (CALCULADA PELA SIMULAÇÃO RIGOROSA)
    e6 = C2_URLI # FRAÇÃO MOLAR DE C2 NO TOPO DA T01 DA URL-1 (CALCULADA PELA SIMULAÇÃO RIGOROSA)  
    e7 = C1_URLII # FRAÇÃO MOLAR de C1 no fundo da T01 NO TOPO DA T01 DA URL-2 (CALCULADA PELA SIMULAÇÃO RIGOROSA)
    e8 = C2_URLII # FRAÇÃO MOLAR DE C2 NO TOPO DA T01 DA URL-2 (CALCULADA PELA SIMULAÇÃO RIGOROSA)       
    e9 =C1_URLIII # FRAÇÃO MOLAR de C1 no fundo da T01 NO TOPO DA T01 DA URL-3 (CALCULADA PELA SIMULAÇÃO RIGOROSA)
    e10 = C2_URLIII # FRAÇÃO MOLAR DE C2 NO TOPO DA T01 DA URL-3 (CALCULADA PELA SIMULAÇÃO RIGOROSA)
    e11 = C3_URLI # FRAÇÃO MOLAR DE Propano no topo da T01 da URL-1 (CALCULADA PELA SIMULAÇÃO RIGOROSA)
    e12 = C3_URLII # FRAÇÃO MOLAR DE Propano no topo da T01 da URL-2 (CALCULADA PELA SIMULAÇÃO RIGOROSA)
    e13 = C3_URLIII # FRAÇÃO MOLAR DE Propano no topo da T01 da URL-3 (CALCULADA PELA SIMULAÇÃO RIGOROSA) 
    
    
    e14  = 1e-10 #espec['Sangria_URLIeII_vazao [m3/d]'] #Em princípio mantemos zeradas
    e15  = 1e-10 #espec['Sangria_URLIII_vazao [m3/d]'] #Em princípio mantemos zeradas
    
    'NOTA: NA versão atual da Otimização Linear Sequencial mantemos as vazões de Sangria zeradas (temos ainda que entender o motiuvo de se utilizar'
    'essas vazões na otimização)'
    
    e16 = espec['UPCGNII_GLP_C2 [%]']
    e17 = espec['UPCGNII_GLP_C2 [%]']
    e18 = espec['UPCGNII_GLP_C2 [%]']
    e19 = espec['UPCGNII_GLP_C2 [%]']
    
    e20 = espec['UPCGNII_vaso1_pressao [kg/cm2_g]']
    e21 = espec['UPCGNIII_vaso1_pressao [kg/cm2_g]']
    e22 = espec['UPCGNIV_vaso1_pressao [kg/cm2_g]']
    
    e23 = espec['UPCGNII_T02_recupercao_C4_topo [-]']
    e24 = espec['UPCGNIII_T02_recupercao_C4_topo [-]']
    e25 = espec['UPCGNIV_T02_recupercao_C4_topo [-]']
    
    e26 = espec['UPCGNII_T02_recupercao_C5p_fundo [-]']
    e27 = espec['UPCGNIII_T02_recupercao_C5p_fundo [-]']
    e28 = espec['UPCGNIV_T02_recupercao_C5p_fundo [-]']
    
    e29 = CO2_URLI
    e30 = CO2_URLII
    e31 = CO2_URLIII
    
    '******************************************************************************************************************************'


    'Exportando as ESPECIFICAÇÕES DA URGN'
    MT_URGN['1'].Pressure.SetValue(e1,'kg/cm2_g')  #ESPEC1 
    MT_URGN['7'].Temperature.SetValue(e2,'C')         #ESPEC2
  
    'Exportando as ESPECIFICAÇÕES DA UPGN'
  
    MT_UPGN['P08'].Temperature.SetValue(e3,'C')    # Temperatura da corrente P08 da UPGN-II
    SS_UPGN.Cell('C2').CellValue = e4 # Temperatura do Topo da T01 da UPGN-II
    SS_UPCGN.Cell('B2').CellValue = 0.9933 *e16/100  #fração molar C2 no GLP da UPGN-II especificado diretamente aqui...
 
    'Exportando as ESPECIFICAÇÕES paras as URLS'
 
    # URLI_T01.values = ((e5,), (e6,), (e11,), (0.0,), (0.0,), (0.0,), (0.0,), (0.0,), (0.0,), (0.0,), (0.0,), (0.0,), (1.0,), (0.00161,), (0.0,), (0.0,), (0.0,)) # Especificação da T01 da URL-I
    # URLII_T01.values = ((e7,), (e8,), (e12,), (0.0,), (0.0,), (0.0,), (0.0,), (0.0,), (0.0,), (0.0,), (0.0,), (0.0,), (1.0,), (0.00161,), (0.0,), (0.0,), (0.0,)) # Especificação da T01 da URL-II
    # URLIII_T01.values = ((e9,), (e10,), (e13,), (0.0,), (0.0,), (0.0,), (0.0,), (0.0,), (0.0,), (0.0,), (0.0,), (0.0,), (1.0,), (0.00161,), (0.0,), (0.0,), (0.0,)) # Especificação da T01 da URL-III
    
    SS_URLI.Cell('C2').CellValue = e5 # FRAÇÃO MOLAR de C1 no fundo da T01 DA URL-1 (CALCULADA PELA SIMULAÇÃO RIGOROSA)
    SS_URLI.Cell('C3').CellValue = e6 # FRAÇÃO MOLAR DE C2 NO TOPO DA T01 DA URL-1 (CALCULADA PELA SIMULAÇÃO RIGOROSA)  
    SS_URLI.Cell('C4').CellValue = e11 # FRAÇÃO MOLAR DE C2 NO TOPO DA T01 DA URL-1 (CALCULADA PELA SIMULAÇÃO RIGOROSA)
    
    SS_URLII.Cell('C2').CellValue = e7 # FRAÇÃO MOLAR de C1 no fundo da T01 DA URL-1 (CALCULADA PELA SIMULAÇÃO RIGOROSA)
    SS_URLII.Cell('C3').CellValue = e8 # FRAÇÃO MOLAR DE C2 NO TOPO DA T01 DA URL-1 (CALCULADA PELA SIMULAÇÃO RIGOROSA)  
    SS_URLII.Cell('C4').CellValue = e12 # FRAÇÃO MOLAR DE C2 NO TOPO DA T01 DA URL-1 (CALCULADA PELA SIMULAÇÃO RIGOROSA) 
    
    SS_URLIII.Cell('C2').CellValue = e9 # FRAÇÃO MOLAR de C1 no fundo da T01 DA URL-1 (CALCULADA PELA SIMULAÇÃO RIGOROSA)
    SS_URLIII.Cell('C3').CellValue = e10 # FRAÇÃO MOLAR DE C2 NO TOPO DA T01 DA URL-1 (CALCULADA PELA SIMULAÇÃO RIGOROSA)  
    SS_URLIII.Cell('C4').CellValue = e13 # FRAÇÃO MOLAR DE C2 NO TOPO DA T01 DA URL-1 (CALCULADA PELA SIMULAÇÃO RIGOROSA) 
    
    
    # Versão 19 [CO2 nas URLs]
    SS_URLI.Cell('C5').CellValue = e29*0.95 # FRAÇÃO MOLAR de C1 no fundo da T01 DA URL-1 (CALCULADA PELA SIMULAÇÃO RIGOROSA)
    SS_URLII.Cell('C5').CellValue = e30*0.95 # FRAÇÃO MOLAR DE C2 NO TOPO DA T01 DA URL-1 (CALCULADA PELA SIMULAÇÃO RIGOROSA)  
    SS_URLIII.Cell('C5').CellValue = e31*0.95 # FRAÇÃO MOLAR DE C2 NO TOPO DA T01 DA URL-1 (CALCULADA PELA SIMULAÇÃO RIGOROSA)
  
    'Exportando as vazões das SANGRIAS'

    MT_main['Sangria URLs I-II'].StdLiqVolFlow.SetValue(e14,'m3/d')
    MT_main['Sangria URL III'].StdLiqVolFlow.SetValue(e15,'m3/d')
  
    
    'Exportando as ESPECIFICAÇÕES para as UPCGNs'
  
    SS_UPCGN.Cell('B3').CellValue = 0.9933 *e17/100 #fração molar  de C2 GLP na T1 da UPCGN-II
    SS_UPCGN.Cell('B4').CellValue = 0.9933 *e18/100 #fração molar C2 GLP na T1 da UPCGN-III
    SS_UPCGN.Cell('B5').CellValue = 0.9933 *e19/100 #fração molar C2 GLP na T1 da UPCGN-IV
    
 
    SS_UPCGN.Cell('C6').CellValue = e20 #Pressão do vaso 1 da UPCGN-II
    SS_UPCGN.Cell('C7').CellValue = e21 #Pressão do vaso 1 da UPCGN-III
    SS_UPCGN.Cell('C8').CellValue = e22 #Pressão do vaso 1 da UPCGN-IV
 
    SS_UPCGN.Cell('B9').CellValue = e23 #Recuperação de C4 no topo da T02 da UPCGN-II
    SS_UPCGN.Cell('B10').CellValue = e24 #Recuperação de C4 no topo da T02 da UPCGN-III
    SS_UPCGN.Cell('B11').CellValue = e25 #Recuperação de C4 no topo da T02 da UPCGN-IV
 
    SS_UPCGN.Cell('B12').CellValue = e26 #Recuperação de C5p no fundo da T02 da UPCGN-II
    SS_UPCGN.Cell('B13').CellValue = e27 #Recuperação de C5p no fundo da T02 da UPCGN-III
    SS_UPCGN.Cell('B14').CellValue = e28 ##Recuperação de C5p no fundo da T02 da UPCGN-IV

    # MT_UPGN['P24'].Temperature.SetValue(T_P24_UPGN,'C')   # corrente de entrada da T01 da UPGN-II
    
    'Exportando o status das unidades'

    SS_STATUS_UNIDADES.Cell('C3').CellValue = status_unit['URGN']
    SS_STATUS_UNIDADES.Cell('C4').CellValue = status_unit['URLI']
    SS_STATUS_UNIDADES.Cell('C5').CellValue = status_unit['URLII']
    SS_STATUS_UNIDADES.Cell('C6').CellValue = status_unit['URLIII']
    SS_STATUS_UNIDADES.Cell('C7').CellValue = status_unit['UPGN']
    SS_STATUS_UNIDADES.Cell('C8').CellValue = status_unit['UPCGNII']
    SS_STATUS_UNIDADES.Cell('C9').CellValue = status_unit['UPCGNIII']
    SS_STATUS_UNIDADES.Cell('C10').CellValue = status_unit['UPCGNIV']
    SS_STATUS_UNIDADES.Cell('C11').CellValue = status_unit['GASDUC']
    SS_STATUS_UNIDADES.Cell('C12').CellValue = status_unit['MIX']
    
    code_EspVar=0 # AINDA NÃO É ÚTIL [SERÁ USADO COM UMA FLAG PARA INDICAR O FUNCIONAMENTO DA ROTINA]

    return code_EspVar

def Hysysconect(filename):
    
    '''
    *************************************************************************************************************************************
    [1] DESCRIÇÃO: Hysysconect: Rotina para ligar o Hysys
    
    [2] EXPLICAÇÃO: Essa é utilizada para realizar a ativação do Hysys e comunicação com a simulação 
    
    [3] DADOS DE ENTRADA: 
        filename-> Nome do arquivo a ser simulado;
    
    [4] DADOS DE SAÍDA: 
        filename  -> Flag para indicar sucesso ou insucesso da comunicação
        simCase   -> Objeto resultante da comunicação
        hyApp     -> Objeto resultante da comunicação
        
    [5] OBSERVAÇÕES: Ao longo do código adicionaremos algumas anotações para facilitar a compreensão das considerações utilizadas
    *************************************************************************************************************************************
    '''
    
    'Conexao com Hysys'
    
    filepath = os.path.dirname(os.path.realpath(__file__)) # pasta caminho do arquivo...
    print("Abrindo o aplicativo Hysys")
    hyFilePath = f"{filepath}\\{filename}" # localização do arquivo
    print(f"Abrindo a simulação {filename}")
    hyApp = win32.Dispatch('Hysys.Application') # Abertura do Hysys
    hyApp.visible = True # Ativar visibilidade
    simCase = hyApp.SimulationCases.Open(hyFilePath) # Abrir a simulação
    simCase.Activate()  # Abrir a simulação
    simCase.Solver.CanSolve = False # Deixar a simulação em hold
      
    
    if simCase.Solver.CanSolve == 'False': # Lógica para averiguar sucesso da comunicação (vou melhorar!!!)
        code_hconect = 0
    else:
        code_hconect = 1
        
    return code_hconect, simCase, hyApp # retorno da função

def Inputdata():
    
    '''
    *************************************************************************************************************************************
    [1] DESCRIÇÃO: Inputdata: Lê dados da planilha excel Input_Data.xlsx.
    
    [2] EXPLICAÇÃO: Essa função lê dados das: 
            condições dos coletores;
            composição de entrada dos coletores;
            valor inicial das variáveis manipuladas;
            preços;
            especificações das unidades;
            restrições de capacidades;
            restrições dos produtos
    
    [3] DADOS DE ENTRADA: 
        A função, na versão atual (V13) não possui dados de entrada.
    
    [4] DADOS DE SAÌDA: 
        edata     -> Dicionário contendo todos os dados lidos.
        code_input -> Flag para averiguar saída de dados [se code_input=0, ok]
         
    [5] OBSERVAÇÕES: A saída dessa função é usada nas funções SpecVar, SpecSLP e SLP
    *************************************************************************************************************************************
    '''
    
    
    
    'Leitura dos dados desde arquivo externo'
    
    Condicoes_coletores = pd.read_excel('Input_Data.xlsx','Condicoes_coletores')
    Comp_entrada_coletores = pd.read_excel('Input_Data.xlsx','Comp_entrada_coletores')
    Valor_inicial_manipuladas = pd.read_excel('Input_Data.xlsx','Valor_inicial_manipuladas')
    Precos = pd.read_excel('Input_Data.xlsx','Precos')
    Especificacoes_plantas = pd.read_excel('Input_Data.xlsx','Especificacoes_plantas')
    Restricoes_capacidade = pd.read_excel('Input_Data.xlsx','Restricoes_capacidade')
    Restricoes_produtos = pd.read_excel('Input_Data.xlsx','Restricoes_produtos')
    Custos  = pd.read_excel('Input_Data.xlsx','Custos')
    
    'Colocando os dados lidos nas dicionários correspondentes'
    
    col_295_condicoes = dict(zip(Condicoes_coletores['Variavel'], Condicoes_coletores['Col 295']))
    col_299_condicoes = dict(zip(Condicoes_coletores['Variavel'], Condicoes_coletores['Col 299']))
    col_302_condicoes = dict(zip(Condicoes_coletores['Variavel'], Condicoes_coletores['Col 302']))

    col_295_comp = dict(zip(Comp_entrada_coletores['Componente'], Comp_entrada_coletores['Col 295']))
    col_299_comp = dict(zip(Comp_entrada_coletores['Componente'], Comp_entrada_coletores['Col 299']))
    col_302_comp = dict(zip(Comp_entrada_coletores['Componente'], Comp_entrada_coletores['Col 302']))

    valor_inicial_manipuladas = dict(zip(Valor_inicial_manipuladas['Variaveis'], Valor_inicial_manipuladas['Valor'])) 

    precos = dict(zip(Precos['Descripcao'], Precos['Valor']))

    espec_plantas = dict(zip(Especificacoes_plantas['Variavel'], Especificacoes_plantas['Valor'])) 

    restricoes_capacidade_LC = dict(zip(Restricoes_capacidade['Unidade'], Restricoes_capacidade['LC']))   
    restricoes_capacidade_UC = dict(zip(Restricoes_capacidade['Unidade'], Restricoes_capacidade['UC']))  

    restricoes_produtos_LC = dict(zip(Restricoes_produtos['Item'], Restricoes_produtos['LC']))   
    restricoes_produtos_UC = dict(zip(Restricoes_produtos['Item'], Restricoes_produtos['UC'])) 
    
    #VERSÃO 19 - INCLUSÃO DOS CUSTOS NA PLANILHA
    
    custos = dict(zip(Custos['Descricao'], Custos['Valor']))
    
    # versão 18/19: inclusão do status das unidades: ESSES STATUS IRÃO PARA A SPREADSHEET STATUS". Dependem dos valores das alocações.
    status_unit = {'URGN': 1, # VALORES DEFAULT
                   'URLI': 1,
                   'URLII': 1,
                   'URLIII': 1,
                   'UPGN': 1,
                   'UPCGNII': 1,
                   'UPCGNIII': 1,
                   'UPCGNIV': 1,
                   'GASDUC': 1,
                   'MIX': 1,}
    
    
    # Checando se as alocações estão zerada....nesse caso, zeramos o status da unidade [VERSÂO 18]
    for unidade, capacidade in restricoes_capacidade_UC.items():
        status = 'não operando' if capacidade < 1e3 else 'operando' # se for menor que 1e3
        print ('')
        print(f'{"#" * 27}\nunidade {unidade} {status}\n{"alocação máxima é nula, forçando status = 0 da unidade" if capacidade < 1e3 else "status=1"}')
        print (f'{"#" * 27}')
        # restricoes_capacidade_UC[unidade] = 0 if capacidade < 1e3 else capacidade  # zerando também alocação mínima
        # restricoes_capacidade_LC[unidade] = 0 if capacidade < 1e3 else restricoes_capacidade_LC[unidade]*1# zerando também alocação mínima
        status_unit[unidade] = 0 if capacidade < 1e3 else 1 # status_unit carrega informações do status das unidades...
    
   
    'Dicionario com valores lidos'
    
    edata = {'col_295_condicoes': col_295_condicoes, # edata contém todas as informações
             'col_299_condicoes': col_299_condicoes,
             'col_302_condicoes': col_302_condicoes,
             'col_295_comp': col_295_comp,
             'col_299_comp': col_299_comp,
             'col_302_comp': col_302_comp,
             'valor_inicial_manipuladas': valor_inicial_manipuladas,
             'precos': precos,
             'espec_plantas': espec_plantas,             
             'restricoes_capacidade_LC': restricoes_capacidade_LC,
             'restricoes_capacidade_UC': restricoes_capacidade_UC, 
             'restricoes_produtos_LC': restricoes_produtos_LC, 
             'restricoes_produtos_UC': restricoes_produtos_UC, 
             'status_unit': status_unit, # versão 19: inclusão do status das unidades
             'custos': custos, 
             }
    
    if edata == 0: # flag para averiguar saída de dados...
        code_input = 1;
    else:
        code_input = 0;
     
    return code_input, edata # retorno da função...

def SimulaLP(x, G_Rec_UPCGNs_in,G_Rec_UPCGNs_C_in,obj):
    
    '''
    *************************************************************************************************************************************
    [1] DESCRIÇÃO: SimulaLP: Rodar Simulação Essencial usando a Programação Linear
    
    [2] EXPLICAÇÃO: Essa rotina é utilizada para atualizar os valores das variáveis de decisão e rodar a simulação Essencial.  Após 
    o cálculo, obtemos os valores da Receita e condições da corrente de reciclo da UPCGN de saída.
    
    
    [3] DADOS DE ENTRADA: 
        x       -> Variáveis de decisão;
        G_Rec_UPCGNs_in    -> Vazão da corrente de reciclo da UPCGN que SAI das UPCGNs
        G_Rec_UPCGNs_C_in  -> Fração molar dos componentes da corrente de reciclo da UPCGN que SAI das UPCGNs
        obj             -> Dicionário contendo os objetos resultantes das variáveis e spreadsheets do hsysys que serão utilizados
    
    [4] DADOS DE SAÌDA: 
        Receita        -> valor da RECEITA
        Reciclo_UPCGN-  > Dicionário contendo dados da corrente de reciclo da UPCGN
        Carga_Unidades -> Dicionário contendo dados de restrições máximas de produtos
        R_cap          -> Dicionário contendo dados de restrições de capacidade das UNIDADES 
        Carga          -> Dicionário contendo valores das vazões de carga das UNIDADES

        
    [5] OBSERVAÇÕES: Ao longo do código adicionaremos algumas anotações para facilitar a compreensão das considerações utilizadas
    *************************************************************************************************************************************
    '''
    

    # x são os valores "estimativas iniciais" das vazões (variávei, s de decisão)

    'Descompactando o dicionário obj'
    
    MT_main = obj['MT_main']
    MT_URGN = obj['MT_URGN']
    MT_URLI = obj['MT_URLI'] 
    MT_URLII = obj['MT_URLII'] 
    MT_URLIII = obj['MT_URLIII'] 
    MT_UPGN  = obj['MT_UPGN']
    Solver = obj['Solver'] 
    
    'Especificaçao das Vazões (ENTRAM DIRETAMENTE NAS VAZÕES A1-A4, B1-B4, C3-C5)'
        
    MT_main['A1'].MolarFlow.SetValue(x[0],'m3/d_(gas)')         #M1
    MT_main['A2'].MolarFlow.SetValue(x[1],'m3/d_(gas)')         #M2
    MT_main['A3'].MolarFlow.SetValue(x[2],'m3/d_(gas)')         #M3
    MT_main['A4'].MolarFlow.SetValue(x[3],'m3/d_(gas)')         #M4
    
    MT_main['B1'].MolarFlow.SetValue(x[4],'m3/d_(gas)')         #M5
    MT_main['B2'].MolarFlow.SetValue(x[5],'m3/d_(gas)')         #M6
    MT_main['B3'].MolarFlow.SetValue(x[6],'m3/d_(gas)')         #M7
    MT_main['B4'].MolarFlow.SetValue(x[7],'m3/d_(gas)')         #M8 
    
    MT_main['C3'].MolarFlow.SetValue(x[8],'m3/d_(gas)')         #M9 
    MT_main['C4'].MolarFlow.SetValue(x[9],'m3/d_(gas)')         #M10 
    MT_main['C5'].MolarFlow.SetValue(x[10],'m3/d_(gas)')        #M11 
    
    MT_main['Gás de Reciclo UPCGNs RCY_BP'].MolarFlow.SetValue(G_Rec_UPCGNs_in,'m3/d_(gas)')  # Setando azão de Reciclo da UPCG de entrada
    MT_main['Gás de Reciclo UPCGNs RCY_BP'].ComponentMolarFractionValue=G_Rec_UPCGNs_C_in # Setando fração molar da corrente de Reciclo da UPCG de entrada

    'Execução da Simulação'
    
    Solver.CanSolve = True  # INICIANDO A SIMULAÇÃO
    status = Solver.CanSolve # STATUS DA SIMULAÇÃOe1
    assert status == True # CHECANDO SE ESTÁ OK
    
    'Cálculo da Receita'
    
    f_OBJ= obj['SS_f_OBJ'].Cell('C6').CellValue # Função objetivo calculada dentro do simulador
    Receita= obj['SS_f_OBJ'].Cell('C2').CellValue
    Custo= obj['SS_f_OBJ'].Cell('C3').CellValue
    
    'Vazão e Composição de Reciclo saindo da UPCGN'
    
    G_Rec_UPCGN_out = obj['MT_main']['Gás de Reciclo UPCGNs RCY'].MolarFlow.GetValue('m3/d_(gas)') # Vazão de Reciclo da UPCG de saída
    G_Rec_UPCGNs_C_out  = obj['MT_main']['Gás de Reciclo UPCGNs RCY'].ComponentMolarFractionValue # Fração molar da corrente de Reciclo da UPCG de saída
    # C4  = obj['MT_main']['GV_20C1atm'].ComponentMolarFractionValue[3:12] # COMPOSICOES De C4+ no GAS de VENDA
    # C4 = np.sum(C4) # em fração
    
    
    'NOVAS VARIÁVEIS v8 [verificar vazões de entrada das unidades]'
    
    'NOTA: Monitoramos as vazões de saída de gás, após o vaso de entrada, de cada unidade de modo a comparar com as vazões'
    'de entrada manipuladas pelo algoritmo de otimização...'
    
    CARGA_G_URGN = MT_URGN['2'].MolarFlow.GetValue("m3/d_(gas)") #Vazão da corrente de gás da URGN
    CARGA_G_URLI = MT_URLI['4outA'].MolarFlow.GetValue("m3/d_(gas)") #Vazão da corrente de gás da URLI
    CARGA_G_URLII = MT_URLII['4outA'].MolarFlow.GetValue("m3/d_(gas)") #Vazão da corrente de gás da URLII
    CARGA_G_URLIII = MT_URLIII['4outA'].MolarFlow.GetValue("m3/d_(gas)") #Vazão da corrente de gás da URLIII
    CARGA_G_UPGN = MT_UPGN['P03'].MolarFlow.GetValue("m3/d_(gas)") #Vazão da corrente de gás da UPGN
    
    Reciclo_UPCGN = {'G_Rec_UPCGN_out': G_Rec_UPCGN_out,
                     'G_Rec_UPCGNs_C_out': G_Rec_UPCGNs_C_out
                     }
    Carga_Unidades = {'CARGA_G_URGN': CARGA_G_URGN,
                     'CARGA_G_URLI': CARGA_G_URLI,
                     'CARGA_G_URLII': CARGA_G_URLII,
                     'CARGA_G_URLIII':CARGA_G_URLIII,
                     'CARGA_G_UPGN': CARGA_G_UPGN,
                     }
    
    return f_OBJ, Reciclo_UPCGN, Carga_Unidades, Receita, Custo # RETORNANDO A Função_Objetivo E VAZÃO DE RECICLO

def SimulaLP_closed(x,MaterialStreams_C, Solver_C,Spread_sheet_C):
    
    'NÃO ESTAMOS UTILIZANDO'
    
    

        # x são os valores "estimativas iniciais" das vazões (variáveis de decisão)
        # y são as temperaturas iniciais
        # z são as pressões iniciais
        
#   Especificaçao das Vazões (ENTRAM DIRETAMENTE NAS VAZÕES A1-A4, B1-B4, C3-C5)
        
    MaterialStreams_C['A1'].MolarFlow.SetValue(x[0],'m3/d_(gas)')         #M1
    MaterialStreams_C['A2'].MolarFlow.SetValue(x[1],'m3/d_(gas)')         #M2
    MaterialStreams_C['A3'].MolarFlow.SetValue(x[2],'m3/d_(gas)')         #M3
    MaterialStreams_C['A4'].MolarFlow.SetValue(x[3],'m3/d_(gas)')         #M4
    
    MaterialStreams_C['B1'].MolarFlow.SetValue(x[4],'m3/d_(gas)')         #M5
    MaterialStreams_C['B2'].MolarFlow.SetValue(x[5],'m3/d_(gas)')         #M6
    MaterialStreams_C['B3'].MolarFlow.SetValue(x[6],'m3/d_(gas)')         #M7
    MaterialStreams_C['B4'].MolarFlow.SetValue(x[7],'m3/d_(gas)')         #M8 
    
    MaterialStreams_C['C3'].MolarFlow.SetValue(x[8],'m3/d_(gas)')         #M9 
    MaterialStreams_C['C4'].MolarFlow.SetValue(x[9],'m3/d_(gas)')         #M10 
    MaterialStreams_C['C5'].MolarFlow.SetValue(x[10],'m3/d_(gas)')        #M11 
    

    # EXECUÇÃO DA SIMULAÇÃO E VERIFICAÇÃO DE CONVERGÊNCIA
    
    Solver_C.CanSolve = True  # INICIANDO A SIMULAÇÃO
    status = Solver_C.CanSolve # STATUS DA SIMULAÇÃO
    assert status == True # CHECANDO SE ESTÁ OK
    Receita=Spread_sheet_C.Cell('B9').CellValue # OBTENÇÃO DA RECEITA
    G_Rec_UPCGNs = MaterialStreams_C['Gás de Reciclo UPCGNs RCY'].MolarFlow.GetValue('m3/d_(gas)') # OBTENÇÃO DA VAZÃO UPCGN
    G_Rec_UPCGNs_C  = MaterialStreams_C['Gás de Reciclo UPCGNs RCY'].ComponentMolarFractionValue # COMPOSICOES DO RECICLO
    C4  = MaterialStreams_C['GV_20C1atm'].ComponentMolarFractionValue[3:12] # COMPOSICOES De C4+ no GAS de VENDA
    C4 = np.sum(C4) # em fração
    
  
    return Receita, G_Rec_UPCGNs, G_Rec_UPCGNs_C,C4 # RETORNANDO A RECEITA E VAZÃO DE RECICLO

def Spec_prods(x, delta_MIX, obj):
    
    '''
    *************************************************************************************************************************************
    [1] DESCRIÇÃO: Spec_prods: Rotina que obtem os valores das restrições de produtos. 
    
    [2] EXPLICAÇÃO: Essa função é utilizada para obter os valores das restrições de produtos, dadas os valores das variáveis
    de decisão. A função tamb´me pode ser utilizada para calcular as derivadas das derivadas das especificações em relação às
    variáveis de decisão.
    
    [3] DADOS DE ENTRADA: 
        x -> variáveis de decisão;
        delta_MIX -> magnitude da perturbação da vazão MIX. OBS: Em princípio, só consideramos essa variável para o cálculo
        da derivada.
        obj       -> Dicionário contendo os objetos resultantes das variáveis e spreadsheets do hsysys que serão utilizados
    
    [4] DADOS DE SAÌDA: 
        specs      -> Dicionário contendo os valores das especificações de produtos resultantes da simulação.
         
    [5] OBSERVAÇÕES: A saída dessa função é usada nas funções SpecVar, SpecSLP e SLP
    *************************************************************************************************************************************
    '''
    
    
    # Desempacotando o dicionário obj
    
    MT_main = obj['MT_main']
    # SS_Receita = obj['SS_Receita'] 
    SS_Rest = obj['SS_Rest']
    
#   Especificaçao das Vazões (ENTRAM DIRETAMENTE NAS VAZÕES A1-A4, B1-B4, C3-C5)
        
    MT_main['A1'].MolarFlow.SetValue(x[0],'m3/d_(gas)')         #M1
    MT_main['A2'].MolarFlow.SetValue(x[1],'m3/d_(gas)')         #M2
    MT_main['A3'].MolarFlow.SetValue(x[2],'m3/d_(gas)')         #M3
    MT_main['A4'].MolarFlow.SetValue(x[3],'m3/d_(gas)')         #M4
    
    MT_main['B1'].MolarFlow.SetValue(x[4],'m3/d_(gas)')         #M5
    MT_main['B2'].MolarFlow.SetValue(x[5],'m3/d_(gas)')         #M6
    MT_main['B3'].MolarFlow.SetValue(x[6],'m3/d_(gas)')         #M7
    MT_main['B4'].MolarFlow.SetValue(x[7],'m3/d_(gas)')         #M8 
    
    MT_main['C3'].MolarFlow.SetValue(x[8],'m3/d_(gas)')         #M9 
    MT_main['C4'].MolarFlow.SetValue(x[9],'m3/d_(gas)')         #M10 
    MT_main['C5'].MolarFlow.SetValue(x[10]+delta_MIX,'m3/d_(gas)')        #M11   
    
    y0=SS_Rest.Cell('G2').CellValue # C1_GV
    y1=SS_Rest.Cell('G3').CellValue # C2_GV
    y2=SS_Rest.Cell('G4').CellValue # C3_GV
    y3=SS_Rest.Cell('G5').CellValue # C4+_GV
    y4=SS_Rest.Cell('G6').CellValue # CO2_GV
    y5=SS_Rest.Cell('G7').CellValue # INERTES_GV
    y6=SS_Rest.Cell('G8').CellValue # PCS_GV
    y7=SS_Rest.Cell('G9').CellValue # POA_GV
    y8=SS_Rest.Cell('G10').CellValue # POH_GV
    y9=SS_Rest.Cell('G11').CellValue # Wobbe_GV
    y10=SS_Rest.Cell('G12').CellValue # No_Metano_GV
    
    y11=SS_Rest.Cell('G14').CellValue # C2_GLP
    y12=SS_Rest.Cell('G15').CellValue # C5+_CLP
    y13=SS_Rest.Cell('G16').CellValue # RHO_GLP
    
    # ppm_H2O_GV = SS_Rest.Cell('G16').CellValue # ppm GV
    # ppm_H2O_GV_max = Spread_sheet_R.Cell('G17').CellValue # ppm Máximo do GV
    # F_agua = MaterialStreams['h2o_in'].MolarFlow.GetValue("m3/d_(gas)") #Vazão necessária de água
    # F_MIX_plus = MaterialStreams['H2O_poa'].MolarFlow.GetValue("m3/d_(gas)")  #Vazão molar do MIX_plus
    
    specs = np.array([y0,y1,y2,y3,y4,y5,y6,y7,y8,y9,y10,y11,y12,y13]) #valores das restrições

    
    return specs

def f_POA(x,F_MIX_plus_est,deltaM,MaterialStreams,Spread_sheet_R):
    
    MaterialStreams['H2O_poa'].MolarFlow.SetValue(F_MIX_plus_est,'m3/d_(gas)') # vazão_base da MIX_plus  
    F_MIX_plus_base = MaterialStreams['H2O_poa'].MolarFlow.GetValue("m3/d_(gas)")  #Vazão molar do MIX_plus       
    POA_base=Spread_sheet_R.Cell('F21').CellValue # POA
    
    MaterialStreams['H2O_poa'].MolarFlow.SetValue(F_MIX_plus_est+deltaM,'m3/d_(gas)') # vazão_base da MIX_plus 
    POA_after=Spread_sheet_R.Cell('F21').CellValue # POA
    
    return POA_base, POA_after, F_MIX_plus_base

def plot_derivatives(dR_dF, index):
    fig, ax = plt.subplots()

    title_list = [f'Derivadas iter{i}' for i in range(11)]
    barlabels = ['dA1', 'dA2', 'dA3', 'dA4' , 'dB1', 'dB2', 'dB3', 'dB4', 'dC3', 'dC4', 'dC5']
    bar_colors = ['tab:red', 'tab:blue', 'tab:green', 'tab:orange', 'tab:pink', 'tab:gray', 'tab:brown', 'tab:olive', 'tab:purple', 'tab:cyan', 'tab:gray']
    n_dec = np.arange(len(dR_dF))

    ax.bar(n_dec, dR_dF, color=bar_colors, label=barlabels, width=0.9)
    ax.set_ylabel('Derivadas')
    ax.set_title(title_list[index])
    ax.legend(title='Derivadas', loc=(1, 0))

    plt.show()

def plot_manipuladas(manip, index):
    

    var_list = ['G_295toGASDUC', 'G_295toURGN', 'G_295toURLs', 'G_295toUPGN',
                'G_299toGASDUC', 'G_299toURGN', 'G_299toURLs', 'G_299toUPGN',
                'G_302toURLs', 'G_302toUPGN', 'G_302toMIX']
    n_manip = manip.shape[1]
    for var in range(n_manip):
        plt.figure()
        plt.title(f'Manipulada {var_list[var]}')
        plt.plot(np.arange(1, index+1), manip[0:index, var])
        plt.ylabel('Vazão [m3/d/_gas]')
        plt.xlabel('Iteração')
    
    return 0

def f_Rel(model, rel_SLP):
    workbook = xlsxwriter.Workbook('demo.xlsx')
    worksheet = workbook.add_worksheet('Iteração_1')
    
    # Widen the first column to make the text clearer.
    worksheet.set_column('A:A', 20)
    
    # Add a bold format to use to highlight cells.
    bold = workbook.add_format({'bold': True})
    
    # Write some simple text.
    worksheet.write('A1', 'Hello')
    
    # Text with formatting.
    worksheet.write('A2', 'World', bold)
    
    # Write some numbers, with row/column notation.
    worksheet.write(2, 0, 123)
    worksheet.write(3, 0, 123.456)
    
    # # Insert an image.
    # worksheet.insert_image('B5', 'logo.png')
    
    workbook.close()
    
    out_rel = 1
    
    return out_rel

def f_Plot(model, rel_SLP):
    
    Receita_base = rel_SLP['Receita_base']
    Receita      = rel_SLP['Receita']
    Iterações    = rel_SLP['Iterações']
    Desvio       = rel_SLP['Desvio']
    
    plt.style.use('_mpl-gallery')
  
    x = Iterações
    y = Receita
    z = Desvio

    fig, ax = plt.subplots()
    fig2, ax2 = plt.subplots()

    ax.plot(x, y, 'o-')
    ax2.plot(x, z, 'o-')

    ax.set(xlabel='Iterações', ylabel='Receita [Mil R$/dia]', title='Receita x Iterações')
    ax.grid()
    ax2.set(xlabel='Iterações', ylabel='Desvio',
    title='Desvio x Iterações')
    ax2.grid()


    plt.show()

    # Plotando as manipuladas
    # plot_manipuladas(manip, index)
    
    code_f_Plot = 1
    return code_f_Plot













