# -*- coding: utf-8 -*-
"""
Created on Fri Jan  5 21:59:39 2024

@author: lizst
"""

import numpy as np
import pandas as pd
import json
import os
from numpy import linalg as la
import win32com.client as win32
from pulp import pulp, LpMaximize, LpProblem, LpVariable
from scipy.optimize import Bounds, minimize

def ler_config():
    file = open("engine_config.json", "r")
    config = json.load(file)
    file.close()
    return config
"*** Fim de Função para ler o arquivo json com as configurações ***"

"*** Função para ler o arquivo json com os dados de entrada ***"
def ler_inputs():
    file = open("engine_inputs.json", "r")
    inputs = json.load(file)
    file.close()
    return inputs
"*** Fim de Função para ler o arquivo json com os dados de entrada ***"

def aloca_cargas(sim_case, config, inputs):
    
    #!# Vazões de entrada aos coletores
    Q_in_295 = inputs["simulacao"]["coletores"]["SG-295"]["carga"]["valor"]
    Q_in_299 = inputs["simulacao"]["coletores"]["SG-299"]["carga"]["valor"]
    Q_in_302 = inputs["simulacao"]["coletores"]["SG-302"]["carga"]["valor"]

    #!# Disponibilidade de cada unidade e total
    Q_max_GASDUC = inputs["simulacao"]["disponibilidade"]["GASDUC II"]["valor"]
    Q_max_MIX_UTGCAB = inputs["simulacao"]["disponibilidade"]["MIX UTGCAB"]["valor"]
    Q_max_UPGN = inputs["simulacao"]["disponibilidade"]["UPGN II"]["valor"]
    Q_max_URGN = inputs["simulacao"]["disponibilidade"]["URGN"]["valor"]
    Q_max_URL_I = inputs["simulacao"]["disponibilidade"]["URL I"]["valor"]
    Q_max_URL_II = inputs["simulacao"]["disponibilidade"]["URL II"]["valor"]
    Q_max_URL_III = inputs["simulacao"]["disponibilidade"]["URL III"]["valor"]
    Q_max_URLs = Q_max_URL_I + Q_max_URL_II + Q_max_URL_III
    Q_max_Total = Q_max_GASDUC + Q_max_MIX_UTGCAB + Q_max_UPGN + Q_max_URGN + Q_max_URLs

    # |# Cargas iniciais para uma primeira simulação
    # Gás do 302
    Q_URCO2toU211 = 0
    Q_MIX_UTGCAB = min(Q_max_MIX_UTGCAB, Q_in_302)
    Q_BypassU211 = min(Q_max_UPGN, (Q_in_302 - Q_max_MIX_UTGCAB))
    # Gás do 299
    Q_299toUPGN = 0
    Q_299to295 = min(Q_in_299, Q_max_GASDUC)
    # Gás do 295
    Q_295toUPGN = 0
    Q_GASDUC_II = Q_max_GASDUC
    Q_Carga_URGN = min(Q_max_URGN, (Q_in_295 + Q_299to295))
        
    # Gás restante para as URLs
    resto_gas = 0.96*(Q_in_295 + Q_in_299 + Q_in_302) - (Q_MIX_UTGCAB + Q_BypassU211
                                                         + Q_295toUPGN + Q_GASDUC_II
                                                         + Q_Carga_URGN + Q_299toUPGN
                                                         + Q_299to295 + Q_URCO2toU211
                                                         )
    Q_Carga_URL_I = 1/3*resto_gas
    Q_Carga_URL_III = 1/3*resto_gas
    
    # Colocando essas cargas iniciais em um dicionario
    cargas = {"Q_MIX_UTGCAB": Q_MIX_UTGCAB,
            "Q_BypassU211": Q_BypassU211,
            "Q_295toUPGN": Q_295toUPGN,
            "Q_GASDUC_II": Q_GASDUC_II,
            "Q_Carga_URGN": Q_Carga_URGN,
            "Q_299toUPGN": Q_299toUPGN,
            "Q_299to295": Q_299to295,
            "Q_URCO2toU211": Q_URCO2toU211,
            "Q_Carga_URL_I": Q_Carga_URL_I,
            "Q_Carga_URL_III": Q_Carga_URL_III,
            }

    # # |# Realizando a primeira simulação
    # x = []
    # resultados_simulacao = simula(sim_case, config, inputs, cargas, x)

    # # |# Lendo as vazões de gás de saida dos coletores, da queima e do gás de reciclo das UPCGNs
    # Q_295Gout = resultados_simulacao["Coletores"]["Col_Qgas"]["COL295_QGAS"]
    # Q_299Gout = resultados_simulacao["Coletores"]["Col_Qgas"]["COL299_QGAS"]
    # Q_302Gout = resultados_simulacao["Coletores"]["Col_Qgas"]["COL302_QGAS"]
    # Q_gas_reciclo_UPCGNs = resultados_simulacao["Coletores"]["Col_Qgas"][
    #     "Gas_Reciclo_UPCGNs"
    # ]
    # Queima295 = resultados_simulacao["producao"]["queima"]["queima_por_coletor"][
    #     "queima295"
    # ]
    # Queima299 = resultados_simulacao["producao"]["queima"]["queima_por_coletor"][
    #     "queima299"
    # ]
    # Queima302 = resultados_simulacao["producao"]["queima"]["queima_por_coletor"][
    #     "queima302"
    # ]

    # # |# Total de Gas a ser distribuido para as unidades
    # Q_gas_total = (
    #     Q_295Gout
    #     - Queima295
    #     + Q_gas_reciclo_UPCGNs / 2
    #     + Q_299Gout
    #     - Queima299
    #     + Q_302Gout
    #     - Queima302
    # )

    # # |# Verificando se a capacidade da planta é suficiente para tratar todo o gás
    # assert (
    #     Q_gas_total <= Q_max_Total
    # ), "A carga de gás é maior do que a capacidade da planta"

    # # |# Otimização linear para encontrar a alocação segundo a priorização do usuário

    # # Lendo as duas primeras colocadas na priorização
    # prioridade = list(inputs["simulacao"]["priorizacaoCarga"])[0:3]
    # lista_opcoes = ["GAS", "MIX", "UPG", "URG", "URL"]
    # indice1 = lista_opcoes.index(prioridade[0][0:3])
    # indice2 = lista_opcoes.index(prioridade[1][0:3])
    # indice3 = lista_opcoes.index(prioridade[2][0:3])

    # # Criando o modelo
    # model = LpProblem(name="Cabiunas_LP", sense=LpMaximize)

    # # Variaveis
    # G_302toUPGN = LpVariable("Q_BypassU211", lowBound=0, upBound=Q_max_UPGN)
    # G_302toMIX = LpVariable("Q_MIX_UTGCAB", lowBound=0, upBound=Q_max_MIX_UTGCAB)
    # G_302toURLs = LpVariable("Gás de 302 para URLs", lowBound=0, upBound=Q_max_URLs)
    # G_299to295 = LpVariable("Q_299to295", lowBound=0, upBound=Q_max_GASDUC)
    # G_299toUPGN = LpVariable("Q_299toUPGN", lowBound=0, upBound=Q_max_UPGN)
    # G_299toURLs = LpVariable("Gás de 299 para URLs", lowBound=0, upBound=Q_max_URLs)
    # G_295toGASDUC = LpVariable("Q_GASDUC_II", lowBound=0, upBound=Q_max_GASDUC)
    # G_295toURGN = LpVariable("Q_Carga_URGN", lowBound=0, upBound=Q_max_URGN)
    # G_295toUPGN = LpVariable("Q_295toUPGN", lowBound=0, upBound=Q_max_UPGN)
    # G_295toURLs = LpVariable("Gás de 295 para URLs", lowBound=0, upBound=Q_max_URLs)
    # Cargas = LpVariable.dicts(
    #     "Cargas", range(0, 5), lowBound=0
    # )  # ['GAS','MIX','UPG','URG','URL']

    # "Função Objetivo (depende da priorizaçao)"
    # model += (
    #     0.5 * Cargas[indice1]
    #     + 0.3 * Cargas[indice2]
    #     + 0.2 * Cargas[indice3]
    #     - 0.1 * G_299to295
    # )

    # "Restrições"
    # model += (
    #     G_302toUPGN + G_302toMIX + G_302toURLs == 0.999999 * (Q_302Gout - Queima302),
    #     "Alocar todo o gás do 302",
    # )
    # model += (
    #     G_299to295 + G_299toUPGN + G_299toURLs == 0.999999 * (Q_299Gout - Queima299),
    #     "Alocar todo o gás do 299",
    # )
    # model += (
    #     G_295toGASDUC + G_295toURGN + G_295toUPGN + G_295toURLs
    #     == Q_295Gout - Queima295 + Q_gas_reciclo_UPCGNs / 2 + G_299to295,
    #     "Alocar todo o gás do 295, do Reciclo e do 299 enviado para 295",
    # )
    # model += (G_295toGASDUC == Cargas[0], "GASDUC")
    # model += (G_302toMIX == Cargas[1], "MIX")
    # model += (G_302toUPGN + G_299toUPGN + G_295toUPGN == Cargas[2], "UPGN")
    # model += (G_295toURGN == Cargas[3], "URGN")
    # model += (G_302toURLs + G_299toURLs + G_295toURLs == Cargas[4], "URLs")
    # model += (Cargas[0] >= config["LB otimizacao"]["Q_GASDUC_II"], "GASDUC mínimo")
    # model += (Cargas[0] <= Q_max_GASDUC, "GASDUC máximo")
    # model += (Cargas[1] >= config["LB otimizacao"]["Q_MIX_UTGCAB"], "MIX mínimo")
    # model += (Cargas[1] <= Q_max_MIX_UTGCAB, "MIX máximo")
    # model += (Cargas[2] >= Q_max_UPGN / 2, "UPGN mínimo")
    # model += (Cargas[2] <= Q_max_UPGN, "UPGN máximo")
    # model += (Cargas[3] >= Q_max_URGN / 2, "URGN mínimo")
    # model += (Cargas[3] <= Q_max_URGN, "URGN máximo")
    # model += (Cargas[4] >= Q_max_URLs / 2, "URLs mínimo")
    # model += (Cargas[4] <= Q_max_URLs, "URLs máximo")

    # "Executa otimização"
    # status = model.solve(pulp.PULP_CBC_CMD(msg=False))

    # # |# Alocação encontrada
    # Q_MIX_UTGCAB = G_302toMIX.value()
    # Q_BypassU211 = G_302toUPGN.value()
    # Q_295toUPGN = G_295toUPGN.value()
    # Q_GASDUC_II = G_295toGASDUC.value()
    # Q_Carga_URGN = G_295toURGN.value()
    # Q_299toUPGN = G_299toUPGN.value()
    # Q_299to295 = G_299to295.value()
    # Q_URCO2toU211 = 0

    # cargas = {
    #     "Q_MIX_UTGCAB": Q_MIX_UTGCAB,
    #     "Q_BypassU211": Q_BypassU211,
    #     "Q_295toUPGN": Q_295toUPGN,
    #     "Q_GASDUC_II": Q_GASDUC_II,
    #     "Q_Carga_URGN": Q_Carga_URGN,
    #     "Q_299toUPGN": Q_299toUPGN,
    #     "Q_299to295": Q_299to295,
    #     "Q_URCO2toU211": Q_URCO2toU211,
    # }

    return cargas

"*** Função para conferir se as restriçoes são satisfeitas e calcular penalização por restriçoes ***"
def confere_restricoes(config, inputs, valor_restricoes):
    # Restriçoes das unidades
    carga_GASDUC = {
        "superior": inputs["simulacao"]["disponibilidade"]["GASDUC II"]["valor"],
        "inferior": 1e-4,
        "unidade": "m3/d_(gas)",
    }
    carga_MIX_UTGCAB = {
        "superior": inputs["simulacao"]["disponibilidade"]["MIX UTGCAB"]["valor"],
        "inferior": 1e-4,
        "unidade": "m3/d_(gas)",
    }
    carga_UPGN_II = {
        "superior": inputs["simulacao"]["disponibilidade"]["UPGN II"]["valor"],
        "inferior": (inputs["simulacao"]["disponibilidade"]["UPGN II"]["valor"]) / 2,
        "unidade": "m3/d_(gas)",
    }
    carga_URGN = {
        "superior": inputs["simulacao"]["disponibilidade"]["URGN"]["valor"],
        "inferior": (inputs["simulacao"]["disponibilidade"]["URGN"]["valor"]) / 2,
        "unidade": "m3/d_(gas)",
    }
    carga_URLs = {
        "superior": (
            inputs["simulacao"]["disponibilidade"]["URL I"]["valor"]
            + inputs["simulacao"]["disponibilidade"]["URL II"]["valor"]
            + inputs["simulacao"]["disponibilidade"]["URL III"]["valor"]
        ),
        "inferior": (
            (
                inputs["simulacao"]["disponibilidade"]["URL I"]["valor"]
                + inputs["simulacao"]["disponibilidade"]["URL II"]["valor"]
                + inputs["simulacao"]["disponibilidade"]["URL III"]["valor"]
            )
        )
        / 2,
        "unidade": "m3/d_(gas)",
    }
    carga_UPCGNs = {"superior": 4500, "inferior": 750, "unidade": "m3/d)"}
    # Agrupando em um dicionario as restriçoes das unidades
    restricoes_unidades = {
        "carga_GASDUC": carga_GASDUC,
        "carga_MIX_UTGCAB": carga_MIX_UTGCAB,
        "carga_UPGN_II": carga_UPGN_II,
        "carga_URGN": carga_URGN,
        "carga_URLs": carga_URLs,
        "carga_UPCGNs": carga_UPCGNs,
    }
    # Criando o dicionario de todas as restricoes (disponibilidade das unidade e qualidade)
    restricoes = restricoes_unidades | config["restricoes_qualidade"]
    # Alocando os array
    c_low = {}
    c_up = {}
    # Calculando as penalizações por violação dos limites
    for key in restricoes:
        c_low[key] = (restricoes[key]["inferior"] - valor_restricoes[key]) / abs(
            restricoes[key]["inferior"]
        )
        c_up[key] = (valor_restricoes[key] - restricoes[key]["superior"]) / abs(
            restricoes[key]["superior"]
        )
    # Calculando c e c_penal
    c = np.r_[list(c_low.values()), list(c_up.values())]
    c_penal = sum(c[np.where(c > 0)[0]])
    # Avaliando se todas as restriçoes foram satisfeitas
    if c_penal == 0:
        viavel = True
    else:
        viavel = False
    # Armazenando resultados em dicionario
    penal_rest = {
        "c_low": c_low,
        "c_up": c_up,
        "c": list(c),
        "c_penal": c_penal,
        "viavel": viavel,
    }
    return penal_rest
"*** Função para conferir se as restriçoes são satisfeitas e calcular penalização por restriçoes ***"


def PCSCalculadoISO(FracaoMolar):
    x = np.array(FracaoMolar)
    ncomp = len(x)
    GrossCalValue = np.zeros(ncomp)
    GrossCalValue[0] = 891.05  # C1
    GrossCalValue[1] = 1561.42  # C2
    GrossCalValue[2] = 2220.13  # C3
    GrossCalValue[3] = 2869.39  # iC4
    GrossCalValue[4] = 2878.58  # nC4
    GrossCalValue[5] = 3530.25  # iC5
    GrossCalValue[6] = 3537.19  # nC5
    GrossCalValue[7] = 4196.6  # C6
    GrossCalValue[8] = 4855.31  # C7
    GrossCalValue[9] = 5513.9  # C8
    GrossCalValue[10] = 6173.48  # C9
    GrossCalValue[11] = 6832.33  # C10
    GrossCalValue[12] = 0  # N2
    GrossCalValue[13] = 0  # CO2
    GrossCalValue[14] = 0  # H2O
    GrossCalValue[15] = 0  # H2S
    GrossCalValue[16] = 0  # EGlycol
    PCSCalculadoISO = sum(GrossCalValue * x)  # kJ/mol
    PCSCalculadoISO = PCSCalculadoISO / 1054  # MMbtu/kmol
    return PCSCalculadoISO
"*** Fim de Função para calcular o PCS (ISO 6976) ***"



"*** Função para executar a simulação detalhada e retornar os resultados ***"
def simula_detalhada(simCase, config, inputs, cargas):
    
    '''
    *************************************************************************************************************************************
    [1] DESCRIÇÃO: Função que roda a simulação detalhada 
    
    [2] EXPLICAÇÃO: Essa rotina é utilizada para especificar algumas variáveis das UNIDADES URLS, URGN, UPGN e UPCGN. Além
    disso importaremos os valores dos preços e condições dos coletores. A maior parte
    das variáveis é especificada no arquivo Input_Data.xlsx. Outras variáveis (frações molares da T02 das URLs) são calculadas na 
    simulação rogorosa (utilizada no modo offline) e importadas para a simulação Essencial.
    
    [3] DADOS DE ENTRADA: 
        edata   -> Dicionário resultante da leitura da dados da planilha Input_Data.xlsx;
        obj     -> Dicionário contendo os objetos resultantes das variáveis e spreadsheets do hysys que serão utilizados
        R_especs-> Dicionário contendo as especificações importadas da simulação rigorosa
    
    [4] DADOS DE SAÌDA: 
        cod_SpecVar  -> Flag para indicar sucesso ou insucesso do cálculo
        
    [5] OBSERVAÇÕES: Ao longo do código adicionaremos algumas anotações para facilitar a compreensão das considerações utilizadas
    
    
    Histórico:
        
    (a) 08/04/2024 - Inclusão de dados da corrente de CO2 das URLS para serem exportadas para a simulação Essencial
    *************************************************************************************************************************************
    '''  
    
    simCase.Solver.CanSolve = False  

    # Criando objetos auxiliares
    Solver = simCase.Solver
    MaterialStreams = (simCase.Flowsheet.MaterialStreams)  # correntes de materia do flowsheet principal
    Operations = simCase.Flowsheet.Operations  # operações da flowsheet principal
    # EnergyStreams = simCase.Flowsheet.EnergyStreams # correntes de energia do flowsheet principal
    MaterialStreams_URLs = simCase.Flowsheet.Flowsheets("TPL4").MaterialStreams  # correntes de materia do subflowsheet URLs
    MaterialStreams_URGN = simCase.Flowsheet.Flowsheets("TPL8").MaterialStreams  # correntes de materia do subflowsheet URGN
    MaterialStreams_UPGN = simCase.Flowsheet.Flowsheets("TPL3").MaterialStreams  # correntes de materia do subflowsheet UPGN
    Operations_UPGN = simCase.Flowsheet.Flowsheets("TPL3").Operations.Item  # operações do subflowsheet UPGN
    Operations_UPCGN = simCase.Flowsheet.Flowsheets("TPL15").Operations.Item  # operações do subflowsheet UPCGN
    Operations_URLs = simCase.Flowsheet.Flowsheets("TPL4").Operations.Item  # operações do subflowsheet URLs
    MaterialStreams_UPCGN_II = (simCase.Flowsheet.Flowsheets("TPL15").Flowsheets("TPL17").MaterialStreams)  # correntes de materia do subflowsheet UPCGNII
    MaterialStreams_UPCGN_III = (simCase.Flowsheet.Flowsheets("TPL15").Flowsheets("TPL11").MaterialStreams)  # correntes de materia do subflowsheet UPCGNIII
    Operations_URL_II = (simCase.Flowsheet.Flowsheets("TPL4").Flowsheets("TPL6").Operations.Item)  # operações do subflowsheet URLII
    Operations_UPCGNIV = (simCase.Flowsheet.Flowsheets("TPL15").Flowsheets("TPL16").Operations.Item)  # operações do subflowsheet UPCGNIV

    #!# Fixando as varaiveis de entrada na simulação
    
    # Corrente de entrada ao Coletor 295
    MaterialStreams["GásRico295"].MolarFlow.SetValue(
        inputs["simulacao"]["coletores"]["SG-295"]["carga"]["valor"], "m3/d_(gas)"
    )  # Carga do coletor 295
    MaterialStreams["GásRico295"].Temperature.SetValue(
        inputs["simulacao"]["coletores"]["SG-295"]["temperatura"]["valor"], "C"
    )  # Temperatura do coletor 295
    MaterialStreams["GásRico295"].Pressure.SetValue(
        inputs["simulacao"]["coletores"]["SG-295"]["pressao"]["valor"], "kg/cm2_g"
    )  # Pressão do coletor 295
    MaterialStreams["GásRico295"].ComponentMolarFractionValue = (
        np.array(
            inputs["simulacao"]["coletores"]["SG-295"]["composicao"] + [0, 0, 0]
        )
        / 100
    )  # Fração molar coletor 295
    
    # Corrente de entrada ao Coletor 299
    MaterialStreams["GásRico299"].MolarFlow.SetValue(
        inputs["simulacao"]["coletores"]["SG-299"]["carga"]["valor"], "m3/d_(gas)"
    )  # Carga do coletor 299
    MaterialStreams["GásRico299"].Temperature.SetValue(
        inputs["simulacao"]["coletores"]["SG-299"]["temperatura"]["valor"], "C"
    )  # Temperatura do coletor 299
    MaterialStreams["GásRico299"].Pressure.SetValue(
        inputs["simulacao"]["coletores"]["SG-299"]["pressao"]["valor"], "kg/cm2_g"
    )  # Pressão do coletor 299
    MaterialStreams["GásRico295"].ComponentMolarFractionValue = (
        np.array(
            inputs["simulacao"]["coletores"]["SG-299"]["composicao"] + [0, 0, 0]
        )
        / 100
    )  # Fração molar coletor 295
    
    # Corrente de entrada ao Coletor 302
    MaterialStreams["GásRico302"].MolarFlow.SetValue(
        inputs["simulacao"]["coletores"]["SG-302"]["carga"]["valor"], "m3/d_(gas)"
    )  # Carga do coletor 302
    MaterialStreams["GásRico302"].Temperature.SetValue(
        inputs["simulacao"]["coletores"]["SG-302"]["temperatura"]["valor"], "C"
    )  # Temperatura do coletor 302
    MaterialStreams["GásRico302"].Pressure.SetValue(
        inputs["simulacao"]["coletores"]["SG-302"]["pressao"]["valor"], "kg/cm2_g"
    )  # Pressão do coletor 302
    MaterialStreams["GásRico302"].ComponentMolarFractionValue = (
        np.array(
            inputs["simulacao"]["coletores"]["SG-302"]["composicao"] + [0, 0, 0]
        )
        / 100
    )  # Fração molar coletor 302
    
    # Gás Combustível e Queima
    Operations.Item("VarModelo").Imports.Item(5).CellValue = inputs["simulacao"][
        "queima"
    ][
        "valor"
    ] /100 # Fraçao em energia queimada por energia de gás rico pré-coletor
    
      # URGN
    MaterialStreams_URGN["1"].Pressure.SetValue(
        inputs["simulacao"]["URGN"]["vaso1"]["pressao"]["valor"], "kg/cm2_g"
    )  # Pressão do Vaso V-01 da URGN
    MaterialStreams_URGN["7"].Temperature.SetValue(
        inputs["simulacao"]["URGN"]["vaso2"]["temperatura"]["valor"], "C"
    )  # Temperatura do Vaso V-02 da URGN
    
    # UPGN II
    MaterialStreams_UPGN["P08"].Temperature.SetValue(
        inputs["simulacao"]["UPGNII"]["P04"]["temperatura"]["valor"], "C"
    )  # Temperatura do chiller de propano P-04 da UPGN
    Operations_UPGN("Spread T-01").Imports.Item(0).CellValue = inputs["simulacao"][
        "UPGNII"
    ]["T01"]["temperaturaTopo"][
        "valor"
    ]  # Temperatura de topo da torre T-01 da UPGN
    Operations_UPCGN("EtGLP_UPCGNs").Imports.Item(3).CellValue = 0.9933 * (
        inputs["simulacao"]["UPGNII"]["GLP"]["etano"]["valor"] / 100
    )  # Etanização do GLP
   
    # URLs
    Operations_URL_II("ModoOpURLs").Imports.Item(7).CellValue = inputs["simulacao"][
        "URLI"
    ]["T01"]["recuperacaoMetanoTopo"][
        "valor"
    ]  # Recuperação de chave leve no topo da URL-I
    Operations_URL_II("ModoOpURLs").Imports.Item(10).CellValue = inputs[
        "simulacao"
    ]["URLI"]["T01"]["recuperacaoEtanoFundo"][
        "valor"
    ]  # Recuperação de chave pesado no fundo da URL-I
    Operations_URL_II("ModoOpURLs").Imports.Item(8).CellValue = inputs["simulacao"][
        "URLII"
    ]["T01"]["recuperacaoMetanoTopo"][
        "valor"
    ]  # Recuperação de chave leve no topo da URL-II
    Operations_URL_II("ModoOpURLs").Imports.Item(12).CellValue = inputs[
        "simulacao"
    ]["URLII"]["T01"]["recuperacaoEtanoFundo"][
        "valor"
    ]  # Recuperação de chave pesado no fundo da URL-II
    Operations_URL_II("ModoOpURLs").Imports.Item(9).CellValue = inputs["simulacao"][
        "URLIII"
    ]["T01"]["recuperacaoMetanoTopo"][
        "valor"
    ]  # Recuperação de chave leve no topo da URL-III
    Operations_URL_II("ModoOpURLs").Imports.Item(14).CellValue = inputs[
        "simulacao"
    ]["URLIII"]["T01"]["recuperacaoEtanoFundo"][
        "valor"
    ]  # Recuperação de chave pesado no fundo da URL-III
    
    # disp_URL_I = inputs["simulacao"]["disponibilidade"]["URL I"]["valor"]
    # disp_URL_II = inputs["simulacao"]["disponibilidade"]["URL II"]["valor"]
    # disp_URL_III = inputs["simulacao"]["disponibilidade"]["URL III"]["valor"]
    # disp_URLs = disp_URL_I + disp_URL_II + disp_URL_III
    
    # Operations_URLs("TEE-100").Splits.SetValues(
    #     [disp_URL_I / disp_URLs, disp_URL_III / disp_URLs, disp_URL_II / disp_URLs],
    #     "",
    # )
    
    # Sangrias
    MaterialStreams["Sangria URLs I-II"].StdLiqVolFlow.SetValue(
        inputs["simulacao"]["sangria"]["URLI/II"]["valor"], "m3/d"
    )  # Sangria URL I e II
    MaterialStreams["Sangria URL III"].StdLiqVolFlow.SetValue(
        inputs["simulacao"]["sangria"]["URLIII"]["valor"], "m3/d"
    )  # Sangria URL I e II
    
    # UPCGNs
    Operations_UPCGN("EtGLP_UPCGNs").Imports.Item(0).CellValue = 0.9933 * (
        inputs["simulacao"]["UPCGNII"]["GLP"]["etano"]["valor"] / 100
    )  # Etanização do GLP da UPCGNII
    MaterialStreams_UPCGN_II["3"].Pressure.SetValue(
        inputs["simulacao"]["UPCGNII"]["vaso1"]["pressao"]["valor"], "kg/cm2_g"
    )  # Pressão do vaso V-01 da UPCGNII
    Operations_UPCGNIV("RecDesbuta").Imports.Item(0).CellValue = inputs[
        "simulacao"
    ]["UPCGNII"]["T02-Desbutanizadora"]["recuperacaoC4GLP"][
        "valor"
    ]  # Recuperação de chave leve no topo da UPCGNII
    Operations_UPCGNIV("RecDesbuta").Imports.Item(1).CellValue = inputs[
        "simulacao"
    ]["UPCGNII"]["T02-Desbutanizadora"]["recuperacaoC5+Fundo"][
        "valor"
    ]  # Recuperação de chave pesado no fundo da UPCGNII
    Operations_UPCGN("EtGLP_UPCGNs").Imports.Item(1).CellValue = 0.9933 * (
        inputs["simulacao"]["UPCGNIII"]["GLP"]["etano"]["valor"] / 100
    )  # Etanização do GLP da UPCGNIII
    MaterialStreams_UPCGN_III["3"].Pressure.SetValue(
        inputs["simulacao"]["UPCGNIII"]["vaso1"]["pressao"]["valor"], "kg/cm2_g"
    )  # Pressão do vaso V-01 da UPCGNIII
    Operations_UPCGNIV("RecDesbuta").Imports.Item(2).CellValue = inputs[
        "simulacao"
    ]["UPCGNIII"]["T02-Desbutanizadora"]["recuperacaoC4GLP"][
        "valor"
    ]  # Recuperação de chave leve no topo da UPCGNIII
    Operations_UPCGNIV("RecDesbuta").Imports.Item(3).CellValue = inputs[
        "simulacao"
    ]["UPCGNIII"]["T02-Desbutanizadora"]["recuperacaoC5+Fundo"][
        "valor"
    ]  # Recuperação de chave pesado no fundo da UPCGNIII
    Operations_UPCGN("EtGLP_UPCGNs").Imports.Item(2).CellValue = 0.9933 * (
        inputs["simulacao"]["UPCGNIV"]["GLP"]["etano"]["valor"] / 100
    )  # Etanização do GLP da UPCGNIV
    MaterialStreams["Carga da U-301"].Pressure.SetValue(
        inputs["simulacao"]["UPCGNIV"]["vaso1"]["pressao"]["valor"], "kg/cm2_g"
    )  # Pressão do vaso V-01 da UPCGNIV
    Operations_UPCGNIV("RecDesbuta").Imports.Item(4).CellValue = inputs[
        "simulacao"
    ]["UPCGNIV"]["T02-Desbutanizadora"]["recuperacaoC4GLP"][
        "valor"
    ]  # Recuperação de chave leve no topo da UPCGNIV
    Operations_UPCGNIV("RecDesbuta").Imports.Item(5).CellValue = inputs[
        "simulacao"
    ]["UPCGNIV"]["T02-Desbutanizadora"]["recuperacaoC5+Fundo"][
        "valor"
    ]  # Recuperação de chave pesado no fundo da UPCGNIV
    
    # Manipuladas
    MaterialStreams["MIX UTGCAB"].MolarFlow.SetValue(cargas["Q_MIX_UTGCAB"], "m3/d_(gas)")              # M1
    MaterialStreams["BypassU211"].MolarFlow.SetValue(cargas["Q_BypassU211"], "m3/d_(gas)")              # M2
    MaterialStreams["295toUPGN"].MolarFlow.SetValue(cargas["Q_295toUPGN"], "m3/d_(gas)")                # M3
    MaterialStreams["GASDUC-II"].MolarFlow.SetValue(cargas["Q_GASDUC_II"], "m3/d_(gas)")                # M4
    MaterialStreams["Carga URGN"].MolarFlow.SetValue(cargas["Q_Carga_URGN"], "m3/d_(gas)")              # M5
    MaterialStreams["299toUPGN"].MolarFlow.SetValue(cargas["Q_299toUPGN"], "m3/d_(gas)")                # M6
    MaterialStreams["299to295"].MolarFlow.SetValue(cargas["Q_299to295"], "m3/d_(gas)")                  # M7
    MaterialStreams["URCO2toU211"].MolarFlow.SetValue(cargas["Q_URCO2toU211"], "m3/d_(gas)")            # M8
    Tee_GR_UPGNII = (inputs["simulacao"]["gasResidualParaReprocessamento"]["TEE_GR_UPGNII"]["valor"] / 100)
    Tee_GR_URGN = (inputs["simulacao"]["gasResidualParaReprocessamento"]["TEE_GR_URGN"]["valor"] / 100)
    Operations.Item("TEE_GR_UPGN").Splits.SetValues([Tee_GR_UPGNII, 1 - Tee_GR_UPGNII], "")             # M9
    Operations.Item("TEE_GR_URGN").Splits.SetValues([1 - Tee_GR_URGN, Tee_GR_URGN], "")                 # M10
    MaterialStreams_URLs["Carga URL I"].MolarFlow.SetValue(cargas["Q_Carga_URL_I"], "m3/d_(gas)")       # M11
    MaterialStreams_URLs["Carga URL III"].MolarFlow.SetValue(cargas["Q_Carga_URL_III"], "m3/d_(gas)")   # M12
    
    #!# EXECUÇÃO DA SIMULAÇÃO
    Solver.CanSolve = True

    # |# Verificação da convergencia
    # TODO
    status = Solver.CanSolve
    assert status == True, "Problemas de convergencia da simulação"

    # |# SAÍDAS
    # GASDUC
    GASDUC_Fv = MaterialStreams["GDC_20C1atm"].ActualVolumeFlow.GetValue("m3/d")
    GASDUC_Z = MaterialStreams["GDC_20C1atm"].ComponentMolarFractionValue
    GASDUC_Vm = MaterialStreams["GDC_20C1atm"].MolarVolume.GetValue("m3/kgmole")
    GASDUC_rho = MaterialStreams["GDC_20C1atm"].StdLiqMassDensity.GetValue("kg/m3")
    assert (
        GASDUC_Fv >= 0
        and all(i >= 0 for i in GASDUC_Z)
        and all(i <= 1 for i in GASDUC_Z)
        and GASDUC_Vm >= 0
    ), 'Erro na leitura das propriedades da corrente "GDC_20C1atm"'
    # GV
    GV_Fv = MaterialStreams["GV_20C1atm"].ActualVolumeFlow.GetValue("m3/d")
    GV_Z = MaterialStreams["GV_20C1atm"].ComponentMolarFractionValue
    GV_Vm = MaterialStreams["GV_20C1atm"].MolarVolume.GetValue("m3/kgmole")
    GV_rho = MaterialStreams["GV_20C1atm"].StdLiqMassDensity.GetValue("kg/m3")
    GV_POA = MaterialStreams["POA"].GetCorrelationValue("Water Dew Point", "Gas").GetValue("C")
    GV_PCS3 = (MaterialStreams["GV_20C1atm"].GetCorrelationValue("HHV Vol. Basis", "Gas").GetValue("MJ/m3"))
    GV_POH = MaterialStreams["POH"].GetCorrelationValue("HC Dew Point", "Gas").GetValue("C")
    GV_Wobbe = MaterialStreams["GV_20C1atm"].GetCorrelationValue("Wobbe Index", "Gas").GetValue("MJ/m3")
    GV_No_metano = (1.445* ((137.78 * GV_Z[0])
                            + (29.948 * GV_Z[1])
                            + (-18.193 * GV_Z[2])
                            + (-167.062 * (GV_Z[3] + GV_Z[4]))
                            + (181.233 * GV_Z[13])
                            + (26.994 * GV_Z[12])
                            ) - 103.42
                    )
    GV_from_URGN_part1 = MaterialStreams["Gás URGN"].MolarFlow.GetValue("m3/d_(gas)")
    GV_from_URGN_part2 = MaterialStreams["GR URGN to Venda"].MolarFlow.GetValue("m3/d_(gas)")
    GV_from_UPGN_part1 = MaterialStreams["Gás UPGN"].MolarFlow.GetValue("m3/d_(gas)")
    GV_from_UPGN_part2 = MaterialStreams["GR UPGN-II to Venda"].MolarFlow.GetValue("m3/d_(gas)")
    GV_from_URLs = MaterialStreams["Gás URLs"].MolarFlow.GetValue("m3/d_(gas)")
    assert (
        GV_Fv >= 0
        and GV_from_URGN_part1 >= 0
        and GV_from_URGN_part2 >= 0
        and GV_from_UPGN_part1 >= 0
        and GV_from_UPGN_part2 >= 0
        and GV_from_URLs >= 0
        and all(i >= 0 for i in GV_Z)
        and all(i <= 1 for i in GV_Z)
        and GV_Vm >= 0
        and GV_rho >= 0
        and GV_PCS3 >= 0
        and GV_Wobbe >= 0
    ), 'Erro na leitura das propriedades da corrente "GV_20C1atm ou POA ou POH"'
    GV_from_URGN = GV_from_URGN_part1 + GV_from_URGN_part2
    GV_from_UPGN = GV_from_UPGN_part1 + GV_from_UPGN_part2
    # LGN
    LGN_Fv = MaterialStreams["LGN"].StdLiqVolFlow.GetValue("m3/d")
    LGN_Z = MaterialStreams["LGN"].ComponentMolarFractionValue
    LGN_Vm = MaterialStreams["LGN"].MolarVolume.GetValue("m3/kgmole")
    LGN_rho = MaterialStreams["LGN"].StdLiqMassDensity.GetValue("kg/m3")
    LGN_Fw = MaterialStreams["LGN"].MassFlow.GetValue("tonne/d")
    LGN_MW = MaterialStreams["LGN"].MolecularWeightValue
    assert (
        LGN_Fv >= 0
        and all(i >= 0 for i in LGN_Z)
        and all(i <= 1 for i in LGN_Z)
        and LGN_Vm >= 0
        and LGN_rho >= 0
    ), 'Erro na leitura das propriedades da corrente "LGN"'
    # GLP
    GLP_Fv = MaterialStreams["GLP"].StdLiqVolFlow.GetValue("m3/d")
    GLP_Z = MaterialStreams["GLP"].ComponentMolarFractionValue
    GLP_Zv = MaterialStreams["GLP"].ComponentVolumeFractionValue
    GLP_Vm = MaterialStreams["GLP"].MolarVolume.GetValue("m3/kgmole")
    GLP_rho = MaterialStreams["GLP"].StdLiqMassDensity.GetValue("kg/m3")
    GLP_from_UPGN = MaterialStreams["GLP UPGN"].StdLiqVolFlow.GetValue("m3/d")
    GLP_from_UPCGNs = MaterialStreams["GLP UPCGNs"].StdLiqVolFlow.GetValue("m3/d")
    assert (
        GLP_Fv >= 0
        and GLP_from_UPGN >= 0
        and GLP_from_UPCGNs >= 0
        and all(i >= 0 for i in GLP_Z)
        and all(i <= 1 for i in GLP_Z)
        and all(i >= 0 for i in GLP_Zv)
        and all(i <= 1 for i in GLP_Zv)
        and GLP_Vm >= 0
        and GLP_rho > 0
    ), 'Erro na leitura das propriedades da corrente "GLP"'
    # C5+
    C5p_Fv = MaterialStreams["C5+"].StdLiqVolFlow.GetValue("m3/d")
    C5p_Z = MaterialStreams["C5+"].ComponentMolarFractionValue
    C5p_Vm = MaterialStreams["C5+"].MolarVolume.GetValue("m3/kgmole")
    C5p_rho = MaterialStreams["C5+"].StdLiqMassDensity.GetValue("kg/m3")
    C5p_PVR = MaterialStreams["C5+"].ColdProperty.ReidVapourPressure.GetValue("kPa")
    C5p_from_UPGN = MaterialStreams["C5+ UPGN"].StdLiqVolFlow.GetValue("m3/d")
    C5p_from_UPCGNs = MaterialStreams["C5+ UPCGNs"].StdLiqVolFlow.GetValue("m3/d")
    assert (
        C5p_Fv >= 0
        and all(i >= 0 for i in C5p_Z)
        and all(i <= 1 for i in C5p_Z)
        and C5p_Vm >= 0
    ), 'Erro na leitura das propriedades da corrente "C5+"'
    # Mix UTGCAB
    Mix_Fv = MaterialStreams["MIXtoScrubber2"].MolarFlow.GetValue("m3/d_(gas)")
    Mix_Z = MaterialStreams["MIXtoScrubber2"].ComponentMolarFractionValue
    assert (
        Mix_Fv >= 0 and all(i >= 0 for i in Mix_Z) and all(i <= 1 for i in Mix_Z)
    ), 'Erro na leitura das propriedades da corrente "MIXtoScrubber2"'
    # Queima
    Queima_Fv = MaterialStreams["Queima_20C1atm"].MolarFlow.GetValue("m3/d_(gas)")
    Queima_Z = MaterialStreams["Queima_20C1atm"].ComponentMolarFractionValue
    Queima_Vm = MaterialStreams["Queima_20C1atm"].MolarVolume.GetValue("m3/kgmole")
    Queima_PCS = PCSCalculadoISO(Queima_Z)
    Queima295_FV = MaterialStreams["295Queima"].MolarFlow.GetValue("m3/d_(gas)")
    Queima299_FV = MaterialStreams["299Queima"].MolarFlow.GetValue("m3/d_(gas)")
    Queima302_FV = MaterialStreams["302Queima"].MolarFlow.GetValue("m3/d_(gas)")
    assert (
        Queima_Fv >= 0
        and all(i >= 0 for i in Queima_Z)
        and all(i <= 1 for i in Queima_Z)
        and Queima295_FV >= 0
        and Queima299_FV >= 0
        and Queima302_FV >= 0
    ), 'Erro na leitura das propriedades da corrente "Queima_20C1atm"'

    # Consumo de eletricidade e Combustivel
    EE = (
        Operations.Item("VarModelo").Cell("B1").CellValue
    )  # Consumo de energia eletrica [kJ/h]
    assert EE >= 0, 'Erro na leitura das propriedades do Spreadsheet "VarModelo"'
    if inputs["simulacao"]["gasCombustivel"]["modoCalculo"] == "Definir %":
        PercentualGasCombustivel = inputs["simulacao"]["gasCombustivel"][
            "GasCombustivel"
        ]["valor"]
        GC_Fv = GV_Fv - GV_Fv * (
            1 - PercentualGasCombustivel / (100 + PercentualGasCombustivel)
        )
    elif (
        inputs["simulacao"]["gasCombustivel"]["modoCalculo"] == "Estimar por Simulacao"
    ):
        GC_Fv = (
            Operations.Item("VarModelo").Cell("B2").CellValue
        )  # Gás utilizado como combustivel [m3/d]
        assert (
            EE >= 0 and GC_Fv >= 0
        ), 'Erro na leitura das propriedades da spreadsheet "VarModelo"'

    # |# Restriçoes
    # Restrições de carga das unidades
    carga_GASDUC = cargas["Q_GASDUC_II"]
    carga_MIX_UTGCAB = cargas["Q_MIX_UTGCAB"]
    carga_UPGN_II = (
        cargas["Q_BypassU211"]
        + cargas["Q_URCO2toU211"]
        + cargas["Q_299toUPGN"]
        + cargas["Q_295toUPGN"]
    )
    carga_URGN = cargas["Q_Carga_URGN"]
    carga_URL_I = MaterialStreams_URLs["Carga URL I"].MolarFlow.GetValue("m3/d_(gas)")
    assert (
        carga_URL_I >= 0
    ), 'Erro na leitura das propriedades da corrente "Carga URL I"'
    carga_URL_II = MaterialStreams_URLs["Carga URL II"].MolarFlow.GetValue("m3/d_(gas)")
    assert (
        carga_URL_II >= 0
    ), 'Erro na leitura das propriedades da corrente "Carga URL II"'
    carga_URL_III = MaterialStreams_URLs["Carga URL III"].MolarFlow.GetValue(
        "m3/d_(gas)"
    )
    assert (
        carga_URL_III >= 0
    ), 'Erro na leitura das propriedades da corrente "Carga URL III"'
    carga_URLs = carga_URL_I + carga_URL_II + carga_URL_III
    carga_U298 = MaterialStreams["Carga da U-298"].StdLiqVolFlow.GetValue("m3/d")
    assert (
        carga_U298 >= 0
    ), 'Erro na leitura das propriedades da corrente "Carga da U-298"'
    carga_U300 = MaterialStreams["Carga da U-300"].StdLiqVolFlow.GetValue("m3/d")
    assert (
        carga_U300 >= 0
    ), 'Erro na leitura das propriedades da corrente "Carga da U-300"'
    carga_U301 = MaterialStreams["Carga da U-301"].StdLiqVolFlow.GetValue("m3/d")
    assert (
        carga_U301 >= 0
    ), 'Erro na leitura das propriedades da corrente "Carga da U-301"'
    sangria_U298 = MaterialStreams["Sangria para U-298"].StdLiqVolFlow.GetValue("m3/d")
    assert (
        sangria_U298 >= 0
    ), 'Erro na leitura das propriedades da corrente "Sangria para U-298"'
    sangria_U300 = MaterialStreams["Sangria para U-300"].StdLiqVolFlow.GetValue("m3/d")
    assert (
        sangria_U300 >= 0
    ), 'Erro na leitura das propriedades da corrente "Sangria para U-300"'
    sangria_U301 = MaterialStreams["Sangria para U-301"].StdLiqVolFlow.GetValue("m3/d")
    assert (
        sangria_U301 >= 0
    ), 'Erro na leitura das propriedades da corrente "Sangria para U-301"'
    LGN_U298 = (
        MaterialStreams["LGN URGN p/ U298"]
        .GetCorrelationValue("Act. Liq. Flow", "Standard")
        .GetValue("m3/d")
    )
    assert (
        LGN_U298 >= 0
    ), 'Erro na leitura das propriedades da corrente "LGN URGN p/ U298"'
    carga_UPCGNs = (
        carga_U298
        + carga_U300
        + carga_U301
        + sangria_U298
        + sangria_U300
        + sangria_U301
        + LGN_U298
    )
    # Restrições do GV
    GV_C1 = GV_Z[0] * 100
    GV_C2 = GV_Z[1] * 100
    GV_C3 = GV_Z[2] * 100
    GV_C4p = sum(GV_Z[3:12]) * 100
    GV_CO2 = GV_Z[13] * 100
    GV_inertes = (GV_Z[12] + GV_Z[13]) * 100
    GV_PCS3 = GV_PCS3
    GV_POA = GV_POA
    GV_POH = GV_POH
    GV_Wobbe = GV_Wobbe    
    GV_No_metano = GV_No_metano
        
    # Restrições do GLP
    GLP_C2 = GLP_Zv[1] * 100
    GLP_C5p = sum(GLP_Zv[5:]) * 100
    GLP_rho = GLP_rho
    Restricoes = {
        "carga_GASDUC": carga_GASDUC,
        "carga_MIX_UTGCAB": carga_MIX_UTGCAB,
        "carga_UPGN_II": carga_UPGN_II,
        "carga_URGN": carga_URGN,
        "carga_URLs": carga_URLs,
        "carga_UPCGNs": carga_UPCGNs,
        "GV_C1": GV_C1,
        "GV_C2": GV_C2,
        "GV_C3": GV_C3,
        "GV_C4p": GV_C4p,
        "GV_CO2": GV_CO2,
        "GV_inertes": GV_inertes,
        "GV_PCS": GV_PCS3,
        "GV_POA": GV_POA,
        "GV_POH": GV_POH,
        "GV_Wobbe": GV_Wobbe,
        "GV_No_metano": GV_No_metano,
        "GLP_C2": GLP_C2,
        "GLP_C5p": GLP_C5p,
        "GLP_rho": GLP_rho,
    }

    # Verificando se todas as restrições são satisfeitas e calculando as penalizações
    penal_rest = confere_restricoes(config, inputs, Restricoes)

    # |# Parte de avaliação economica

    # Taxa de cambio [R$/USD]
    taxa_cambio = inputs["otimizacao"]["preco"]["cambio"]["valor"]

    # Receita GASDUC
    GASDUC_PCS = PCSCalculadoISO(GASDUC_Z)  # MMbtu/kmol
    GASDUC_PCS2 = GASDUC_PCS / GASDUC_Vm  # MMbtu/m3
    GASDUC_energia = GASDUC_Fv * GASDUC_PCS2  # MMbtu/d
    GASDUC_preco = (
        taxa_cambio
        * inputs["otimizacao"]["preco"]["custoOportunidade"]["GASDUC"]["valor"]
    )  # [R$/MMbtu]
    GASDUC_receita = GASDUC_preco * GASDUC_energia  # [R$/d]
    # Receita GV
    GV_PCS = PCSCalculadoISO(GV_Z)  # MMbtu/kmol
    GV_PCS2 = GV_PCS / GV_Vm  # MMbtu/m3
    GV_Fv_net = (
        GV_Fv - GC_Fv
    )  # Gas de Venta liquido, descontado o utilizado como combustivel [m3/d]
    GV_energia = GV_Fv_net * GV_PCS2  # MMbtu/d
    GV_preco = taxa_cambio*inputs["otimizacao"]["preco"]["custoOportunidade"]["GV"][
        "valor"
    ]  # [R$/MMbtu]
    GV_receita = GV_preco * GV_energia  # [R$/d]
    # Receita LGN
    LGN_PCS = PCSCalculadoISO(LGN_Z)  # MMbtu/kmol
    LGN_PCS2 = LGN_PCS / LGN_Vm  # MMbtu/m3
    LGN_energia = LGN_Fv * LGN_PCS2  # MMbtu/d
    LGN_preco = taxa_cambio*inputs["otimizacao"]["preco"]["custoOportunidade"]["LGN"][
        "valor"
    ]  # [R$/MMbtu]
    LGN_receita = LGN_preco * LGN_energia  # [R$/d]
    # Receita GLP
    GLP_PCS = PCSCalculadoISO(GLP_Z)  # MMbtu/kmol
    GLP_PCS2 = GLP_PCS / GLP_Vm  # MMbtu/m3
    GLP_energia = GLP_Fv * GLP_PCS2  # MMbtu/d
    GLP_preco = taxa_cambio*inputs["otimizacao"]["preco"]["custoOportunidade"]["GLP"][
        "valor"
    ]  # [R$/MMbtu]
    GLP_receita = GLP_preco * GLP_energia  # [R$/d]
    # Receita C5+
    C5p_PCS = PCSCalculadoISO(C5p_Z)  # MMbtu/kmol
    C5p_PCS2 = C5p_PCS / C5p_Vm  # MMbtu/m3
    C5p_energia = C5p_Fv * C5p_PCS2  # MMbtu/d
    C5p_preco = taxa_cambio*inputs["otimizacao"]["preco"]["custoOportunidade"]["C5+"][
        "valor"
    ]  # [R$/MMbtu]
    C5p_receita = C5p_preco * C5p_energia  # [R$/d]
    # Receita_total
    Receita_total = (
        GASDUC_receita + GV_receita + LGN_receita + GLP_receita + C5p_receita
    )  # [R$/d]
    
    print('Receita ponto Base - Rigorosa')
    print(Receita_total)
    Receitas = {"gasDeVenda": {"economico": {"valor": GV_receita, "unidade": "R$/d"},
                               "producao": {"valor": GV_Fv_net, "unidade": "m3/d_(gas)"}},
                "GASDUC II": {"economico": {"valor": GASDUC_receita, "unidade": "R$/d"},
                                           "producao": {"valor": GASDUC_Fv, "unidade": "m3/d_(gas)"}},
                "LGN": {"economico": {"valor": LGN_receita, "unidade": "R$/d"},
                        "producao": {"valor": LGN_Fv, "unidade": "m3/d"}},
                "GLP": {"economico": {"valor": GLP_receita, "unidade": "R$/d"},
                        "producao": {"valor": GLP_Fv, "unidade": "m3/d"}},
                "C5+": {"economico": {"valor": C5p_receita, "unidade": "R$/d"},
                        "producao": {"valor": C5p_Fv, "unidade": "m3/d"}},
                "Total": {"economico": {"valor": Receita_total, "unidade": "R$/d"}},
                }

    # Custos eletricidade e combustivel
    EE2 = EE * 2.77786e-7  # Consumo de energia eletrica [MW]
    Custo_EE = (
        inputs["otimizacao"]["preco"]["custoEnergiaEletrica"]["custoUnitario"]["valor"]
        * EE2
        * 24
    )  # Custo energia eletrica [R$/d]
    Custo_GC = (
        inputs["otimizacao"]["preco"]["custoGasCombustivel"]["custoUnitario"]["valor"]
        * GC_Fv
        * 1e-3
    )  # Custo gas combustivel [R$/d]
    # Custos de produtos químicos
    Q_302ToURCO2 = MaterialStreams["302ToURCO2"].MolarFlow.GetValue("m3/d_(gas)")
    assert (
        Q_302ToURCO2 >= 0
    ), 'Erro na leitura das propriedades da corrente "302ToURCO2"'
    custo_amina_ativada_URCO2 = (
        inputs["otimizacao"]["preco"]["custosProdutosQuimicos"]["aminaAtivadaURCO2"][
            "valor"
        ]
        * 1.5
        * 1000
        * 2
        / 30
        * taxa_cambio
        * Q_302ToURCO2
        / 9e6
    )  # [R$/d]
    custo_propano_URGN = (
        716
        * inputs["otimizacao"]["preco"]["custoOportunidade"]["GLP"]["valor"]
        * GLP_PCS2
        * taxa_cambio
        / 1.5
        / 30
        * cargas["Q_Carga_URGN"]
        / 3e6
        * 0.6
    )  # [R$/d]
    custo_propano_UPGNII = (
        716
        * inputs["otimizacao"]["preco"]["custoOportunidade"]["GLP"]["valor"]
        * GLP_PCS2
        * taxa_cambio
        / 1.5
        / 30
        * carga_UPGN_II
        / 5.8e6
        * 0.4
    )  # [R$/d]
    custo_peneira_molecular_URLs = (
        inputs["otimizacao"]["preco"]["custosProdutosQuimicos"]["peneiraMolecularURLs"][
            "valor"
        ]
        * 1000
        * 3
        / (3 * 360)
        * (carga_URLs / (5.4e6 * 3))
    )  # [R$/d]
    custo_carvao_activado_URCO2 = (
        (343 * 24 * 0.453592)
        * inputs["otimizacao"]["preco"]["custosProdutosQuimicos"]["carvaoAtivadoURCO2"][
            "valor"
        ]
        / (1.5 * 360)
        * 2
        * Q_302ToURCO2
        / 9e6
    )  # [R$/d]
    custo_produtos_quimicos = (
        custo_amina_ativada_URCO2
        + custo_carvao_activado_URCO2
        + custo_propano_URGN
        + custo_propano_UPGNII
        + custo_peneira_molecular_URLs
    )
    # Custo total
    Custo_total = Custo_EE + Custo_GC  # Custo total [R$/d]
    print('Custo ponto Base - Rigorosa')
    print(Custo_total)
    Custos = {"energiaEletrica": {"economico": {"valor": Custo_EE, "unidade": "R$/d"},
                                  "potenciaSimulada": {"valor": EE2, "unidade": "MW"}},
              "gasCombustivel": {"economico": {"valor": Custo_GC, "unidade": "R$/d"},
                                 "demandaSimulada": {"valor": GC_Fv, "unidade": "m3/d"}},
              "produtosQuimicos": {"economico": {"valor": custo_produtos_quimicos, "unidade": "R$/d"}},
              "Total": {"economico": {"valor": Custo_total, "unidade": "R$/d"}}
            }

    # Margem
    margen = Receita_total - Custo_total  # [R$/d]
    Margem = {"margem": {"economico": {"valor": margen, "unidade": "R$/d"} }}

    # Vazão de gás que sai dos coletores
    COL295_QGAS = MaterialStreams["295Gout"].MolarFlow.GetValue(
        "m3/d_(gas)"
    )  # Vazão de gás do Coletor 295
    COL299_QGAS = MaterialStreams["299Gout"].MolarFlow.GetValue(
        "m3/d_(gas)"
    )  # Vazão de gás do Coletor 299
    COL302_QGAS = MaterialStreams["302Gout"].MolarFlow.GetValue(
        "m3/d_(gas)"
    )  # Vazão de gás do Coletor 302
    Gas_Reciclo_UPCGNs = MaterialStreams["Gás de Reciclo UPCGNs"].MolarFlow.GetValue(
        "m3/d_(gas)"
    )  # Vazão de gás de reciclo das UPCGs
    Col_Qgas = {
        "COL295_QGAS": COL295_QGAS,
        "COL299_QGAS": COL299_QGAS,
        "COL302_QGAS": COL302_QGAS,
        "Gas_Reciclo_UPCGNs": Gas_Reciclo_UPCGNs,
    }
    Coletores = {"Col_Qgas": Col_Qgas}

    # Vazaõ de condensado que sai dos coletores
    COL295_QCOND = (
        MaterialStreams["295CondOut"]
        .GetCorrelationValue("Act. Liq. Flow", "Standard")
        .GetValue("m3/d")
    )
    COL299_QCOND = (
        MaterialStreams["299CondOut"]
        .GetCorrelationValue("Act. Liq. Flow", "Standard")
        .GetValue("m3/d")
    )
    COL302_QCOND = (
        MaterialStreams["302CondOut"]
        .GetCorrelationValue("Act. Liq. Flow", "Standard")
        .GetValue("m3/d")
    )

    # Vazão de sangria total para as UPCNGs
    Q_sangria_total = MaterialStreams["LGNSangria"].StdLiqVolFlow.GetValue("m3/d")
    # Vazão de liuido que sai do vaso de Carga da UPGN II
    Q_vasoDeCargaUPGN = MaterialStreams["Cond MIX V-21101"].ActualVolumeFlow.GetValue("m3/d")

    # Dicionario para cada produto
    gasDeVenda = {
        "composicao": list(np.array(GV_Z) * 100),
        "vazao": {"valor": GV_Fv_net, "unidade": "m3/d_(gas)"},
        "propriedades": {
            "POA": {"valor": GV_POA, "unidade": "C"},
            "POH": {"valor": GV_POH, "unidade": "C"},
            "PCS molar": {"valor": GV_PCS, "unidade": "MMbtu/kgmole"},
            "Ind. Wobbe": {"valor": GV_Wobbe, "unidade": "MJ/m3"},
            "Numero de metano": {"valor": GV_No_metano, "unidade": ""},
            "Volume molar": {"valor": GV_Vm, "unidade": "m3/kgmole"},
            "PCS volumetrico": {"valor": GV_PCS2 * 1054, "unidade": "MJ/m3"},
            "Massa especifica": {"valor": GV_rho, "unidade": "kg/m3"},
        },
        "unidadeOrigem": {
            "URGN": {"valor": GV_from_URGN, "unidade": "m3/d_(gas)"},
            "UPGN II": {"valor": GV_from_UPGN, "unidade": "m3/d_(gas)"},
            "URLs": {"valor": GV_from_URLs, "unidade": "m3/d_(gas)"},
        },
    }
    LGN = {
        "composicao": list(np.array(LGN_Z) * 100),
        "vazao": {"valor": LGN_Fv, "unidade": "m3/d"},
        "propriedades": {
            "Volume molar": {"valor": LGN_Vm, "unidade": "m3/kgmole"},
            "PCS molar": {"valor": LGN_PCS, "unidade": "MMbtu/kgmole"},
            "Massa especifica": {"valor": LGN_rho, "unidade": "kg/m3"},
        },
    }
    GLP = {
        "composicao": list(np.array(GLP_Z) * 100),
        "vazao": {"valor": GLP_Fv, "unidade": "m3/d"},
        "propriedades": {
            "Volume molar": {"valor": GLP_Vm, "unidade": "m3/kgmole"},
            "PCS molar": {"valor": GLP_PCS, "unidade": "MMbtu/kgmole"},
            "Massa especifica": {"valor": GLP_rho, "unidade": "kg/m3"},
        },
        "unidadeOrigem": {
            "UPGN II": {"valor": GLP_from_UPGN, "unidade": "m3/d"},
            "UPCGNs": {"valor": GLP_from_UPCGNs, "unidade": "m3/d"},
        },
    }
    C5p = {
        "composicao": list(np.array(C5p_Z) * 100),
        "vazao": {"valor": C5p_Fv, "unidade": "m3/d"},
        "propriedades": {
            "Volume molar": {"valor": C5p_Vm, "unidade": "m3/kgmole"},
            "PCS molar": {"valor": C5p_PCS, "unidade": "MMbtu/kgmole"},
            "Massa especifica": {"valor": C5p_rho, "unidade": "kg/m3"},
            "PVR": {"valor": C5p_PVR, "unidade": "kPa"},
        },
        "unidadeOrigem": {
            "UPGN II": {"valor": C5p_from_UPGN, "unidade": "m3/d"},
            "UPCGNs": {"valor": C5p_from_UPCGNs, "unidade": "m3/d"},
        },
    }
    GASDUC_II = {
        "composicao": list(np.array(GASDUC_Z) * 100),
        "vazao": {"valor": GASDUC_Fv, "unidade": "m3/d_(gas)"},
        "propriedades": {
            "Volume molar": {"valor": GASDUC_Vm, "unidade": "m3/kgmole"},
            "PCS molar": {"valor": GASDUC_PCS, "unidade": "MMbtu/kgmole"},
            "Massa especifica": {"valor": GASDUC_rho, "unidade": "kg/m3"},
        },
    }
    MIX_UTGCAB = {
        "composicao": list(np.array(Mix_Z) * 100),
        "vazao": {"valor": Mix_Fv, "unidade": "m3/d_(gas)"},
    }
    queima = {
        "composicao": list(np.array(Queima_Z) * 100),
        "vazao": {"valor": Queima_Fv, "unidade": "m3/d_(gas)"},
        "propriedades": {
            "Volume molar": {"valor": Queima_Vm, "unidade": "m3/kgmole"},
            "PCS molar": {"valor": Queima_PCS, "unidade": "MMbtu/kgmole"},
        },
        "queima_por_coletor": {
            "queima295": Queima295_FV,
            "queima299": Queima299_FV,
            "queima302": Queima302_FV,
        },
    }
    gasCombustivel = {"vazao": {"valor": GC_Fv, "unidade": "m3/d"}}
    # Dicionario de produção
    producao = {
        "gasDeVenda": gasDeVenda,
        "LGN": LGN,
        "GLP": GLP,
        "C5+": C5p,
        "GASDUC II": GASDUC_II,
        "MIX UTGCAB": MIX_UTGCAB,
        "queima": queima,
        "gasCombustivel": gasCombustivel,
    }
    # Cargas das Unidades
    cargasDasUnidades = {
        "URGN": {"valor": carga_URGN, "unidade": "m3/d_(gas)"},
        "UPGN II": {"valor": carga_UPGN_II, "unidade": "m3/d_(gas)"},
        "URL I": {"valor": carga_URL_I, "unidade": "m3/d_(gas)"},
        "URL II": {"valor": carga_URL_II, "unidade": "m3/d_(gas)"},
        "URL III": {"valor": carga_URL_III, "unidade": "m3/d_(gas)"},
        "UPCGNs": {
            "condensadoSG-295": {"valor": COL295_QCOND, "unidade": "m3/d"},
            "condensadoSG-299": {"valor": COL299_QCOND, "unidade": "m3/d"},
            "condensadoSG-302": {"valor": COL302_QCOND, "unidade": "m3/d"},
            "LGNdaURGN": {"valor": LGN_U298, "unidade": "m3/d"},
            "sangriaDaURLs": {"valor": Q_sangria_total, "unidade": "m3/d"},
            "vasoDeCargaUPGN-II": {"valor": Q_vasoDeCargaUPGN, "unidade": "m3/d"},
        },
    }
    
    # Maipuladas
    manipuladas = {"299to295": {"valor": cargas['Q_299to295'], "unidade": "m3/d_(gas)"},
                   "GASDUC-II": {"valor": cargas['Q_GASDUC_II'], "unidade": "m3/d_(gas)"},
                   "MIX UTGCAB": {"valor": cargas['Q_MIX_UTGCAB'], "unidade": "m3/d_(gas)"},
                   "295toUPGN": {"valor": cargas['Q_295toUPGN'], "unidade": "m3/d_(gas)"},
                   "299toUPGN": {"valor": cargas['Q_299toUPGN'], "unidade": "m3/d_(gas)"},
                   "URCO2toU211": {"valor": cargas['Q_URCO2toU211'], "unidade": "m3/d_(gas)"},
                   "BypassU211": {"valor": cargas['Q_BypassU211'], "unidade": "m3/d_(gas)"},
                   "Carga URGN": {"valor": cargas['Q_Carga_URGN'], "unidade": "m3/d_(gas)"},
                   "Carga URL I": {"valor": cargas['Q_Carga_URL_I'], "unidade": "m3/d_(gas)"},
                   "Carga URL III": {"valor": cargas['Q_Carga_URL_III'], "unidade": "m3/d_(gas)"},
                   "TEE_GR_UPGN": {"valor": Tee_GR_UPGNII, "unidade": "%"},
                   "TEE_GR_URGN": {"valor": Tee_GR_URGN, "unidade": "%"},
                   "Carga URL II": {"valor": carga_URL_II, "unidade": "m3/d_(gas)"},
                   "Carga UPGN II": {"valor": carga_UPGN_II, "unidade": "m3/d_(gas)"},
                   "295toURLs": {"valor": MaterialStreams["295toURLs"].MolarFlow.GetValue("m3/d_(gas)"), "unidade": "m3/d_(gas)"},
                   "299toURLs": {"valor": MaterialStreams["299toURLs"].MolarFlow.GetValue("m3/d_(gas)"), "unidade": "m3/d_(gas)"},
                   "302toURLs": {"valor": MaterialStreams["URCO2ToURL"].MolarFlow.GetValue("m3/d_(gas)"), "unidade": "m3/d_(gas)"},
                   }
    
    # Contrato Braskem
    etano_LGN = LGN_Z[1] * (LGN_Fw * 1000 / LGN_MW) * 30 / 1000
    propano_LGN = LGN_Z[2] * (LGN_Fw * 1000 / LGN_MW) * 44 / 1000
    eteno_equivalente_LGN = propano_LGN/2.14 + etano_LGN/1.24
    contratoBraskem = {"etano": {"valor": etano_LGN, "unidade": "tonne/d"},
                        "propano": {"valor": propano_LGN, "unidade": "tonne/d"},
                        "etenoEquivalente": {"valor": eteno_equivalente_LGN, "unidade": "tonne/d"}
                      }
    # Dicionario com o resumo dos resultados da simulação
    resultados_simulacao = {
                            "convergiu": status,
                            "atendeRestricoes": penal_rest["viavel"],
                            "producao": producao,
                            "cargasDasUnidades": cargasDasUnidades,
                            "contratoBraskem": contratoBraskem,
                            "receita": Receitas,
                            "custo": Custos,
                            "margem": Margem,
                            "alocacaoCargas": manipuladas,
                            "restricoes": Restricoes,
                            "coletores": Coletores,
                            "penal_rest": penal_rest,
                            }

    'Especificações que serão utilizadas na simulação essencial [LIZANDRO]'
    C2_URLI = simCase.Flowsheet.Flowsheets('TPL4').Flowsheets('TPL5').MaterialStreams["26"].ComponentMolarFractionValue[1]   # Fração molar de C2 no topo da T01 da URL-I
    C3_URLI = simCase.Flowsheet.Flowsheets('TPL4').Flowsheets('TPL5').MaterialStreams["26"].ComponentMolarFractionValue[2]   # Fração molar de C3 no topo da T01 da URL-II
    C1_URLI = simCase.Flowsheet.Flowsheets('TPL4').Flowsheets('TPL5').MaterialStreams["42"].ComponentMolarFractionValue[0]   # Fração molar de C1 no fundo da T01 da URL-I
    CO2_URLI = simCase.Flowsheet.Flowsheets('TPL4').Flowsheets('TPL5').MaterialStreams["26"].ComponentMolarFractionValue[13]   # Fração molar de CO2 no topo da T01 da URL-I
    
    
    C2_URLII = simCase.Flowsheet.Flowsheets('TPL4').Flowsheets('TPL6').MaterialStreams["26"].ComponentMolarFractionValue[1]   # Fração molar de C2 no topo da T01 da URL-I
    C3_URLII = simCase.Flowsheet.Flowsheets('TPL4').Flowsheets('TPL6').MaterialStreams["26"].ComponentMolarFractionValue[2]   # Fração molar de C3 no topo da T01 da URL-II
    C1_URLII = simCase.Flowsheet.Flowsheets('TPL4').Flowsheets('TPL6').MaterialStreams["42"].ComponentMolarFractionValue[0]   # Fração molar de C1 no fundo da T01 da URL-I
    CO2_URLII = simCase.Flowsheet.Flowsheets('TPL4').Flowsheets('TPL6').MaterialStreams["26"].ComponentMolarFractionValue[13]  # Fração molar de CO2 no topo da T01 da URL-I
    
    
    C2_URLIII = simCase.Flowsheet.Flowsheets('TPL4').Flowsheets('TPL10').MaterialStreams["26"].ComponentMolarFractionValue[1]   # Fração molar de C2 no topo da T01 da URL-I
    C3_URLIII= simCase.Flowsheet.Flowsheets('TPL4').Flowsheets('TPL10').MaterialStreams["26"].ComponentMolarFractionValue[2]   # Fração molar de C3 no topo da T01 da URL-II
    C1_URLIII = simCase.Flowsheet.Flowsheets('TPL4').Flowsheets('TPL10').MaterialStreams["42"].ComponentMolarFractionValue[0]   # Fração molar de C1 no fundo da T01 da URL-I
    CO2_URLIII = simCase.Flowsheet.Flowsheets('TPL4').Flowsheets('TPL10').MaterialStreams["26"].ComponentMolarFractionValue[13]  # Fração molar de CO2 no topo da T01 da URL-I
    
    
    Temp_V03 = MaterialStreams_UPGN['Fundo do V-03'].Temperature.Getvalue("C") # Para a simulação Essencial
    T_P24_UPGN = MaterialStreams_UPGN['P24'].Temperature.Getvalue("C") # Para a simulação Essencial [Entrada da torre T01]
    
    R_especs = {'C2_URLI':C2_URLI, # Dicionário com as especifciação que serão repassadas à simulação Essencial
                'C3_URLI':C3_URLI,
                'C1_URLI':C1_URLI,
                'C2_URLII':C2_URLII,
                'C3_URLII':C3_URLII,
                'C1_URLII':C1_URLII,
                'C2_URLIII':C2_URLIII,
                'C3_URLIII':C3_URLIII,
                'C1_URLIII':C1_URLIII,
                'CO2_URLI':CO2_URLI,
                'CO2_URLII':CO2_URLII,
                'CO2_URLIII':CO2_URLIII,                
                'Temp_V03':Temp_V03,   
                'T_P24_UPGN': T_P24_UPGN,
                }

    return resultados_simulacao, R_especs 