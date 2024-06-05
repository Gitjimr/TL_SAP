import numpy as np
import pandas as pd
from unidecode import unidecode
import datetime
from datetime import datetime

data_hoje = str(datetime.today().strftime('%d.%m.%Y'))

import streamlit as st

st.session_state.update(st.session_state)
for k, v in st.session_state.items():
    st.session_state[k] = v


def checagem_df(df):
    df = df.drop_duplicates()
    df = df.dropna(how='all')
    df = df.reset_index(drop=True)

    return df


import os

path = os.path.dirname(__file__)
my_path = path + '/files/'
uploaded_file0 = st.sidebar.file_uploader("Carregar Dados Chave",
                                         help="Carregar arquivo com dados necessários do SAP. Caso precise recarregá-lo, atualize a página. Este arquivo deve ser continuamente atualizado conforme novos dados sejam inseridos no SAP"
                                         )
if uploaded_file0 is not None:
    if 'SAP_CTPM' not in st.session_state:
        with st.spinner('Carregando Lista de Equipamentos...'):
            SAP_EQP_N6 = pd.read_excel(uploaded_file0, sheet_name="EQP", skiprows=0, dtype=str)
        with st.spinner('Carregando IE03 SAP...'):
            SAP_EQP = pd.read_excel(uploaded_file0, sheet_name="IE03", skiprows=0, dtype=str)
        with st.spinner('Carregando IA39 SAP...'):
            SAP_TL = pd.read_excel(uploaded_file0, sheet_name="IA39", skiprows=0, dtype=str)
        with st.spinner('Carregando IP18 SAP...'):
            SAP_ITEM = pd.read_excel(uploaded_file0, sheet_name="IP18", skiprows=0, dtype=str)
        with st.spinner('Carregando IP24 SAP...'):
            SAP_PMI = pd.read_excel(uploaded_file0, sheet_name="IP24", skiprows=0, dtype=str)
            SAP_PMI['CONCAT CENTRO_DESC'] = SAP_PMI["Planning Plant"].map(str, na_action=None) + SAP_PMI["Maintenance Plan Desc"].map(str, na_action='ignore')
            SAP_PMI['CONCAT TL_EQP'] = np.where(  # Incluído 05/06/2024
                SAP_PMI['Equipment'].notna(),  # condição: se 'Equipment' não for NaN
                SAP_PMI["Group"].map(str) + SAP_PMI["Equipment"].map(str),  # se verdadeiro: group + equipment
                SAP_PMI["Group"].map(str) + SAP_PMI["Functional Location"].map(str)  # se falso: group + functional location
            )
            SAP_PMI = pd.DataFrame(SAP_PMI)
        with st.spinner('Carregando Centros de Trabalho SAP...'):
            SAP_CTPM = pd.read_excel(uploaded_file0, sheet_name="CTPM", skiprows=0, dtype=str)
        with st.spinner('Carregando Materiais SAP...'):
            SAP_MATERIAIS = pd.read_excel(uploaded_file0, sheet_name="MATERIAIS", skiprows=0, dtype=str)
            SAP_MATERIAIS.dropna(subset='Material', inplace=True)
            SAP_MATERIAIS.reset_index(drop=True, inplace=True)

            st.session_state.SAP_EQP_N6 = SAP_EQP_N6
            st.session_state.SAP_EQP = SAP_EQP
            st.session_state.SAP_TL = SAP_TL
            st.session_state.SAP_ITEM = SAP_ITEM
            st.session_state.SAP_PMI = SAP_PMI
            st.session_state.SAP_MATERIAIS = SAP_MATERIAIS
            st.session_state.SAP_CTPM = SAP_CTPM
    else:
        SAP_EQP_N6 = st.session_state['SAP_EQP_N6']
        SAP_EQP = st.session_state['SAP_EQP']
        SAP_TL = st.session_state['SAP_TL']
        SAP_ITEM = st.session_state['SAP_ITEM']
        SAP_PMI = st.session_state['SAP_PMI']
        SAP_CTPM = st.session_state['SAP_CTPM']
        SAP_MATERIAIS = st.session_state['SAP_MATERIAIS']

#*-*-*-*-OK ACIMA

uploaded_file1 = st.file_uploader("Carregar 'CONCAT_LSMW_CAB' com cabeçalho das listas de tarefas geradas na página TL_SAP",help="Precisa estar com o número das task lists")
uploaded_file2 = st.file_uploader("Carregar 'TABELAO...' com listas de tarefas geradas na página TL")

if uploaded_file1 is not None and uploaded_file2 is not None and uploaded_file0 is not None:

    with st.spinner('Carregando Cabeçalho dos Planos...'):

        # ADD PLANO:
        
        # Nomes dos arquivos de cabecalho da LSMW a serem lidos:
        
        arquivos_cabecalho_tl = []   # Nome dos arquivos (CONCAT_LSMW_CAB)
        
        arquivos_cabecalho_tl.append(uploaded_file1)
        
        
        # Nomes dos arquivos TABELAO_SAP (COM CABECALHOS) a serem lidos:
        
        arquivos_cabecalho_planos = []  # Nome dos arquivos (CABEÇALHO DO TABELAO)
        
        arquivos_cabecalho_planos.append(uploaded_file2)
        
        
        # Carregando planilhas necessárias
        
        for i in range(len(arquivos_cabecalho_planos)):
        
            ## Ler a tabela específica da planilha
            
            ### MODELO TABELAO_SAP:
            df_cabecalhos_pm_i = pd.read_excel(arquivos_cabecalho_planos[i], sheet_name='CABECALHO PLANO', skiprows=0,dtype = str)
            
            df_cabecalhos_pm_i["CONCAT"] = df_cabecalhos_pm_i["Centro de planejamento*"].map(str, na_action=None) + df_cabecalhos_pm_i["Descrição"].map(str, na_action='ignore')    # A descrição do cabecalho é a "descrição antiga", sem revisão na qtd de caracteres
            df_cabecalhos_pm_i = df_cabecalhos_pm_i[pd.notna(df_cabecalhos_pm_i['Chave do grupo de listas de tarefas*'])].copy().reset_index(drop=True).loc[:, ~df_cabecalhos_pm_i.columns.str.contains('^Unnamed')]
            
            if i == 0:
                df_cabecalhos_pm = df_cabecalhos_pm_i
            else:
                df_cabecalhos_pm = pd.concat([df_cabecalhos_pm, df_cabecalhos_pm_i], ignore_index=True, sort=False)
        
        
        for i in range(len(arquivos_cabecalho_tl)):
        
            ## Ler a tabela específica da planilha
            
            ### MODELO TABELAO_SAP:
            df_cabecalhos_tl_i = pd.read_excel(arquivos_cabecalho_tl[i], sheet_name='Cabeçalho da lista de tarefas', skiprows=0,dtype = str)
            
            df_cabecalhos_tl_i["CONCAT"] = df_cabecalhos_tl_i["Centro"].map(str, na_action=None) + df_cabecalhos_tl_i["Descrição antiga"].map(str, na_action='ignore')
            
            df_cabecalhos_tl_i = df_cabecalhos_tl_i[pd.notna(df_cabecalhos_tl_i['Chave de grupo'])].copy().reset_index(drop=True).loc[:, ~df_cabecalhos_tl_i.columns.str.contains('^Unnamed')]
            
            if i == 0:
                df_cabecalhos_tl = df_cabecalhos_tl_i
            else:
                df_cabecalhos_tl = pd.concat([df_cabecalhos_tl, df_cabecalhos_tl_i], ignore_index=True, sort=False)
            
    
         
        LSMW_ADD_PM = {
        
                    'Plano de manutenção': [],
        
                    'Ctg.plano de manutenção': [],
                    'Texto do plano de manutenção':[],
                    'Ciclo de manutenção':[],
                    'Unidade para a execução de medidas':[],
                    'Texto p/o pacote ou ciclo de manute':[],
                    'Grupo de autorizações referente ao':[],
                    'Intervalo de solicitação de manuten':[],
                    'Unidade para o intervalo de solicit':[],
                    'Horizonte de abertura p/ solicitaçõ':[],
                    'Data de início':[],
                    'Texto breve do item':[],
                    'Nº equipamento':[],
                    'Centro de planejamento de manutenção':[],
                    'Grupo de planejamento para serviços cliente e manutenção':[],
                    'Tipo de ordem':[],
                    'Tipo de atividade de manutenção':[],
                    'Centro de trabalho principal para medidas de manutenção': [],
                    'Tipo de roteiro': [],
                    'Chave para grupo de listas de tarefas': [],
                    'Numerador de grupos': [],
                    'Local de instalação': [],
                    'Centro relativo ao centro de trabalho responsável': [],
                    'Prioridade': []
                    #'CONCAT': []
        
        }
        
        
        for j in range(len(df_cabecalhos_pm['Descrição'])):
        
          # Chave de grupo e descrição do item: **************
        
          try:
            valor_procurado = df_cabecalhos_pm['CONCAT'].iloc[j]
            index_tl = df_cabecalhos_tl[df_cabecalhos_tl['CONCAT'] == valor_procurado].index[0]
          except:
            index_tl = np.nan
        
          if not pd.isna(index_tl):
            chave_gp_tl = df_cabecalhos_tl['Chave de grupo'][index_tl]
            desc_tl = df_cabecalhos_tl['Descrição roteiro'][index_tl]
            LSMW_ADD_PM['Chave para grupo de listas de tarefas'].append(chave_gp_tl)
            LSMW_ADD_PM['Texto breve do item'].append(desc_tl)
          else:
            LSMW_ADD_PM['Chave para grupo de listas de tarefas'].append(np.nan)
            LSMW_ADD_PM['Texto breve do item'].append(df_cabecalhos_pm['Descrição'][j])
        
        
          # Outros:
        
          LSMW_ADD_PM['Tipo de ordem'].append('ZM04' if 'PRED' in str(LSMW_ADD_PM['Texto breve do item'][-1])[0:5] else 'ZM03')
          LSMW_ADD_PM['Ctg.plano de manutenção'].append('PM')
          LSMW_ADD_PM['Intervalo de solicitação de manuten'].append(335)
          LSMW_ADD_PM['Unidade para o intervalo de solicit'].append('DIA')
          LSMW_ADD_PM['Horizonte de abertura p/ solicitaçõ'].append(100)
          LSMW_ADD_PM['Data de início'].append(np.nan)
          LSMW_ADD_PM['Tipo de roteiro'].append('A')
          LSMW_ADD_PM['Numerador de grupos'].append(1)
        
          LSMW_ADD_PM['Nº equipamento'].append(df_cabecalhos_pm['ID_SAP'][j])
        
          LSMW_ADD_PM['Local de instalação'].append(df_cabecalhos_pm['Local de instalação'][j])
        
          LSMW_ADD_PM['Centro de planejamento de manutenção'].append(df_cabecalhos_pm['LI_N3'][j][0:4])
        
          LSMW_ADD_PM['Centro relativo ao centro de trabalho responsável'].append(df_cabecalhos_pm['LI_N3'][j][0:4])
        
          LSMW_ADD_PM['Grupo de autorizações referente ao'].append('MN01')
        
          LSMW_ADD_PM['Centro de trabalho principal para medidas de manutenção'].append(df_cabecalhos_pm['Centro de trabalho'][j])
        
          LSMW_ADD_PM['Grupo de planejamento para serviços cliente e manutenção'].append(df_cabecalhos_pm['LI_N3'][j][9:14])
        
        
          # Tipo de atividade:
        
          if 'TRO' in str(LSMW_ADD_PM['Texto breve do item'][-1])[0:4]:
            LSMW_ADD_PM['Tipo de atividade de manutenção'].append('Z02')
          elif 'REA' in str(LSMW_ADD_PM['Texto breve do item'][-1])[0:4]:
            LSMW_ADD_PM['Tipo de atividade de manutenção'].append('Z03')
          elif 'LIM' in str(LSMW_ADD_PM['Texto breve do item'][-1])[0:4]:
            LSMW_ADD_PM['Tipo de atividade de manutenção'].append('Z04')
          elif 'REV' in str(LSMW_ADD_PM['Texto breve do item'][-1])[0:4]:
            LSMW_ADD_PM['Tipo de atividade de manutenção'].append('Z05')
          elif 'LUB' in str(LSMW_ADD_PM['Texto breve do item'][-1])[0:4]:
            LSMW_ADD_PM['Tipo de atividade de manutenção'].append('Z06')
          elif 'AJU' in str(LSMW_ADD_PM['Texto breve do item'][-1])[0:4]:
            LSMW_ADD_PM['Tipo de atividade de manutenção'].append('Z07')
          elif 'CAL' in str(LSMW_ADD_PM['Texto breve do item'][-1])[0:4]:
            LSMW_ADD_PM['Tipo de atividade de manutenção'].append('Z08')
          elif 'INS' in str(LSMW_ADD_PM['Texto breve do item'][-1])[0:4]:
            LSMW_ADD_PM['Tipo de atividade de manutenção'].append('Z11')
          elif 'TES' in str(LSMW_ADD_PM['Texto breve do item'][-1])[0:4]:
            LSMW_ADD_PM['Tipo de atividade de manutenção'].append('Z12')
          elif 'PRED' in str(LSMW_ADD_PM['Texto breve do item'][-1])[0:5]:
            LSMW_ADD_PM['Tipo de atividade de manutenção'].append('Z14')
          else:
            LSMW_ADD_PM['Tipo de atividade de manutenção'].append(np.nan)
        
        
          # Prioridade:
        
          if any(item in str(LSMW_ADD_PM['Texto breve do item'][-1])[0:4] for item in ['REV','LUB']):
            LSMW_ADD_PM['Prioridade'].append(2)
          else:
            LSMW_ADD_PM['Prioridade'].append(3)
        
        
          # Periodicidade:
        
          try:
            if float(df_cabecalhos_pm['Condições da instalação'][j]) == 0:
              periodicidade_pm = df_cabecalhos_pm['Descrição'][j].split(' ')[2]
              LSMW_ADD_PM['Ciclo de manutenção'].append(periodicidade_pm)
            elif float(df_cabecalhos_pm['Condições da instalação'][j]) == 1:
              periodicidade_pm = df_cabecalhos_pm['Descrição'][j].split(' ')[3]
              LSMW_ADD_PM['Ciclo de manutenção'].append(periodicidade_pm)
            else:
              periodicidade_pm = np.nan
              LSMW_ADD_PM['Ciclo de manutenção'].append(np.nan)
          except:
            periodicidade_pm = np.nan
            LSMW_ADD_PM['Ciclo de manutenção'].append(np.nan)
        
        
          ## Desc Periodicidade
          if any(item in str(LSMW_ADD_PM['Texto breve do item'][-1]) for item in [' 1M ',' 21D ']):
            LSMW_ADD_PM['Texto p/o pacote ou ciclo de manute'].append('MENSAL')
          elif ' 2M ' in str(LSMW_ADD_PM['Texto breve do item'][-1]):
            LSMW_ADD_PM['Texto p/o pacote ou ciclo de manute'].append('BIMESTRAL')
          elif ' 3M ' in str(LSMW_ADD_PM['Texto breve do item'][-1]):
            LSMW_ADD_PM['Texto p/o pacote ou ciclo de manute'].append('TRIMESTRAL')
          elif ' 4M ' in str(LSMW_ADD_PM['Texto breve do item'][-1]):
            LSMW_ADD_PM['Texto p/o pacote ou ciclo de manute'].append('QUADRIMESTRAL')
          elif ' 5M ' in str(LSMW_ADD_PM['Texto breve do item'][-1]):
            LSMW_ADD_PM['Texto p/o pacote ou ciclo de manute'].append('QUINQUIMESTRAL')
          elif ' 6M ' in str(LSMW_ADD_PM['Texto breve do item'][-1]):
            LSMW_ADD_PM['Texto p/o pacote ou ciclo de manute'].append('SEMESTRAL')
          elif any(item in str(LSMW_ADD_PM['Texto breve do item'][-1]) for item in [' 12M ',' 1A ']):
            LSMW_ADD_PM['Texto p/o pacote ou ciclo de manute'].append('ANUAL')
          elif ' 2A ' in str(LSMW_ADD_PM['Texto breve do item'][-1]):
            LSMW_ADD_PM['Texto p/o pacote ou ciclo de manute'].append('BIENAL')
          elif ' 3A ' in str(LSMW_ADD_PM['Texto breve do item'][-1]):
            LSMW_ADD_PM['Texto p/o pacote ou ciclo de manute'].append('TRIENAL')
          elif ' 4A ' in str(LSMW_ADD_PM['Texto breve do item'][-1]):
            LSMW_ADD_PM['Texto p/o pacote ou ciclo de manute'].append('QUADRIENAL')
          elif ' 5A ' in str(LSMW_ADD_PM['Texto breve do item'][-1]):
            LSMW_ADD_PM['Texto p/o pacote ou ciclo de manute'].append('QUINQUENAL')
          elif ' 6A ' in str(LSMW_ADD_PM['Texto breve do item'][-1]):
            LSMW_ADD_PM['Texto p/o pacote ou ciclo de manute'].append('SEXENAL')
          elif ' 7A ' in str(LSMW_ADD_PM['Texto breve do item'][-1]):
            LSMW_ADD_PM['Texto p/o pacote ou ciclo de manute'].append('SEPTENAL')
          elif ' 10A ' in str(LSMW_ADD_PM['Texto breve do item'][-1]):
            LSMW_ADD_PM['Texto p/o pacote ou ciclo de manute'].append('DECENAL')
          elif ' 12A ' in str(LSMW_ADD_PM['Texto breve do item'][-1]):
            LSMW_ADD_PM['Texto p/o pacote ou ciclo de manute'].append('DOZENAL')
          elif any(item in str(LSMW_ADD_PM['Texto breve do item'][-1]) for item in [' 7D ',' 5D ',' 1S ']):
            LSMW_ADD_PM['Texto p/o pacote ou ciclo de manute'].append('SEMANAL')
          elif ' 1D ' in str(LSMW_ADD_PM['Texto breve do item'][-1]):
            LSMW_ADD_PM['Texto p/o pacote ou ciclo de manute'].append('DIARIO')
          elif ' 13D ' in str(LSMW_ADD_PM['Texto breve do item'][-1]):
            LSMW_ADD_PM['Texto p/o pacote ou ciclo de manute'].append('TREZENAL')
          elif any(item in str(LSMW_ADD_PM['Texto breve do item'][-1]) for item in [' 2S ',' 14D ',' 15D ']):
            LSMW_ADD_PM['Texto p/o pacote ou ciclo de manute'].append('QUINZENAL')
        
          elif isinstance(LSMW_ADD_PM['Ciclo de manutenção'][-1],str):
            if LSMW_ADD_PM['Ciclo de manutenção'][-1][-1] == 'M':
              LSMW_ADD_PM['Texto p/o pacote ou ciclo de manute'].append(LSMW_ADD_PM['Ciclo de manutenção'][-1][:-1]+' '+'MESES')
            elif LSMW_ADD_PM['Ciclo de manutenção'][-1][-1] == 'S':
              LSMW_ADD_PM['Texto p/o pacote ou ciclo de manute'].append(LSMW_ADD_PM['Ciclo de manutenção'][-1][:-1]+' '+'SEMANAS')
            elif LSMW_ADD_PM['Ciclo de manutenção'][-1][-1] == 'D':
              LSMW_ADD_PM['Texto p/o pacote ou ciclo de manute'].append(LSMW_ADD_PM['Ciclo de manutenção'][-1][:-1]+' '+'DIAS')
            elif LSMW_ADD_PM['Ciclo de manutenção'][-1][-1] == 'H':
              LSMW_ADD_PM['Texto p/o pacote ou ciclo de manute'].append(LSMW_ADD_PM['Ciclo de manutenção'][-1][:-1]+' '+'HORAS')
            elif LSMW_ADD_PM['Ciclo de manutenção'][-1][-1] == 'A':
              LSMW_ADD_PM['Texto p/o pacote ou ciclo de manute'].append(LSMW_ADD_PM['Ciclo de manutenção'][-1][:-1]+' '+'ANOS')
            else:
              LSMW_ADD_PM['Texto p/o pacote ou ciclo de manute'].append(np.nan)
          else:
            LSMW_ADD_PM['Texto p/o pacote ou ciclo de manute'].append(np.nan)
        
        
          ##  Valor Periodicidade
          try:
            if LSMW_ADD_PM['Ciclo de manutenção'][-1][-1] == 'M':
              LSMW_ADD_PM['Ciclo de manutenção'][-1] = float(LSMW_ADD_PM['Ciclo de manutenção'][-1].split('M')[0])*28
              LSMW_ADD_PM['Unidade para a execução de medidas'].append('DIA')
            elif LSMW_ADD_PM['Ciclo de manutenção'][-1][-1] == 'S':
              LSMW_ADD_PM['Ciclo de manutenção'][-1] = float(LSMW_ADD_PM['Ciclo de manutenção'][-1].split('S')[0])*7
              LSMW_ADD_PM['Unidade para a execução de medidas'].append('DIA')
            elif LSMW_ADD_PM['Ciclo de manutenção'][-1][-1] == 'D':
              LSMW_ADD_PM['Ciclo de manutenção'][-1] = round(round(float(LSMW_ADD_PM['Ciclo de manutenção'][-1].split('D')[0]))/7)*7
              LSMW_ADD_PM['Unidade para a execução de medidas'].append('DIA')
            elif LSMW_ADD_PM['Ciclo de manutenção'][-1][-1] == 'H':
              LSMW_ADD_PM['Ciclo de manutenção'][-1] = float(LSMW_ADD_PM['Ciclo de manutenção'][-1].split('H')[0])
              LSMW_ADD_PM['Unidade para a execução de medidas'].append('H')
            elif LSMW_ADD_PM['Ciclo de manutenção'][-1][-1] == 'A' and float(LSMW_ADD_PM['Ciclo de manutenção'][j].split('A')[0])>=3:
              LSMW_ADD_PM['Ciclo de manutenção'][-1] = float(LSMW_ADD_PM['Ciclo de manutenção'][-1].split('A')[0])*12
              LSMW_ADD_PM['Unidade para a execução de medidas'].append('MES')
            elif LSMW_ADD_PM['Ciclo de manutenção'][-1][-1] == 'A' and float(LSMW_ADD_PM['Ciclo de manutenção'][j].split('A')[0])<3:
              LSMW_ADD_PM['Ciclo de manutenção'][-1] = float(LSMW_ADD_PM['Ciclo de manutenção'][-1].split('A')[0])*28*12
              LSMW_ADD_PM['Unidade para a execução de medidas'].append('DIA')
          except:
              LSMW_ADD_PM['Ciclo de manutenção'][-1] = np.nan
              LSMW_ADD_PM['Unidade para a execução de medidas'].append(np.nan)
        
        
          # Nome plano:
        
          tipo_op = LSMW_ADD_PM['Texto breve do item'][-1]
        
          if isinstance(tipo_op, str):
        
            concat_centro_desc_prd = str(LSMW_ADD_PM['Centro de planejamento de manutenção'][-1]) + 'PREV PRD '+str(periodicidade_pm)+' '+tipo_op.split(' ')[-1]
        
            if ' FUNC ' in tipo_op:
              cod_prev = 'PREV FUNC '+str(periodicidade_pm)+' '+tipo_op.split(' ')[-1]
        
            elif concat_centro_desc_prd in SAP_PMI['CONCAT CENTRO_DESC'].values:
              cod_prev = 'PREV PRD '+str(periodicidade_pm)+' '+tipo_op.split(' ')[-1]
        
            else:
              cod_prev = tipo_op + '### REVISAR PREV OU SE EXISTE ### PREV PRD '+str(periodicidade_pm)+' '+tipo_op.split(' ')[-1]
        
          else:
            cod_prev = np.nan
        
          LSMW_ADD_PM['Texto do plano de manutenção'].append(cod_prev)
          LSMW_ADD_PM['Plano de manutenção'].append(np.nan)
        
        
        ############################
        
        
        LSMW_ADD_PM = pd.DataFrame(LSMW_ADD_PM)
        LSMW_ADD_PM['CONCAT'] = LSMW_ADD_PM["Centro de planejamento de manutenção"].map(str, na_action=None) + LSMW_ADD_PM["Texto do plano de manutenção"].map(str, na_action='ignore')
        
        # Remover linhas NaN e adicioná-las a DF de ERRO CARGA
        
        ## Máscara booleana para linhas onde a coluna está vazia
        mask = LSMW_ADD_PM['Chave para grupo de listas de tarefas'].isna()
        #mask = LSMW_ADD_PM['Texto do plano de manutenção'].isna()
        
        ## Criar um novo DataFrame com as linhas onde a coluna está vazia
        LSMW_ADD_PM_ERRO_CARGA = LSMW_ADD_PM[mask]
        
        ## Remover as linhas onde a coluna está vazia do DataFrame original
        LSMW_ADD_PM = LSMW_ADD_PM[~mask]
        
        
        ############################
        
        
        # Criar índices e separar planilha LSMW para ADD_ITEM (ADD_PLANO deve manter apenas os índices = 1, e a ADD_ITEM os demais)
        
        LSMW_ADD_PM = LSMW_ADD_PM.sort_values(by= ['CONCAT','Texto breve do item']).reset_index(drop=True)    # Ordenar valores
        
        indice_inserir_coluna = LSMW_ADD_PM.columns.get_loc(LSMW_ADD_PM.columns[-1]) + 1
        LSMW_ADD_PM.insert(loc=indice_inserir_coluna, column='indice_pm', value=np.nan)
        
        for i in range(len(LSMW_ADD_PM['indice_pm'])):
          #try:
          ##  Trazer número de planos já existentes
          if LSMW_ADD_PM['CONCAT'][i] in SAP_PMI['CONCAT CENTRO_DESC'].values:
            valor_procurado = LSMW_ADD_PM['CONCAT'][i]
            index_procurado = SAP_PMI[SAP_PMI['CONCAT CENTRO_DESC'] == valor_procurado].index
            if not index_procurado.empty:
                LSMW_ADD_PM.at[i, 'Plano de manutenção'] = SAP_PMI.at[index_procurado[0], 'Maintenance Plan']
                print(SAP_PMI['Maintenance Plan'][index_procurado[0]])############
          #except:
           # pass
        
          ##  Criar índices para enviar índices maiores que 1 para a planilha de ADD_ITEM
          if i == 0:
            LSMW_ADD_PM['indice_pm'][i] = 1
            continue
        
          if LSMW_ADD_PM['CONCAT'][i] != LSMW_ADD_PM['CONCAT'][i-1]:
            LSMW_ADD_PM['indice_pm'][i] = 1
        
          #try:
          ### Verificar se o plano do índice 1 já existe (será enviado direto para a planilha de itens):
          if LSMW_ADD_PM['indice_pm'][i] == 1 and LSMW_ADD_PM['CONCAT'][i] in SAP_PMI['CONCAT CENTRO_DESC'].values:
            print(2)
            LSMW_ADD_PM['indice_pm'][i] = 2
          #except:
           # pass
        
          if LSMW_ADD_PM['CONCAT'][i] == LSMW_ADD_PM['CONCAT'][i-1]:
            LSMW_ADD_PM['indice_pm'][i] = 1 + LSMW_ADD_PM['indice_pm'][i-1]
        
        
        ##  Separando planilha para a ADD_ITEM
        
        LSMW_ADD_PM_ITEM = LSMW_ADD_PM[LSMW_ADD_PM['indice_pm'] != 1].copy().reset_index(drop=True)
        LSMW_ADD_PM1 = LSMW_ADD_PM
        LSMW_ADD_PM = LSMW_ADD_PM[LSMW_ADD_PM['indice_pm'] == 1].copy().reset_index(drop=True)

        ###

        #   Fazer concat group + equipment or functional location - Incluído 05/06/2024

        LSMW_ADD_PM['CONCAT TL_EQP'] = np.where(  # Incluído 05/06/2024
          LSMW_ADD_PM['Nº equipamento'].notna(),  # condição: se 'Equipment' não for NaN
          LSMW_ADD_PM["Chave para grupo de listas de tarefas"].map(str) + LSMW_ADD_PM["Nº equipamento"].map(str),  # se verdadeiro: group + equipment
          LSMW_ADD_PM["Chave para grupo de listas de tarefas"].map(str) + LSMW_ADD_PM["Local de instalação"].map(str)  # se falso: group + functional location
        )
        
        #

        ############################

        import io
        buffer1 = io.BytesIO()
        
        # Salvando em arquivo excel
        
        with pd.ExcelWriter(buffer1, engine="xlsxwriter") as excel_writer:
            ## Crie um objeto ExcelWriter
            nome_arquivo_sap = 'CONCAT_LSMW_ADD_PM_'
        
            ## Salve cada DataFrame em uma planilha diferente
            LSMW_ADD_PM.to_excel(excel_writer, sheet_name='ADD_PM', index=False)
            LSMW_ADD_PM_ERRO_CARGA.to_excel(excel_writer, sheet_name='ERRO_CARGA', index=False)
            LSMW_ADD_PM1.to_excel(excel_writer, sheet_name='ADD_PM+ITEM_REVISAR', index=False)
        
            ## Feche o objeto ExcelWriter
            excel_writer.close()
        
            st.download_button(
                label="Download "+nome_arquivo_sap,
                data=buffer1,
                file_name=nome_arquivo_sap+'.xlsx',
            )
        
            ###


    with st.spinner('Carregando Itens dos Planos...'):

        # ADD ITEM:
        
        LSMW_ADD_ITEM = {
        
                    'Plano de manutenção': [],
        
                    'Texto breve do item': [],
                    'Nº equipamento':[],
                    'Centro de planejamento de manutenção':[],
                    'Grupo de planejamento para serviços cliente e manutenção':[],
                    'Tipo de ordem':[],
                    'Tipo de atividade de manutenção':[],
                    'Centro de trabalho principal para medidas de manutenção':[],
                    'Tipo de roteiro':[],
                    'Chave para grupo de listas de tarefas':[],
                    'Numerador de grupos':[],
                    'Local de instalação':[],
                    'Centro relativo ao centro de trabalho responsável':[],
                    'Prioridade':[]
        
        }
        
        
        LSMW_ADD_ITEM['Plano de manutenção'] = LSMW_ADD_PM_ITEM['Plano de manutenção']
        
        LSMW_ADD_ITEM['Texto breve do item'] = LSMW_ADD_PM_ITEM['Texto breve do item']
        
        LSMW_ADD_ITEM['Nº equipamento'] = LSMW_ADD_PM_ITEM['Nº equipamento']
        
        LSMW_ADD_ITEM['Centro de planejamento de manutenção'] = LSMW_ADD_PM_ITEM['Centro de planejamento de manutenção']
        
        LSMW_ADD_ITEM['Grupo de planejamento para serviços cliente e manutenção'] = LSMW_ADD_PM_ITEM['Grupo de planejamento para serviços cliente e manutenção']
        
        LSMW_ADD_ITEM['Tipo de ordem'] = LSMW_ADD_PM_ITEM['Tipo de ordem']
        
        LSMW_ADD_ITEM['Tipo de atividade de manutenção'] = LSMW_ADD_PM_ITEM['Tipo de atividade de manutenção']
        
        LSMW_ADD_ITEM['Centro de trabalho principal para medidas de manutenção'] = LSMW_ADD_PM_ITEM['Centro de trabalho principal para medidas de manutenção']
        
        LSMW_ADD_ITEM['Tipo de roteiro'] = LSMW_ADD_PM_ITEM['Tipo de roteiro']
        
        LSMW_ADD_ITEM['Chave para grupo de listas de tarefas'] = LSMW_ADD_PM_ITEM['Chave para grupo de listas de tarefas']
        
        LSMW_ADD_ITEM['Numerador de grupos'] = LSMW_ADD_PM_ITEM['Numerador de grupos']
        
        LSMW_ADD_ITEM['Local de instalação'] = LSMW_ADD_PM_ITEM['Local de instalação']
        
        LSMW_ADD_ITEM['Centro relativo ao centro de trabalho responsável'] = LSMW_ADD_PM_ITEM['Centro relativo ao centro de trabalho responsável']
        
        LSMW_ADD_ITEM['Prioridade'] = LSMW_ADD_PM_ITEM['Prioridade']
        
        LSMW_ADD_ITEM['TXT PLANO'] = LSMW_ADD_PM_ITEM['Texto do plano de manutenção']
        
        LSMW_ADD_ITEM = pd.DataFrame(LSMW_ADD_ITEM)
        
        
        ############################
        
        
        # Verificar itens que já subiram

        indice_inserir_coluna = LSMW_ADD_ITEM.columns.get_loc(LSMW_ADD_ITEM.columns[-1]) + 1
        LSMW_ADD_ITEM.insert(loc=indice_inserir_coluna, column='carregado?', value=np.nan)
        
        LSMW_ADD_ITEM['CONCAT'] = np.where(  # Incluído 05/06/2024
          LSMW_ADD_ITEM['Nº equipamento'].notna(),  # condição: se 'Equipment' não for NaN
          LSMW_ADD_ITEM["Chave para grupo de listas de tarefas"].map(str) + LSMW_ADD_ITEM["Nº equipamento"].map(str),  # se verdadeiro: group + equipment
          LSMW_ADD_ITEM["Chave para grupo de listas de tarefas"].map(str) + LSMW_ADD_ITEM["Local de instalação"].map(str)  # se falso: group + functional location
        )
        
        
        for i in range(len(LSMW_ADD_ITEM['carregado?'])):
          #try:
          ##  Trazer número de planos já existentes
          if LSMW_ADD_ITEM['CONCAT'][i] in SAP_PMI['CONCAT TL_EQP'].values:
            valor_procurado = LSMW_ADD_ITEM['CONCAT'][i]
            index_procurado = SAP_PMI[SAP_PMI['CONCAT TL_EQP'] == valor_procurado].index
            if not index_procurado.empty and not pd.isna(valor_procurado):
                LSMW_ADD_ITEM.at[i, 'carregado?'] = SAP_PMI.at[index_procurado[0], 'Maintenance Item']
        
          if LSMW_ADD_ITEM['CONCAT'][i] in LSMW_ADD_PM['CONCAT TL_EQP'].values:   # Para o ADD_PM - Incluído 05/06/2024
            valor_procurado = LSMW_ADD_ITEM['CONCAT'][i]
            index_procurado = LSMW_ADD_PM[LSMW_ADD_PM['CONCAT TL_EQP'] == valor_procurado].index
            if not index_procurado.empty and not pd.isna(valor_procurado):
                LSMW_ADD_ITEM.at[i, 'carregado?'] = LSMW_ADD_PM.at[index_procurado[0], 'CONCAT TL_EQP']
          #except:
           # pass
        
        LSMW_ADD_ITEM.drop_duplicates(inplace=True)#.reset_index(drop = True)    # Incluído 05/06/2024
        
        LSMW_ADD_ITEM_NOVO = LSMW_ADD_ITEM[pd.isna(LSMW_ADD_ITEM['carregado?'])]
        LSMW_ADD_ITEM_EXIS = LSMW_ADD_ITEM[~pd.isna(LSMW_ADD_ITEM['carregado?'])]
        
        
        ############################
        
        import io
        buffer2 = io.BytesIO()
        
        # Salvando em arquivo excel
        
        with pd.ExcelWriter(buffer2, engine="xlsxwriter") as excel_writer:
            ## Crie um objeto ExcelWriter
            nome_arquivo_sap2 = 'CONCAT_LSMW_ADD_ITEM_'
        
            ## Salve cada DataFrame em uma planilha diferente
            LSMW_ADD_ITEM_NOVO.to_excel(excel_writer, sheet_name='ADD_ITEM', index=False)
            LSMW_ADD_ITEM_EXIS.to_excel(excel_writer, sheet_name='ITENS CARREGADOS', index=False)
        
            ## Feche o objeto ExcelWriter
            excel_writer.close()
        
            st.download_button(
                label="Download "+nome_arquivo_sap2,
                data=buffer2,
                file_name=nome_arquivo_sap2+'.xlsx',
            )
        
            ###
