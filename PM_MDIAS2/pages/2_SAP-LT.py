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


try:
    if 'SAP_EQP' not in st.session_state:
        with st.spinner('Carregando Equipamentos SAP...'):
            SAP_EQP = pd.read_excel(my_path + 'SAP_EQP_05-04.xlsx', sheet_name='Sheet1', skiprows=0, dtype=str)
            SAP_EQP['CONCAT CENTRO_DESC'] = SAP_EQP["Centro planejamento"].map(str, na_action=None) + SAP_EQP[
                "Denominação do objeto técnico"].map(str, na_action='ignore')
            SAP_EQP = pd.DataFrame(SAP_EQP)
            st.session_state.SAP_EQP = SAP_EQP
    SAP_EQP = st.session_state['SAP_EQP']
except:
    pass

try:
    if 'SAP_TL' not in st.session_state:
        with st.spinner('Carregando Task Lists SAP...'):
            SAP_TL = pd.read_excel(my_path + 'SAP_TL_04-04.xlsx', sheet_name='Sheet1', skiprows=0, dtype=str)
            SAP_TL['CONCAT CENTRO_DESC'] = SAP_TL["Centro planejamento"].map(str, na_action=None) + SAP_TL["Descrição"].map(str,na_action='ignore')
            SAP_TL = pd.DataFrame(SAP_TL)
            st.session_state.SAP_TL = SAP_TL
    SAP_TL = st.session_state['SAP_TL']
except:
    pass

try:
    if 'SAP_PMI' not in st.session_state:
        with st.spinner('Carregando Planos SAP...'):
            SAP_PMI = pd.read_excel(my_path + 'SAP_PMI_08-04_.xlsx', sheet_name='Sheet1', skiprows=0, dtype=str)
            SAP_PMI['CONCAT CENTRO_DESC'] = SAP_PMI["Planning Plant"].map(str, na_action=None) + SAP_PMI[
                "Maintenance Plan Desc"].map(str, na_action='ignore')
            SAP_PMI['CONCAT TL_EQP'] = SAP_PMI["Group"].map(str, na_action=None) + SAP_PMI["Equipment"].map(str,na_action='ignore')

            SAP_PMI = pd.DataFrame(SAP_PMI)
            st.session_state.SAP_PMI = SAP_PMI
    SAP_PMI = st.session_state['SAP_PMI']
except:
    pass

    if 'SAP_EQP' in st.session_state:
        SAP_EQP = st.session_state['SAP_EQP']
    if 'SAP_TL' in st.session_state:
        SAP_TL = st.session_state['SAP_TL']
    if 'SAP_PMI' in st.session_state:
        SAP_PMI = st.session_state['SAP_PMI']


uploaded_file = st.file_uploader("Carregar 'TABELAO...' com listas de tarefas geradas na página TL")
if uploaded_file is not None:

    with st.spinner('Carregando Cabeçalho Task List...'):

        # ADD TASK LIST:

        # Nomes dos arquivos e planilhas a serem lidos:

        arquivos_cabecalho_operacoes = []

        arquivos_cabecalho_operacoes.append(uploaded_file)   # Nome do arquivo

        # TAREFAS SEM LUBRIFICAÇÃO E CALIBRAÇÃO

        LSMW_CAB = pd.DataFrame({

                    'Chave de grupo': [],
                    'Data fixada':[],
                    'Numerador de grupos':[],
                    'Descrição roteiro':[],
                    'Centro':[],
                    'A&D: ID externo da lista de tarefas':[],
                    'Centro de trabalho':[],
                    'Utilização do plano':[],
                    'Grupo de planejamento ou departamento responsável':[],
                    'Status':[],
                    'Condições da instalação':[],

                    'Descrição antiga':[],
                    'CONCAT CENTRO_DESC':[]

        })

        for i in range(len(arquivos_cabecalho_operacoes)):

          LSMW_CAB_i = pd.DataFrame({

                      'Chave de grupo': [],
                      'Data fixada':[],
                      'Numerador de grupos':[],
                      'Descrição roteiro':[],
                      'Centro':[],
                      'A&D: ID externo da lista de tarefas':[],
                      'Centro de trabalho':[],
                      'Utilização do plano':[],
                      'Grupo de planejamento ou departamento responsável':[],
                      'Status':[],
                      'Condições da instalação':[],

                      'Descrição antiga':[],
                      'CONCAT CENTRO_DESC':[]

          })


          # Leitura do arquivo baixado

          ## Ler a tabela específica da planilha

          ### MODELO PLANILHA DE CARGA:
          #df_cabecalhos_i = pd.read_excel(arquivos_cabecalho_operacoes[i]+'.xlsx', sheet_name='Cabeçalho da lista de tarefas', usecols="A:Q", skiprows=7,dtype = str)

          ### MODELO TABELAO_SAP:
          df_cabecalhos_i = pd.read_excel(arquivos_cabecalho_operacoes[i], sheet_name='CABECALHO S CALIB LUB', skiprows=0,dtype = str)
          df_cabecalhos_i = pd.concat([df_cabecalhos_i , pd.read_excel(arquivos_cabecalho_operacoes[i], sheet_name='CABECALHO LUB', skiprows=0,dtype = str)], ignore_index=True, sort=False)
          df_cabecalhos_i = df_cabecalhos_i[pd.notna(df_cabecalhos_i['Chave do grupo de listas de tarefas*'])].copy().reset_index(drop=True).loc[:, ~df_cabecalhos_i.columns.str.contains('^Unnamed')]

          df_cabecalhos_i_columns = list(df_cabecalhos_i.columns)

          LSMW_CAB_i['CONCAT CENTRO_DESC'] = df_cabecalhos_i["Centro de centro de trabalho"].map(str, na_action=None) + df_cabecalhos_i["Descrição antiga"].map(str, na_action='ignore')


          ## Transformando tabelas em DataFrame do Python
          LSMW_CAB_i['A&D: ID externo da lista de tarefas'] = df_cabecalhos_i[df_cabecalhos_i_columns[0]]
          LSMW_CAB_i['Numerador de grupos'] = df_cabecalhos_i[df_cabecalhos_i_columns[1]]
          LSMW_CAB_i['Data fixada'] = [data_hoje]*len(LSMW_CAB_i['A&D: ID externo da lista de tarefas'])
          LSMW_CAB_i['Descrição roteiro'] = df_cabecalhos_i[df_cabecalhos_i_columns[4]]
          LSMW_CAB_i['Centro'] = df_cabecalhos_i[df_cabecalhos_i_columns[5]]
          LSMW_CAB_i['Centro de trabalho'] = df_cabecalhos_i[df_cabecalhos_i_columns[6]]
          LSMW_CAB_i['Utilização do plano'] = df_cabecalhos_i[df_cabecalhos_i_columns[8]]
          LSMW_CAB_i['Grupo de planejamento ou departamento responsável'] = df_cabecalhos_i[df_cabecalhos_i_columns[11]]
          LSMW_CAB_i['Status'] = df_cabecalhos_i[df_cabecalhos_i_columns[9]]
          LSMW_CAB_i['Condições da instalação'] = df_cabecalhos_i[df_cabecalhos_i_columns[12]]

          LSMW_CAB_i['Chave de grupo'] = np.nan

          try:
            LSMW_CAB_i['Descrição antiga'] = df_cabecalhos_i['Descrição antiga']
          except:
            pass

          if i == 0:
            LSMW_CAB = LSMW_CAB_i
          else:
            LSMW_CAB = pd.concat([LSMW_CAB, LSMW_CAB_i], ignore_index=True, sort=False)


        # Checando de lista de tarefa já foi carregada (OBS: Não funciona se ela passar de 40 caracteres)
        try:
          LSMW_CAB.insert(loc=LSMW_CAB.columns.get_loc('CONCAT CENTRO_DESC') + 1, column='CARREGADO?', value=np.nan)

          for i in range(len(LSMW_CAB['Centro'])):

            if len(LSMW_CAB['Descrição antiga'][i]) <= 40:
                selected_rows = SAP_TL.loc[SAP_TL['CONCAT CENTRO_DESC'] == LSMW_CAB['CONCAT CENTRO_DESC'][i], 'Grupo']
                if not selected_rows.empty:
                    num_tl = selected_rows.iloc[0]
                    if isinstance(num_tl, str) or isinstance(num_tl, int):
                        LSMW_CAB.at[i, 'CARREGADO?'] = 1
                        LSMW_CAB.at[i, 'Chave de grupo'] = num_tl
                    else:
                        LSMW_CAB.at[i, 'CARREGADO?'] = 0
                else:
                    LSMW_CAB.at[i, 'CARREGADO?'] = 0
            else:
                LSMW_CAB.at[i, 'CARREGADO?'] = 'NECESSÁRIO CHECAGEM MANUAL'


          LSMW_CAB_CAR = LSMW_CAB[LSMW_CAB['CARREGADO?'] == 1]
          LSMW_CAB = LSMW_CAB[LSMW_CAB['CARREGADO?'] != 1]

        except:
          print("ERRO: VERIFICAÇÃO DE LISTAS EXISTENTES")
          pass
        #

        import io
        buffer = io.BytesIO()

        # Salvando em arquivo excel

        with pd.ExcelWriter(buffer, engine="xlsxwriter") as excel_writer:
            ## Crie um objeto ExcelWriter
            nome_arquivo_sap = 'CONCAT_LSMW_CAB_'

            ## Salve cada DataFrame em uma planilha diferente
            LSMW_CAB.to_excel(excel_writer, sheet_name='Cabeçalho da lista de tarefas', index=False)
            try:
                LSMW_CAB_CAR.to_excel(excel_writer, sheet_name='CARREGADO Cabeçalho',
                                      index=False)  # Listas já carregadas
            except:
                pass

            ## Feche o objeto ExcelWriter
            excel_writer.close()

            st.download_button(
                label="Download "+nome_arquivo_sap,
                data=buffer,
                file_name=nome_arquivo_sap+'.xlsx',
            )

            ###


    with st.spinner('Carregando Operações Task List...'):

        #   EDIT TASK LIST:

        # TAREFAS SEM LUBRIFICAÇÃO E CALIBRAÇÃO

        LSMW_OP = pd.DataFrame({

                    'Chave para grupo de listas de tarefas': [],
                    'Data fixada':[],
                    'Centro':[],
                    'Numerador de grupos':[],
                    'Entrada':[],
                    'Nº operação':[],
                    'Suboperação':[],
                    'Txt.breve operação':[],
                    'Trabalho da operação':[],
                    'Unidade de trabalho':[],
                    'Núm.capacidade necessária':[],
                    'Duração normal da operação':[],
                    'Unidade da duração normal':[],
                    'Chave de cálculo':[],
                    'Porcentagem de trabalho':[],
                    'Fator de execução':[],
                    'Nº equipamento':[],

                    'A&D: ID externo da lista de tarefas':[],
                    'CONCAT CENTRO_DESC':[]

        })


        for i in range(len(arquivos_cabecalho_operacoes)):

          LSMW_OP_i = pd.DataFrame({
        
                    'Chave para grupo de listas de tarefas': [],
                    'Data fixada':[],
                    'Centro':[],
                    'Numerador de grupos':[],
                    'Entrada':[],
                    'Nº operação':[],
                    'Suboperação':[],
                    'Txt.breve operação':[],
                    'Trabalho da operação':[],
                    'Unidade de trabalho':[],
                    'Núm.capacidade necessária':[],
                    'Duração normal da operação':[],
                    'Unidade da duração normal':[],
                    'Chave de cálculo':[],
                    'Porcentagem de trabalho':[],
                    'Fator de execução':[],
                    'Nº equipamento':[],
        
                    'A&D: ID externo da lista de tarefas':[],
                    'CONCAT CENTRO_DESC':[]
          })
        
          # Leitura do arquivo baixado
        
          ## Ler a tabela específica da planilha
        
          ### MODELO PLANILHA DE CARGA:
          #df_cabecalhos_i = pd.read_excel(arquivos_cabecalho_operacoes[i]+'.xlsx', sheet_name='Cabeçalho da lista de tarefas', usecols="A:Q", skiprows=7,dtype = str)
          #df_operacoes_i = pd.read_excel(arquivos_cabecalho_operacoes[i]+'.xlsx', sheet_name='Operações e atividades', usecols="A:AJ", skiprows=7,dtype = str)
        
          ### MODELO TABELAO_SAP:
          df_cabecalhos_i = pd.read_excel(arquivos_cabecalho_operacoes[i]+'.xlsx', sheet_name='CABECALHO S CALIB LUB', usecols="A:Q", skiprows=0,dtype = str)
          df_cabecalhos_i = pd.concat([df_cabecalhos_i , pd.read_excel(arquivos_cabecalho_operacoes[i]+'.xlsx', sheet_name='CABECALHO LUB', skiprows=0,dtype = str)], ignore_index=True, sort=False)
          df_operacoes_i = pd.read_excel(arquivos_cabecalho_operacoes[i]+'.xlsx', sheet_name='TAREFAS S CALIB LUB', usecols="A:AJ", skiprows=0,dtype = str)
          df_operacoes_i = pd.concat([df_operacoes_i , pd.read_excel(arquivos_cabecalho_operacoes[i]+'.xlsx', sheet_name='TAREFAS LUB', skiprows=0,dtype = str)], ignore_index=True, sort=False)
        
          df_cabecalhos_i = df_cabecalhos_i[pd.notna(df_cabecalhos_i['Chave do grupo de listas de tarefas*'])].copy().reset_index(drop=True).loc[:, ~df_cabecalhos_i.columns.str.contains('^Unnamed')]
          df_operacoes_i = df_operacoes_i[pd.notna(df_operacoes_i['Chave do grupo de listas de tarefas*'])].copy().reset_index(drop=True).loc[:, ~df_operacoes_i.columns.str.contains('^Unnamed')]
        
          df_operacoes_i_columns = list(df_operacoes_i.columns)
        
        
          list_entrada = []
          ent = 1
          for j in range(len(df_operacoes_i[df_operacoes_i_columns[0]])):
            if j == 0:
              list_entrada.append(ent)
            elif df_operacoes_i[df_operacoes_i_columns[0]][j] == df_operacoes_i[df_operacoes_i_columns[0]][j-1]:
              ent = ent + 1
              list_entrada.append(ent)
            else:
              ent = 1
              list_entrada.append(ent)
        
        
          ## Transformando tabelas em DataFrame do Python
          LSMW_OP_i['Numerador de grupos'] = np.nan
        
          LSMW_OP_i['Centro'] = df_operacoes_i[df_operacoes_i_columns[6]]
          LSMW_OP_i['Data fixada'] = [data_hoje]*len(LSMW_OP_i['Centro'])
          LSMW_OP_i['Entrada'] = list_entrada
          LSMW_OP_i['Nº operação'] = df_operacoes_i[df_operacoes_i_columns[2]]
          LSMW_OP_i['Suboperação'] = df_operacoes_i[df_operacoes_i_columns[4]]
          LSMW_OP_i['Txt.breve operação'] = df_operacoes_i[df_operacoes_i_columns[8]]
          LSMW_OP_i['Trabalho da operação'] = df_operacoes_i[df_operacoes_i_columns[13]]
          LSMW_OP_i['Unidade de trabalho'] = df_operacoes_i[df_operacoes_i_columns[14]]
        
          LSMW_OP_i['Núm.capacidade necessária'] = df_operacoes_i[df_operacoes_i_columns[16]]
          LSMW_OP_i['Duração normal da operação'] = df_operacoes_i[df_operacoes_i_columns[17]]
          LSMW_OP_i['Unidade da duração normal'] = df_operacoes_i[df_operacoes_i_columns[18]]
          LSMW_OP_i['Chave de cálculo'] = df_operacoes_i[df_operacoes_i_columns[12]]
          LSMW_OP_i['Porcentagem de trabalho'] = df_operacoes_i[df_operacoes_i_columns[19]]
          LSMW_OP_i['Fator de execução'] = df_operacoes_i[df_operacoes_i_columns[9]]
          LSMW_OP_i['Nº equipamento'] = df_operacoes_i[df_operacoes_i_columns[10]]
        
          LSMW_OP_i['Chave para grupo de listas de tarefas'] = np.nan
        
          LSMW_OP_i['A&D: ID externo da lista de tarefas'] = df_operacoes_i[df_operacoes_i_columns[0]]    #**
        
          if i == 0:
            LSMW_OP = LSMW_OP_i
          else:
            LSMW_OP = pd.concat([LSMW_OP, LSMW_OP_i], ignore_index=True, sort=False)
        
        
        
        #####################
        
        try:
          # Checando de lista de tarefa já foi carregada
        
          LSMW_OP.insert(loc=LSMW_OP.columns.get_loc('A&D: ID externo da lista de tarefas') + 1, column='CARREGADO?', value=np.nan)
        
        
          for i in range(len(LSMW_OP['Centro'])):
        
            ## Trazer o Concat CENTRO_DESC com base no cabeçalho LSMW da TL:
            valor_procurado = LSMW_OP['A&D: ID externo da lista de tarefas'][i]
            index_tl = LSMW_CAB_CAR[LSMW_CAB_CAR['A&D: ID externo da lista de tarefas'] == valor_procurado].index
            if not index_tl.empty:
                LSMW_OP.at[i, 'CONCAT CENTRO_DESC'] = LSMW_CAB_CAR.at[index_tl[0], 'CONCAT CENTRO_DESC']
                LSMW_OP.at[i, 'Chave para grupo de listas de tarefas'] = LSMW_CAB_CAR.at[index_tl[0], 'Chave de grupo']
        
            if pd.isna(LSMW_OP['CONCAT CENTRO_DESC'][i]):   # Trazer informações de carregamento para os que ainda não foram carregados ou devem ser checados manualmente
                index_tl = LSMW_CAB[LSMW_CAB['A&D: ID externo da lista de tarefas'] == valor_procurado].index
                if not index_tl.empty:
                    LSMW_OP.at[i, 'CONCAT CENTRO_DESC'] = LSMW_CAB.at[index_tl[0], 'CONCAT CENTRO_DESC']
                    LSMW_OP.at[i, 'CARREGADO?'] = LSMW_CAB.at[index_tl[0], 'CARREGADO?']
        
        
            ## Checar Equipamento + Subop repetidos
            if LSMW_OP['A&D: ID externo da lista de tarefas'][i] in LSMW_CAB_CAR['A&D: ID externo da lista de tarefas'].values:
        
              LSMW_OP['CARREGADO?'][i] = 1  # Indentificando operações já carregadas
        
              LSMW_OP_un = SAP_TL[SAP_TL['CONCAT CENTRO_DESC'] == LSMW_OP['CONCAT CENTRO_DESC'][i]].reset_index(drop=True)
              LSMW_OP_un['CONCAT EQP_OP'] = LSMW_OP_un["Equipamento operação"].astype(str) + LSMW_OP_un["Txt.breve operação"].astype(str) #*
              CONCAT_EQP_OP_i = str(LSMW_OP['Nº equipamento'][i])+str(LSMW_OP['Txt.breve operação'][i])
              if CONCAT_EQP_OP_i in LSMW_OP_un['CONCAT EQP_OP'].values:
                LSMW_OP['CARREGADO?'][i] = 'EQUIP+OP REPETIDO: EXCLUIR LINHA E ALTERAR NUM DA OP'
        
        
            ##  Se todas as operações são repetidas, excluir cabeçalho, se não, mantê-lo na aba de carregados (CARRGADOS? = 1)
          for i in range(len(LSMW_OP['Centro'])):
        
            if LSMW_OP['CARREGADO?'][i] == 'EQUIP+OP REPETIDO: EXCLUIR LINHA E ALTERAR NUM DA OP':    # Verifica se já é item que foi carregado
              if pd.isna(LSMW_OP['Suboperação'][i]) and 'LUB' not in str(LSMW_OP['Txt.breve operação'][i])[0:4]:   # Saber se é cabealho
                LSMW_OP_un = LSMW_OP[LSMW_OP['Chave para grupo de listas de tarefas'] == LSMW_OP['Chave para grupo de listas de tarefas'][i]].reset_index(drop=True)
                todos_iguais = LSMW_OP_un['CARREGADO?'].nunique() == 1    # Verificar se todos os termos iguais
                if not todos_iguais:
                  LSMW_OP['CARREGADO?'][i] = 1
        
          LSMW_OP = LSMW_OP[LSMW_OP['CARREGADO?'] != 'EQUIP+OP REPETIDO: EXCLUIR LINHA E ALTERAR NUM DA OP'].reset_index(drop=True)
        
        
            ## Número e tempo das operações:
          for i in range(len(LSMW_OP['Centro'])):
        
            if LSMW_OP['A&D: ID externo da lista de tarefas'][i] in LSMW_CAB_CAR['A&D: ID externo da lista de tarefas'].values:
        
              ### Trazer cabeçalho da lista i
              valor_procurado = LSMW_OP['A&D: ID externo da lista de tarefas'][i]
              index_tl = LSMW_CAB_CAR[LSMW_CAB_CAR['A&D: ID externo da lista de tarefas'] == valor_procurado].index
              cabecalho_lista_i = LSMW_CAB_CAR.at[index_tl[0], 'Descrição roteiro']
        
              ### Trazer número da última suboperação (não LUB) - OK FUNFOU
              if not pd.isna(LSMW_OP['Suboperação'][i]) and 'LUB' not in str(cabecalho_lista_i)[0:4]:
                if pd.isna(LSMW_OP['Suboperação'][i-1]):    # Saber se antes desta subop é o cabeçalho
                  LSMW_OP_un = SAP_TL[SAP_TL['CONCAT CENTRO_DESC'] == LSMW_OP['CONCAT CENTRO_DESC'][i]].reset_index(drop=True)
                  LSMW_OP_un['Suboperação'] = pd.to_numeric(LSMW_OP_un['Suboperação'], errors='coerce').fillna(0).astype(int)   # Coerce = erros viram NaN
                  nova_subop = max(LSMW_OP_un['Suboperação']) + 10
                  LSMW_OP['Suboperação'][i] = nova_subop
                  LSMW_OP['Entrada'][i] = (nova_subop/10) + 1 ###* testar
                  #LSMW_OP['CARREGADO?'][i] = 1
                else:
                  nova_subop = nova_subop + 10
                  LSMW_OP['Suboperação'][i] = nova_subop
                  LSMW_OP['Entrada'][i] = (nova_subop/10) + 1 ###* testar
        
              ### Trazer número da última operação (LUB) - OK FUNFOU
              elif 'LUB' in str(cabecalho_lista_i)[0:4]:
                if LSMW_OP['A&D: ID externo da lista de tarefas'][i-1] != LSMW_OP['A&D: ID externo da lista de tarefas'][i]:    # Saber se antes desta subop é outra op
                  LSMW_OP_un = SAP_TL[SAP_TL['CONCAT CENTRO_DESC'] == LSMW_OP['CONCAT CENTRO_DESC'][i]].reset_index(drop=True)
                  LSMW_OP_un['Operação'] = pd.to_numeric(LSMW_OP_un['Operação'], errors='coerce').fillna(0).astype(int)   # Coerce = erros viram NaN
                  nova_subop = max(LSMW_OP_un['Operação']) + 10
                  LSMW_OP['Nº operação'][i] = nova_subop
                  LSMW_OP['Entrada'][i] = (nova_subop/10) + 1 ###* testar
                else:
                  nova_subop = nova_subop + 10
                  LSMW_OP['Nº operação'][i] = nova_subop
                  LSMW_OP['Entrada'][i] = (nova_subop/10) + 1 ###* testar
        
        
              ### Alterar tempo total para cabeçalho não lubrificação
              if pd.isna(LSMW_OP['Suboperação'][i]) and 'LUB' not in str(LSMW_OP['Txt.breve operação'][i])[0:4]:
        
                ####  Somando tempo das ops a serem adicionadas
                LSMW_OP_un = LSMW_OP[LSMW_OP['A&D: ID externo da lista de tarefas'] == LSMW_OP['A&D: ID externo da lista de tarefas'][i]].reset_index(drop=True)
                soma_tempo_dn = 0
                soma_tempo_t = 0
                for t in range(len(LSMW_OP_un['Duração normal da operação'])):
                    if LSMW_OP_un['Txt.breve operação'][t] == cabecalho_lista_i:   # Pular tempo do cabeçalho
                      continue
                    tempo_int_dn = int(LSMW_OP_un['Duração normal da operação'][t])
                    tempo_int_t = int(LSMW_OP_un['Trabalho da operação'][t])
                    soma_tempo_dn += tempo_int_dn
                    soma_tempo_t += tempo_int_t
                ####  Somando tempo das ops já existentes
                LSMW_OP_un = SAP_TL[SAP_TL['CONCAT CENTRO_DESC'] == LSMW_OP['CONCAT CENTRO_DESC'][i]].reset_index(drop=True)
                for t in range(len(LSMW_OP_un['Duração normal'])):
                    if LSMW_OP_un['Txt.breve operação'][t] == cabecalho_lista_i:   # Pular tempo do cabeçalho
                      continue
                    tempo_int_dn = int(LSMW_OP_un['Duração normal'][t])
                    tempo_int_t = int(LSMW_OP_un['Trabalho'][t])
                    soma_tempo_dn += tempo_int_dn
                    soma_tempo_t += tempo_int_t
        
                LSMW_OP['Duração normal da operação'][i] = soma_tempo_dn
                LSMW_OP['Trabalho da operação'][i] = soma_tempo_t
        
        
          LSMW_OP_CAR = LSMW_OP[LSMW_OP['CARREGADO?'] == 1]
          LSMW_OP = LSMW_OP[LSMW_OP['CARREGADO?'] != 1]
        
        except:
          pass


        #####################

        import io
        buffer2 = io.BytesIO()

        # Salvando em arquivo excel

        with pd.ExcelWriter(buffer2, engine="xlsxwriter") as excel_writer:
            ## Crie um objeto ExcelWriter
            nome_arquivo_sap2 = 'CONCAT_LSMW_OP_'

            ## Salve cada DataFrame em uma planilha diferente
            LSMW_OP.to_excel(excel_writer, sheet_name='Operações e atividades', index=False)
            try:
                LSMW_OP_CAR.to_excel(excel_writer, sheet_name='CARREGADO OP', index=False)  # Listas já carregadas
            except:
                pass

            ## Feche o objeto ExcelWriter
            excel_writer.close()

            st.download_button(
                label="Download "+nome_arquivo_sap2,
                data=buffer2,
                file_name=nome_arquivo_sap2+'.xlsx',
            )

            ###
