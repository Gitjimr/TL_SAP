import numpy as np
import pandas as pd
from unidecode import unidecode

import streamlit as st
st.session_state.update(st.session_state)
for k, v in st.session_state.items():
    st.session_state[k] = v

hide_menu = '''
        <style>
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        </style>
'''
st.markdown(hide_menu, unsafe_allow_html=True)

st.markdown(
    """
    <style>
    .css-1jc7ptx, .e1ewe7hr3, .viewerBadge_container__1QSob,
    .styles_viewerBadge__1yB5_, .viewerBadge_link__1S137,
    .viewerBadge_text__1JaDK {
        display: none;
    }
    </style>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
    <style>
    .st-emotion-cache-ch5dnh
    {
        visibility: hidden;
    }
    .st-emotion-cache-q16mip
    {
        visibility: hidden;
    }
    .st-emotion-cache-ztfqz8
    {
        visibility: hidden;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# Função para realizar checagem: Exluir linhas completamente nulas, duplicatas e resetar o índice

def checagem_df(df):
    df = df.drop_duplicates()
    df = df.dropna(how='all')
    df = df.reset_index(drop=True)

    return df


#

# Função para salvar arquivo em .xlsx:

def gerar_xlsx(df, nome_arquivo):
    df.to_excel(nome_arquivo,
                index=False)  # O parâmetro 'index=False' evita a gravação do índice do DataFrame no arquivo


#


# Função para deixar a string com menos de 40 caracteres:

def cortar_string(string):
    palavras = string.split()  # Divide a string em uma lista de palavras

    i = 0  # Número de iterações até cortar a string

    # Loop até que a frase tenha menos de 40 caracteres
    while len(' '.join(palavras)) >= 40:
        maior_palavra = ''  # Variável para armazenar a palavra mais longa

        # Encontre a palavra mais longa
        for palavra in palavras:
            if len(palavra) > len(maior_palavra):
                maior_palavra = palavra

        # Reduza a palavra mais longa para 4 caracteres
        palavras[palavras.index(maior_palavra)] = palavras[palavras.index(maior_palavra)][:4]

        i += 1

        if i >= 10:
            # Reduza a palavra mais longa para 3 caracteres caso passe do limite de iterações
            palavras[palavras.index(maior_palavra)] = palavras[palavras.index(maior_palavra)][:3]
            i = 0

    # Agora que a frase tem menos de 40 caracteres, una as palavras novamente
    result_string = ' '.join(palavras)
    return result_string

prcnt_width = 80
max_width_str = f"max-width: {prcnt_width}%;"
st.markdown(f""" 
            <style> 
            .reportview-container .main .block-container{{{max_width_str}}}
            </style>    
            """,
            unsafe_allow_html=True)

#
import os

path = os.path.dirname(__file__)
my_path = path + '/pages/files/'
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


#   SETUP
#   CONCLUIR EDIÇÃO AQUI PARA PEGAR CERTO O ARQUIVO DE UPLOAD:
uploaded_file = st.file_uploader("Carregar planilha 'Op_padrao'")
if uploaded_file is not None and uploaded_file0 is not None:

    with st.spinner('Carregando Op_padrao...'):

        # Leitura do arquivo baixado

        ## Ler a tabela específica da planilha

        try:
            tabela = pd.read_excel( uploaded_file, sheet_name="EQUIPAMENTOS", skiprows=0, dtype=str)
        except:
            tabela = pd.read_excel( uploaded_file, sheet_name="EQUIPAMENTOS-AZ", skiprows=0, dtype=str)

        ## Ler a tabela op_padrao verificando a linha em que se encontra o cabeçalho

        skiprow = 0
        while 1:
            op_padrao = pd.read_excel( uploaded_file, sheet_name="op_padrao", skiprows=skiprow,
                                      dtype=str)
            lista_colunas_df_op = op_padrao.columns.tolist()
            if 'TEXTO TIPO EQUIPAMENTO' in lista_colunas_df_op:  # Identificar se é o cabeçalho
                break
            if skiprow > 10:
                st.write('ERRO: REVISAR POSIÇÃO DO CABEÇALHO DA TABELA')
                break
            skiprow += 1

        ## Transformando tabelas em DataFrame do Python

        df = pd.DataFrame(tabela)  # Criando df para a base de equipamentos
        # df = df[intervalo_planilha_equip[0]:intervalo_planilha_equip[1]]
        # df = df.head(500)    # teste
        # df = df.head(5117)    # BIS FF (M001)
        # df = df[6251:7310]    # WAF FF (M001)

        df_op = pd.DataFrame(op_padrao)  # Criando df para planilha das operações padrão para cada equipamento

        df_colunas = df.keys()
        df_colunas = df_colunas.tolist()

        if 'TIP. OBJETO_N7_procv' in df.columns:
            df = df.rename(columns={'TIP. OBJETO_N7_procv': 'TIP. OBJETO_N7'})


        ## Remover acentos do TEXTO DESCRITIVO, OPERAÇÃO PADRÃO E TASK LIST (df_op):

        def unidecode_texto(texto):
            if isinstance(texto, str):
                texto = unidecode(texto)  # Remover acentos, ç, etc
                texto = " ".join(texto.split())  # Remover espaços múltiplos
                text = texto.strip()  # Fazer strip para evitar espaços no fim (incluído 01/03/2024)
            return texto


        # Aplicar a função a todas as células do DataFrame
        df_op = df_op.applymap(unidecode_texto).loc[:, ~df_op.columns.str.contains('^Unnamed')]
        df_op = df_op[pd.notna(df_op['TEXTO TIPO EQUIPAMENTO'])].copy().reset_index(drop=True)
        df_op['OPERACAO PADRAO'] = df_op['OPERACAO PADRAO'].apply(
            lambda x: x.replace(';', ' ') if isinstance(x, str) else x)
        df_op['TEXTO DESCRITIVO'] = df_op['TEXTO DESCRITIVO'].apply(
            lambda x: x.replace(';', ' ') if isinstance(x, str) else x)
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
        ###


        num_carga = st.number_input(
            label='Número de Carga',
            min_value=0,
            step=1,
            key='num_carga',
            help="Último número de carga do tabelão das listas de tarefa no Sharepoint."
        )

    lista_nsubir = []

    #   CHECAR NA PLANILHA 'EQUIPAMENTOS' SE HÁ ALGUM COM STATUS 'MREL' OU SE NÃO EXISTE
    with st.spinner('Checando equipamentos inativos...'):
        try:
            MREL = SAP_EQP.loc[SAP_EQP['Status do sistema'].str.contains('|'.join(['MREL', 'INAT']))]
            MREL = MREL['Equipamento'].to_list()
            def verificar_equip_inativos(linha_df, df_colunas, MREL, SAP_EQP):
                if not pd.isna(linha_df['ID_SAP_N8']) and (
                        linha_df['ID_SAP_N8'] in MREL or linha_df['ID_SAP_N8'] not in SAP_EQP['Equipamento'].values):
                    linha_df[df_colunas.index('ID_SAP_N8'):] = np.nan

                if not pd.isna(linha_df['ID_SAP_N7']) and (
                        linha_df['ID_SAP_N7'] in MREL or linha_df['ID_SAP_N7'] not in SAP_EQP['Equipamento'].values):
                    linha_df[df_colunas.index('ID_SAP_N7'):] = np.nan

                if not pd.isna(linha_df['ID_SAP_N6']) and (
                        linha_df['ID_SAP_N6'] in MREL or linha_df['ID_SAP_N6'] not in SAP_EQP['Equipamento'].values):
                    linha_df[0:] = np.nan

                return linha_df
            df = df.apply(verificar_equip_inativos, axis=1, args=(df_colunas, MREL, SAP_EQP))
        except:
            pass


    # Função para checar variações com regras para "e's" e "ou's":
    with st.spinner('Link de equipamentos com planos...'):
        def check_variacao(desc, variacao):
            try:
                palavras_divididas = variacao.split()
                palavras_check = []

                for palavra in palavras_divididas:
                    if '+' in palavra:
                        palavras_check.append(
                            palavra.split('+'))  # Palavras unidas por '+' deverão estar ao mesmo tempo
                    else:
                        palavras_check.append(palavra)

                # Avaliar se pelo menos uma das palavras separadas por espaço está na descrição, considerando que palavras unidas por '+' deverão estar ao mesmo tempo
                resultado = any(
                    all(p in desc for p in item) if isinstance(item, list) else item in desc for item in palavras_check)

                return resultado  # True ou False
            except:
                return False


        #

        # Criando colunas de tipo de operação:

        if 'Tipo OP Padrão N6' not in df.columns or 'Tipo OP Padrão N7' not in df.columns or 'Tipo OP Padrão N8' not in df.columns:
            df.insert(loc=df.keys().tolist().index('EQUIPAMENTO PRINCIPAL'), column='Tipo OP Padrão N6', value=np.nan)
            df.insert(loc=df.keys().tolist().index('SISTEMA FUNCIONAL / CONJUNTO'), column='Tipo OP Padrão N7',
                      value=np.nan)
            df.insert(loc=df.keys().tolist().index('EQUIPAMENTO FUNCIONAL / SUB-CONJUNTO'), column='Tipo OP Padrão N8',
                      value=np.nan)

            ## Zerando tipos de operação, caso já preenchidos, a fim de preenchê-los novamente:

            df['Tipo OP Padrão N6'] = np.nan
            df['Tipo OP Padrão N7'] = np.nan
            df['Tipo OP Padrão N8'] = np.nan

        else:
            df['Tipo OP Padrão N6'] = [i.strip() if isinstance(i, str) else i for i in df['Tipo OP Padrão N6']]
            df['Tipo OP Padrão N7'] = [i.strip() if isinstance(i, str) else i for i in df['Tipo OP Padrão N7']]
            df['Tipo OP Padrão N8'] = [i.strip() if isinstance(i, str) else i for i in df['Tipo OP Padrão N8']]

            df['Tipo OP Padrão N6'] = df['Tipo OP Padrão N6'].apply(lambda x: unidecode(x) if isinstance(x, str) else x)
            df['Tipo OP Padrão N7'] = df['Tipo OP Padrão N7'].apply(lambda x: unidecode(x) if isinstance(x, str) else x)
            df['Tipo OP Padrão N8'] = df['Tipo OP Padrão N8'].apply(lambda x: unidecode(x) if isinstance(x, str) else x)

        ###

        # Lógica de classificação:
        for i in range(len(df)):
            # print(i)
            n4 = str(df['LINHAS / DIAG / SUB PROCESS'][i])
            equipamento_principal = str(df['EQUIPAMENTO PRINCIPAL'][i])
            sistema_funcional = str(df['SISTEMA FUNCIONAL / CONJUNTO'][i])
            equipamento_funcional = str(df['EQUIPAMENTO FUNCIONAL / SUB-CONJUNTO'][i])
            sistemas_etapas_process = str(df['SISTEMAS / ETAPAS PROCESS'][i])

            ## N6
            if isinstance(equipamento_principal, str) and not isinstance(df['Tipo OP Padrão N6'][i],
                                                                         str):  # Saber se já está classificado
                obj_n6 = df['TIP. OBJETO_N6'][i]
                desc = equipamento_principal  # descrição equipamento para checar variação
                n4n5 = sistemas_etapas_process + " " + n4  # descrição da etapa (n5) + n4 para checar variação

                for j in range(len(df_op)):
                    tipo_equipamento = str(df_op['TIPO EQUIPAMENTO'][j])
                    variacao_desc = df_op['VARIACAO DESC'][j]  # Variações para a descrição do equipamento ou da etapa
                    variacao_n4n5 = df_op['VARIACAO N4/N5'][j]  # Variações para a descrição do N4

                    if isinstance(tipo_equipamento, str) and isinstance(obj_n6, str):
                        if tipo_equipamento in obj_n6:
                            if check_variacao(desc, variacao_desc) and check_variacao(n4n5, variacao_n4n5):
                                df.loc[i, 'Tipo OP Padrão N6'] = df_op['TEXTO TIPO EQUIPAMENTO'][j]
                            elif check_variacao(desc, variacao_desc) and not isinstance(variacao_n4n5, str):
                                df.loc[i, 'Tipo OP Padrão N6'] = df_op['TEXTO TIPO EQUIPAMENTO'][j]
                            elif check_variacao(n4n5, variacao_n4n5) and not isinstance(variacao_desc, str):
                                df.loc[i, 'Tipo OP Padrão N6'] = df_op['TEXTO TIPO EQUIPAMENTO'][j]
                            elif not isinstance(variacao_desc, str) and not isinstance(variacao_n4n5, str):
                                df.loc[i, 'Tipo OP Padrão N6'] = df_op['TEXTO TIPO EQUIPAMENTO'][j]

            ## N7
            if isinstance(sistema_funcional, str) and not isinstance(df['Tipo OP Padrão N7'][i],
                                                                     str):  # Saber se já está classificado
                obj_n7 = df['TIP. OBJETO_N7'][i]
                desc = sistema_funcional + " " + equipamento_principal  # descrição do sistema + equipamento para checar variação
                n4n5 = sistemas_etapas_process + " " + n4  # descrição da etapa (n5) + n4 para checar variação

                for j in range(len(df_op)):
                    tipo_equipamento = str(df_op['TIPO EQUIPAMENTO'][j])
                    variacao_desc = df_op['VARIACAO DESC'][j]
                    variacao_n4n5 = df_op['VARIACAO N4/N5'][j]

                    if isinstance(tipo_equipamento, str) and isinstance(obj_n7, str):
                        if tipo_equipamento in obj_n7:
                            if check_variacao(desc, variacao_desc) and check_variacao(n4n5, variacao_n4n5):
                                df.loc[i, 'Tipo OP Padrão N7'] = df_op['TEXTO TIPO EQUIPAMENTO'][j]
                            elif check_variacao(desc, variacao_desc) and not isinstance(variacao_n4n5, str):
                                df.loc[i, 'Tipo OP Padrão N7'] = df_op['TEXTO TIPO EQUIPAMENTO'][j]
                            elif check_variacao(n4n5, variacao_n4n5) and not isinstance(variacao_desc, str):
                                df.loc[i, 'Tipo OP Padrão N7'] = df_op['TEXTO TIPO EQUIPAMENTO'][j]
                            elif not isinstance(variacao_desc, str) and not isinstance(variacao_n4n5, str):
                                df.loc[i, 'Tipo OP Padrão N7'] = df_op['TEXTO TIPO EQUIPAMENTO'][j]

            ## N8
            if isinstance(equipamento_funcional, str) and not isinstance(df['Tipo OP Padrão N8'][i],
                                                                         str):  # Saber se já está classificado
                obj_n8 = df['TIP. OBJETO_N8'][i]
                desc = equipamento_funcional + " " + equipamento_principal  # descrição dos equipamentos funcional + principal para checar variação
                n4n5 = sistemas_etapas_process + " " + n4  # descrição da etapa (n5) + n4 para checar variação

                for j in range(len(df_op)):
                    tipo_equipamento = str(df_op['TIPO EQUIPAMENTO'][j])
                    variacao_desc = df_op['VARIACAO DESC'][j]
                    variacao_n4n5 = df_op['VARIACAO N4/N5'][j]

                    if isinstance(tipo_equipamento, str) and isinstance(obj_n8, str):
                        if tipo_equipamento in obj_n8:
                            if check_variacao(desc, variacao_desc) and check_variacao(n4n5, variacao_n4n5):
                                df.loc[i, 'Tipo OP Padrão N8'] = df_op['TEXTO TIPO EQUIPAMENTO'][j]
                            elif check_variacao(desc, variacao_desc) and not isinstance(variacao_n4n5, str):
                                df.loc[i, 'Tipo OP Padrão N8'] = df_op['TEXTO TIPO EQUIPAMENTO'][j]
                            elif check_variacao(n4n5, variacao_n4n5) and not isinstance(variacao_desc, str):
                                df.loc[i, 'Tipo OP Padrão N8'] = df_op['TEXTO TIPO EQUIPAMENTO'][j]
                            elif not isinstance(variacao_desc, str) and not isinstance(variacao_n4n5, str):
                                df.loc[i, 'Tipo OP Padrão N8'] = df_op['TEXTO TIPO EQUIPAMENTO'][j]

        ###

        # *-*-*-*-OK ACIMA

    with st.spinner('Associando planos a equipamentos...'):
        # Criando df de equipamentos sem planos atribuídos

        df_n6splano = df[df['Tipo OP Padrão N6'].isna()]
        df_n6splano = df_n6splano[['EQUIPAMENTO PRINCIPAL', 'TIP. OBJETO_N6']]
        df_n6splano = checagem_df(df_n6splano)
        df_n6splano = df_n6splano.sort_values(by='TIP. OBJETO_N6')

        df_n8splano = df[df['Tipo OP Padrão N8'].isna()]
        df_n8splano = df_n8splano[['EQUIPAMENTO FUNCIONAL / SUB-CONJUNTO', 'TIP. OBJETO_N8']]
        df_n8splano = checagem_df(df_n8splano)
        df_n8splano = df_n8splano.sort_values(by='TIP. OBJETO_N8')

        #

        df_class = df

        # Pegando todas as colunas da planilha de operações, menos "TIPO DE EQUIPAMENTO", que é repetida da base de equipamentos
        df_op_colunas = df_op.columns.values
        df_op_colunas = df_op_colunas[
                        1:]  # Pegando todas as colunas da planilha de operações, menos "TIPO DE EQUIPAMENTO", que é repetida
        df_op_colunas = df_op_colunas.tolist()
        #

        # Lista de colunas a serem adicionadas na base de equipamentos lida contidas na planilha de operações padrão
        colunas_para_adicionar = ['EXCLUIR?']  # Adicionando coluna "EXCLUIR?" para operações posteriores
        colunas_para_adicionar = colunas_para_adicionar + sorted(df_op_colunas, key=df_op_colunas.index, reverse=True)
        #

        # Índice onde você deseja inserir as novas colunas (após a última coluna)
        # indice_inserir_coluna = df.columns.get_loc('EQUIPAMENTO FUNCIONAL / SUB-CONJUNTO') + 1
        indice_inserir_coluna = df.columns.get_loc(df_colunas[-1]) + 1
        #

        # Adiciona as colunas se não existirem e atribuir, inicialmente, o valor "NaN" a elas
        for coluna in colunas_para_adicionar:
            if coluna not in df.columns:
                df.insert(loc=indice_inserir_coluna, column=coluna, value=np.nan)
        #
        ################################
        #              N8
        ################################

        dict_nova_linha_N8 = {}
        dict_nova_linha_N8 = pd.DataFrame(dict_nova_linha_N8)

        # Loop para atualizar o DataFrame com base em condições
        for i in range(len(df['LI_N5'])):
            tipo_op = df['Tipo OP Padrão N8'][i]
            # Verificar se 'Tipo OP Padrão N8' não está vazio antes de usar o 'state_df'
            if tipo_op:
                # Pegar todas as linhas de operação padrão para cada linha do df de equipamentos
                state_df = df_op[df_op['TEXTO TIPO EQUIPAMENTO'] == str(df['Tipo OP Padrão N8'][i])]
                # Adicionar valores da linha 'i' do DataFrame 'df' ao 'state_df'
                for coluna in df.columns.difference(df_op.columns):
                    state_df[coluna] = df.at[i, coluna]
                # Reorganizar as colunas de 'state_df' na ordem de 'df'
                state_df = state_df[df.columns]
                # Adicionar 'state_df' para 'i' atual ao dataframe de linhas a serem adicionadas no final
                # dict_nova_linha_N8 = dict_nova_linha_N8.append(state_df, ignore_index=True)
                dict_nova_linha_N8 = pd.concat([dict_nova_linha_N8, state_df], ignore_index=True, sort=False)
                # Atribuir 'EXCLUIR' para o item 'i' do df de equipamentos, sem as operações
                df.at[i, 'EXCLUIR?'] = 'EXCLUIR'

        dict_nova_linha_N8['Tipo OP Padrão N6'] = np.nan
        dict_nova_linha_N8['Tipo OP Padrão N7'] = np.nan

        # df_operacoes_i = pd.concat([df_operacoes_i , ], ignore_index=True, sort=False)

        ################################
        #              N7
        ################################

        dict_nova_linha_N7 = {}
        dict_nova_linha_N7 = pd.DataFrame(dict_nova_linha_N7)

        # Loop para atualizar o DataFrame com base em condições
        for i in range(len(df['LI_N5'])):
            tipo_op = df['Tipo OP Padrão N7'][i]
            # Verificar se 'Tipo OP Padrão N7' não está vazio antes de usar str.startswith
            if tipo_op:
                # Pegar todas as linhas de operação padrão para cada linha do df de equipamentos
                state_df = df_op[df_op['TEXTO TIPO EQUIPAMENTO'] == str(df['Tipo OP Padrão N7'][i])]
                # Adicionar valores da linha 'i' do DataFrame 'df' ao 'state_df'
                for coluna in df.columns.difference(df_op.columns):
                    state_df[coluna] = df.at[i, coluna]
                # Reorganizar as colunas de 'state_df' na ordem de 'df'
                state_df = state_df[df.columns]
                # Adicionar 'state_df' para 'i' atual ao dataframe de linhas a serem adicionadas no final
                # dict_nova_linha_N7 = dict_nova_linha_N7.append(state_df, ignore_index=True)
                dict_nova_linha_N7 = pd.concat([dict_nova_linha_N7, state_df], ignore_index=True, sort=False)
                # Atribuir 'EXCLUIR' para o item 'i' do df de equipamentos, sem as operações
                df.at[i, 'EXCLUIR?'] = 'EXCLUIR'

        for coluna in dict_nova_linha_N7:
            try:
                if 'N8' in coluna or 'EQUIPAMENTO FUNCIONAL' in coluna:  # MUDEI DE EQUIPAMENTO PRINCIPAL -> EQUIPAMENTO FUNCIONAL
                    dict_nova_linha_N7[coluna] = np.nan
                    for col in range(len(df_op_colunas)):
                        for linha in range(len(dict_nova_linha_N7[col])):
                            dict_nova_linha_N7[col][linha] = np.nan
            except:
                continue

        dict_nova_linha_N7['Tipo OP Padrão N6'] = np.nan
        dict_nova_linha_N7['Tipo OP Padrão N8'] = np.nan

        ################################
        #              N6
        ################################

        dict_nova_linha_N6 = {}
        dict_nova_linha_N6 = pd.DataFrame(dict_nova_linha_N6)

        # Loop para atualizar o DataFrame com base em condições
        for i in range(len(df['LI_N5'])):
            tipo_op = df['Tipo OP Padrão N6'][i]
            # Verificar se 'Tipo OP Padrão N6' não está vazio antes de usar str.startswith
            if tipo_op:
                # Pegar todas as linhas de operação padrão para cada linha do df de equipamentos
                state_df = df_op[df_op['TEXTO TIPO EQUIPAMENTO'] == str(df['Tipo OP Padrão N6'][i])]
                # Adicionar valores da linha 'i' do DataFrame 'df' ao 'state_df'
                for coluna in df.columns.difference(df_op.columns):
                    state_df[coluna] = df.at[i, coluna]
                # Reorganizar as colunas de 'state_df' na ordem de 'df'
                state_df = state_df[df.columns]
                # Adicionar 'state_df' para 'i' atual ao dataframe de linhas a serem adicionadas no final
                # dict_nova_linha_N6 = dict_nova_linha_N6.append(state_df, ignore_index=True)
                dict_nova_linha_N6 = pd.concat([dict_nova_linha_N6, state_df], ignore_index=True, sort=False)
                # Atribuir 'EXCLUIR' para o item 'i' do df de equipamentos, sem as operações
                df.at[i, 'EXCLUIR?'] = 'EXCLUIR'

        for coluna in dict_nova_linha_N6:
            try:
                if 'N7' in coluna or 'N8' in coluna or 'SISTEMA FUNCIONAL' in coluna or 'EQUIPAMENTO FUNCIONAL' in coluna:
                    print(coluna)
                    dict_nova_linha_N6[coluna] = np.nan
                    for col in range(len(df_op_colunas)):
                        for linha in range(len(dict_nova_linha_N6[col])):
                            dict_nova_linha_N6[col][linha] = np.nan
            except:
                continue

        dict_nova_linha_N6['Tipo OP Padrão N7'] = np.nan
        dict_nova_linha_N6['Tipo OP Padrão N8'] = np.nan

        ################################

        ##  Excluir linhas em que há "EXCLUIR" no dataframe original

        for i in range(len(df['LI_N5'])):
            if df['EXCLUIR?'][i] == 'EXCLUIR':
                df = df.drop(i)

        ###

        ## Checagem

        df = checagem_df(df)

        ##

        ## Adicionar linhas criadas do dicionário à tabela

        # df = df.append(dict_nova_linha_N6, ignore_index=True)
        df = pd.concat([df, dict_nova_linha_N6], ignore_index=True, sort=False)
        # df = df.append(dict_nova_linha_N7, ignore_index=True)
        df = pd.concat([df, dict_nova_linha_N7], ignore_index=True, sort=False)
        # df = df.append(dict_nova_linha_N8, ignore_index=True)
        df = pd.concat([df, dict_nova_linha_N8], ignore_index=True, sort=False)

        ###

        df = df.drop('EXCLUIR?', axis=1)  # Deletando coluna 'EXCLUIR?'


    with st.spinner('Tratando dados...'):

        # Adiciona a coluna concatenando "OPERACAO PADRAO" + "TEXTO DESCRITIVO"
        try:
            df.insert(loc=df.keys().tolist().index('TEXTO DESCRITIVO') + 1, column='OP: TEXTO DESCRITIVO', value=np.nan)

            for i in range(len(df['OPERACAO PADRAO'])):
                if isinstance(df['OPERACAO PADRAO'][i], str) and isinstance(df['TEXTO DESCRITIVO'][i], str):
                    # df['OP: TEXTO DESCRITIVO'][i] = df['OPERACAO PADRAO'][i]+': '+df['TEXTO DESCRITIVO'][i]
                    df['OP: TEXTO DESCRITIVO'][i] = ': ' + df['TEXTO DESCRITIVO'][i]
                # elif isinstance(df['OPERACAO PADRAO'][i], str):
                # df['OP: TEXTO DESCRITIVO'][i] = df['OPERACAO PADRAO'][i]
                elif isinstance(df['TEXTO DESCRITIVO'][i], str):
                    df['OP: TEXTO DESCRITIVO'][i] = df['TEXTO DESCRITIVO'][i]

        except:
            st.write('TEXTO DESCRITIVO' in df)
            st.write('OPERACAO PADRAO' in df)
            df.loc[-1] = 'ERRO: CONCATENAÇÃO DE TEXTO OPERAÇÃO + TEXTO DESCRITIVO'
            st.write('ERRO: CONCATENAÇÃO DE TEXTO OPERAÇÃO + TEXTO DESCRITIV')

        #

        # Adiciona a coluna da "TASK LIST" efetiva
        try:
            df.insert(loc=df.keys().tolist().index('TASK LIST_PARCIAL') + 1, column='TASK LIST',
                      value=df['TASK LIST_PARCIAL'] + ' ' + df['LINHAS / DIAG / SUB PROCESS'])

            ## Adicionar N4 para caso de 'ROTA'='SIM'
            for i in range(len(df['ROTA?'])):
                if str(df['ROTA?'][i]) == 'SIM':
                    df['TASK LIST'][i] = df['TASK LIST_PARCIAL'][i] + ' ' + df['DESC SISTEMAS / ETAPAS PROCESS'][
                        i] + ' ' + df['LINHAS / DIAG / SUB PROCESS'][i]
                if str(df['ROTA?'][i]) == 'PERSONALIZADO':
                    df['TASK LIST'][i] = df['TASK LIST_PARCIAL'][i]
            ##

            df = df.drop('TASK LIST_PARCIAL', axis=1)

        except:
            st.write('LINHAS / DIAG / SUB PROCESS' in df)
            st.write('DESC SISTEMAS / ETAPAS PROCESS' in df)
            df.loc[-1] = 'ERRO: CRIAÇÃO DA TASK LIST EFETIVA'
            st.write('ERRO: CRIAÇÃO DA TASK LIST EFETIVA')


        df['TASK LIST'].str.replace('REF', 'REF', regex=True)

        # TRATAMENTO DE DADOS FINAL:

        ## Lista de colunas que vieram com a tabela a serem excluídas:

        col_excluir = ['#CHAR', '#CHAR2', '#COD_MATERIAL', '#QTD_MATERIAL', 'TAG_IDENT_N', '#CARCTR N6',
                       'FABRICANTE N6', 'MODELO', 'ANO', 'TAG_IDENT_N7', '#CARCTR N7', 'TIPO EQUIPAMENTO',
                       'VARIACAO DESC', 'VARIACAO N4/N5', 'TASK LIST_TRECHO 01', 'TRECHO 02_TASK LIST', '#CHAR.1']
        for col in col_excluir:
            try:
                df = df.drop(col, axis=1)
            except:
                continue

        ##

        ## Checagem final:

        df = checagem_df(df)

        ##

        ## Abreviar coluna 'TASK LIST' se maior que 40 caracteres:

        for i in range(len(df['TASK LIST'])):
            if isinstance(df['TASK LIST'][i], str):
                if len(df['TASK LIST'][i]) > 40:
                    df['TASK LIST'][i] = df['TASK LIST'][i].replace('.', '')  # Remove pontos
                    df['TASK LIST'][i] = df['TASK LIST'][i].replace(' N0', ' ')
                    df['TASK LIST'][i] = df['TASK LIST'][i].replace(' DE ', ' ')
                    df['TASK LIST'][i] = df['TASK LIST'][i].replace(' DA ', ' ')
                    df['TASK LIST'][i] = df['TASK LIST'][i].replace(' DO ', ' ')
                    df['TASK LIST'][i] = df['TASK LIST'][i].replace('  ', ' ')  # Remove espaços supérfluos
                if len(df['TASK LIST'][i]) > 40:
                    df['TASK LIST'][i] = df['TASK LIST'][i].replace(' RESFRIAMENTO ', ' RESF ')
                    df['TASK LIST'][i] = df['TASK LIST'][i].replace(' RESFRIADOR ', ' RESF ')
                    df['TASK LIST'][i] = df['TASK LIST'][i].replace(' REFRIGERACAO ', ' REFRIG ')
                    df['TASK LIST'][i] = df['TASK LIST'][i].replace(' EMBALAGEM ', ' EMB ')
                    df['TASK LIST'][i] = df['TASK LIST'][i].replace(' EMBALADORA ', ' EMB ')
                    df['TASK LIST'][i] = df['TASK LIST'][i].replace(' EMBALADOR ', ' EMB ')
                    df['TASK LIST'][i] = df['TASK LIST'][i].replace(' ENCAIXOTAMENTO ', ' ENCX ')
                    df['TASK LIST'][i] = df['TASK LIST'][i].replace(' ENCAIXOTADOR ', ' ENCX ')
                    df['TASK LIST'][i] = df['TASK LIST'][i].replace(' ENCAIXOTADORA ', ' ENCX ')
                    df['TASK LIST'][i] = df['TASK LIST'][i].replace(' EMPILHAMENTO ', ' EMPI ')
                    df['TASK LIST'][i] = df['TASK LIST'][i].replace(' APLICACAO ', ' APL ')
                    df['TASK LIST'][i] = df['TASK LIST'][i].replace(' PREPARO ', ' PREP ')
                    df['TASK LIST'][i] = df['TASK LIST'][i].replace(' PREPARACAO ', ' PREP ')
                    df['TASK LIST'][i] = df['TASK LIST'][i].replace(' BISCOITO ', ' BISC ')
                    df['TASK LIST'][i] = df['TASK LIST'][i].replace(' WAFER ', ' WAF ')
                    df['TASK LIST'][i] = df['TASK LIST'][i].replace(' TRANSPORTE ', ' TRANSP ')
                    df['TASK LIST'][i] = df['TASK LIST'][i].replace(' TRANSFERENCIA ', ' TRANSF ')
                    df['TASK LIST'][i] = df['TASK LIST'][i].replace(' ARMAZENAGEM ', ' ARMZ ')
                    df['TASK LIST'][i] = df['TASK LIST'][i].replace(' ALIMENTACAO ', ' ALIM ')
                    df['TASK LIST'][i] = df['TASK LIST'][i].replace(' SEPARACAO ', ' SEP ')
                    df['TASK LIST'][i] = df['TASK LIST'][i].replace(' RECEBIMENTO ', ' RECEB ')
                    df['TASK LIST'][i] = df['TASK LIST'][i].replace(' MARGARINA ', ' MARG ')
                    df['TASK LIST'][i] = df['TASK LIST'][i].replace(' DISTRIBUICAO ', ' DISTRIB ')
                    df['TASK LIST'][i] = df['TASK LIST'][i].replace(' GERACAO ', ' GER ')
                    df['TASK LIST'][i] = df['TASK LIST'][i].replace(' COMPRIMIDO ', ' COMP ')
                    df['TASK LIST'][i] = df['TASK LIST'][i].replace('  ', ' ')  # Remove espaços supérfluos
        ##

    with st.spinner('Gerando planilha de listas de tarefas...'):
        if 'EQUIPAMENTO PRINCIPAL OG' in df.columns:
            df['EQUIPAMENTO PRINCIPAL'] = df['EQUIPAMENTO PRINCIPAL OG']

        df_tl = {
            'LI_N3': [],
            'OS': [],  # Conta e diferencia ordens de manutenção diferentes
            'ROTA?': [],
            'TASK LIST': [],
            'ID N6/N5': [],
            'ID SAP N6/N5': [],
            'N6/N5': [],
            'NÚMERO OP': [],
            'OPERAÇÃO: TEXTO CURTO': [],
            'ID EQUIP OP': [],
            'ID SAP EQUIP OP': [],
            'EQUIPAMENTO OP': [],
            'OPERAÇÃO: TEXTO LONGO': [],
            'DURAÇÃO (min)': [],
            'CT PM': [],
            'ESTADO MAQ': [],
            'ID SAP N6': []
        }

        df_tl['LI_N3'] = df['LI_N3']
        df_tl['TASK LIST'] = df['TASK LIST']
        df_tl['NÚMERO OP'] = np.nan
        df_tl['OPERAÇÃO: TEXTO CURTO'] = df['OPERACAO PADRAO']
        df_tl['N6/N5'] = np.nan
        df_tl['OPERAÇÃO: TEXTO LONGO'] = df['OP: TEXTO DESCRITIVO']
        df_tl['DURAÇÃO (min)'] = df['DUR_NORMAL_MIN']
        df_tl['OPERADORES'] = df['QTD']
        df_tl['ROTA?'] = df['ROTA?']
        df_tl['ID N6/N5'] = np.nan
        df_tl['ID SAP N6/N5'] = np.nan
        df_tl['ID EQUIP OP'] = np.nan
        df_tl['ID SAP EQUIP OP'] = np.nan
        df_tl['EQUIPAMENTO OP'] = np.nan
        df_tl['OS'] = np.nan
        df_tl['CT PM'] = df['CT PM']
        df_tl['ESTADO MAQ'] = df['ESTADO MAQ']
        df_tl['ID SAP N6'] = df['ID_SAP_N6']

        df_tl = pd.DataFrame(df_tl)

        df_tl['DURAÇÃO (min)'] = pd.to_numeric(df_tl['DURAÇÃO (min)'], errors='coerce')

        # FIZ ESSE CÓDIGO PARA PADRONIZAR A COLUNA 'EQUIPAMENTO' DO PADRÃO DE TABELA SURGERIDO PARA O PM VISTO QUE O EXCEL ESTAVA DEMORANDO MUITO

        # Preenchendo colunas 'EQUIPAMENTO':

        for i in range(len(df['ROTA?'])):

            if df['ROTA?'][i] == 'SIM':
                df_tl['ID N6/N5'][i] = df['LI_N5'][i]
                df_tl['ID SAP N6/N5'][i] = np.nan
                df_tl['N6/N5'][i] = df['DESC SISTEMAS / ETAPAS PROCESS'][i]
                if isinstance(df['Tipo OP Padrão N6'][i], str):
                    df_tl['ID EQUIP OP'][i] = df['NR_TECNICO_N6'][i]
                    df_tl['ID SAP EQUIP OP'][i] = df['ID_SAP_N6'][i]
                    df_tl['EQUIPAMENTO OP'][i] = df['EQUIPAMENTO PRINCIPAL'][i]
                elif isinstance(df['Tipo OP Padrão N7'][i], str):
                    df_tl['ID EQUIP OP'][i] = df['NR_TECNICO_N7'][i]
                    df_tl['ID SAP EQUIP OP'][i] = df['ID_SAP_N7'][i]
                    df_tl['EQUIPAMENTO OP'][i] = df['SISTEMA FUNCIONAL / CONJUNTO'][i]
                elif isinstance(df['Tipo OP Padrão N8'][i], str):
                    df_tl['ID EQUIP OP'][i] = df['NR_TECNICO_N8'][i]
                    df_tl['ID SAP EQUIP OP'][i] = df['ID_SAP_N8'][i]
                    df_tl['EQUIPAMENTO OP'][i] = df['EQUIPAMENTO FUNCIONAL / SUB-CONJUNTO'][i]

            else:
                df_tl['ID N6/N5'][i] = df['NR_TECNICO_N6'][i]
                df_tl['ID SAP N6/N5'][i] = df['ID_SAP_N6'][i]
                df_tl['N6/N5'][i] = df['EQUIPAMENTO PRINCIPAL'][i]
                df_tl['ID EQUIP OP'][i] = np.nan
                df_tl['EQUIPAMENTO OP'][i] = np.nan
                if isinstance(df['Tipo OP Padrão N6'][i], str):
                    df_tl['ID EQUIP OP'][i] = df['NR_TECNICO_N6'][i]
                    df_tl['ID SAP EQUIP OP'][i] = df['ID_SAP_N6'][i]
                    df_tl['EQUIPAMENTO OP'][i] = df['EQUIPAMENTO PRINCIPAL'][i]
                elif isinstance(df['Tipo OP Padrão N7'][i], str):
                    df_tl['ID EQUIP OP'][i] = df['NR_TECNICO_N7'][i]
                    df_tl['ID SAP EQUIP OP'][i] = df['ID_SAP_N7'][i]
                    df_tl['EQUIPAMENTO OP'][i] = df['SISTEMA FUNCIONAL / CONJUNTO'][i]
                elif isinstance(df['Tipo OP Padrão N8'][i], str):
                    df_tl['ID EQUIP OP'][i] = df['NR_TECNICO_N8'][i]
                    df_tl['ID SAP EQUIP OP'][i] = df['ID_SAP_N8'][i]
                    df_tl['EQUIPAMENTO OP'][i] = df['EQUIPAMENTO FUNCIONAL / SUB-CONJUNTO'][i]


        #
        
        # LOCAL DE INSTALAÇÃO PARA ROTA PERSONALIZADA:
        
        for i in range(len(df['ROTA?'])):   # *********** ADICIONADO 24/05/2024
        
          if df['ROTA?'][i] == 'PERSONALIZADO':
            df__copia = df[df['TASK LIST'] == df_tl['TASK LIST'][i]].copy()
            print(df__copia)
            if len(df__copia.drop_duplicates( subset = ['LI_N5'] ).reset_index(drop = True)) == 1:
              df_tl['ID N6/N5'][i] = df['LI_N5'][i]
            elif len(df__copia.drop_duplicates( subset = ['LI_N4'] ).reset_index(drop = True)) == 1:
              df_tl['ID N6/N5'][i] = df['LI_N4'][i]
            elif len(df__copia.drop_duplicates( subset = ['LI_N3'] ).reset_index(drop = True)) == 1:
              df_tl['ID N6/N5'][i] = df['LI_N3'][i]
            elif len(df__copia.drop_duplicates( subset = ['LI_N2'] ).reset_index(drop = True)) == 1:
              df_tl['ID N6/N5'][i] = df['LI_N2'][i]
            else:
              df_tl['ID N6/N5'][i] = 'ERRO: EQUIPAMENTOS DE SETORES (IND, MOI, ETC) DIFERENTES PRESENTES NA MESMA ROTA.'
        #
        
        # Ordenando 'TASK LIST' em ordem alfabética:

        df_tl = df_tl.sort_values(by=['TASK LIST', 'N6/N5', 'ID SAP N6'])

        df_tl = df_tl.reset_index(drop=True)

        #

        df_tl = checagem_df(df_tl)

        # Preenchendo colunas 'NÚMERO OP' e 'OS':

        os = 1
        for i in range(len(df_tl['ROTA?'])):

            if i == 0:
                numero_op = 10

            if i > 0:
                if df_tl['TASK LIST'][i] != df_tl['TASK LIST'][i - 1] or df_tl['N6/N5'][i] != df_tl['N6/N5'][i - 1] or \
                        df_tl['LI_N3'][i] != df_tl['LI_N3'][i - 1]:
                    numero_op = 10
                    os = os + 1
                else:
                    numero_op = numero_op + 10

            df_tl['OS'][i] = os

            if len('00' + str(numero_op)) == 4:
                df_tl['NÚMERO OP'][i] = '00' + str(numero_op)
            elif len('00' + str(numero_op)) == 5:
                df_tl['NÚMERO OP'][i] = '0' + str(numero_op)
            elif len('00' + str(numero_op)) == 6:
                df_tl['NÚMERO OP'][i] = str(numero_op)
        #

        # Quantidade de operações e duração

        top_num_op = df_tl.groupby(['TASK LIST', 'N6/N5', 'LI_N3'])['NÚMERO OP'].count().reset_index()
        top_num_op = top_num_op.sort_values(by='NÚMERO OP', ascending=False).head(10)

        top_duracao = df_tl.groupby(['TASK LIST', 'N6/N5', 'LI_N3'])['DURAÇÃO (min)'].sum().reset_index()
        top_duracao = top_duracao.sort_values(by='DURAÇÃO (min)', ascending=False).head(10)

        #

        # Criando df de planos que não foram atribuídos a nenhum equipamento

        df_planosobrando = {'PLANO': []}

        lista_planos_df = list(set(df['Tipo OP Padrão N6'])) + list(set(df['Tipo OP Padrão N7'])) + list(
            set(df['Tipo OP Padrão N8']))
        lista_planos_df = [x for x in lista_planos_df if str(x) != 'nan']
        planos_nao_atribuidos = set(df_op['TEXTO TIPO EQUIPAMENTO']) - set(lista_planos_df)
        df_planosobrando['PLANO'] = list(planos_nao_atribuidos)

        # Criando o DataFrame df_planosobrando
        df_planosobrando = pd.DataFrame(df_planosobrando)
        df_planosobrando = checagem_df(df_planosobrando)
        #


        import datetime
        from datetime import datetime
        data_hoje = str(datetime.today().strftime('%d.%m.%Y'))


        #   Preenchendo CABECALHO DOS PLANOS DE MANUTENCAO

        df_cabecalho_eqp = {
            'Chave do grupo de listas de tarefas*': [],
            'Contador de grupos*': [],
            'Data de criação registro': [],
            'Data de início da validade': [],
            'Descrição': [],
            'Centro de planejamento*': [],
            'Centro de trabalho': [],
            'Centro de centro de trabalho': [],
            'Utilização lista de tarefas': [],
            'Status global': [],
            'Estratégia de manutenção': [],
            'Grupo de planejamento': [],
            'Condições da instalação': [],
            'Conjunto': [],
            'Txt.descr.cabeçalho': [],
            'Numeração Externa': [],
            'Equipamento': [],
            'ID_SAP': [],
            'Local de instalação': [],
            'LI_N3': []
        }
        df_1 = df_tl.copy()   # *********** ALTERADO 24/05/2024
        df_1 = df_1.drop_duplicates( subset = ['TASK LIST','EQUIPAMENTO OP'], keep = 'last').reset_index(drop = True)    # ALTERADO: O CABEÇALHO PUXAVA SÓ O EQP PRINCIPAL
        # df_1 = df_tl.drop_duplicates( subset = ['TASK LIST','N6/N5'], keep = 'last')
        # df_1 = df_1.sort_values(by=['Índices'])
        # df_1 = df_1.reset_index(drop=True)

        indice = 0
        for i in range(len(df_1['LI_N3'])):
            if i > 0 and df_1['TASK LIST'][i] == df_1['TASK LIST'][i - 1] and df_1['ID SAP EQUIP OP'][i] == \
                    df_1['ID SAP EQUIP OP'][i - 1]:
                continue

            df_cabecalho_eqp['Chave do grupo de listas de tarefas*'].append(num_carga + indice)
            df_cabecalho_eqp['Data de criação registro'].append(data_hoje)
            df_cabecalho_eqp['Data de início da validade'].append('01.12.2023')
            df_cabecalho_eqp['Contador de grupos*'].append(1)
            df_cabecalho_eqp['Utilização lista de tarefas'].append(4)
            df_cabecalho_eqp['Status global'].append(4)
            df_cabecalho_eqp['Condições da instalação'].append(1 if 'FUNC' in df_1['TASK LIST'][i] else 0)
            df_cabecalho_eqp['Descrição'].append(df_1['TASK LIST'][i])
            df_cabecalho_eqp['Centro de planejamento*'].append(df_1['LI_N3'][i][0:4])
            df_cabecalho_eqp['Centro de trabalho'].append(df_1['CT PM'][i])
            df_cabecalho_eqp['Centro de centro de trabalho'].append(df_1['LI_N3'][i][0:4])
            df_cabecalho_eqp['Grupo de planejamento'].append(df_1['LI_N3'][i][9:12])
            df_cabecalho_eqp['Txt.descr.cabeçalho'].append(np.nan)

            # if str(0) in df_1['ESTADO MAQ'][i]:
            #  df_cabecalho_eqp['Txt.descr.cabeçalho'].append('ANTES DE INICIAR DEVE BLOQUEAR O EQUIPAMENTO, REALIZAR A APR E UTILIZAR TODOS OS EPIs (LUVAS, LUVEX, ÓCULOS SEG. E CAPACETE). TODAS AS NORMAS DE SEGURANÇA DO ALIMENTO E QUALIDADE DEVEM SER SEGUIDAS. \nAPÓS CONCLUIR AS ATIVIDADES DEVE SER GARANTIDA A LIMPEZA DA ÁREA, DO EQUIPAMENTO E RETIRADA TODAS FERRAMENTAS E RESÍDUOS CONTAMINANTES. PARA EQUIPAMENTOS CRÍTICOS, É EXIGIDA A LIBERAÇÃO DA QUALIDADE.')
            # else:
            #  df_cabecalho_eqp['Txt.descr.cabeçalho'].append('É CRUCIAL UTILIZAR TODOS OS EPIs (LUVEX, LUVAS, ÓCULOS SEG. E CAPACETE) E SEMPRE MANTER A DISTANCIA DAS PARTES MÓVEIS DOS EQUIPAMENTOS.')

            df_cabecalho_eqp['Numeração Externa'].append(
                str(df_cabecalho_eqp['Chave do grupo de listas de tarefas*'][-1]) + '_' + str(
                    df_cabecalho_eqp['Contador de grupos*'][-1]))
            if str(df_1['ROTA?'][i]) == 'NAO':
                df_cabecalho_eqp['Equipamento'].append(df_1['EQUIPAMENTO OP'][i])
                df_cabecalho_eqp['ID_SAP'].append(df_1['ID SAP EQUIP OP'][i])
                df_cabecalho_eqp['Local de instalação'].append(np.nan)
            else:
                df_cabecalho_eqp['Equipamento'].append(np.nan)
                df_cabecalho_eqp['ID_SAP'].append(np.nan)
                df_cabecalho_eqp['Local de instalação'].append(df_1['ID N6/N5'][i])
            df_cabecalho_eqp['LI_N3'].append(df_1['LI_N3'][i])

            indice = indice + 1

        for chave, valor in df_cabecalho_eqp.items():
            if not valor:  # Verifica se a lista está vazia
                df_cabecalho_eqp[chave] = [np.nan] * len(
                    df_cabecalho_eqp['Chave do grupo de listas de tarefas*'])  # Substitui por np.nan

        df_cabecalho_eqp = checagem_df(pd.DataFrame(df_cabecalho_eqp))


        #   Preenchendo lista de tarefas sem lub e calib

        df_opativ = {
            'Chave do grupo de listas de tarefas*': [],
            'Contador de grupos*': [],
            'Número da atividade*': [],
            'Sequencial*': [],
            'Sub Operacao': [],
            'Centro de trabalho': [],
            'Centro': [],
            'Chave de controle': [],
            'Descrição da operação': [],
            'Fator de execução': [],
            'N do equipamento': [],
            'Local de instalação': [],
            'Chave de cálculo': [],
            'Trabalho envolvido na atividade': [],
            'Unidade para trabalho (formato ISO)': [],
            'Tipo de atividade': [],
            'Número de capacidades necessárias': [],
            'Duração normal da atividade': [],
            'Duração/unidade normal (formato ISO)': [],
            'Porcentagem de trabalho': [],
            'Quantidade da ordem': [],
            'Unidade quantidade ordem (formato ISO)': [],
            'Crit.ordenação': [],
            'Preço por unidade': [],
            'Moeda': [],
            'Unidade de preço': [],
            'Registro info para compras': [],
            'Fornecedor': [],
            'Prazo de entrega previsto em dias': [],
            'Acordo de compra': [],
            'Item de acordo de compra': [],
            'Classe de custo': [],
            'Grupo de mercadorias': [],
            'Grupo de compradores': [],
            'Organização de compras': [],
            'Texto descritivo de operação': [],
            'NC?': []
        }
        df_2 = df_tl.copy()   # *********** ALTERADO 24/05/2024
        df_2['ID EQUIP OP'] = np.where((df_2['ROTA?'] == 'SIM') | (df_2['ROTA?'] == 'PERSONALIZADO'), df_2['ID EQUIP OP'], np.nan)
        df_2['ID SAP EQUIP OP'] = np.where((df_2['ROTA?'] == 'SIM') | (df_2['ROTA?'] == 'PERSONALIZADO'), df_2['ID SAP EQUIP OP'], np.nan)
        df_2['EQUIPAMENTO OP'] = np.where((df_2['ROTA?'] == 'SIM') | (df_2['ROTA?'] == 'PERSONALIZADO'), df_2['EQUIPAMENTO OP'], np.nan)
        df_2 = df_2.drop_duplicates(subset=['TASK LIST', 'OPERAÇÃO: TEXTO CURTO', 'ID SAP EQUIP OP', 'EQUIPAMENTO OP'],
                                     keep='last').reset_index(drop=True)
        # df_2 = df_tl.drop_duplicates( subset = ['TASK LIST','OPERAÇÃO: TEXTO CURTO'], keep = 'last')
        # df_2 = df_2.sort_values(by=['Índices'])
        # df_2 = df_2.reset_index(drop=True)

        indice = -1

        for i in range(len(df_2['LI_N3'])):

            if 'CALI' in df_2['TASK LIST'][i][0:4] or 'LUB' in df_2['TASK LIST'][i][0:4]:
                continue

            if i == 0 or df_2['TASK LIST'][i] != df_2['TASK LIST'][i - 1]:

                indice = indice + 1

                # Operação de cabeçalho
                df_opativ['Chave do grupo de listas de tarefas*'].append(num_carga + indice)
                i_cabecalho = df_opativ['Chave do grupo de listas de tarefas*'].index(
                    num_carga + indice)  # Salvar índice do cabeçalho pra somar tempos

                df_opativ['Contador de grupos*'].append(1)
                df_opativ['Número da atividade*'].append('0010')
                df_opativ['Sequencial*'].append(indice)
                df_opativ['Sub Operacao'].append(np.nan)
                df_opativ['Centro de trabalho'].append(df_2['CT PM'][i])
                df_opativ['Centro'].append(df_2['LI_N3'][i][0:4])
                df_opativ['Chave de controle'].append('PM01')
                df_opativ['Descrição da operação'].append(df_2['TASK LIST'][i])
                # df_opativ['Fator de execução'].append(1 if 'FUNC' in df_2['TASK LIST'][i] else 0)
                df_opativ['Fator de execução'].append(1)
                df_opativ['N do equipamento'].append(np.nan)
                df_opativ['NC?'].append(np.nan)
                df_opativ['Chave de cálculo'].append(2)
                try:
                    df_opativ['Trabalho envolvido na atividade'].append(
                        int(df_2['DURAÇÃO (min)'][i]))  ## SOMAR TODAS AS SUBS
                except:
                    df_opativ['Trabalho envolvido na atividade'].append(0)  ## SOMAR TODAS AS SUBS
                df_opativ['Unidade para trabalho (formato ISO)'].append('MIN')
                df_opativ['Tipo de atividade'].append('MANUT')
                df_opativ['Número de capacidades necessárias'].append(int(df_2['OPERADORES'][i]))
                df_opativ['Duração normal da atividade'].append(
                    df_opativ['Trabalho envolvido na atividade'][-1] * df_opativ['Número de capacidades necessárias'][
                        -1])
                df_opativ['Duração/unidade normal (formato ISO)'].append('MIN')
                df_opativ['Porcentagem de trabalho'].append(100)
                if str(0) in str(df_2['ESTADO MAQ'][i]):
                    df_opativ['Texto descritivo de operação'].append(
                        'ANTES DE INICIAR DEVE BLOQUEAR O EQUIPAMENTO, REALIZAR A APR E UTILIZAR TODOS OS EPIs (LUVAS, LUVEX, ÓCULOS SEG. E CAPACETE). TODAS AS NORMAS DE SEGURANÇA DO ALIMENTO E QUALIDADE DEVEM SER SEGUIDAS. \nAPÓS CONCLUIR AS ATIVIDADES DEVE SER GARANTIDA A LIMPEZA DA ÁREA, DO EQUIPAMENTO E RETIRADA TODAS FERRAMENTAS E RESÍDUOS CONTAMINANTES. PARA EQUIPAMENTOS CRÍTICOS, É EXIGIDA A LIBERAÇÃO DA QUALIDADE.')
                else:
                    df_opativ['Texto descritivo de operação'].append(
                        'É CRUCIAL UTILIZAR TODOS OS EPIs (LUVEX, LUVAS, ÓCULOS SEG. E CAPACETE) E SEMPRE MANTER A DISTANCIA DAS PARTES MÓVEIS DOS EQUIPAMENTOS.')

                i_sub = 10

                ###

                df_opativ['Chave do grupo de listas de tarefas*'].append(num_carga + indice)
                df_opativ['Contador de grupos*'].append(1)
                df_opativ['Número da atividade*'].append('0010')
                df_opativ['Sequencial*'].append(indice)

                if i_sub < 100:
                    df_opativ['Sub Operacao'].append('00' + str(i_sub))
                elif i_sub < 1000:
                    df_opativ['Sub Operacao'].append('0' + str(i_sub))
                elif i_sub < 10000:
                    df_opativ['Sub Operacao'].append(str(i_sub))
                else:
                    print(f"ERRO: SUB OPERAÇÕES ULTRAPASSARAM O VALOR MÁXIMO PERMITIDO NA TL '{df_2['TASK LIST'][i]}' ")

                df_opativ['Centro de trabalho'].append(df_2['CT PM'][i])
                df_opativ['Centro'].append(df_2['LI_N3'][i][0:4])
                df_opativ['Chave de controle'].append('PM01')
                df_opativ['Descrição da operação'].append(df_2['OPERAÇÃO: TEXTO CURTO'][i])
                # df_opativ['Fator de execução'].append(1 if 'FUNC' in df_2['TASK LIST'][i] else 0)
                df_opativ['Fator de execução'].append(1)
                df_opativ['NC?'].append('NC' if pd.isna(df_2['ID SAP EQUIP OP'][i]) and str('NAO') not in str(df_2['ROTA?'][i]) else df_2['ID SAP EQUIP OP'][i])   # 'NA' SE EQP NÃO SUBIU

                # Checar se está dentro da lista das task list que não irão subir:
                if df_2['TASK LIST'][i] in lista_nsubir:
                    df_opativ['NC?'][-1] = 'NC'
                    df_opativ['Sequencial*'][-1] = 'INDICADO PARA NÃO SUBIR'
                #

                df_opativ['N do equipamento'].append(df_2['ID SAP EQUIP OP'][i] if str('NAO') not in str(df_2['ROTA?'][i]) else np.nan)
                df_opativ['Chave de cálculo'].append(2)
                try:
                    df_opativ['Trabalho envolvido na atividade'].append(
                        int(df_2['DURAÇÃO (min)'][i]))  ## SOMAR TODAS AS SUBS
                except:
                    df_opativ['Trabalho envolvido na atividade'].append(0)  ## SOMAR TODAS AS SUBS
                df_opativ['Unidade para trabalho (formato ISO)'].append('MIN')
                df_opativ['Tipo de atividade'].append('MANUT')
                df_opativ['Número de capacidades necessárias'].append(int(df_2['OPERADORES'][i]))
                df_opativ['Duração normal da atividade'].append(
                    df_opativ['Trabalho envolvido na atividade'][-1] * df_opativ['Número de capacidades necessárias'][
                        -1])
                df_opativ['Duração/unidade normal (formato ISO)'].append('MIN')
                df_opativ['Porcentagem de trabalho'].append(100)
                df_opativ['Texto descritivo de operação'].append(df_2['OPERAÇÃO: TEXTO LONGO'][i])

                i_sub = i_sub + 10

                df_opativ['Trabalho envolvido na atividade'][i_cabecalho] = \
                df_opativ['Trabalho envolvido na atividade'][-1]
                df_opativ['Duração normal da atividade'][i_cabecalho] = df_opativ['Duração normal da atividade'][-1]

                ###

            else:
                df_opativ['Chave do grupo de listas de tarefas*'].append(num_carga + indice)
                df_opativ['Contador de grupos*'].append(1)
                df_opativ['Número da atividade*'].append('0010')
                df_opativ['Sequencial*'].append(indice)

                if i_sub < 100:
                    df_opativ['Sub Operacao'].append('00' + str(i_sub))
                elif i_sub < 1000:
                    df_opativ['Sub Operacao'].append('0' + str(i_sub))
                elif i_sub < 10000:
                    df_opativ['Sub Operacao'].append(str(i_sub))
                else:
                    print(f"ERRO: SUB OPERAÇÕES ULTRAPASSARAM O VALOR MÁXIMO PERMITIDO NA TL '{df_2['TASK LIST'][i]}' ")

                df_opativ['Centro de trabalho'].append(df_2['CT PM'][i])
                df_opativ['Centro'].append(df_2['LI_N3'][i][0:4])
                df_opativ['Chave de controle'].append('PM01')
                df_opativ['Descrição da operação'].append(df_2['OPERAÇÃO: TEXTO CURTO'][i])
                # df_opativ['Fator de execução'].append(1 if 'FUNC' in df_2['TASK LIST'][i] else 0)
                df_opativ['Fator de execução'].append(1)
                df_opativ['NC?'].append('NC' if pd.isna(df_2['ID SAP EQUIP OP'][i]) and str('NAO') not in str(df_2['ROTA?'][i]) else df_2['ID SAP EQUIP OP'][i])   # 'NA' SE EQP NÃO SUBIU

                # Checar se está dentro da lista das task list que não irão subir:
                if df_2['TASK LIST'][i] in lista_nsubir:
                    df_opativ['NC?'][-1] = 'NC'
                    df_opativ['Sequencial*'][-1] = 'INDICADO PARA NÃO SUBIR'
                #

                df_opativ['N do equipamento'].append(df_2['ID SAP EQUIP OP'][i] if str('NAO') not in str(df_2['ROTA?'][i]) else np.nan)
                df_opativ['Chave de cálculo'].append(2)
                try:
                    df_opativ['Trabalho envolvido na atividade'].append(
                        int(df_2['DURAÇÃO (min)'][i]))  ## SOMAR TODAS AS SUBS
                except:
                    df_opativ['Trabalho envolvido na atividade'].append(0)  ## SOMAR TODAS AS SUBS
                df_opativ['Unidade para trabalho (formato ISO)'].append('MIN')
                df_opativ['Tipo de atividade'].append('MANUT')
                df_opativ['Número de capacidades necessárias'].append(int(df_2['OPERADORES'][i]))
                df_opativ['Duração normal da atividade'].append(
                    df_opativ['Trabalho envolvido na atividade'][-1] * df_opativ['Número de capacidades necessárias'][
                        -1])
                df_opativ['Duração/unidade normal (formato ISO)'].append('MIN')
                df_opativ['Porcentagem de trabalho'].append(100)
                df_opativ['Texto descritivo de operação'].append(df_2['OPERAÇÃO: TEXTO LONGO'][i])

                i_sub = i_sub + 10

                df_opativ['Trabalho envolvido na atividade'][i_cabecalho] = \
                df_opativ['Trabalho envolvido na atividade'][i_cabecalho] + \
                df_opativ['Trabalho envolvido na atividade'][-1]
                df_opativ['Duração normal da atividade'][i_cabecalho] = df_opativ['Duração normal da atividade'][
                                                                            i_cabecalho] + \
                                                                        df_opativ['Duração normal da atividade'][-1]

        for chave, valor in df_opativ.items():
            if not valor:  # Verifica se a lista está vazia
                df_opativ[chave] = [np.nan] * len(
                    df_opativ['Chave do grupo de listas de tarefas*'])  # Substitui por np.nan

        df_opativ = pd.DataFrame(df_opativ)

        #
        # df_opativ = df_opativ.drop_duplicates( subset = ['Descrição da operação','N do equipamento'], keep = 'last').reset_index(drop = True)
        #

        # Checagem dos itens não carregados:

        ## Identifica as linhas onde 'N do equipamento' é igual a 'NC'
        linhas_na = df_opativ[df_opativ['NC?'] == 'NC']['Chave do grupo de listas de tarefas*']

        ## Marca todos os registros na coluna 'NC?' como 'NC' para os números iguais
        df_opativ.loc[df_opativ['Chave do grupo de listas de tarefas*'].isin(linhas_na), 'NC?'] = 'NC'

        ## Crie um novo DataFrame para armazenar todos os registros com 'NC'
        df_nc = df_opativ[df_opativ['NC?'] == 'NC'].copy()

        ## Remove as linhas com o mesmo número na coluna 1 quando 'NA' está presente
        df_opativ = df_opativ[~df_opativ['Chave do grupo de listas de tarefas*'].isin(linhas_na)].reset_index(drop=True)

        ## Resetar o índice
        df_opativ = df_opativ.reset_index(drop=True)

        df_opativ['Nova Chave'] = -1  # Crie uma nova coluna para armazenar os valores ajustados

        indice = -1

        for i in range(len(df_opativ['Chave do grupo de listas de tarefas*'])):
            if i == 0 or df_opativ['Chave do grupo de listas de tarefas*'][i] != \
                    df_opativ['Chave do grupo de listas de tarefas*'][i - 1]:
                indice = indice + 1
            df_opativ['Nova Chave'][i] = num_carga + indice

        df_opativ['Chave do grupo de listas de tarefas*'] = df_opativ['Nova Chave']

        ## Remova a coluna original e renomeie a nova coluna
        df_opativ = df_opativ.drop(columns=['Nova Chave'])

        ###

        # Numerar sequencial

        seq = 1
        for i in range(len(df_opativ['Sequencial*'])):
            df_opativ['Sequencial*'][i] = seq
            seq += 1
            if df_opativ['Sequencial*'][i] == 9999:
                seq = 1

        #


        #   Cabeçalho sem lub e calib

        df_cabecalho = {
            'Chave do grupo de listas de tarefas*': [],
            'Contador de grupos*': [],
            'Data de criação registro': [],
            'Data de início da validade': [],
            'Descrição': [],
            'Centro de planejamento*': [],
            'Centro de trabalho': [],
            'Centro de centro de trabalho': [],
            'Utilização lista de tarefas': [],
            'Status global': [],
            'Estratégia de manutenção': [],
            'Grupo de planejamento': [],
            'Condições da instalação': [],
            'Conjunto': [],
            'Txt.descr.cabeçalho': [],
            'Numeração Externa': [],
            'LI_N3': []
        }

        df_1 = df_1.drop_duplicates(subset=['TASK LIST'], keep='last').reset_index(drop=True)
        # df_1 = df_1.sort_values(by=['índices'])
        # df_1 = df_1.reset_index(drop=True)

        indice = 0
        for i in range(len(df_1['LI_N3'])):
            if i > 0 and df_1['TASK LIST'][i] == df_1['TASK LIST'][i - 1]:
                continue
            if 'CALI' in df_1['TASK LIST'][i][0:4] or 'LUB' in df_1['TASK LIST'][i][0:4]:
                continue
            df_cabecalho['Chave do grupo de listas de tarefas*'].append(num_carga + indice)
            df_cabecalho['Data de criação registro'].append(data_hoje)
            df_cabecalho['Data de início da validade'].append('01.12.2023')
            df_cabecalho['Contador de grupos*'].append(1)
            df_cabecalho['Utilização lista de tarefas'].append(4)
            df_cabecalho['Status global'].append(4)
            df_cabecalho['Condições da instalação'].append(1 if 'FUNC' in df_1['TASK LIST'][i] else 0)
            df_cabecalho['Descrição'].append(df_1['TASK LIST'][i])
            df_cabecalho['Centro de planejamento*'].append(df_1['LI_N3'][i][0:4])
            df_cabecalho['Centro de trabalho'].append(df_1['CT PM'][i])
            df_cabecalho['Centro de centro de trabalho'].append(df_1['LI_N3'][i][0:4])
            df_cabecalho['Grupo de planejamento'].append(df_1['LI_N3'][i][9:12])
            df_cabecalho['LI_N3'].append(df_1['LI_N3'][i])

            if str(0) in str(df_1['ESTADO MAQ'][i]):
                df_cabecalho['Txt.descr.cabeçalho'].append(
                    'ANTES DE INICIAR DEVE BLOQUEAR O EQUIPAMENTO, REALIZAR A APR E UTILIZAR TODOS OS EPIs (LUVAS, ÓCULOS SEG. E CAPACETE). TODAS AS NORMAS DE SEGURANÇA DO ALIMENTO E QUALIDADE DEVEM SER SEGUIDAS. \nAPÓS CONCLUIR AS ATIVIDADES DEVE SER GARANTIDA A LIMPEZA DA ÁREA, DO EQUIPAMENTO E RETIRADA TODAS FERRAMENTAS E RESÍDUOS CONTAMINANTES. PARA EQUIPAMENTOS CRÍTICOS, É EXIGIDA A LIBERAÇÃO DA QUALIDADE.')
            else:
                df_cabecalho['Txt.descr.cabeçalho'].append(
                    'É CRUCIAL UTILIZAR TODOS OS EPIs (LUVAS, ÓCULOS SEG. E CAPACETE) E SEMPRE MANTER A DISTANCIA DAS PARTES MÓVEIS DOS EQUIPAMENTOS.')

            df_cabecalho['Numeração Externa'].append(
                str(df_cabecalho['Chave do grupo de listas de tarefas*'][-1]) + '_' + str(
                    df_cabecalho['Contador de grupos*'][-1]))
            indice = indice + 1

        for chave, valor in df_cabecalho.items():
            if not valor:  # Verifica se a lista está vazia
                df_cabecalho[chave] = [np.nan] * len(
                    df_cabecalho['Chave do grupo de listas de tarefas*'])  # Substitui por np.nan

        df_cabecalho = checagem_df(pd.DataFrame(df_cabecalho))

        # Checagem dos itens não carregados:

        # Identificar as linhas a serem removidas
        linhas_para_remover = df_cabecalho[df_cabecalho['Descrição'].isin(df_nc['Descrição da operação'])].index

        # Remover as linhas do DataFrame original
        df_cabecalho = df_cabecalho.drop(linhas_para_remover).reset_index(drop=True)

        # Resetar o índice
        df_cabecalho = df_cabecalho.reset_index(drop=True)

        df_cabecalho['Nova Chave'] = -1  # Crie uma nova coluna para armazenar os valores ajustados

        indice = -1

        for i in range(len(df_cabecalho['Chave do grupo de listas de tarefas*'])):
            if i == 0 or df_cabecalho['Chave do grupo de listas de tarefas*'][i] != \
                    df_cabecalho['Chave do grupo de listas de tarefas*'][i - 1]:
                indice = indice + 1
            df_cabecalho['Nova Chave'][i] = num_carga + indice

        df_cabecalho['Chave do grupo de listas de tarefas*'] = df_cabecalho['Nova Chave']
        df_cabecalho['Numeração Externa'] = df_cabecalho['Chave do grupo de listas de tarefas*'].astype(str) + '_' + \
                                            df_cabecalho['Contador de grupos*'].astype(str)

        # Remova a coluna original e renomeie a nova coluna
        df_cabecalho = df_cabecalho.drop(columns=['Nova Chave'])

        ###

        indice_inserir_coluna = df_cabecalho.columns.get_loc(df_cabecalho.columns[-1]) + 1
        df_cabecalho.insert(loc=indice_inserir_coluna, column='Descrição antiga', value=df_cabecalho['Descrição'])


        #   Preenchendo lista de tarefas com lub

        df_opativ_lub = {
            'Chave do grupo de listas de tarefas*': [],
            'Contador de grupos*': [],
            'Número da atividade*': [],
            'Sequencial*': [],
            'Sub Operacao': [],
            'Centro de trabalho': [],
            'Centro': [],
            'Chave de controle': [],
            'Descrição da operação': [],
            'Fator de execução': [],
            'N do equipamento': [],
            'Local de instalação': [],
            'Chave de cálculo': [],
            'Trabalho envolvido na atividade': [],
            'Unidade para trabalho (formato ISO)': [],
            'Tipo de atividade': [],
            'Número de capacidades necessárias': [],
            'Duração normal da atividade': [],
            'Duração/unidade normal (formato ISO)': [],
            'Porcentagem de trabalho': [],
            'Quantidade da ordem': [],
            'Unidade quantidade ordem (formato ISO)': [],
            'Crit.ordenação': [],
            'Preço por unidade': [],
            'Moeda': [],
            'Unidade de preço': [],
            'Registro info para compras': [],
            'Fornecedor': [],
            'Prazo de entrega previsto em dias': [],
            'Acordo de compra': [],
            'Item de acordo de compra': [],
            'Classe de custo': [],
            'Grupo de mercadorias': [],
            'Grupo de compradores': [],
            'Organização de compras': [],
            'Texto descritivo de operação': [],
            'NC?': []
        }
        try:
            num_carga = list(df_opativ[df_opativ.columns[0]])[-1] + 1
        except:
            num_carga = 0
            
        df_2 = df_tl.copy()   # *********** ALTERADO 24/05/2024
        df_2['ID EQUIP OP'] = np.where((df_2['ROTA?'] == 'SIM') | (df_2['ROTA?'] == 'PERSONALIZADO'), df_2['ID EQUIP OP'], np.nan)
        df_2['ID SAP EQUIP OP'] = np.where((df_2['ROTA?'] == 'SIM') | (df_2['ROTA?'] == 'PERSONALIZADO'), df_2['ID SAP EQUIP OP'], np.nan)
        df_2['EQUIPAMENTO OP'] = np.where((df_2['ROTA?'] == 'SIM') | (df_2['ROTA?'] == 'PERSONALIZADO'), df_2['EQUIPAMENTO OP'], np.nan)
        
        # df_2 = df_tl.sort_values(by=['TASK LIST','N6/N5'])
        df_2 = df_2.drop_duplicates(subset=['TASK LIST', 'OPERAÇÃO: TEXTO CURTO', 'ID SAP EQUIP OP', 'EQUIPAMENTO OP'],
                                     keep='last').reset_index(drop=True)
        # df_2 = df_tl.drop_duplicates( subset = ['TASK LIST','OPERAÇÃO: TEXTO CURTO'], keep = 'last')
        # df_2 = df_2.sort_values(by=['Índices'])
        # df_2 = df_2.reset_index(drop=True)

        indice = -1

        for i in range(len(df_2['LI_N3'])):

            if 'LUB' not in df_2['TASK LIST'][i][0:4]:
                continue

            if i == 0 or df_2['TASK LIST'][i] != df_2['TASK LIST'][i - 1]:

                indice = indice + 1

                i_sub = 10

                ###
                df_opativ_lub['Chave do grupo de listas de tarefas*'].append(num_carga + indice)
                df_opativ_lub['Contador de grupos*'].append(1)
                df_opativ_lub['Sub Operacao'].append(np.nan)
                df_opativ_lub['Sequencial*'].append(indice)

                if i_sub < 100:
                    df_opativ_lub['Número da atividade*'].append('00' + str(i_sub))
                elif i_sub < 1000:
                    df_opativ_lub['Número da atividade*'].append('0' + str(i_sub))
                elif i_sub < 10000:
                    df_opativ_lub['Número da atividade*'].append(str(i_sub))
                else:
                    print(f"ERRO: SUB OPERAÇÕES ULTRAPASSARAM O VALOR MÁXIMO PERMITIDO NA TL '{df_2['TASK LIST'][i]}' ")

                df_opativ_lub['Centro de trabalho'].append(df_2['CT PM'][i])
                df_opativ_lub['Centro'].append(df_2['LI_N3'][i][0:4])
                df_opativ_lub['Chave de controle'].append('PM01')
                df_opativ_lub['Descrição da operação'].append(df_2['OPERAÇÃO: TEXTO CURTO'][i])
                # df_opativ['Fator de execução'].append(1 if 'FUNC' in df_2['TASK LIST'][i] else 0)
                df_opativ_lub['Fator de execução'].append(1)
                df_opativ_lub['NC?'].append('NC' if pd.isna(df_2['ID SAP EQUIP OP'][i]) and str('NAO') not in str(df_2['ROTA?'][i]) else df_2['ID SAP EQUIP OP'][i])

                # Checar se está dentro da lista das task list que não irão subir:
                if df_2['TASK LIST'][i] in lista_nsubir:
                    df_opativ_lub['NC?'][-1] = 'NC'
                    df_opativ_lub['Sequencial*'][-1] = 'INDICADO PARA NÃO SUBIR'
                #

                df_opativ_lub['N do equipamento'].append(df_2['ID SAP EQUIP OP'][i] if str('NAO') not in str(df_2['ROTA?'][i]) else np.nan)
                df_opativ_lub['Chave de cálculo'].append(2)
                df_opativ_lub['Trabalho envolvido na atividade'].append(
                    int(df_2['DURAÇÃO (min)'][i]))  ## SOMAR TODAS AS SUBS
                df_opativ_lub['Unidade para trabalho (formato ISO)'].append('MIN')
                df_opativ_lub['Tipo de atividade'].append('MANUT')
                df_opativ_lub['Número de capacidades necessárias'].append(int(df_2['OPERADORES'][i]))
                df_opativ_lub['Duração normal da atividade'].append(
                    df_opativ_lub['Trabalho envolvido na atividade'][-1] *
                    df_opativ_lub['Número de capacidades necessárias'][-1])
                df_opativ_lub['Duração/unidade normal (formato ISO)'].append('MIN')
                df_opativ_lub['Porcentagem de trabalho'].append(100)

                if pd.isna(df_2['OPERAÇÃO: TEXTO LONGO'][i]):  # Caso o texto longo não exista na operação
                    if str(0) in df_2['ESTADO MAQ'][i] and i_sub == 10:  # Máquina parada
                        texto_l = str(
                            '\nANTES DE INICIAR DEVE BLOQUEAR O EQUIPAMENTO, REALIZAR A APR E UTILIZAR TODOS OS EPIs (LUVAS, LUVEX, ÓCULOS SEG. E CAPACETE). TODAS AS NORMAS DE SEGURANÇA DO ALIMENTO E QUALIDADE DEVEM SER SEGUIDAS.\nAPÓS CONCLUIR AS ATIVIDADES DEVE SER GARANTIDA A LIMPEZA DA ÁREA, DO EQUIPAMENTO E RETIRADA TODAS FERRAMENTAS E RESÍDUOS CONTAMINANTES. PARA EQUIPAMENTOS CRÍTICOS, É EXIGIDA A LIBERAÇÃO DA QUALIDADE.')
                        df_opativ_lub['Texto descritivo de operação'].append(texto_l)
                    elif str(1) in df_2['ESTADO MAQ'][i] and i_sub == 10:  # Máquina funcional
                        texto_l = str(
                            '\nÉ CRUCIAL UTILIZAR TODOS OS EPIs (LUVEX, LUVAS, ÓCULOS SEG. E CAPACETE) E SEMPRE MANTER A DISTANCIA DAS PARTES MÓVEIS DOS EQUIPAMENTOS.')
                        df_opativ_lub['Texto descritivo de operação'].append(texto_l)
                    else:
                        df_opativ_lub['Texto descritivo de operação'].append(np.nan)

                else:  # Caso exista texto longo na operação
                    if str(0) in df_2['ESTADO MAQ'][i] and i_sub == 10:
                        texto_l = str(str(df_2['OPERAÇÃO: TEXTO LONGO'][
                                              i]) + '\nANTES DE INICIAR DEVE BLOQUEAR O EQUIPAMENTO, REALIZAR A APR E UTILIZAR TODOS OS EPIs (LUVAS, LUVEX, ÓCULOS SEG. E CAPACETE). TODAS AS NORMAS DE SEGURANÇA DO ALIMENTO E QUALIDADE DEVEM SER SEGUIDAS.\nAPÓS CONCLUIR AS ATIVIDADES DEVE SER GARANTIDA A LIMPEZA DA ÁREA, DO EQUIPAMENTO E RETIRADA TODAS FERRAMENTAS E RESÍDUOS CONTAMINANTES. PARA EQUIPAMENTOS CRÍTICOS, É EXIGIDA A LIBERAÇÃO DA QUALIDADE.')
                        df_opativ_lub['Texto descritivo de operação'].append(texto_l)
                    elif str(1) in df_2['ESTADO MAQ'][i] and i_sub == 10:
                        texto_l = str(str(df_2['OPERAÇÃO: TEXTO LONGO'][
                                              i]) + '\nÉ CRUCIAL UTILIZAR TODOS OS EPIs (LUVEX, LUVAS, ÓCULOS SEG. E CAPACETE) E SEMPRE MANTER A DISTANCIA DAS PARTES MÓVEIS DOS EQUIPAMENTOS.')
                        df_opativ_lub['Texto descritivo de operação'].append(texto_l)
                    else:
                        df_opativ_lub['Texto descritivo de operação'].append(df_2['OPERAÇÃO: TEXTO LONGO'][i])

                i_sub = i_sub + 10

                ###

            else:
                df_opativ_lub['Chave do grupo de listas de tarefas*'].append(num_carga + indice)
                df_opativ_lub['Contador de grupos*'].append(1)
                df_opativ_lub['Sub Operacao'].append(np.nan)
                df_opativ_lub['Sequencial*'].append(indice)

                if i_sub < 100:
                    df_opativ_lub['Número da atividade*'].append('00' + str(i_sub))
                elif i_sub < 1000:
                    df_opativ_lub['Número da atividade*'].append('0' + str(i_sub))
                elif i_sub < 10000:
                    df_opativ_lub['Número da atividade*'].append(str(i_sub))
                else:
                    print(f"ERRO: SUB OPERAÇÕES ULTRAPASSARAM O VALOR MÁXIMO PERMITIDO NA TL '{df_2['TASK LIST'][i]}' ")

                df_opativ_lub['Centro de trabalho'].append(df_2['CT PM'][i])
                df_opativ_lub['Centro'].append(df_2['LI_N3'][i][0:4])
                df_opativ_lub['Chave de controle'].append('PM01')
                df_opativ_lub['Descrição da operação'].append(df_2['OPERAÇÃO: TEXTO CURTO'][i])
                # df_opativ['Fator de execução'].append(1 if 'FUNC' in df_2['TASK LIST'][i] else 0)
                df_opativ_lub['Fator de execução'].append(1)
                df_opativ_lub['NC?'].append('NC' if pd.isna(df_2['ID SAP EQUIP OP'][i]) and str('NAO') not in str(df_2['ROTA?'][i]) else df_2['ID SAP EQUIP OP'][i])   # 'NA' SE EQP NÃO SUBIU

                # Checar se está dentro da lista das task list que não irão subir:
                if df_2['TASK LIST'][i] in lista_nsubir:
                    df_opativ_lub['NC?'][-1] = 'NC'
                    df_opativ_lub['Sequencial*'][-1] = 'INDICADO PARA NÃO SUBIR'
                #

                df_opativ_lub['N do equipamento'].append(df_2['ID SAP EQUIP OP'][i] if str('NAO') not in str(df_2['ROTA?'][i]) else np.nan)
                df_opativ_lub['Chave de cálculo'].append(2)

                try:
                    df_opativ_lub['Trabalho envolvido na atividade'].append(
                        int(df_2['DURAÇÃO (min)'][i]))  ## SOMAR TODAS AS SUBS
                except:
                    df_opativ_lub['Trabalho envolvido na atividade'].append(0)  ## SOMAR TODAS AS SUBS

                df_opativ_lub['Unidade para trabalho (formato ISO)'].append('MIN')
                df_opativ_lub['Tipo de atividade'].append('MANUT')
                df_opativ_lub['Número de capacidades necessárias'].append(int(df_2['OPERADORES'][i]))
                df_opativ_lub['Duração normal da atividade'].append(
                    df_opativ_lub['Trabalho envolvido na atividade'][-1] *
                    df_opativ_lub['Número de capacidades necessárias'][-1])
                df_opativ_lub['Duração/unidade normal (formato ISO)'].append('MIN')
                df_opativ_lub['Porcentagem de trabalho'].append(100)
                df_opativ_lub['Texto descritivo de operação'].append(df_2['OPERAÇÃO: TEXTO LONGO'][i])

                i_sub = i_sub + 10

        for chave, valor in df_opativ_lub.items():
            if not valor:  # Verifica se a lista está vazia
                df_opativ_lub[chave] = [np.nan] * len(
                    df_opativ_lub['Chave do grupo de listas de tarefas*'])  # Substitui por np.nan

        df_opativ_lub = pd.DataFrame(df_opativ_lub)

        # Checagem dos itens não carregados (eles serão excluídos apenas quando forem excluídos do cabeçalho de lub):

        ## Filtrar o DataFrame para obter apenas linhas onde 'NC?' é igual a 'NC'
        df_nc_lub = df_opativ_lub[df_opativ_lub['NC?'] == 'NC']

        ## Obter os valores únicos da coluna 'Chave do grupo de listas de tarefas* (antes da reordenação)'
        chaves_nc_lub = df_nc_lub['Chave do grupo de listas de tarefas*'].unique()

        ## Salvar os valores únicos em uma lista
        chaves_nc_lub_lista = chaves_nc_lub.tolist()


        #   Cabeçalho lub

        df_cabecalho_lub = {
            'Chave do grupo de listas de tarefas*': [],
            'Contador de grupos*': [],
            'Data de criação registro': [],
            'Data de início da validade': [],
            'Descrição': [],
            'Centro de planejamento*': [],
            'Centro de trabalho': [],
            'Centro de centro de trabalho': [],
            'Utilização lista de tarefas': [],
            'Status global': [],
            'Estratégia de manutenção': [],
            'Grupo de planejamento': [],
            'Condições da instalação': [],
            'Conjunto': [],
            'Txt.descr.cabeçalho': [],
            'Numeração Externa': [],
            'NC?': [],
            'LI_N3': []
        }

        df_1 = df_1.drop_duplicates(subset=['TASK LIST'], keep='last').reset_index(drop=True)
        # df_1 = df_1.sort_values(by=['índices'])
        # df_1 = df_1.reset_index(drop=True)

        indice = 0
        for i in range(len(df_1['LI_N3'])):
            if i > 0 and df_1['TASK LIST'][i] == df_1['TASK LIST'][i - 1]:
                continue
            if 'LUB' not in df_1['TASK LIST'][i][0:4]:
                continue
            df_cabecalho_lub['Chave do grupo de listas de tarefas*'].append(num_carga + indice)
            df_cabecalho_lub['Data de criação registro'].append(data_hoje)
            df_cabecalho_lub['Data de início da validade'].append('01.12.2023')
            df_cabecalho_lub['Contador de grupos*'].append(1)
            df_cabecalho_lub['Utilização lista de tarefas'].append(4)
            df_cabecalho_lub['Status global'].append(4)
            df_cabecalho_lub['Condições da instalação'].append(1 if 'FUNC' in df_1['TASK LIST'][i] else 0)
            df_cabecalho_lub['Descrição'].append(df_1['TASK LIST'][i])
            df_cabecalho_lub['Centro de planejamento*'].append(df_1['LI_N3'][i][0:4])
            df_cabecalho_lub['Centro de trabalho'].append(df_1['CT PM'][i])
            df_cabecalho_lub['Centro de centro de trabalho'].append(df_1['LI_N3'][i][0:4])
            df_cabecalho_lub['Grupo de planejamento'].append(df_1['LI_N3'][i][9:12])
            df_cabecalho_lub['LI_N3'].append(df_1['LI_N3'][i])

            if str(0) in df_1['ESTADO MAQ'][i]:
                df_cabecalho_lub['Txt.descr.cabeçalho'].append(
                    'ANTES DE INICIAR DEVE BLOQUEAR O EQUIPAMENTO, REALIZAR A APR E UTILIZAR TODOS OS EPIs (LUVAS, LUVEX, ÓCULOS SEG. E CAPACETE). TODAS AS NORMAS DE SEGURANÇA DO ALIMENTO E QUALIDADE DEVEM SER SEGUIDAS. \nAPÓS CONCLUIR AS ATIVIDADES DEVE SER GARANTIDA A LIMPEZA DA ÁREA, DO EQUIPAMENTO E RETIRADA TODAS FERRAMENTAS E RESÍDUOS CONTAMINANTES. PARA EQUIPAMENTOS CRÍTICOS, É EXIGIDA A LIBERAÇÃO DA QUALIDADE.')
            else:
                df_cabecalho_lub['Txt.descr.cabeçalho'].append(
                    'É CRUCIAL UTILIZAR TODOS OS EPIs (LUVEX, LUVAS, ÓCULOS SEG. E CAPACETE) E SEMPRE MANTER A DISTANCIA DAS PARTES MÓVEIS DOS EQUIPAMENTOS.')

            df_cabecalho_lub['Numeração Externa'].append(
                str(df_cabecalho_lub['Chave do grupo de listas de tarefas*'][-1]) + '_' + str(
                    df_cabecalho_lub['Contador de grupos*'][-1]))

            df_cabecalho_lub['NC?'].append('NC' if pd.isna(df_1['ID SAP EQUIP OP'][i]) and str('NAO') not in df_1['ROTA?'][i] else df_1['ID SAP EQUIP OP'][i])   # 'NA' SE EQP NÃO SUBIU

            # Checar se está dentro da lista das task list que não irão subir:
            if df_1['TASK LIST'][i] in lista_nsubir:
                df_opativ_lub['NC?'][-1] = 'NC'
                df_opativ_lub['Contador de grupos*'][-1] = 'INDICADO PARA NÃO SUBIR'
            #

            indice = indice + 1

        for chave, valor in df_cabecalho_lub.items():
            if not valor:  # Verifica se a lista está vazia
                df_cabecalho_lub[chave] = [np.nan] * len(
                    df_cabecalho_lub['Chave do grupo de listas de tarefas*'])  # Substitui por np.nan

        df_cabecalho_lub = checagem_df(pd.DataFrame(df_cabecalho_lub))
        print(len(df_cabecalho_lub['Chave do grupo de listas de tarefas*']))

        ##################################################################


        # REMOCAO DOS ITENS NAO CARREGADOS DE LUB (OPERACOES E CABECALHO)


        ##################################################################

        # Filtrar df_opativ_lub e df_cabecalho_lub para remover as linhas com 'Chave do grupo de listas de tarefas*' presentes em chaves_nc_lub
        df_opativ_lub = df_opativ_lub[~df_opativ_lub['Chave do grupo de listas de tarefas*'].isin(chaves_nc_lub)]
        df_cabecalho_lub = df_cabecalho_lub[
            ~df_cabecalho_lub['Chave do grupo de listas de tarefas*'].isin(chaves_nc_lub)]
        # Resetar o índice
        df_opativ_lub = df_opativ_lub.reset_index(drop=True)
        df_cabecalho_lub = df_cabecalho_lub.reset_index(drop=True)

        # REMOVER 'NC' PARA OPERACOES LUB E RENUMERAR CHAVE:
        df_opativ_lub['Nova Chave'] = -1  # Crie uma nova coluna para armazenar os valores ajustados

        indice = -1

        for i in range(len(df_opativ_lub['Chave do grupo de listas de tarefas*'])):
            if i == 0 or df_opativ_lub['Chave do grupo de listas de tarefas*'][i] != \
                    df_opativ_lub['Chave do grupo de listas de tarefas*'][i - 1]:
                indice = indice + 1
            df_opativ_lub['Nova Chave'][i] = num_carga + indice

        df_opativ_lub['Chave do grupo de listas de tarefas*'] = df_opativ_lub['Nova Chave']

        ## Remova a coluna original e renomeie a nova coluna
        df_opativ_lub = df_opativ_lub.drop(columns=['Nova Chave'])

        ## Numerar sequencial

        seq = 1
        for i in range(len(df_opativ_lub['Sequencial*'])):
            df_opativ_lub['Sequencial*'][i] = seq
            seq += 1
            if df_opativ_lub['Sequencial*'][i] == 9999:
                seq = 1

        ###

        # REMOVER 'NC' PARA CABEÇALHO LUB E RENUMERAR CHAVE:
        df_cabecalho_lub['Nova Chave'] = -1  # Crie uma nova coluna para armazenar os valores ajustados

        indice = -1

        for i in range(len(df_cabecalho_lub['Chave do grupo de listas de tarefas*'])):
            if i == 0 or df_cabecalho_lub['Chave do grupo de listas de tarefas*'][i] != \
                    df_cabecalho_lub['Chave do grupo de listas de tarefas*'][i - 1]:
                indice = indice + 1
            df_cabecalho_lub['Nova Chave'][i] = num_carga + indice

        df_cabecalho_lub['Chave do grupo de listas de tarefas*'] = df_cabecalho_lub['Nova Chave']
        df_cabecalho_lub['Numeração Externa'] = df_cabecalho_lub['Chave do grupo de listas de tarefas*'].astype(
            str) + '_' + df_cabecalho_lub['Contador de grupos*'].astype(str)

        ## Remova a coluna original e renomeie a nova coluna
        df_cabecalho_lub = df_cabecalho_lub.drop(columns=['Nova Chave'])

        ###

        indice_inserir_coluna = df_cabecalho_lub.columns.get_loc(df_cabecalho_lub.columns[-1]) + 1
        df_cabecalho_lub.insert(loc=indice_inserir_coluna, column='Descrição antiga',
                                value=df_cabecalho_lub['Descrição'])


        #   Auto-Revisão do Plano S/LUB

        import statistics

        DF = df_opativ
        indice_inserir_coluna = DF.columns.get_loc(DF.columns[-1]) + 1
        DF.insert(loc=indice_inserir_coluna, column='REVISAR', value=1)
        indice_inserir_coluna = DF.columns.get_loc(DF.columns[-1]) + 1
        DF.insert(loc=indice_inserir_coluna, column='PTS', value=1)

        media_len_oppadrao = statistics.mean(len(str(descricao).split()) for descricao in DF['Descrição da operação'])
        media_len_textodesc = statistics.mean(len(str(descricao).split()) for descricao in DF['Texto descritivo de operação'])

        for i in range(len(DF['Chave do grupo de listas de tarefas*'])):
            if pd.isna(DF['Sub Operacao'][i]):  # Checar se e TL (cabeçalho)
                # TL com apenas uma operação
                if DF['Chave do grupo de listas de tarefas*'].value_counts()[
                    DF['Chave do grupo de listas de tarefas*'][i]] <= 2:
                    DF['REVISAR'][i] = 'APENAS UMA OPERAÇÃO'
                    DF['PTS'][i] = 0.5

                # TL com máquina parada e peridiocidade menor que 1M
                if ' FUNC ' not in DF['Descrição da operação'][i] and (
                        DF['Descrição da operação'][i].split(' ')[2][-1] == 'D' or
                        DF['Descrição da operação'][i].split(' ')[2][-1] == 'S'):
                    DF['REVISAR'][i] = 'MÁQUINA PARADA COM PERIODICIDADE < 1M'
                    DF['PTS'][i] = 0

                # TL sem peridiocidade
                try:
                    if ' FUNC ' not in DF['Descrição da operação'][i]:
                        if not any(car == DF['Descrição da operação'][i].split(' ')[2][-1] for car in
                                   ['H', 'D', 'S', 'M', 'A']):
                            DF['REVISAR'][i] = "SEM PERIODICIDADE"
                            DF['PTS'][i] = -50

                        ## Checar se não faltou outra informação na task list
                        elif len(DF['Descrição da operação'][i]) < 17:
                            DF['REVISAR'][i] = "FALTOU INFORMACAO NO TITULO DA TASK LIST"
                            DF['PTS'][i] = -50

                    elif ' FUNC ' in DF['Descrição da operação'][i]:
                        if not any(car == DF['Descrição da operação'][i].split(' ')[3][-1] for car in
                                   ['H', 'D', 'S', 'M', 'A']):
                            DF['REVISAR'][i] = "SEM PERIODICIDADE"
                            DF['PTS'][i] = -50

                        ## Checar se não faltou outra informação na task list
                        elif len(DF['Descrição da operação'][i]) < 22:
                            DF['REVISAR'][i] = "FALTOU INFORMACAO NO TITULO DA TASK LIST"
                            DF['PTS'][i] = -50

                except:
                    DF['REVISAR'][i] = np.nan


            else:
                # OP sem tempo, CT-PM
                if pd.isna(DF['Trabalho envolvido na atividade'][i]) or pd.isna(DF['Centro de trabalho'][i]) or pd.isna(
                        DF['Descrição da operação'][i]):
                    DF['REVISAR'][i] = "SEM TEMPO OU CENTRO DE TRABALHO"
                    DF['PTS'][i] = -50
                if not pd.isna(DF['Trabalho envolvido na atividade'][i]):
                    if int(DF['Trabalho envolvido na atividade'][i]) == 0:
                        DF['REVISAR'][i] = "SEM TEMPO OU CENTRO DE TRABALHO"
                        DF['PTS'][i] = -50

                # Sem texto longo
                if pd.isna(DF['Texto descritivo de operação'][i]):
                    DF['REVISAR'][i] = "VERIFICAR SE TEXTO LONGO (DETALHADO) É NECESSÁRIO"
                    DF['PTS'][i] = 0

                # OP com num de palavras em Op Padrao abaixo da media?
                if len(str(DF['Descrição da operação'][i]).split()) < media_len_oppadrao:
                    if len(str(DF['Texto descritivo de operação'][i]).split()) < 1.5 * len(
                            str(DF['Descrição da operação'][i]).split()):
                        DF['REVISAR'][i] = "TEXTO DESCRITIVO NÃO DETALHADO"
                        DF['PTS'][i] = 0

                # Suboperação não especifica tipo de atividade?
                if isinstance(DF['Descrição da operação'][i], str):
                    if not any(name in DF['Descrição da operação'][i] for name in
                               ['INSP', 'TROC', 'REV', 'BLOQ', 'DESBLOQ', 'DOSA', 'SUB', 'VERI', 'DESL', 'LIG', 'CONEC',
                                'DESCON', 'REAL', 'LIMP', 'EFET', 'EXEC', 'TEST', 'ISOL', 'RETI', 'ABR', 'FECH', 'LUB',
                                'ESGOT', 'SOLIC', 'ENSA', 'CALI', 'AJUS', 'REAP', 'FIX', 'PARAF', 'INSER',
                                'COLOC']) and not any(
                            letr_2 in DF['Descrição da operação'][i].split()[0][-2:] for letr_2 in ['AR', 'ER', 'IR']):
                        DF['REVISAR'][i] = "VERIFICAR SE SUBOPERACAO ESPECIFICA TIPO DE ATIVIDADE (INSP, REVI, ETC)"
                        DF['PTS'][i] = -10
                if not isinstance(DF['Descrição da operação'][i], str):
                    DF['REVISAR'][i] = "SUBOP VAZIA"
                    DF['PTS'][i] = -50

        df_opativ = DF


        # Salvando em arquivo excel

        import io
        buffer = io.BytesIO()

        with pd.ExcelWriter(buffer, engine="xlsxwriter") as excel_writer:
            ## Crie um objeto ExcelWriter
            nome_arquivo_sap = 'TABELAO_SAP_' + str(df_tl['LI_N3'][0][0:4]) + '-' + str(df_tl['LI_N3'][0][9:12])

            ## Salve cada DataFrame em uma planilha diferente
            df_cabecalho.to_excel(excel_writer, sheet_name='CABECALHO S CALIB LUB', index=False)
            df_opativ.to_excel(excel_writer, sheet_name='TAREFAS S CALIB LUB', index=False)
            df_cabecalho_lub.to_excel(excel_writer, sheet_name='CABECALHO LUB', index=False)
            df_opativ_lub.to_excel(excel_writer, sheet_name='TAREFAS LUB', index=False)

            df_cabecalho_eqp.to_excel(excel_writer, sheet_name='CABECALHO PLANO')  # CABEÇALHO PLANO

            df_tl.to_excel(excel_writer, sheet_name='PM-TABELA', index=False)
            if not df_n6splano.empty:
                df_n6splano.to_excel(excel_writer, sheet_name='EQUIP SEM PLANO N6', index=False)
            if not df_planosobrando.empty:
                df_planosobrando.to_excel(excel_writer, sheet_name='PLANOS SEM EQUIP', index=False)

            df_nc.to_excel(excel_writer, sheet_name='NC TAREFAS S CALIB LUB', index=False)
            df_nc_lub.to_excel(excel_writer, sheet_name='NC TAREFAS LUB', index=False)

            ## Feche o objeto ExcelWriter
            excel_writer.close()

            st.download_button(
                label="Download "+nome_arquivo_sap,
                data=buffer,
                file_name=nome_arquivo_sap+'.xlsx',
            )

            ###
