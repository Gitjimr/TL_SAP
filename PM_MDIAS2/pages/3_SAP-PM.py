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
            SAP_EQP = pd.read_excel(my_path + 'SAP_EQP_05-04.xlsx', sheet_name='Sheet1', usecols="A:H", skiprows=0,
                                    dtype=str)
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
            SAP_TL = pd.read_excel(my_path + 'SAP_TL_04-04.xlsx', sheet_name='Sheet1', usecols="A:N", skiprows=0,
                                   dtype=str)
            SAP_TL['CONCAT CENTRO_DESC'] = SAP_TL["Centro planejamento"].map(str, na_action=None) + SAP_TL[
                "Descrição"].map(str, na_action='ignore')
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
            SAP_PMI['CONCAT TL_EQP'] = SAP_PMI["Group"].map(str, na_action=None) + SAP_PMI["Equipment"].map(str,
                                                                                                            na_action='ignore')

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

#*-*-*-*-OK ACIMA