import streamlit as st
import pandas as pd
from office365.sharepoint.files.file import File
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
import io
from datetime import datetime, timedelta


### Autenticação ao Sharepoint

sharepoint_base_url = 'https://gertecsao.sharepoint.com/sites/PowerBi-Estoque/'
folder_in_sharepoint = '/sites/PowerBi-Estoque/Documentos%20Compartilhados/General/Expedi%C3%A7%C3%A3o/Planilhas/'
sharepoint_user = st.secrets.sharepoint.USER
sharepoint_password = st.secrets.sharepoint.SENHA

auth = AuthenticationContext(sharepoint_base_url)
auth.acquire_token_for_user(sharepoint_user, sharepoint_password)
ctx = ClientContext(sharepoint_base_url, auth)
web = ctx.web
ctx.execute_query()

### Links Úteis

saldo_exp_url = '/sites/PowerBi-Estoque/Documentos%20Compartilhados/General/Expedi%C3%A7%C3%A3o/Bases%20de%20Dados/saldo_exp.parquet'  # Arquivo Parquest
afericao_rom_url = '/sites/PowerBi-Estoque/Documentos%20Compartilhados/General/Expedi%C3%A7%C3%A3o/Planilhas/Aferi%C3%A7%C3%A3o%20de%20Romaneios.xlsx'  # Planilha em xlsx
stone_bo = '/sites/PowerBi-Estoque/Documentos%20Compartilhados/General/Expedi%C3%A7%C3%A3o/Planilhas/STONE.xlsx'  # Planilha em xlsx
metas_lab_url = '/sites/PowerBi-Estoque/Documentos%20Compartilhados/General/Expedi%C3%A7%C3%A3o//M%C3%A9tricas/Metas%20Diarias.csv'  # Criar arquivo em CSV


### Funções

def df_sharep(file_url, header=0, format='parquet'):
    """Gera um DataFrame a partir de um diretório do SharePoint."""
    file_response = File.open_binary(ctx, file_url)
    bytes_file_obj = io.BytesIO()
    bytes_file_obj.write(file_response.content)
    bytes_file_obj.seek(0)
    if format == 'parquet':
        return pd.read_parquet(bytes_file_obj)
    elif format == 'csv':
        return pd.read_csv(bytes_file_obj, header=header, sep=';')
    elif format == 'excel':
        return pd.read_excel(bytes_file_obj, header=header, dtype='str')


st.set_page_config('ESTOQUE • STREAMLIT', page_icon='https://i.imgur.com/mOEfCM8.png', layout='wide')

st.header("Liberações diárias do laboratório", divider='grey')

r0c1, r0c2, r0c3, r0c4 = st.columns(4, gap='large')
st.write("")

st.sidebar.title('MÓDULOS')
st.sidebar.page_link('main.py', label='FUNÇÕES')
st.sidebar.page_link('pages/fila.py', label='FILA')
st.sidebar.page_link('pages/exp.py', label='EXPEDIÇÃO', disabled=True)

cont = r0c1.checkbox('CONTRATOS', value=True)
varj = r0c2.checkbox('VAREJOS', value=False)
osin = r0c3.checkbox('OS INTERNA', value=False)

filtro = ["000001" if cont else "^",
          "000002" if varj else "^",
          "000003" if varj else "^",
          "000004" if osin else "^"]

saldo_exp_df = df_sharep(saldo_exp_url)
saldo_exp_df = saldo_exp_df.loc[saldo_exp_df['Fluxo'].isin(filtro)]
saldo_exp_df.loc[saldo_exp_df["Dt Entrada"] != "Nenhum registro encontrado", "Dt Entrada"] = pd.to_datetime(
    saldo_exp_df.loc[saldo_exp_df["Dt Entrada"] != "Nenhum registro encontrado", "Dt Entrada"])

date = r0c4.date_input('DATA ENTRADA PRÉ-EXPEDIÇÃO', value=max(saldo_exp_df['ENTRADA PRÉ-EXPEDIÇÃO']), format='DD/MM/YYYY')

saldo_exp_df = saldo_exp_df[saldo_exp_df['ENTRADA PRÉ-EXPEDIÇÃO'] >= pd.to_datetime(date)]
saldo_exp_df = saldo_exp_df[saldo_exp_df['ENTRADA PRÉ-EXPEDIÇÃO'] <= pd.to_datetime(date) + timedelta(days=1)]
saldo_exp_resumido_caixa_df = saldo_exp_df.groupby(['Client Final', 'Desc Prod'])[['Nr Serie']].count()
saldo_exp_resumido_equip_df = saldo_exp_df.groupby(['Client Final', 'Desc Prod', 'CAIXA'])[
    ['Nr Serie']].count().reset_index()
saldo_exp_resumido_equip_df = saldo_exp_resumido_equip_df.groupby(['Client Final', 'Desc Prod'])[
    ['CAIXA']].count().reset_index()

saldo_exp_resumido_df = saldo_exp_resumido_equip_df.join(saldo_exp_resumido_caixa_df,
                                                         on=['Client Final', 'Desc Prod'])

saldo_exp_resumido_df.columns = ['CLIENTE', 'EQUIPAMENTO', 'QTD CAIXAS', 'QTD TERMINAIS']

st.dataframe(saldo_exp_resumido_df, hide_index=True, use_container_width=True)
