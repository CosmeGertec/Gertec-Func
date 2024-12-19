import streamlit as st
import pandas as pd
from office365.sharepoint.files.file import File
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
import io


### CONFIGURAÇÕES INICIAS DO STREAMLIT
st.set_page_config('ESTOQUE • STREAMLIT', page_icon='https://i.imgur.com/mOEfCM8.png')

st.header("Funções", divider='grey')

st.sidebar.title('MÓDULOS')
st.sidebar.page_link('main.py', label='FUNÇÕES', disabled=True)
st.sidebar.page_link('pages/fila.py', label='FILA')
st.sidebar.page_link('pages/exp.py', label='EXPEDIÇÃO')

### LINKS ONDE SÃO ARMAZENADOS OS DADOS DO FILA
sharepoint_fila_url = 'https://gertecsao.sharepoint.com/sites/PowerBi-Estoque/'
sharepoint_os_url = 'https://gertecsao.sharepoint.com/sites/RecebimentoLogstica/'
folder_in_sharepoint = '/sites/PowerBi-Estoque/Documentos%20Compartilhados/General/Fila'
#sharepoint_user = st.secrets.sharepoint.USER
#sharepoint_password = st.secrets.sharepoint.SENHA
sharepoint_user = "gertec.visualizador@gertec.com.br"
sharepoint_password = "G@9012hdsOQJH215"

saldo_fila_url = '/sites/PowerBi-Estoque/Documentos%20Compartilhados/General/Fila/saldo.parquet'
varejo_liberado_url = '/sites/PowerBi-Estoque/Documentos%20Compartilhados/General/Fila/Varejo%20Liberado/'
sla_contratos_url = '/sites/PowerBi-Estoque/Documentos%20Compartilhados/General/Fila/SlaContratos.csv'
abertura_os_url = '/sites/RecebimentoLogstica/Documentos%20Compartilhados/General/Recebimento%20-%20Abertura%20de%20OS.xlsx'


### FUNÇÕES
def df_sharep(file_url, tipo='parquet', sheet='', site=sharepoint_fila_url):
    """Gera um DataFrame a partir de um diretório do SharePoint."""
    auth = AuthenticationContext(site)
    auth.acquire_token_for_user(sharepoint_user, sharepoint_password)
    ctx = ClientContext(saldo_fila_url, auth)
    web = ctx.web
    ctx.execute_query()

    file_response = File.open_binary(ctx, file_url)
    bytes_file_obj = io.BytesIO()
    bytes_file_obj.write(file_response.content)
    bytes_file_obj.seek(0)
    if tipo == 'parquet':
        return pd.read_parquet(bytes_file_obj)
    elif tipo == 'csv':
        return pd.read_csv(bytes_file_obj, sep=";")
    elif tipo == 'excel':
        if sheet != '':
            return pd.read_excel(bytes_file_obj, sheet, dtype='str')
        else:
            return pd.read_excel(bytes_file_obj, dtype='str')


def create_df_historico_movimentações():

    # Saldo geral
    historico_fila = df_sharep(saldo_fila_url)

    historico_fila['FLUXO'] = historico_fila['FLUXO'].str.replace('000001', 'CONTRATO')
    historico_fila['FLUXO'] = historico_fila['FLUXO'].str.replace('000002', 'VAREJO')
    historico_fila['FLUXO'] = historico_fila['FLUXO'].str.replace('000003', 'VAREJO')
    historico_fila['FLUXO'] = historico_fila['FLUXO'].str.replace('000004', 'OS INTERNA')

    historico_fila['GARANTIA'] = historico_fila['GARANTIA'].str.upper()
    historico_fila['CLIENTE'] = historico_fila['CLIENTE'].str.upper()

    historico_fila = historico_fila[historico_fila['ENTRADA GERFLOOR'] != 'Nenhum registro encontrado']
    historico_fila['ENTRADA GERFLOOR'] = pd.to_datetime(
        historico_fila.loc[historico_fila['ENTRADA GERFLOOR'] != 'Nenhum registro encontrado', 'ENTRADA GERFLOOR'],
        format='%d/%m/%Y %I:%M:%S %p')

    historico_fila = historico_fila[['ENDEREÇO',
                                     'CAIXA',
                                     'SERIAL',
                                     'CLIENTE',
                                     'EQUIPAMENTO',
                                     'NUM OS',
                                     'FLUXO',
                                     'GARANTIA',
                                     'ENTRADA GERFLOOR',
                                     'ENTRADA FILA',
                                     'SAÍDA FILA']]


    return historico_fila


def create_df_saldo_contratos(df):
    df_saldo_atual_contratos = df.copy()
    df_saldo_atual_contratos = df_saldo_atual_contratos[(df_saldo_atual_contratos['FLUXO'] == 'CONTRATO') & (
        ~df_saldo_atual_contratos['ENDEREÇO'].isin(
            ['LAB', 'EQUIPE TECNICA', 'QUALIDADE', 'RETRIAGEM', 'GESTAO DE ATIVOS']))]

    return df_saldo_atual_contratos


def create_df_saldo_contratos_resumido(df):

    abertura_os = df_sharep(abertura_os_url, 'excel', 'BASE', sharepoint_os_url)
    abertura_os = abertura_os[abertura_os['ABRIR O.S'] != "0"]
    abertura_os.reset_index(drop=True, inplace=True)
    abertura_os.loc[abertura_os['CLIENTE GERFLOOR'].isna(), 'CLIENTE GERFLOOR'] = abertura_os.loc[
        abertura_os['CLIENTE GERFLOOR'].isna(), 'CLIENTES'].apply(lambda x: x.split(" - ", maxsplit=1)[0])
    abertura_os.loc[abertura_os['EQUIPAMENTO GERFLOOR'].isna(), 'EQUIPAMENTO GERFLOOR'] = abertura_os.loc[
        abertura_os['EQUIPAMENTO GERFLOOR'].isna(), 'CLIENTES'].apply(lambda x: x.split(" - ", maxsplit=1)[1])
    abertura_os = abertura_os.rename(columns={'CLIENTE GERFLOOR': 'CLIENTE',
                                              'EQUIPAMENTO GERFLOOR': 'EQUIPAMENTO'}).set_index(
        ['CLIENTE', 'EQUIPAMENTO']).drop(['O.S ABERTA', 'CLIENTES'], axis=1)

    df.loc[df['CLIENTE'].str.startswith('COBRA'), 'CLIENTE'] = 'COBRA'
    df.loc[df['CLIENTE'].str.startswith('BB'), 'CLIENTE'] = 'COBRA'

    df_saldo_atual_contratos_resumido = df.groupby(['CLIENTE', 'EQUIPAMENTO'])[['SERIAL']].count().reset_index()

    df_saldo_atual_contratos_resumido = df_saldo_atual_contratos_resumido.join(abertura_os,
                                                                               on=['CLIENTE', 'EQUIPAMENTO'],
                                                                               how='outer')
    df_saldo_atual_contratos_resumido.loc[df_saldo_atual_contratos_resumido['SERIAL'].isna(), 'SERIAL'] = 0
    df_saldo_atual_contratos_resumido.SERIAL = df_saldo_atual_contratos_resumido.SERIAL.astype(int)
    df_saldo_atual_contratos_resumido.loc[df_saldo_atual_contratos_resumido['ABRIR O.S'].isna(), 'ABRIR O.S'] = 0
    df_saldo_atual_contratos_resumido['ABRIR O.S'] = df_saldo_atual_contratos_resumido['ABRIR O.S'].astype(int)
    df_saldo_atual_contratos_resumido.loc[
        df_saldo_atual_contratos_resumido['DIVERGÊNCIA'].isna(), 'DIVERGÊNCIA'] = 0
    df_saldo_atual_contratos_resumido['DIVERGÊNCIA'] = df_saldo_atual_contratos_resumido['DIVERGÊNCIA'].astype(int)
    df_saldo_atual_contratos_resumido.rename(columns={'SERIAL': 'QTD FILA',
                                                      'ABRIR O.S': 'QTD OS'}, inplace=True)
    df_saldo_atual_contratos_resumido = df_saldo_atual_contratos_resumido[
        ['CLIENTE', 'EQUIPAMENTO', 'QTD OS', 'QTD FILA', 'DIVERGÊNCIA']]
    try:
        df_saldo_atual_contratos_resumido.sort_values(['CLIENTE', 'EQUIPAMENTO'], inplace=True)
    except:
        pass

    return df_saldo_atual_contratos_resumido


def html_saldo_contrato():
    df = create_df_historico_movimentações()
    df = create_df_saldo_contratos(df)
    df = create_df_saldo_contratos_resumido(df)

    df.loc[df['CLIENTE'].str.startswith('COBRA'), 'CLIENTE'] = 'COBRA'
    df.loc[df['CLIENTE'].str.startswith('BB'), 'CLIENTE'] = 'COBRA'
    df.loc[df['CLIENTE'].str.startswith('MERCADO'), 'CLIENTE'] = 'MERCADO PAGO'

    df.loc[df['EQUIPAMENTO'].str.contains('PPC930'), 'EQUIPAMENTO'] = 'PPC930'
    df.loc[df['EQUIPAMENTO'].str.contains('MP35P'), 'EQUIPAMENTO'] = 'MP35P'

    df = df.groupby(['CLIENTE', 'EQUIPAMENTO'])[['QTD OS', 'QTD FILA', 'DIVERGÊNCIA']].sum().reset_index()

    html_contratos = df[['CLIENTE', 'EQUIPAMENTO', 'QTD OS', 'QTD FILA', 'DIVERGÊNCIA']].to_html(index=False,
                                                                                                 index_names=False,
                                                                                                 justify='left',
                                                                                                 na_rep='')
    html_contratos = html_contratos.replace('<table border="1" class="dataframe">',
                                            '<style>\ntable {\n  border-collapse: collapse;\n  width: 100%;\n}\n\nth, td {\n  text-align: center;\n  padding-top: 2px;\n  padding-bottom: 1px;\n  padding-left: 8px;\n  padding-right: 8px;\n}\n\ntr:nth-child(even) {\n  background-color: #DCDCDC;\n}\n\ntable, th, td {\n  border: 2px solid black;\n  border-collapse: collapse;\n}\n</style>\n<table border="1" class="dataframe">')

    return html_contratos


def create_df_varejo_liberado(data_liberacao):
    try:
        auth = AuthenticationContext(sharepoint_fila_url)
        auth.acquire_token_for_user(sharepoint_user, sharepoint_password)
        ctx = ClientContext(saldo_fila_url, auth)
        web = ctx.web
        ctx.execute_query()

        file_response = File.open_binary(ctx, varejo_liberado_url + str(data_liberacao) + ".xlsx")
        bytes_file_obj = io.BytesIO()
        bytes_file_obj.write(file_response.content)
        bytes_file_obj.seek(0)
        df = pd.read_excel(bytes_file_obj,
                           sheet_name='LAB. - SEPARAÇÃO',
                           dtype='str')
        processos_varejo = df
        processos_varejo = processos_varejo[['Nr Serie', 'Num OS', 'Produto_1', 'Client Final', 'Dt Aber. OS']]
        processos_varejo.rename(columns={'Nr Serie': 'SERIAL', 'Num OS': 'NUM OS'}, inplace=True)
        processos_varejo.set_index(['SERIAL', 'NUM OS'], inplace=True)

        varejo_liberado = create_df_historico_movimentações().join(processos_varejo,
                                              on=['SERIAL', 'NUM OS'],
                                              how='right')
        varejo_liberado.loc[varejo_liberado['ENDEREÇO'].isna(), 'CLIENTE'] = varejo_liberado.loc[
            varejo_liberado['ENDEREÇO'].isna(), 'Client Final']
        varejo_liberado.loc[varejo_liberado['ENDEREÇO'].isna(), 'EQUIPAMENTO'] = varejo_liberado.loc[
            varejo_liberado['ENDEREÇO'].isna(), 'Produto_1']
        varejo_liberado.loc[varejo_liberado['ENDEREÇO'].isna(), 'ENTRADA GERFLOOR'] = varejo_liberado.loc[
            varejo_liberado['ENDEREÇO'].isna(), 'Dt Aber. OS']
        varejo_liberado.drop(columns=['Produto_1',
                                      'Dt Aber. OS',
                                      'Client Final',
                                      'FLUXO'], inplace=True)
        varejo_liberado.sort_values('ENDEREÇO', inplace=True)

        st.session_state['data_liberação'] = data_liberacao
        st.session_state['varejo_liberado'] = varejo_liberado
    except:
        varejo_liberado = pd.DataFrame(columns=['!'])
        st.session_state['data_liberação'] = data_liberacao
        st.session_state['varejo_liberado'] = varejo_liberado

    return varejo_liberado


def html_varejo(data_liberacao):
    df = create_df_varejo_liberado(data_liberacao)
    varejo_compactado = df.groupby(['NUM OS', 'CLIENTE', 'ENDEREÇO'])['SERIAL'].count().reset_index().copy()
    varejo_compactado['SERIAL'] = varejo_compactado['SERIAL'].apply(lambda x: "TOTAL: " + str(x))

    df = pd.concat([varejo_compactado, df])
    df['SEPARADO'] = ''
    df = df[['NUM OS', 'SERIAL',
             'CAIXA', 'CLIENTE',
             'EQUIPAMENTO', 'ENDEREÇO',
             'SEPARADO', 'GARANTIA']].sort_values(['ENDEREÇO', 'NUM OS', 'SERIAL'])
    df.loc[
        df['SERIAL'].str.startswith('TOTAL'), ['NUM OS', 'CAIXA', 'CLIENTE', 'EQUIPAMENTO', 'ENDEREÇO', 'SEPARADO',
                                               'GARANTIA']] = ''

    html_content = df.to_html(index=False, index_names=False, justify='left', na_rep='')
    html_content = html_content.replace('<table border="1" class="dataframe">',
                                        '<style>\ntable {\n  border-collapse: collapse;\n  width: 100%;\n}\n\nth, td {\n  text-align: center;\n  padding-top: 2px;\n  padding-bottom: 1px;\n  padding-left: 8px;\n  padding-right: 8px;\n}\n\ntr:nth-child(even) {\n  background-color: #DCDCDC;\n}\n\ntable, th, td {\n  border: 2px solid black;\n  border-collapse: collapse;\n}\n</style>\n<table border="1" class="dataframe">')
    return html_content

st.download_button('BAIXAR RESUMO DE SALDO DOS CONTRATOS (FILA)', html_saldo_contrato(), use_container_width=True,
                           file_name='Contratos.html')

st.write("")

date = st.date_input('Data de liberação do varejo')
try:
    st.download_button('BAIXAR TABELA DE VAREJOS LIBERADOS (FILA)', html_varejo(date), use_container_width=True, file_name=f'Varejo {str(date)}.html')
except:
    st.warning("Sem liberação para a data informada!")