from dash import Dash, html, dcc, dash_table, Input, Output
import pandas as pd
import numpy as np
from datetime import datetime, time, timedelta

def horas_comerciais(inicio, fim, hora_abertura=time(8, 0), hora_fechamento=time(18, 0)):
    if inicio > fim:
        return 0

    # Ajustar os limites ao horário comercial
    if inicio.time() < hora_abertura:
        inicio = datetime.combine(inicio.date(), hora_abertura)
    if inicio.time() > hora_fechamento:
        inicio = datetime.combine(inicio.date() + timedelta(days=1), hora_abertura)
    if fim.time() > hora_fechamento:
        fim = datetime.combine(fim.date(), hora_fechamento)
    if fim.time() < hora_abertura:
        fim = datetime.combine(fim.date() - timedelta(days=1), hora_fechamento)

    if inicio.date() == fim.date():
        return max(0, (fim - inicio).total_seconds() / 3600)

    primeiro_dia = datetime.combine(inicio.date(), hora_fechamento) - inicio
    horas_primeiro_dia = primeiro_dia.total_seconds() / 3600

    ultimo_dia = fim - datetime.combine(fim.date(), hora_abertura)
    horas_ultimo_dia = ultimo_dia.total_seconds() / 3600

    dias_completos = max(0, (fim.date() - inicio.date()).days - 1)
    horas_completas = dias_completos * (hora_fechamento.hour - hora_abertura.hour)

    return round(horas_primeiro_dia + horas_completas + horas_ultimo_dia, 2)

app = Dash(__name__)

app.layout = html.Div(children=[
    dcc.Dropdown(['Bradesco', 'Santander', 'Itaú'], 'Bradesco', id='banco-dropdown'),
    html.H1(id='banco'),

    dash_table.DataTable(
        id='table',
        columns=[],
        page_size=25,
        data=[],
        style_data_conditional=[
            {
                'if': {
                    'filter_query': '{Diferença_Horas} < 10',
                    'column_id': 'Diferença_Horas'
                },
                'backgroundColor': 'red',
                'color': 'white'
            },

            {
                'if': {
                    'filter_query': '{Diferença_Horas} > 10',
                    'column_id': 'Diferença_Horas'
                },
                'backgroundColor': 'yellow',
                'color': 'black'
            },

            {
                'if': {
                    'filter_query': '{Diferença_Horas} >= 20',
                    'column_id': 'Diferença_Horas'
                },
                'backgroundColor': 'green',
                'color': 'white'
            },
        ]
    ),

    dcc.Interval(
        id='interval-component',
        interval=60 * 1000,
        n_intervals=0
    ),

    html.Div(id='live-update-text'),
])

# Callback para atualizar o texto com a hora atualizada
@app.callback(
    Output('live-update-text', 'children'),
    [Input('interval-component', 'n_intervals')]
)
def update_layout(n):
    now = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    return f'Última atualização: {now}'

# Callback para atualizar o banco selecionado e a tabela de dados
@app.callback(
    [Output('banco', 'children'),
        Output('table', 'data'),
        Output('table', 'columns')],
    [Input('banco-dropdown', 'value')]
)
def update_output(value):
    #Inseri a Planilha do Bradesco
    if value == "Bradesco":
        #Abertura das planilha do Bradesco
        df_bradesco = pd.read_excel('M:\\Thiago\\Controle de Laudos\\Dados.xlsx', sheet_name='Bradesco')
        df_bradesco = df_bradesco.drop_duplicates()
        df_bradesco_conc = pd.read_excel('M:\\Thiago\\Controle de Laudos\\Dados.xlsx', sheet_name='Bradesco Concluidos')

        #Faz uma especie de 'PROCV' na planilha bradesco_concluidos para á planilha bradesco
        solicitacao_para_situacao = df_bradesco_conc.set_index('Solicitação')['Situação'].to_dict()
        df_bradesco['Concluidos'] = df_bradesco['Solicitação'].map(solicitacao_para_situacao)
        
        #Tranforma a data em formato dd/mm/aa hh:mm:ss
        df_bradesco['Vencimento'] = pd.to_datetime(df_bradesco['Vencimento'], format = '%d/%m/%y %H:%M') # Não está funcionando essa mudança, irei verificar na planilha se é algum erro de formatação nela
        now = datetime.now()

        # Cria uma coluna com a quantidade de horas faltantes para vencer o laudo
        df_bradesco['Diferença_Horas'] = df_bradesco['Vencimento'].apply(lambda venc: horas_comerciais(now, venc))

        # Seleciona as colunas que irei utilizar na visualização e imprimi na tela
        sel_columns = df_bradesco[['Solicitação', 'Cidade', 'Vencimento','Concluidos', 'Diferença_Horas']]
        columns = [{"name": i, "id": i} for i in sel_columns.columns]

        # Retorna no topo o nome da tabela que esta sendo exibida
        return f'Tabela de Vencimentos do Banco: {value}', sel_columns.to_dict('records'), columns

    #Inseri a planilha do Itáu
    elif value == "Itaú":
        # Faz a abertura da planilha
        df_itau = pd.read_excel('M:\\Thiago\\Controle de Laudos\\Dados.xlsx', sheet_name='Cetip')

        # Transforma a coluna de vencimento no formato dd/mm/aa hh:mm:ss
        df_itau['Data Vencimento - Empresa de Avaliação'] = pd.to_datetime(df_itau['Data Vencimento - Empresa de Avaliação'],format='mixed', dayfirst=True) # Tambem não está funcionando, irei verificar na planilha se é algum erro de formatação nela
        now = datetime.now()

        # Cria uma coluna com a quantiade de horas faltantes para o vencimento dos laudos
        df_itau['Diferença_Horas'] = df_itau['Data Vencimento - Empresa de Avaliação'].apply(lambda venc: horas_comerciais(now, venc))

        # Seleciona as colunas que irão aparecer no site
        sel_columns = df_itau[
            ['Nº Controle Interno / Ordem de Serviço', 'Cidade', 'Data Vencimento - Empresa de Avaliação','Status','Diferença_Horas']]
        columns = [{"name": i, "id": i} for i in sel_columns.columns]

        # Retorna no topo o nome da tabela que esta sendo exibida
        return f'Tabela de Vencimentos do Banco: {value}', sel_columns.to_dict('records'), columns

    #Inseri a planilha do Santander
    elif value == 'Santander':
        # Faz a leitura da planilha do Santander
        df_santander = pd.read_excel('M:\\Thiago\\Controle de Laudos\\Dados.xlsx', sheet_name='Inspectos')
        df_santander = df_santander.drop_duplicates()
        
        # Separa as colunas que irei mostrar no site
        sel_columns = df_santander[
            ['Nro. Proposta', 'Município','Status', 'Data Limite']]
        columns = [{"name": i, "id": i} for i in sel_columns.columns]

        # Retorna no topo o nome da tabela que esta sendo exibida
        return f'Tabela de Vencimentos do Banco: {value}', sel_columns.to_dict('records'), columns
    return f'Selecionado: {value}', [], []

if __name__ == '__main__':
    app.run_server(debug=True)
