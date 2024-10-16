from dash import Dash, html, dcc, dash_table, Input, Output, State
import pandas as pd
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


    dcc.DatePickerRange(
        id='date-picker-range',
        start_date=datetime.now().date() - timedelta(days=30),
        end_date=datetime.now().date(),
        display_format='DD/MM/YYYY',
        start_date_placeholder_text='Data Inicial',
        end_date_placeholder_text='Data Final'
    ),

    dash_table.DataTable(
        id='table',
        columns=[],
        page_size=25,
        data=[],
        style_data_conditional=[
            {
                'if': {
                    'filter_query': '{Diferença_Horas} >= 20',
                    'column_id': 'Diferença_Horas'
                },
                'backgroundColor': 'green',
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
                    'filter_query': '{Diferença_Horas} < 10',
                    'column_id': 'Diferença_Horas'
                },
                'backgroundColor': 'red',
                'color': 'white'
            },
            {
                'if': {
                    'filter_query': '{Concluidos} = "Concluído"',
                    'column_id': 'Concluidos'
                },
                'backgroundColor': 'green',
                'color': 'white'
            },
            {
                'if': {
                    'filter_query': '{Concluidos} = "Cancelado"',
                    'column_id': 'Concluidos'
                },
                'backgroundColor': 'rgb(250,128,114)',
                'color': 'white'
            },
        ],
    ),

    dcc.Interval(
        id='interval-component',
        interval=60 * 1000,
        n_intervals=0
    ),
    html.Div(id='update-time'),
])

# Callback para atualizar o texto com a hora atualizada
@app.callback(
    Output('update-time', 'children'),
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
    [Input('banco-dropdown', 'value'),
        Input('date-picker-range', 'start_date'),
        Input('date-picker-range', 'end_date'),])


def update_output(value, start_date, end_date):
    # Inserir a Planilha do Bradesco
    if value == "Bradesco":
        # Leitura das planilhas conforme o seu código
        df_bradesco = pd.read_excel('C:\\Users\\thiago.oliveira\\Downloads\\EmAndamento_Atualizado.xlsx')
        df_bradesco = df_bradesco.drop_duplicates()
        df_bradesco_conc = pd.read_excel('C:\\Users\\thiago.oliveira\\Downloads\\Concluidos.xlsx')
        df_numero_proposta = pd.read_excel('C:\\Users\\thiago.oliveira\\Downloads\\bradesco_viva.xlsx')

        # Realizando os merges e transformações como no seu código
        df_bradesco['Concluidos'] = df_bradesco['Solicitação'].map(df_bradesco_conc.set_index('Solicitação')['Situação'].to_dict())
        
        df_numero_proposta['Numero'] = df_numero_proposta['Título'].str.extract(r'^(\d+)').astype(str)
        df_bradesco['Solicitação'] = df_bradesco['Solicitação'].astype(str)
        
        # Junta as planilhas em uma unica
        df_merged = pd.merge(
            df_bradesco[['Solicitação', 'Cidade', 'Vencimentos', 'Concluidos']],
            df_numero_proposta[['Numero', 'Situação', 'Responsável']],
            left_on='Solicitação',
            right_on='Numero',
            how='left'
        )
            
        #Ajusta o formato de hora
        df_merged['Vencimentos'] = pd.to_datetime(df_merged['Vencimentos'], format='%d/%m/%Y', dayfirst= True)
        df_merged['Data'] = df_merged['Vencimentos'].dt.date
        df_merged['Hora'] = df_merged['Vencimentos'].dt.time
        df_merged['Data_Hora'] = df_merged.apply(lambda row: f"{row['Data']} {row['Hora']}", axis=1)

        # Filtro de datas com base no DatePickerRange
        if start_date and end_date:
            start_date = pd.to_datetime(start_date).date()
            end_date = pd.to_datetime(end_date).date()
            df_merged = df_merged[(df_merged['Data'] >= start_date) & (df_merged['Data'] <= end_date)]

        # Verifica a quantidade de tempo restante
        now = datetime.now()
        df_merged['Diferença_Horas'] = df_merged['Vencimentos'].apply(lambda venc: horas_comerciais(now, venc) if not pd.isna(venc) else None)
        df_merged['Diferença_Horas'] = pd.to_numeric(df_merged['Diferença_Horas'], errors='coerce').fillna(0)

        sel_columns = df_merged[['Solicitação', 'Cidade', 'Data_Hora', 'Diferença_Horas', 'Concluidos', 'Situação', 'Responsável']]
        columns = [{"name": i, "id": i} for i in sel_columns.columns]

        return f'Tabela de Vencimentos do Banco: {value}', sel_columns.to_dict('records'), columns



    elif value == "Itaú":
        # Faz a abertura da planilha
        df_itau = pd.read_excel('C:\\Users\\thiago.oliveira\\Downloads\\cetip.xlsx')
        df_numero_proposta = pd.read_excel('C:\\Users\\thiago.oliveira\\Downloads\\itau_viva.xlsx')
        
        df_numero_proposta['Título'] = df_numero_proposta['Título'].astype(str).str.strip()
        df_itau['Nº Controle Interno / Ordem de Serviço'] = df_itau['Nº Controle Interno / Ordem de Serviço'].astype(str).str.strip()
        
        df_numero_proposta['Numero'] = df_numero_proposta['Título'].str.extract(r'^(\d+)').astype(str)
        
        # Realiza o merge entre as tabelas
        df_merged_itau = pd.merge(
            df_itau[['Nº Controle Interno / Ordem de Serviço', 'Cidade', 'Data Vencimento - Empresa de Avaliação', 'Status']],
            df_numero_proposta[['Numero', 'Situação', 'Responsável']],
            left_on='Nº Controle Interno / Ordem de Serviço',
            right_on='Numero',
            how='left'
        )
        
        # Transforma a coluna de vencimento no formato dd/mm/aa hh:mm:ss
        df_merged_itau['Data Vencimento - Empresa de Avaliação'] = pd.to_datetime(
            df_merged_itau['Data Vencimento - Empresa de Avaliação'],
            format='%d/%m/%y %H:%M', dayfirst=True, errors='coerce'
        )
        
        df_merged_itau['Data'] = df_merged_itau['Data Vencimento - Empresa de Avaliação'].dt.date
        df_merged_itau['Hora'] = df_merged_itau['Data Vencimento - Empresa de Avaliação'].dt.time
        df_merged_itau['Data_Hora'] = df_merged_itau.apply(lambda row: f"{row['Data']} {row['Hora']}", axis=1)

        print(df_merged_itau)
        if start_date and end_date:
            start_date = pd.to_datetime(start_date).date()
            end_date = pd.to_datetime(end_date).date()
            df_merged_itau = df_merged_itau[(df_merged_itau['Data'] >= start_date) & (df_merged_itau['Data'] <= end_date)]

        now = datetime.now()
        
        # Cria uma coluna com a quantidade de horas faltantes para o vencimento dos laudos, ignorando valores NaT
        df_merged_itau['Diferença_Horas'] = df_merged_itau['Data Vencimento - Empresa de Avaliação'].apply(
            lambda venc: horas_comerciais(now, venc) if not pd.isna(venc) else None
        )
        
        # Seleciona as colunas que irão aparecer no site, incluindo 'Atribuída' e 'Responsável'
        sel_columns = df_merged_itau[
            ['Nº Controle Interno / Ordem de Serviço', 'Cidade', 'Data_Hora', 'Status', 
            'Diferença_Horas', 'Situação', 'Responsável']
        ]
        
        # Definição das colunas para exibição no site
        columns = [{"name": i, "id": i} for i in sel_columns.columns]

        # Retorna o nome da tabela e os dados para exibição
        return f'Tabela de Vencimentos do Banco: {value}', sel_columns.to_dict('records'), columns  



    #Inseri a planilha do Santander
    elif value == 'Santander':
        # Faz a leitura da planilha do Santander
        df_santander = pd.read_excel('C:\\Users\\thiago.oliveira\\Downloads\\inspectos.xlsx', sheet_name='Crédito imobiliário')
        df_numero_proposta = pd.read_excel('C:\\Users\\thiago.oliveira\\Downloads\\presenciais.xlsx')
        df_santander = df_santander.drop_duplicates()
        
        df_numero_proposta['Numero'] = df_numero_proposta['Título'].str.extract(r'^(\d+)').astype(str)
        
        df_santander['Nro. Proposta'] = df_santander['Nro. Proposta'].astype(str)
        
        # Realiza o merge entre as tabelas
        df_merged = pd.merge(
            df_numero_proposta[['Numero', 'Situação', 'Responsável']],
            df_santander[['Nro. Proposta', 'Município', 'Data Limite', 'Status']],
            left_on='Numero',
            right_on='Nro. Proposta',
            how = 'left'
        )
        
        # Separa as colunas que irei mostrar no site
        sel_columns = df_merged[
            ['Nro. Proposta', 'Município','Status', 'Data Limite','Situação', 'Responsável']]
        
        columns = [{"name": i, "id": i} for i in sel_columns.columns]

        # Retorna no topo o nome da tabela que esta sendo exibida
        return f'Tabela de Vencimentos do Banco: {value}', sel_columns.to_dict('records'), columns
    
    return f'Selecionado: {value}', [], []

if __name__ == '__main__':
    app.run_server(debug=True)
