import pandas as pd
import numpy as np
import warnings
import plotly.express as px
import plotly.subplots as sp
import plotly.graph_objs as go
import datetime
import time
from flask import Flask, render_template, send_file
import io
import plotly.io as pio
import logging

app = Flask(__name__)


# Leitura das dataframes
df_eam = pd.read_excel('eam.xlsx')
df_custo = pd.read_excel('custo.xlsx')
df_pcfactory = pd.read_excel('pcfactory.xlsx')

# #Manipulação dos dados - Custo
#Corrigir espaços e maiúsculas por colunas
df_custo['Usuario'] = df_custo['Usuario'].apply(lambda x: x.lower() if isinstance(x, str) else x)
df_custo['Usuario'] = df_custo['Usuario'].apply(lambda x: x.replace(" ", "") if isinstance(x, str) else x)
df_custo['Conta'] = df_custo['Conta'].apply(lambda x: x.lower() if isinstance(x, str) else x)
df_custo['Conta'] = df_custo['Conta'].apply(lambda x: x.replace(" ", "") if isinstance(x, str) else x)
#Verificar o usuario e separar por manutenção
df_custo.loc[df_custo['Usuario'] == 'pcmbtu', 'Usuario'] = 'manutenção'
df_custo.loc[df_custo['Usuario'] == 'sergiob', 'Usuario'] = 'manutenção'
df_custo.loc[df_custo['Usuario'] == 'cledson', 'Usuario'] = 'manutenção'
df_custo.loc[df_custo['Usuario'] == 'felipini', 'Usuario'] = 'manutenção'
df_custo.loc[df_custo['Usuario'] == 'ferram-btu', 'Usuario'] = 'manutenção'
df_custo.loc[df_custo['Usuario'] == 'minichello', 'Usuario'] = 'manutenção'
df_custo.loc[df_custo['Usuario'] == 'cignoni', 'Usuario'] = 'manutenção'
df_custo.loc[df_custo['Usuario'] == 'ptavares', 'Usuario'] = 'manutenção'
df_custo.loc[df_custo['Usuario'] == 'silesio', 'Usuario'] = 'manutenção'
df_custo.loc[df_custo['Usuario'] == 'marcioc', 'Usuario'] = 'manutenção'
df_custo.loc[df_custo['Usuario'] == 'rodrigocesar', 'Usuario'] = 'manutenção'
df_custo.loc[df_custo['Usuario'] == 'catharino', 'Usuario'] = 'manutenção'
df_custo.loc[df_custo['Usuario'] == 'vernini', 'Usuario'] = 'manutenção'
df_custo.loc[df_custo['Usuario'] == 'lucianom', 'Usuario'] = 'manutenção'
df_custo.loc[df_custo['Usuario'] == 'lfedele', 'Usuario'] = 'manutenção'
#df_custo.loc[df_custo['Usuario'] == 'mfrederik', 'Usuario'] = 'manutenção' #colocar o do joão quando tiver
df_custo.loc[df_custo['Usuario'] == 'albert', 'Usuario'] = 'manutenção'
#Verificacao do codigo
df_custo.loc[df_custo['Conta'] == 13201033, 'Conta'] = 'IAF'
df_custo.loc[df_custo['Conta'] == 10821508, 'Conta'] = '13J'
df_custo.loc[df_custo['Conta'] == 41405000, 'Conta'] = '1J'
df_custo.loc[df_custo['Conta'] == 41651000, 'Conta'] = '2J'
df_custo.loc[df_custo['Conta'] == 41663010, 'Conta'] = '3J'
df_custo.loc[df_custo['Conta'] == 41650000, 'Conta'] = '4J'
df_custo.loc[df_custo['Conta'] == 41663009, 'Conta'] = '5J'
df_custo.loc[df_custo['Conta'] == 41649000, 'Conta'] = '6J'
df_custo.loc[df_custo['Conta'] == 41663008, 'Conta'] = '7J'
df_custo.loc[df_custo['Conta'] == 41653000, 'Conta'] = '8J'
df_custo.loc[df_custo['Conta'] == 41663011, 'Conta'] = '9J'
df_custo.loc[df_custo['Conta'] == 41663002, 'Conta'] = '11J'
df_custo.loc[df_custo['Conta'] == 41647000, 'Conta'] = '12J'
df_custo.loc[df_custo['Conta'] == 41647000, 'Conta'] = '13J'
df_custo.loc[df_custo['Conta'] == 41655000, 'Conta'] = '14J'
df_custo.loc[df_custo['Conta'] == 41663099, 'Conta'] = '15J'
df_custo.loc[df_custo['Conta'] == 41663007, 'Conta'] = '16J'
df_custo.loc[df_custo['Conta'] == 41625001, 'Conta'] = '17J'
df_custo.loc[df_custo['Conta'] == 41625004, 'Conta'] = '18J'
df_custo.loc[df_custo['Conta'] == 41625005, 'Conta'] = '19J'
df_custo.loc[df_custo['Conta'] == 41655001, 'Conta'] = '20J'
df_custo.loc[df_custo['Conta'] == 41606000, 'Conta'] = '21J'
df_custo.loc[df_custo['Conta'] == 41607000, 'Conta'] = '22J'
df_custo.loc[df_custo['Conta'] == 41647002, 'Conta'] = '23J'
df_custo.loc[df_custo['Conta'] == 41663016, 'Conta'] = '24J'
df_custo.loc[df_custo['Conta'] == 21747008, 'Conta'] = 'Elimina'
df_custo.loc[df_custo['Conta'] == 41505000, 'Conta'] = 'Elimina'
df_custo.loc[df_custo['Conta'] == 41663005, 'Conta'] = 'Elimina'
df_custo.loc[df_custo['CC'] == 12050, 'CC'] = 'Movimentação de madeira'
df_custo.loc[df_custo['CC'] == 12090, 'CC'] = 'Picador'
df_custo.loc[df_custo['CC'] == 12200, 'CC'] = 'Moinhos'
df_custo.loc[df_custo['CC'] == 12250, 'CC'] = 'Secadores'
df_custo.loc[df_custo['CC'] == 12300, 'CC'] = 'Encoladeira'
df_custo.loc[df_custo['CC'] == 12400, 'CC'] = 'Prensa HD'
df_custo.loc[df_custo['CC'] == 12410, 'CC'] = 'Laqueadora'
df_custo.loc[df_custo['CC'] == 12420, 'CC'] = 'Impregnadora'
df_custo.loc[df_custo['CC'] == 12430, 'CC'] = 'Tocchio'
df_custo.loc[df_custo['CC'] == 12506, 'CC'] = 'Barberan'
df_custo.loc[df_custo['CC'] == 12507, 'CC'] = 'Wemhoner'
df_custo.loc[df_custo['CC'] == 12508, 'CC'] = 'Siempelkamp'
df_custo.loc[df_custo['CC'] == 12510, 'CC'] = 'Torwegge'
df_custo.loc[df_custo['CC'] == 12520, 'CC'] = 'Serra Giben'
df_custo.loc[df_custo['CC'] == 12530, 'CC'] = 'Homag'
df_custo.loc[df_custo['CC'] == 12540, 'CC'] = 'Acessórios'
df_custo.loc[df_custo['CC'] == 12550, 'CC'] = 'Extrusoras'
df_custo.loc[df_custo['CC'] == 12600, 'CC'] = 'Lixadeira'
df_custo.loc[df_custo['CC'] == 12800, 'CC'] = 'Cyklop'
df_custo.loc[df_custo['CC'] == 13900, 'CC'] = 'Utilidades'
df_custo.loc[df_custo['CC'] == 16090, 'CC'] = 'Projetos'
df_custo.loc[df_custo['CC'] == 16100, 'CC'] = 'Manutenção mecânica'
df_custo.loc[df_custo['CC'] == 16200, 'CC'] = 'Manutenção de linhas'
df_custo.loc[df_custo['CC'] == 16600, 'CC'] = 'PCM'
df_custo.loc[df_custo['CC'] == 16700, 'CC'] = 'Manutenção elétrica'
df_custo.loc[df_custo['CC'] == 16850, 'CC'] = 'Construção e conservação fabril'
for index, row in df_custo.iterrows():
    if row['Usuario'] != 'manutenção':
        df_custo.loc[index, 'Conta'] = 'Elimina'



def gastos_setores_4j (df_custo): #Gastos por setores 4J
    Setores = df_custo[df_custo['Conta'] == '4J'].groupby('CC')['Val.Item Req.'].sum().reset_index()
    cc_desejados = ['Movimentação de madeira', 'Picador', 'Moinhos', 'Secadores', 'Encoladeira', 'Prensa HD', 'Laqueadora', 'Impregnadora', 'Tocchio', 'Barberan', 'Wemhoner', 'Siempelkamp', 'Torwegge', 'Serra Giben', 'Homag', 'Extrusoras', 'Lixadeira', 'Cyklop', 'Utilidades', 'Projetos', 'Manutenção mecânica', 'Manutenção de linhas', 'PCM','Manutenção elétrica','Construção e conservação fabril']
    Setores_filtrados = Setores[Setores['CC'].isin(cc_desejados) & (Setores['Val.Item Req.'] > 0)]
    Setores_filtrados = Setores_filtrados.sort_values(by='Val.Item Req.', ascending=False)
    fig_gastos_4j = sp.make_subplots(rows=1, cols=1, subplot_titles=["Gastos por setores - 41650000"])
    trace = go.Bar(x=Setores_filtrados['CC'], y=Setores_filtrados['Val.Item Req.'],
                text=Setores_filtrados['Val.Item Req.'].apply(lambda x: f'R$ {x:.2f}'),
                textposition='outside')
    fig_gastos_4j.add_trace(trace)
    fig_gastos_4j.update_xaxes(title_text='')
    fig_gastos_4j.update_yaxes(title_text='', showticklabels=False)
    fig_gastos_4j.update_layout(showlegend=False, margin=dict(l=0, r=0, t=40, b=0), autosize=True)
    fig_gastos_4j.update_layout(
        showlegend=False,
        margin=dict(l=0, r=0, t=40, b=0),
        autosize=True,
        height=1080,
        width=1920,
    )
    return pio.to_html(fig_gastos_4j, full_html=False)

def gastos_setores_5j (df_custo): #Gastos por setores 5J
    Setores = df_custo[df_custo['Conta'] == '5J'].groupby('CC')['Val.Item Req.'].sum().reset_index()
    cc_desejados = ['Movimentação de madeira', 'Picador', 'Moinhos', 'Secadores', 'Encoladeira', 'Prensa HD', 'Laqueadora', 'Impregnadora', 'Tocchio', 'Barberan', 'Wemhoner', 'Siempelkamp', 'Torwegge', 'Serra Giben', 'Homag', 'Extrusoras', 'Lixadeira', 'Cyklop', 'Utilidades', 'Projetos', 'Manutenção mecânica', 'Manutenção de linhas', 'PCM','Manutenção elétrica','Construção e conservação fabril']
    Setores_filtrados = Setores[Setores['CC'].isin(cc_desejados) & (Setores['Val.Item Req.'] > 0)]
    Setores_filtrados = Setores_filtrados.sort_values(by='Val.Item Req.', ascending=False)
    fig_gastos_5j = sp.make_subplots(rows=1, cols=1, subplot_titles=["Gastos por setores - 41663009"])
    trace = go.Bar(x=Setores_filtrados['CC'], y=Setores_filtrados['Val.Item Req.'],
                text=Setores_filtrados['Val.Item Req.'].apply(lambda x: f'R$ {x:.2f}'),
                textposition='outside')
    fig_gastos_5j.add_trace(trace)
    fig_gastos_5j.update_xaxes(title_text='')
    fig_gastos_5j.update_yaxes(title_text='', showticklabels=False)
    fig_gastos_5j.update_layout(showlegend=False, margin=dict(l=0, r=0, t=40, b=0), autosize=True)
    fig_gastos_5j.update_layout(
        showlegend=False,
        margin=dict(l=0, r=0, t=40, b=0),
        autosize=True,
        height=1080,
        width=1920,
    )
    return pio.to_html(fig_gastos_5j, full_html=False)

# def custo_geral (df_custo): #Tabela com todas as contas
#     df_custo_filtrado = df_custo[(df_custo['Conta'] != 'Elimina') & (df_custo['Conta'] != 'IAF')].copy()
#     Contas = df_custo_filtrado.groupby('Conta')['Val.Item Req.'].sum()
#     Valor_maximo = {'1J': 18901, '2J': 35000, '3J': 70000, '4J': 460000, '5J': 250000, '6J': 2500, '7J': 7500,
#                     '8J': 4000, '9J': 15000, '11J': 10000, '12J': 28000, '13J': 35000, '14J': 2000, '15J': 20000,
#                     '16J': 10000, '17J': 1500, '18J': 3000, '19J': 3000, '20J': 2000, '21J': 10000, '22J': 2000,
#                     '23J': 1000, '24J': 8500}
#     Contas_com_maximo = Contas.reset_index().assign(Orçado=Contas.index.map(Valor_maximo))
#     Contas_com_maximo['Saldo'] = Contas_com_maximo['Orçado'] - Contas_com_maximo['Val.Item Req.']
#     Contas_com_maximo = Contas_com_maximo.rename(columns={'Val.Item Req.': 'Utilizado'})
#     ordem_contas = ['1J', '2J', '3J', '4J', '5J', '6J', '7J', '8J', '9J', '11J', '12J', '13J', '14J', '15J', '16J', '17J', '18J', '19J', '20J', '21J', '22J', '23J', '24J']
#     Contas_com_maximo['Conta'] = pd.Categorical(Contas_com_maximo['Conta'], categories=ordem_contas, ordered=True)
#     Contas_com_maximo = Contas_com_maximo.sort_values('Conta')
#     Contas_com_maximo['Disp %'] = ((Contas_com_maximo['Orçado'] - Contas_com_maximo['Utilizado']) / Contas_com_maximo['Orçado']) * 100
#     Contas_com_maximo['Disp %'] = Contas_com_maximo['Disp %'].astype(int)
#     Contas_com_maximo.reset_index(inplace=True)
#     return pio.to_html(Contas_com_maximo, full_html=False)
    
# #Manipulação dos dados - EAM
# Troca do RM para o nome
df_eam.loc[df_eam['Supervisor'] == 'RM007', 'Supervisor'] = 'Felipini'
df_eam.loc[df_eam['Supervisor'] == 'RM008', 'Supervisor'] = 'Sergio'
df_eam.loc[df_eam['Supervisor'] == 'RM011', 'Supervisor'] = 'Albert'
df_eam.loc[df_eam['Supervisor'] == 'RM012', 'Supervisor'] = 'Marcio'
df_eam.loc[df_eam['Supervisor'] == 'RM013', 'Supervisor'] = 'Adilson'
df_eam.loc[df_eam['Supervisor'] == 'RM021', 'Supervisor'] = 'Rodrigo'
df_eam.loc[df_eam['Supervisor'] == 'RM022', 'Supervisor'] = 'Gilson'
df_eam.loc[df_eam['Supervisor'] == 'RM023', 'Supervisor'] = 'Cignoni'
df_eam.loc[df_eam['Supervisor'] == 'RM029', 'Supervisor'] = 'Cledson'
df_eam.loc[df_eam['Supervisor'] == 'RM033', 'Supervisor'] = 'Lucas'
df_eam.loc[df_eam['Supervisor'] == 'RM073', 'Supervisor'] = 'Luigi'

# Homem hora por ordem de serviço
df_eam['Homem hora'] = (df_eam['Horas estimadas'] * df_eam['Pessoal requerido'])

# Separação por setor
Local = df_eam['Local do serviço'].str.lower().str.contains("picador", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Picador"
Local = df_eam['Local do serviço'].str.lower().str.contains("moinhos", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Moinhos"
Local = df_eam['Local do serviço'].str.lower().str.contains("secadores", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Secadores"
Local = df_eam['Local do serviço'].str.lower().str.contains("linha hydrodyn", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Linha Hydrodyn"
Local = df_eam['Local do serviço'].str.lower().str.contains("lixadeira", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Lixadeira Bison"
Local = df_eam['Local do serviço'].str.lower().str.contains("linha de classificação", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Linha de Classificação"
Local = df_eam['Local do serviço'].str.lower().str.contains("laqueadora", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Laqueadora Vits"
Local = df_eam['Local do serviço'].str.lower().str.contains("impregnadora 1", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Impregnadora Vits"
Local = df_eam['Local do serviço'].str.lower().str.contains("impregnadora 2", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Impregnadora Tocchio"
Local = df_eam['Local do serviço'].str.lower().str.contains("preparação química", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Preparação Química"
Local = df_eam['Local do serviço'].str.lower().str.contains("prensa ciclo curto 1", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Wemhoner"
Local = df_eam['Local do serviço'].str.lower().str.contains("prensa ciclo curto 2", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Siempelkamp"
Local = df_eam['Local do serviço'].str.lower().str.contains("laminadora", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Barberãn"
Local = df_eam['Local do serviço'].str.lower().str.contains("torwegge", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Torwegge"
Local = df_eam['Local do serviço'].str.lower().str.contains("cyklop", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Cyklop"
Local = df_eam['Local do serviço'].str.lower().str.contains("voma", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Voma"
Local = df_eam['Local do serviço'].str.lower().str.contains("tecno", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Extrusora Tecno"
Local = df_eam['Local do serviço'].str.lower().str.contains("wpc 1", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Extrusora WPC 1"
Local = df_eam['Local do serviço'].str.lower().str.contains("wpc 2", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Extrusora WPC 2"
Local = df_eam['Local do serviço'].str.lower().str.contains("torwegge 3", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Torwegge 3"
Local = df_eam['Local do serviço'].str.lower().str.contains("utilidades", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Utilidades"
Local = df_eam['Local do serviço'].str.lower().str.contains("sistema de proteção contra incêndios", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Sistema de proteção contra incêndios"
Local = df_eam['Local do serviço'].str.lower().str.contains("sistema elétrico", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Sistema Elétrico"
Local = df_eam['Local do serviço'].str.lower().str.contains("predial", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Predial"
Local = df_eam['Local do serviço'].str.lower().str.contains("fábrica de acessórios", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Fábrica de acessórios"
Local = df_eam['Local do serviço'].str.lower().str.contains("exaustão geral e processamento de pó", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Exaustão geral e processamento de pó"
Local = df_eam['Local do serviço'].str.lower().str.contains("Pisos e acessórios", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Pisos e acessórios"
Local = df_eam['Local do serviço'].str.lower().str.contains("Unidade Painéis e Pisos", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Unidade Painéis e Pisos"
#===============================================================================================================
Local = df_eam['Descrição do equipamento'].str.lower().str.contains("picador", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Picador"
Local = df_eam['Descrição do equipamento'].str.lower().str.contains("moinhos", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Moinhos"
Local = df_eam['Descrição do equipamento'].str.lower().str.contains("secadores", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Secadores"
Local = df_eam['Descrição do equipamento'].str.lower().str.contains("linha hydrodyn", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Linha Hydrodyn"
Local = df_eam['Descrição do equipamento'].str.lower().str.contains("lixadeira", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Lixadeira Bison"
Local = df_eam['Descrição do equipamento'].str.lower().str.contains("linha de classificação", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Linha de Classificação"
Local = df_eam['Descrição do equipamento'].str.lower().str.contains("laqueadora", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Laqueadora Vits"
Local = df_eam['Descrição do equipamento'].str.lower().str.contains("impregnadora 1", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Impregnadora Vits"
Local = df_eam['Descrição do equipamento'].str.lower().str.contains("impregnadora 2", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Impregnadora Tocchio"
Local = df_eam['Descrição do equipamento'].str.lower().str.contains("preparação química", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Preparação Química"
Local = df_eam['Descrição do equipamento'].str.lower().str.contains("prensa ciclo curto 1", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Wemhoner"
Local = df_eam['Descrição do equipamento'].str.lower().str.contains("prensa ciclo curto 2", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Siempelkamp"
Local = df_eam['Descrição do equipamento'].str.lower().str.contains("laminadora", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Barberãn"
Local = df_eam['Descrição do equipamento'].str.lower().str.contains("torwegge", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Torwegge"
Local = df_eam['Descrição do equipamento'].str.lower().str.contains("cyklop", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Cyklop"
Local = df_eam['Descrição do equipamento'].str.lower().str.contains("voma", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Voma"
Local = df_eam['Descrição do equipamento'].str.lower().str.contains("tecno", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Extrusora Tecno"
Local = df_eam['Descrição do equipamento'].str.lower().str.contains("wpc 1", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Extrusora WPC 1"
Local = df_eam['Descrição do equipamento'].str.lower().str.contains("wpc 2", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Extrusora WPC 2"
Local = df_eam['Descrição do equipamento'].str.lower().str.contains("torwegge 3", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Torwegge 3"
Local = df_eam['Descrição do equipamento'].str.lower().str.contains("utilidades", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Utilidades"
Local = df_eam['Descrição do equipamento'].str.lower().str.contains("sistema de proteção contra incêndios", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Sistema de proteção contra incêndios"
Local = df_eam['Descrição do equipamento'].str.lower().str.contains("sistema elétrico", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Sistema Elétrico"
Local = df_eam['Descrição do equipamento'].str.lower().str.contains("predial", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Predial"
Local = df_eam['Descrição do equipamento'].str.lower().str.contains("fábrica de acessórios", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Fábrica de acessórios"
Local = df_eam['Descrição do equipamento'].str.lower().str.contains("exaustão geral e processamento de pó", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Exaustão geral e processamento de pó"
Local = df_eam['Descrição do equipamento'].str.lower().str.contains("Pisos e acessórios", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Pisos e acessórios"
Local = df_eam['Descrição do equipamento'].str.lower().str.contains("Unidade Painéis e Pisos", na=False)
df_eam.loc[Local ,'Local do serviço'] = "Unidade Painéis e Pisos"

for index, row in df_eam.iterrows():
    if row['Local do serviço'] in ['Picador', 'Moinhos', 'Secadores', 'Linha Hydrodyn', 'Lixadeira Bison', 'Linha de Classificação']:
        df_eam.loc[index, 'Relatado por'] = 'HD'
    elif row['Local do serviço'] in ['Laqueadora Vits', 'Impregnadora Vits', 'Impregnadora Tocchio', 'Preparação Química', 'Wemhoner', 'Siempelkamp', 'Barberãn', 'Torwegge', 'Cyklop', 'Voma', 'Extrusora Tecno', 'Extrusora WPC 1', 'Extrusora WPC 2', 'Torwegge 3', 'Pisos e acessórios', 'Fábrica de acessórios']:
        df_eam.loc[index, 'Relatado por'] = 'LPR'
    else:
        df_eam.loc[index, 'Relatado por'] = 'Elimina'

def status_lpr (df_eam): #Status por áreas (LPR)
    df_LPR = df_eam[df_eam['Relatado por'].isin(['LPR'])]
    df_LPR = df_LPR[df_LPR['Status'] != 'Concluído']
    df_LPR = df_LPR.groupby('Status').size().reset_index(name='Count')
    df_LPR.columns = ['Status', 'Ordens de serviço']
    fig_LPR = px.bar(df_LPR, x='Ordens de serviço', y='Status', title='Status por área - LPR',text= 'Ordens de serviço')
    fig_LPR.update_layout(xaxis_title='', yaxis_title='',showlegend=False)
    fig_LPR.update_layout(
        showlegend=False,
        margin=dict(l=0, r=0, t=40, b=0),
        autosize=True,
        height=1080,
        width=1920,
    )
    return pio.to_html(fig_LPR, full_html=False)  

def status_hd (df_eam): #Status por áreas (HD)
    df_HD = df_eam[df_eam['Relatado por'].isin(['HD'])]
    df_HD = df_HD[df_HD['Status'] != 'Concluído']
    df_HD = df_HD.groupby('Status').size().reset_index(name='Count')
    df_HD.columns = ['Status', 'Ordens de serviço']
    fig_HD = px.bar(df_HD, x='Ordens de serviço', y='Status', title='Status por área - HD',text= 'Ordens de serviço')
    fig_HD.update_layout(xaxis_title='', yaxis_title='',showlegend=False)
    fig_HD.update_layout(
        showlegend=False,
        margin=dict(l=0, r=0, t=40, b=0),
        autosize=True,
        height=1080,
        width=1920,
    )
    return pio.to_html(fig_HD, full_html=False)  

def graf_tipos (df_eam):  #Graficos OS - Tipos
    Filtro_tipos = df_eam['Tipo'].value_counts()
    porcentagem_tipos = (Filtro_tipos / Filtro_tipos.sum()) * 100
    df_porcentagem = pd.DataFrame({'Tipo': porcentagem_tipos.index, 'Porcentagem': porcentagem_tipos.values})
    fig_os_tipos = px.pie(df_porcentagem, values='Porcentagem', names='Tipo', title='Ordens de serviço por tipos',
                labels={'Porcentagem': '%'}, hole=0.3)
    fig_os_tipos.update_traces(textposition='inside', textinfo='percent+label')
    fig_os_tipos.update_layout(showlegend=False)
    fig_os_tipos.update_layout(
        showlegend=False,
        margin=dict(l=0, r=0, t=40, b=0),
        autosize=True,
        height=1080,
        width=1920,
    )
    return pio.to_html(fig_os_tipos, full_html=False) 

def graf_depart (df_eam): #Graficos OS - OS TOTAL por departamento
    departamentos_excluir = df_eam['Departamento'].value_counts()[df_eam['Departamento'].value_counts() < 10].index
    df_filtrado = df_eam[~df_eam['Departamento'].isin(departamentos_excluir)]
    filtro_depart = df_filtrado['Departamento'].value_counts()
    df_OS_depart = pd.DataFrame({'Departamento': filtro_depart.index, 'Qtd OS': filtro_depart.values})
    fig_os_ostotal = px.pie(df_OS_depart, values='Qtd OS', names='Departamento', 
                labels={'Qtd OS': '%'}, hole=0.3)
    fig_os_ostotal.update_traces(textposition='inside', textinfo='percent+label')
    fig_os_ostotal.update_layout(showlegend=False)
    fig_os_ostotal.update_layout(title='Ordens de serviço por departamento')
    fig_os_ostotal.update_layout(
        showlegend=False,
        margin=dict(l=0, r=0, t=40, b=0),
        autosize=True,
        height=1080,
        width=1920,
    )
    return pio.to_html(fig_os_ostotal, full_html=False) 

def graf_status (df_eam): #Graficos OS - Status
    excluir = ['Concluído'] 
    df_filtrado = df_eam[~df_eam['Status'].isin(excluir)]
    filtro_status = df_filtrado['Status'].value_counts().reset_index()
    filtro_status.columns = ['Status', 'Quantidade']
    fig_os_status = px.bar(filtro_status, x='Status', y='Quantidade', text='Quantidade',
                labels={'Quantidade': 'Quantidade de OS'}, color='Status')
    fig_os_status.update_traces(texttemplate='%{text:.2s}', textposition='outside', textfont_size=14,
                    hovertemplate='%{x}<br>Quantidade: %{y}')
    fig_os_status.update_layout(title='Ordens de serviço por status', xaxis_title='', yaxis_title='',showlegend=False)
    fig_os_status.update_yaxes(title_text='', showticklabels=False)
    fig_os_status.update_xaxes(tickangle=45)
    fig_os_status.update_layout(
        showlegend=False,
        margin=dict(l=0, r=0, t=40, b=0),
        autosize=True,
        height=1080,
        width=1920,
    )
    return pio.to_html(fig_os_status, full_html=False) 

def graf_corr (df_eam): #Corretiva
    corr_ = ['COR-EMER', 'COR-PROG', 'COR-REPR']
    corr_df_filtrado = df_eam[df_eam['Classe'].isin(corr_)]
    corr_Filtro_status = corr_df_filtrado['Classe'].value_counts()
    corr_porcentagem_tipos = (corr_Filtro_status / corr_Filtro_status.sum()) * 100
    corr_ = corr_porcentagem_tipos.reset_index()
    corr_.columns = ['Classe','Porcentagem']
    fig_corr = px.pie(corr_, values='Porcentagem', names='Classe', 
                labels={'Porcentagem': '%'}, hole=0.3)
    fig_corr.update_traces(textposition='inside', textinfo='percent+label')
    fig_corr.update_layout(showlegend=False)
    fig_corr.update_layout(title='Ordens de serviço - Corretiva')
    fig_corr.update_layout(
        showlegend=False,
        margin=dict(l=0, r=0, t=40, b=0),
        autosize=True,
        height=1080,
        width=1920,
    )
    return pio.to_html(fig_corr, full_html=False) 

def graf_pm (df_eam): #Projetos e melhorias
    Pm_ = ['PM-REDC', 'PM-PROD', 'PM-ORGA', 'PM-MAQN']  
    Pm_df_filtrado = df_eam[df_eam['Classe'].isin(Pm_)]
    Pm_Filtro_status = Pm_df_filtrado['Classe'].value_counts()
    Pm_porcentagem_tipos = (Pm_Filtro_status / Pm_Filtro_status.sum()) * 100
    Pm_ = Pm_porcentagem_tipos.reset_index()
    Pm_.columns = ['Classe','Porcentagem']
    fig_Pm = px.pie(Pm_, values='Porcentagem', names='Classe', 
                labels={'Porcentagem': '%'}, hole=0.3)
    fig_Pm.update_traces(textposition='inside', textinfo='percent+label')
    fig_Pm.update_layout(showlegend=False)
    fig_Pm.update_layout(title='Ordens de serviço - Projetos e melhorias')
    fig_Pm.update_layout(
        showlegend=False,
        margin=dict(l=0, r=0, t=40, b=0),
        autosize=True,
        height=1080,
        width=1920,
    )
    return pio.to_html(fig_Pm, full_html=False) 

def graf_est (df_eam): #Estratégica
    Est_ = ['EST_TRN', 'EST_ANFA', 'EST-ORGA', 'EST-PLAN']
    Est_df_filtrado = df_eam[df_eam['Classe'].isin(Est_)]
    Est_Filtro_status = Est_df_filtrado['Classe'].value_counts()
    Est_porcentagem_tipos = (Est_Filtro_status / Est_Filtro_status.sum()) * 100 
    Est_ = Est_porcentagem_tipos.reset_index()
    Est_.columns = ['Classe','Porcentagem']
    fig_Est = px.pie(Est_, values='Porcentagem', names='Classe', 
                labels={'Porcentagem': '%'}, hole=0.3)
    fig_Est.update_traces(textposition='inside', textinfo='percent+label')
    fig_Est.update_layout(showlegend=False)
    fig_Est.update_layout(title='Ordens de serviço - Estratégica')
    fig_Est.update_layout(
        showlegend=False,
        margin=dict(l=0, r=0, t=40, b=0),
        autosize=True,
        height=1080,
        width=1920,
    )
    return pio.to_html(fig_Est, full_html=False) 

# def graf_pdt (df_eam): #Preditiva
#     Pdt_ = ['PDT_ANCI', 'PDT_ANVI', 'PDT_TERM']
#     Pdt_df_filtrado = df_eam[df_eam['Classe'].isin(Pdt_)]
#     Pdt_Filtro_status = Pdt_df_filtrado['Classe'].value_counts()
#     Pdt_porcentagem_tipos = (Pdt_Filtro_status / Pdt_Filtro_status.sum()) * 100 
#     Pdt_ = Pdt_porcentagem_tipos.reset_index()
#     Pdt_.columns = ['Classe','Porcentagem']
#     fig_Pdt = px.pie(Pdt_, values='Porcentagem', names='Classe', 
#                 labels={'Porcentagem': '%'}, hole=0.3)
#     fig_Pdt.update_traces(textposition='inside', textinfo='percent+label')
#     fig_Pdt.update_layout(showlegend=False)
#     fig_Pdt.update_layout(title='Ordens de serviço - Preditiva')
#     return pio.to_html(fig_Pdt, full_html=False) 

def graf_prv (df_eam): #Preventiva
    Prv_ = ['PRV-REFO', 'PRV-CALI', 'PRV-CONS', 'PRV-REVI', 'PRV-AJUS', 'PRV-TEST', 'PRV-SUBS', 'PRV-INSP']  
    Prv_df_filtrado = df_eam[df_eam['Classe'].isin(Prv_)]
    Prv_Filtro_status = Prv_df_filtrado['Classe'].value_counts()
    Prv_porcentagem_tipos = (Prv_Filtro_status / Prv_Filtro_status.sum()) * 100
    Prv_ = Prv_porcentagem_tipos.reset_index()
    Prv_.columns = ['Classe','Porcentagem']
    fig_Prv = px.pie(Prv_, values='Porcentagem', names='Classe', 
                labels={'Porcentagem': '%'}, hole=0.3)
    fig_Prv.update_traces(textposition='inside', textinfo='percent+label')
    fig_Prv.update_layout(showlegend=False)
    fig_Prv.update_layout(title='Ordens de serviço - Preventiva')
    fig_Prv.update_layout(
        showlegend=False,
        margin=dict(l=0, r=0, t=40, b=0),
        autosize=True,
        height=1080,
        width=1920,
    )
    return pio.to_html(fig_Prv, full_html=False) 


#Manipulando dados do PC Factory
df_pcfactory['T.Decorrido'] = df_pcfactory['T.Decorrido'].replace('24:00:00', '00:00:00')
df_pcfactory['T.Decorrido'] = pd.to_datetime(df_pcfactory['T.Decorrido'], format='%H:%M:%S').dt.time
df_pcfactory['T.Decorrido'] = df_pcfactory['T.Decorrido'].apply(lambda x: datetime.timedelta(hours=x.hour, minutes=x.minute, seconds=x.second).total_seconds())

def mtbf (df_pcfactory): # MTBF - Tempo médio entre falhas
    df_parada = df_pcfactory[(df_pcfactory['Status de Recurso'].isin(['[0202] Manutenção Elétrica', '[0203] Manutenção Mecânica','[1101] Parada não Identificada']))]
    total_parada = df_parada.groupby('Recurso')['T.Decorrido'].sum()
    total_parada = total_parada / 3600
    total_parada = round(total_parada,0)
    Falhas = df_parada['Recurso'].value_counts()
    df_produzindo = df_pcfactory[(df_pcfactory['Status de Recurso'].isin(['[0101] Producao', '[0902] Refeição', '[0838] Término de Produção', '[0837] Início de Produção']))]
    total_produzido = df_produzindo.groupby('Recurso')['T.Decorrido'].sum()
    total_produzido = total_produzido / 3600
    total_produzido = round(total_produzido,0)
    mtbf = (total_produzido - total_parada) / Falhas
    mtbf = round(mtbf,0)
    mtbf = mtbf.reset_index()
    mtbf.columns = ['Recurso', 'MTBF']
    fig_mtbf = px.bar(mtbf, x='Recurso', y='MTBF', title='Tempo médio entre falhas por setor')
    fig_mtbf.update_layout(xaxis_title='', yaxis_title='',showlegend=False)
    fig_mtbf.update_layout(
        showlegend=False,
        margin=dict(l=0, r=0, t=40, b=0),
        autosize=True,
        height=1080,
        width=1920,
    )
    return pio.to_html(fig_mtbf, full_html=False) 

def mttr (df_pcfactory): # MTTR - Tempo médio para reparo de um equipamento
    df_mttr = df_pcfactory[(df_pcfactory['Status de Recurso'].isin(['[0202] Manutenção Elétrica', '[0203] Manutenção Mecânica']))]
    total_reparos = df_mttr.groupby('Recurso')['T.Decorrido'].sum()
    reparos= df_mttr['Recurso'].value_counts()
    mttr = total_reparos / reparos
    mttr = mttr / 3600
    mttr = round(mttr,2)
    mttr = mttr[mttr >= 0]
    mttr = mttr.reset_index()
    mttr.columns = ['Recurso', 'MTTR']
    fig_mttr = px.bar(mttr, x='Recurso', y='MTTR', title='Tempo médio para reparo por setor')
    fig_mttr.update_layout(xaxis_title='', yaxis_title='',showlegend=False)
    fig_mttr.update_layout(
        showlegend=False,
        margin=dict(l=0, r=0, t=40, b=0),
        autosize=True,
        height=1080,
        width=1920,
    )
    return pio.to_html(fig_mttr, full_html=False) 

def disponibilidade (df_pcfactory): # (Mede a eficácia do tempo de inatividade em um sistema ou equipamento e também é o tempo que a máquina está disponível para funcionar, conforme o programado.)
    #MTBF 
    df_parada = df_pcfactory[(df_pcfactory['Status de Recurso'].isin(['[0202] Manutenção Elétrica', '[0203] Manutenção Mecânica','[1101] Parada não Identificada']))]
    total_parada = df_parada.groupby('Recurso')['T.Decorrido'].sum()
    total_parada = total_parada / 3600
    total_parada = round(total_parada,0)
    Falhas = df_parada['Recurso'].value_counts()
    df_produzindo = df_pcfactory[(df_pcfactory['Status de Recurso'].isin(['[0101] Producao', '[0902] Refeição', '[0838] Término de Produção', '[0837] Início de Produção']))]
    total_produzido = df_produzindo.groupby('Recurso')['T.Decorrido'].sum()
    total_produzido = total_produzido / 3600
    total_produzido = round(total_produzido,0)
    mtbf = (total_produzido - total_parada) / Falhas
    mtbf = round(mtbf,0)
    #MTTR
    df_mttr = df_pcfactory[(df_pcfactory['Status de Recurso'].isin(['[0202] Manutenção Elétrica', '[0203] Manutenção Mecânica']))]
    total_reparos = df_mttr.groupby('Recurso')['T.Decorrido'].sum()
    reparos= df_mttr['Recurso'].value_counts()
    mttr = total_reparos / reparos
    mttr = mttr / 3600
    mttr = round(mttr,2)
    D = (mtbf / (mtbf + mttr)) * 100
    D = round(D,1)
    D = D[D >= 0]
    Disponibilidade = D.reset_index()
    Disponibilidade.columns = ['Recurso', 'Disponibilidade']
    fig_disp = px.bar(Disponibilidade, x='Recurso', y='Disponibilidade', title='Disponibilidade dos setores')
    fig_disp.update_layout(xaxis_title='', yaxis_title='',showlegend=False)
    fig_disp.update_layout(
        showlegend=False,
        margin=dict(l=0, r=0, t=40, b=0),
        autosize=True,
        height=1080,
        width=1920,
    )
    return pio.to_html(fig_disp, full_html=False) 

def confiabilidade (df_pcfactory): # (Probabilidade de um setor desempenhar a sua função de acordo com as condições de operação e durante um intervalo específico)
    df_parada = df_pcfactory[(df_pcfactory['Status de Recurso'].isin(['[0202] Manutenção Elétrica', '[0203] Manutenção Mecânica','[1101] Parada não Identificada']))]
    total_parada = df_parada.groupby('Recurso')['T.Decorrido'].sum()
    total_parada = total_parada / 3600
    total_parada = round(total_parada,0)
    Falhas = df_parada['Recurso'].value_counts()
    df_produzindo = df_pcfactory[(df_pcfactory['Status de Recurso'].isin(['[0101] Producao', '[0902] Refeição', '[0838] Término de Produção', '[0837] Início de Produção']))]
    total_produzido = df_produzindo.groupby('Recurso')['T.Decorrido'].sum()
    total_produzido = total_produzido / 3600
    total_produzido = round(total_produzido,0)
    mtbf = (total_produzido - total_parada) / Falhas
    mtbf = round(mtbf,0)
    taxa_de_falha = 1 / mtbf
    tempo = 168 #a cada uma semana
    Confiabilidade = 2.7182 ** (-taxa_de_falha * tempo)* 100
    Confiabilidade = round(Confiabilidade, 2)
    Confiabilidade = 100 - Confiabilidade
    Confiabilidade = Confiabilidade[Confiabilidade >= 0]
    Confiabilidade = Confiabilidade.reset_index()
    Confiabilidade.columns = ['Recurso', 'Confiabilidade']
    fig_Confiabilidade = px.bar(Confiabilidade, x='Recurso', y='Confiabilidade', title='Confiabilidade da semana por setores')
    fig_Confiabilidade.update_layout(xaxis_title='', yaxis_title='',showlegend=False)
    fig_Confiabilidade.update_layout(
        showlegend=False,
        margin=dict(l=0, r=0, t=40, b=0),
        autosize=True,
        height=1080,
        width=1920,
    )
    return pio.to_html(fig_Confiabilidade, full_html=False)

# def oee (df_pcfactory): # Calculo de OEE Simplificado (Eficácia geral do equipamento)
#     #OEE simplificado pois não temos os dados que estão faltado
#     #(Performance[velocidade real/ velocidade padrão])
#     #(Qualidade[quantidade de produtos bons/ total produtos produzidos])
#     df_produzindo = df_pcfactory[(df_pcfactory['Status de Recurso'].isin(['[0101] Producao', '[0902] Refeição', '[0838] Término de Produção', '[0837] Início de Produção']))]
#     total_produzido = df_produzindo.groupby('Recurso')['T.Decorrido'].sum()
#     oee= (total_produzido / 720) * 100
#     oee = oee / 3600  
#     oee = round(oee,2)
#     oee = oee.reset_index()
#     oee.columns = ['Recurso', 'OEE']
#     fig_oee = px.bar(oee, x='Recurso', y='OEE', title='Oee simplificado por setor')
#     fig_oee.update_layout(xaxis_title='', yaxis_title='',showlegend=False)
#     return pio.to_html(fig_oee, full_html=False)

def indice_prod (df_pcfactory): # IP - Índice de produtividade da máquina
    h_disponiveis = 720 #padrão do mês - 30 dias
    df_nprogramada = df_pcfactory[(df_pcfactory['Status de Recurso'].isin(['[0202] Manutenção Elétrica', '[0203] Manutenção Mecânica']))]
    df_programada = df_pcfactory[(df_pcfactory['Status de Recurso'].isin(['[0201] Manutenção Programada']))]
    h_nprogramada = df_nprogramada.groupby('Recurso')['T.Decorrido'].sum()
    h_programada = df_programada.groupby('Recurso')['T.Decorrido'].sum()
    h_nprogramada = h_nprogramada /3600
    h_programada = h_programada / 3600
    h_trabalhas = h_nprogramada.add(h_programada, fill_value=0)
    h_trabalhas = 720 - h_trabalhas
    ip = h_trabalhas / h_disponiveis * 100
    ip = round(ip,2)
    ip = ip.reset_index()
    ip.columns = ['Recurso', 'IP']
    fig_ip = px.line(ip, x='Recurso', y='IP', title='Índice de produtividade das máquinas',text= 'IP')
    fig_ip.update_traces(textposition='top center')
    fig_ip.update_layout(xaxis_title='', yaxis_title='',showlegend=False)
    fig_ip.update_layout(
        showlegend=False,
        margin=dict(l=0, r=0, t=40, b=0),
        autosize=True,
        height=1080,
        width=1920,
    )
    return pio.to_html(fig_ip, full_html=False)


@app.route('/')
def index():
    return render_template('grafico_teste.html',gastos_setores_4j=gastos_setores_4j (df_custo), gastos_setores_5j=gastos_setores_5j (df_custo),status_lpr=status_lpr (df_eam),status_hd=status_hd (df_eam),graf_tipos=graf_tipos (df_eam),graf_depart=graf_depart (df_eam),graf_status=graf_status (df_eam),graf_corr=graf_corr (df_eam),graf_pm=graf_pm (df_eam),graf_est=graf_est (df_eam),graf_prv=graf_prv (df_eam),mtbf=mtbf (df_pcfactory),mttr=mttr (df_pcfactory),disponibilidade=disponibilidade (df_pcfactory),confiabilidade=confiabilidade (df_pcfactory),indice_prod=indice_prod (df_pcfactory))
#não esta aparecendo - oee (df_pcfactory) e graf_pdt (df_eam) - custo_geral=custo_geral (df_custo) por enquanto nao ta funfando

if __name__ == '__main__':
    app.run(debug=True)




#Comentários
    #Criar uma formula para a cada 3 meses gerar um grafico trimestral e a cada 12 meses gerar um anual para assim colocar uma linha de tendência ou até mesmo
    # uma previsão 
    #No momento tem alguns dados que posteriormente serão inseridos no EAM que possibilitaram fazer analises mais
    #precisas de custos de mão de obra e custos de tipos de ordens (preventiva, corretiva entre outros)
    #Quando o EAM estiver rodando tudo limpo, com todas as informações devidas será muito melhor para 
    #analisar e gerar certos gráficos