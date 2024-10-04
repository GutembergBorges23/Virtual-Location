import getpass
import os.path
import pandas as pd
from datetime import datetime, timedelta
import datetime as dt
# from datetime import timedelta
import time
import warnings
import sqlite3

# Use a biblioteca openpyxl normalmente
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.styles.stylesheet")


def print_hi(name):
    print(f'Deus é bom o tempo todo, o tempo todo {name}')


if __name__ == '__main__':

    print_hi('Deus é bom!')

    # obtendo e identificando usuário atual
    user = getpass.getuser()
    # definindo o caminho para a pasta
    path = f'C:\\Users\\{user}\\Procter and Gamble\\Database_Digitization - Documents\\Reporte Locação Virtual\\'
    print(f'Usuário identificado, {user}')
    # validando o usuário ativo para o caminho da pasta
    if not os.path.exists(path):
        print('O path não existe')
        print('Tentando borges.g.2')
        user = 'borges.g.2'
        path = f'C:\\Users\\{user}\\Procter and Gamble\\Database_Digitization - Documents\\Reporte Locação Virtual\\'

    # Definindo acesso para a planilha de Master Data
    md_file = 'MD Value Stream.xlsx'
    mard_file = 'MARD.xlsx'
    mbew_file = 'MBEW.xlsx'
    # mseg_file = 'MSEG.xlsx'
    mchb_file = 'MCHB.xlsx'
    location_mapping_file = 'Virtual Locations Mapping.xlsx'

    # Tabela MD RN e MD Solimoes
    df_consulta_md_rn = pd.read_excel(path + md_file, sheet_name='MD RN', dtype=str)
    df_consulta_md_sol = pd.read_excel(path + md_file, sheet_name='MD Solimoes', dtype=str)
    df_consulta_md_camp = pd.read_excel(path + md_file, sheet_name='MD CAMPUS', dtype=str)
    df_consulta_mard = pd.read_excel(path + mard_file, sheet_name='default_1', dtype=str)
    df_consulta_mbew = pd.read_excel(path + mbew_file, sheet_name='default_1', dtype=str)
    # df_consulta_mseg = pd.read_excel(path + mseg_file, sheet_name='default_1', dtype=str)
    # Inicio da leitura do df_consulta_mchb
    inicio_tempo = time.time()
    df_consulta_mchb = pd.read_excel(path + mchb_file, sheet_name='default_1', dtype=str)
    # Inicio da leitura do df_consulta_mchb
    fim_tempo = time.time()
    # Cálculo de tempo decorrido
    tempo_decorrido = fim_tempo - inicio_tempo
    print(f"Tempo de leitura do arquivo Excel: {tempo_decorrido} segundos")
    df_consulta_location = pd.read_excel(path + location_mapping_file, sheet_name='Sheet1', dtype=str)

    df_consulta_mchb = df_consulta_mchb.astype({
        'valuated_unrestricted_use_stock': float,
        'stock_in_quality_inspection': float,
        'blocked_stock': float,
    })

    # Crie uma série booleana para identificar linhas com todas as colunas iguais a 0
    serie_todas_colunas_zero = (df_consulta_mchb['valuated_unrestricted_use_stock'] == 0) & \
                               (df_consulta_mchb['stock_in_quality_inspection'] == 0) & \
                               (df_consulta_mchb['blocked_stock'] == 0)

    # Filtre o DataFrame usando a série booleana
    df_consulta_mchb = df_consulta_mchb[~serie_todas_colunas_zero]

    # tratando data na mseg e mchb
    df_consulta_mchb['date_of_last_change'] = pd.to_datetime(
        df_consulta_mchb['date_of_last_change'],
        yearfirst=True,
        errors='coerce'
    )

    # Agrupar dados por código e locação e apontar para ultima data
    # df_consulta_mchb = df_consulta_mchb.sort_values('Data_Mov', ascending=False).groupby(['Código', 'Locação'])
    # df_consulta_mchb = df_consulta_mchb.groupby(['Código', 'Planta', 'Locação'])['Data_Mov'].max().reset_index()
    # mchb= (df_consulta_mchb.loc[df_consulta_mchb['date_of_last_change'].dt.year > 2000].reset_index(drop=True))

    # Escolher e renomear colunas mseg
    """dicicinario4_renomeando = {
        'id_of_product_material': 'Código',
        'plant_id': 'Planta',
        'movement_type_inventory_management': 'Tipo_Mov',
        'storage_location': 'Locação',
        'quantity_1': 'Qtd',
        'batch_id': 'Lote',
        'posting_date_in_the_document': 'Data_Mov',
        'year_of_material_document': 'Ano_Mov',
        'ConcaMat+SL': 'ID_Codxloc',
    }"""

    # Escolher e renomear colunas mchb
    dicionario5_renomeando = {
        'id_of_product_material': 'Código',
        'plant_id': 'Planta',
        'storage_location': 'Locação',
        'batch_id': 'Lote',
        'date_of_last_change': 'Data_Mov',
        'valuated_unrestricted_use_stock': 'Estoque_Livre',
        'stock_in_quality_inspection': 'Estoque_Qualidade',
        'blocked_stock': 'Estoque_Bloqueado'
    }

    # Renomeando coluna mseg e mchb
    # df_consulta_mseg.rename(columns=dicicinario4_renomeando, inplace=True)
    df_consulta_mchb.rename(columns=dicionario5_renomeando, inplace=True)

    # Selecionado Colunas da tabela mseg e mchb
    df_md_columns1 = ['Código', 'Planta', 'Locação', 'Qtd', 'Lote', 'Data_Mov', 'ID_Codxloc']
    df_md_columns5 = ['Código', 'Planta', 'Locação', 'Lote', 'Estoque_Livre', 'Estoque_Qualidade',
                      'Estoque_Bloqueado', 'Data_Mov']

    # Atribuindo colunas na propria tabelas mseg e mchb
    # df_consulta_mseg = df_consulta_mseg[df_md_columns1]
    df_consulta_mchb = df_consulta_mchb[df_md_columns5]

    # Escolher e renomear colunas MD Geral
    dicicinario_renomeando = {
        'Plant': 'Planta',
        'Material Code': 'Código',
        'Material Desc': 'Descrição',
        'Material Type': 'Tipo',
        'Base unit of measure': 'UN',
        'MRP Controller': 'C. MRP',
        'VS': 'VS',
        'VS Bloqueio': 'VS Bloq.',
        'Area': 'Área',
        'Area Bloqueio': 'Área Bloq.',
        'Grupos': 'Grupo',
        'ODM': 'ODM',
        'PE': 'PE',
        'Analista': 'Analista',
    }

    # Escolher e renomear colunas mard
    dicicinario2_renomeando = {
        'id_of_product_material': 'Código',
        'plant_id': 'Planta',
        'valuated_unrestricted_use_stock': 'Estoque_Livre',
        'storage_location': 'Locação',
        'stock_in_quality_inspection': 'Estoque_Qualidade',
        'blocked_stock': 'Estoque_Bloqueado',
        'ConcaMat+SL': 'ID_Codxloc',
    }

    # Escolher e renomear colunas mbew
    dicicinario3_renomeando = {
        'id_of_product_material': 'Código',
        'site_plant_code': 'Planta',
        'standard_price': 'Preço_Padrão',
        'standard_price_in_the_previous_period': 'Preço_Unit_Ant',
        'price_unit': 'Preço_Unit',
    }

    # Mesclando as planilhas MD RN e MD Solimoes
    df_md_geral = pd.concat([df_consulta_md_rn, df_consulta_md_sol]).reset_index(drop=True)

    # Renomeando colunas
    df_md_geral.rename(columns=dicicinario_renomeando, inplace=True)
    df_consulta_mard.rename(columns=dicicinario2_renomeando, inplace=True)
    df_consulta_mbew.rename(columns=dicicinario3_renomeando, inplace=True)

    # Selecionando colunas
    df_md_columns = ['Planta', 'Código', 'Descrição', 'Tipo', 'UN', 'C. MRP', 'VS', 'VS Bloq.', 'Área', 'Área Bloq.',
                     'Grupo', 'ODM', 'PE', 'Analista']

    df_md_columns8 = ['Código', 'Planta', 'Locação', 'Estoque_Livre', 'Estoque_Qualidade', 'Estoque_Bloqueado']

    # Selecionando colunas na tabela df_consulta_MB52
    df_consulta_mard = df_consulta_mard[df_md_columns8]

    # Atribuindo colunas selecionada no proprio arquivo
    df_md_geral = df_md_geral[df_md_columns]

    # Tipagem de colunas
    # df_md_geral = df_md_geral.astype({
    #    'Código': int
    # })

    # Retirar códigos repetidos
    df_md_geral.drop_duplicates(subset='Código', inplace=True)
    df_consulta_md_camp.drop_duplicates(subset='Id', inplace=True)
    df_consulta_mbew.drop_duplicates(subset='Código', inplace=True)

    # Mesclar mard e mchb
    df_consulta_mb52 = df_consulta_mard.merge(
        df_consulta_mchb,
        left_on=['Código', 'Locação', 'Planta'],
        right_on=['Código', 'Locação', 'Planta'],
        how='inner'
    )

    df_colunas_selecionadas_mb52 = ['Código', 'Planta', 'Locação', 'Lote', 'Estoque_Livre_y', 'Estoque_Qualidade_y',
                                    'Estoque_Bloqueado_y', 'Data_Mov']

    df_consulta_mb52 = df_consulta_mb52[df_colunas_selecionadas_mb52]

    # data hoje - 3 dias
    data_atual_menos3 = pd.to_datetime('today') - timedelta(days=3)

    # Preenchendo campos vazios na coluna Data_Mov com a data calculada
    df_consulta_mb52.update(df_consulta_mb52['Data_Mov'].fillna(data_atual_menos3))

    # Coluna Data_Mov como datetime
    df_consulta_mb52['Data_Mov'] = pd.to_datetime(df_consulta_mb52['Data_Mov'])

    # Apontando para data de hoje
    dataAtual = datetime.now()
    # dataAtual = dt.datetime.strptime(dataAtual, '%d/%m/%Y').date()

    # Criando uma coluna 'QtDias' no df_consulta_mb52
    df_consulta_mb52['QtDias'] = (dataAtual - df_consulta_mb52['Data_Mov']).dt.days

    # Coluna Data_Mov como date
    df_consulta_mb52['Data_Mov'] = df_consulta_mb52['Data_Mov'].dt.date

    # df_virtual_hist
    df_virtual_hist = df_consulta_mb52.merge(
        df_consulta_location,
        left_on=['Locação'],
        right_on=['Locação'],
        how='inner'
    )

    # Exportar para um arquivo em excel
    # df_virtual_hist.to_excel('df_virtual_hist.xlsx')

    # Selecionando determinada colunas df_virtaul_hist
    df_hist_columns = [
        'Código',
        'Planta',
        'Locação',
        'Lote',
        'Estoque_Livre_y',
        'Estoque_Qualidade_y',
        'Estoque_Bloqueado_y',
        'Data_Mov',
        'QtDias',
        'Target/Dia',
        'Date_Update',
    ]

    # Criando coluna data_update
    df_virtual_hist['Date_Update'] = dt.date.today()

    df_virtual_hist = df_virtual_hist.astype({'Target/Dia': int})
    df_virtual_hist = df_virtual_hist[df_virtual_hist['QtDias'] > df_virtual_hist['Target/Dia']]
    df_virtual_hist = df_virtual_hist.reset_index(drop=True)    # Resetando as linhas filtradas
    df_virtual_hist = df_virtual_hist[df_hist_columns]

    # Verificar se os aruivos foram atualizados na pasta
    arquivo_mchb = os.path.join(path, 'MCHB.xlsx')
    arquivo_mbew = os.path.join(path, 'MBEW.xlsx')
    arquivo_mard = os.path.join(path, 'MARD.xlsx')
    
    # obtendo a data e hora do arquivo na pas
    data_modificacao_mchb = os.path.getmtime(arquivo_mchb)
    data_modificacao_mbew = os.path.getmtime(arquivo_mbew)
    data_modificacao_mard = os.path.getmtime(arquivo_mard)

    # Convertendo de timestamp apara objeto datime
    data_modificacao_mchb = dt.datetime.fromtimestamp(data_modificacao_mchb)
    data_modificacao_mchb = data_modificacao_mchb.date()
    df_virtual_hist['Date_Update'] = data_modificacao_mchb

    # Configurando o caminho para o banco de dados
    db_path = path + 'bd_location\\db_virtual.db'

    # Criando conexão com o banco de dados
    conexao = sqlite3.connect(db_path)
    ponteiro = conexao.cursor()

    # Consultando tabela df_historico_virtual no db
    ponteiro.execute('SELECT * FROM df_historico_virtual')
    # df_historico_virtual = df_virtual_hist.to_sql('df_historico_virtual', conexao, if_exists='append', index=False)
    df_historico_virtual = pd.DataFrame(ponteiro.fetchall(), columns=df_hist_columns)

    # Identificando linhas duplicadas
    print("Identificando possíveis linhas duplicadas")
    linhas_duplicadas = df_historico_virtual.duplicated()

    # Criando dataframe com linhas duplciadas
    print("Criando uma dataframe para identificar possíveis linhas duplicadas...")
    df_duplicadas = df_historico_virtual[linhas_duplicadas]

    # Removendo linhas duplicadas
    print("Removendo possíveis linhas duplicadas...")
    df_historico_virtual.drop_duplicates(inplace=True)
    conexao.commit()

    # Verificando data para input no db
    data_hoje = datetime.today().date()

    print("Verificando a data da última gravação...")
    # Verificar se a data do arquivo é mais recente que a data mínima no histórico
    if pd.to_datetime(df_historico_virtual['Date_Update'].max()).date() < data_modificacao_mchb:
        print('Inserindo novos dados...')

        # Inserindo novos dados no banco de dados
        df_virtual_hist.to_sql('df_historico_virtual', conexao, if_exists='append', index=False)
        ponteiro.execute('SELECT * FROM df_historico_virtual')
        df_historico_virtual = pd.DataFrame(ponteiro.fetchall(), columns=df_hist_columns)
        conexao.commit()
    else:
        print('Nenhuma informação foi gravada, Os Dados foram gravados anteriormente...')

    # Fechar a Conexão
    print("Fechando a conexão...")
    conexao.close()
    print('Conexão com o banco de dados fechada...')

    """df_historico_virtual = remover_duplicatas(df_historico_virtual, db_path, table_name='df_historico_virtual',
                                              param_duplicates=['Código', 'Planta', 'Locação', 'Lote',
                                                                'Estoque_Livre_y', 'Estoque_Qualidade_y',
                                                                'Estoque_Bloqueado_y', 'Data_Mov',
                                                                'Target/Dia', 'Date_Update', 'Data_x'])"""

    df_historico_virtual.to_excel('virtual_analisar.xlsx', index=False)

    # deletando tabelas que nao serao utilizadas no projeto
    print('Deletando dataframes que não serão utilizados...')
    del (
        df_consulta_md_rn,
        df_consulta_md_sol,
        df_consulta_mchb,
        df_virtual_hist,
        df_consulta_mard,
        df_duplicadas,
        df_md_geral,
    )

    print('====== Script executado com sucesso ======')
