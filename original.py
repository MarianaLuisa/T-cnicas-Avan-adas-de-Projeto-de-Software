import pandas as pd

caminho_arquivo = r"C:\Users\mariana.silva\Downloads\sispass_2018_2023.xlsx"
df = pd.read_excel(caminho_arquivo)

colunas_anos = [coluna for coluna in df.columns if coluna.startswith('s')]
dados_linhas = []

def calculate_bird_rate(row, prefix, year):
    """
    Calcula a taxa de aves com base nos dados da linha.

    Args:
        row (Series): Linha do DataFrame.
        prefix (str): Prefixo das colunas.
        year (int): Ano dos dados.

    Returns:
        float: Taxa de aves.
    """
    col_criadores = f"{prefix}.qtd_criadores_{year}0801"
    col_aves = f"{prefix}.qtd_aves_{year}0801"
    criadores = row[col_criadores]
    aves = row[col_aves]
    return round(aves / criadores, 2) if criadores != 0 else 0

def process_data(file_path):
    """
    Processa os dados do arquivo Excel e retorna um DataFrame final organizado.

    Args:
        file_path (str): Caminho do arquivo Excel.

    Returns:
        DataFrame: DataFrame final organizado.
    """
    df = pd.read_excel(file_path)
    colunas_anos = [coluna for coluna in df.columns if coluna.startswith('s')]
    dados_linhas = []

    for _, row in df.iterrows():
        for coluna in colunas_anos:
            ano = int(coluna.split('_')[-1][:4])
            prefixo = coluna.split('.')[0]

            nova_linha = {
                'm.cod_municipio': row['m.cod_municipio'],
                'm.nom_municipio': row['m.nom_municipio'],
                'm.sig_uf': row['m.sig_uf'],
                'm.nom_regiao': row['m.nom_regiao'],
                'ano': ano,
                'criadores': row[f"{prefixo}.qtd_criadores_{ano}0801"],
                'aves': row[f"{prefixo}.qtd_aves_{ano}0801"],
                'tx_aves': calculate_bird_rate(row, prefixo, ano)
            }
            dados_linhas.append(nova_linha)

    df_final = pd.DataFrame(dados_linhas)
    return df_final.drop_duplicates()

colunas_ordenadas = ['m.cod_municipio', 'm.nom_municipio', 'm.sig_uf', 'm.nom_regiao', 'ano', 'criadores', 'aves',
                     'tx_aves']
df_final = df_final[colunas_ordenadas]

# tava duplicando a planilha inteira, então isso daqui remove duplicatas
df_final = df_final.drop_duplicates()

df_final.to_excel('sispassAtualizado.xlsx', index=False)
