import pandas as pd

def process_zges_file(file_path='zges.txt', output_path='zges.xlsx'):
    """
    Processa o arquivo zges.txt, extrai e transforma os dados,
    e salva o resultado em um arquivo Excel.

    Args:
        file_path (str): O caminho para o arquivo de entrada 'zges.txt'.
        output_path (str): O caminho para o arquivo de saída 'zges.xlsx'.
    """
    try:
        df = pd.read_csv(file_path, sep='\t', encoding='latin1', engine='python', on_bad_lines='skip', header=1)
    except FileNotFoundError:
        print(f"Erro: O arquivo '{file_path}' não foi encontrado. Certifique-se de que ele está no mesmo diretório do script.")
        return
    except Exception as e:
        print(f"Ocorreu um erro ao ler o arquivo '{file_path}': {e}")
        return

    df2 = df[['Código de Projeto', 'Descrição do Projeto', 'Gestor']]

    df2.rename(columns={
        'Código de Projeto': 'PEP1',
        'Descrição do Projeto': 'DESCR_PEP1',
        'Gestor': 'GESTOR'
    }, inplace=True)

    # Verifica se a coluna 'PEP1' existe antes de tentar dividi-la
    if 'PEP1' in df2.columns:
        df2[['col1', 'col2', 'Ano_pep']] = df2['PEP1'].astype(str).str.split('-', n=2, expand=True)
        df2['col22'] = df2['col2'].astype(str).str.slice(0, 3)
        df2['NUMERO'] = df2['col2'].astype(str).str.slice(-3)
        df2['SIGLA'] = df2['col1'].astype(str) + '-' + df2['col22'].astype(str)
    else:
        print("A coluna 'PEP1' não foi encontrada após a renomeação. Verifique o nome das colunas originais.")
        return

    df3 = df2[['PEP1', 'Ano_pep', 'SIGLA', 'NUMERO', 'DESCR_PEP1', 'GESTOR']]
    df3.dropna(inplace=True) # remove linhas com valores NaN em qualquer coluna
    df3 = df3.dropna(subset=['PEP1']) # garante que 'PEP1' não tenha NaNs, embora o dropna anterior já o faça

    # Expressão regular para validar o formato de 'PEP1'
    pattern = r'^[A-Z]{3}-[A-Z0-9]{6}-[0-9]{2}$'
    df3 = df3[df3['PEP1'].astype(str).str.match(pattern)]

    try:
        df3.to_excel(output_path, index=False)
        print(f"Arquivo '{output_path}' criado com sucesso!")
    except Exception as e:
        print(f"Ocorreu um erro ao salvar o arquivo '{output_path}': {e}")

# Executa a função
if __name__ == "__main__":
    process_zges_file()