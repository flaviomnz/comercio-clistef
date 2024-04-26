import os
import configparser
import pandas as pd

def ler_arquivo_ini(caminho_arquivo):
    config = configparser.ConfigParser()
    config.read(caminho_arquivo)
    return config

def percorrer_pastas(diretorio_raiz):
    dados = []
    for pasta_atual, _, arquivos in os.walk(diretorio_raiz):
        for arquivo in arquivos:
            if arquivo.lower().endswith('.ini'):
                caminho_completo = os.path.join(pasta_atual, arquivo)
                config = ler_arquivo_ini(caminho_completo)
                
                # Extrair todos os campos e valores do arquivo INI
                for secao in config.sections():
                    for chave, valor in config.items(secao):
                        dados.append({'Arquivo': caminho_completo, 'Seção': secao, 'Chave': chave, 'Valor': valor})
    
    # Criar um DataFrame com os dados
    df = pd.DataFrame(dados)
    
    # Salvar em um arquivo Excel com formatação aprimorada
    nome_arquivo_excel = 'dados_ini.xlsx'
    df.to_excel(nome_arquivo_excel, index=False)

    print(f"Dados extraídos e salvos em '{nome_arquivo_excel}'")

# Exemplo de uso
pasta_raiz = r'C:\Users\dti.flaviob\Desktop\Estrutura'
percorrer_pastas(pasta_raiz)
