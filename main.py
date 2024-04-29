import os
import configparser
import pandas as pd

def ler_arquivo_ini(caminho_arquivo):
    config = configparser.ConfigParser()
    config.read(caminho_arquivo)
    return config

def extrair_informacoes(caminho_arquivo):
    config = ler_arquivo_ini(caminho_arquivo)
    informacoes = []
    
    campos = ['Empresa', 'Terminal', 'Operador', 'Porta']
    info_secao = {'Arquivo': caminho_arquivo}
    for secao in config.sections():
        for campo in campos:
            if config.has_option(secao, campo):
                info_secao[campo] = config.get(secao, campo)
        if len(info_secao) > 1:  # Se encontramos pelo menos um campo
            informacoes.append(info_secao)
    
    return informacoes

def percorrer_pastas(diretorio_raiz):
    dados = []
    for pasta_atual, _, arquivos in os.walk(diretorio_raiz):
        for arquivo in arquivos:
            if arquivo.lower().endswith('.ini'):
                caminho_completo = os.path.join(pasta_atual, arquivo)
                dados.extend(extrair_informacoes(caminho_completo))
    
    # Criar um DataFrame com os dados
    df = pd.DataFrame(dados)
    
    # Salvar em um arquivo Excel com formatação aprimorada
    nome_arquivo_excel = 'dados_ini.xlsx'
    df.to_excel(nome_arquivo_excel, index=False)

    print(f"Dados extraídos e salvos em '{nome_arquivo_excel}'")

# Exemplo de uso
pasta_raiz = r'C:\Users\dti.flaviob\Desktop\Estrutura'
percorrer_pastas(pasta_raiz)
