import os
import configparser
import pandas as pd

def ler_arquivo_ini(caminho_arquivo):
    config = configparser.ConfigParser()
    config.read(caminho_arquivo)
    return config

def extrair_informacoes(caminho_arquivo, nome_subpasta):
    config = ler_arquivo_ini(caminho_arquivo)
    informacoes = []
    
    campos = ['Empresa', 'Terminal', 'Operador', 'Porta', 'CNPJ', 'IP']
    info_secao = {}
    for secao in config.sections():
        for campo in campos:
            if config.has_option(secao, campo):
                info_secao[campo] = config.get(secao, campo)
        # Verificar se o campo "Operador" existe e, se não, adicionar o nome da subpasta
        if 'Operador' not in info_secao or not info_secao['Operador']:
            info_secao['Operador'] = nome_subpasta
        
        if len(info_secao) > 0:
            informacoes.append(info_secao)
    
    return informacoes

def processar_pastas(diretorios_raiz):
    for pasta_raiz in diretorios_raiz:
        dados = []
        for pasta_atual, _, arquivos in os.walk(pasta_raiz):
            nome_subpasta = os.path.basename(pasta_atual)
            for arquivo in arquivos:
                if arquivo.lower().endswith('.ini'):
                    caminho_completo = os.path.join(pasta_atual, arquivo)
                    dados.extend(extrair_informacoes(caminho_completo, nome_subpasta))
        
        # Criar um DataFrame com os dados e remover linhas duplicadas
        df = pd.DataFrame(dados).drop_duplicates()
        
        # Salvar em um arquivo Excel com formatação aprimorada
        nome_arquivo_excel = f'dados_ini_{os.path.basename(pasta_raiz)}.xlsx'
        df.to_excel(nome_arquivo_excel, index=False)
        print(f"Dados extraídos e salvos em '{nome_arquivo_excel}'")

# Exemplo de uso com três diretórios raiz diferentes
diretorios_raiz = [
    r'C:\Users\flaviob\Desktop\Estrutura',
]
processar_pastas(diretorios_raiz)