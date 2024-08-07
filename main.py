import os
import configparser
import pandas as pd
import logging
from datetime import datetime
import tkinter as tk
from tkinter import filedialog

log_filename = datetime.now().strftime('log_%Y-%m-%d_%H-%M-%S.log')
logging.basicConfig(filename=log_filename, filemode='w', level=logging.DEBUG, format='%(asctime)s - %(message)s')

def ler_arquivo_ini(caminho_arquivo):
    try:
        config = configparser.ConfigParser()
        config.read(caminho_arquivo)
        return config
    except Exception as e:
        logging.error(f"Erro ao ler o arquivo INI: {caminho_arquivo} - {e}")
        raise

def extrair_informacoes(caminho_arquivo, nome_subpasta):
    try:
        config = ler_arquivo_ini(caminho_arquivo)
        informacoes = []
        
        campos = ['Empresa', 'Terminal', 'Operador', 'Porta', 'CNPJ', 'IP']
        info_secao = {}
        for secao in config.sections():
            for campo in campos:
                if config.has_option(secao, campo):
                    info_secao[campo] = config.get(secao, campo)
            if 'Operador' not in info_secao or not info_secao['Operador']:
                info_secao['Operador'] = nome_subpasta
            
            if len(info_secao) > 0:
                informacoes.append(info_secao)
        
        logging.info(f"Extraídas informações do arquivo: {caminho_arquivo}")
        return informacoes
    except Exception as e:
        logging.error(f"Erro ao extrair informações do arquivo: {caminho_arquivo} - {e}")
        return []

def processar_pastas(diretorios_raiz):
    for pasta_raiz in diretorios_raiz:
        dados = []
        for pasta_atual, _, arquivos in os.walk(pasta_raiz):
            nome_subpasta = os.path.basename(pasta_atual)
            for arquivo in arquivos:
                if arquivo.lower().endswith('.ini'):
                    caminho_completo = os.path.join(pasta_atual, arquivo)
                    dados.extend(extrair_informacoes(caminho_completo, nome_subpasta))
        df = pd.DataFrame(dados).drop_duplicates()
        
        try:
            nome_arquivo_excel = f'dados_ini_{os.path.basename(pasta_raiz)}.xlsx'
            df.to_excel(nome_arquivo_excel, index=False)
            logging.info(f"Dados extraídos e salvos em '{nome_arquivo_excel}'")
        except Exception as e:
            logging.error(f"Erro ao salvar os dados no arquivo Excel: {nome_arquivo_excel} - {e}")

def selecionar_diretorio():
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal
    diretorio_raiz = filedialog.askdirectory(title="Selecione o diretório raiz onde se encontram os arquivos")
    return diretorio_raiz

# Solicitar ao usuário o diretório raiz
diretorio_raiz = selecionar_diretorio()
if diretorio_raiz:
    diretorios_raiz = [diretorio_raiz]
    processar_pastas(diretorios_raiz)
else:
    logging.error("Nenhum diretório foi selecionado.")
