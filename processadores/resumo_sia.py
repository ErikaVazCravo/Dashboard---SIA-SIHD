"""Regras migradas de Resumo da Producao Ambulatorial.py; leitura sem interface e sem persistencia."""
from .comum import abrir_texto
import os
import re
import sys
import pandas as pd
regex_data = re.compile('\\b\\d{2}/\\d{2}/\\d{4}\\b')

def limpar_cnes(valor):
    if pd.isna(valor):
        return ''
    return re.sub('\\D', '', str(valor)).strip()

def limpar_texto_excel(valor):
    if isinstance(valor, str):
        return re.sub('[\\x00-\\x1F]', '', valor)
    return valor

def parse_quantidade(valor):
    if valor is None:
        return 0
    valor = str(valor).strip()
    if not valor:
        return 0
    valor = valor.replace('.', '').replace(' ', '').replace(',', '.')
    try:
        return int(float(valor))
    except:
        return 0

def parse_valor(valor):
    if valor is None:
        return 0.0
    valor = str(valor).strip()
    if not valor:
        return 0.0
    valor = valor.replace('.', '').replace(' ', '').replace(',', '.')
    try:
        return round(float(valor), 2)
    except:
        return 0.0

def obter_nome_estabelecimento(cnes, nome_relatorio, mapa):
    cnes = limpar_cnes(cnes)
    if cnes in mapa:
        return mapa[cnes]
    return str(nome_relatorio).strip()

def processar_rsp_sintetico(caminho, tipo_relatorio, mapa_nomes, competencia_informada):
    mes_informado = int(competencia_informada[:2])
    ano_informado = int(competencia_informada[3:])
    with abrir_texto(caminho) as f:
        linhas = [linha.rstrip('\n') for linha in f]
    dados = []
    cnes = ''
    prestador = ''
    for linha in linhas:
        linha_strip = linha.strip()
        if not linha_strip:
            continue
        if linha_strip.startswith('Pagina') or linha_strip.startswith('SMS-ARARAQ') or regex_data.search(linha_strip) or linha_strip.startswith('********') or linha_strip.startswith('  PROCED.'):
            continue
        if linha_strip.startswith('Unidade'):
            cnes = limpar_cnes(linha[8:15].strip())
            nome_relatorio = linha[16:51].strip()
            prestador = obter_nome_estabelecimento(cnes, nome_relatorio, mapa_nomes)
            continue
        if linha_strip.startswith('TOTAL DA UNIDADE'):
            total_apresentado = linha[75:87].strip()
            total_aprovado = linha[88:102].strip()
            if tipo_relatorio == 'quantidade':
                dados.append({'ANO': ano_informado, 'Mes': mes_informado, 'Competencia': competencia_informada, 'CNES': cnes, 'Prestador': prestador, 'Qtd. Apres.': parse_quantidade(total_apresentado), 'Qtd. Aprov.': parse_quantidade(total_aprovado)})
            elif tipo_relatorio == 'valor':
                dados.append({'ANO': ano_informado, 'Mes': mes_informado, 'Competencia': competencia_informada, 'CNES': cnes, 'Prestador': prestador, 'Vlr. Apres.': parse_valor(total_apresentado), 'Vlr. Aprov.': parse_valor(total_aprovado)})
    df = pd.DataFrame(dados)
    if not df.empty:
        df = df.map(limpar_texto_excel)
    return df
