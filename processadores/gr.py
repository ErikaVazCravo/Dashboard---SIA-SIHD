"""Regras migradas de Gr Procedimentos - Quantitativos e Valores.py; leitura sem interface e sem persistencia."""
from .comum import abrir_texto
import os
import re
import sys
import pandas as pd

def extrair_competencia_linha(competencia_raw):
    competencia_raw = str(competencia_raw).strip()
    if len(competencia_raw) == 6 and competencia_raw.isdigit():
        ano = int(competencia_raw[:4])
        mes = int(competencia_raw[4:6])
        competencia = f'{mes:02d}/{ano}'
        return (ano, mes, competencia)
    return (None, None, '')
regex_data = re.compile('\\b\\d{2}/\\d{2}/\\d{4}\\b')

def limpar_cnes(valor):
    if pd.isna(valor):
        return ''
    return re.sub('\\D', '', str(valor)).strip()

def clean_excel_string(s):
    if isinstance(s, str):
        return re.sub('[\\x00-\\x1F]', '', s)
    return s

def parse_int_from_field(s):
    if s is None:
        return 0
    s = str(s).strip()
    if not s:
        return 0
    s = s.replace('.', '').replace(' ', '').replace(',', '.')
    try:
        return int(float(s))
    except:
        return 0

def parse_float_from_field(s):
    if s is None:
        return 0.0
    s = str(s).strip()
    if not s:
        return 0.0
    s = s.replace('.', '').replace(' ', '').replace(',', '.')
    try:
        return round(float(s), 2)
    except:
        return 0.0

def obter_nome_estabelecimento(cnes, nome_relatorio, mapa):
    cnes = limpar_cnes(cnes)
    if cnes in mapa:
        return mapa[cnes]
    return str(nome_relatorio).strip()

def grupo_procedimento(procedimento):
    procedimento = str(procedimento).strip()
    if len(procedimento) >= 2 and procedimento[:2].isdigit():
        return procedimento[:2]
    return ''

def processar_raprocqt(caminho, processamento, MAPA_NOMES):
    with abrir_texto(caminho) as f:
        linhas = [ln.rstrip('\n') for ln in f]
    registros = []
    cnes_atual = ''
    estabelecimento_atual = ''
    for linha_raw in linhas:
        if not linha_raw:
            continue
        linha_strip = linha_raw.strip()
        if linha_strip.startswith('Pagina') or linha_strip.startswith('SMS-ARARAQ') or regex_data.search(linha_strip) or linha_strip.startswith('********') or linha_strip.startswith('CMP') or linha_strip.startswith('_ SINTETICO') or linha_strip.upper().startswith('CMP   PROCEDIMENTO'):
            continue
        if linha_strip.startswith('Unidade'):
            cnes_atual = limpar_cnes(linha_raw[8:15].strip())
            estabelecimento_relatorio = linha_raw[16:].strip()
            estabelecimento_atual = obter_nome_estabelecimento(cnes_atual, estabelecimento_relatorio, MAPA_NOMES)
            continue
        if linha_strip.startswith('TOTAL DA UNIDADE'):
            continue
        if len(linha_raw) >= 6:
            competencia_raw = linha_raw[0:6].strip()
            ano_linha, mes_linha, competencia_linha = extrair_competencia_linha(competencia_raw)
            if not competencia_linha:
                continue
            procedimento = linha_raw[7:17].strip()
            apac = linha_raw[18:31].strip()
            qtd_apres_raw = linha_raw[39:50].strip()
            qtd_aprov_raw = linha_raw[51:62].strip()
            tipo_reg = 'APAC' if apac else 'BPA/RAAS'
            registros.append({'Processamento': processamento, 'ANO': ano_linha, 'Mes': mes_linha, 'Competencia': competencia_linha, 'CNES': cnes_atual, 'Prestador': estabelecimento_atual, 'Tipo Reg.': tipo_reg, 'Procedimento': procedimento, 'Grupo Proc.': grupo_procedimento(procedimento), 'Qtd. Apres.': parse_int_from_field(qtd_apres_raw), 'Qtd. Aprov.': parse_int_from_field(qtd_aprov_raw)})
    df = pd.DataFrame(registros)
    if df.empty:
        return df
    df = df.map(clean_excel_string)
    df = df.groupby(['Processamento', 'ANO', 'Mes', 'Competencia', 'CNES', 'Prestador', 'Tipo Reg.', 'Procedimento', 'Grupo Proc.'], as_index=False, dropna=False).agg({'Qtd. Apres.': 'sum', 'Qtd. Aprov.': 'sum'})
    return df

def processar_raproced(caminho, processamento, MAPA_NOMES):
    with abrir_texto(caminho) as f:
        linhas = [ln.rstrip('\n') for ln in f]
    registros = []
    cnes_atual = ''
    estabelecimento_atual = ''
    for linha_raw in linhas:
        if not linha_raw:
            continue
        linha_strip = linha_raw.strip()
        if linha_strip.startswith('Pagina') or linha_strip.startswith('SMS-ARARAQ') or regex_data.search(linha_strip) or linha_strip.startswith('********') or linha_strip.startswith('CMP') or linha_strip.startswith('_ SINTETICO') or linha_strip.upper().startswith('CMP   PROCEDIMENTO'):
            continue
        if linha_strip.startswith('Unidade'):
            cnes_atual = limpar_cnes(linha_raw[8:15].strip())
            estabelecimento_relatorio = linha_raw[16:].strip()
            estabelecimento_atual = obter_nome_estabelecimento(cnes_atual, estabelecimento_relatorio, MAPA_NOMES)
            continue
        if linha_strip.startswith('TOTAL DA UNIDADE'):
            continue
        if len(linha_raw) >= 6:
            competencia_raw = linha_raw[0:6].strip()
            ano_linha, mes_linha, competencia_linha = extrair_competencia_linha(competencia_raw)
            if not competencia_linha:
                continue
            procedimento = linha_raw[7:17].strip()
            apac = linha_raw[18:31].strip()
            vlr_apres_raw = linha_raw[39:53].strip()
            vlr_aprov_raw = linha_raw[54:68].strip()
            tipo_reg = 'APAC' if apac else 'BPA/RAAS'
            registros.append({'Processamento': processamento, 'ANO': ano_linha, 'Mes': mes_linha, 'Competencia': competencia_linha, 'CNES': cnes_atual, 'Prestador': estabelecimento_atual, 'Tipo Reg.': tipo_reg, 'Procedimento': procedimento, 'Grupo Proc.': grupo_procedimento(procedimento), 'Vlr. Apres.': parse_float_from_field(vlr_apres_raw), 'Vlr. Aprov.': parse_float_from_field(vlr_aprov_raw)})
    df = pd.DataFrame(registros)
    if df.empty:
        return df
    df = df.map(clean_excel_string)
    df = df.groupby(['Processamento', 'ANO', 'Mes', 'Competencia', 'CNES', 'Prestador', 'Tipo Reg.', 'Procedimento', 'Grupo Proc.'], as_index=False, dropna=False).agg({'Vlr. Apres.': 'sum', 'Vlr. Aprov.': 'sum'})
    return df
