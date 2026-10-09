import re
import os
import pdfplumber
import pandas as pd
prestadores = {'2079763': 'CAIRBAR SCHUTEL', '6943284': 'MATERNIDADE GOTA DE LEITE', '2082527': 'SANTA CASA', '9267263': 'UNIDADE DE RETAGUARDA'}

def extrair_texto_pdf(pdf_path):
    texto = []
    with pdfplumber.open(pdf_path) as pdf:
        for pagina in pdf.pages:
            conteudo = pagina.extract_text()
            if conteudo:
                texto.extend(conteudo.split('\n'))
    return texto

def obter_competencia(linhas):
    for linha in linhas:
        m = re.search('Competência:\\s*(\\d{2}/\\d{4})', linha)
        if m:
            return m.group(1)
    return ''

def converter_valor(valor):
    return float(str(valor).replace('.', '').replace(',', '.'))

def processar_pdf(pdf_path, tipo_atendimento):
    linhas = extrair_texto_pdf(pdf_path)
    competencia = obter_competencia(linhas)
    registros = []
    cnes = None
    estabelecimento = None
    municipio_atual = None
    cod_municipio = None
    dentro_bloco = False
    for linha in linhas:
        linha = linha.strip()
        if linha.startswith('MS/DATASUS') or linha.startswith('DEMONSTRATIVO') or linha.startswith('Página:') or linha.startswith('SISTEMA DE INFORMAÇÕES') or linha.startswith('M350320001') or linha.startswith('Gestor:') or linha.startswith('MUNICIPIO') or linha.startswith('FILTRO'):
            continue
        m = re.match('CNES\\s+(\\d{7})', linha)
        if m:
            cnes = m.group(1)
            estabelecimento = prestadores.get(cnes, f'CNES {cnes}')
            dentro_bloco = False
            continue
        if linha.startswith('MUN. ORIGEM'):
            dentro_bloco = True
            continue
        if linha.startswith('Total do CNES'):
            dentro_bloco = False
            municipio_atual = None
            cod_municipio = None
            continue
        if not dentro_bloco:
            continue
        if linha.startswith('Total do Municipio') or linha.startswith('Total do Gestor') or linha.startswith('TODAS AS ESPECIALIDADES'):
            continue
        padrao_municipio = '^(\\d{6})\\s+(.+?)\\s+(Psiquiatria|Clínico|Cirúrgico|Pediátricos|Obstétricos)\\s+(\\d+)\\s+([\\d\\.,]+)\\s+([\\d\\.,]+)$'
        m = re.match(padrao_municipio, linha)
        if m:
            cod_municipio = m.group(1)
            municipio_atual = m.group(2).strip()
            especialidade = m.group(3)
            qtd = int(m.group(4))
            total = converter_valor(m.group(5))
            registros.append({'Competencia': competencia, 'CNES': cnes, 'Prestador': estabelecimento, 'Tipo de Atendimento': tipo_atendimento, 'Cod. Mun.': cod_municipio, 'Municipio': municipio_atual, 'Especialidade': especialidade, 'Qtde. de AIHs': qtd, 'Valor': total})
            continue
        padrao_especialidade = '^(Psiquiatria|Clínico|Cirúrgico|Pediátricos|Obstétricos)\\s+(\\d+)\\s+([\\d\\.,]+)\\s+([\\d\\.,]+)$'
        m = re.match(padrao_especialidade, linha)
        if m and municipio_atual:
            especialidade = m.group(1)
            qtd = int(m.group(2))
            total = converter_valor(m.group(3))
            registros.append({'Competencia': competencia, 'CNES': cnes, 'Prestador': estabelecimento, 'Tipo de Atendimento': tipo_atendimento, 'Cod. Mun.': cod_municipio, 'Municipio': municipio_atual, 'Especialidade': especialidade, 'Qtde. de AIHs': qtd, 'Valor': total})
    return pd.DataFrame(registros)

def agrupar_total_municipios(df):
    return df.groupby(['Competencia', 'CNES', 'Prestador', 'Tipo de Atendimento', 'Cod. Mun.', 'Municipio', 'Especialidade'], as_index=False).agg({'Qtde. de AIHs': 'sum', 'Valor': 'sum'})

def agrupar_total_prestador(df):
    return df.groupby(['Competencia', 'CNES', 'Prestador', 'Tipo de Atendimento', 'Especialidade'], as_index=False).agg({'Qtde. de AIHs': 'sum', 'Valor': 'sum'})

def calcular_outros_tipos(df_sem_filtro, df_eletivo, df_urgencia, colunas_chave):
    total_sem_filtro = df_sem_filtro.groupby(colunas_chave, as_index=False).agg({'Qtde. de AIHs': 'sum', 'Valor': 'sum'})
    total_eletivo_urgencia = pd.concat([df_eletivo, df_urgencia], ignore_index=True)
    total_eletivo_urgencia = total_eletivo_urgencia.groupby(colunas_chave, as_index=False).agg({'Qtde. de AIHs': 'sum', 'Valor': 'sum'})
    resultado = total_sem_filtro.merge(total_eletivo_urgencia, on=colunas_chave, how='left', suffixes=('_sem_filtro', '_eletivo_urgencia'))
    resultado['Qtde. de AIHs_eletivo_urgencia'] = resultado['Qtde. de AIHs_eletivo_urgencia'].fillna(0)
    resultado['Valor_eletivo_urgencia'] = resultado['Valor_eletivo_urgencia'].fillna(0)
    resultado['Qtde. de AIHs'] = resultado['Qtde. de AIHs_sem_filtro'] - resultado['Qtde. de AIHs_eletivo_urgencia']
    resultado['Valor'] = resultado['Valor_sem_filtro'] - resultado['Valor_eletivo_urgencia']
    resultado = resultado[(resultado['Qtde. de AIHs'] != 0) | (resultado['Valor'].round(2) != 0)].copy()
    resultado['Tipo de Atendimento'] = 'Outros Tipos'
    resultado['Valor'] = resultado['Valor'].round(2)
    return resultado

def adicionar_ano_mes(df):
    """
    Adiciona as colunas ANO e MÊS a partir da Competencia.
    Competencia deve estar no formato MM/AAAA.
    """
    if df.empty:
        return df
    df = df.copy()
    df.insert(0, 'ANO', df['Competencia'].str[-4:].astype(int))
    df.insert(1, 'MÊS', df['Competencia'].str[:2].astype(int))
    return df
