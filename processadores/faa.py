"""Contagem FAA das UPAs: regras do script original, sem interface desktop."""
import re
import pandas as pd
from io import BytesIO
from hashlib import sha256
from processamento import excel, validar_competencia

COLUNAS = ['ANO', 'MÊS', 'Competencia', 'CNES', 'Prestador', 'Especialidade', 'Qtde.']
UPAS = {'2064731': {'Prestador': 'UPA Central', 'Palavras': ['central', 'centro', 'upa ii']}, '4047184': {'Prestador': 'UPA Vila Xavier', 'Palavras': ['vila xavier', 'xavier']}, '9099581': {'Prestador': 'UPA Selmi Dei', 'Palavras': ['selmi', 'selmi dei']}}

def normalizar_texto(valor):
    if pd.isna(valor):
        return ''
    texto = str(valor).strip().lower()
    substituicoes = {'á': 'a', 'à': 'a', 'â': 'a', 'ã': 'a', 'é': 'e', 'ê': 'e', 'í': 'i', 'ó': 'o', 'ô': 'o', 'õ': 'o', 'ú': 'u', 'ç': 'c'}
    for antigo, novo in substituicoes.items():
        texto = texto.replace(antigo, novo)
    texto = re.sub('\\s+', ' ', texto)
    return texto

def padronizar_especialidades(df):
    df2 = df.copy()
    df2['Especialidade'] = df2['Especialidade'].replace({'AUXILIAR DE ENFERMAGEM': 'ENFERMAGEM', 'ENFERMEIRA (O)': 'ENFERMAGEM', 'TECNICO DE ENFERMAGEM': 'ENFERMAGEM', 'TECNICO DE ENFERMAGEM PSF': 'ENFERMAGEM', 'FARMACEUTICO': 'ATEND. FARMÁRCIA', 'TECNICO EM FARMACIA': 'ATEND. FARMÁRCIA', 'MEDICINA DO TRABALHO': 'MÉDICO – CLÍNICA GERAL', 'CLINICO GERAL (225125)': 'MÉDICO – CLÍNICA GERAL', 'MÉDICO CARDIOLOGISTA': 'MÉDICO – CLÍNICA GERAL', 'MEDICO GENERALISTA': 'MÉDICO – CLÍNICA GERAL', 'CIRURGIAO DENTISTA C.GERA': 'ODONTOLOGIA', 'TECNICO EM HIGIENE DENTAL': 'ODONTOLOGIA', 'ASSISTENTE SOCIAL': 'ASSISTÊNCIA SOCIAL', 'ORTOPEDIA/TRAUMATOLOGIA': 'MÉDICO - ORTOPEDIA', 'PEDIATRIA': 'MÉDICO - PEDIATRIA'})
    return df2


def ler_relatorio(dados, cnes, competencia, nome):
    """Lê a primeira aba, como o script original; a UPA vem do campo de upload."""
    try:
        preview = pd.read_excel(BytesIO(dados), header=None, nrows=20, engine='openpyxl')
        cabecalho = next((idx for idx, linha in preview.iterrows()
            if {'Especialidade', 'Quantidade'}.issubset({str(v).strip() for v in linha.dropna()})), None)
        if cabecalho is None:
            raise ValueError('Cabeçalho com Especialidade e Quantidade não localizado nas primeiras 20 linhas.')
        df = pd.read_excel(BytesIO(dados), header=cabecalho, engine='openpyxl')
    except ValueError:
        raise
    except Exception as erro:
        raise ValueError('Não foi possível ler a planilha. Envie o XLSX original emitido pelo FAA.') from erro
    df.columns = [str(c).strip() for c in df.columns]
    obrigatorias = {'Unidade Prestadora', 'Especialidade', 'Quantidade'}
    faltantes = obrigatorias - set(df.columns)
    if faltantes:
        raise ValueError('Colunas ausentes: ' + ', '.join(sorted(faltantes)))
    df = df[df['Especialidade'].notna() & df['Quantidade'].notna()].copy()
    df['Especialidade'] = df['Especialidade'].astype(str).str.strip()
    df = df[df['Especialidade'] != ''].copy()
    if df.empty:
        raise ValueError('Nenhuma especialidade com quantidade preenchida foi encontrada.')
    for unidade in df['Unidade Prestadora'].dropna().unique():
        texto = normalizar_texto(unidade)
        identificadas = {codigo for codigo, upa in UPAS.items()
            if codigo in texto or any(palavra in texto for palavra in upa['Palavras'])}
        if identificadas and identificadas != {cnes}:
            raise ValueError(f'A coluna Unidade Prestadora contém uma UPA diferente de {UPAS[cnes]["Prestador"]}. Confira o campo de upload.')
    quantidades = pd.to_numeric(df['Quantidade'], errors='coerce')
    if quantidades.isna().any() or (quantidades < 0).any() or (quantidades % 1 != 0).any():
        raise ValueError('Quantidade deve conter números inteiros não negativos. Há valores inválidos na planilha.')
    df['Qtde.'] = quantidades.astype('int64')
    df['ANO'] = int(competencia[3:])
    df['MÊS'] = int(competencia[:2])
    df['Competencia'] = competencia
    df['CNES'] = cnes
    df['Prestador'] = UPAS[cnes]['Prestador']
    return df[COLUNAS]


def processar_faa(competencia, arquivos):
    """Recebe {CNES: [(nome, bytes)]}; retorna XLSX e resumo sem gravar dados."""
    competencia = validar_competencia(competencia)
    if set(arquivos) - set(UPAS):
        raise ValueError('Foi informada uma UPA não cadastrada.')
    frames, vistos = [], set()
    for cnes, itens in arquivos.items():
        for nome, dados in itens:
            if not nome.lower().endswith('.xlsx'):
                raise ValueError(f'{nome}: envie um arquivo .xlsx emitido pelo FAA.')
            digest = sha256(dados).digest()
            if digest in vistos:
                raise ValueError(f'{nome}: o mesmo arquivo foi enviado mais de uma vez.')
            vistos.add(digest)
            try:
                frames.append(ler_relatorio(dados, cnes, competencia, nome))
            except ValueError as erro:
                raise ValueError(f'{UPAS[cnes]["Prestador"]} — {nome}: {erro}') from erro
    if not frames:
        raise ValueError('Envie pelo menos um relatório XLSX do FAA.')
    df = pd.concat(frames, ignore_index=True)
    chaves = COLUNAS[:-1]
    original = df.groupby(chaves, as_index=False)['Qtde.'].sum()
    padronizado = padronizar_especialidades(df).groupby(chaves, as_index=False)['Qtde.'].sum()
    saida = excel({'ESPECIALIDADES_PADRONIZADAS': padronizado, 'ORIGINAL': original})
    resumo = original.groupby(['CNES', 'Prestador'], as_index=False)['Qtde.'].sum()
    return saida, resumo
