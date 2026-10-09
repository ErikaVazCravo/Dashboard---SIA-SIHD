import argparse
from decimal import Decimal
from pathlib import Path
import re
import unicodedata
import pandas as pd
import pdfplumber
PRESTADORES = {'2079763': 'CAIRBAR SCHUTEL', '6943284': 'MATERNIDADE GOTA DE LEITE', '2082527': 'SANTA CASA', '9267263': 'UNIDADE DE RETAGUARDA'}
ORIGEM = 'Municipio / UF de ORIGEM'
COD_MUNICIPIO = 'Cód. Município'
COLUNAS = ['ANO', 'MÊS', 'Comp', 'CNES', 'Prestador', COD_MUNICIPIO, ORIGEM, 'AIH', 'Cod. Procedimento', 'Procedimento', 'Valor']
VALOR = '-?(?:\\d{1,3}(?:\\.\\d{3})+|\\d+),\\d{2}'
LINHA = re.compile(f'^(?P<municipio>\\d{{6}})\\s+(?P<origem>.+?\\s*/\\s*[A-Z]{{2}})\\s+(?P<aih>\\d{{13}})\\s+(?P<codigo>\\d{{10}})\\s*-\\s*(?P<descricao>.+?)\\s+(?:ALTA|M[EÉ]DIA|BAIXA)\\s+(?:FAEC|MAC)\\s+(?P<cnes>\\d{{7}})\\s*-\\s*(?P<prestador>.+?)\\s+(?P<valor>{VALOR})$', re.IGNORECASE)

def normalizar(texto):
    return ''.join((c for c in unicodedata.normalize('NFKD', texto) if not unicodedata.combining(c))).upper()

def converter_valor(valor):
    return Decimal(valor.replace('.', '').replace(',', '.'))

def extrair_texto_pdf(caminho_pdf):
    linhas = []
    with pdfplumber.open(caminho_pdf) as pdf:
        for numero, pagina in enumerate(pdf.pages, 1):
            texto = pagina.extract_text()
            if not texto or not texto.strip():
                raise ValueError(f'Página {numero} sem texto extraível. Verifique se o PDF precisa de OCR.')
            linhas.extend((' '.join(l.split()) for l in texto.splitlines()))
    return linhas

def extrair_competencia(linha):
    encontrado = re.search('COMPETENCIA\\s*:\\s*(\\d{2}/\\d{4})', normalizar(linha))
    return encontrado.group(1) if encontrado else ''

def interpretar_linha_producao(linha):
    encontrado = LINHA.fullmatch(' '.join(linha.split()))
    if not encontrado:
        return None
    d = encontrado.groupdict()
    return {'CNES': d['cnes'], 'Prestador': PRESTADORES.get(d['cnes'], d['prestador']), COD_MUNICIPIO: d['municipio'], ORIGEM: d['origem'], 'AIH': d['aih'], 'Cod. Procedimento': d['codigo'], 'Procedimento': d['descricao'], 'Valor': converter_valor(d['valor'])}

def interpretar_linhas(linhas):
    texto = normalizar('\n'.join(linhas))
    if 'DEMONSTRATIVO DE PROCEDENCIA ANALITICO' not in texto:
        raise ValueError('Selecione o DEMONSTRATIVO DE PROCEDÊNCIA ANALÍTICO, em lugar dos relatórios individuais por estabelecimento.')
    competencias = {c for l in linhas if (c := extrair_competencia(l))}
    if len(competencias) != 1:
        raise ValueError('O PDF deve conter uma única competência identificável.')
    competencia = competencias.pop()
    if not 1 <= int(competencia[:2]) <= 12:
        raise ValueError('Competência inválida.')
    registros = []
    quantidade_bloco = 0
    valor_bloco = Decimal('0')
    qtd_conferida = False
    municipios_conferidos = 0
    totais_executantes = []
    for linha in linhas:
        simples = normalizar(linha)
        dados = interpretar_linha_producao(linha)
        if dados:
            if qtd_conferida:
                raise ValueError('Subtotal de município ausente antes do próximo registro.')
            registros.append({'ANO': competencia[-4:], 'MÊS': competencia[:2], 'Comp': competencia, **dados})
            quantidade_bloco += 1
            valor_bloco += dados['Valor']
        elif re.match('^\\d{6,7}\\s+', linha) or re.search('\\b\\d{13}\\b', linha):
            raise ValueError(f'Linha de produção não reconhecida; a exportação foi interrompida: {linha}')
        elif simples.startswith('QTD TOTAL DE AIH'):
            qtd = re.fullmatch('QTD TOTAL DE AIHS?\\s+(\\d+)', simples)
            if not qtd or int(qtd.group(1)) != quantidade_bloco or qtd_conferida:
                raise ValueError('Quantidade de AIHs diverge do subtotal do município de origem.')
            qtd_conferida = True
        elif simples.startswith('TOTAL DO MUNICIPIO DE ORIGEM'):
            total = re.fullmatch(f'TOTAL DO MUNICIPIO DE ORIGEM\\s+({VALOR})', simples)
            if not total or not qtd_conferida or converter_valor(total.group(1)) != valor_bloco:
                raise ValueError('Valor extraído diverge do subtotal do município de origem.')
            municipios_conferidos += 1
            quantidade_bloco, valor_bloco, qtd_conferida = (0, Decimal('0'), False)
        elif simples.startswith('TOTAL DO MUNICIPIO EXECUTANTE'):
            total = re.fullmatch(f'TOTAL DO MUNICIPIO EXECUTANTE\\s+(\\d+)\\s+((?:{VALOR}\\s+){{4}}{VALOR})', simples)
            if not total:
                raise ValueError('Total do município executante não reconhecido.')
            numeros = [converter_valor(v) for v in total.group(2).split()]
            if sum(numeros[:-1]) != numeros[-1]:
                raise ValueError('Total do município executante inconsistente.')
            totais_executantes.append((int(total.group(1)), numeros[-1]))
    if not registros or not municipios_conferidos or quantidade_bloco or qtd_conferida:
        raise ValueError('PDF sem registros válidos ou com subtotais incompletos.')
    if not totais_executantes or sum((q for q, v in totais_executantes)) != len(registros) or sum((v for q, v in totais_executantes)) != sum((r['Valor'] for r in registros)):
        raise ValueError('A extração não confere com os totais dos municípios executantes.')
    return pd.DataFrame(registros, columns=COLUNAS)

def processar_pdf(caminho_pdf):
    return interpretar_linhas(extrair_texto_pdf(caminho_pdf))
