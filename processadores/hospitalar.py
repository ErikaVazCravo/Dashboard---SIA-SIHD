import argparse
from copy import copy
from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path
import re
import sys
import unicodedata
CABECALHOS = ['ANO', 'MÊS', 'COMP', 'CNES', 'Prestador', 'Qtde. Apres.', 'Qtde. Aprov', 'Compl. Federal', 'TOTAL']
PRESTADORES = {'2079763': 'CAIRBAR SCHUTEL', '6943284': 'MATERNIDADE GOTA DE LEITE', '2082527': 'SANTA CASA', '9267263': 'UNIDADE DE RETAGUARDA'}
INTEIRO = '(?:\\d{1,3}(?:\\.\\d{3})+|\\d+)'
DINHEIRO = '(?:\\d{1,3}(?:\\.\\d{3})+|\\d+),\\d{2}'
INICIO = re.compile(f'^(\\d{{7}})\\s*-\\s*(.+?)\\s+({INTEIRO})\\s+({INTEIRO})\\s+({DINHEIRO})$')

def normalizar(texto):
    return ''.join((c for c in unicodedata.normalize('NFKD', texto) if not unicodedata.combining(c))).lower()

def moeda(texto):
    return Decimal(texto.replace('.', '').replace(',', '.'))

def copiar_estilo(origem, destino):
    for atributo in ('font', 'fill', 'border', 'alignment', 'protection'):
        setattr(destino, atributo, copy(getattr(origem, atributo)))
    destino.number_format = origem.number_format

def valores(linha, quantidade):
    numeros = re.findall(DINHEIRO, linha)
    if len(numeros) != quantidade:
        raise ValueError(f'Linha com valores incompletos ou formato inesperado: {linha}')
    numeros = [moeda(n) for n in numeros]
    if sum(numeros[:-1]) != numeros[-1]:
        raise ValueError(f'Soma dos valores diverge do total: {linha}')
    return numeros[-1]

@dataclass(frozen=True)
class Registro:
    ano: int
    mes: int
    cnes: str
    nome: str
    apresentadas: int
    aprovadas: int
    federal: Decimal
    total: Decimal
    origem: str

    def linha(self):
        return [self.ano, self.mes, f'{self.mes:02d}/{self.ano}', self.cnes, PRESTADORES.get(self.cnes, self.nome), self.apresentadas, self.aprovadas, self.federal, self.total]

def interpretar(texto, origem='PDF'):
    """Falha em relatórios incompletos; não transforma campos ausentes em zero."""
    simples = normalizar(texto)
    if 'resumo dos valores aprovados por cnes' not in simples:
        raise ValueError('O PDF não é um resumo dos valores aprovados por CNES do SIHD2.')
    competencias = set(re.findall('competencia\\s*:\\s*(\\d{2})/(\\d{4})', simples))
    if len(competencias) != 1:
        raise ValueError('É necessário haver uma única competência identificável por PDF.')
    mes, ano = map(int, next(iter(competencias)))
    if not 1 <= mes <= 12:
        raise ValueError('Mês da competência inválido.')
    registros, atual, campos, vistos = ([], None, {}, set())

    def concluir():
        if atual is None:
            return
        cnes, nome, apresentadas, aprovadas, _ = atual.groups()
        if set(campos) != {'base', 'federal', 'local', 'total'}:
            raise ValueError(f"CNES {cnes}: bloco incompleto ({', '.join(campos)}).")
        if campos['base'] + campos['federal'] + campos['local'] != campos['total']:
            raise ValueError(f'CNES {cnes}: total diverge dos valores e complementos.')
        if cnes in vistos:
            raise ValueError(f'CNES {cnes} repetido no PDF. Verifique o relatório.')
        ap, av = (int(apresentadas.replace('.', '')), int(aprovadas.replace('.', '')))
        if av > ap:
            raise ValueError(f'CNES {cnes}: quantidade aprovada maior que apresentada.')
        vistos.add(cnes)
        registros.append(Registro(ano, mes, cnes, nome, ap, av, campos['federal'], campos['total'], origem))
    for linha in texto.splitlines():
        linha = ' '.join(linha.split())
        chave = normalizar(linha)
        inicio = INICIO.fullmatch(linha)
        if inicio:
            concluir()
            atual, campos = (inicio, {})
            continue
        if re.match('^\\d{7}\\s*-', linha):
            raise ValueError(f'Cabeçalho de prestador não reconhecido: {linha}')
        if chave.startswith(('total do municipio', 'total do gestor')):
            concluir()
            atual, campos = (None, {})
        if atual is None:
            continue
        for rotulo, campo, quantidade in [('valor sem complemento', 'base', 5), ('complemento federal', 'federal', 5), ('complemento local', 'local', 5), ('total ', 'total', 3)]:
            if chave.startswith(rotulo):
                if campo in campos:
                    raise ValueError(f'Linha repetida para o CNES {atual.group(1)}: {rotulo}')
                campos[campo] = valores(linha, quantidade)
                break
    concluir()
    if not registros:
        raise ValueError('Nenhum prestador encontrado. PDFs digitalizados precisam de OCR.')
    if 'total do gestor' not in simples:
        raise ValueError('Total do gestor ausente: não é possível conferir a extração completa.')
    rodape = simples.rsplit('total do gestor', 1)[1]
    totais = re.findall(f'(?:^|\\n)\\s*(?:total\\s+)?({INTEIRO})\\s+({INTEIRO})\\s+{DINHEIRO}\\s+({DINHEIRO})\\s+({DINHEIRO})\\s+({DINHEIRO})\\s*(?:\\n|$)', rodape)
    if len(totais) != 1:
        raise ValueError('Não foi possível identificar o total geral do gestor.')
    ap, av, hospitalar, profissional, total = totais[0]
    if moeda(hospitalar) + moeda(profissional) != moeda(total):
        raise ValueError('Total do gestor inconsistente.')
    if sum((r.apresentadas for r in registros)) != int(ap.replace('.', '')) or sum((r.aprovadas for r in registros)) != int(av.replace('.', '')) or sum((r.total for r in registros)) != moeda(total):
        raise ValueError('A soma dos prestadores não confere com o total do gestor.')
    return registros
