"""Regras migradas de OCIs - procedimentos principais e estab executor.py; leitura sem interface e sem persistencia."""
from .comum import abrir_texto
import os
import re
import sys
import pandas as pd

def limpar_procedimento(valor):
    if pd.isna(valor):
        return ''
    valor = re.sub('\\D', '', str(valor)).strip()
    if valor:
        return valor.zfill(10)
    return ''

def limpar_grupo_proc(valor):
    if pd.isna(valor):
        return ''
    valor = str(valor).strip()
    if valor.endswith('.0'):
        valor = valor.replace('.0', '')
    valor = re.sub('\\D', '', valor).strip()
    if valor:
        return valor.zfill(2)
    return ''

def limpar_texto_excel(valor):
    if isinstance(valor, str):
        return re.sub('[\\x00-\\x1F]', '', valor)
    return valor

def ler_tb_procedimento(caminho):
    """
    Layout tb_procedimento.txt:

    CO_PROCEDIMENTO: posições 1 a 10
    NO_PROCEDIMENTO: posições 11 a 260
    """
    mapa = {}
    with abrir_texto(caminho) as f:
        for linha in f:
            if not linha.strip():
                continue
            co_procedimento = linha[0:10].strip()
            no_procedimento = linha[10:260].strip()
            co_procedimento = limpar_procedimento(co_procedimento)
            if co_procedimento:
                mapa[co_procedimento] = no_procedimento
    return mapa
