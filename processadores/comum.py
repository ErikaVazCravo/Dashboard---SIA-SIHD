"""Leitura de caminhos SIGTAP e relatórios recebidos em memória."""
from io import StringIO
from pathlib import Path


def ler_texto(fonte):
    if isinstance(fonte, str):
        dados = Path(fonte).read_bytes()
    elif isinstance(fonte, Path):
        dados = fonte.read_bytes()
    elif isinstance(fonte, bytes):
        dados = fonte
    else:
        dados = fonte.getvalue()
    if isinstance(dados, str):
        return dados
    for encoding in ('utf-8-sig', 'cp1252', 'latin-1'):
        try:
            return dados.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise ValueError('Não foi possível ler o texto do arquivo.')


def abrir_texto(fonte):
    return StringIO(ler_texto(fonte))
