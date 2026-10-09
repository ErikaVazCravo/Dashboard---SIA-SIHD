"""Orquestra as sete automações; entradas e planilhas permanecem em memória."""
from io import BytesIO
from pathlib import Path
import re
from zipfile import ZipFile, ZIP_DEFLATED
import pandas as pd
import pdfplumber
from processadores import gr, resumo_sia, oci_principais, oci, procedencia, pmae, hospitalar
from processadores.comum import ler_texto

RELATORIOS = {
    'rsproced': 'Sintético de Procedimentos por Unidade (RSPROCED)',
    'raproced': 'Analítico de Procedimentos por Unidade (RAPROCED)',
    'rsprocqt': 'Sintético de Proced. / Unidade Físico (RSPROCQT)',
    'raprocqt': 'Analítico de Proced. / Unidade Físico (RAPROCQT)',
    'r0570': 'Emissão Síntese da Produção — APAC (R0570)',
    'resumo_sihd': 'Resumo dos valores aprovados por CNES — todos os estabelecimentos',
    'sem_filtro': 'Demonstrativo de Procedência — SEM FILTRO',
    'eletivo': 'Demonstrativo de Procedência — INTERNAÇÃO / ELETIVO',
    'urgencia': 'Demonstrativo de Procedência — INTERNAÇÃO / URGÊNCIA',
    'pmae': 'Procedência Analítico — FAEC: Programa Agora Tem Especialistas / Componente Cirúrgico',
}
CHAVES = ['ANO', 'Mes', 'Competencia', 'CNES', 'Prestador']
MEDIDAS = ['Qtd. Apres.', 'Qtd. Aprov.', 'Vlr. Apres.', 'Vlr. Aprov.']


def validar_competencia(comp):
    if not re.fullmatch(r'(0[1-9]|1[0-2])/\d{4}', comp.strip()):
        raise ValueError('Informe a competência no formato MM/AAAA, com mês entre 01 e 12.')
    return comp.strip()


def validar_sigtap(pasta, comp):
    pasta = Path(pasta.strip().strip('"')).expanduser()
    if not pasta.is_dir():
        raise ValueError('A pasta SIGTAP não existe neste computador. Informe a pasta descompactada.')
    for nome in ('tb_procedimento.txt', 'rl_procedimento_compativel.txt'):
        if not (pasta / nome).is_file():
            raise ValueError(f'Não foi encontrado {nome} na pasta SIGTAP informada.')
    tabela = oci.carregar_tb_procedimento(pasta / 'tb_procedimento.txt')
    esperado = comp[3:] + comp[:2]
    if tabela.empty or set(tabela['DT_COMPETENCIA']) != {esperado}:
        raise ValueError('A competência de tb_procedimento.txt não corresponde à competência informada.')
    return pasta


def mapa_estabelecimentos(fonte=None):
    if not fonte:
        return {}
    mapa = {}
    for linha in ler_texto(fonte).splitlines():
        m = re.match(r'^\s*(\d{7})\s*[-;\t|]?\s*(.+?)\s*$', linha)
        if m:
            mapa[m[1]] = m[2]
    if not mapa:
        raise ValueError('O arquivo de nomes deve conter CNES de sete dígitos seguido do nome, um por linha.')
    return mapa


def excel(abas):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter', engine_kwargs={'options': {
        'in_memory': True, 'strings_to_formulas': False, 'strings_to_urls': False,
    }}) as writer:
        cabecalho = writer.book.add_format({'bold': True, 'bg_color': '#E6F2F3', 'border': 1})
        texto = writer.book.add_format({'num_format': '@'})
        dinheiro = writer.book.add_format({'num_format': '#,##0.00'})
        destaque = writer.book.add_format({'bg_color': '#FFF2CC'})
        for nome, df in abas.items():
            df.to_excel(writer, sheet_name=nome, index=False)
            ws = writer.sheets[nome]
            ws.freeze_panes(1, 0)
            if len(df.columns):
                ws.autofilter(0, 0, len(df), len(df.columns) - 1)
            for idx, coluna in enumerate(df.columns):
                ws.write_string(0, idx, str(coluna), cabecalho)
                formato = dinheiro if pd.api.types.is_float_dtype(df[coluna]) else texto if pd.api.types.is_object_dtype(df[coluna]) else None
                ws.set_column(idx, idx, min(48, max(14, len(str(coluna)) + 3)), formato)
            # Os identificadores continuam como texto, incluindo zeros à esquerda.
            if nome == 'Total_Terceiros' and not df.empty:
                for idx, row in df.iterrows():
                    if row['Tp. Proc.'] == 'PRINC' and row['Total SIGTAP'] < row['Vl. Proc. OCI']:
                        ws.write_number(df.index.get_loc(idx) + 1, df.columns.get_loc('Vl. Proc. OCI'), float(row['Vl. Proc. OCI']), destaque)
    return buffer.getvalue()


def juntar(qtd, valor, chaves):
    if qtd.empty or valor.empty:
        raise ValueError('Um dos relatórios SIA não retornou registros. Confira o arquivo e seu layout de texto.')
    qtd = qtd.groupby(chaves, as_index=False, dropna=False)[MEDIDAS[:2]].sum()
    valor = valor.groupby(chaves, as_index=False, dropna=False)[MEDIDAS[2:]].sum()
    df = qtd.merge(valor, on=chaves, how='outer')
    for col in MEDIDAS:
        df[col] = df[col].fillna(0).astype(int if col.startswith('Qtd') else float)
        if col.startswith('Vlr'):
            df[col] = df[col].round(2)
    if not df['CNES'].astype(str).str.fullmatch(r'\d{7}').all():
        raise ValueError('Há registros SIA sem CNES válido. Confira o layout dos relatórios.')
    return df[chaves + MEDIDAS].sort_values(chaves).reset_index(drop=True)


def texto_pdf(dados):
    with pdfplumber.open(BytesIO(dados)) as pdf:
        paginas = [p.extract_text(x_tolerance=2) or '' for p in pdf.pages]
    if not paginas or any(not p.strip() for p in paginas):
        raise ValueError('PDF vazio ou com página sem texto extraível. É necessário um PDF textual.')
    return '\n'.join(paginas)


def conferir_comp(df, coluna, comp, nome):
    if not df.empty and set(df[coluna].dropna().astype(str)) != {comp}:
        raise ValueError(f'{nome}: a competência não corresponde à informada ({comp}).')


def internacoes(arquivos, comp):
    frames = {}
    colunas = ['Competencia', 'CNES', 'Prestador', 'Tipo de Atendimento', 'Cod. Mun.', 'Municipio', 'Especialidade', 'Qtde. de AIHs', 'Valor']
    for chave, tipo in [('sem_filtro', 'Sem Filtro'), ('eletivo', 'Eletivo'), ('urgencia', 'Urgência')]:
        texto = texto_pdf(arquivos[chave])
        if 'DEMONSTRATIVO DE PROCEDENCIA' not in pmae.normalizar(texto) or 'ANALITICO' in pmae.normalizar(texto):
            raise ValueError(f'{RELATORIOS[chave]}: selecione o demonstrativo sintético correto.')
        if set(re.findall(r'COMPETENCIA\s*:\s*(\d{2}/\d{4})', pmae.normalizar(texto))) != {comp}:
            raise ValueError(f'{RELATORIOS[chave]}: competência ausente ou diferente da informada.')
        df = procedencia.processar_pdf(BytesIO(arquivos[chave]), tipo)
        if df.empty:
            # Um relatório filtrado pode não ter internações; somente aceite zero explícito.
            if not re.search(r'TOTAL DO GESTOR\s+0\s+0,00\s+0,00', pmae.normalizar(texto)):
                raise ValueError(f'{RELATORIOS[chave]}: nenhum registro reconhecido; confira o layout.')
            df = pd.DataFrame(columns=colunas)
        if not df.empty and (df['CNES'].isna().any() or not df['CNES'].str.fullmatch(r'\d{7}').all()):
            raise ValueError(f'{RELATORIOS[chave]}: CNES ausente ou inválido.')
        frames[chave] = df
    total, eletivo, urgencia = (frames[k] for k in ('sem_filtro', 'eletivo', 'urgencia'))
    if total.empty:
        vazio = pd.DataFrame(columns=colunas)
        return {'Total Municipios': vazio, 'Total Prestador': vazio.drop(columns=['Cod. Mun.', 'Municipio'])}
    chaves_m = ['Competencia', 'CNES', 'Prestador', 'Cod. Mun.', 'Municipio', 'Especialidade']
    # Detecta grupos filtrados ausentes no total, antes da subtração.
    base = set(total[chaves_m].itertuples(index=False, name=None))
    for df in (eletivo, urgencia):
        if set(df[chaves_m].itertuples(index=False, name=None)) - base:
            raise ValueError('Eletivo/Urgência contém grupos ausentes no relatório sem filtro.')
    outros = procedencia.calcular_outros_tipos(total, eletivo, urgencia, chaves_m)
    if (outros['Qtde. de AIHs'] < 0).any() or (outros['Valor'] < -0.01).any():
        raise ValueError('Eletivo + Urgência excede o relatório sem filtro. Confira os filtros e a competência.')
    tipos = pd.concat([eletivo, urgencia, outros], ignore_index=True)
    municipios = procedencia.adicionar_ano_mes(procedencia.agrupar_total_municipios(tipos))
    prestadores = procedencia.adicionar_ano_mes(procedencia.agrupar_total_prestador(tipos))
    return {'Total Municipios': municipios, 'Total Prestador': prestadores}


def processar(pasta_sigtap, comp, arquivos, nomes=None, progresso=None):
    comp = validar_competencia(comp)
    faltantes = [label for key, label in RELATORIOS.items() if not arquivos.get(key)]
    if faltantes:
        raise ValueError('Selecione todos os relatórios: ' + '; '.join(faltantes))
    pasta = validar_sigtap(pasta_sigtap, comp)
    mapa = mapa_estabelecimentos(nomes)
    saidas, avisos = {}, []
    def etapa(texto):
        if progresso:
            progresso(texto)

    etapa('SIA: resumo e procedimentos')
    resumo = juntar(
        resumo_sia.processar_rsp_sintetico(arquivos['rsprocqt'], 'quantidade', mapa, comp),
        resumo_sia.processar_rsp_sintetico(arquivos['rsproced'], 'valor', mapa, comp), CHAVES)
    saidas[f'Resumo_Producao_SIA_{comp.replace("/", "-")}.xlsx'] = excel({'Sheet1': resumo})
    chaves_gr = ['Processamento'] + CHAVES + ['Tipo Reg.', 'Procedimento', 'Grupo Proc.']
    procedimentos = juntar(gr.processar_raprocqt(arquivos['raprocqt'], comp, mapa),
                           gr.processar_raproced(arquivos['raproced'], comp, mapa), chaves_gr)
    if not procedimentos['Procedimento'].str.fullmatch(r'\d{10}').all():
        raise ValueError('RAPROCED/RAPROCQT: código de procedimento inválido.')
    saidas['GR_Procedimentos_SIA.xlsx'] = excel({'Sheet1': procedimentos})

    etapa('SIA: SIGTAP e produção OCI')
    base = oci.gerar_base_oci_procedimentos_e_valores(str(pasta))
    conferir_comp(base, 'Comp.', comp, 'Relações SIGTAP')
    if base[['Nome_Proc_Principal', 'Nome_Proc_Secundario', 'Vl_Proc_Principal', 'Vl_Proc_Secundario']].isna().any().any():
        raise ValueError('Há procedimentos compatíveis ausentes na tabela SIGTAP.')
    saidas['OCI_Procedimentos_e_Valores.xlsx'] = excel({'Sheet1': base})
    principais = procedimentos[procedimentos['Grupo Proc.'] == '09'].copy()
    descricoes = oci_principais.ler_tb_procedimento(pasta / 'tb_procedimento.txt')
    principais['Descr. Proced.'] = principais['Procedimento'].map(descricoes).fillna('')
    if (principais['Descr. Proced.'] == '').any():
        raise ValueError('Há procedimentos OCI do SIA ausentes no SIGTAP.')
    colunas = ['Processamento'] + CHAVES + ['Procedimento', 'Descr. Proced.'] + MEDIDAS
    saidas['SIA_OCIs.xlsx'] = excel({'Sheet1': principais[colunas]})
    dados, n_apacs, n_extraidas = oci.parse_apac(arquivos['r0570'])
    if dados.empty:
        raise ValueError('R0570: nenhuma produção OCI reconhecida. Confira o relatório e o layout.')
    conferir_comp(dados, 'Processamento', comp, 'R0570')
    if n_apacs != n_extraidas:
        raise ValueError('R0570: existem APACs OCI identificadas sem procedimentos extraídos.')
    # Usa os nomes dos próprios relatórios, complementados pelo arquivo opcional.
    mapa_relatorios = dict(zip(procedimentos['CNES'], procedimentos['Prestador']))
    mapa_relatorios.update(mapa)
    dados = oci.preencher_nomes_estabelecimentos(dados, mapa_relatorios)
    dados = oci.adicionar_nome_proced(dados, *oci.carregar_mapas_nomes_procedimentos(base))
    if (dados['Nome_Proced'] == '').any():
        raise ValueError('R0570: há procedimentos OCI sem correspondência na base SIGTAP.')
    dados = oci.organizar_coluna_data_nasc(dados)
    mapa_principal, mapa_secundario = oci.carregar_mapas_valores_sigtap(base)
    principal_apac = dados[dados['Tp. Proc.'] == 'PRINC'].set_index('N. APAC')['Procedimento'].to_dict()
    for _, row in dados.iterrows():
        codigo = oci.normalizar_proc(row['Procedimento'])
        if row['Tp. Proc.'] == 'PRINC':
            valido = codigo in mapa_principal
        else:
            valido = (oci.normalizar_proc(principal_apac.get(row['N. APAC'], '')), codigo) in mapa_secundario
        if not valido:
            raise ValueError('R0570: procedimento principal ou compatibilidade ausente na base SIGTAP.')
    vl = oci.montar_aba_vl_sigtap(dados, mapa_principal, mapa_secundario)
    vl_apac = oci.montar_aba_vl_apac_sigtap(vl)
    vl = oci.aplicar_percentual_vl_sigtap(vl, vl_apac)
    terceiros = oci.montar_aba_total_terceiros(dados, vl)
    finais = oci.organizar_abas_finais(dados, vl, vl_apac, terceiros)
    saidas['OCI_Producao_Prestadores.xlsx'] = excel(dict(zip(
        ['DADOS_OCI', 'Vl_SIGTAP', 'Vl_APAC_SIGTAP', 'Total_Terceiros'], finais)))

    etapa('SIHD: produção hospitalar')
    registros = hospitalar.interpretar(texto_pdf(arquivos['resumo_sihd']))
    hospital = pd.DataFrame([r.linha() for r in registros], columns=hospitalar.CABECALHOS)
    conferir_comp(hospital, 'COMP', comp, 'Resumo SIHD')
    for coluna in ['Compl. Federal', 'TOTAL']:
        hospital[coluna] = hospital[coluna].map(float)
    saidas['Resumo da Produção Hospitalar.xlsx'] = excel({'SIHD_Prestadores': hospital})
    etapa('SIHD: procedência e AIHs')
    saidas['SIHD_Internacoes.xlsx'] = excel(internacoes(arquivos, comp))
    aihs = pmae.processar_pdf(BytesIO(arquivos['pmae']))
    conferir_comp(aihs, 'Comp', comp, 'Procedência Analítico')
    if aihs.duplicated(['Comp', 'CNES', 'AIH']).any():
        raise ValueError('Existem AIHs repetidas no relatório analítico.')
    aihs['Valor'] = aihs['Valor'].map(float)
    saidas['SIHD_AIHs_Procedimentos.xlsx'] = excel({'AIHs': aihs})
    if not mapa:
        avisos.append('Nomes de estabelecimentos obtidos dos relatórios e dos cadastros já existentes nas automações.')
    return saidas, avisos


def compactar(saidas):
    buffer = BytesIO()
    with ZipFile(buffer, 'w', ZIP_DEFLATED) as zipfile:
        for nome, conteudo in saidas.items():
            zipfile.writestr(nome, conteudo)
    return buffer.getvalue()
