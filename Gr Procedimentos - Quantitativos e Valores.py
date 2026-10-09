import os
import re
import sys
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog


# ============================================================
# INTERFACE INICIAL
# ============================================================

root = tk.Tk()
root.withdraw()

# ============================================================
# COMPETÊNCIA DO PROCESSAMENTO
# ============================================================

messagebox.showinfo(
    "Compilador Analítico do SIA",
    "COMPILADOR ANALÍTICO DO SIA\n\n"

    "Antes de continuar, observe as informações abaixo:\n\n"

    "1. Processamento corresponde à competência em que a produção foi processada pelo SIA.\n\n"

    "2. Competência corresponde ao mês de realização da produção ambulatorial. "
    "O relatório pode conter mais de uma competência."

    "3. O arquivo Nomes_estabelecimentos.txt deve estar na mesma pasta deste programa.\n\n"

    "4. O arquivo Nomes_estabelecimentos.txt deve possuir o seguinte formato:\n\n"

    "5. Selecione os arquivos RAPROCQT e RAPROCED referentes ao mesmo processamento.\n\n"

    "CNES (7 dígitos) + TABULAÇÃO + Nome do estabelecimento\n\n"

    "Exemplo:\n\n"

    "2082527\tSANTA CASA\n"
    "5860059\tPREF - CTA\n"
    "7581114\tCEO\n\n"

    "Importante: entre o CNES e o nome deve existir uma TABULAÇÃO, e não apenas espaços."
)


processamento = simpledialog.askstring(
    "Competência do processamento",
    "Qual é a competência do processamento?\n\n"
    "Informe no formato MM/AAAA.\n\n"
    "Exemplo:\n"
    "06/2026"
)

if not processamento:
    raise SystemExit("Competência do processamento não informada.")

processamento = processamento.strip()

if not re.fullmatch(r"\d{2}/\d{4}", processamento):
    messagebox.showerror(
        "Competência inválida",
        "A competência deve ser informada no formato MM/AAAA."
    )
    raise SystemExit("Competência inválida.")

# ============================================================
# LOCAL DO CÓDIGO E ARQUIVO DE NOMES
# ============================================================

if getattr(sys, "frozen", False):
    pasta_codigo = os.path.dirname(sys.executable)
else:
    pasta_codigo = os.path.dirname(os.path.abspath(__file__))

arquivo_nomes = os.path.join(pasta_codigo, "Nomes_estabelecimentos.txt")

if not os.path.exists(arquivo_nomes):
    messagebox.showerror(
        "Arquivo não encontrado",
        f"O arquivo Nomes_estabelecimentos.txt não foi encontrado no mesmo local do programa.\n\n"
        f"Local esperado:\n{arquivo_nomes}"
    )
    raise SystemExit("Arquivo Nomes_estabelecimentos.txt não encontrado.")

def extrair_competencia_linha(competencia_raw):
    competencia_raw = str(competencia_raw).strip()

    if len(competencia_raw) == 6 and competencia_raw.isdigit():
        ano = int(competencia_raw[:4])
        mes = int(competencia_raw[4:6])
        competencia = f"{mes:02d}/{ano}"
        return ano, mes, competencia

    return None, None, ""

# ============================================================
# SELEÇÃO DOS ARQUIVOS
# ============================================================

messagebox.showinfo(
    "Selecionar relatórios",
    "Selecione os arquivos RAPROCQT e RAPROCED.\n\n"
    "Pode selecionar arquivos .JAN, .FEV, .MAR, .ABR, .MAI, .JUN, .JUL etc."
)

arquivos_relatorios = filedialog.askopenfilenames(
    title="Selecione os arquivos RAPROCQT e RAPROCED",
    filetypes=[
        ("Relatórios SIA", "*.*"),
        ("Todos os arquivos", "*.*")
    ]
)

if not arquivos_relatorios:
    raise SystemExit("Operação cancelada pelo usuário.")


# ============================================================
# SELEÇÃO DA PASTA DE SAÍDA
# ============================================================

pasta_saida = filedialog.askdirectory(
    title="Selecione a pasta onde o Excel final será salvo"
)

if not pasta_saida:
    raise SystemExit("Operação cancelada pelo usuário.")


# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================

regex_data = re.compile(r"\b\d{2}/\d{2}/\d{4}\b")


def limpar_cnes(valor):
    if pd.isna(valor):
        return ""
    return re.sub(r"\D", "", str(valor)).strip()


def clean_excel_string(s):
    if isinstance(s, str):
        return re.sub(r"[\x00-\x1F]", "", s)
    return s


def parse_int_from_field(s):
    if s is None:
        return 0

    s = str(s).strip()

    if not s:
        return 0

    s = s.replace(".", "").replace(" ", "").replace(",", ".")

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

    s = s.replace(".", "").replace(" ", "").replace(",", ".")

    try:
        return round(float(s), 2)
    except:
        return 0.0


def carregar_mapa_estabelecimentos(caminho):
    mapa = {}

    with open(caminho, "r", encoding="utf-8-sig", errors="ignore") as f:
        for numero_linha, linha in enumerate(f, start=1):
            linha = linha.rstrip("\n")

            if not linha.strip():
                continue

            if linha.strip().startswith("-"):
                continue

            if "CNES" in linha.upper() and "NOME" in linha.upper():
                continue

            partes = linha.split("\t", maxsplit=1)

            if len(partes) != 2:
                print(
                    f"⚠️ Linha ignorada no Nomes_estabelecimentos.txt "
                    f"(sem TAB): linha {numero_linha} -> {repr(linha)}"
                )
                continue

            cnes = limpar_cnes(partes[0])
            nome = partes[1].strip()

            if len(cnes) == 7 and nome:
                mapa[cnes] = nome
            else:
                print(
                    f"⚠️ Linha ignorada no Nomes_estabelecimentos.txt "
                    f"(CNES inválido ou nome vazio): linha {numero_linha} -> {repr(linha)}"
                )

    return mapa


def obter_nome_estabelecimento(cnes, nome_relatorio, mapa):
    cnes = limpar_cnes(cnes)

    if cnes in mapa:
        return mapa[cnes]

    return str(nome_relatorio).strip()


def grupo_procedimento(procedimento):
    procedimento = str(procedimento).strip()

    if len(procedimento) >= 2 and procedimento[:2].isdigit():
        return procedimento[:2]

    return ""


# ============================================================
# CARREGA NOMES DOS ESTABELECIMENTOS
# ============================================================

MAPA_NOMES = carregar_mapa_estabelecimentos(arquivo_nomes)

print(f"\n✅ Nomes carregados do arquivo: {len(MAPA_NOMES)}")
print(f"📄 Arquivo usado: {arquivo_nomes}")


# ============================================================
# PROCESSA RAPROCQT - QUANTIDADES
# ============================================================

def processar_raprocqt(caminho):
    with open(caminho, "r", encoding="utf-8", errors="ignore") as f:
        linhas = [ln.rstrip("\n") for ln in f]

    registros = []

    cnes_atual = ""
    estabelecimento_atual = ""

    for linha_raw in linhas:
        if not linha_raw:
            continue

        linha_strip = linha_raw.strip()

        if (
            linha_strip.startswith("Pagina")
            or linha_strip.startswith("SMS-ARARAQ")
            or regex_data.search(linha_strip)
            or linha_strip.startswith("********")
            or linha_strip.startswith("CMP")
            or linha_strip.startswith("_ SINTETICO")
            or linha_strip.upper().startswith("CMP   PROCEDIMENTO")
        ):
            continue

        if linha_strip.startswith("Unidade"):
            cnes_atual = limpar_cnes(linha_raw[8:15].strip())
            estabelecimento_relatorio = linha_raw[16:].strip()
            estabelecimento_atual = obter_nome_estabelecimento(
                cnes_atual,
                estabelecimento_relatorio,
                MAPA_NOMES
            )
            continue

        if linha_strip.startswith("TOTAL DA UNIDADE"):
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

            tipo_reg = "APAC" if apac else "BPA/RAAS"

            registros.append({
                "Processamento": processamento,
                "ANO": ano_linha,
                "Mes": mes_linha,
                "Competencia": competencia_linha,
                "CNES": cnes_atual,
                "Prestador": estabelecimento_atual,
                "Tipo Reg.": tipo_reg,
                "Procedimento": procedimento,
                "Grupo Proc.": grupo_procedimento(procedimento),
                "Qtd. Apres.": parse_int_from_field(qtd_apres_raw),
                "Qtd. Aprov.": parse_int_from_field(qtd_aprov_raw),
            })

    df = pd.DataFrame(registros)

    if df.empty:
        return df

    df = df.map(clean_excel_string)

    df = df.groupby(
    [
        "Processamento",
        "ANO",
        "Mes",
        "Competencia",
        "CNES",
        "Prestador",
        "Tipo Reg.",
        "Procedimento",
        "Grupo Proc."
    ],
    as_index=False,
    dropna=False
).agg({
    "Qtd. Apres.": "sum",
    "Qtd. Aprov.": "sum"
})

    return df


# ============================================================
# PROCESSA RAPROCED - VALORES
# ============================================================

def processar_raproced(caminho):
    with open(caminho, "r", encoding="utf-8", errors="ignore") as f:
        linhas = [ln.rstrip("\n") for ln in f]

    registros = []

    cnes_atual = ""
    estabelecimento_atual = ""

    for linha_raw in linhas:
        if not linha_raw:
            continue

        linha_strip = linha_raw.strip()

        if (
            linha_strip.startswith("Pagina")
            or linha_strip.startswith("SMS-ARARAQ")
            or regex_data.search(linha_strip)
            or linha_strip.startswith("********")
            or linha_strip.startswith("CMP")
            or linha_strip.startswith("_ SINTETICO")
            or linha_strip.upper().startswith("CMP   PROCEDIMENTO")
        ):
            continue

        if linha_strip.startswith("Unidade"):
            cnes_atual = limpar_cnes(linha_raw[8:15].strip())
            estabelecimento_relatorio = linha_raw[16:].strip()
            estabelecimento_atual = obter_nome_estabelecimento(
                cnes_atual,
                estabelecimento_relatorio,
                MAPA_NOMES
            )
            continue

        if linha_strip.startswith("TOTAL DA UNIDADE"):
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

            tipo_reg = "APAC" if apac else "BPA/RAAS"

            registros.append({
                "Processamento": processamento,
                "ANO": ano_linha,
                "Mes": mes_linha,
                "Competencia": competencia_linha,
                "CNES": cnes_atual,
                "Prestador": estabelecimento_atual,
                "Tipo Reg.": tipo_reg,
                "Procedimento": procedimento,
                "Grupo Proc.": grupo_procedimento(procedimento),
                "Vlr. Apres.": parse_float_from_field(vlr_apres_raw),
                "Vlr. Aprov.": parse_float_from_field(vlr_aprov_raw),
            })

    df = pd.DataFrame(registros)

    if df.empty:
        return df

    df = df.map(clean_excel_string)

    df = df.groupby(
    [
        "Processamento",
        "ANO",
        "Mes",
        "Competencia",
        "CNES",
        "Prestador",
        "Tipo Reg.",
        "Procedimento",
        "Grupo Proc."
    ],
    as_index=False,
    dropna=False
).agg({
    "Vlr. Apres.": "sum",
    "Vlr. Aprov.": "sum"
})
    
    return df


# ============================================================
# PROCESSAMENTO PRINCIPAL
# ============================================================

dados_qtde = []
dados_valor = []

print("\n========================================")
print(" PROCESSANDO RELATÓRIOS ANALÍTICOS DO SIA")
print("========================================")

for caminho in arquivos_relatorios:
    nome_arquivo = os.path.basename(caminho)
    nome_upper = nome_arquivo.upper()

    try:
        if "RAPROCQT" in nome_upper:
            print(f"\n📊 Processando quantidade: {nome_arquivo}")
            df_qt = processar_raprocqt(caminho)

            if not df_qt.empty:
                dados_qtde.append(df_qt)
                print(f"✅ Linhas de quantidade: {len(df_qt)}")
            else:
                print("⚠️ Nenhuma linha de quantidade encontrada.")

        elif "RAPROCED" in nome_upper:
            print(f"\n💰 Processando valor: {nome_arquivo}")
            df_vl = processar_raproced(caminho)

            if not df_vl.empty:
                dados_valor.append(df_vl)
                print(f"✅ Linhas de valor: {len(df_vl)}")
            else:
                print("⚠️ Nenhuma linha de valor encontrada.")

        else:
            print(f"\n⚠️ Arquivo ignorado, pois não contém RAPROCQT ou RAPROCED no nome:")
            print(nome_arquivo)

    except Exception as e:
        print(f"\n❌ Erro ao processar {nome_arquivo}: {e}")


# ============================================================
# CONSOLIDAÇÃO
# ============================================================

if not dados_qtde and not dados_valor:
    print("\n❌ Nenhum relatório RAPROCQT ou RAPROCED foi processado.")
    raise SystemExit("Nenhum relatório processado.")


df_qtde_total = pd.concat(dados_qtde, ignore_index=True) if dados_qtde else pd.DataFrame()
df_valor_total = pd.concat(dados_valor, ignore_index=True) if dados_valor else pd.DataFrame()

chaves = [
    "Processamento",
    "ANO",
    "Mes",
    "Competencia",
    "CNES",
    "Prestador",
    "Tipo Reg.",
    "Procedimento",
    "Grupo Proc."
]

if not df_qtde_total.empty:
    df_qtde_total = df_qtde_total.groupby(
        chaves,
        as_index=False,
        dropna=False
    ).agg({
        "Qtd. Apres.": "sum",
        "Qtd. Aprov.": "sum"
    })

if not df_valor_total.empty:
    df_valor_total = df_valor_total.groupby(
        chaves,
        as_index=False,
        dropna=False
    ).agg({
        "Vlr. Apres.": "sum",
        "Vlr. Aprov.": "sum"
    })


if not df_qtde_total.empty and not df_valor_total.empty:
    df_final = pd.merge(
        df_qtde_total,
        df_valor_total,
        on=chaves,
        how="outer"
    )
elif not df_qtde_total.empty:
    df_final = df_qtde_total.copy()
    df_final["Vlr. Apres."] = 0.0
    df_final["Vlr. Aprov."] = 0.0
else:
    df_final = df_valor_total.copy()
    df_final["Qtd. Apres."] = 0
    df_final["Qtd. Aprov."] = 0


df_final["Qtd. Apres."] = df_final["Qtd. Apres."].fillna(0).astype(int)
df_final["Qtd. Aprov."] = df_final["Qtd. Aprov."].fillna(0).astype(int)
df_final["Vlr. Apres."] = df_final["Vlr. Apres."].fillna(0).astype(float).round(2)
df_final["Vlr. Aprov."] = df_final["Vlr. Aprov."].fillna(0).astype(float).round(2)

df_final = df_final[
    [
        "Processamento",
        "ANO",
        "Mes",
        "Competencia",
        "CNES",
        "Prestador",
        "Tipo Reg.",
        "Procedimento",
        "Grupo Proc.",
        "Qtd. Apres.",
        "Qtd. Aprov.",
        "Vlr. Apres.",
        "Vlr. Aprov."
    ]
]

df_final = df_final.sort_values(
    by=[
        "Processamento",
        "ANO",
        "Mes",
        "CNES",
        "Prestador",
        "Tipo Reg.",
        "Procedimento"
    ]
)


# ============================================================
# VERIFICAÇÕES
# ============================================================

print("\n========================================")
print(" VERIFICAÇÃO DE CONSISTÊNCIA")
print("========================================")

cnes_arquivo_nomes = set(MAPA_NOMES.keys())
cnes_relatorio = set(df_final["CNES"].dropna().astype(str).unique())

cnes_nao_localizados = sorted(cnes_relatorio - cnes_arquivo_nomes)

if cnes_nao_localizados:
    print("\n⚠️ CNES não localizados no Nomes_estabelecimentos.txt:")
    for cnes in cnes_nao_localizados:
        print(f" - {cnes}")
else:
    print("\n✅ Todos os CNES foram localizados no Nomes_estabelecimentos.txt.")


duplicados = df_final.groupby(
    [
        "ANO",
        "Mes",
        "Competencia",
        "CNES",
        "Tipo Reg.",
        "Procedimento"
    ]
)["Prestador"].nunique().reset_index()

duplicados = duplicados[duplicados["Prestador"] > 1]

if not duplicados.empty:
    print("\n⚠️ Atenção: existe mais de um nome de prestador para o mesmo CNES/procedimento.")
    print(duplicados)
else:
    print("\n✅ Quantidades e valores foram vinculados pelo mesmo CNES, competência, tipo e procedimento.")


# ============================================================
# SALVAR EXCEL FINAL
# ============================================================

nome_saida = f"GR_Procedimentos_SIA.xlsx"

caminho_saida = os.path.join(
    pasta_saida,
    nome_saida
)

df_final.to_excel(caminho_saida, index=False)

print("\n========================================")
print(" PROCESSAMENTO CONCLUÍDO")
print("========================================")
print("✅ Arquivo Excel gerado com sucesso:")
print(caminho_saida)
print(f"\nTotal de linhas no arquivo final: {len(df_final)}")