import os
import re
import sys
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox


# ============================================================
# INTERFACE INICIAL
# ============================================================

root = tk.Tk()
root.withdraw()

messagebox.showinfo(
    "Gerador SIA OCIs",
    "GERADOR SIA_OCIs.xlsx\n\n"
    "Este sistema irá gerar um novo arquivo Excel chamado:\n\n"
    "SIA_OCIs.xlsx\n\n"
    "Requisitos:\n\n"
    "1. Selecione o arquivo GR_Procedimentos_SIA.xlsx gerado pela automação Gr Procedimentos - Quantitativos e Valores.py.\n\n"
    "2. Deve existir uma pasta chamada SIGTAP no mesmo local deste programa.\n\n"
    "3. Dentro da pasta SIGTAP deve existir o arquivo tb_procedimento.txt.\n\n"
    "4. Serão importadas somente as linhas em que Grupo Proc. = 09.\n\n"
    "5. A coluna Descr. Proced. será preenchida conforme o código do procedimento "
    "existente na tabela tb_procedimento.txt."
)


# ============================================================
# LOCAL DO CÓDIGO E PASTA SIGTAP
# ============================================================

if getattr(sys, "frozen", False):
    pasta_codigo = os.path.dirname(sys.executable)
else:
    pasta_codigo = os.path.dirname(os.path.abspath(__file__))

pasta_sigtap = os.path.join(pasta_codigo, "SIGTAP")
arquivo_tb_procedimento = os.path.join(pasta_sigtap, "tb_procedimento.txt")


if not os.path.exists(pasta_sigtap):
    messagebox.showerror(
        "Pasta SIGTAP não encontrada",
        f"A pasta SIGTAP não foi encontrada no mesmo local do programa.\n\n"
        f"Local esperado:\n{pasta_sigtap}"
    )
    raise SystemExit("Pasta SIGTAP não encontrada.")


if not os.path.exists(arquivo_tb_procedimento):
    messagebox.showerror(
        "Arquivo tb_procedimento.txt não encontrado",
        f"O arquivo tb_procedimento.txt não foi encontrado dentro da pasta SIGTAP.\n\n"
        f"Local esperado:\n{arquivo_tb_procedimento}"
    )
    raise SystemExit("Arquivo tb_procedimento.txt não encontrado.")


# ============================================================
# SELEÇÃO DO ARQUIVO GR_Procedimentos_SIA.xlsx
# ============================================================

arquivo_gr = filedialog.askopenfilename(
    title="Selecione o arquivo GR_Procedimentos_SIA.xlsx",
    filetypes=[
        ("Arquivo Excel", "*.xlsx"),
        ("Todos os arquivos", "*.*")
    ]
)

if not arquivo_gr:
    raise SystemExit("Operação cancelada pelo usuário.")


# ============================================================
# SELEÇÃO DA PASTA DE SAÍDA
# ============================================================

pasta_saida = filedialog.askdirectory(
    title="Selecione a pasta onde o arquivo SIA_OCIs.xlsx será salvo"
)

if not pasta_saida:
    raise SystemExit("Operação cancelada pelo usuário.")


# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================

def limpar_procedimento(valor):
    if pd.isna(valor):
        return ""

    valor = re.sub(r"\D", "", str(valor)).strip()

    if valor:
        return valor.zfill(10)

    return ""


def limpar_grupo_proc(valor):
    if pd.isna(valor):
        return ""

    valor = str(valor).strip()

    if valor.endswith(".0"):
        valor = valor.replace(".0", "")

    valor = re.sub(r"\D", "", valor).strip()

    if valor:
        return valor.zfill(2)

    return ""


def limpar_texto_excel(valor):
    if isinstance(valor, str):
        return re.sub(r"[\x00-\x1F]", "", valor)
    return valor


def ler_tb_procedimento(caminho):
    """
    Layout tb_procedimento.txt:

    CO_PROCEDIMENTO: posições 1 a 10
    NO_PROCEDIMENTO: posições 11 a 260
    """

    mapa = {}

    with open(caminho, "r", encoding="latin1", errors="ignore") as f:
        for linha in f:
            if not linha.strip():
                continue

            co_procedimento = linha[0:10].strip()
            no_procedimento = linha[10:260].strip()

            co_procedimento = limpar_procedimento(co_procedimento)

            if co_procedimento:
                mapa[co_procedimento] = no_procedimento

    return mapa


# ============================================================
# CARREGA TABELA SIGTAP
# ============================================================

print("\n========================================")
print(" CARREGANDO TABELA SIGTAP")
print("========================================")

mapa_procedimentos = ler_tb_procedimento(arquivo_tb_procedimento)

print(f"✅ Procedimentos carregados: {len(mapa_procedimentos)}")
print(f"📄 Arquivo usado: {arquivo_tb_procedimento}")


# ============================================================
# CARREGA ARQUIVO GR_Procedimentos_SIA.xlsx
# ============================================================

print("\n========================================")
print(" LENDO GR_Procedimentos_SIA.xlsx")
print("========================================")

df = pd.read_excel(
    arquivo_gr,
    dtype={
        "Processamento": str,
        "ANO": str,
        "Mes": str,
        "Competencia": str,
        "CNES": str,
        "Prestador": str,
        "Tipo Reg.": str,
        "Procedimento": str,
        "Grupo Proc.": str,
    }
)

print("\n===== Valores encontrados na coluna Grupo Proc. =====")
print(df["Grupo Proc."].unique())

df.columns = [str(col).strip() for col in df.columns]

# ============================================================
# AJUSTA PROCESSAMENTO E COMPETÊNCIA
# ============================================================

df["Processamento"] = pd.to_datetime(
    df["Processamento"],
    errors="coerce"
).dt.strftime("%m/%Y")

df["Competencia"] = pd.to_datetime(
    df["Competencia"],
    errors="coerce"
).dt.strftime("%m/%Y")

colunas_obrigatorias = [
    "Processamento",
    "ANO",
    "Mes",
    "Competencia",
    "CNES",
    "Prestador",
    "Procedimento",
    "Grupo Proc.",
    "Qtd. Apres.",
    "Qtd. Aprov.",
    "Vlr. Apres.",
    "Vlr. Aprov."
]

colunas_faltantes = [
    col for col in colunas_obrigatorias
    if col not in df.columns
]

if colunas_faltantes:
    messagebox.showerror(
        "Colunas ausentes",
        "O arquivo selecionado não possui todas as colunas obrigatórias.\n\n"
        "Colunas ausentes:\n"
        + "\n".join(colunas_faltantes)
    )
    raise SystemExit("Arquivo GR_Procedimentos_SIA.xlsx inválido.")


# ============================================================
# FILTRA SOMENTE GRUPO 09
# ============================================================

df["Grupo Proc."] = (
    df["Grupo Proc."]
        .astype(str)
        .str.replace(".0", "", regex=False)
        .str.strip()
        .str.zfill(2)
)

print("\n===== Grupo Proc. após conversão =====")
print(sorted(df["Grupo Proc."].unique()))

df["Procedimento"] = df["Procedimento"].apply(limpar_procedimento)

df_oci = df[df["Grupo Proc."] == "09"].copy()

print(f"✅ Linhas encontradas com Grupo Proc. = 09: {len(df_oci)}")

if df_oci.empty:
    messagebox.showwarning(
        "Sem registros OCI",
        "Nenhuma linha com Grupo Proc. = 09 foi encontrada no arquivo selecionado."
    )


# ============================================================
# DESCRIÇÃO DO PROCEDIMENTO
# ============================================================

df_oci["Descr. Proced."] = df_oci["Procedimento"].map(mapa_procedimentos).fillna("")

procedimentos_sem_descricao = sorted(
    df_oci.loc[
        df_oci["Descr. Proced."] == "",
        "Procedimento"
    ].dropna().unique()
)

if procedimentos_sem_descricao:
    print("\n⚠️ Procedimentos sem descrição na tb_procedimento.txt:")
    for proc in procedimentos_sem_descricao:
        print(f" - {proc}")
else:
    print("\n✅ Todos os procedimentos OCI foram localizados na tb_procedimento.txt.")


# ============================================================
# MONTA RESULTADO FINAL
# ============================================================

df_final = df_oci[
    [
        "Processamento",
        "ANO",
        "Mes",
        "Competencia",
        "CNES",
        "Prestador",
        "Procedimento",
        "Descr. Proced.",
        "Qtd. Apres.",
        "Qtd. Aprov.",
        "Vlr. Apres.",
        "Vlr. Aprov."
    ]
].copy()


df_final = df_final.map(limpar_texto_excel)

df_final["Qtd. Apres."] = pd.to_numeric(df_final["Qtd. Apres."], errors="coerce").fillna(0).astype(int)
df_final["Qtd. Aprov."] = pd.to_numeric(df_final["Qtd. Aprov."], errors="coerce").fillna(0).astype(int)
df_final["Vlr. Apres."] = pd.to_numeric(df_final["Vlr. Apres."], errors="coerce").fillna(0).astype(float).round(2)
df_final["Vlr. Aprov."] = pd.to_numeric(df_final["Vlr. Aprov."], errors="coerce").fillna(0).astype(float).round(2)

df_final = df_final.sort_values(
    by=[
        "Processamento",
        "ANO",
        "Mes",
        "CNES",
        "Prestador",
        "Procedimento"
    ]
)


# ============================================================
# SALVA EXCEL FINAL
# ============================================================

caminho_saida = os.path.join(
    pasta_saida,
    "SIA_OCIs.xlsx"
)

df_final.to_excel(caminho_saida, index=False)


# ============================================================
# FINALIZAÇÃO
# ============================================================

print("\n========================================")
print(" PROCESSAMENTO CONCLUÍDO")
print("========================================")
print("✅ Arquivo Excel gerado com sucesso:")
print(caminho_saida)
print(f"\nTotal de linhas no arquivo final: {len(df_final)}")

messagebox.showinfo(
    "Processamento concluído",
    f"Arquivo SIA_OCIs.xlsx gerado com sucesso!\n\n"
    f"Local:\n{caminho_saida}\n\n"
    f"Total de linhas: {len(df_final)}"
)