import os
import re
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog


# ============================================================
# CONFIGURAÇÕES DAS UPAS
# ============================================================

UPAS = {
    "2064731": {
        "Prestador": "UPA Central",
        "Palavras": ["central", "centro", "upa ii"]
    },
    "4047184": {
        "Prestador": "UPA Vila Xavier",
        "Palavras": ["vila xavier", "xavier"]
    },
    "9099581": {
        "Prestador": "UPA Selmi Dei",
        "Palavras": ["selmi", "selmi dei"]
    }
}


# ============================================================
# INTERFACE INICIAL
# ============================================================

root = tk.Tk()
root.withdraw()

messagebox.showinfo(
    "FAA UPAs",
    "AUTOMAÇÃO FAA - UPAs\n\n"
    "Este sistema irá ler diretamente os relatórios XLSX emitidos pelo FAA.\n\n"
    "Selecione os arquivos das UPAs:\n"
    "- UPA Central\n"
    "- UPA Vila Xavier\n"
    "- UPA Selmi Dei\n\n"
    "Não é necessário gerar CSV."
)


# ============================================================
# COMPETÊNCIA
# ============================================================

competencia = simpledialog.askstring(
    "Competência",
    "Informe a competência que está sendo trabalhada no formato MM/AAAA.\n\n"
    "Exemplo:\n"
    "06/2026"
)

if not competencia:
    raise SystemExit("Competência não informada.")

competencia = competencia.strip()

if not re.fullmatch(r"\d{2}/\d{4}", competencia):
    messagebox.showerror(
        "Competência inválida",
        "A competência deve estar no formato MM/AAAA.\n\nExemplo: 06/2026"
    )
    raise SystemExit("Competência inválida.")

mes = int(competencia[:2])
ano = int(competencia[3:])

if mes < 1 or mes > 12:
    messagebox.showerror(
        "Competência inválida",
        "O mês deve estar entre 01 e 12."
    )
    raise SystemExit("Competência inválida.")


# ============================================================
# SELEÇÃO DOS XLSX DO FAA
# ============================================================

arquivos_xlsx = filedialog.askopenfilenames(
    title="Selecione os relatórios XLSX do FAA das UPAs",
    filetypes=[
        ("Arquivos Excel", "*.xlsx"),
        ("Todos os arquivos", "*.*")
    ]
)

if not arquivos_xlsx:
    raise SystemExit("Nenhum arquivo selecionado.")


# ============================================================
# PASTA DE SAÍDA
# ============================================================

pasta_saida = filedialog.askdirectory(
    title="Selecione a pasta onde será salvo o FAA_UPAs-RESULTADO.xlsx"
)

if not pasta_saida:
    raise SystemExit("Pasta de saída não selecionada.")


# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================

def normalizar_texto(valor):
    if pd.isna(valor):
        return ""

    texto = str(valor).strip().lower()

    substituicoes = {
        "á": "a", "à": "a", "â": "a", "ã": "a",
        "é": "e", "ê": "e",
        "í": "i",
        "ó": "o", "ô": "o", "õ": "o",
        "ú": "u",
        "ç": "c"
    }

    for antigo, novo in substituicoes.items():
        texto = texto.replace(antigo, novo)

    texto = re.sub(r"\s+", " ", texto)

    return texto


def identificar_upa(caminho_arquivo, df):
    texto_busca = normalizar_texto(os.path.basename(caminho_arquivo))

    if "Unidade Prestadora" in df.columns:
        unidades = (
            df["Unidade Prestadora"]
            .dropna()
            .astype(str)
            .unique()
        )

        texto_busca += " " + " ".join(
            normalizar_texto(u) for u in unidades
        )

    for cnes, dados in UPAS.items():
        for palavra in dados["Palavras"]:
            if normalizar_texto(palavra) in texto_busca:
                return cnes, dados["Prestador"]

    return "", ""


def localizar_linha_cabecalho(caminho_arquivo):
    preview = pd.read_excel(
        caminho_arquivo,
        header=None,
        nrows=20
    )

    for indice, linha in preview.iterrows():
        valores = [
            str(v).strip() for v in linha.tolist()
            if not pd.isna(v)
        ]

        if "Especialidade" in valores and "Quantidade" in valores:
            return indice

    return None


def ler_relatorio_faa(caminho_arquivo):
    linha_cabecalho = localizar_linha_cabecalho(caminho_arquivo)

    if linha_cabecalho is None:
        print(f"⚠️ Cabeçalho não localizado no arquivo: {os.path.basename(caminho_arquivo)}")
        return pd.DataFrame()

    df = pd.read_excel(
        caminho_arquivo,
        header=linha_cabecalho
    )

    df.columns = [str(c).strip() for c in df.columns]

    colunas_necessarias = [
        "Unidade Prestadora",
        "Especialidade",
        "Quantidade"
    ]

    for coluna in colunas_necessarias:
        if coluna not in df.columns:
            print(f"⚠️ Coluna '{coluna}' não encontrada em: {os.path.basename(caminho_arquivo)}")
            return pd.DataFrame()

    cnes, prestador = identificar_upa(caminho_arquivo, df)

    if not cnes:
        print(f"⚠️ Não foi possível identificar a UPA do arquivo: {os.path.basename(caminho_arquivo)}")
        print("   O arquivo será ignorado.")
        return pd.DataFrame()

    df = df[
        df["Especialidade"].notna()
        & df["Quantidade"].notna()
    ].copy()

    df["Especialidade"] = df["Especialidade"].astype(str).str.strip()
    df["Qtde."] = pd.to_numeric(df["Quantidade"], errors="coerce").fillna(0)

    df["ANO"] = ano
    df["MÊS"] = mes
    df["Competencia"] = competencia
    df["CNES"] = cnes
    df["Prestador"] = prestador

    df_final = df[
        [
            "ANO",
            "MÊS",
            "Competencia",
            "CNES",
            "Prestador",
            "Especialidade",
            "Qtde."
        ]
    ].copy()

    return df_final


def padronizar_especialidades(df):
    df2 = df.copy()

    df2["Especialidade"] = df2["Especialidade"].replace({
        "AUXILIAR DE ENFERMAGEM": "ENFERMAGEM",
        "ENFERMEIRA (O)": "ENFERMAGEM",
        "TECNICO DE ENFERMAGEM": "ENFERMAGEM",
        "TECNICO DE ENFERMAGEM PSF": "ENFERMAGEM",

        "FARMACEUTICO": "ATEND. FARMÁRCIA",
        "TECNICO EM FARMACIA": "ATEND. FARMÁRCIA",

        "MEDICINA DO TRABALHO": "MÉDICO – CLÍNICA GERAL",
        "CLINICO GERAL (225125)": "MÉDICO – CLÍNICA GERAL",
        "MÉDICO CARDIOLOGISTA": "MÉDICO – CLÍNICA GERAL",
        "MEDICO GENERALISTA": "MÉDICO – CLÍNICA GERAL",

        "CIRURGIAO DENTISTA C.GERA": "ODONTOLOGIA",
        "TECNICO EM HIGIENE DENTAL": "ODONTOLOGIA",

        "ASSISTENTE SOCIAL": "ASSISTÊNCIA SOCIAL",
        "ORTOPEDIA/TRAUMATOLOGIA": "MÉDICO - ORTOPEDIA",
        "PEDIATRIA": "MÉDICO - PEDIATRIA"
    })

    return df2


# ============================================================
# PROCESSAMENTO DOS ARQUIVOS
# ============================================================

todos_dados = []

print("\n========================================")
print(" PROCESSANDO RELATÓRIOS XLSX DO FAA")
print("========================================")

for arquivo in arquivos_xlsx:
    print(f"\n📄 Processando: {os.path.basename(arquivo)}")

    df_arquivo = ler_relatorio_faa(arquivo)

    if not df_arquivo.empty:
        todos_dados.append(df_arquivo)
        print(f"✅ Linhas extraídas: {len(df_arquivo)}")
    else:
        print("⚠️ Nenhuma linha extraída.")


if not todos_dados:
    messagebox.showerror(
        "Sem dados",
        "Nenhum dado foi extraído dos arquivos selecionados."
    )
    raise SystemExit("Nenhum dado extraído.")


df = pd.concat(todos_dados, ignore_index=True)

df["Qtde."] = pd.to_numeric(df["Qtde."], errors="coerce").fillna(0)


# ============================================================
# ABA ORIGINAL
# ============================================================

resultado_original = (
    df.groupby(
        [
            "ANO",
            "MÊS",
            "Competencia",
            "CNES",
            "Prestador",
            "Especialidade"
        ],
        as_index=False
    )["Qtde."].sum()
)


# ============================================================
# ABA PADRONIZADA
# ============================================================

df_padronizado = padronizar_especialidades(df)

resultado_padronizado = (
    df_padronizado.groupby(
        [
            "ANO",
            "MÊS",
            "Competencia",
            "CNES",
            "Prestador",
            "Especialidade"
        ],
        as_index=False
    )["Qtde."].sum()
)


# ============================================================
# SALVAR RESULTADO
# ============================================================

arquivo_saida = os.path.join(
    pasta_saida,
    "FAA_UPAs-RESULTADO.xlsx"
)

with pd.ExcelWriter(arquivo_saida, engine="openpyxl") as writer:
    resultado_padronizado.to_excel(
        writer,
        sheet_name="ESPECIALIDADES_PADRONIZADAS",
        index=False
    )

    resultado_original.to_excel(
        writer,
        sheet_name="ORIGINAL",
        index=False
    )


print("\n========================================")
print(" PROCESSAMENTO CONCLUÍDO")
print("========================================")
print("✅ Arquivo gerado com sucesso:")
print(arquivo_saida)
print(f"\nTotal de linhas consolidadas: {len(df)}")

messagebox.showinfo(
    "Processamento concluído",
    f"Arquivo FAA_UPAs-RESULTADO.xlsx gerado com sucesso!\n\n"
    f"Local:\n{arquivo_saida}\n\n"
    f"Total de linhas consolidadas: {len(df)}"
)