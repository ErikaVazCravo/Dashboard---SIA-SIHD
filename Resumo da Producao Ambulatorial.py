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

messagebox.showinfo(
    "Compilador Sintético do SIA",
    "COMPILADOR SINTÉTICO DO SIA\n\n"
    "Antes de continuar:\n\n"
    "1. O arquivo Nomes_estabelecimentos.txt deve estar na mesma pasta deste programa.\n\n"
    "2. O formato obrigatório do arquivo é:\n\n"
    "CNES com 7 dígitos + TAB + Nome do estabelecimento\n\n"
    "Exemplo:\n\n"
    "2082527\tSANTA CASA\n"
    "5860059\tPREF - CTA\n"
    "7581114\tCEO\n\n"
    "Importante: entre o CNES e o nome deve haver uma TABULAÇÃO, não apenas espaço.\n\n"
    "Na próxima tela será solicitada a competência dos relatórios."
)


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


# ============================================================
# COMPETÊNCIA INFORMADA PELO USUÁRIO
# ============================================================

competencia_informada = simpledialog.askstring(
    "Competência",
    "Informe a competência dos relatórios no formato MM/AAAA.\n\n"
    "Exemplos:\n"
    "05/2026\n"
    "06/2026"
)

if not competencia_informada:
    raise SystemExit("Competência não informada.")

competencia_informada = competencia_informada.strip()

if not re.fullmatch(r"\d{2}/\d{4}", competencia_informada):
    messagebox.showerror(
        "Competência inválida",
        "A competência deve estar no formato MM/AAAA.\n\n"
        "Exemplo: 05/2026"
    )
    raise SystemExit("Competência inválida.")

mes_informado = int(competencia_informada[:2])
ano_informado = int(competencia_informada[3:])

if mes_informado < 1 or mes_informado > 12:
    messagebox.showerror(
        "Competência inválida",
        "O mês deve estar entre 01 e 12."
    )
    raise SystemExit("Competência inválida.")


# ============================================================
# SELEÇÃO DOS RELATÓRIOS
# ============================================================

messagebox.showinfo(
    "Selecionar relatórios",
    "Selecione os arquivos RSPROCQT e RSPROCED.\n\n"
    "Pode selecionar arquivos .JAN, .FEV, .MAR, .ABR, .MAI, .JUN, .JUL etc."
)

arquivos_relatorios = filedialog.askopenfilenames(
    title="Selecione os arquivos RSPROCQT e RSPROCED",
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


def limpar_texto_excel(valor):
    if isinstance(valor, str):
        return re.sub(r"[\x00-\x1F]", "", valor)
    return valor


def parse_quantidade(valor):
    if valor is None:
        return 0

    valor = str(valor).strip()

    if not valor:
        return 0

    valor = valor.replace(".", "").replace(" ", "").replace(",", ".")

    try:
        return int(float(valor))
    except:
        return 0


def parse_valor(valor):
    if valor is None:
        return 0.0

    valor = str(valor).strip()

    if not valor:
        return 0.0

    valor = valor.replace(".", "").replace(" ", "").replace(",", ".")

    try:
        return round(float(valor), 2)
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


def processar_rsp_sintetico(caminho, tipo_relatorio, mapa_nomes):
    with open(caminho, "r", encoding="utf-8", errors="ignore") as f:
        linhas = [linha.rstrip("\n") for linha in f]

    dados = []

    cnes = ""
    prestador = ""

    for linha in linhas:
        linha_strip = linha.strip()

        if not linha_strip:
            continue

        if (
            linha_strip.startswith("Pagina")
            or linha_strip.startswith("SMS-ARARAQ")
            or regex_data.search(linha_strip)
            or linha_strip.startswith("********")
            or linha_strip.startswith("  PROCED.")
        ):
            continue

        if linha_strip.startswith("Unidade"):
            cnes = limpar_cnes(linha[8:15].strip())
            nome_relatorio = linha[16:51].strip()
            prestador = obter_nome_estabelecimento(cnes, nome_relatorio, mapa_nomes)
            continue

        if linha_strip.startswith("TOTAL DA UNIDADE"):
            total_apresentado = linha[75:87].strip()
            total_aprovado = linha[88:102].strip()

            if tipo_relatorio == "quantidade":
                dados.append({
                    "ANO": ano_informado,
                    "Mes": mes_informado,
                    "Competencia": competencia_informada,
                    "CNES": cnes,
                    "Prestador": prestador,
                    "Qtd. Apres.": parse_quantidade(total_apresentado),
                    "Qtd. Aprov.": parse_quantidade(total_aprovado),
                })

            elif tipo_relatorio == "valor":
                dados.append({
                    "ANO": ano_informado,
                    "Mes": mes_informado,
                    "Competencia": competencia_informada,
                    "CNES": cnes,
                    "Prestador": prestador,
                    "Vlr. Apres.": parse_valor(total_apresentado),
                    "Vlr. Aprov.": parse_valor(total_aprovado),
                })

    df = pd.DataFrame(dados)

    if not df.empty:
        df = df.map(limpar_texto_excel)

    return df


# ============================================================
# CARREGA NOMES DOS ESTABELECIMENTOS
# ============================================================

mapa_nomes = carregar_mapa_estabelecimentos(arquivo_nomes)

print(f"\n✅ Nomes carregados do arquivo: {len(mapa_nomes)}")
print(f"📄 Arquivo usado: {arquivo_nomes}")


# ============================================================
# PROCESSA OS ARQUIVOS SELECIONADOS
# ============================================================

dados_qtd = []
dados_vlr = []

print("\n========================================")
print(" PROCESSANDO RELATÓRIOS SINTÉTICOS DO SIA")
print("========================================")

for caminho in arquivos_relatorios:
    nome_arquivo = os.path.basename(caminho)
    nome_upper = nome_arquivo.upper()

    if "RSPROCQT" in nome_upper:
        print(f"\n📊 Processando quantidade: {nome_arquivo}")
        df = processar_rsp_sintetico(caminho, "quantidade", mapa_nomes)

        if not df.empty:
            dados_qtd.append(df)
            print(f"✅ Linhas extraídas: {len(df)}")
        else:
            print("⚠️ Nenhuma linha extraída.")

    elif "RSPROCED" in nome_upper:
        print(f"\n💰 Processando valor: {nome_arquivo}")
        df = processar_rsp_sintetico(caminho, "valor", mapa_nomes)

        if not df.empty:
            dados_vlr.append(df)
            print(f"✅ Linhas extraídas: {len(df)}")
        else:
            print("⚠️ Nenhuma linha extraída.")

    else:
        print(f"\n⚠️ Arquivo ignorado, pois não contém RSPROCQT ou RSPROCED no nome:")
        print(nome_arquivo)


# ============================================================
# CONSOLIDAÇÃO FINAL
# ============================================================

if not dados_qtd and not dados_vlr:
    print("\n❌ Nenhum relatório RSPROCQT ou RSPROCED foi processado.")
    raise SystemExit("Nenhum relatório processado.")


df_qtd = pd.concat(dados_qtd, ignore_index=True) if dados_qtd else pd.DataFrame()
df_vlr = pd.concat(dados_vlr, ignore_index=True) if dados_vlr else pd.DataFrame()

chaves = [
    "ANO",
    "Mes",
    "Competencia",
    "CNES",
    "Prestador"
]

if not df_qtd.empty:
    df_qtd = df_qtd.groupby(chaves, as_index=False, dropna=False).agg({
        "Qtd. Apres.": "sum",
        "Qtd. Aprov.": "sum"
    })

if not df_vlr.empty:
    df_vlr = df_vlr.groupby(chaves, as_index=False, dropna=False).agg({
        "Vlr. Apres.": "sum",
        "Vlr. Aprov.": "sum"
    })


if not df_qtd.empty and not df_vlr.empty:
    df_final = pd.merge(
        df_qtd,
        df_vlr,
        on=chaves,
        how="outer"
    )
elif not df_qtd.empty:
    df_final = df_qtd.copy()
    df_final["Vlr. Apres."] = 0.0
    df_final["Vlr. Aprov."] = 0.0
else:
    df_final = df_vlr.copy()
    df_final["Qtd. Apres."] = 0
    df_final["Qtd. Aprov."] = 0


df_final["Qtd. Apres."] = df_final["Qtd. Apres."].fillna(0).astype(int)
df_final["Qtd. Aprov."] = df_final["Qtd. Aprov."].fillna(0).astype(int)
df_final["Vlr. Apres."] = df_final["Vlr. Apres."].fillna(0).astype(float).round(2)
df_final["Vlr. Aprov."] = df_final["Vlr. Aprov."].fillna(0).astype(float).round(2)

df_final = df_final[
    [
        "ANO",
        "Mes",
        "Competencia",
        "CNES",
        "Prestador",
        "Qtd. Apres.",
        "Qtd. Aprov.",
        "Vlr. Apres.",
        "Vlr. Aprov."
    ]
]

df_final = df_final.sort_values(
    by=[
        "ANO",
        "Mes",
        "CNES",
        "Prestador"
    ]
)


# ============================================================
# VERIFICAÇÕES
# ============================================================

print("\n========================================")
print(" VERIFICAÇÃO DE CONSISTÊNCIA")
print("========================================")

cnes_mapa = set(mapa_nomes.keys())
cnes_resultado = set(df_final["CNES"].dropna().astype(str).unique())

cnes_nao_localizados = sorted(cnes_resultado - cnes_mapa)

if cnes_nao_localizados:
    print("\n⚠️ CNES não localizados no Nomes_estabelecimentos.txt:")
    for cnes in cnes_nao_localizados:
        print(f" - {cnes}")
else:
    print("\n✅ Todos os CNES foram localizados no Nomes_estabelecimentos.txt.")

duplicados = df_final.groupby(
    ["ANO", "Mes", "Competencia", "CNES"]
)["Prestador"].nunique().reset_index()

duplicados = duplicados[duplicados["Prestador"] > 1]

if not duplicados.empty:
    print("\n⚠️ Atenção: existe mais de um nome de prestador para o mesmo CNES.")
    print(duplicados)
else:
    print("\n✅ Quantidades e valores foram vinculados pelo mesmo CNES e competência.")


# ============================================================
# SALVA EXCEL FINAL
# ============================================================

nome_saida = f"Resumo_Producao_SIA_{competencia_informada.replace('/', '-')}.xlsx"

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