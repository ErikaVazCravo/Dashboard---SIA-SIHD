import re
import os
import pdfplumber
import pandas as pd
from tkinter import Tk, filedialog, messagebox


prestadores = {
    "2079763": "CAIRBAR SCHUTEL",
    "6943284": "MATERNIDADE GOTA DE LEITE",
    "2082527": "SANTA CASA",
    "9267263": "UNIDADE DE RETAGUARDA"
}


def extrair_texto_pdf(pdf_path):
    texto = []
    with pdfplumber.open(pdf_path) as pdf:
        for pagina in pdf.pages:
            conteudo = pagina.extract_text()
            if conteudo:
                texto.extend(conteudo.split("\n"))
    return texto


def obter_competencia(linhas):
    for linha in linhas:
        m = re.search(r"Competência:\s*(\d{2}/\d{4})", linha)
        if m:
            return m.group(1)
    return ""


def converter_valor(valor):
    return float(str(valor).replace(".", "").replace(",", "."))


def processar_pdf(pdf_path, tipo_atendimento):
    linhas = extrair_texto_pdf(pdf_path)
    competencia = obter_competencia(linhas)

    registros = []

    cnes = None
    estabelecimento = None
    municipio_atual = None
    cod_municipio = None
    dentro_bloco = False

    for linha in linhas:
        linha = linha.strip()

        if (
            linha.startswith("MS/DATASUS")
            or linha.startswith("DEMONSTRATIVO")
            or linha.startswith("Página:")
            or linha.startswith("SISTEMA DE INFORMAÇÕES")
            or linha.startswith("M350320001")
            or linha.startswith("Gestor:")
            or linha.startswith("MUNICIPIO")
            or linha.startswith("FILTRO")
        ):
            continue

        m = re.match(r"CNES\s+(\d{7})", linha)

        if m:
            cnes = m.group(1)
            estabelecimento = prestadores.get(cnes, f"CNES {cnes}")
            dentro_bloco = False
            continue

        if linha.startswith("MUN. ORIGEM"):
            dentro_bloco = True
            continue

        if linha.startswith("Total do CNES"):
            dentro_bloco = False
            municipio_atual = None
            cod_municipio = None
            continue

        if not dentro_bloco:
            continue

        if (
            linha.startswith("Total do Municipio")
            or linha.startswith("Total do Gestor")
            or linha.startswith("TODAS AS ESPECIALIDADES")
        ):
            continue

        padrao_municipio = (
            r"^(\d{6})\s+(.+?)\s+"
            r"(Psiquiatria|Clínico|Cirúrgico|Pediátricos|Obstétricos)\s+"
            r"(\d+)\s+"
            r"([\d\.,]+)\s+"
            r"([\d\.,]+)$"
        )

        m = re.match(padrao_municipio, linha)

        if m:
            cod_municipio = m.group(1)
            municipio_atual = m.group(2).strip()
            especialidade = m.group(3)
            qtd = int(m.group(4))
            total = converter_valor(m.group(5))

            registros.append({
                "Competencia": competencia,
                "CNES": cnes,
                "Prestador": estabelecimento,
                "Tipo de Atendimento": tipo_atendimento,
                "Cod. Mun.": cod_municipio,
                "Municipio": municipio_atual,
                "Especialidade": especialidade,
                "Qtde. de AIHs": qtd,
                "Valor": total
            })

            continue

        padrao_especialidade = (
            r"^(Psiquiatria|Clínico|Cirúrgico|Pediátricos|Obstétricos)\s+"
            r"(\d+)\s+"
            r"([\d\.,]+)\s+"
            r"([\d\.,]+)$"
        )

        m = re.match(padrao_especialidade, linha)

        if m and municipio_atual:
            especialidade = m.group(1)
            qtd = int(m.group(2))
            total = converter_valor(m.group(3))

            registros.append({
                "Competencia": competencia,
                "CNES": cnes,
                "Prestador": estabelecimento,
                "Tipo de Atendimento": tipo_atendimento,
                "Cod. Mun.": cod_municipio,
                "Municipio": municipio_atual,
                "Especialidade": especialidade,
                "Qtde. de AIHs": qtd,
                "Valor": total
            })

    return pd.DataFrame(registros)


def agrupar_total_municipios(df):
    return (
        df.groupby(
            [
                "Competencia",
                "CNES",
                "Prestador",
                "Tipo de Atendimento",
                "Cod. Mun.",
                "Municipio",
                "Especialidade"
            ],
            as_index=False
        )
        .agg({
            "Qtde. de AIHs": "sum",
            "Valor": "sum"
        })
    )


def agrupar_total_prestador(df):
    return (
        df.groupby(
            [
                "Competencia",
                "CNES",
                "Prestador",
                "Tipo de Atendimento",
                "Especialidade"
            ],
            as_index=False
        )
        .agg({
            "Qtde. de AIHs": "sum",
            "Valor": "sum"
        })
    )


def calcular_outros_tipos(df_sem_filtro, df_eletivo, df_urgencia, colunas_chave):
    total_sem_filtro = (
        df_sem_filtro.groupby(colunas_chave, as_index=False)
        .agg({
            "Qtde. de AIHs": "sum",
            "Valor": "sum"
        })
    )

    total_eletivo_urgencia = pd.concat(
        [df_eletivo, df_urgencia],
        ignore_index=True
    )

    total_eletivo_urgencia = (
        total_eletivo_urgencia.groupby(colunas_chave, as_index=False)
        .agg({
            "Qtde. de AIHs": "sum",
            "Valor": "sum"
        })
    )

    resultado = total_sem_filtro.merge(
        total_eletivo_urgencia,
        on=colunas_chave,
        how="left",
        suffixes=("_sem_filtro", "_eletivo_urgencia")
    )

    resultado["Qtde. de AIHs_eletivo_urgencia"] = resultado["Qtde. de AIHs_eletivo_urgencia"].fillna(0)
    resultado["Valor_eletivo_urgencia"] = resultado["Valor_eletivo_urgencia"].fillna(0)

    resultado["Qtde. de AIHs"] = (
        resultado["Qtde. de AIHs_sem_filtro"]
        - resultado["Qtde. de AIHs_eletivo_urgencia"]
    )

    resultado["Valor"] = (
        resultado["Valor_sem_filtro"]
        - resultado["Valor_eletivo_urgencia"]
    )

    resultado = resultado[
        (resultado["Qtde. de AIHs"] != 0)
        | (resultado["Valor"].round(2) != 0)
    ].copy()

    resultado["Tipo de Atendimento"] = "Outros Tipos"
    resultado["Valor"] = resultado["Valor"].round(2)

    return resultado


def adicionar_ano_mes(df):
    """
    Adiciona as colunas ANO e MÊS a partir da Competencia.
    Competencia deve estar no formato MM/AAAA.
    """

    if df.empty:
        return df

    df = df.copy()

    df.insert(
        0,
        "ANO",
        df["Competencia"].str[-4:].astype(int)
    )

    df.insert(
        1,
        "MÊS",
        df["Competencia"].str[:2].astype(int)
    )

    return df


def main():
    root = Tk()
    root.withdraw()

    messagebox.showinfo(
        "Procedência SIHD",
        "O sistema solicitará 3 arquivos PDF:\n\n"
        "1. Internações ELETIVAS\n"
        "2. Internações de URGÊNCIA\n"
        "3. Relatório SEM FILTRO\n\n"
        "O tipo 'Outros Tipos' será calculado por diferença:\n"
        "SEM FILTRO - ELETIVO - URGÊNCIA."
    )

    arquivo_eletivo = filedialog.askopenfilename(
        title="Selecione o arquivo de internações ELETIVAS",
        filetypes=[("Arquivos PDF", "*.pdf"), ("Todos os arquivos", "*.*")]
    )

    if not arquivo_eletivo:
        return

    arquivo_urgencia = filedialog.askopenfilename(
        title="Selecione o arquivo de internações de URGÊNCIA",
        filetypes=[("Arquivos PDF", "*.pdf"), ("Todos os arquivos", "*.*")]
    )

    if not arquivo_urgencia:
        return

    arquivo_sem_filtro = filedialog.askopenfilename(
        title="Selecione o arquivo SEM FILTRO",
        filetypes=[("Arquivos PDF", "*.pdf"), ("Todos os arquivos", "*.*")]
    )

    if not arquivo_sem_filtro:
        return

    pasta_saida = filedialog.askdirectory(
        title="Selecione a pasta onde será salvo o arquivo procedencia.xlsx"
    )

    if not pasta_saida:
        return

    print("Processando ELETIVO...")
    df_eletivo = processar_pdf(arquivo_eletivo, "Eletivo")

    print("Processando URGÊNCIA...")
    df_urgencia = processar_pdf(arquivo_urgencia, "Urgência")

    print("Processando SEM FILTRO...")
    df_sem_filtro = processar_pdf(arquivo_sem_filtro, "Sem Filtro")

    if df_sem_filtro.empty:
        messagebox.showerror("Erro", "O relatório SEM FILTRO não retornou registros.")
        return

    colunas_municipios = [
        "Competencia",
        "CNES",
        "Prestador",
        "Cod. Mun.",
        "Municipio",
        "Especialidade"
    ]

    colunas_prestador = [
        "Competencia",
        "CNES",
        "Prestador",
        "Especialidade"
    ]

    outros_municipios = calcular_outros_tipos(
        df_sem_filtro,
        df_eletivo,
        df_urgencia,
        colunas_municipios
    )

    outros_prestador = calcular_outros_tipos(
        df_sem_filtro,
        df_eletivo,
        df_urgencia,
        colunas_prestador
    )

    df_total_tipos = pd.concat(
        [
            df_eletivo,
            df_urgencia,
            outros_municipios
        ],
        ignore_index=True
    )

    total_municipios = agrupar_total_municipios(df_total_tipos)

    total_prestador = pd.concat(
        [
            agrupar_total_prestador(df_eletivo),
            agrupar_total_prestador(df_urgencia),
            outros_prestador
        ],
        ignore_index=True
    )

    total_prestador = (
        total_prestador.groupby(
            [
                "Competencia",
                "CNES",
                "Prestador",
                "Tipo de Atendimento",
                "Especialidade"
            ],
            as_index=False
        )
        .agg({
            "Qtde. de AIHs": "sum",
            "Valor": "sum"
        })
    )

    total_municipios = adicionar_ano_mes(total_municipios)
    total_prestador = adicionar_ano_mes(total_prestador)

    arquivo_saida = os.path.join(pasta_saida, "SIHD_Internacoes.xlsx")

    with pd.ExcelWriter(arquivo_saida, engine="openpyxl") as writer:
        total_municipios.to_excel(writer, sheet_name="Total Municipios", index=False)
        total_prestador.to_excel(writer, sheet_name="Total Prestador", index=False)

    messagebox.showinfo(
        "Processamento concluído",
        f"Arquivo SIHD_Internacoes.xlsx gerado com sucesso!\n\n{arquivo_saida}"
    )

    print("Arquivo gerado com sucesso:")
    print(arquivo_saida)


if __name__ == "__main__":
    main()