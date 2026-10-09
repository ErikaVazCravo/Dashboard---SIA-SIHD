# Dashboard SIA + SIHD

Aplicação Streamlit local que reúne as sete automações existentes. Os scripts originais foram preservados; as funções migradas estão em `processadores/` e o encadeamento está em `processamento.py`.

## Iniciar

Na primeira utilização, abra um terminal nesta pasta e execute:

```powershell
python -m pip install -r requirements.txt
```

Depois, dê dois cliques em **Iniciar Dashboard.bat**, ou execute:

```powershell
python -m streamlit run app.py
```

O aplicativo abre no navegador em `http://localhost:8501`. Mantenha o terminal aberto durante o uso.

## Usar

1. Informe a competência/processamento no formato MM/AAAA.
2. Informe o caminho da pasta SIGTAP descompactada da mesma competência. Ela deve conter `tb_procedimento.txt` e `rl_procedimento_compativel.txt` diretamente.
3. Envie um arquivo em cada um dos cinco campos SIA e dos cinco campos SIHD, conforme os rótulos e filtros. Os relatórios SIA devem ser os textos originais de largura fixa, incluindo extensões mensais; os SIHD devem ser PDFs com texto selecionável e todos os estabelecimentos.
4. Opcionalmente, envie `Nomes_estabelecimentos.txt`, com CNES de sete dígitos seguido do nome em cada linha. Sem ele, os nomes vêm dos relatórios e dos cadastros SIHD das automações.
5. Clique em **Processar e gerar todos os arquivos** e baixe o ZIP ou as planilhas individuais.
6. Clique em **Limpar arquivos e resultados da sessão** ao terminar.

O nome dos arquivos enviados não precisa conter os códigos dos relatórios: cada campo determina sua finalidade. PDFs digitalizados precisam ser convertidos para texto antes do envio. A aplicação não aplica filtros a um relatório enviado no campo errado: exporte os filtros indicados no SIHD.

## Resultados

| Arquivo | Conteúdo |
| --- | --- |
| `Resumo_Producao_SIA_MM-AAAA.xlsx` | Quantidades e valores por prestador |
| `GR_Procedimentos_SIA.xlsx` | Quantidades e valores por procedimento |
| `SIA_OCIs.xlsx` | Procedimentos do grupo 09 e descrição SIGTAP |
| `OCI_Procedimentos_e_Valores.xlsx` | Base OCI e compatibilidades da tabela SIGTAP |
| `OCI_Producao_Prestadores.xlsx` | DADOS_OCI, Vl_SIGTAP, Vl_APAC_SIGTAP, Total_Terceiros |
| `Resumo da Produção Hospitalar.xlsx` | Aba SIHD_Prestadores |
| `SIHD_Internacoes.xlsx` | Total Municipios e Total Prestador; Outros = Sem filtro − Eletivo − Urgência |
| `SIHD_AIHs_Procedimentos.xlsx` | Aba AIHs do demonstrativo analítico |

## Dados e validação

Entradas, planilhas e ZIP são mantidos somente em memória da sessão, sem banco de dados, cache de dados, arquivos temporários ou gravação de resultados na pasta SIGTAP. O download salva no destino escolhido pelo navegador. A leitura da SIGTAP é somente leitura. A aplicação fica acessível apenas no computador local por padrão; estatísticas de uso do Streamlit estão desativadas. O botão de limpeza remove as referências da sessão e troca os campos de upload; isso não promete apagamento físico imediato de RAM pelo sistema operacional.

Os nomes e as abas dos resultados são preservados. As regras de extração e cálculo das automações foram reaproveitadas, com encadeamento em memória e verificações de competência, identificação dos relatórios e diferenças negativas nas internações. O resumo hospitalar e o demonstrativo analítico mantêm as conferências de subtotais dos scripts originais. Os relatórios SIA e a procedência sintética ainda dependem dos layouts fixos das automações; mudança de layout exige revisão e conferência com os relatórios originais. A formatação dos resumos foi padronizada; a produção OCI mantém os destaques da automação.

Uma competência pode incluir produção SIA de meses anteriores: os analíticos mantêm a competência de produção de cada linha e recebem a competência informada na coluna Processamento. Não é exigida igualdade entre competência de produção e processamento.

Os testes automatizados usam dados sintéticos e verificam o encadeamento, cálculos, arquivos Excel e ausência de gravação de dados durante o processamento. A extração do demonstrativo analítico também foi conferida com um relatório real de 08/2026: 126 AIHs, total de R$ 827.508,73, com subtotais consistentes e sem AIHs duplicadas. A validação de uma competência completa, incluindo SIA e SIGTAP, ainda depende dos respectivos arquivos. Os relatórios reais e cadastros locais não são incluídos no repositório.

```powershell
python -m unittest discover -s tests -v
```
