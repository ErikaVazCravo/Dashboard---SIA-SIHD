"""Aplicação local SIA + SIHD, sem persistência de arquivos enviados."""
import streamlit as st
from processamento import RELATORIOS, processar, compactar

st.set_page_config(page_title='Dashboard • SIA + SIHD', page_icon='📊', layout='wide')
st.title('Dashboard · SIA + SIHD')
st.write('Envie os relatórios da competência para gerar todas as planilhas das automações.')
st.caption('Processamento em memória, sem salvar entradas ou resultados no disco. Faça o download antes de limpar a sessão.')


def invalidar():
    st.session_state.pop('resultados', None)
    st.session_state.pop('avisos', None)
    st.session_state.pop('zip', None)


def limpar():
    invalidar()
    st.session_state['geracao'] = st.session_state.get('geracao', 0) + 1
    for chave in list(st.session_state):
        if chave.startswith('entrada_'):
            del st.session_state[chave]


st.button('Limpar arquivos e resultados da sessão', on_click=limpar)
geracao = st.session_state.get('geracao', 0)
st.subheader('Competência e SIGTAP')
col1, col2 = st.columns([1, 3])
comp = col1.text_input('Competência / processamento (MM/AAAA)', placeholder='09/2026',
                       key=f'entrada_comp_{geracao}', on_change=invalidar)
pasta = col2.text_input('Localização da pasta SIGTAP descompactada',
                         placeholder=r'C:\Tabelas\SIGTAP\202609', key=f'entrada_pasta_{geracao}', on_change=invalidar)
st.caption('A pasta precisa conter tb_procedimento.txt e rl_procedimento_compativel.txt. O caminho é lido no computador onde a aplicação está executando.')

arquivos = {}
col_sia, col_sihd = st.columns(2, gap='large')
for coluna, titulo, chaves in [
    (col_sia, 'Relatórios SIA', list(RELATORIOS)[:5]),
    (col_sihd, 'Relatórios SIHD', list(RELATORIOS)[5:]),
]:
    with coluna:
        st.subheader(titulo)
        st.caption('Arquivos de texto originais, inclusive extensões mensais (.JAN, .FEV etc.).' if titulo.endswith('SIA') else 'PDFs textuais, contendo todos os estabelecimentos, com os filtros indicados.')
        for chave in chaves:
            arquivos[chave] = st.file_uploader(RELATORIOS[chave],
                type=None if chave in list(RELATORIOS)[:5] else ['pdf'],
                accept_multiple_files=False, key=f'entrada_{chave}_{geracao}', on_change=invalidar)

with st.expander('Nomes de estabelecimentos — opcional'):
    st.write('Se você utiliza Nomes_estabelecimentos.txt, envie aqui para manter a padronização dos nomes. Sem esse arquivo, serão usados os nomes dos relatórios e das automações SIHD.')
    nomes = st.file_uploader('Nomes_estabelecimentos.txt', type=['txt'],
                             key=f'entrada_nomes_{geracao}', on_change=invalidar)

pronto = bool(comp.strip() and pasta.strip() and all(arquivos.values()))
if not pronto:
    st.info('Preencha a competência, a pasta SIGTAP e os dez campos de relatórios para gerar todos os arquivos.')
if st.button('Processar e gerar todos os arquivos', type='primary', disabled=not pronto):
    invalidar()
    with st.status('Processando relatórios…', expanded=True) as status:
        try:
            saidas, avisos = processar(pasta, comp, {k: v.getvalue() for k, v in arquivos.items()},
                                      nomes.getvalue() if nomes else None, progresso=st.write)
            st.session_state['resultados'] = saidas
            st.session_state['avisos'] = avisos
            st.session_state['zip'] = compactar(saidas)
        except Exception as erro:
            status.update(label='O processamento não foi concluído', state='error')
            st.error(str(erro))
        else:
            status.update(label=f'{len(saidas)} arquivos gerados', state='complete', expanded=False)

if 'resultados' in st.session_state:
    st.subheader('Arquivos para download')
    st.success('Processamento concluído. Baixe o ZIP completo ou as planilhas individualmente.')
    for aviso in st.session_state.get('avisos', []):
        st.info(aviso)
    st.download_button('Baixar todos os arquivos (.zip)', data=st.session_state['zip'],
        file_name=f'Dashboard_SIA_SIHD_{comp.replace("/", "-")}.zip', mime='application/zip', on_click='ignore', type='primary')
    for nome, dados in st.session_state['resultados'].items():
        st.download_button(nome, data=dados, file_name=nome,
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', on_click='ignore')
