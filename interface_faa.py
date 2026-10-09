"""Fluxo Streamlit independente para contagem FAA das UPAs."""
import streamlit as st
from processadores.faa import UPAS, processar_faa


def mostrar_faa():
    def invalidar():
        st.session_state.pop('faa_resultado', None)
        st.session_state.pop('faa_resumo', None)

    def limpar():
        invalidar()
        st.session_state['faa_geracao'] = st.session_state.get('faa_geracao', 0) + 1
        for chave in list(st.session_state):
            if chave.startswith('faa_entrada_'):
                del st.session_state[chave]

    st.subheader('Contagem de procedimentos · UPAs FAA')
    st.write('Informe a competência e envie os relatórios XLSX originais do FAA nos campos das respectivas UPAs.')
    st.caption('Você pode processar uma ou mais UPAs nesta opção. Ela funciona independentemente dos relatórios SIA/SIHD e da SIGTAP.')
    st.button('Limpar arquivos e resultados FAA', on_click=limpar, key='faa_limpar')
    geracao = st.session_state.get('faa_geracao', 0)
    comp = st.text_input('Competência FAA (MM/AAAA)', placeholder='08/2026',
        key=f'faa_entrada_comp_{geracao}', on_change=invalidar)
    uploads = {}
    for coluna, (cnes, dados) in zip(st.columns(3), UPAS.items()):
        with coluna:
            uploads[cnes] = st.file_uploader(f'{dados["Prestador"]} · CNES {cnes}', type=['xlsx'],
                accept_multiple_files=True, key=f'faa_entrada_{cnes}_{geracao}', on_change=invalidar)
    st.caption('A competência informada será aplicada a todos os arquivos. Confira o período dos relatórios antes de enviar. Envie arquivos distintos, sem sobreposição de dados.')
    pronto = bool(comp.strip() and any(uploads.values()))
    if st.button('Processar contagem FAA', type='primary', disabled=not pronto, key='faa_processar'):
        invalidar()
        with st.spinner('Consolidando procedimentos das UPAs…'):
            try:
                resultado, resumo = processar_faa(comp, {
                    cnes: [(arquivo.name, arquivo.getvalue()) for arquivo in arquivos]
                    for cnes, arquivos in uploads.items()})
            except Exception as erro:
                st.error(str(erro))
            else:
                st.session_state['faa_resultado'] = resultado
                st.session_state['faa_resumo'] = resumo
    if 'faa_resultado' in st.session_state:
        st.success('Contagem concluída. A planilha contém as abas ESPECIALIDADES_PADRONIZADAS e ORIGINAL.')
        st.dataframe(st.session_state['faa_resumo'], hide_index=True)
        st.download_button('Baixar FAA_UPAs-RESULTADO.xlsx', data=st.session_state['faa_resultado'],
            file_name='FAA_UPAs-RESULTADO.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', on_click='ignore')
