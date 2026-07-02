import streamlit as st
from engine import make_excel

st.set_page_config(page_title="Gerador de Fluxo de Caixa", layout="wide")

st.title("Gerador de Fluxo de Caixa Mensal")
st.caption("Gera a planilha no mesmo modelo do arquivo fluxo_caixa_mensal_infinite_IA.xlsx")

with st.expander("Arquivos esperados", expanded=True):
    st.markdown(
        """
        **1. Orçamento / PLS:** arquivo `.xls` ou `.xlsx` com códigos e valores do orçamento.  
        **2. Cronograma:** arquivo `.xlsx` com atividades, datas de início/fim, lote/pavimento e serviço.

        A planilha de saída sempre terá exatamente estas abas:
        - Dashboard
        - Fluxo Mensal
        - Macroserviços
        - Atividades Base
        - Premissas
        """
    )

col1, col2 = st.columns(2)
with col1:
    budget_file = st.file_uploader("Orçamento / PLS (.xls ou .xlsx)", type=["xls", "xlsx"])
with col2:
    schedule_file = st.file_uploader("Cronograma físico (.xlsx)", type=["xlsx"])

st.divider()

if st.button("Gerar planilha", type="primary", disabled=not (budget_file and schedule_file)):
    try:
        with st.spinner("Processando orçamento, cronograma e gerando modelo..."):
            output = make_excel(budget_file, schedule_file)
        st.success("Planilha gerada com sucesso.")
        st.download_button(
            label="Baixar planilha gerada",
            data=output,
            file_name="fluxo_caixa_mensal_obra.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as exc:
        st.error("Erro ao gerar a planilha. Verifique se os arquivos seguem a mesma estrutura do modelo usado.")
        st.exception(exc)

st.info("Observação: o motor usa regras por macroserviço. Para ajustar a lógica de outras obras, edite MACRO_RULES no arquivo engine.py.")
