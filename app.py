import io
import streamlit as st
from engine import make_excel

st.set_page_config(page_title="Gerador de Fluxo de Caixa de Obra", layout="wide")
st.title("Gerador de Fluxo de Caixa Mensal da Obra")
st.caption("Envie o orçamento/PLS e o cronograma no mesmo padrão dos arquivos do Infinite. O sistema gera uma planilha físico-financeira automaticamente.")

with st.expander("Como o cálculo funciona", expanded=True):
    st.write(
        "O sistema agrupa o orçamento por macroserviço, identifica as atividades correspondentes no cronograma "
        "e distribui os valores por período, por pavimento/frente ou por marcos de fornecimento. "
        "Macroserviços repetitivos, como estrutura e alvenaria, são rateados pelas frentes/pavimentos encontrados no cronograma."
    )

col1, col2 = st.columns(2)
with col1:
    budget = st.file_uploader("1. Orçamento / PLS (.xls ou .xlsx)", type=["xls", "xlsx"])
with col2:
    schedule = st.file_uploader("2. Cronograma físico (.xlsx)", type=["xlsx"])

if st.button("Gerar planilha", type="primary", disabled=not (budget and schedule)):
    try:
        with st.spinner("Processando arquivos e montando fluxo de caixa..."):
            data = make_excel(budget, schedule)
        st.success("Planilha gerada com sucesso.")
        st.download_button(
            "Baixar fluxo de caixa em Excel",
            data=data,
            file_name="fluxo_caixa_mensal_obra.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.error("Não consegui gerar a planilha com esses arquivos.")
        st.exception(e)

st.divider()
st.subheader("Observação técnica")
st.write(
    "Este site gera uma previsão gerencial. Para orçamento executivo fechado, revise a aba 'Macroserviços' e ajuste regras específicas de compra, fabricação e medição quando necessário."
)
