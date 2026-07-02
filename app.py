import streamlit as st
from engine import make_excel

st.set_page_config(page_title="Gerador de Fluxo de Caixa", layout="wide")

st.title("Gerador de Fluxo de Caixa Mensal de Obra")
st.write(
    "Envie o orçamento/PLS e o cronograma no mesmo padrão dos arquivos modelo. "
    "O sistema gera uma planilha com as mesmas abas do modelo validado."
)

col1, col2 = st.columns(2)
with col1:
    budget_file = st.file_uploader("1. Orçamento / PLS (.xls ou .xlsx)", type=["xls", "xlsx"])
with col2:
    schedule_file = st.file_uploader("2. Cronograma físico (.xlsx)", type=["xlsx"])

st.divider()

if st.button("Gerar planilha", type="primary", disabled=not (budget_file and schedule_file)):
    try:
        with st.spinner("Lendo arquivos, classificando macroserviços e gerando Excel..."):
            output = make_excel(budget_file, schedule_file)
        st.success("Planilha gerada com sucesso.")
        st.download_button(
            label="Baixar planilha gerada",
            data=output,
            file_name="fluxo_caixa_mensal_obra.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as exc:
        st.error("Erro ao gerar a planilha. Veja o detalhe abaixo.")
        st.exception(exc)

st.info(
    "Abas geradas: Dashboard, Fluxo Mensal, Macroserviços, Atividades Base e Premissas. "
    "Para adaptar a outras obras, ajuste as regras MACRO_RULES no arquivo engine.py."
)
