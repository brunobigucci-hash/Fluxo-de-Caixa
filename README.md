# Gerador de Fluxo de Caixa Mensal de Obra

Este app gera uma planilha físico-financeira no mesmo modelo do arquivo `fluxo_caixa_mensal_infinite_IA.xlsx`.

## Abas geradas

1. Dashboard
2. Fluxo Mensal
3. Macroserviços
4. Atividades Base
5. Premissas

## Rodar localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Publicar no Streamlit Community Cloud

1. Suba estes arquivos para um repositório GitHub.
2. Acesse https://share.streamlit.io/
3. Clique em **New app**.
4. Selecione o repositório.
5. Main file: `app.py`.
6. Clique em **Deploy**.

## Ajuste das regras

As regras de classificação ficam no arquivo `engine.py`, variável `MACRO_RULES`.
