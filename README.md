# Gerador de Fluxo de Caixa Mensal de Obra

Este app recebe dois arquivos:

1. **Orçamento / PLS** em `.xls` ou `.xlsx`
2. **Cronograma físico** em `.xlsx`

E gera uma planilha Excel com o mesmo padrão do modelo validado:

- Dashboard
- Fluxo Mensal
- Macroserviços
- Atividades Base
- Premissas

## Como rodar localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Como publicar no Streamlit Cloud

1. Suba estes arquivos no GitHub:
   - `app.py`
   - `engine.py`
   - `requirements.txt`
   - `README.md`
2. Acesse https://share.streamlit.io/
3. Crie um novo app apontando para `app.py`.

## Como adaptar para outras obras

Edite as regras `MACRO_RULES` no arquivo `engine.py`.
Cada regra define:

- Nome do macroserviço
- Códigos do orçamento
- Critério de distribuição
- Palavras-chave do cronograma
- Fallback quando não encontrar atividade no cronograma

## Observação técnica

O modelo é gerencial: distribui o orçamento por macroserviço e usa o cronograma para lançar os custos mensalmente conforme as frentes executivas/pavimentos.
