# Gerador de Fluxo de Caixa Mensal de Obra

Site local em Streamlit para gerar automaticamente uma planilha físico-financeira a partir de:

1. Orçamento/PLS em `.xls` ou `.xlsx`
2. Cronograma físico em `.xlsx`

## Como rodar

No terminal, dentro desta pasta:

```bash
pip install -r requirements.txt
streamlit run app.py
```

O navegador abrirá o site. Envie os dois arquivos e clique em **Gerar planilha**.

## O que a planilha gerada contém

- Dashboard
- Fluxo Mensal
- Macroserviços
- Base Distribuição
- Resumo Orçamento
- Itens Orçamento
- Atividades Cronograma
- Riscos

## Lógica de cálculo

- O orçamento é classificado automaticamente por macroserviço.
- O cronograma é classificado automaticamente pelas atividades e lotes/pavimentos.
- Serviços repetitivos, como estrutura, alvenaria, instalações e acabamentos, são distribuídos por frente/pavimento.
- Serviços de fornecimento, como elevadores e esquadrias, são tratados por marcos.
- Serviços sem atividade direta são distribuídos no período total da obra e aparecem nas premissas.

## Onde ajustar regras

As regras ficam no arquivo `engine.py`, na lista `MACRO_RULES`.

Você pode alterar palavras-chave, critério de distribuição e premissas por macroserviço.
