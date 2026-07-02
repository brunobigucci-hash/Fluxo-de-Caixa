from __future__ import annotations

import io
import re
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from typing import BinaryIO, Iterable

import numpy as np
import pandas as pd

# Ordem e nomes exatamente iguais ao modelo fluxo_caixa_mensal_infinite_IA.xlsx
MACRO_ORDER = [
    "Serviços técnicos",
    "Implantação do canteiro",
    "Segurança/proteções",
    "Operação do canteiro",
    "Administração/despesas gerais",
    "Preparo do terreno",
    "Contenção",
    "Locação",
    "Fundação/blocos/baldrames",
    "Reservatórios de água",
    "Piso subsolo",
    "Superestrutura",
    "Alvenaria/encunhamento",
    "Impermeabilização",
    "Revestimento parede/teto",
    "Revestimento piso",
    "Fachada",
    "Forros e molduras",
    "Granito/bancadas",
    "Louças e metais",
    "Esquadrias madeira/portas",
    "Esquadrias alumínio",
    "Serralheria",
    "Pintura",
    "Cobertura",
    "Escada emergência",
    "Instalações elétricas",
    "Instalações hidráulicas/sanitárias",
    "Instalações incêndio",
    "Instalações diversas",
    "Elevadores",
    "Ar condicionado",
    "Ligações concessionárias",
    "Periferia/pisos externos",
    "Desmobilização",
    "Serviços finais",
]

# Regras de classificação. Ajuste aqui para novas obras.
MACRO_RULES = {
    "Serviços técnicos": {"codes": ["01.01"], "keywords": ["projeto", "servico tecnico", "levantamento", "topografia", "consultoria"], "criterion": "Linear no prazo total da obra"},
    "Implantação do canteiro": {"codes": ["01.02"], "keywords": ["implantacao do canteiro", "canteiro", "tapume", "barracao"], "criterion": "Rateio por pavimento/frente conforme cronograma"},
    "Segurança/proteções": {"codes": ["01.03"], "keywords": ["seguranca", "protecao", "protecoes", "epi", "epc"], "criterion": "Linear no prazo total da obra"},
    "Operação do canteiro": {"codes": ["01.04"], "keywords": ["operacao do canteiro", "agua", "energia", "limpeza", "andaime"], "criterion": "Linear no prazo total da obra"},
    "Administração/despesas gerais": {"codes": ["01.05"], "keywords": ["administracao", "despesa geral", "engenheiro", "mestre", "almoxarife"], "criterion": "Linear no prazo total da obra"},
    "Preparo do terreno": {"codes": ["01.06"], "keywords": ["escavacao", "terraplenagem", "preparo do terreno", "demolicao", "movimento de terra"], "criterion": "Rateio por pavimento/frente conforme cronograma"},
    "Contenção": {"codes": ["02.01"], "keywords": ["contencao", "solo grampeado", "tirante", "parede diafragma"], "criterion": "Rateio por pavimento/frente conforme cronograma"},
    "Locação": {"codes": ["02.02"], "keywords": ["locacao", "gabarito"], "criterion": "Alocação por evento executivo correlato"},
    "Fundação/blocos/baldrames": {"codes": ["02.04"], "keywords": ["fundacao", "sapata", "bloco", "baldrame", "estaca", "tubulao"], "criterion": "Rateio por pavimento/frente conforme cronograma"},
    "Reservatórios de água": {"codes": ["02.05"], "keywords": ["reservatorio", "caixa d agua", "caixa d'agua"], "criterion": "Alocação por evento executivo correlato"},
    "Piso subsolo": {"codes": ["02.06"], "keywords": ["piso do subsolo", "piso subsolo"], "criterion": "Rateio por pavimento/frente conforme cronograma"},
    "Superestrutura": {"codes": ["02.07"], "keywords": ["estrutura", "superestrutura", "concreto", "forma", "armacao", "laje", "pilar", "viga", "platibanda", "barrilete"], "criterion": "Rateio por pavimento/frente conforme cronograma"},
    "Alvenaria/encunhamento": {"codes": ["02.08"], "keywords": ["alvenaria", "encunhamento", "bloco", "vedacao"], "criterion": "Rateio por pavimento/frente conforme cronograma"},
    "Impermeabilização": {"codes": ["02.10"], "keywords": ["impermeabilizacao", "manta", "argamassa polimerica"], "criterion": "Rateio por pavimento/frente conforme cronograma"},
    "Revestimento parede/teto": {"codes": ["02.11"], "keywords": ["revestimento parede", "revestimento teto", "emboço", "reboco", "massa unica", "azulejo"], "criterion": "Rateio por pavimento/frente conforme cronograma"},
    "Revestimento piso": {"codes": ["02.12"], "keywords": ["revestimento piso", "piso", "porcelanato", "ceramica", "contrapiso"], "criterion": "Rateio por pavimento/frente conforme cronograma"},
    "Fachada": {"codes": ["02.13"], "keywords": ["fachada", "pastilha", "textura externa", "caixilho"], "criterion": "Rateio por pavimento/frente conforme cronograma"},
    "Forros e molduras": {"codes": ["02.14"], "keywords": ["forro", "moldura", "gesso"], "criterion": "Rateio por pavimento/frente conforme cronograma"},
    "Granito/bancadas": {"codes": ["02.15"], "keywords": ["granito", "bancada", "soleira", "peitoril"], "criterion": "Rateio por pavimento/frente conforme cronograma"},
    "Louças e metais": {"codes": ["02.16"], "keywords": ["louca", "metal", "torneira", "bacia", "cuba"], "criterion": "Rateio por pavimento/frente conforme cronograma"},
    "Esquadrias madeira/portas": {"codes": ["02.17"], "keywords": ["porta de madeira", "esquadria madeira", "batente", "guarnicao"], "criterion": "Rateio por pavimento/frente conforme cronograma"},
    "Esquadrias alumínio": {"codes": ["02.18"], "keywords": ["esquadria aluminio", "aluminio", "contramarco", "janela"], "criterion": "Rateio por pavimento/frente conforme cronograma"},
    "Serralheria": {"codes": ["02.19"], "keywords": ["serralheria", "gradil", "guarda corpo", "corrimao"], "criterion": "Rateio por pavimento/frente conforme cronograma"},
    "Pintura": {"codes": ["02.20"], "keywords": ["pintura", "massa corrida", "tinta"], "criterion": "Rateio por pavimento/frente conforme cronograma"},
    "Cobertura": {"codes": ["02.21"], "keywords": ["cobertura", "telhado", "telha"], "criterion": "Rateio por pavimento/frente conforme cronograma"},
    "Escada emergência": {"codes": ["02.22"], "keywords": ["escada emergencia", "escada de emergencia", "pressurizacao escada"], "criterion": "Rateio por pavimento/frente conforme cronograma"},
    "Instalações elétricas": {"codes": ["03.01"], "keywords": ["eletrica", "eletroduto", "fiacao", "barramento", "quadros", "iluminacao"], "criterion": "Rateio por pavimento/frente conforme cronograma"},
    "Instalações hidráulicas/sanitárias": {"codes": ["03.02"], "keywords": ["hidraulica", "sanitaria", "esgoto", "agua fria", "agua quente", "pluvial", "prumada", "ramal"], "criterion": "Rateio por pavimento/frente conforme cronograma"},
    "Instalações incêndio": {"codes": ["03.03"], "keywords": ["incendio", "hidrante", "sprinkler", "alarme"], "criterion": "Rateio junto às frentes de instalações"},
    "Instalações diversas": {"codes": ["03.04"], "keywords": ["instalacoes diversas", "cftv", "interfone", "telefonia", "dados"], "criterion": "Rateio junto às frentes de instalações"},
    "Elevadores": {"codes": ["03.05"], "keywords": ["elevador", "elevadores"], "criterion": "Rateio por pavimento/frente conforme cronograma"},
    "Ar condicionado": {"codes": ["03.06"], "keywords": ["ar condicionado", "split", "evaporadora", "condensadora"], "criterion": "Rateio junto às frentes de instalações"},
    "Ligações concessionárias": {"codes": ["03.07"], "keywords": ["concessionaria", "ligacao", "sabesp", "enel", "comgas"], "criterion": "Alocação concentrada no final da obra"},
    "Periferia/pisos externos": {"codes": ["04.01", "04.02", "04.03"], "keywords": ["periferia", "piso externo", "paisagismo", "calcada", "urbanizacao"], "criterion": "Alocação concentrada no final da obra"},
    "Desmobilização": {"codes": ["05.01"], "keywords": ["desmobilizacao"], "criterion": "Alocação concentrada no final da obra"},
    "Serviços finais": {"codes": ["05.02"], "keywords": ["servicos finais", "limpeza final", "entrega", "arremate", "teste", "comissionamento"], "criterion": "Rateio por pavimento/frente conforme cronograma"},
}

LINEAR_MACROS = {"Serviços técnicos", "Segurança/proteções", "Operação do canteiro", "Administração/despesas gerais"}
FINAL_MACROS = {"Ligações concessionárias", "Periferia/pisos externos", "Desmobilização"}


def _norm(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip().lower()
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    text = re.sub(r"\s+", " ", text)
    return text


def _as_date(value):
    if pd.isna(value) or value == "":
        return pd.NaT
    if isinstance(value, datetime):
        return pd.Timestamp(value)
    if isinstance(value, (int, float)) and value > 30000:
        return pd.Timestamp("1899-12-30") + pd.to_timedelta(int(value), unit="D")
    return pd.to_datetime(value, errors="coerce", dayfirst=True)


def _read_excel_any(uploaded_file, header=None) -> dict[str, pd.DataFrame]:
    uploaded_file.seek(0)
    data = uploaded_file.read()
    bio = io.BytesIO(data)
    name = getattr(uploaded_file, "name", "arquivo.xlsx").lower()
    engine = "xlrd" if name.endswith(".xls") else "openpyxl"
    return pd.read_excel(bio, sheet_name=None, header=header, engine=engine)


def _last_numeric(row: Iterable) -> float:
    vals = []
    for v in row:
        if isinstance(v, str):
            cleaned = v.replace("R$", "").replace(".", "").replace(",", ".").strip()
            num = pd.to_numeric(cleaned, errors="coerce")
        else:
            num = pd.to_numeric(v, errors="coerce")
        if pd.notna(num):
            vals.append(float(num))
    return vals[-1] if vals else 0.0


def _classify_text(text: str) -> str | None:
    n = _norm(text)
    for macro, rule in MACRO_RULES.items():
        for kw in rule["keywords"]:
            if _norm(kw) in n:
                return macro
    return None


def _extract_budget_values(budget_file) -> pd.DataFrame:
    sheets = _read_excel_any(budget_file, header=None)
    totals = {m: 0.0 for m in MACRO_ORDER}
    origins = {m: [] for m in MACRO_ORDER}

    for sheet_name, df in sheets.items():
        df = df.dropna(how="all").dropna(axis=1, how="all")
        for _, row in df.iterrows():
            cells = ["" if pd.isna(x) else str(x) for x in row.tolist()]
            row_text = " | ".join(cells)
            row_norm = _norm(row_text)
            amount = _last_numeric(row.tolist())
            if abs(amount) < 0.001:
                continue
            chosen = None
            origin = ""
            # Prioridade por código, porque o orçamento costuma ser hierárquico.
            for macro, rule in MACRO_RULES.items():
                for code in rule["codes"]:
                    if re.search(rf"(^|\D){re.escape(code)}(\D|$)", row_text):
                        chosen = macro
                        origin = code
                        break
                if chosen:
                    break
            if not chosen:
                chosen = _classify_text(row_text)
                origin = "palavra-chave" if chosen else ""
            if chosen:
                totals[chosen] += amount
                if origin and origin not in origins[chosen]:
                    origins[chosen].append(origin)

    # Proteção: se a leitura por linhas capturar subtotais e total geral duplicados, ainda assim mantém o modelo operacional.
    # Em bases muito específicas, ajuste códigos/palavras em MACRO_RULES.
    rows = []
    for macro in MACRO_ORDER:
        rows.append({
            "Macroserviço": macro,
            "Valor Orçado": float(totals.get(macro, 0.0)),
            "Critério usado": MACRO_RULES[macro]["criterion"],
            "Código orçamento / origem": "+".join(origins[macro]) if origins[macro] else "+".join(MACRO_RULES[macro]["codes"]),
        })
    return pd.DataFrame(rows)


def _find_col(columns, options):
    norm_cols = {_norm(c): c for c in columns}
    for opt in options:
        o = _norm(opt)
        for n, original in norm_cols.items():
            if o == n or o in n:
                return original
    return None


def _extract_schedule(schedule_file) -> pd.DataFrame:
    sheets = _read_excel_any(schedule_file, header=0)
    best = None
    best_score = -1
    for _, df in sheets.items():
        score = 0
        cols = [_norm(c) for c in df.columns]
        for key in ["atividade", "tarefa", "nome", "inicio", "fim", "termino", "lote", "pavimento"]:
            if any(key in c for c in cols):
                score += 1
        if len(df) > 20:
            score += 2
        if score > best_score:
            best_score = score
            best = df.copy()

    if best is None or best.empty:
        raise ValueError("Não encontrei uma aba de cronograma legível.")

    df = best.copy()
    id_col = _find_col(df.columns, ["id", "codigo"])
    activity_col = _find_col(df.columns, ["atividade", "tarefa", "nome", "name"])
    service_col = _find_col(df.columns, ["servico", "serviço"])
    lot_col = _find_col(df.columns, ["lote", "pavimento", "andar", "local"])
    start_col = _find_col(df.columns, ["inicio", "início", "start"])
    end_col = _find_col(df.columns, ["fim", "termino", "término", "finish"])

    if activity_col is None:
        # Usa a primeira coluna textual se o cronograma vier sem cabeçalho claro.
        text_cols = [c for c in df.columns if df[c].astype(str).str.len().mean() > 5]
        activity_col = text_cols[0] if text_cols else df.columns[0]
    if start_col is None or end_col is None:
        raise ValueError("Não consegui identificar colunas de início e fim no cronograma.")

    out = pd.DataFrame()
    out["ID"] = df[id_col] if id_col else np.arange(1, len(df) + 1)
    out["Atividade"] = df[activity_col].fillna("").astype(str)
    out["Serviço"] = df[service_col].fillna("-").astype(str) if service_col else "-"
    out["Lote"] = df[lot_col].fillna("-").astype(str) if lot_col else "-"
    out["Início"] = df[start_col].apply(_as_date)
    out["Fim"] = df[end_col].apply(_as_date)
    out = out.dropna(subset=["Início", "Fim"])
    out = out[out["Atividade"].str.strip() != ""]
    out["Macroserviço adotado"] = out.apply(lambda r: _classify_text(f"{r['Atividade']} {r['Serviço']} {r['Lote']}") or "Serviços finais", axis=1)
    return out.reset_index(drop=True)


def _month_range(start: pd.Timestamp, end: pd.Timestamp) -> pd.DatetimeIndex:
    s = pd.Timestamp(start.year, start.month, 1)
    e = pd.Timestamp(end.year, end.month, 1)
    return pd.date_range(s, e, freq="MS")


def _allocate(months, macros_df, activities_df) -> tuple[pd.DataFrame, pd.DataFrame]:
    flow = pd.DataFrame({"Mês": months})
    for macro in MACRO_ORDER:
        flow[macro] = 0.0

    fronts = activities_df.groupby("Macroserviço adotado").size().to_dict()
    macros_df["Nº frentes/atividades"] = macros_df["Macroserviço"].map(fronts).fillna(0).astype(int)
    macros_df["Valor médio por frente"] = np.where(
        macros_df["Nº frentes/atividades"] > 0,
        macros_df["Valor Orçado"] / macros_df["Nº frentes/atividades"],
        np.nan,
    )

    month_index = {m: i for i, m in enumerate(months)}
    start_month = months[0]
    end_month = months[-1]
    total_months = max(len(months), 1)

    for _, row in macros_df.iterrows():
        macro = row["Macroserviço"]
        value = float(row["Valor Orçado"] or 0)
        if value == 0:
            continue
        if macro in LINEAR_MACROS:
            flow[macro] += value / total_months
        elif macro in FINAL_MACROS:
            flow.loc[flow["Mês"] == end_month, macro] += value
        else:
            acts = activities_df[activities_df["Macroserviço adotado"] == macro]
            if len(acts) == 0:
                # Sem frente direta: concentra no mês final, exceto itens técnicos já tratados linearmente.
                flow.loc[flow["Mês"] == end_month, macro] += value
            else:
                unit = value / len(acts)
                for _, act in acts.iterrows():
                    m = pd.Timestamp(act["Início"].year, act["Início"].month, 1)
                    if m in month_index:
                        flow.at[month_index[m], macro] += unit

    flow["Total Mensal"] = flow[MACRO_ORDER].sum(axis=1)
    flow["Acumulado"] = flow["Total Mensal"].cumsum()
    total = flow["Total Mensal"].sum()
    flow["% Física Acum."] = np.where(total > 0, flow["Acumulado"] / total, 0)
    flow["% Financeira Acum."] = np.where(total > 0, flow["Acumulado"] / total, 0)
    return flow, macros_df


def _write_excel(flow: pd.DataFrame, macros: pd.DataFrame, acts: pd.DataFrame, budget_name: str, schedule_name: str) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="dd/mm/yyyy", date_format="dd/mm/yyyy") as writer:
        wb = writer.book
        # Formats
        title_fmt = wb.add_format({"bold": True, "font_size": 16, "font_color": "white", "bg_color": "#1F4E78", "align": "center", "valign": "vcenter"})
        header_fmt = wb.add_format({"bold": True, "font_color": "white", "bg_color": "#1F4E78", "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True})
        subheader_fmt = wb.add_format({"bold": True, "font_color": "white", "bg_color": "#5B9BD5", "border": 1, "align": "center"})
        money_fmt = wb.add_format({"num_format": 'R$ #,##0.00', "border": 1})
        num_fmt = wb.add_format({"num_format": '#,##0.00', "border": 1})
        int_fmt = wb.add_format({"num_format": '#,##0', "border": 1})
        pct_fmt = wb.add_format({"num_format": '0.00%', "border": 1})
        date_fmt = wb.add_format({"num_format": 'mm/yyyy', "border": 1, "align": "center"})
        date_full_fmt = wb.add_format({"num_format": 'dd/mm/yyyy', "border": 1, "align": "center"})
        text_fmt = wb.add_format({"border": 1, "valign": "top"})
        note_fmt = wb.add_format({"text_wrap": True, "valign": "top", "border": 1})
        kpi_label = wb.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1})
        kpi_val_money = wb.add_format({"num_format": 'R$ #,##0.00', "border": 1, "bg_color": "#F2F2F2"})
        kpi_val = wb.add_format({"border": 1, "bg_color": "#F2F2F2"})

        # 1 Dashboard
        dash = wb.add_worksheet("Dashboard")
        writer.sheets["Dashboard"] = dash
        dash.merge_range("A1:H1", "Fluxo de Caixa Mensal — Obra", title_fmt)
        total_budget = float(flow["Total Mensal"].sum())
        start = flow["Mês"].min()
        end = flow["Mês"].max()
        peak_idx = int(flow["Total Mensal"].idxmax()) if len(flow) else 0
        peak_value = float(flow.loc[peak_idx, "Total Mensal"]) if len(flow) else 0
        peak_month = flow.loc[peak_idx, "Mês"] if len(flow) else pd.NaT
        kpis = [
            ["Orçamento total", total_budget],
            ["Início do cronograma", start],
            ["Fim do cronograma", end],
            ["Prazo analisado (meses)", len(flow)],
            ["Pico mensal de caixa", peak_value],
            ["Mês do pico", peak_month],
        ]
        for r, (label, val) in enumerate(kpis, start=2):
            dash.write(r, 0, label, kpi_label)
            if isinstance(val, (int, float)) and "Pico" in label or label == "Orçamento total":
                dash.write(r, 1, val, kpi_val_money)
            elif isinstance(val, pd.Timestamp):
                dash.write_datetime(r, 1, val.to_pydatetime(), date_fmt)
            else:
                dash.write(r, 1, val, kpi_val)
        dash.write_row("D3", ["Ranking", "Mês", "Desembolso"], header_fmt)
        top = flow.sort_values("Total Mensal", ascending=False).head(5).reset_index(drop=True)
        for i, row in top.iterrows():
            dash.write_number(3+i, 3, i+1, int_fmt)
            dash.write_datetime(3+i, 4, row["Mês"].to_pydatetime(), date_fmt)
            dash.write_number(3+i, 5, row["Total Mensal"], money_fmt)
        dash.write("A11", "Leitura executiva", header_fmt)
        dash.merge_range("A12:H12", f"O maior desembolso mensal ocorre em {peak_month.strftime('%m/%Y') if pd.notna(peak_month) else '-'} com aproximadamente R$ {peak_value:,.2f}.", note_fmt)
        dash.merge_range("A13:H13", "O modelo usa rateio por macroserviço: quando há atividades por pavimento/frente no cronograma, o valor do macroserviço é dividido pela quantidade de ocorrências e alocado no mês de início da frente.", note_fmt)
        dash.merge_range("A14:H14", "Serviços sem atividade direta foram distribuídos por premissa gerencial: administração/serviços técnicos linear no prazo; concessionárias e periferia no final; desmobilização no último mês.", note_fmt)
        # Hidden data for chart
        chart_row = 17
        dash.write_row(chart_row, 0, ["Mês", "Total Mensal", "Acumulado"], header_fmt)
        for i, row in flow.iterrows():
            dash.write_datetime(chart_row+1+i, 0, row["Mês"].to_pydatetime(), date_fmt)
            dash.write_number(chart_row+1+i, 1, row["Total Mensal"], money_fmt)
            dash.write_number(chart_row+1+i, 2, row["Acumulado"], money_fmt)
        chart = wb.add_chart({"type": "column"})
        chart.add_series({"name": "Desembolso mensal", "categories": ["Dashboard", chart_row+1, 0, chart_row+len(flow), 0], "values": ["Dashboard", chart_row+1, 1, chart_row+len(flow), 1]})
        chart.set_title({"name": "Desembolso Mensal"})
        chart.set_legend({"none": True})
        chart.set_y_axis({"num_format": 'R$ #,##0'})
        dash.insert_chart("A17", chart, {"x_scale": 1.45, "y_scale": 1.1})
        chart2 = wb.add_chart({"type": "line"})
        chart2.add_series({"name": "Acumulado", "categories": ["Dashboard", chart_row+1, 0, chart_row+len(flow), 0], "values": ["Dashboard", chart_row+1, 2, chart_row+len(flow), 2]})
        chart2.set_title({"name": "Curva S Financeira"})
        chart2.set_y_axis({"num_format": 'R$ #,##0'})
        dash.insert_chart("E17", chart2, {"x_scale": 1.25, "y_scale": 1.1})
        dash.set_column("A:A", 28); dash.set_column("B:B", 18); dash.set_column("D:F", 16); dash.set_column("G:H", 12)

        # 2 Fluxo Mensal
        fluxo = flow.copy()
        fluxo.to_excel(writer, sheet_name="Fluxo Mensal", index=False)
        ws = writer.sheets["Fluxo Mensal"]
        ws.set_row(0, 32, header_fmt)
        ws.freeze_panes(1, 1)
        for c, col in enumerate(fluxo.columns):
            width = 14 if c else 12
            fmt = date_fmt if col == "Mês" else pct_fmt if "%" in col else money_fmt
            ws.set_column(c, c, width, fmt)
        ws.autofilter(0, 0, len(fluxo), len(fluxo.columns)-1)
        if len(fluxo) > 0:
            ws.conditional_format(1, fluxo.columns.get_loc("Total Mensal"), len(fluxo), fluxo.columns.get_loc("Total Mensal"), {"type": "data_bar", "bar_color": "#5B9BD5"})

        # 3 Macroserviços
        m = macros[["Macroserviço", "Valor Orçado", "Critério usado", "Nº frentes/atividades", "Valor médio por frente", "Código orçamento / origem"]].copy()
        m.to_excel(writer, sheet_name="Macroserviços", index=False, startrow=0, startcol=0)
        ms = writer.sheets["Macroserviços"]
        ms.set_row(0, 28, header_fmt)
        ms.set_column("A:A", 30, text_fmt)
        ms.set_column("B:B", 16, money_fmt)
        ms.set_column("C:C", 42, note_fmt)
        ms.set_column("D:D", 16, int_fmt)
        ms.set_column("E:E", 18, money_fmt)
        ms.set_column("F:F", 22, text_fmt)
        ms.write("H1", "Resumo", header_fmt); ms.write("I1", "Valor", header_fmt)
        summary = [["Total orçamento", total_budget], ["Macros cadastrados", len(MACRO_ORDER)], ["Atividades cronograma", len(acts)]]
        for i, (label, val) in enumerate(summary, start=1):
            ms.write(i, 7, label, kpi_label)
            ms.write(i, 8, val, kpi_val_money if i == 1 else kpi_val)
        ms.freeze_panes(1, 0)
        ms.autofilter(0, 0, len(m), 5)

        # 4 Atividades Base
        a = acts[["ID", "Atividade", "Serviço", "Lote", "Início", "Fim", "Macroserviço adotado"]].copy()
        a.to_excel(writer, sheet_name="Atividades Base", index=False)
        ab = writer.sheets["Atividades Base"]
        ab.set_row(0, 28, header_fmt)
        widths = [12, 38, 22, 20, 12, 12, 30]
        for c, width in enumerate(widths):
            fmt = date_full_fmt if c in [4, 5] else text_fmt
            ab.set_column(c, c, width, fmt)
        ab.freeze_panes(1, 0)
        ab.autofilter(0, 0, len(a), len(a.columns)-1)

        # 5 Premissas
        premissas = pd.DataFrame([
            ["Modelo", "Fluxo gerencial por macroserviço, não por composição analítica."],
            ["Rateio por pavimento/frente", "O valor do macroserviço foi dividido pela quantidade de ocorrências no cronograma e lançado no mês de início da atividade."],
            ["Estrutura", "Superestrutura foi distribuída em todas as atividades de estrutura, trechos, platibandas, barrilete e caixa d’água."],
            ["Alvenaria", "Valor de alvenaria foi rateado entre Alvenaria de Vedação e Encunhamento."],
            ["Esquadrias de alumínio", "Contramarco foi considerado dentro de esquadrias de alumínio; caixilhos foram considerados dentro de fachada quando aparecem junto da fachada."],
            ["Instalações", "Elétrica usa conduítes, fiação e barramento; hidráulica usa prumadas/ramais; incêndio/diversas/ar condicionado foram rateados junto das frentes de instalações."],
            ["Administração/serviços técnicos/canteiro", "Itens sem atividade direta foram distribuídos linearmente no prazo da obra ou no evento mais coerente."],
            ["Periferia, ligações, desmobilização", "Itens concentrados no final da obra, por serem desembolsos típicos de entrega/fechamento."],
            ["% Física", "Nesta versão, a % física acumulada é proxy da % financeira acumulada. A aba permite substituir depois por medição física real."],
            ["Fonte orçamento", budget_name],
            ["Fonte cronograma", schedule_name],
        ], columns=["Item", "Premissa adotada"])
        premissas.to_excel(writer, sheet_name="Premissas", index=False)
        pr = writer.sheets["Premissas"]
        pr.set_row(0, 28, header_fmt)
        pr.set_column("A:A", 30, text_fmt)
        pr.set_column("B:B", 120, note_fmt)
        pr.freeze_panes(1, 0)
    output.seek(0)
    return output.getvalue()


def make_excel(budget_file, schedule_file) -> bytes:
    budget_name = getattr(budget_file, "name", "orçamento enviado")
    schedule_name = getattr(schedule_file, "name", "cronograma enviado")
    macros = _extract_budget_values(budget_file)
    acts = _extract_schedule(schedule_file)
    if acts.empty:
        raise ValueError("Cronograma sem atividades válidas com datas.")
    months = _month_range(acts["Início"].min(), acts["Fim"].max())
    flow, macros = _allocate(months, macros, acts)
    return _write_excel(flow, macros, acts, budget_name, schedule_name)
