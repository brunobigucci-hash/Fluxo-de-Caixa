from __future__ import annotations

import io
import re
import unicodedata
from dataclasses import dataclass
from typing import Dict, List, Tuple, Optional

import numpy as np
import pandas as pd


@dataclass
class MacroRule:
    macro: str
    budget_patterns: List[str]
    schedule_patterns: List[str]
    distribution: str
    unit_mode: str  # pavimento | periodo | evento | marco
    notes: str


MACRO_RULES: List[MacroRule] = [
    MacroRule("Serviços iniciais / canteiro", ["servico inicial", "canteiro", "instalacao provisoria", "tapume", "mobilizacao", "projeto", "sondagem"], ["implantacao", "canteiro", "mobilizacao"], "Evento/linear no período", "periodo", "Desembolso no período das atividades iniciais."),
    MacroRule("Contenção e escavação", ["contencao", "solo grampeado", "escavacao", "movimento de terra", "terra"], ["contencao", "solo grampeado", "escavacao"], "Linear no período", "periodo", "Distribuição linear entre início e fim."),
    MacroRule("Fundações", ["fundacao", "fundacoes", "sapata", "baldrame", "estaca", "bloco", "radier"], ["sapata", "baldrame", "fundacao", "fundacoes", "estaca", "bloco"], "Linear no período", "periodo", "Distribuição linear no período de fundações."),
    MacroRule("Estrutura", ["estrutura", "concreto", "forma", "arma", "aco", "escoramento", "laje", "pilar", "viga"], ["estrutura", "concreto", "pilar", "viga", "laje"], "Por pavimento/frente", "pavimento", "Valor total dividido pela quantidade de lotes/pavimentos com atividade de estrutura."),
    MacroRule("Alvenaria", ["alvenaria", "bloco", "verga", "contraverga", "encunhamento"], ["alvenaria"], "Por pavimento/frente", "pavimento", "Valor total dividido pelos lotes/pavimentos de alvenaria."),
    MacroRule("Impermeabilização", ["impermeabilizacao", "manta", "asfaltica", "primer", "proteção mecanica", "protecao mecanica"], ["impermeabilizacao", "manta"], "Por frente + linear", "pavimento", "Distribuição conforme frentes/pavimentos liberados."),
    MacroRule("Instalações hidráulicas", ["hidraul", "agua fria", "agua quente", "esgoto", "pluvial", "louca", "metais", "bomba", "reservatorio"], ["hidraul", "agua fria", "agua quente", "esgoto", "pluvial", "barrilete", "caixa d"], "Por pavimento/frente", "pavimento", "Valor por frente conforme cronograma."),
    MacroRule("Instalações elétricas", ["eletric", "eletroduto", "cabo", "quadro", "disjuntor", "tomada", "interruptor", "iluminacao", "luminaria"], ["eletric", "eletroduto", "enfiacao", "quadro", "infra seca", "instalacoes eletricas"], "Por pavimento/frente", "pavimento", "Valor por frente conforme cronograma."),
    MacroRule("Instalações especiais / incêndio / gás", ["incendio", "sprinkler", "hidrante", "gas", "glp", "detecção", "deteccao", "alarme", "spda", "cftv", "telefon", "interfon"], ["incendio", "gas", "glp", "spda", "cftv", "alarme", "telefon", "interfon"], "Por pavimento/frente", "pavimento", "Valor por frente conforme cronograma."),
    MacroRule("Esquadrias", ["esquadria", "aluminio", "ferro", "porta", "janela", "guarda corpo", "gradil", "contramarco", "vidro"], ["esquadria", "contramarco", "porta", "janela", "guarda corpo", "vidro"], "Fornecimento + instalação", "marco", "Padrão: 30% compra, 40% fabricação/entrega, 30% instalação no período das atividades."),
    MacroRule("Revestimentos / acabamentos", ["revestimento", "argamassa", "emboço", "emboco", "reboco", "ceramica", "porcelanato", "piso", "azulejo", "granito", "soleira", "rodape", "forro", "gesso"], ["revestimento", "emboço", "emboco", "reboco", "piso", "ceramica", "porcelanato", "forro", "gesso", "acabamento"], "Por pavimento/frente", "pavimento", "Valor por pavimento/frente conforme cronograma."),
    MacroRule("Pintura", ["pintura", "massa corrida", "textura", "selador", "tinta"], ["pintura", "massa corrida", "textura"], "Por pavimento/frente", "pavimento", "Valor por frente conforme cronograma."),
    MacroRule("Fachada", ["fachada", "pastilha", "brise", "pele de vidro", "revestimento externo"], ["fachada", "revestimento externo", "pastilha"], "Por trecho/frente", "pavimento", "Distribuição por trechos/frentes de fachada."),
    MacroRule("Elevadores", ["elevador"], ["elevador"], "Marcos de fornecimento", "marco", "Padrão: 30% pedido, 40% fabricação, 20% entrega, 10% instalação."),
    MacroRule("Áreas externas / paisagismo", ["paisagismo", "jardim", "grama", "urbanizacao", "urbanização", "pavimentacao", "calcada", "calçada", "playground", "piscina"], ["paisagismo", "urbanizacao", "urbanização", "area externa", "área externa", "piscina", "jardim"], "Concentrado no final", "periodo", "Linear no período, normalmente mais concentrado no fim."),
    MacroRule("Limpeza / entrega", ["limpeza", "desmobilizacao", "desmobilização", "entrega", "habite", "as built", "comissionamento"], ["limpeza", "desmobilizacao", "desmobilização", "entrega", "comissionamento"], "Evento final", "periodo", "Desembolso no período final."),
]

DEFAULT_MACRO = "Outros serviços / indiretos"


def norm(value) -> str:
    if pd.isna(value):
        return ""
    s = str(value).lower().strip()
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"\s+", " ", s)
    return s


def classify(text: str, mode: str) -> str:
    t = norm(text)
    best = DEFAULT_MACRO
    best_score = 0
    for rule in MACRO_RULES:
        pats = rule.budget_patterns if mode == "budget" else rule.schedule_patterns
        score = sum(1 for p in pats if norm(p) in t)
        if score > best_score:
            best_score = score
            best = rule.macro
    return best


def read_excel_any(file_obj) -> pd.ExcelFile:
    # pandas can read .xlsx directly. For old .xls, xlrd must be installed.
    return pd.ExcelFile(file_obj)


def find_header_row(raw: pd.DataFrame, expected: List[str]) -> int:
    for idx, row in raw.iterrows():
        row_text = " | ".join(norm(x) for x in row.tolist())
        if all(e in row_text for e in expected):
            return idx
    return 0


def parse_budget(file_obj) -> Tuple[pd.DataFrame, pd.DataFrame]:
    xl = read_excel_any(file_obj)
    sheet = "ORÇAMENTO" if "ORÇAMENTO" in xl.sheet_names else xl.sheet_names[0]
    raw = xl.parse(sheet, header=None)
    header_row = find_header_row(raw.head(20), ["descricao", "total"])
    df = xl.parse(sheet, header=header_row)
    df.columns = [str(c).strip() for c in df.columns]

    desc_col = next((c for c in df.columns if norm(c) in ["descricao", "descrição"] or "descr" in norm(c)), None)
    total_col = next((c for c in df.columns if norm(c) == "total" or "valor" in norm(c)), None)
    code_col = next((c for c in df.columns if "cod" in norm(c)), None)
    if desc_col is None or total_col is None:
        raise ValueError("Não encontrei colunas de Descrição e Total no orçamento.")

    out = df[[c for c in [code_col, desc_col, total_col] if c is not None]].copy()
    out.columns = ["Código", "Descrição", "Valor"] if code_col else ["Descrição", "Valor"]
    if "Código" not in out.columns:
        out.insert(0, "Código", "")
    out["Valor"] = pd.to_numeric(out["Valor"], errors="coerce").fillna(0)
    out["Descrição"] = out["Descrição"].astype(str)

    # Keep structured budget rows with actual values. Avoid duplicate leaf rows where possible by prioritizing coded rows.
    out = out[out["Valor"] > 0].copy()
    out["Código_norm"] = out["Código"].astype(str).str.strip()
    # Prefer rows with code like 01.02.003 (package lines); if none, keep all.
    coded = out[out["Código_norm"].str.match(r"^\d{2}(\.\d{2,3})+", na=False)].copy()
    if not coded.empty:
        out = coded
    out["Macroserviço"] = out["Descrição"].apply(lambda x: classify(x, "budget"))
    out = out.drop(columns=["Código_norm"])
    summary = out.groupby("Macroserviço", as_index=False)["Valor"].sum().sort_values("Valor", ascending=False)
    return out, summary


def parse_schedule(file_obj) -> pd.DataFrame:
    xl = read_excel_any(file_obj)
    sheet = xl.sheet_names[0]
    raw = xl.parse(sheet, header=None)
    header_row = find_header_row(raw.head(20), ["data de inicio", "data de termino"])
    df = xl.parse(sheet, header=header_row)
    df.columns = [str(c).strip() for c in df.columns]

    task_col = next((c for c in df.columns if "pacote" in norm(c) or "tarefa" in norm(c) or "atividade" in norm(c)), None)
    start_col = next((c for c in df.columns if "inicio" in norm(c)), None)
    end_col = next((c for c in df.columns if "termino" in norm(c) or "fim" in norm(c)), None)
    lot_col = next((c for c in df.columns if "lote" in norm(c) or "pav" in norm(c)), None)
    service_col = next((c for c in df.columns if norm(c) == "servico" or norm(c) == "serviço"), None)

    if task_col is None or start_col is None or end_col is None:
        raise ValueError("Não encontrei colunas de atividade, início e término no cronograma.")

    cols = [task_col, start_col, end_col]
    if lot_col: cols.append(lot_col)
    if service_col: cols.append(service_col)
    out = df[cols].copy()
    rename = {task_col: "Atividade", start_col: "Início", end_col: "Término"}
    if lot_col: rename[lot_col] = "Lote/Pavimento"
    if service_col: rename[service_col] = "Serviço"
    out = out.rename(columns=rename)
    if "Lote/Pavimento" not in out.columns: out["Lote/Pavimento"] = "Frente única"
    if "Serviço" not in out.columns: out["Serviço"] = ""
    out["Início"] = pd.to_datetime(out["Início"], errors="coerce")
    out["Término"] = pd.to_datetime(out["Término"], errors="coerce")
    out = out.dropna(subset=["Atividade", "Início", "Término"]).copy()
    out["Texto Classificação"] = out["Atividade"].astype(str) + " " + out["Serviço"].astype(str)
    out["Macroserviço"] = out["Texto Classificação"].apply(lambda x: classify(x, "schedule"))
    out["Mês Início"] = out["Início"].dt.to_period("M").dt.to_timestamp()
    out["Mês Término"] = out["Término"].dt.to_period("M").dt.to_timestamp()
    return out.drop(columns=["Texto Classificação"])


def month_range(start: pd.Timestamp, end: pd.Timestamp) -> List[pd.Timestamp]:
    if pd.isna(start) or pd.isna(end):
        return []
    return list(pd.period_range(start.to_period("M"), end.to_period("M"), freq="M").to_timestamp())


def distribute_period(value: float, start: pd.Timestamp, end: pd.Timestamp) -> Dict[pd.Timestamp, float]:
    months = month_range(start, end)
    if not months:
        return {}
    return {m: value / len(months) for m in months}


def distribute_milestone(value: float, start: pd.Timestamp, end: pd.Timestamp, macro: str) -> Dict[pd.Timestamp, float]:
    months = month_range(start, end)
    if not months:
        return {}
    if len(months) == 1:
        return {months[0]: value}
    if macro == "Elevadores":
        weights = [0.30, 0.40, 0.20, 0.10]
    else:
        weights = [0.30, 0.40, 0.30]
    selected_idx = np.linspace(0, len(months)-1, num=len(weights)).round().astype(int)
    result: Dict[pd.Timestamp, float] = {}
    for idx, w in zip(selected_idx, weights):
        result[months[idx]] = result.get(months[idx], 0) + value * w
    return result


def build_cashflow(budget_items: pd.DataFrame, budget_summary: pd.DataFrame, schedule: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    total_budget = float(budget_summary["Valor"].sum())
    rows = []
    allocation_rows = []

    macros = sorted(set(budget_summary["Macroserviço"]).union(set(schedule["Macroserviço"])))
    for macro in macros:
        value = float(budget_summary.loc[budget_summary["Macroserviço"] == macro, "Valor"].sum())
        acts = schedule[schedule["Macroserviço"] == macro].copy()
        rule = next((r for r in MACRO_RULES if r.macro == macro), None)
        if rule is None:
            rule = MacroRule(DEFAULT_MACRO, [], [], "Linear no período", "periodo", "Itens não classificados automaticamente.")
        if value <= 0:
            continue
        if acts.empty:
            # no schedule activity: place as indirect spread over whole schedule
            start, end = schedule["Início"].min(), schedule["Término"].max()
            dist = distribute_period(value, start, end)
            for m, v in dist.items(): rows.append([m, macro, v])
            allocation_rows.append([macro, value, 0, "Sem atividade direta", start, end, rule.distribution, rule.notes + " Sem atividade direta: distribuído no período total da obra."])
            continue

        start, end = acts["Início"].min(), acts["Término"].max()
        if rule.unit_mode == "pavimento":
            units = acts[["Atividade", "Lote/Pavimento", "Início", "Término"]].drop_duplicates().copy()
            # For each activity/front, count one equivalent unit. If too many rows, this still follows the schedule rhythm.
            unit_value = value / max(len(units), 1)
            for _, a in units.iterrows():
                dist = distribute_period(unit_value, a["Início"], a["Término"])
                for m, v in dist.items(): rows.append([m, macro, v])
            allocation_rows.append([macro, value, len(units), value/max(len(units),1), start, end, rule.distribution, rule.notes])
        elif rule.unit_mode == "marco":
            dist = distribute_milestone(value, start, end, macro)
            for m, v in dist.items(): rows.append([m, macro, v])
            allocation_rows.append([macro, value, len(acts), "Marcos", start, end, rule.distribution, rule.notes])
        else:
            dist = distribute_period(value, start, end)
            for m, v in dist.items(): rows.append([m, macro, v])
            allocation_rows.append([macro, value, len(acts), "Linear", start, end, rule.distribution, rule.notes])

    monthly_macro = pd.DataFrame(rows, columns=["Mês", "Macroserviço", "Valor"])
    if monthly_macro.empty:
        raise ValueError("Não foi possível gerar distribuição mensal.")
    pivot = monthly_macro.pivot_table(index="Mês", columns="Macroserviço", values="Valor", aggfunc="sum", fill_value=0).sort_index()
    pivot["Valor previsto"] = pivot.sum(axis=1)
    pivot["Acumulado"] = pivot["Valor previsto"].cumsum()
    pivot["% Financeira"] = pivot["Acumulado"] / total_budget
    # Física gerencial: usa a mesma base de macroserviços, mas ponderada por valor.
    pivot["% Física"] = pivot["% Financeira"]
    flow = pivot.reset_index()

    allocation = pd.DataFrame(allocation_rows, columns=["Macroserviço", "Valor total", "Qtd. frentes/atividades", "Valor unitário/base", "Início", "Término", "Critério", "Premissa"])
    risks = analyze_risks(flow, total_budget)
    return flow, allocation, monthly_macro, risks


def analyze_risks(flow: pd.DataFrame, total_budget: float) -> pd.DataFrame:
    avg = flow["Valor previsto"].mean()
    rows = []
    for _, r in flow.iterrows():
        month = r["Mês"]
        val = float(r["Valor previsto"])
        pct = val / total_budget if total_budget else 0
        if val > 1.35 * avg:
            rows.append([month, "Pico de desembolso", val, pct, "Mês acima de 135% da média. Avaliar compras/parcelamentos."])
        if val < 0.35 * avg:
            rows.append([month, "Baixo desembolso", val, pct, "Mês abaixo de 35% da média. Confirmar se cronograma não está subalocado."])
    return pd.DataFrame(rows, columns=["Mês", "Risco", "Valor", "% do orçamento", "Comentário"])


def make_excel(budget_file, schedule_file) -> bytes:
    budget_items, budget_summary = parse_budget(budget_file)
    schedule = parse_schedule(schedule_file)
    flow, allocation, monthly_macro, risks = build_cashflow(budget_items, budget_summary, schedule)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="dd/mm/yyyy", date_format="dd/mm/yyyy") as writer:
        flow.to_excel(writer, sheet_name="Fluxo Mensal", index=False)
        allocation.to_excel(writer, sheet_name="Macroserviços", index=False)
        monthly_macro.to_excel(writer, sheet_name="Base Distribuição", index=False)
        budget_summary.to_excel(writer, sheet_name="Resumo Orçamento", index=False)
        budget_items.to_excel(writer, sheet_name="Itens Orçamento", index=False)
        schedule.to_excel(writer, sheet_name="Atividades Cronograma", index=False)
        risks.to_excel(writer, sheet_name="Riscos", index=False)

        wb = writer.book
        money = wb.add_format({"num_format": 'R$ #,##0.00', "align": "right"})
        pct = wb.add_format({"num_format": '0.00%', "align": "right"})
        datefmt = wb.add_format({"num_format": 'mmm/yy', "align": "center"})
        header = wb.add_format({"bold": True, "font_color": "white", "bg_color": "#17365D", "border": 1, "align": "center"})
        title = wb.add_format({"bold": True, "font_size": 16, "font_color": "#17365D"})
        normal = wb.add_format({"text_wrap": True, "valign": "top"})

        for sheet_name, df in [("Fluxo Mensal", flow), ("Macroserviços", allocation), ("Base Distribuição", monthly_macro), ("Resumo Orçamento", budget_summary), ("Itens Orçamento", budget_items), ("Atividades Cronograma", schedule), ("Riscos", risks)]:
            ws = writer.sheets[sheet_name]
            ws.freeze_panes(1, 0)
            for col_num, col_name in enumerate(df.columns):
                ws.write(0, col_num, col_name, header)
                width = min(max(len(str(col_name)) + 4, 14), 42)
                ws.set_column(col_num, col_num, width, normal)
            # formats by column name
            for idx, col in enumerate(df.columns):
                n = norm(col)
                if "valor" in n or "acumulado" in n:
                    ws.set_column(idx, idx, 16, money)
                if "%" in str(col) or "financeira" in n or "fisica" in n:
                    ws.set_column(idx, idx, 14, pct)
                if "mes" in n or "inicio" in n or "termino" in n:
                    ws.set_column(idx, idx, 14, datefmt)
            if not df.empty:
                ws.autofilter(0, 0, len(df), max(len(df.columns)-1, 0))

        # Dashboard
        dash = wb.add_worksheet("Dashboard")
        writer.sheets["Dashboard"] = dash
        total = float(budget_summary["Valor"].sum())
        peak_idx = flow["Valor previsto"].idxmax()
        peak_month = pd.to_datetime(flow.loc[peak_idx, "Mês"]).strftime("%b/%Y")
        peak_value = float(flow.loc[peak_idx, "Valor previsto"])
        dash.write("A1", "Dashboard - Fluxo de Caixa da Obra", title)
        dash.write("A3", "Orçamento total", header); dash.write("B3", total, money)
        dash.write("A4", "Mês de pico", header); dash.write("B4", peak_month)
        dash.write("A5", "Valor do pico", header); dash.write("B5", peak_value, money)
        dash.write("A6", "Nº de meses", header); dash.write("B6", len(flow))
        dash.write("A8", "Premissa central", header)
        dash.write("B8", "Orçamento agrupado por macroserviço; macroserviços repetitivos distribuídos conforme frentes/pavimentos do cronograma; itens sem atividade direta distribuídos ao longo do período da obra.", normal)
        dash.set_column("A:A", 24)
        dash.set_column("B:B", 80)

        # Charts from Fluxo Mensal
        nrows = len(flow) + 1
        chart1 = wb.add_chart({"type": "column"})
        chart1.add_series({"name": "Desembolso mensal", "categories": ["Fluxo Mensal", 1, 0, nrows-1, 0], "values": ["Fluxo Mensal", 1, flow.columns.get_loc("Valor previsto"), nrows-1, flow.columns.get_loc("Valor previsto")]})
        chart1.set_title({"name": "Desembolso mensal"})
        chart1.set_y_axis({"num_format": 'R$ #,##0'})
        dash.insert_chart("A10", chart1, {"x_scale": 1.4, "y_scale": 1.2})

        chart2 = wb.add_chart({"type": "line"})
        chart2.add_series({"name": "Curva S financeira", "categories": ["Fluxo Mensal", 1, 0, nrows-1, 0], "values": ["Fluxo Mensal", 1, flow.columns.get_loc("% Financeira"), nrows-1, flow.columns.get_loc("% Financeira")]})
        chart2.add_series({"name": "Curva S física", "categories": ["Fluxo Mensal", 1, 0, nrows-1, 0], "values": ["Fluxo Mensal", 1, flow.columns.get_loc("% Física"), nrows-1, flow.columns.get_loc("% Física")]})
        chart2.set_title({"name": "Curva S"})
        chart2.set_y_axis({"num_format": "0%"})
        dash.insert_chart("I10", chart2, {"x_scale": 1.4, "y_scale": 1.2})

    output.seek(0)
    return output.read()
