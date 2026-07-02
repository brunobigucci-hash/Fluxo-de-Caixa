from __future__ import annotations

import io
import re
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from typing import BinaryIO, Dict, List, Tuple

import numpy as np
import pandas as pd


# ============================================================
# GERADOR DE FLUXO DE CAIXA MENSAL DE OBRA
# ============================================================
# Entrada:
#   1) Orçamento/PLS em .xls ou .xlsx
#   2) Cronograma em .xlsx exportado do MS Project/Prevision/etc.
# Saída:
#   Excel com exatamente as abas:
#      Dashboard
#      Fluxo Mensal
#      Macroserviços
#      Atividades Base
#      Premissas
# ============================================================


@dataclass
class MacroRule:
    name: str
    codes: List[str]
    criterio: str
    keywords: List[str]
    fallback: str = "linear"  # linear | first | last


# Regras principais. Para outras obras, ajuste nomes/códigos/palavras-chave aqui.
MACRO_RULES: List[MacroRule] = [
    MacroRule("Serviços técnicos", ["01.01"], "Linear no prazo total da obra", ["projeto", "consultoria", "sondagem", "laudo"], "linear"),
    MacroRule("Implantação do canteiro", ["01.02"], "Rateio por pavimento/frente conforme cronograma", ["implantacao", "implantaçao", "canteiro", "mobilizacao"], "first"),
    MacroRule("Segurança/proteções", ["01.03"], "Linear no prazo total da obra", ["seguranca", "protecoes", "protecao", "bandeja", "tela", "guarda corpo"], "linear"),
    MacroRule("Operação do canteiro", ["01.04"], "Linear no prazo total da obra", ["operacao do canteiro", "operacao", "agua", "energia", "limpeza canteiro"], "linear"),
    MacroRule("Administração/despesas gerais", ["01.05"], "Linear no prazo total da obra", ["administracao", "despesas gerais", "engenheiro", "mestre", "almoxarife"], "linear"),
    MacroRule("Preparo do terreno", ["01.06"], "Rateio por pavimento/frente conforme cronograma", ["escavacao", "terreno", "movimento de terra", "demolicao", "preparo"], "first"),
    MacroRule("Contenção", ["02.01"], "Rateio por pavimento/frente conforme cronograma", ["contencao", "solo grampeado", "tirante", "cortina"], "first"),
    MacroRule("Locação", ["02.02"], "Alocação por evento executivo correlato", ["locacao", "gabarito", "marcacao"], "first"),
    MacroRule("Fundação/blocos/baldrames", ["02.04"], "Rateio por pavimento/frente conforme cronograma", ["fundacao", "sapata", "sapatas", "bloco", "baldrame", "estaca", "radier"], "first"),
    MacroRule("Reservatórios de água", ["02.05"], "Alocação por evento executivo correlato", ["reservatorio", "caixa d agua", "caixa d'agua"], "last"),
    MacroRule("Piso subsolo", ["02.06"], "Rateio por pavimento/frente conforme cronograma", ["piso do subsolo", "piso subsolo", "piso garagem"], "first"),
    MacroRule("Superestrutura", ["02.07"], "Rateio por pavimento/frente conforme cronograma", ["estrutura", "superestrutura", "pilar", "viga", "laje", "platibanda", "barrilete", "caixa d agua", "caixa d'agua", "piso do pav"], "linear"),
    MacroRule("Alvenaria/encunhamento", ["02.08"], "Rateio por pavimento/frente conforme cronograma", ["alvenaria", "encunhamento", "vedacao", "parede"], "linear"),
    MacroRule("Impermeabilização", ["02.10"], "Rateio por pavimento/frente conforme cronograma", ["impermeabilizacao", "impermeabilizaçao", "manta", "argamassa polimerica"], "linear"),
    MacroRule("Revestimento parede/teto", ["02.11"], "Rateio por pavimento/frente conforme cronograma", ["revestimento parede", "revestimentos parede", "revestimento teto", "emboço", "reboco", "argamassa interna", "chapisco", "massa unica"], "linear"),
    MacroRule("Revestimento piso", ["02.12"], "Rateio por pavimento/frente conforme cronograma", ["revestimento piso", "piso interno", "contrapiso", "porcelanato", "ceramica"], "linear"),
    MacroRule("Fachada", ["02.13"], "Rateio por pavimento/frente conforme cronograma", ["fachada", "revestimento externo", "caixilho fachada", "l1", "l2"], "linear"),
    MacroRule("Forros e molduras", ["02.14"], "Rateio por pavimento/frente conforme cronograma", ["forro", "moldura", "gesso", "sanca"], "linear"),
    MacroRule("Granito/bancadas", ["02.15"], "Rateio por pavimento/frente conforme cronograma", ["granito", "bancada", "soleira", "peitoril"], "linear"),
    MacroRule("Louças e metais", ["02.16"], "Rateio por pavimento/frente conforme cronograma", ["louca", "louças", "metais", "torneira", "ralo", "carenagem"], "linear"),
    MacroRule("Esquadrias madeira/portas", ["02.17"], "Rateio por pavimento/frente conforme cronograma", ["porta madeira", "esquadria madeira", "batente", "guarnicao", "folha de porta"], "linear"),
    MacroRule("Esquadrias alumínio", ["02.18"], "Rateio por pavimento/frente conforme cronograma", ["esquadria aluminio", "esquadrias aluminio", "contramarco", "caixilho", "janela", "porta aluminio"], "linear"),
    MacroRule("Serralheria", ["02.19"], "Rateio por pavimento/frente conforme cronograma", ["serralheria", "gradil", "guarda corpo", "corrimao", "portao", "ferro"], "linear"),
    MacroRule("Pintura", ["02.20"], "Rateio por pavimento/frente conforme cronograma", ["pintura", "massa corrida", "textura", "selador"], "linear"),
    MacroRule("Cobertura", ["02.21"], "Rateio por pavimento/frente conforme cronograma", ["cobertura", "telhado", "rufo", "calha"], "last"),
    MacroRule("Escada emergência", ["02.22"], "Rateio por pavimento/frente conforme cronograma", ["escada", "emergencia"], "linear"),
    MacroRule("Instalações elétricas", ["03.01"], "Rateio por pavimento/frente conforme cronograma", ["eletrica", "elétrica", "conduite", "conduíte", "fiação", "fiacao", "barramento", "quadro", "tomada", "interruptor"], "linear"),
    MacroRule("Instalações hidráulicas/sanitárias", ["03.02"], "Rateio por pavimento/frente conforme cronograma", ["hidraulica", "hidráulica", "sanitaria", "sanitária", "esgoto", "agua fria", "agua quente", "pluvial", "prumada", "ramal", "dry wall", "shafts"], "linear"),
    MacroRule("Instalações incêndio", ["03.03"], "Rateio junto às frentes de instalações", ["incendio", "incêndio", "sprinkler", "hidrante", "alarme"], "linear"),
    MacroRule("Instalações diversas", ["03.04"], "Rateio junto às frentes de instalações", ["instalacoes diversas", "instalações diversas", "dados", "telefonia", "cftv", "interfone"], "linear"),
    MacroRule("Elevadores", ["03.05"], "Rateio por pavimento/frente conforme cronograma", ["elevador", "elevadores"], "last"),
    MacroRule("Ar condicionado", ["03.06"], "Rateio junto às frentes de instalações", ["ar condicionado", "condensadora", "evaporadora", "dreno ar"], "linear"),
    MacroRule("Ligações concessionárias", ["03.07"], "Alocação concentrada no final da obra", ["ligacao definitiva", "ligação definitiva", "concessionaria", "concessionárias", "enel", "sabesp"], "last"),
    MacroRule("Periferia/pisos externos", ["04.01", "04.02", "04.03"], "Alocação concentrada no final da obra", ["periferia", "piso externo", "paisagismo", "calcada", "calçada", "area externa", "área externa", "diversos"], "last"),
    MacroRule("Desmobilização", ["05.01"], "Alocação concentrada no final da obra", ["desmobilizacao", "desmobilização"], "last"),
    MacroRule("Serviços finais", ["05.02"], "Rateio por pavimento/frente conforme cronograma", ["servicos finais", "serviços finais", "limpeza final", "vistoria", "entrega", "habite-se"], "last"),
]


PREMISSAS = [
    ["Modelo", "Fluxo gerencial por macroserviço, não por composição analítica."],
    ["Rateio por pavimento/frente", "O valor do macroserviço foi dividido pela quantidade de ocorrências no cronograma e lançado no mês de início da atividade."],
    ["Estrutura", "Superestrutura foi distribuída em todas as atividades de estrutura, trechos, platibandas, barrilete e caixa d’água identificadas no cronograma."],
    ["Alvenaria", "Valor de alvenaria foi rateado entre Alvenaria de Vedação e Encunhamento quando tais atividades existirem no cronograma."],
    ["Esquadrias de alumínio", "Contramarco foi considerado dentro de esquadrias de alumínio; caixilhos podem ser classificados em fachada quando aparecerem junto da fachada."],
    ["Instalações", "Elétrica, hidráulica, incêndio, diversas e ar condicionado foram rateadas pelas frentes executivas correspondentes do cronograma."],
    ["Administração/serviços técnicos/canteiro", "Itens sem atividade direta foram distribuídos linearmente no prazo da obra ou no evento executivo mais coerente."],
    ["Serviços finais/periferia/ligações", "Itens finais foram concentrados no final do cronograma quando não existir atividade executiva explícita."],
    ["Validação", "Esta é uma previsão gerencial. Para uso contratual, valide os critérios de compras por marcos e fornecimentos especiais."],
]


def _norm(value) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"\s+", " ", text)
    return text


def _read_excel_any(file_obj, sheet_name=None, header=None) -> pd.DataFrame:
    """Lê xls/xlsx via pandas. No Streamlit o objeto tem .name."""
    name = getattr(file_obj, "name", "uploaded.xlsx")
    engine = "xlrd" if str(name).lower().endswith(".xls") else "openpyxl"
    try:
        file_obj.seek(0)
    except Exception:
        pass
    return pd.read_excel(file_obj, sheet_name=sheet_name, header=header, engine=engine)


def _excel_sheet_names(file_obj) -> List[str]:
    name = getattr(file_obj, "name", "uploaded.xlsx")
    engine = "xlrd" if str(name).lower().endswith(".xls") else "openpyxl"
    try:
        file_obj.seek(0)
    except Exception:
        pass
    xls = pd.ExcelFile(file_obj, engine=engine)
    return xls.sheet_names


def _find_header_row(raw: pd.DataFrame, required_groups: List[List[str]]) -> int:
    for idx in range(min(len(raw), 80)):
        row = [_norm(x) for x in raw.iloc[idx].tolist()]
        joined = " | ".join(row)
        ok = True
        for group in required_groups:
            if not any(term in joined for term in group):
                ok = False
                break
        if ok:
            return idx
    raise ValueError("Não consegui identificar o cabeçalho da planilha.")


def _find_col(columns: List[str], candidates: List[str]) -> str | None:
    norm_map = {col: _norm(col) for col in columns}
    for cand in candidates:
        nc = _norm(cand)
        for col, ncol in norm_map.items():
            if nc == ncol or nc in ncol:
                return col
    return None


def extract_schedule(schedule_file) -> pd.DataFrame:
    sheets = _excel_sheet_names(schedule_file)
    # Preferência por abas usuais de cronograma
    preferred = None
    for sh in sheets:
        if _norm(sh) in {"activities", "atividades", "cronograma", "schedule"}:
            preferred = sh
            break
    if preferred is None:
        preferred = sheets[0]

    raw = _read_excel_any(schedule_file, sheet_name=preferred, header=None)
    header_row = _find_header_row(raw, [["inicio", "start", "data de inicio", "data início"], ["fim", "termino", "término", "finish", "data de termino", "data de término"]])

    data = raw.iloc[header_row + 1 :].copy()
    data.columns = [str(x).strip() if pd.notna(x) else f"col_{i}" for i, x in enumerate(raw.iloc[header_row].tolist())]
    data = data.dropna(how="all")

    cols = list(data.columns)
    col_id = _find_col(cols, ["ID", "Código", "Codigo", "WBS", "Identificador"])
    col_activity = _find_col(cols, ["Pacote de trabalho/tarefas", "Atividade", "Tarefa", "Nome da tarefa", "Task Name", "Name"])
    col_service = _find_col(cols, ["Serviço", "Servico", "Disciplina"])
    col_lote = _find_col(cols, ["Lote", "Pavimento", "Local", "Trecho"])
    col_start = _find_col(cols, ["Data de Início", "Data Inicio", "Início", "Inicio", "Start", "Start Date", "Early Start", "Início Agendado"])
    col_end = _find_col(cols, ["Data de Término", "Data Termino", "Término", "Termino", "Fim", "Finish", "Finish Date", "Early Finish", "Término Agendado"])

    if col_activity is None:
        raise ValueError("Não consegui identificar a coluna de atividade/tarefa no cronograma.")
    if col_start is None or col_end is None:
        raise ValueError("Não consegui identificar colunas de início e fim no cronograma.")

    out = pd.DataFrame({
        "ID": data[col_id] if col_id else range(1, len(data) + 1),
        "Atividade": data[col_activity],
        "Serviço": data[col_service] if col_service else "-",
        "Lote": data[col_lote] if col_lote else "-",
        "Início": pd.to_datetime(data[col_start], errors="coerce"),
        "Fim": pd.to_datetime(data[col_end], errors="coerce"),
    })
    out = out.dropna(subset=["Atividade", "Início", "Fim"]).copy()
    out["Atividade"] = out["Atividade"].astype(str).str.strip()
    out["Serviço"] = out["Serviço"].fillna("-").astype(str).str.strip()
    out["Lote"] = out["Lote"].fillna("-").astype(str).str.strip()
    out["Mês"] = out["Início"].values.astype("datetime64[M]")
    out["Macroserviço adotado"] = out.apply(classify_activity, axis=1)
    return out


def extract_budget(budget_file) -> pd.DataFrame:
    sheets = _excel_sheet_names(budget_file)
    sheet = None
    for s in sheets:
        if _norm(s) in {"orcamento", "orçamento", "budget"}:
            sheet = s
            break
    if sheet is None:
        # se não houver ORÇAMENTO, usa primeira aba que tenha código/descrição/total
        sheet = sheets[0]

    raw = _read_excel_any(budget_file, sheet_name=sheet, header=None)
    header_row = _find_header_row(raw, [["descri"], ["total"]])
    data = raw.iloc[header_row + 1 :].copy()
    data.columns = [str(x).strip() if pd.notna(x) else f"col_{i}" for i, x in enumerate(raw.iloc[header_row].tolist())]
    data = data.dropna(how="all")

    cols = list(data.columns)
    col_code = _find_col(cols, ["Cód. Estruturado", "Cod. Estruturado", "Código", "Codigo", "Cód", "Cod"])
    col_desc = _find_col(cols, ["Descrição", "Descricao", "Serviço", "Servico", "Item"])
    col_total = _find_col(cols, ["Total", "Valor", "Valor Orçado", "Valor Orcado", "Preço Total", "Preco Total"])
    if col_desc is None or col_total is None:
        raise ValueError("Não consegui identificar descrição e total no orçamento.")

    # Linhas de nível macro: 01.01, 02.07 etc.
    df = data[[col_code, col_desc, col_total]].copy() if col_code else data[[col_desc, col_total]].copy()
    if col_code:
        df.columns = ["Código", "Descrição", "Valor"]
        df["Código"] = df["Código"].astype(str).str.strip()
    else:
        df.columns = ["Descrição", "Valor"]
        df["Código"] = ""
    df["Valor"] = pd.to_numeric(df["Valor"], errors="coerce").fillna(0.0)
    df["Descrição"] = df["Descrição"].fillna("").astype(str).str.strip()
    level2 = df[df["Código"].str.match(r"^\d{2}\.\d{2}$", na=False)].copy()

    # Se a estrutura não tiver código 01.01, tenta usar linhas com valor e descrição como macros.
    if level2.empty:
        level2 = df[(df["Descrição"] != "") & (df["Valor"] > 0)].copy()

    macros = []
    used_codes = set()
    for rule in MACRO_RULES:
        value = 0.0
        origins = []
        for code in rule.codes:
            hit = level2[level2["Código"].astype(str) == code]
            if not hit.empty:
                value += float(hit["Valor"].sum())
                origins.append(code)
                used_codes.add(code)
        # fallback por descrição, se código não existir
        if value == 0:
            for _, row in level2.iterrows():
                desc = _norm(row["Descrição"])
                if any(_norm(k) in desc for k in rule.keywords):
                    value += float(row["Valor"])
                    origins.append(str(row["Código"]) or row["Descrição"])
                    used_codes.add(str(row["Código"]))
        macros.append({
            "Macroserviço": rule.name,
            "Valor Orçado": value,
            "Critério usado": rule.criterio,
            "Código orçamento / origem": "+".join(origins) if origins else "+".join(rule.codes),
            "fallback": rule.fallback,
        })

    # Qualquer macro do orçamento não capturada entra como "Outros"
    leftovers = level2[~level2["Código"].astype(str).isin(used_codes)].copy()
    if not leftovers.empty:
        val = float(leftovers["Valor"].sum())
        if val > 0:
            macros.append({
                "Macroserviço": "Outros itens do orçamento",
                "Valor Orçado": val,
                "Critério usado": "Linear no prazo total da obra",
                "Código orçamento / origem": "+".join(leftovers["Código"].astype(str).tolist()),
                "fallback": "linear",
            })

    out = pd.DataFrame(macros)
    return out


def classify_activity(row) -> str:
    text = _norm(" ".join([str(row.get("Atividade", "")), str(row.get("Serviço", "")), str(row.get("Lote", ""))]))

    # Regras específicas antes das genéricas
    priority = [
        "Implantação do canteiro", "Contenção", "Preparo do terreno", "Fundação/blocos/baldrames",
        "Piso subsolo", "Superestrutura", "Alvenaria/encunhamento", "Impermeabilização",
        "Fachada", "Instalações elétricas", "Instalações hidráulicas/sanitárias", "Instalações incêndio",
        "Elevadores", "Forros e molduras", "Revestimento parede/teto", "Revestimento piso",
        "Esquadrias alumínio", "Esquadrias madeira/portas", "Serralheria", "Pintura", "Cobertura",
        "Granito/bancadas", "Louças e metais", "Periferia/pisos externos", "Serviços finais", "Desmobilização",
    ]
    rules = sorted(MACRO_RULES, key=lambda r: priority.index(r.name) if r.name in priority else 999)
    for rule in rules:
        if any(_norm(k) in text for k in rule.keywords):
            return rule.name
    return "Não classificado"


def month_range(start: pd.Timestamp, end: pd.Timestamp) -> pd.DatetimeIndex:
    start_m = pd.Timestamp(start).replace(day=1)
    end_m = pd.Timestamp(end).replace(day=1)
    return pd.date_range(start_m, end_m, freq="MS")


def allocate_cashflow(macros: pd.DataFrame, activities: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    start = activities["Início"].min()
    end = activities["Fim"].max()
    months = month_range(start, end)

    flux = pd.DataFrame({"Mês": months})
    macro_rows = []

    total_activities = len(activities)
    for _, macro in macros.iterrows():
        name = macro["Macroserviço"]
        value = float(macro["Valor Orçado"] or 0)
        criterio = macro["Critério usado"]
        fallback = macro.get("fallback", "linear")
        col = np.zeros(len(months), dtype=float)

        acts = activities[activities["Macroserviço adotado"] == name].copy()
        n = len(acts)
        if value <= 0:
            pass
        elif "Linear no prazo total" in criterio or name in ["Serviços técnicos", "Segurança/proteções", "Operação do canteiro", "Administração/despesas gerais", "Outros itens do orçamento"]:
            col[:] = value / len(months)
        elif n > 0 and ("Rateio" in criterio or "evento" in criterio.lower()):
            value_per = value / n
            for m, count in acts.groupby("Mês").size().items():
                idx = list(months).index(pd.Timestamp(m)) if pd.Timestamp(m) in months else None
                if idx is not None:
                    col[idx] += value_per * int(count)
        else:
            # fallback técnico quando não há atividade correspondente
            if fallback == "first":
                col[0] = value
            elif fallback == "last":
                col[-1] = value
            else:
                col[:] = value / len(months)

        # Correção de arredondamento para fechar exatamente o valor do macro
        diff = value - float(col.sum())
        if abs(diff) > 0.01:
            # joga diferença no mês com maior valor ou último mês
            ix = int(np.argmax(col)) if col.sum() > 0 else len(col) - 1
            col[ix] += diff

        flux[name] = col
        macro_rows.append({
            "Macroserviço": name,
            "Valor Orçado": value,
            "Critério usado": criterio,
            "Nº frentes/atividades": int(n),
            "Valor médio por frente": (value / n if n > 0 else np.nan),
            "Código orçamento / origem": macro["Código orçamento / origem"],
        })

    macro_base = pd.DataFrame(macro_rows)
    service_cols = [c for c in flux.columns if c != "Mês"]
    flux["Total Mensal"] = flux[service_cols].sum(axis=1)
    flux["Acumulado"] = flux["Total Mensal"].cumsum()
    budget_total = float(macro_base["Valor Orçado"].sum())
    flux["% Física"] = flux["Acumulado"] / budget_total if budget_total else 0
    flux["% Financeira"] = flux["Acumulado"] / budget_total if budget_total else 0
    return flux, macro_base


def make_excel(budget_file, schedule_file) -> bytes:
    activities = extract_schedule(schedule_file)
    macros = extract_budget(budget_file)
    flux, macro_base = allocate_cashflow(macros, activities)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="dd/mm/yyyy", date_format="dd/mm/yyyy") as writer:
        workbook = writer.book
        # Formats
        fmt_title = workbook.add_format({"bold": True, "font_size": 18, "font_color": "#FFFFFF", "bg_color": "#1F4E78", "align": "center", "valign": "vcenter"})
        fmt_header = workbook.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#1F4E78", "border": 1, "align": "center", "valign": "vcenter"})
        fmt_subheader = workbook.add_format({"bold": True, "font_color": "#1F4E78", "bg_color": "#D9EAF7", "border": 1})
        fmt_money = workbook.add_format({"num_format": 'R$ #,##0.00', "border": 1})
        fmt_money_bold = workbook.add_format({"num_format": 'R$ #,##0.00', "bold": True, "border": 1, "bg_color": "#E2F0D9"})
        fmt_pct = workbook.add_format({"num_format": "0.00%", "border": 1})
        fmt_date = workbook.add_format({"num_format": "mmm/yy", "border": 1})
        fmt_date_full = workbook.add_format({"num_format": "dd/mm/yyyy", "border": 1})
        fmt_text = workbook.add_format({"border": 1, "valign": "top"})
        fmt_note = workbook.add_format({"text_wrap": True, "valign": "top", "border": 1})
        fmt_center = workbook.add_format({"border": 1, "align": "center"})

        # ---------------- Dashboard ----------------
        ws = workbook.add_worksheet("Dashboard")
        writer.sheets["Dashboard"] = ws
        ws.merge_range("A1:H1", "Fluxo de Caixa Mensal — Obra", fmt_title)
        ws.set_row(0, 30)
        ws.set_column("A:A", 28)
        ws.set_column("B:B", 18)
        ws.set_column("D:D", 10)
        ws.set_column("E:E", 16)
        ws.set_column("F:F", 18)
        ws.write("A3", "Orçamento total", fmt_subheader)
        ws.write_number("B3", float(macro_base["Valor Orçado"].sum()), fmt_money_bold)
        ws.write("A4", "Início do cronograma", fmt_subheader)
        ws.write_datetime("B4", pd.Timestamp(activities["Início"].min()).to_pydatetime(), fmt_date_full)
        ws.write("A5", "Fim do cronograma", fmt_subheader)
        ws.write_datetime("B5", pd.Timestamp(activities["Fim"].max()).to_pydatetime(), fmt_date_full)
        ws.write("A6", "Prazo analisado (meses)", fmt_subheader)
        ws.write_number("B6", len(flux), fmt_center)
        ws.write("A7", "Pico mensal de caixa", fmt_subheader)
        ws.write_number("B7", float(flux["Total Mensal"].max()), fmt_money_bold)
        ws.write("A8", "Mês do pico", fmt_subheader)
        ws.write_datetime("B8", pd.Timestamp(flux.loc[flux["Total Mensal"].idxmax(), "Mês"]).to_pydatetime(), fmt_date)
        ws.write("A9", "Quantidade de atividades", fmt_subheader)
        ws.write_number("B9", int(len(activities)), fmt_center)
        ws.write("A10", "Macroserviços", fmt_subheader)
        ws.write_number("B10", int(len(macro_base)), fmt_center)

        ws.write_row("D3", ["Ranking", "Mês", "Desembolso"], fmt_header)
        ranking = flux[["Mês", "Total Mensal"]].sort_values("Total Mensal", ascending=False).head(10).reset_index(drop=True)
        for i, row in ranking.iterrows():
            ws.write_number(3 + i, 3, i + 1, fmt_center)
            ws.write_datetime(3 + i, 4, pd.Timestamp(row["Mês"]).to_pydatetime(), fmt_date)
            ws.write_number(3 + i, 5, float(row["Total Mensal"]), fmt_money)

        ws.write("A13", "Análise automática", fmt_header)
        analysis_lines = []
        pico = float(flux["Total Mensal"].max())
        media = float(flux["Total Mensal"].mean())
        if pico > media * 1.8:
            analysis_lines.append("Há concentração relevante de desembolso no mês de pico.")
        if (flux["Total Mensal"] < media * 0.25).sum() > 0:
            analysis_lines.append("Existem meses com baixo desembolso relativo; verifique paralisações ou lacunas de cronograma.")
        if not analysis_lines:
            analysis_lines.append("O fluxo mensal está relativamente distribuído, considerando as premissas adotadas.")
        ws.merge_range("A14:F16", "\n".join(analysis_lines), fmt_note)

        # Data auxiliar para gráficos no Dashboard
        aux_start_row = 20
        ws.write_row(aux_start_row, 0, ["Mês", "Total Mensal", "Acumulado", "% Financeira"], fmt_header)
        for i, row in flux.iterrows():
            r = aux_start_row + 1 + i
            ws.write_datetime(r, 0, pd.Timestamp(row["Mês"]).to_pydatetime(), fmt_date)
            ws.write_number(r, 1, float(row["Total Mensal"]), fmt_money)
            ws.write_number(r, 2, float(row["Acumulado"]), fmt_money)
            ws.write_number(r, 3, float(row["% Financeira"]), fmt_pct)
        chart1 = workbook.add_chart({"type": "column"})
        chart1.add_series({"name": "Desembolso mensal", "categories": ["Dashboard", aux_start_row + 1, 0, aux_start_row + len(flux), 0], "values": ["Dashboard", aux_start_row + 1, 1, aux_start_row + len(flux), 1]})
        chart1.set_title({"name": "Desembolso Mensal"})
        chart1.set_legend({"none": True})
        chart1.set_y_axis({"num_format": 'R$ #,##0'})
        ws.insert_chart("A18", chart1, {"x_scale": 1.35, "y_scale": 1.05})

        chart2 = workbook.add_chart({"type": "line"})
        chart2.add_series({"name": "Curva S Financeira", "categories": ["Dashboard", aux_start_row + 1, 0, aux_start_row + len(flux), 0], "values": ["Dashboard", aux_start_row + 1, 3, aux_start_row + len(flux), 3]})
        chart2.set_title({"name": "Curva S Financeira"})
        chart2.set_y_axis({"num_format": "0%"})
        chart2.set_legend({"none": True})
        ws.insert_chart("E18", chart2, {"x_scale": 1.2, "y_scale": 1.05})

        # ---------------- Fluxo Mensal ----------------
        flux_to_write = flux.copy()
        flux_to_write.to_excel(writer, sheet_name="Fluxo Mensal", index=False)
        ws2 = writer.sheets["Fluxo Mensal"]
        ws2.freeze_panes(1, 1)
        ws2.set_column(0, 0, 13)
        ws2.set_column(1, len(flux_to_write.columns) - 1, 18)
        for c, col in enumerate(flux_to_write.columns):
            ws2.write(0, c, col, fmt_header)
        for r in range(1, len(flux_to_write) + 1):
            ws2.write_datetime(r, 0, pd.Timestamp(flux_to_write.iloc[r - 1, 0]).to_pydatetime(), fmt_date)
        total_idx = flux_to_write.columns.get_loc("Total Mensal")
        acum_idx = flux_to_write.columns.get_loc("Acumulado")
        fis_idx = flux_to_write.columns.get_loc("% Física")
        fin_idx = flux_to_write.columns.get_loc("% Financeira")
        ws2.set_column(1, acum_idx, 16, fmt_money)
        ws2.set_column(fis_idx, fin_idx, 14, fmt_pct)
        # Formata corpo monetário
        for row in range(1, len(flux_to_write) + 1):
            for col in range(1, acum_idx + 1):
                val = flux_to_write.iloc[row - 1, col]
                ws2.write_number(row, col, float(val), fmt_money)
            ws2.write_number(row, fis_idx, float(flux_to_write.iloc[row - 1, fis_idx]), fmt_pct)
            ws2.write_number(row, fin_idx, float(flux_to_write.iloc[row - 1, fin_idx]), fmt_pct)
        ws2.autofilter(0, 0, len(flux_to_write), len(flux_to_write.columns) - 1)

        # ---------------- Macroserviços ----------------
        macro_write = macro_base[["Macroserviço", "Valor Orçado", "Critério usado", "Nº frentes/atividades", "Valor médio por frente", "Código orçamento / origem"]].copy()
        macro_write.to_excel(writer, sheet_name="Macroserviços", index=False, startrow=0, startcol=0)
        ws3 = writer.sheets["Macroserviços"]
        ws3.freeze_panes(1, 0)
        ws3.set_column("A:A", 34)
        ws3.set_column("B:B", 18)
        ws3.set_column("C:C", 42)
        ws3.set_column("D:D", 18)
        ws3.set_column("E:E", 20)
        ws3.set_column("F:F", 24)
        for c, col in enumerate(macro_write.columns):
            ws3.write(0, c, col, fmt_header)
        for r in range(1, len(macro_write) + 1):
            ws3.write(r, 0, macro_write.iloc[r - 1, 0], fmt_text)
            ws3.write_number(r, 1, float(macro_write.iloc[r - 1, 1]), fmt_money)
            ws3.write(r, 2, macro_write.iloc[r - 1, 2], fmt_text)
            ws3.write_number(r, 3, int(macro_write.iloc[r - 1, 3]), fmt_center)
            v = macro_write.iloc[r - 1, 4]
            if pd.notna(v):
                ws3.write_number(r, 4, float(v), fmt_money)
            else:
                ws3.write_blank(r, 4, None, fmt_money)
            ws3.write(r, 5, macro_write.iloc[r - 1, 5], fmt_text)
        ws3.write("H1", "Resumo", fmt_header)
        ws3.write("H2", "Total orçamento", fmt_subheader); ws3.write_number("I2", float(macro_base["Valor Orçado"].sum()), fmt_money_bold)
        ws3.write("H3", "Macros cadastrados", fmt_subheader); ws3.write_number("I3", int(len(macro_base)), fmt_center)
        ws3.write("H4", "Atividades cronograma", fmt_subheader); ws3.write_number("I4", int(len(activities)), fmt_center)
        ws3.write("H5", "Não classificadas", fmt_subheader); ws3.write_number("I5", int((activities["Macroserviço adotado"] == "Não classificado").sum()), fmt_center)

        # ---------------- Atividades Base ----------------
        act_write = activities[["ID", "Atividade", "Serviço", "Lote", "Início", "Fim", "Macroserviço adotado"]].copy()
        act_write.to_excel(writer, sheet_name="Atividades Base", index=False)
        ws4 = writer.sheets["Atividades Base"]
        ws4.freeze_panes(1, 0)
        ws4.set_column("A:A", 12)
        ws4.set_column("B:B", 38)
        ws4.set_column("C:D", 18)
        ws4.set_column("E:F", 14)
        ws4.set_column("G:G", 28)
        for c, col in enumerate(act_write.columns):
            ws4.write(0, c, col, fmt_header)
        for r in range(1, len(act_write) + 1):
            ws4.write(r, 0, act_write.iloc[r - 1, 0], fmt_text)
            ws4.write(r, 1, act_write.iloc[r - 1, 1], fmt_text)
            ws4.write(r, 2, act_write.iloc[r - 1, 2], fmt_text)
            ws4.write(r, 3, act_write.iloc[r - 1, 3], fmt_text)
            ws4.write_datetime(r, 4, pd.Timestamp(act_write.iloc[r - 1, 4]).to_pydatetime(), fmt_date_full)
            ws4.write_datetime(r, 5, pd.Timestamp(act_write.iloc[r - 1, 5]).to_pydatetime(), fmt_date_full)
            ws4.write(r, 6, act_write.iloc[r - 1, 6], fmt_text)
        ws4.autofilter(0, 0, len(act_write), len(act_write.columns) - 1)

        # ---------------- Premissas ----------------
        ws5 = workbook.add_worksheet("Premissas")
        writer.sheets["Premissas"] = ws5
        ws5.set_column("A:A", 28)
        ws5.set_column("B:B", 110)
        ws5.write_row("A1", ["Item", "Premissa adotada"], fmt_header)
        for i, row in enumerate(PREMISSAS, start=1):
            ws5.write(i, 0, row[0], fmt_text)
            ws5.write(i, 1, row[1], fmt_note)
        ws5.freeze_panes(1, 0)

    output.seek(0)
    return output.getvalue()
