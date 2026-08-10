# -*- coding: utf-8 -*-
"""
Force Plate Test Report — app Streamlit (versione a file singolo)

Carica gli export XLSX di ForceMate/ForceDecks (IMTP, SJ, CMJ, CMJ RE),
calcola automaticamente medie, T-score rispetto alla popolazione di
riferimento e produce un profilo di forza con grafici Plotly e report Word.

Struttura del file:
  PARTE 1 — Costanti e dati di popolazione
  PARTE 2 — Caricamento file (sidebar upload)
  PARTE 3 — Lettura/parsing dei file XLSX
  PARTE 4 — Analisi dati e confronto con la popolazione
  PARTE 5 — Report live (UI a schede)
  PARTE 6 — Report scaricabile (Word)
"""

from __future__ import annotations

import io
import math
import datetime as dt
from dataclasses import dataclass, field

import streamlit as st
import plotly.graph_objects as go
import openpyxl
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH


# ============================================================================
# PARTE 1 — COSTANTI E DATI DI POPOLAZIONE
# ============================================================================
# Equivalente ai fogli "DATI POP" (norme di popolazione) e "DATI TEST"
# (definizione delle metriche) del Google Sheet originale.

CATEGORIES = ["FORZA MAX", "POTENZA", "ESPLOSIVITA'", "REATTIVITA'"]

# Bande di valutazione basate sul T-score (media 50, deviazione standard 10)
BANDS = [
    (float("-inf"), 30, "MOLTO SOTTO LA MEDIA", "#B00020"),
    (30, 40, "SOTTO LA MEDIA", "#E4572E"),
    (40, 60, "NELLA MEDIA", "#F2A007"),
    (60, 70, "SOPRA LA MEDIA", "#8BC34A"),
    (70, 80, "BUONO", "#4CAF50"),
    (80, float("inf"), "OTTIMO", "#2E7D32"),
]


def banda_da_tscore(t):
    if t is None:
        return None, None
    for lo, hi, label, color in BANDS:
        if lo <= t < hi:
            return label, color
    return "OTTIMO", "#2E7D32"


# sd sempre positiva: la direzione "minore è meglio" è gestita dal flag
# lower_is_better nella metrica, non dal segno della deviazione standard.
POP = {
    "imtp_peak_force":        dict(mean_m=2606.8,      sd_m=646.1163194,  mean_f=1575.0,      sd_f=386.145544),
    "imtp_rel_peak_force":    dict(mean_m=34.56,        sd_m=5.27,         mean_f=34.56,        sd_f=5.27),
    "sj_mean_power":          dict(mean_m=1180.0,       sd_m=414.1638849,  mean_f=823.0,        sd_f=240.2447942),
    "sj_height":              dict(mean_m=26.9,         sd_m=6.57,         mean_f=19.35,        sd_f=5.51),
    "sj_contraction_time":    dict(mean_m=0.445,        sd_m=0.127835797,  mean_f=0.46,         sd_f=0.137032268),
    "cmj_height":             dict(mean_m=30.01,        sd_m=6.5,          mean_f=20.91,        sd_f=5.84),
    "mrsi_cmj":               dict(mean_m=0.419,        sd_m=0.098531539,  mean_f=0.308,        sd_f=0.093831019),
    "dsi":                    dict(mean_m=0.7,          sd_m=0.074074074,  mean_f=0.7,          sd_f=0.074074074),
    "eur":                    dict(mean_m=0.107708605,  sd_m=0.038034873,  mean_f=0.090563116,  sd_f=0.02656191),
    "cmj_re_rebound_height":  dict(mean_m=38.5,         sd_m=5.5996817,    mean_f=38.5,         sd_f=5.5996817),
    "cmj_re_contact_time":    dict(mean_m=0.25,         sd_m=0.148005087,  mean_f=0.25,         sd_f=0.148005087),
    "cmj_re_rebound_impulse": dict(mean_m=541.5,        sd_m=53.82161458,  mean_f=541.5,        sd_f=53.82161458),
    "mrsi_cmj_re":            dict(mean_m=1.37,         sd_m=0.331705729,  mean_f=1.37,         sd_f=0.331705729),
}


def _safe_div(a, b):
    try:
        if a is None or b in (None, 0):
            return None
        return float(a) / float(b)
    except (TypeError, ValueError):
        return None


# Definizione metriche: key, label, category, jump_type (o None se derivata),
# raw_var (nome variabile grezza ForceMate) o derive (funzione per-rep),
# unit, pop_key (chiave in POP o None), lower_is_better, kind
#   kind: "score" (T-score da distribuzione pooled), "score_single"
#         (T-score da singolo valore aggregato, es. DSI/EUR), "info"
#         (valore mostrato ma senza T-score)
METRICS = [
    dict(key="imtp_peak_force", label="IMTP Peak Force", category="FORZA MAX",
         jump_type="imtp", raw_var="peak force", unit="N",
         pop_key="imtp_peak_force", lower_is_better=False, kind="score"),
    dict(key="imtp_rel_peak_force", label="IMTP Rel Peak Force", category="FORZA MAX",
         jump_type="imtp", raw_var=None,
         derive=lambda rep: _safe_div(rep.get("peak force"), rep.get("body mass")),
         unit="N/kg", pop_key="imtp_rel_peak_force", lower_is_better=False, kind="score"),

    dict(key="sj_mean_power", label="SJ Mean Power", category="POTENZA",
         jump_type="sj", raw_var="avg. propulsive power", unit="W",
         pop_key="sj_mean_power", lower_is_better=False, kind="score"),
    dict(key="sj_height", label="SJ Height", category="POTENZA",
         jump_type="sj", raw_var="jump height ft", unit="cm",
         pop_key="sj_height", lower_is_better=False, kind="score"),
    dict(key="sj_contraction_time", label="SJ Contraction Time", category="POTENZA",
         jump_type="sj", raw_var="time to takeoff", unit="s",
         pop_key="sj_contraction_time", lower_is_better=True, kind="score"),
    dict(key="sj_net_impulse", label="SJ Net Impulse", category="POTENZA",
         jump_type="sj", raw_var="net impulse", unit="N\u00b7s",
         pop_key=None, lower_is_better=False, kind="info"),
    dict(key="sj_net_rel_impulse", label="SJ Net Rel Impulse", category="POTENZA",
         jump_type="sj", raw_var=None,
         derive=lambda rep: _safe_div(rep.get("net impulse"), rep.get("body mass")),
         unit="N\u00b7s/kg", pop_key=None, lower_is_better=False, kind="info"),

    dict(key="cmj_net_impulse", label="CMJ Net Impulse", category="ESPLOSIVITA'",
         jump_type="cmj", raw_var="net impulse", unit="N\u00b7s",
         pop_key=None, lower_is_better=False, kind="info"),
    dict(key="cmj_net_rel_impulse", label="CMJ Net Rel Impulse", category="ESPLOSIVITA'",
         jump_type="cmj", raw_var=None,
         derive=lambda rep: _safe_div(rep.get("net impulse"), rep.get("body mass")),
         unit="N\u00b7s/kg", pop_key=None, lower_is_better=False, kind="info"),
    dict(key="cmj_contraction_time", label="CMJ Contraction Time", category="ESPLOSIVITA'",
         jump_type="cmj", raw_var="time to takeoff", unit="s",
         pop_key=None, lower_is_better=True, kind="info"),
    dict(key="cmj_height", label="CMJ Height", category="ESPLOSIVITA'",
         jump_type="cmj", raw_var="jump height ft", unit="cm",
         pop_key="cmj_height", lower_is_better=False, kind="score"),
    dict(key="mrsi_cmj", label="mRSI-CMJ", category="ESPLOSIVITA'",
         jump_type="cmj", raw_var="rsi modified", unit="m/s",
         pop_key="mrsi_cmj", lower_is_better=False, kind="score"),
    dict(key="cmj_peak_force", label="CMJ Peak Force", category="ESPLOSIVITA'",
         jump_type="cmj", raw_var="peak propulsive force", unit="N",
         pop_key=None, lower_is_better=False, kind="info"),

    dict(key="cmj_re_initial_height", label="CMJ RE Jump Height (iniziale)", category="REATTIVITA'",
         jump_type="cmrj", raw_var="jump height ft", unit="cm",
         pop_key=None, lower_is_better=False, kind="info"),
    dict(key="cmj_re_rebound_height", label="CMJ RE Rebound Jump Height", category="REATTIVITA'",
         jump_type="cmrj", raw_var="rebound jump height ft", unit="cm",
         pop_key="cmj_re_rebound_height", lower_is_better=False, kind="score"),
    dict(key="cmj_re_contact_time", label="CMJ RE Contact Time", category="REATTIVITA'",
         jump_type="cmrj", raw_var="rebound contact time", unit="s",
         pop_key="cmj_re_contact_time", lower_is_better=True, kind="score"),
    dict(key="cmj_re_rebound_impulse", label="CMJ RE Rebound Propulsive Impulse", category="REATTIVITA'",
         jump_type="cmrj", raw_var="rebound propulsive impulse", unit="N\u00b7s",
         pop_key="cmj_re_rebound_impulse", lower_is_better=False, kind="score"),
    dict(key="mrsi_cmj_re", label="mRSI-CMJ RE", category="REATTIVITA'",
         jump_type="cmrj", raw_var="rebound rsi modified", unit="m/s",
         pop_key="mrsi_cmj_re", lower_is_better=False, kind="score"),
    dict(key="unbalanced_landing_raw", label="Braking Impulse Sym. Index (CMJ RE)", category="REATTIVITA'",
         jump_type="cmrj", raw_var="braking impulse sym. index", unit="%",
         pop_key=None, lower_is_better=False, kind="info"),

    # Indici derivati da medie aggregate cross-test (non per-rep)
    dict(key="dsi", label="DSI (Dynamic Strength Index)", category="INDICI",
         jump_type=None, raw_var=None, unit="",
         pop_key="dsi", lower_is_better=False, kind="score_single"),
    dict(key="eur", label="EUR (Eccentric Utilisation Ratio)", category="INDICI",
         jump_type=None, raw_var=None, unit="",
         pop_key="eur", lower_is_better=False, kind="score_single"),
]

# Controlli a soglia (equivalenti a "Jump to rebound Ratio", "CMJ to CMJ RE
# check", "Unbalanced landing check" del foglio originale)
CHECKS = [
    dict(key="jump_to_rebound_ratio", label="Jump to Rebound Ratio",
         desc="Rapporto tra altezza del rimbalzo e altezza del salto iniziale nel CMJ RE.",
         threshold=0.60, direction="min"),
    dict(key="cmj_to_cmjre_check", label="CMJ to CMJ RE Check",
         desc="Rapporto tra l'altezza del salto iniziale nel CMJ RE e l'altezza del CMJ standard.",
         threshold=0.85, direction="min"),
    dict(key="unbalanced_landing_check", label="Unbalanced Landing Check",
         desc="Indice di simmetria dell'impulso frenante in atterraggio (CMJ RE). Valori assoluti alti indicano un atterraggio sbilanciato.",
         threshold=0.50, direction="max"),
]

JUMP_TYPE_LABELS = {"imtp": "IMTP", "sj": "Squat Jump", "cmj": "CMJ", "cmrj": "CMJ Rebound"}


# ============================================================================
# PARTE 2 — CARICAMENTO FILE (sidebar upload)
# ============================================================================

st.set_page_config(page_title="Force Plate Test Report", page_icon="🏋️", layout="wide")
PRIMARY = "#1f77b4"

st.sidebar.title("📂 Import dati")
uploaded = st.sidebar.file_uploader(
    "Carica i file XLSX esportati da ForceMate", type=["xlsx"], accept_multiple_files=True
)

if not uploaded:
    st.title("🏋️ Force Plate Test Report")
    st.info("Carica uno o più file XLSX esportati da ForceMate (IMTP, SJ, CMJ, CMJ RE) dalla barra laterale per iniziare.")
    st.stop()


# ============================================================================
# PARTE 3 — LETTURA / PARSING DEI FILE XLSX
# ============================================================================
# Ogni file ForceMate ha: colonne A-F = metadati (generali, attributi
# atleta, campi custom); colonna G = unità di misura; blocchi da 3 colonne
# ("Jump 1", "Jump 2", ...) con [variabile, valore, spacer] per ripetizione.
# Il "jump type" è taggato per singola ripetizione, non per file (es. un
# file SJ può contenere per errore una rep taggata "cmj").

@dataclass
class ParsedFile:
    filename: str
    metadata: dict
    reps: list = field(default_factory=list)  # [{"jump_type","vars","units"}, ...]


def _parse_date(raw):
    if raw is None:
        return None
    if isinstance(raw, dt.datetime):
        return raw
    s = str(raw).strip()
    for fmt in ("%d.%m.%Y %H:%M:%S", "%d.%m.%Y", "%Y-%m-%d"):
        try:
            return dt.datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def parse_forcemate_workbook(file_like, filename: str) -> ParsedFile:
    wb = openpyxl.load_workbook(file_like, data_only=True, read_only=False)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return ParsedFile(filename=filename, metadata={}, reps=[])

    max_col = max(len(r) for r in rows)

    def cell(r_idx, c_idx):
        row = rows[r_idx - 1] if 0 < r_idx <= len(rows) else ()
        return row[c_idx - 1] if 0 < c_idx <= len(row) else None

    def kv_lookup(col_key, col_val, max_r=15):
        out = {}
        for r in range(1, max_r + 1):
            k = cell(r, col_key)
            if k:
                out[str(k).strip().lower()] = cell(r, col_val)
        return out

    generali = kv_lookup(1, 2)
    attributi = kv_lookup(3, 4)
    custom = kv_lookup(5, 6)

    metadata = {
        "nome": attributi.get("name") or generali.get("name") or "Atleta",
        "sesso": (attributi.get("gender") or "").strip().upper() if attributi.get("gender") else None,
        "altezza_cm": attributi.get("height"),
        "peso_kg_input": attributi.get("weight"),
        "data_test": _parse_date(generali.get("date")),
        "device": generali.get("device"),
        "team": generali.get("team"),
        "test_period": custom.get("test period"),
        "test_type": custom.get("test type"),
    }

    block_starts = [c for c in range(8, max_col + 1)
                    if cell(1, c) and str(cell(1, c)).strip().lower().startswith("jump")]

    reps = []
    for bs in block_starts:
        var_col, val_col, unit_col = bs, bs + 1, 7
        variables, units, jump_type = {}, {}, None
        for r in range(2, len(rows) + 1):
            var_name = cell(r, var_col)
            if not var_name:
                continue
            var_name = str(var_name).strip().lower()
            val = cell(r, val_col)
            unit = cell(r, unit_col)
            if var_name == "jump type":
                jump_type = str(val).strip().lower() if val else None
                continue
            if var_name == "tags":
                continue
            num = None
            if isinstance(val, (int, float)):
                num = float(val)
            elif isinstance(val, str):
                try:
                    num = float(val.replace(",", "."))
                except ValueError:
                    num = None
            variables[var_name] = num if num is not None else val
            if unit:
                units[var_name] = str(unit).strip()
        reps.append({"jump_type": jump_type, "vars": variables, "units": units})

    return ParsedFile(filename=filename, metadata=metadata, reps=reps)


def apply_type_override(parsed: ParsedFile, override_type):
    if not override_type:
        return parsed
    for rep in parsed.reps:
        if not rep["jump_type"]:
            rep["jump_type"] = override_type
    return parsed


def guess_type_from_filename(filename: str):
    fn = filename.upper()
    for t in ("CMRJ", "IMTP", "CMJ", "SJ", "DJ"):
        if t in fn:
            return t.lower()
    return None


parsed_files = []
st.sidebar.markdown("---")
st.sidebar.subheader("Tipo di esercizio per file")
st.sidebar.caption("Verificare soprattutto l'IMTP: spesso non viene taggato automaticamente dal software.")

for f in uploaded:
    pf = parse_forcemate_workbook(io.BytesIO(f.getvalue()), f.name)
    guess = guess_type_from_filename(f.name)
    detected_types = {r["jump_type"] for r in pf.reps if r["jump_type"]}
    default = guess or (list(detected_types)[0] if len(detected_types) == 1 else None)
    options = ["imtp", "sj", "cmj", "cmrj"]
    idx = options.index(default) if default in options else 0
    choice = st.sidebar.selectbox(
        f.name, options, index=idx, format_func=lambda x: JUMP_TYPE_LABELS.get(x, x),
        key=f"type_{f.name}",
        help="Applicato solo alle ripetizioni prive di tag automatico nel file."
    )
    parsed_files.append(apply_type_override(pf, choice))


# ============================================================================
# PARTE 4 — ANALISI DATI E CONFRONTO CON LA POPOLAZIONE
# ============================================================================
# Equivalente alle funzioni MEDIA_DINAMICA_JUMP / DEV_STD_DINAMICA_JUMP /
# CONTA_DINAMICA_JUMP dello script Apps Script originale: i valori vengono
# aggregati per tipo di test attraverso TUTTI i file caricati, poi
# confrontati con le norme di popolazione tramite Z-score / T-score.

def pooled_stats(values):
    values = [v for v in values if isinstance(v, (int, float)) and not math.isnan(v)]
    n = len(values)
    if n == 0:
        return dict(n=0, mean=None, sd=None)
    mean = sum(values) / n
    if n < 2:
        return dict(n=n, mean=mean, sd=None)
    var = sum((v - mean) ** 2 for v in values) / n  # ddof=0, come nello script originale
    return dict(n=n, mean=mean, sd=math.sqrt(var))


def z_t_score(value, pop_mean, pop_sd, lower_is_better):
    if value is None or pop_mean is None or not pop_sd:
        return None, None
    z = (value - pop_mean) / abs(pop_sd)
    if lower_is_better:
        z = -z
    return z, 50 + 10 * z


def collect_reps(files, jump_type):
    return [rep for pf in files for rep in pf.reps if rep["jump_type"] == jump_type]


def compute_metric_values(files, metric):
    reps = collect_reps(files, metric["jump_type"])
    values = []
    for rep in reps:
        if metric.get("raw_var"):
            v = rep["vars"].get(metric["raw_var"])
        elif metric.get("derive"):
            v = metric["derive"](rep["vars"])
        else:
            v = None
        if isinstance(v, (int, float)):
            values.append(v)
    return values


def sesso_da_file(files):
    for pf in files:
        if pf.metadata.get("sesso"):
            return pf.metadata["sesso"]
    return "UOMO"


def build_results(files):
    sesso_raw = sesso_da_file(files)
    is_female = sesso_raw and sesso_raw.upper().startswith("F")

    support = {}
    results = []

    for metric in METRICS:
        if metric["kind"] in ("score", "info") and metric.get("jump_type"):
            values = compute_metric_values(files, metric)
            stats = pooled_stats(values)
            support[metric["key"]] = stats

            z = t = banda = colore = None
            pop_mean = pop_sd = None
            if metric.get("pop_key"):
                pop = POP[metric["pop_key"]]
                pop_mean = pop["mean_f"] if is_female else pop["mean_m"]
                pop_sd = pop["sd_f"] if is_female else pop["sd_m"]
                if metric["kind"] == "score":
                    z, t = z_t_score(stats["mean"], pop_mean, pop_sd, metric["lower_is_better"])
                    banda, colore = banda_da_tscore(t)

            results.append(dict(
                key=metric["key"], label=metric["label"], category=metric["category"],
                unit=metric["unit"], kind=metric["kind"], n=stats["n"], mean=stats["mean"], sd=stats["sd"],
                z=z, t=t, banda=banda, colore=colore, pop_mean=pop_mean, pop_sd=pop_sd,
            ))

    # --- Indici derivati da medie aggregate (DSI, EUR) ---
    cmj_peak = support.get("cmj_peak_force", {}).get("mean")
    imtp_peak = support.get("imtp_peak_force", {}).get("mean")
    dsi_val = (cmj_peak / imtp_peak) if (cmj_peak and imtp_peak) else None

    # EUR = (CMJ height - SJ height) / SJ height. Coerente con la media di
    # popolazione del foglio originale (~0,11): un rapporto puro CMJ/SJ
    # varrebbe tipicamente ~1,1 e sarebbe incompatibile con quella norma.
    cmj_h = support.get("cmj_height", {}).get("mean")
    sj_h = support.get("sj_height", {}).get("mean")
    eur_val = ((cmj_h - sj_h) / sj_h) if (cmj_h and sj_h) else None

    for key, val in (("dsi", dsi_val), ("eur", eur_val)):
        pop = POP[key]
        pop_mean = pop["mean_f"] if is_female else pop["mean_m"]
        pop_sd = pop["sd_f"] if is_female else pop["sd_m"]
        z, t = z_t_score(val, pop_mean, pop_sd, False)
        banda, colore = banda_da_tscore(t)
        metric_def = next(m for m in METRICS if m["key"] == key)
        results.append(dict(
            key=key, label=metric_def["label"], category="INDICI", unit="",
            kind="score_single", n=(1 if val is not None else 0), mean=val, sd=None,
            z=z, t=t, banda=banda, colore=colore, pop_mean=pop_mean, pop_sd=pop_sd,
        ))

    # --- Checks a soglia ---
    def make_check(defn, ratio):
        passed = None
        if ratio is not None:
            passed = (ratio >= defn["threshold"]) if defn["direction"] == "min" else (abs(ratio) <= defn["threshold"])
        return dict(**defn, value=ratio, passed=passed)

    initial_h = support.get("cmj_re_initial_height", {}).get("mean")
    rebound_h = support.get("cmj_re_rebound_height", {}).get("mean")
    landing_sym = support.get("unbalanced_landing_raw", {}).get("mean")
    cmj_h_standalone = support.get("cmj_height", {}).get("mean")

    checks_out = [
        make_check(CHECKS[0], (rebound_h / initial_h) if (initial_h and rebound_h) else None),
        make_check(CHECKS[1], (initial_h / cmj_h_standalone) if (initial_h and cmj_h_standalone) else None),
        make_check(CHECKS[2], (landing_sym / 100.0) if landing_sym is not None else None),
    ]

    return dict(results=results, support=support, checks=checks_out, sesso=sesso_raw)


def profilo_forza(results):
    out = {}
    for cat in CATEGORIES:
        ts = [r["t"] for r in results if r["category"] == cat and r["t"] is not None]
        out[cat] = (sum(ts) / len(ts)) if ts else None
    return out


results_bundle = build_results(parsed_files)
results = results_bundle["results"]
checks = results_bundle["checks"]
profilo = profilo_forza(results)

meta0 = parsed_files[0].metadata
nome = meta0.get("nome") or "Atleta"
sesso = results_bundle["sesso"] or "-"
date_tests = [pf.metadata.get("data_test") for pf in parsed_files if pf.metadata.get("data_test")]
data_min = min(date_tests).strftime("%d/%m/%Y") if date_tests else "-"
data_max = max(date_tests).strftime("%d/%m/%Y") if date_tests else "-"
periodo = data_min if data_min == data_max else f"{data_min} → {data_max}"


# ============================================================================
# PARTE 5 — REPORT LIVE (interfaccia a schede)
# ============================================================================

st.title(f"🏋️ Force Plate Test Report — {nome}")
c1, c2, c3, c4 = st.columns(4)
c1.metric("Sesso", sesso)
c2.metric("Data test", periodo)
c3.metric("File caricati", len(parsed_files))
c4.metric("Ripetizioni totali", sum(len(pf.reps) for pf in parsed_files))

tab_profilo, tab_dettaglio, tab_checks, tab_report = st.tabs(
    ["📊 Profilo di Forza", "🔍 Dettaglio Test", "✅ Controlli Tecnici", "📄 Report"]
)

with tab_profilo:
    cats_valide = [c for c in CATEGORIES if profilo.get(c) is not None]
    if not cats_valide:
        st.warning("Nessuna metrica con confronto di popolazione disponibile: carica almeno un test tra IMTP, SJ, CMJ o CMJ RE.")
    else:
        vals = [profilo[c] for c in cats_valide]
        vals_closed = vals + [vals[0]]
        cats_closed = cats_valide + [cats_valide[0]]

        fig = go.Figure()
        fig.add_trace(go.Scatterpolar(
            r=[50] * (len(cats_valide) + 1), theta=cats_closed, mode="lines",
            line=dict(color="rgba(150,150,150,0.6)", dash="dash"), name="Media popolazione (T=50)"
        ))
        fig.add_trace(go.Scatterpolar(
            r=vals_closed, theta=cats_closed, fill="toself",
            line=dict(color=PRIMARY, width=3), fillcolor="rgba(31,119,180,0.25)", name=nome
        ))
        fig.update_layout(
            polar=dict(radialaxis=dict(range=[0, 100], showticklabels=True, ticks="")),
            showlegend=True, height=520, margin=dict(t=40, b=40),
        )
        st.plotly_chart(fig, use_container_width=True)

        cols = st.columns(len(cats_valide))
        for col, cat in zip(cols, cats_valide):
            t = profilo[cat]
            banda, colore = banda_da_tscore(t)
            col.markdown(f"**{cat}**")
            col.markdown(f"<span style='font-size:28px;color:{colore}'>{t:.0f}</span>", unsafe_allow_html=True)
            col.caption(banda)

    indici = [r for r in results if r["category"] == "INDICI" and r["mean"] is not None]
    if indici:
        st.markdown("#### Indici")
        icols = st.columns(len(indici))
        for col, r in zip(icols, indici):
            col.metric(r["label"], f"{r['mean']:.3f}",
                       help=f"T-score: {r['t']:.0f} ({r['banda']})" if r["t"] else None)

with tab_dettaglio:
    for cat in CATEGORIES:
        cat_results = [r for r in results if r["category"] == cat and r["mean"] is not None]
        if not cat_results:
            continue
        st.markdown(f"### {cat}")
        cols = st.columns(2)

        scored = [r for r in cat_results if r["t"] is not None]
        if scored:
            fig = go.Figure()
            fig.add_trace(go.Bar(
                x=[r["label"] for r in scored], y=[r["t"] for r in scored],
                marker_color=[r["colore"] for r in scored],
                text=[f"{r['t']:.0f}" for r in scored], textposition="outside",
            ))
            fig.add_hline(y=50, line_dash="dash", line_color="gray", annotation_text="Media pop.")
            fig.update_layout(yaxis_title="T-score", height=380, margin=dict(t=20, b=20))
            cols[0].plotly_chart(fig, use_container_width=True)

        pop_comp = [r for r in cat_results if r.get("pop_mean") is not None]
        if pop_comp:
            fig2 = go.Figure()
            fig2.add_trace(go.Bar(name=nome, x=[r["label"] for r in pop_comp], y=[r["mean"] for r in pop_comp], marker_color=PRIMARY))
            fig2.add_trace(go.Bar(name="Media popolazione", x=[r["label"] for r in pop_comp], y=[r["pop_mean"] for r in pop_comp], marker_color="lightgray"))
            fig2.update_layout(barmode="group", height=380, margin=dict(t=20, b=20))
            cols[1].plotly_chart(fig2, use_container_width=True)

        table_rows = [{
            "Metrica": r["label"], "Unità": r["unit"], "N": r["n"],
            "Media": round(r["mean"], 3) if r["mean"] is not None else None,
            "Dev.Std": round(r["sd"], 3) if r["sd"] else None,
            "T-score": round(r["t"], 1) if r["t"] is not None else "—",
            "Valutazione": r["banda"] or "—",
        } for r in cat_results]
        st.dataframe(table_rows, use_container_width=True, hide_index=True)
        st.markdown("---")

with tab_checks:
    st.caption("Controlli a soglia sul CMJ Rebound, per individuare asimmetrie o esecuzioni tecnicamente scorrette.")
    for c in checks:
        col1, col2 = st.columns([3, 1])
        with col1:
            st.markdown(f"**{c['label']}**")
            st.caption(c["desc"])
        with col2:
            if c["value"] is None:
                st.markdown("—")
            else:
                icon = "✅" if c["passed"] else "⚠️"
                st.markdown(f"### {icon} {c['value']*100:.1f}%")
                soglia_lbl = "min" if c["direction"] == "min" else "max"
                st.caption(f"soglia {soglia_lbl} {c['threshold']*100:.0f}%")
        st.markdown("---")


# ============================================================================
# PARTE 6 — REPORT SCARICABILE (Word)
# ============================================================================

def genera_report_docx(nome, sesso, periodo, results, profilo, checks):
    doc = Document()
    title = doc.add_heading("FORCE PLATE TEST REPORT", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p = doc.add_paragraph()
    p.add_run("Atleta: ").bold = True
    p.add_run(f"{nome}    ")
    p.add_run("Sesso: ").bold = True
    p.add_run(f"{sesso}    ")
    p.add_run("Data test: ").bold = True
    p.add_run(f"{periodo}")

    doc.add_paragraph(
        "Per valutare i risultati del test è stato utilizzato il T-Score, un indice standardizzato "
        "che confronta la prestazione dell'atleta rispetto a un gruppo di riferimento, esprimendo la "
        "distanza dalla media in deviazioni standard. Punteggi tra 0 e 50 indicano valori inferiori "
        "alla media, mentre punteggi tra 50 e 100 indicano valori superiori alla media."
    )

    doc.add_heading("Profilo di Forza", level=1)
    tbl = doc.add_table(rows=1, cols=2)
    tbl.style = "Light Grid Accent 1"
    hdr = tbl.rows[0].cells
    hdr[0].text, hdr[1].text = "Categoria", "T-score"
    for cat in CATEGORIES:
        t = profilo.get(cat)
        row = tbl.add_row().cells
        row[0].text = cat
        row[1].text = f"{t:.0f}" if t is not None else "N/D"

    for cat in CATEGORIES:
        cat_results = [r for r in results if r["category"] == cat and r["mean"] is not None]
        if not cat_results:
            continue
        doc.add_heading(cat, level=2)
        tbl = doc.add_table(rows=1, cols=5)
        tbl.style = "Light List Accent 1"
        hdr = tbl.rows[0].cells
        for i, h in enumerate(["Metrica", "Unità", "Media", "N", "T-score / Valutazione"]):
            hdr[i].text = h
        for r in cat_results:
            row = tbl.add_row().cells
            row[0].text = r["label"]
            row[1].text = r["unit"] or ""
            row[2].text = f"{r['mean']:.3f}" if r["mean"] is not None else "N/D"
            row[3].text = str(r["n"])
            row[4].text = f"{r['t']:.0f} ({r['banda']})" if r["t"] is not None else "—"
        doc.add_paragraph()

    indici = [r for r in results if r["category"] == "INDICI" and r["mean"] is not None]
    if indici:
        doc.add_heading("Indici", level=1)
        for r in indici:
            doc.add_paragraph(
                f"{r['label']}: {r['mean']:.3f}"
                + (f"  —  T-score {r['t']:.0f} ({r['banda']})" if r["t"] is not None else ""),
                style="List Bullet",
            )

    doc.add_heading("Controlli Tecnici", level=1)
    for c in checks:
        if c["value"] is None:
            continue
        stato = "OK" if c["passed"] else "DA VERIFICARE"
        doc.add_paragraph(
            f"{c['label']}: {c['value']*100:.1f}% (soglia {'min' if c['direction']=='min' else 'max'} "
            f"{c['threshold']*100:.0f}%) — {stato}",
            style="List Bullet",
        )

    doc.add_heading("Analisi", level=1)
    doc.add_paragraph(
        "Spazio riservato al commento tecnico del preparatore, con osservazioni su punti di forza, "
        "aree di miglioramento e indicazioni di lavoro in palestra sulla base del profilo emerso."
    )

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


with tab_report:
    st.markdown("Genera un report Word (.docx) con profilo di forza, tabelle riassuntive e analisi testuale, in stile analogo al foglio REPORT originale.")
    if st.button("📄 Genera report Word", type="primary"):
        docx_bytes = genera_report_docx(nome, sesso, periodo, results, profilo, checks)
        st.download_button(
            "⬇️ Scarica report (.docx)", data=docx_bytes,
            file_name=f"Report_{nome.replace(' ', '_')}_{dt.date.today().isoformat()}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )