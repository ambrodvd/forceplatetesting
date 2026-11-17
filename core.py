"""Palette/tema, parsing dei file, analisi (T-score, profilo) e grafici Plotly."""

from __future__ import annotations

from dataclasses import dataclass
from dataclasses import field
import datetime as dt
import math
import re
import textwrap

import openpyxl
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from data import (
    BANDS,
    CATEGORIES,
    CHECKS,
    INDEX_ZONE_COLORS,
    INDEX_ZONE_LABELS,
    METRICS,
    banda_da_tscore,
    extra_metric_label,
    get_metric_description,
    zona_da_indice,
)

# ############################################################################
# TEMA — palette DU Coaching e CSS
# ############################################################################
# Palette brand DU Coaching (Destination Unknown), allineata al tema
# Streamlit fornito dall'utente:
#   primaryColor           #e94a26  -> ACCENT (riferimenti di popolazione/benchmark)
#   secondaryBackgroundColor #0bb6ff -> PRIMARY (dati dell'atleta: linee/punti nei grafici)
#   textColor               #065678 -> TEXT_COLOR (testi, titoli, chrome UI)
#   backgroundColor          #ffffff -> BG_COLOR (sfondo grafici/report)
PRIMARY = "#0bb6ff"     # azzurro — dati dell'atleta (linee/punti nei grafici)
ACCENT = "#e94a26"      # arancio — riferimenti di popolazione/benchmark
TEXT_COLOR = "#065678"  # blu scuro — testi, titoli, chrome dei widget
BG_COLOR = "#ffffff"    # sfondo


def apply_brand_css():
    # Applica la palette brand ai widget nativi di Streamlit (tab attive,
    # pulsanti, download button, checkbox) iniettando CSS direttamente nel file,
    # così il branding funziona subito senza dover distribuire/ricordare anche
    # un file .streamlit/config.toml separato. Nota: usa selettori CSS interni
    # di Streamlit che potrebbero cambiare in futuri aggiornamenti della libreria.
    st.markdown(f"""
    <style>
        .stTabs [data-baseweb="tab-highlight"] {{ background-color: {TEXT_COLOR} !important; }}
        .stTabs [aria-selected="true"] {{ color: {TEXT_COLOR} !important; }}
        .stButton > button, .stDownloadButton > button {{
            border-color: {TEXT_COLOR} !important;
            color: {TEXT_COLOR} !important;
        }}
        .stButton > button:hover, .stDownloadButton > button:hover,
        .stButton > button:active, .stDownloadButton > button:active {{
            border-color: {ACCENT} !important;
            color: {ACCENT} !important;
        }}
        input[type="checkbox"] {{ accent-color: {TEXT_COLOR} !important; }}
    </style>
    """, unsafe_allow_html=True)

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

    if _is_imtp_trial_layout(rows):
        return _parse_imtp_trial_rows(rows, filename)

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


def _is_imtp_trial_layout(rows) -> bool:
    """Rileva il layout Unit/Trial N (IMTP senza intestazione ForceMate
    standard): riga 1, colonna A = 'Unit', e almeno una colonna successiva
    che inizia con 'Trial'."""
    if not rows or not rows[0]:
        return False
    header = rows[0]
    if not header or str(header[6]).strip().lower() != "unit":
        return False
    return any(v and str(v).strip().lower().startswith("trial") for v in header[6:])


def _to_num(v):
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        try:
            return float(v.strip().replace(",", "."))
        except ValueError:
            return None
    return None


def _parse_imtp_trial_rows(rows, filename: str) -> ParsedFile:
    max_col = max(len(r) for r in rows)

    def cell(r_idx, c_idx):
        row = rows[r_idx - 1] if 0 < r_idx <= len(rows) else ()
        return row[c_idx - 1] if 0 < c_idx <= len(row) else None

    # Ogni blocco "Trial N" occupa 3 colonne: variabile, valore, spacer.
    # La colonna A contiene sempre l'unità di misura della riga.
    block_starts = [c for c in range(2, max_col + 1)
                    if cell(1, c) and str(cell(1, c)).strip().lower().startswith("trial")]

    reps = []
    for bs in block_starts:
        var_col, val_col = bs, bs + 1
        variables, units = {}, {}
        for r in range(2, len(rows) + 1):
            var_name = cell(r, var_col)
            if not var_name:
                continue
            var_name = str(var_name).strip().lower()
            val = cell(r, val_col)
            unit = cell(r, 1)
            num = _to_num(val)
            variables[var_name] = num if num is not None else val
            if unit:
                units[var_name] = str(unit).strip()

        # Non c'è un campo "body mass" in kg: c'è "system weight" (N),
        # cioè il peso corporeo come forza. Lo convertiamo per poter
        # calcolare IMTP Rel Peak Force. ASSUNZIONE DA VERIFICARE: g=9.80665.
        sw = variables.get("system weight")
        if "body mass" not in variables and isinstance(sw, (int, float)):
            variables["body mass"] = sw / 9.80665

        # Nessun tag "jump type" nel file: resta None e viene assegnato
        # dal selettore in sidebar (apply_type_override), come richiesto.
        reps.append({"jump_type": None, "vars": variables, "units": units})

    # Questo layout non contiene nome/sesso/peso/data: verranno ereditati
    # dagli altri file caricati insieme (vedi sesso_da_file()).
    metadata = {
        "nome": None, "sesso": None, "altezza_cm": None, "peso_kg_input": None,
        "data_test": None, "device": None, "team": None,
        "test_period": None, "test_type": None,
    }
    return ParsedFile(filename=filename, metadata=metadata, reps=reps)


FULL_META_COLS = {"Nome", "Sesso", "Data test", "File", "Jump Type", "Indice Ripetizione"}


def parse_full_export_to_parsed(file_like) -> ParsedFile:
    """Ricostruisce un ParsedFile dall'export completo (una riga per
    ripetizione, una colonna per variabile grezza), per rivedere grafici e
    tabelle senza i file XLSX originali. I dati sono già per ripetizione e
    completi: nulla va ricostruito a ritroso, nemmeno il peso corporeo
    delle metriche 'Rel'."""
    df = pd.read_csv(file_like, encoding="utf-8-sig")
    if "Jump Type" not in df.columns:
        raise ValueError("Il file non sembra un export completo: manca la colonna 'Jump Type'.")

    var_cols = [c for c in df.columns if c not in FULL_META_COLS]
    if not var_cols:
        raise ValueError("Nessuna colonna di variabili trovata nel file.")

    reps = []
    for _, row in df.iterrows():
        jump_type = str(row["Jump Type"]).strip().lower()
        if jump_type in ("", "—", "-", "nan", "none"):
            jump_type = None
        variables = {}
        for c in var_cols:
            v = pd.to_numeric(row[c], errors="coerce")
            if pd.notna(v):
                variables[c] = float(v)
        if variables:
            reps.append({"jump_type": jump_type, "vars": variables, "units": {}})

    if not reps:
        raise ValueError("Nessuna ripetizione valida nel file.")

    def _primo(colonna):
        if colonna in df.columns and len(df):
            v = str(df[colonna].iloc[0]).strip()
            if v and v.lower() not in ("nan", "none", "-"):
                return v
        return None

    metadata = {
        "nome": _primo("Nome"), "sesso": _primo("Sesso"), "altezza_cm": None,
        "peso_kg_input": None, "data_test": None, "device": None, "team": None,
        "test_period": None, "test_type": None, "periodo_override": _primo("Data test"),
    }
    return ParsedFile(filename="(da CSV)", metadata=metadata, reps=reps)


def _slug_nome(text):
    """Nome atleta utilizzabile in un nome file."""
    s = re.sub(r"[^A-Za-z0-9]+", "_", str(text or "").strip())
    return re.sub(r"_+", "_", s).strip("_") or "atleta"


def _data_file_token(files, periodo_str):
    """Data del test in formato ISO per il nome file: si preferisce la data
    reale dei file XLSX, poi quella scritta nel CSV ricaricato, infine oggi."""
    date_tests = [pf.metadata.get("data_test") for pf in files if pf.metadata.get("data_test")]
    if date_tests:
        return min(date_tests).strftime("%Y-%m-%d")
    m = re.search(r"(\d{2})/(\d{2})/(\d{4})", str(periodo_str or ""))
    if m:
        return f"{m.group(3)}-{m.group(2)}-{m.group(1)}"
    return dt.date.today().isoformat()


def build_full_raw_export(files, nome, sesso, periodo):
    """Dump grezzo di TUTTE le variabili presenti nei file XLSX caricati,
    non solo quelle riconosciute dal catalogo dell'app (curate o extra):
    pensato per studiare nuove metriche e costruire in futuro nuovi dati
    di popolazione. Una riga per ripetizione; le colonne sono l'unione di
    tutte le variabili incontrate in tutti i file/ripetizioni, quindi
    normale che compaiano molte celle vuote (una ripetizione non ha tutte
    le variabili di tutti gli altri tipi di test).

    Le ripetizioni deselezionate nella scheda Dettaglio Test vengono
    escluse, coerentemente con l'altro export e con medie/T-score."""
    all_vars = set()
    rep_entries = []
    tipo_counter = {}
    for pf in files:
        for i, rep in enumerate(pf.reps):
            jt = rep.get("jump_type")
            # Indice progressivo PER TIPO DI TEST attraverso tutti i file,
            # nello stesso ordine di collect_reps_all(): è la chiave con cui
            # sono salvate le checkbox "Prova N" (incl_{jump_type}_{indice}).
            # Va incrementato anche per le ripetizioni scartate, altrimenti
            # le successive si disallineano.
            idx_tipo = tipo_counter.get(jt, 0)
            tipo_counter[jt] = idx_tipo + 1
            if jt and not is_rep_included(jt, idx_tipo):
                continue
            rep_entries.append((pf.filename, i, rep))
            all_vars.update(rep["vars"].keys())
    all_vars_sorted = sorted(all_vars)

    rows = []
    for filename, idx, rep in rep_entries:
        row = {
            "Nome": nome, "Sesso": sesso, "Data test": periodo,
            "File": filename, "Jump Type": rep.get("jump_type") or "—",
            "Indice Ripetizione": idx + 1,
        }
        for var in all_vars_sorted:
            row[var] = rep["vars"].get(var)
        rows.append(row)
    return pd.DataFrame(rows)



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


def norma_extra(pop_dict, jump_type, raw_var, is_female):
    """(media, dev.std) di popolazione per una variabile grezza, letta dalle
    chiavi 'tipo::variabile' (presenti solo nelle norme Team). (None, None)
    se la norma non esiste o non c'è per quel sesso."""
    if not jump_type or not raw_var:
        return None, None
    c = pop_dict.get(f"{jump_type}::{raw_var}")
    if not c:
        return None, None
    return (c["mean_f"], c["sd_f"]) if is_female else (c["mean_m"], c["sd_m"])


# Tipo di test di ciascuna categoria (per le metriche extra).
CATEGORY_JUMP_TYPE = {
    "ISOMETRIC PULL TEST": "imtp", "SQUAT JUMP TEST": "sj",
    "COUNTERMOVEMENT JUMP TEST": "cmj", "COUNTERMOVEMENT JUMP REBOUND TEST": "cmrj",
}


def build_extra_defs(jump_type, category):
    """Definizioni delle metriche extra scelte con la ricerca in Dettaglio
    Test, lette dallo stato del multiselect: così build_results() le calcola
    insieme alle curate. La direzione non è nota: si assume 'più alto è
    meglio'."""
    return [
        dict(key=f"extra::{jump_type}::{name}", label=extra_metric_label(name), category=category,
             jump_type=jump_type, raw_var=name, unit="", pop_key=None,
             lower_is_better=False, kind="info", desc=get_metric_description(jump_type, name))
        for name in st.session_state.get(f"extra_metrics_{jump_type}", [])
    ]


def nel_profilo(r):
    """True se il T-score della metrica entra nel Profilo di Forza. Scelta
    dell'utente (checkbox in Dettaglio Test); di default TUTTE le metriche
    con un T-score."""
    return st.session_state.get(f"prof_{r['key']}", True)


def collect_reps_all(files, jump_type):
    """Tutte le ripetizioni di un tipo, nell'ordine di caricamento — senza
    applicare le esclusioni scelte dall'utente. Usata per popolare le
    colonne 'Salto N' nel Dettaglio Test."""
    return [rep for pf in files for rep in pf.reps if rep["jump_type"] == jump_type]


def is_rep_included(jump_type, index, default=True):
    return st.session_state.get(f"incl_{jump_type}_{index}", default)


def collect_reps(files, jump_type):
    """Ripetizioni di un tipo, filtrate escludendo i salti deselezionati
    dall'utente nella scheda Dettaglio Test (checkbox 'Salto N')."""
    reps_all = collect_reps_all(files, jump_type)
    return [rep for i, rep in enumerate(reps_all) if is_rep_included(jump_type, i)]


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


def per_rep_metric_values(files, metric):
    """Valore della metrica per OGNI ripetizione (inclusi i salti esclusi
    dal calcolo), nello stesso ordine delle colonne 'Salto N'."""
    reps_all = collect_reps_all(files, metric["jump_type"])
    values = []
    for rep in reps_all:
        if metric.get("raw_var"):
            v = rep["vars"].get(metric["raw_var"])
        elif metric.get("derive"):
            v = metric["derive"](rep["vars"])
        else:
            v = None
        values.append(v if isinstance(v, (int, float)) else None)
    return values


def best_worst_indices(values, incl_mask, lower_is_better):
    """Indice del salto migliore e peggiore tra quelli inclusi e numerici."""
    candidates = [(i, v) for i, (v, inc) in enumerate(zip(values, incl_mask))
                  if inc and isinstance(v, (int, float))]
    if len(candidates) < 2:
        return None, None
    key = (lambda t: t[1])
    best = (min if lower_is_better else max)(candidates, key=key)[0]
    worst = (max if lower_is_better else min)(candidates, key=key)[0]
    if best == worst:
        return None, None
    return best, worst


def sesso_da_file(files):
    for pf in files:
        if pf.metadata.get("sesso"):
            return pf.metadata["sesso"]
    return "UOMO"


def build_results(files, pop_dict, thresholds):
    sesso_raw = sesso_da_file(files)
    is_female = sesso_raw and sesso_raw.upper().startswith("F")

    support = {}
    results = []

    # Metriche curate + metriche extra scelte con la ricerca in Dettaglio Test.
    extra_defs = [m for cat, jt in CATEGORY_JUMP_TYPE.items() for m in build_extra_defs(jt, cat)]

    for metric in METRICS + extra_defs:
        if metric["kind"] in ("score", "info", "support") and metric.get("jump_type"):
            values = compute_metric_values(files, metric)
            stats = pooled_stats(values)
            support[metric["key"]] = stats

            # Norma: chiave curata se esiste, altrimenti 'tipo::variabile'
            # (solo nelle norme Team). Il T-score si calcola ovunque ci sia una
            # norma; se entra nel profilo lo decide l'utente (nel_profilo).
            # Le metriche "support" (solo per i CHECKS) non hanno mai norma,
            # quindi niente T-score anche se esiste una norma Team.
            if metric["kind"] == "support":
                pop_mean, pop_sd = None, None
            elif metric.get("pop_key") and metric["pop_key"] in pop_dict:
                # Non tutti i set hanno tutte le chiavi (es. Net Rel Impulse
                # esiste solo nelle norme Team): se manca, niente T-score.
                pop = pop_dict[metric["pop_key"]]
                pop_mean = pop["mean_f"] if is_female else pop["mean_m"]
                pop_sd = pop["sd_f"] if is_female else pop["sd_m"]
            else:
                pop_mean, pop_sd = norma_extra(pop_dict, metric.get("jump_type"),
                                               metric.get("raw_var"), is_female)
            z, t = z_t_score(stats["mean"], pop_mean, pop_sd, metric["lower_is_better"])
            banda, colore = banda_da_tscore(t)

            results.append(dict(
                key=metric["key"], label=metric["label"], category=metric["category"],
                unit=metric["unit"], kind=metric["kind"], n=stats["n"], mean=stats["mean"], sd=stats["sd"],
                z=z, t=t, banda=banda, colore=colore, pop_mean=pop_mean, pop_sd=pop_sd,
                desc=metric.get("desc"),
            ))

    # --- Indici derivati da medie aggregate (DSI, EUR) ---
    cmj_peak = support.get("cmj_peak_force", {}).get("mean")
    imtp_peak = support.get("imtp_peak_force", {}).get("mean")
    dsi_val = (cmj_peak / imtp_peak) if (cmj_peak and imtp_peak) else None

    # EUR = CMJ height / SJ height (rapporto puro), coerente con la tabella
    # costanti fornita dall'utente (~1,11 uomini, ~1,09 donne).
    cmj_h = support.get("cmj_height", {}).get("mean")
    sj_h = support.get("sj_height", {}).get("mean")
    eur_val = (cmj_h / sj_h) if (cmj_h and sj_h) else None

    for key, val in (("dsi", dsi_val), ("eur", eur_val)):
        pop = pop_dict[key]
        pop_mean = pop["mean_f"] if is_female else pop["mean_m"]
        metric_def = next(m for m in METRICS if m["key"] == key)
        thr_low, thr_high = thresholds[key]
        zona, zona_colore = zona_da_indice(key, val, thr_low, thr_high)
        # Niente T-score per DSI/EUR: non sono metriche "più alto = meglio"
        # ma indici diagnostici, dove entrambi gli estremi descrivono un
        # profilo diverso, non uno migliore. Resta la media di popolazione,
        # mostrata come riferimento nell'help e nel report.
        results.append(dict(
            key=key, label=metric_def["label"], category="INDICI", unit="",
            kind="score_single", n=(1 if val is not None else 0), mean=val, sd=None,
            z=None, t=None, banda=None, colore=None, pop_mean=pop_mean, pop_sd=None,
            zona=zona, zona_colore=zona_colore, thr_low=thr_low, thr_high=thr_high,
            desc=metric_def.get("desc"),
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
    cmj_re_contact_time = support.get("cmj_re_contact_time", {}).get("mean")

    checks_out = [
        make_check(CHECKS[0], (rebound_h / initial_h) if (initial_h and rebound_h) else None),
        make_check(CHECKS[1], (initial_h / cmj_h_standalone) if (initial_h and cmj_h_standalone) else None),
        make_check(CHECKS[2], (landing_sym / 100.0) if landing_sym is not None else None),
        make_check(CHECKS[3], cmj_re_contact_time if cmj_re_contact_time is not None else None),
    ]

    return dict(results=results, support=support, checks=checks_out, sesso=sesso_raw)


def profilo_forza(results):
    """Media dei T-score per categoria, solo delle metriche spuntate in
    'Metriche incluse nel Profilo di Forza' (vedi nel_profilo)."""
    out = {}
    for cat in CATEGORIES:
                ts = [r["t"] for r in results
              if r["category"] == cat and r["t"] is not None and nel_profilo(r)]
                out[cat] = (sum(ts) / len(ts)) if ts else None
    return out


# ============================================================================

# PARTE 4bis — GRAFICI A QUADRANTI E A BANDE
# ============================================================================
# Due famiglie di grafici bidimensionali:
#
#  - quadrant_chart (RSQ): incrocia due metriche e divide il piano nei
#    quattro quadranti definiti dal crosshair centrato sulla media di
#    popolazione. Legge il livello assoluto nelle due metriche.
#
#  - ratio_wedge_chart (DSI, EUR): per un indice che è un RAPPORTO tra le
#    due metriche. Qui i quadranti non servono: tutti i punti con lo stesso
#    rapporto stanno su una retta per l'origine, quindi il piano va diviso
#    da rette diagonali (y = x / soglia), non da una croce. Il crosshair di
#    popolazione resta come contesto sul livello assoluto.

def _wrap_label(text, width=16):
    """Spezza un'etichetta lunga su più righe (per gli assi angolari del
    radar e per le etichette delle zone), così non deborda oltre il bordo
    del grafico invece di essere tagliata."""
    return "<br>".join(textwrap.wrap(text, width=width, break_long_words=False))


def quadrant_chart(x_mean, y_mean, x_sd, y_sd, x_label, y_label, quadrant_defs,
                   pop_x=None, pop_y=None, point_color=PRIMARY, height=430):
    """quadrant_defs: dict con chiavi 'tl','tr','bl','br' -> (etichetta, colore).
    Il crosshair (linee tratteggiate e confini dei quadranti) è centrato sulla
    media di popolazione (pop_x, pop_y); se non disponibile per una metrica si
    usa la media del test come fallback. Il punto mostrato è la media del test
    con barre d'errore pari alla deviazione standard delle ripetizioni incluse.
    Ritorna una go.Figure oppure None se manca la media del test su uno dei due assi."""
    if x_mean is None or y_mean is None:
        return None
    x_sd = x_sd or 0
    y_sd = y_sd or 0
    cx = pop_x if pop_x is not None else x_mean
    cy = pop_y if pop_y is not None else y_mean

    # Raggio simmetrico intorno al crosshair (cx, cy): copre la distanza
    # massima tra il crosshair e gli estremi del punto (media ± dev.std),
    # più un margine di respiro. Essendo simmetrico, il crosshair cade
    # sempre esattamente al centro del grafico e i quattro quadranti
    # risultano sempre della stessa dimensione.
    x_extent = abs(x_mean - cx) + x_sd
    y_extent = abs(y_mean - cy) + y_sd
    x_r = x_extent * 1.3 if x_extent > 0 else (abs(cx) * 0.15 or 1)
    y_r = y_extent * 1.3 if y_extent > 0 else (abs(cy) * 0.15 or 1)
    x0, x1 = cx - x_r, cx + x_r
    y0, y1 = cy - y_r, cy + y_r

    tl_label, tl_color = quadrant_defs["tl"]
    tr_label, tr_color = quadrant_defs["tr"]
    bl_label, bl_color = quadrant_defs["bl"]
    br_label, br_color = quadrant_defs["br"]

    fig = go.Figure()
    for x_a, x_b, y_a, y_b, color in (
        (x0, cx, cy, y1, tl_color), (cx, x1, cy, y1, tr_color),
        (x0, cx, y0, cy, bl_color), (cx, x1, y0, cy, br_color),
    ):
        fig.add_shape(type="rect", x0=x_a, x1=x_b, y0=y_a, y1=y_b,
                      fillcolor=color, opacity=0.15, line_width=0)

    for x_pos, y_pos, text, anchor in (
        ((x0 + cx) / 2, y1, tl_label, "top"), ((cx + x1) / 2, y1, tr_label, "top"),
        ((x0 + cx) / 2, y0, bl_label, "bottom"), ((cx + x1) / 2, y0, br_label, "bottom"),
    ):
        fig.add_annotation(x=x_pos, y=y_pos, text=text, showarrow=False,
                           font=dict(size=10, color="#666"), yanchor=anchor)

    fig.add_vline(x=cx, line_dash="dash", line_color=ACCENT)
    fig.add_hline(y=cy, line_dash="dash", line_color=ACCENT)

    fig.add_trace(go.Scatter(
        x=[x_mean], y=[y_mean], mode="markers",
        marker=dict(size=12, color=point_color, line=dict(width=1.5, color="white")),
        error_x=dict(type="data", array=[x_sd], visible=True, color=point_color, thickness=1.5, width=6),
        error_y=dict(type="data", array=[y_sd], visible=True, color=point_color, thickness=1.5, width=6),
        name="Media test \u00b1 Dev.Std",
        hovertemplate=(f"Media test<br>{x_label}: %{{x:.2f}} \u00b1 {x_sd:.2f}"
                       f"<br>{y_label}: %{{y:.2f}} \u00b1 {y_sd:.2f}<extra></extra>"),
    ))

    fig.update_layout(
        xaxis_title=x_label, yaxis_title=y_label,
        xaxis_range=[x0, x1], yaxis_range=[y0, y1],
        xaxis_showgrid=False, yaxis_showgrid=False,
        xaxis=dict(automargin=True, title=dict(standoff=12)),
        yaxis=dict(automargin=True, title=dict(standoff=12)),
        height=height, margin=dict(t=30, b=50, l=70, r=20), showlegend=False,
        plot_bgcolor=BG_COLOR, paper_bgcolor=BG_COLOR, font=dict(color=TEXT_COLOR),
    )
    return fig


def _clip_half_plane(poly, f, keep_positive):
    """Sutherland-Hodgman: ritaglia un poligono contro il semipiano f>=0
    (keep_positive) oppure f<=0. f(x, y) è lineare, quindi l'intersezione con
    ogni lato si trova per interpolazione lineare dei valori di f."""
    out = []
    for i in range(len(poly)):
        cur, prv = poly[i], poly[i - 1]
        fc, fp = f(*cur), f(*prv)
        inc = fc >= 0 if keep_positive else fc <= 0
        inp = fp >= 0 if keep_positive else fp <= 0
        if inc != inp and (fp - fc) != 0:
            t = fp / (fp - fc)
            out.append((prv[0] + t * (cur[0] - prv[0]), prv[1] + t * (cur[1] - prv[1])))
        if inc:
            out.append(cur)
    return out


def _poly_path(poly):
    if len(poly) < 3:
        return None
    return "M " + " L ".join(f"{x},{y}" for x, y in poly) + " Z"


def _poly_centroid(poly):
    """Baricentro del poligono (per posizionare l'etichetta della zona)."""
    n, a, cx, cy = len(poly), 0.0, 0.0, 0.0
    for i in range(n):
        x0, y0 = poly[i - 1]
        x1, y1 = poly[i]
        cr = x0 * y1 - x1 * y0
        a += cr
        cx += (x0 + x1) * cr
        cy += (y0 + y1) * cr
    if a == 0:
        return sum(p[0] for p in poly) / n, sum(p[1] for p in poly) / n
    return cx / (3 * a), cy / (3 * a)


def ratio_wedge_chart(x_mean, y_mean, x_sd, y_sd, x_label, y_label, key,
                      thr_low, thr_high, pop_x=None, pop_y=None,
                      point_color=PRIMARY, height=470):
    """Grafico a bande per un indice-rapporto (indice = x / y).

    A differenza di quadrant_chart le zone NON sono i quattro quadranti del
    crosshair, ma tre spicchi delimitati dalle rette y = x/soglia uscenti
    dall'origine: essendo l'indice un rapporto, tutti i punti con lo stesso
    indice stanno su una retta per l'origine, quindi è la posizione rispetto
    alle diagonali a rappresentare l'indice. Le linee tratteggiate della media
    di popolazione restano come contesto sul livello assoluto e vengono
    disegnate SOLO se quel dato di popolazione esiste davvero.

    L'origine può restare fuori dal range: gli spicchi sono ritagliati sul
    riquadro visibile. Il range verticale viene però allargato quanto basta
    perché entrambe le diagonali attraversino il riquadro, altrimenti un
    atleta lontano dalle soglie vedrebbe un grafico tinto di un solo colore,
    senza alcun confine visibile su cui orientarsi."""
    if x_mean is None or y_mean is None or not thr_low or not thr_high:
        return None
    x_sd, y_sd = x_sd or 0, y_sd or 0
    cx = pop_x if pop_x is not None else x_mean
    cy = pop_y if pop_y is not None else y_mean

    x_ext, y_ext = abs(x_mean - cx) + x_sd, abs(y_mean - cy) + y_sd
    x_r = x_ext * 1.3 if x_ext > 0 else (abs(cx) * 0.15 or 1)
    y_r = y_ext * 1.3 if y_ext > 0 else (abs(cy) * 0.15 or 1)
    for thr in (thr_low, thr_high):
        y_r = max(y_r, abs(cx / thr - cy) * 1.08)
    x0, x1 = cx - x_r, cx + x_r
    y0, y1 = cy - y_r, cy + y_r

    box = [(x0, y0), (x1, y0), (x1, y1), (x0, y1)]
    f_low = lambda x, y: y - x / thr_low
    f_high = lambda x, y: y - x / thr_high
    zones = [
        (_clip_half_plane(box, f_low, True), 0),
        (_clip_half_plane(_clip_half_plane(box, f_low, False), f_high, True), 1),
        (_clip_half_plane(box, f_high, False), 2),
    ]
    labels = INDEX_ZONE_LABELS[key]

    fig = go.Figure()
    for poly, i in zones:
        path = _poly_path(poly)
        if not path:
            continue
        fig.add_shape(type="path", path=path, fillcolor=INDEX_ZONE_COLORS[i],
                      opacity=0.18, line_width=0, layer="below")
        lx, ly = _poly_centroid(poly)
        fig.add_annotation(x=lx, y=ly, text=_wrap_label(labels[i], 20), showarrow=False,
                           font=dict(size=10, color="#666"), align="center")

    for thr in (thr_low, thr_high):
        fig.add_shape(type="line", x0=x0, y0=x0 / thr, x1=x1, y1=x1 / thr,
                      line=dict(color="#8d8d8d", dash="dot", width=3))

    # Le linee di popolazione compaiono solo dove il dato esiste: oggi
    # cmj_peak_force non ha una norma, quindi il DSI mostra la sola
    # orizzontale (IMTP). Aggiungendo la norma mancante a DEFAULT_POP /
    # CONST_LABELS e impostando pop_key sulla metrica, la verticale
    # comparirà da sola senza altre modifiche.
    # Peso visivo: le diagonali (che rappresentano l'indice) sono spesse,
    # il crosshair di popolazione è sottile — è l'indice il soggetto del
    # grafico, la popolazione è solo contesto.
    if pop_x is not None:
        fig.add_vline(x=pop_x, line_dash="dash", line_color=ACCENT, line_width=1)
    if pop_y is not None:
        fig.add_hline(y=pop_y, line_dash="dash", line_color=ACCENT, line_width=1)

    fig.add_trace(go.Scatter(
        x=[x_mean], y=[y_mean], mode="markers",
        marker=dict(size=12, color=point_color, line=dict(width=1.5, color="white")),
        error_x=dict(type="data", array=[x_sd], visible=True, color=point_color, thickness=1.5, width=6),
        error_y=dict(type="data", array=[y_sd], visible=True, color=point_color, thickness=1.5, width=6),
        name="Media test \u00b1 Dev.Std",
        hovertemplate=(f"Media test<br>{x_label}: %{{x:.2f}} \u00b1 {x_sd:.2f}"
                       f"<br>{y_label}: %{{y:.2f}} \u00b1 {y_sd:.2f}<extra></extra>"),
    ))

    fig.update_layout(
        xaxis_title=x_label, yaxis_title=y_label,
        xaxis_range=[x0, x1], yaxis_range=[y0, y1],
        xaxis=dict(showgrid=False, automargin=True, title=dict(standoff=12)),
        yaxis=dict(showgrid=False, automargin=True, title=dict(standoff=12)),
        height=height, margin=dict(t=30, b=50, l=70, r=20), showlegend=False,
        plot_bgcolor=BG_COLOR, paper_bgcolor=BG_COLOR, font=dict(color=TEXT_COLOR),
    )
    return fig


def ratio_band_strip(key, value, thr_low, thr_high, label, decimals=2, height=120):
    """Barra orizzontale a tre bande con il valore dell'indice: la zona si
    legge a colpo d'occhio, senza dover interpretare una pendenza. Range =
    soglie ± 0.2, allargato automaticamente se il valore cade fuori.

    Lo spessore della fascia NON si imposta direttamente: add_vrect riempie
    sempre tutta l'area di plot, quindi la fascia è alta esattamente
    (height - margin.t - margin.b) = 10px. Per spostare le etichette senza
    ingrassare la fascia si alzano margin_t E height della stessa quantità.
    Il marcatore (13px) e il suo valore sono più alti della fascia: senza
    cliponaxis=False verrebbero tagliati dai bordi dell'area di plot."""
    if value is None or not thr_low or not thr_high:
        return None
    lo = min(thr_low - 0.2, value - 0.05)
    hi = max(thr_high + 0.2, value + 0.05)
    labels = INDEX_ZONE_LABELS[key]

    fig = go.Figure()
    for a, b, i in ((lo, thr_low, 0), (thr_low, thr_high, 1), (thr_high, hi, 2)):
        fig.add_vrect(x0=a, x1=b, fillcolor=INDEX_ZONE_COLORS[i], opacity=0.35, line_width=0)
        # Le etichette sono annotazioni in coordinate "paper" sopra l'area di
        # plot: non partecipano all'automargin, quindi lo spazio per loro va
        # riservato a mano con margin_t. Wrap a 22 caratteri = due righe al
        # massimo per le etichette attuali.
        fig.add_annotation(x=(a + b) / 2, y=1.10, xref="x", yref="paper", yanchor="bottom",
                           text=_wrap_label(labels[i], 22), showarrow=False,
                           font=dict(size=11, color="#484343"))
    fig.add_trace(go.Scatter(
        x=[value], y=[0], mode="markers+text",
        marker=dict(size=13, color=PRIMARY, symbol="diamond", line=dict(width=1.5, color="white")),
        text=[f"{value:.{decimals}f}"], textposition="bottom center",
        textfont=dict(color=TEXT_COLOR, size=13),
        cliponaxis=False,
        hovertemplate=f"{label}: %{{x:.{decimals}f}}<extra></extra>",
    ))
    fig.update_layout(
        xaxis=dict(range=[lo, hi], showgrid=False, tickvals=[thr_low, thr_high],
                   ticktext=[f"{thr_low:.2f}", f"{thr_high:.2f}"], automargin=True,
                   title=dict(text=label, standoff=12)),
        yaxis=dict(range=[-1, 1], showgrid=False, showticklabels=False, zeroline=False),
        height=height, margin=dict(t=60, b=50, l=20, r=20), showlegend=False,
        plot_bgcolor=BG_COLOR, paper_bgcolor=BG_COLOR, font=dict(color=TEXT_COLOR),
    )
    return fig


RSQ_QUADRANTS = dict(
    tl=("Alta Reattività", "#4FC3F7"),
    tr=("Forza-Dominante", "#EF5350"),
    bl=("Elastico-Dominante", "#66BB6A"),
    br=("Bassa Reattività", "#FFEE58"),
)

# Metriche (altezza, tempo) usate per il grafico RSQ di ciascuna categoria
# che include un mRSI. Definito a livello di modulo per essere condiviso
# tra la scheda Dettaglio Test e il report HTML.
RSQ_METRIC_KEYS = {
    "COUNTERMOVEMENT JUMP TEST": ("cmj_height", "cmj_contraction_time"),
    "COUNTERMOVEMENT JUMP REBOUND TEST": ("cmj_re_rebound_height", "cmj_re_contact_time"),
}


def build_rsq_chart(cat, results):
    """Grafico RSQ per una categoria, se applicabile. Ritorna None altrimenti."""
    if cat not in RSQ_METRIC_KEYS:
        return None
    h_key, t_key = RSQ_METRIC_KEYS[cat]
    h_metric = next(m for m in METRICS if m["key"] == h_key)
    t_metric = next(m for m in METRICS if m["key"] == t_key)
    r_h = next(r for r in results if r["key"] == h_key)
    r_t = next(r for r in results if r["key"] == t_key)
    return quadrant_chart(
        x_mean=r_t["mean"], y_mean=r_h["mean"], x_sd=r_t["sd"], y_sd=r_h["sd"],
        x_label=f"{t_metric['label']} ({t_metric['unit']})",
        y_label=f"{h_metric['label']} ({h_metric['unit']})",
        quadrant_defs=RSQ_QUADRANTS, pop_x=r_t["pop_mean"], pop_y=r_h["pop_mean"],
        point_color=PRIMARY,
    )


def build_dsi_chart(results, thresholds):
    """DSI: CMJ Peak Propulsive Force (x) vs IMTP Peak Force (y). La linea
    verticale compare solo quando cmj_peak_force avrà un dato di popolazione
    (oggi pop_key=None -> pop_mean None -> linea nascosta)."""
    r_cmj_peak = next((r for r in results if r["key"] == "cmj_peak_force"), None)
    r_imtp_peak = next((r for r in results if r["key"] == "imtp_peak_force"), None)
    if r_cmj_peak is None or r_imtp_peak is None:
        return None
    thr_low, thr_high = thresholds["dsi"]
    return ratio_wedge_chart(
        x_mean=r_cmj_peak["mean"], y_mean=r_imtp_peak["mean"],
        x_sd=r_cmj_peak["sd"], y_sd=r_imtp_peak["sd"],
        x_label="CMJ Peak Propulsive Force (N)", y_label="IMTP Peak Force (N)",
        key="dsi", thr_low=thr_low, thr_high=thr_high,
        pop_x=r_cmj_peak["pop_mean"], pop_y=r_imtp_peak["pop_mean"],
    )


def build_eur_chart(results, thresholds):
    """EUR: CMJ Height (x) vs SJ Height (y). Entrambe hanno un dato di
    popolazione, quindi il crosshair è completo."""
    r_sj_h = next((r for r in results if r["key"] == "sj_height"), None)
    r_cmj_h = next((r for r in results if r["key"] == "cmj_height"), None)
    if r_sj_h is None or r_cmj_h is None:
        return None
    thr_low, thr_high = thresholds["eur"]
    return ratio_wedge_chart(
        x_mean=r_cmj_h["mean"], y_mean=r_sj_h["mean"], x_sd=r_cmj_h["sd"], y_sd=r_sj_h["sd"],
        x_label="CMJ Height (cm)", y_label="SJ Height (cm)",
        key="eur", thr_low=thr_low, thr_high=thr_high,
        pop_x=r_cmj_h["pop_mean"], pop_y=r_sj_h["pop_mean"],
    )


def build_tscore_bar_chart(cat_results):
    """Grafico a barre orizzontali del T-score per una categoria (usato sia
    nella scheda Dettaglio Test sia nel report HTML). Ritorna None se nessuna
    metrica della categoria ha un T-score calcolabile."""
    scored = [r for r in cat_results if r["t"] is not None]
    if not scored:
        return None

    scored_sorted = sorted(scored, key=lambda r: r["t"])
    labels = [r["label"] for r in scored_sorted]
    tvals = [r["t"] for r in scored_sorted]
    colors = [r["colore"] for r in scored_sorted]

    fig = go.Figure()

    for lo, hi, band_label, band_color in BANDS:
        lo_c, hi_c = max(lo, 0), min(hi, 100)
        if lo_c >= hi_c:
            continue
        # Il rettangolo di sfondo (banda) resta ancorato ai dati (asse x),
        # ma l'etichetta viene disegnata separatamente appena sopra l'area
        # di plot (yref="paper", y>1) invece che dentro di essa: così non
        # si sovrappone mai alla barra più in alto, indipendentemente dal
        # numero di metriche mostrate.
        fig.add_vrect(x0=lo_c, x1=hi_c, fillcolor=band_color, opacity=0.3, line_width=0)
        fig.add_annotation(
            x=(lo_c + hi_c) / 2, y=1.02, xref="x", yref="paper", yanchor="bottom",
            text=band_label, showarrow=False, font=dict(size=9, color="#484343"),
        )

    fig.add_trace(go.Bar(
        x=tvals, y=labels, base=0, orientation="h",
        marker_color=colors, customdata=tvals,
        text=[f"{t:.0f}" for t in tvals], textposition="outside",
        textfont=dict(color="#484343", size=12),
        hovertemplate="%{y}: T-score %{customdata:.0f}<extra></extra>",
    ))
    fig.add_vline(x=50, line_dash="dash", line_color=ACCENT)
    fig.update_layout(
        xaxis_title="T-score", xaxis_range=[0, 100],
        xaxis=dict(showgrid=False, ticks="", automargin=True, title=dict(standoff=18)),
        yaxis=dict(showgrid=False, automargin=True),
        height=max(360, 70 * len(scored_sorted) + 150),
        margin=dict(t=70, b=60, l=10, r=20), showlegend=False,
        plot_bgcolor=BG_COLOR, paper_bgcolor=BG_COLOR,
        font=dict(color=TEXT_COLOR),
    )
    return fig


def build_radar_chart(cats_valide, profilo, nome):
    """Radar (Scatterpolar) del profilo di forza (usato sia nella scheda
    Profilo di Forza sia nel report HTML). Ritorna None se non ci sono
    categorie valide."""
    if not cats_valide:
        return None
    vals = [profilo[c] for c in cats_valide]
    vals_closed = vals + [vals[0]]
    cats_closed = [_wrap_label(c) for c in cats_valide] + [_wrap_label(cats_valide[0])]

    fig = go.Figure()
    fig.add_trace(go.Scatterpolar(
        r=[50] * (len(cats_valide) + 1), theta=cats_closed, mode="lines",
        line=dict(color="rgba(233,74,38,0.6)", dash="dash"), name="Media popolazione (T=50)"
    ))
    fig.add_trace(go.Scatterpolar(
        r=vals_closed, theta=cats_closed, fill="toself",
        line=dict(color=PRIMARY, width=3), fillcolor="rgba(11,182,255,0.25)", name=nome
    ))
    fig.update_layout(
        polar=dict(
            domain=dict(x=[0.12, 0.88], y=[0.08, 0.92]),
            radialaxis=dict(range=[0, 100], showticklabels=True, ticks="", tickfont=dict(color=TEXT_COLOR)),
        ),
        font=dict(color=TEXT_COLOR),
        paper_bgcolor=BG_COLOR,
        showlegend=True, height=560, margin=dict(t=60, b=60, l=60, r=60),
    )
    return fig