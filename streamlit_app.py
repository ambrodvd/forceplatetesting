# -*- coding: utf-8 -*-
"""
Force Plate Test Report — app Streamlit (versione a file singolo)

Carica gli export XLSX di ForceMate/ForceDecks (IMTP, SJ, CMJ, CMJ RE),
calcola automaticamente medie, T-score rispetto alla popolazione di
riferimento e produce un profilo di forza con grafici Plotly e report HTML interattivo.

Struttura del file:
  PARTE 1 — Costanti e dati di popolazione
  PARTE 2 — Caricamento file (sidebar upload)
  PARTE 3 — Lettura/parsing dei file XLSX
  PARTE 4 — Analisi dati e confronto con la popolazione
  PARTE 5 — Report live (UI a schede)
  PARTE 6 — Report scaricabile (HTML interattivo)
"""

from __future__ import annotations

import io
import math
import re
import textwrap
import html as _html
import datetime as dt
from dataclasses import dataclass, field

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import openpyxl


# ============================================================================
# PARTE 1 — COSTANTI E DATI DI POPOLAZIONE
# ============================================================================
# Equivalente ai fogli "DATI POP" (norme di popolazione) e "DATI TEST"
# (definizione delle metriche) del Google Sheet originale.

CATEGORIES = ["ISOMETRIC PULL TEST", "SQUAT JUMP TEST", "COUNTERMOVEMENT JUMP TEST", "COUNTERMOVEMENT JUMP REBOUND TEST"]

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
#
# Questi sono i valori DI DEFAULT: l'utente può visionarli, modificarli,
# scaricarli e ricaricarli da un file Excel nella scheda "⚙️ Costanti"
# (vedi PARTE 4bis). A runtime l'app usa sempre st.session_state["pop"],
# inizializzato da questo dizionario alla prima esecuzione.
#
# NB: EUR è definito come rapporto puro CMJ height / SJ height (non come
# differenza percentuale), coerente con la tabella costanti fornita
# dall'utente (~1.11 uomini, ~1.09 donne).
DEFAULT_POP = {
    "imtp_peak_force":        dict(mean_m=2606.8,      sd_m=646.1163194,  mean_f=1575.0,      sd_f=386.145544),
    "imtp_rel_peak_force":    dict(mean_m=34.56,        sd_m=5.27,         mean_f=34.56,        sd_f=5.27),
    "sj_mean_power":          dict(mean_m=1180.0,       sd_m=414.1638849,  mean_f=823.0,        sd_f=240.2447942),
    "sj_height":              dict(mean_m=26.9,         sd_m=6.57,         mean_f=19.35,        sd_f=5.51),
    "sj_contraction_time":    dict(mean_m=0.445,        sd_m=0.127835797,  mean_f=0.46,         sd_f=0.137032268),
    "cmj_height":             dict(mean_m=30.01,        sd_m=6.5,          mean_f=20.91,        sd_f=5.84),
    "mrsi_cmj":               dict(mean_m=0.419,        sd_m=0.098531539,  mean_f=0.308,        sd_f=0.093831019),
    "dsi":                    dict(mean_m=0.7,          sd_m=0.074074074,  mean_f=0.7,          sd_f=0.074074074),
    "eur":                    dict(mean_m=1.107708605,  sd_m=0.038034873,  mean_f=1.090563116,  sd_f=0.02656191),
    "cmj_re_rebound_height":  dict(mean_m=38.5,         sd_m=5.5996817,    mean_f=38.5,         sd_f=5.5996817),
    "cmj_re_contact_time":    dict(mean_m=0.25,         sd_m=0.148005087,  mean_f=0.25,         sd_f=0.148005087),
    "cmj_re_rebound_impulse": dict(mean_m=541.5,        sd_m=53.82161458,  mean_f=541.5,        sd_f=53.82161458),
    "mrsi_cmj_re":            dict(mean_m=1.37,         sd_m=0.331705729,  mean_f=1.37,         sd_f=0.331705729),
}

# Etichetta e unità di misura di ciascuna costante, usate sia per la
# tabella a schermo sia per il file Excel scaricabile/ricaricabile.
# Le etichette ricalcano quelle del file "DATO" fornito dall'utente, in
# modo che un file con quella struttura venga riconosciuto automaticamente.
CONST_LABELS = {
    "imtp_peak_force":        ("IMTP abs", "N", "ISOMETRIC PULL TEST"),
    "imtp_rel_peak_force":    ("IMTP rel", "N/kg", "ISOMETRIC PULL TEST"),
    "sj_mean_power":          ("SJ mean power", "W", "SQUAT JUMP TEST"),
    "sj_height":               ("SJ height", "cm", "SQUAT JUMP TEST"),
    "sj_contraction_time":     ("SJ contraction time", "s", "SQUAT JUMP TEST"),
    "cmj_height":              ("CMJ height", "cm", "COUNTERMOVEMENT JUMP TEST"),
    "mrsi_cmj":                ("Mrsi-CMJ", "m/s", "COUNTERMOVEMENT JUMP TEST"),
    "cmj_re_rebound_height":   ("CMJ RE Rebound Jump height", "cm", "COUNTERMOVEMENT JUMP REBOUND TEST"),
    "cmj_re_contact_time":     ("CMJ RE Contact Time", "s", "COUNTERMOVEMENT JUMP REBOUND TEST"),
    "cmj_re_rebound_impulse":  ("CMJ RE propulsive impulse", "N\u00b7s", "COUNTERMOVEMENT JUMP REBOUND TEST"),
    "mrsi_cmj_re":             ("mRSI-CMJ RE", "m/s", "COUNTERMOVEMENT JUMP REBOUND TEST"),
    "dsi":                     ("DSI", "", "INDICI"),
    "eur":                     ("EUR", "", "INDICI"),
}


def costanti_dataframe(pop_dict):
    """Costruisce il DataFrame mostrato/editato nella scheda Costanti."""
    rows = []
    for key, (label, unit, categoria) in CONST_LABELS.items():
        c = pop_dict[key]
        rows.append({
            "Categoria": categoria, "Costante": label, "Unità": unit,
            "Media Uomini": c["mean_m"], "Dev.Std Uomini": c["sd_m"],
            "Media Donne": c["mean_f"], "Dev.Std Donne": c["sd_f"],
        })
    return pd.DataFrame(rows)


def genera_costanti_xlsx(pop_dict):
    """Esporta le costanti nello stesso formato del file caricabile:
    DATO | UdM | MEDIA | DEV ST | MEDIA2 | DEV ST3, con riga SESSO."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "TABELLA metrics data"
    ws.append(["DATO", "UdM", "MEDIA", "DEV ST", "MEDIA2", "DEV ST3"])
    ws.append(["SESSO", None, "UOMO", "UOMO", "DONNA", "DONNA"])
    for key, (label, unit, _categoria) in CONST_LABELS.items():
        c = pop_dict[key]
        ws.append([label, unit, c["mean_m"], c["sd_m"], c["mean_f"], c["sd_f"]])
    for col in ("A", "C", "D", "E", "F"):
        ws.column_dimensions[col].width = 16
    ws.column_dimensions["A"].width = 30
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


def parse_constants_workbook(file_like):
    """Legge un file Excel nel formato DATO/UdM/MEDIA/DEV ST/MEDIA2/DEV ST3
    e restituisce (updates, unmatched): un dict {key: {mean_m, sd_m, mean_f,
    sd_f}} per le costanti riconosciute, e la lista di righe non riconosciute."""
    wb = openpyxl.load_workbook(file_like, data_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))

    label_to_key = {label.strip().lower(): key for key, (label, _u, _c) in CONST_LABELS.items()}

    header_idx = None
    for i, row in enumerate(rows):
        if row and row[0] and str(row[0]).strip().upper() == "DATO":
            header_idx = i
            break
    if header_idx is None:
        return {}, []

    def to_num(v):
        if isinstance(v, (int, float)):
            return float(v)
        if isinstance(v, str):
            try:
                return float(v.replace(",", "."))
            except ValueError:
                return None
        return None

    updates, unmatched = {}, []
    for row in rows[header_idx + 1:]:
        if not row or row[0] is None:
            continue
        name = str(row[0]).strip()
        if not name or name.upper() == "SESSO":
            continue
        if "DATI DA LETTERATURA" in name.upper():
            break
        key = label_to_key.get(name.lower())
        mean_m = to_num(row[2]) if len(row) > 2 else None
        sd_m = to_num(row[3]) if len(row) > 3 else None
        mean_f = to_num(row[4]) if len(row) > 4 else None
        sd_f = to_num(row[5]) if len(row) > 5 else None
        if key is None:
            unmatched.append(name)
            continue
        if None not in (mean_m, sd_m, mean_f, sd_f):
            updates[key] = dict(mean_m=mean_m, sd_m=abs(sd_m), mean_f=mean_f, sd_f=abs(sd_f))
    return updates, unmatched


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
    dict(key="imtp_peak_force", label="IMTP Peak Force", category="ISOMETRIC PULL TEST",
         jump_type="imtp", raw_var="peak force", unit="N",
         pop_key="imtp_peak_force", lower_is_better=False, kind="score"),
    dict(key="imtp_rel_peak_force", label="IMTP Rel Peak Force", category="ISOMETRIC PULL TEST",
         jump_type="imtp", raw_var=None,
         derive=lambda rep: _safe_div(rep.get("peak force"), rep.get("body mass")),
         unit="N/kg", pop_key="imtp_rel_peak_force", lower_is_better=False, kind="score"),

    dict(key="sj_mean_power", label="SJ Mean Power", category="SQUAT JUMP TEST",
         jump_type="sj", raw_var="avg. propulsive power", unit="W",
         pop_key="sj_mean_power", lower_is_better=False, kind="score"),
    dict(key="sj_height", label="SJ Height", category="SQUAT JUMP TEST",
         jump_type="sj", raw_var="jump height ft", unit="cm",
         pop_key="sj_height", lower_is_better=False, kind="score"),
    dict(key="sj_contraction_time", label="SJ Contraction Time", category="SQUAT JUMP TEST",
         jump_type="sj", raw_var="time to takeoff", unit="s",
         pop_key="sj_contraction_time", lower_is_better=True, kind="score"),
    dict(key="sj_net_impulse", label="SJ Net Impulse", category="SQUAT JUMP TEST",
         jump_type="sj", raw_var="net impulse", unit="N\u00b7s",
         pop_key=None, lower_is_better=False, kind="info"),
    dict(key="sj_net_rel_impulse", label="SJ Net Rel Impulse", category="SQUAT JUMP TEST",
         jump_type="sj", raw_var=None,
         derive=lambda rep: _safe_div(rep.get("net impulse"), rep.get("body mass")),
         unit="N\u00b7s/kg", pop_key=None, lower_is_better=False, kind="info"),

    dict(key="cmj_net_impulse", label="CMJ Net Impulse", category="COUNTERMOVEMENT JUMP TEST",
         jump_type="cmj", raw_var="net impulse", unit="N\u00b7s",
         pop_key=None, lower_is_better=False, kind="info"),
    dict(key="cmj_net_rel_impulse", label="CMJ Net Rel Impulse", category="COUNTERMOVEMENT JUMP TEST",
         jump_type="cmj", raw_var=None,
         derive=lambda rep: _safe_div(rep.get("net impulse"), rep.get("body mass")),
         unit="N\u00b7s/kg", pop_key=None, lower_is_better=False, kind="info"),
    dict(key="cmj_contraction_time", label="CMJ Contraction Time", category="COUNTERMOVEMENT JUMP TEST",
         jump_type="cmj", raw_var="time to takeoff", unit="s",
         pop_key=None, lower_is_better=True, kind="info"),
    dict(key="cmj_height", label="CMJ Height", category="COUNTERMOVEMENT JUMP TEST",
         jump_type="cmj", raw_var="jump height ft", unit="cm",
         pop_key="cmj_height", lower_is_better=False, kind="score"),
    dict(key="mrsi_cmj", label="mRSI-CMJ", category="COUNTERMOVEMENT JUMP TEST",
         jump_type="cmj", raw_var="rsi modified", unit="m/s",
         pop_key="mrsi_cmj", lower_is_better=False, kind="score"),
    dict(key="cmj_peak_force", label="CMJ Peak Force", category="COUNTERMOVEMENT JUMP TEST",
         jump_type="cmj", raw_var="peak propulsive force", unit="N",
         pop_key=None, lower_is_better=False, kind="info"),

    dict(key="cmj_re_initial_height", label="CMJ RE Jump Height (iniziale)", category="COUNTERMOVEMENT JUMP REBOUND TEST",
         jump_type="cmrj", raw_var="jump height ft", unit="cm",
         pop_key=None, lower_is_better=False, kind="info"),
    dict(key="cmj_re_rebound_height", label="CMJ RE Rebound Jump Height", category="COUNTERMOVEMENT JUMP REBOUND TEST",
         jump_type="cmrj", raw_var="rebound jump height ft", unit="cm",
         pop_key="cmj_re_rebound_height", lower_is_better=False, kind="score"),
    dict(key="cmj_re_contact_time", label="CMJ RE Contact Time", category="COUNTERMOVEMENT JUMP REBOUND TEST",
         jump_type="cmrj", raw_var="rebound contact time", unit="s",
         pop_key="cmj_re_contact_time", lower_is_better=True, kind="score"),
    dict(key="cmj_re_rebound_impulse", label="CMJ RE Rebound Propulsive Impulse", category="COUNTERMOVEMENT JUMP REBOUND TEST",
         jump_type="cmrj", raw_var="rebound propulsive impulse", unit="N\u00b7s",
         pop_key="cmj_re_rebound_impulse", lower_is_better=False, kind="score"),
    dict(key="mrsi_cmj_re", label="mRSI-CMJ RE", category="COUNTERMOVEMENT JUMP REBOUND TEST",
         jump_type="cmrj", raw_var="rebound rsi modified", unit="m/s",
         pop_key="mrsi_cmj_re", lower_is_better=False, kind="score"),
    dict(key="unbalanced_landing_raw", label="Braking Impulse Sym. Index (CMJ RE)", category="COUNTERMOVEMENT JUMP REBOUND TEST",
         jump_type="cmrj", raw_var="braking impulse sym. index", unit="%",
         pop_key=None, lower_is_better=False, kind="info"),

    # Indici derivati da medie aggregate cross-test (non per-rep)
    dict(key="dsi", label="DSI (Dynamic Strength Index - Peak force CMJ/Peak force IMTP)", category="INDICI",
         jump_type=None, raw_var=None, unit="",
         pop_key="dsi", lower_is_better=False, kind="score_single"),
    dict(key="eur", label="EUR (Eccentric Utilisation Ratio - CMJ Height / SJ Height)", category="INDICI",
         jump_type=None, raw_var=None, unit="",
         pop_key="eur", lower_is_better=False, kind="score_single"),
]

# Controlli a soglia (equivalenti a "Jump to rebound Ratio", "CMJ to CMJ RE
# check", "Unbalanced landing check" del foglio originale, più il controllo
# sul tempo di contatto). Riguardano esclusivamente il CMJ Rebound, quindi
# vengono mostrati insieme alla categoria COUNTERMOVEMENT JUMP REBOUND TEST
# nel Dettaglio Test.
# scale/decimals/suffix controllano solo la formattazione a schermo: il
# "value" resta sempre nell'unità nativa del calcolo (rapporto puro per i
# primi tre, secondi per il tempo di contatto).
CHECKS = [
    dict(key="jump_to_rebound_ratio", label="Jump to Rebound Ratio",
         desc="Rapporto tra altezza del rimbalzo e altezza del salto iniziale nel CMJ RE.",
         threshold=0.60, direction="min", scale=100, decimals=1, suffix="%"),
    dict(key="cmj_to_cmjre_check", label="CMJ to CMJ RE Check",
         desc="Rapporto tra l'altezza del salto iniziale nel CMJ RE e l'altezza del CMJ standard.",
         threshold=0.85, direction="min", scale=100, decimals=1, suffix="%"),
    dict(key="unbalanced_landing_check", label="Unbalanced Landing Check",
         desc="Indice di simmetria dell'impulso frenante in atterraggio (CMJ RE). Valori assoluti alti indicano un atterraggio sbilanciato.",
         threshold=0.50, direction="max", scale=100, decimals=1, suffix="%"),
    dict(key="cmj_re_contact_time_check", label="CMJ RE Contact Time Check",
         desc="Tempo di contatto nel rimbalzo del CMJ RE: un tempo di contatto troppo lungo indica un rimbalzo poco reattivo.",
         threshold=0.250, direction="max", scale=1000, decimals=0, suffix=" ms"),
]

JUMP_TYPE_LABELS = {"imtp": "IMTP", "sj": "Squat Jump", "cmj": "CMJ", "cmrj": "CMJ Rebound"}


# ============================================================================
# PARTE 1bis — CATALOGO METRICHE EXTRA (ricerca libera, senza dato di
# popolazione/T-score)
# ============================================================================
# L'app mostra di default solo un set curato di metriche (vedi METRICS
# sopra), ciascuna con confronto di popolazione. Questi elenchi coprono
# TUTTE le variabili che ForceMate/ForceDecks può esportare per ciascun
# tipo di test, così l'utente può cercare e aggiungere qualsiasi altra
# metrica dalla scheda Dettaglio Test — senza T-score (nessun dato di
# popolazione disponibile per queste), solo media/dev.std/CV% per
# ripetizione, con le stesse esclusioni already impostate sui salti.
# Nomi già lowercase per combaciare con le chiavi salvate in rep["vars"]
# (il parsing normalizza sempre con .strip().lower()).
_BASE_JUMP_METRICS_CATALOG = [
    'avg braking force sym. index', 'avg braking power sym. index', 'avg eccentric force',
    'avg eccentric force sym. index', 'avg eccentric power', 'avg eccentric power sym. index',
    'avg eccentric velocity', 'avg landing rfd', 'avg landing rfd sym. index',
    'avg propulsive force sym. index', 'avg propulsive power sym. index', 'avg. braking force',
    'avg. braking power', 'avg. braking velocity', 'avg. propulsive force',
    'avg. propulsive power', 'avg. propulsive velocity', 'avg. rfd', 'avg. rfd sym. index',
    'body mass', 'body weight', 'body weight sd', 'braking duration', 'braking end time',
    'braking impulse', 'braking impulse sym. index', 'braking rfd', 'braking rfd sym. index',
    'braking work', 'braking work sym. index', 'contact time', 'decel rfd', 'decel rfd sym. index',
    'displacement depth', 'eccentric impulse', 'eccentric impulse sym. index', 'eccentric rfd',
    'eccentric rfd sym. index', 'eccentric work', 'eccentric work sym. index', 'flight threshold',
    'flight time', 'force at min displacement', 'force peak power', 'initiation threshold',
    'jump height ft', 'jump height ni', 'jump momentum', 'jump start time', 'jump threshold time',
    'landing peak force time', 'landing rfd 0-20ms', 'landing rfd 0-20ms sym. index',
    'landing rfd 0-40ms', 'landing rfd 0-40ms sym. index', 'landing rfd 0-60ms',
    'landing rfd 0-60ms sym. index', 'landing rfd 0-80ms', 'landing rfd 0-80ms sym. index',
    'landing time', 'left avg braking force', 'left avg braking power', 'left avg eccentric force',
    'left avg eccentric power', 'left avg landing rfd', 'left avg propulsive force',
    'left avg propulsive power', 'left avg rfd', 'left braking impulse', 'left braking rfd',
    'left braking work', 'left decel rfd', 'left eccentric impulse', 'left eccentric rfd',
    'left eccentric work', 'left landing rfd 0-20ms', 'left landing rfd 0-40ms',
    'left landing rfd 0-60ms', 'left landing rfd 0-80ms', 'left net impulse', 'left p1 impulse',
    'left p2 impulse', 'left peak braking force', 'left peak braking power',
    'left peak eccentric force', 'left peak eccentric power', 'left peak force',
    'left peak landing force', 'left peak propulsive force', 'left peak propulsive power',
    'left propulsive impulse', 'left propulsive rfd', 'left propulsive work',
    'left time to peak braking force', 'left time to peak braking power',
    'left time to peak eccentric force', 'left time to peak eccentric power',
    'left time to peak force', 'left time to peak landing force', 'left time to peak power',
    'left time to peak propulsive force', 'left time to peak propulsive power',
    'min braking velocity', 'min eccentric velocity', 'min unweight force', 'net impulse',
    'net impulse sym. index', 'p1 avg force', 'p1 avg power', 'p1 avg velocity', 'p1 duration',
    'p1 impulse', 'p1 impulse sym. index', 'p1 p2 duration ratio', 'p1 p2 force ratio',
    'p1 p2 power ratio', 'p1 p2 velocity ratio', 'p1 peak force', 'p1 peak power',
    'p1 peak velocity', 'p2 avg force', 'p2 avg power', 'p2 avg velocity', 'p2 duration',
    'p2 impulse', 'p2 impulse sym. index', 'p2 peak force', 'p2 peak power', 'p2 peak velocity',
    'peak braking force', 'peak braking force sym. index', 'peak braking power',
    'peak braking power sym. index', 'peak eccentric force', 'peak eccentric force sym. index',
    'peak eccentric power', 'peak eccentric power sym. index', 'peak force',
    'peak force sym. index', 'peak force time', 'peak landing force',
    'peak landing force sym. index', 'peak power', 'peak propulsive force',
    'peak propulsive force sym. index', 'peak propulsive power',
    'peak propulsive power sym. index', 'peak propulsive velocity', 'peak velocity',
    'propulsive duration', 'propulsive impulse', 'propulsive impulse sym. index', 'propulsive rfd',
    'propulsive rfd sym. index', 'propulsive start time', 'propulsive work',
    'propulsive work sym. index', 'rel. avg. propulsive force', 'rel. min unweight force',
    'rel. propulsive impulse', 'relative force at min displacement', 'relative peak force',
    'relative peak landing force', 'relative peak power', 'right avg braking force',
    'right avg braking power', 'right avg eccentric force', 'right avg eccentric power',
    'right avg landing rfd', 'right avg propulsive force', 'right avg propulsive power',
    'right avg rfd', 'right braking impulse', 'right braking rfd', 'right braking work',
    'right decel rfd', 'right eccentric impulse', 'right eccentric rfd', 'right eccentric work',
    'right landing rfd 0-20ms', 'right landing rfd 0-40ms', 'right landing rfd 0-60ms',
    'right landing rfd 0-80ms', 'right net impulse', 'right p1 impulse', 'right p2 impulse',
    'right peak braking force', 'right peak braking power', 'right peak eccentric force',
    'right peak eccentric power', 'right peak force', 'right peak landing force',
    'right peak propulsive force', 'right peak propulsive power', 'right propulsive impulse',
    'right propulsive rfd', 'right propulsive work', 'right time to peak braking force',
    'right time to peak braking power', 'right time to peak eccentric force',
    'right time to peak eccentric power', 'right time to peak force',
    'right time to peak landing force', 'right time to peak power',
    'right time to peak propulsive force', 'right time to peak propulsive power', 'rsi',
    'rsi exponential', 'rsi modified', 'takeoff time', 'takeoff velocity',
    'time to peak braking force', 'time to peak braking force sym. index',
    'time to peak braking power', 'time to peak braking power sym. index',
    'time to peak eccentric force', 'time to peak eccentric force sym. index',
    'time to peak eccentric power', 'time to peak eccentric power sym. index',
    'time to peak force', 'time to peak landing force', 'time to peak landing force sym. index',
    'time to peak power', 'time to peak power sym. index', 'time to peak propulsive force',
    'time to peak propulsive force sym. index', 'time to peak propulsive power',
    'time to peak propulsive power sym. index', 'time to peak si', 'time to takeoff',
    'unweighted duration', 'velocity peak power',
]

_CMRJ_EXTRA_METRICS_CATALOG = [
    'left rebound avg braking force', 'left rebound avg braking power',
    'left rebound avg propulsive force', 'left rebound avg propulsive power',
    'left rebound braking impulse', 'left rebound braking rfd', 'left rebound braking work',
    'left rebound p1 impulse', 'left rebound p2 impulse', 'left rebound peak braking force',
    'left rebound peak braking power', 'left rebound peak propulsive force',
    'left rebound peak propulsive power', 'left rebound propulsive impulse',
    'left rebound propulsive rfd', 'left rebound propulsive work',
    'left rebound time to peak braking force', 'left rebound time to peak braking power',
    'left rebound time to peak propulsive force', 'left rebound time to peak propulsive power',
    'rebound avg braking force', 'rebound avg braking force symmetry index',
    'rebound avg braking power', 'rebound avg braking power symmetry index',
    'rebound avg braking velocity', 'rebound avg propulsive force',
    'rebound avg propulsive force symmetry index', 'rebound avg propulsive power',
    'rebound avg propulsive power symmetry index', 'rebound avg propulsive velocity',
    'rebound braking impulse', 'rebound braking impulse symmetry index', 'rebound braking rfd',
    'rebound braking rfd symmetry index', 'rebound braking work',
    'rebound braking work symmetry index', 'rebound contact time', 'rebound flight time',
    'rebound force at min displacement', 'rebound jump height ft', 'rebound jump momentum',
    'rebound min braking velocity', 'rebound p1 avg force', 'rebound p1 avg power',
    'rebound p1 avg velocity', 'rebound p1 duration', 'rebound p1 impulse',
    'rebound p1 impulse symmetry index', 'rebound p1 p2 duration ratio',
    'rebound p1 p2 force ratio', 'rebound p1 p2 power ratio', 'rebound p1 p2 velocity ratio',
    'rebound p1 peak force', 'rebound p1 peak power', 'rebound p1 peak velocity',
    'rebound p2 avg force', 'rebound p2 avg power', 'rebound p2 avg velocity',
    'rebound p2 duration', 'rebound p2 impulse', 'rebound p2 impulse symmetry index',
    'rebound p2 peak force', 'rebound p2 peak power', 'rebound p2 peak velocity',
    'rebound peak braking force', 'rebound peak braking force symmetry index',
    'rebound peak braking power', 'rebound peak braking power symmetry index',
    'rebound peak propulsive force', 'rebound peak propulsive force symmetry index',
    'rebound peak propulsive power', 'rebound peak propulsive power symmetry index',
    'rebound peak propulsive velocity', 'rebound propulsive impulse',
    'rebound propulsive impulse symmetry index', 'rebound propulsive rfd',
    'rebound propulsive rfd symmetry index', 'rebound propulsive work',
    'rebound propulsive work symmetry index', 'rebound relative force at min displacement',
    'rebound rsi', 'rebound rsi modified', 'rebound time to peak braking force',
    'rebound time to peak braking force symmetry index', 'rebound time to peak braking power',
    'rebound time to peak braking power symmetry index', 'rebound time to peak propulsive force',
    'rebound time to peak propulsive force symmetry index',
    'rebound time to peak propulsive power',
    'rebound time to peak propulsive power symmetry index', 'rebound time to takeoff',
    'right rebound avg braking force', 'right rebound avg braking power',
    'right rebound avg propulsive force', 'right rebound avg propulsive power',
    'right rebound braking impulse', 'right rebound braking rfd', 'right rebound braking work',
    'right rebound p1 impulse', 'right rebound p2 impulse', 'right rebound peak braking force',
    'right rebound peak braking power', 'right rebound peak propulsive force',
    'right rebound peak propulsive power', 'right rebound propulsive impulse',
    'right rebound propulsive rfd', 'right rebound propulsive work',
    'right rebound time to peak braking force', 'right rebound time to peak braking power',
    'right rebound time to peak propulsive force', 'right rebound time to peak propulsive power',
    'with armswing',
]

_IMTP_METRICS_CATALOG = [
    'end of rise force', 'end of rise force net', 'end of rise force relative',
    'end of rise force relative net', 'end of rise torque', 'force 100ms', 'force 100ms left',
    'force 100ms left net', 'force 100ms left relative', 'force 100ms left relative net',
    'force 100ms net', 'force 100ms relative', 'force 100ms relative net', 'force 100ms right',
    'force 100ms right net', 'force 100ms right relative', 'force 100ms right relative net',
    'force 100ms si', 'force 150ms', 'force 150ms left', 'force 150ms left net',
    'force 150ms left relative', 'force 150ms left relative net', 'force 150ms net',
    'force 150ms relative', 'force 150ms relative net', 'force 150ms right',
    'force 150ms right net', 'force 150ms right relative', 'force 150ms right relative net',
    'force 150ms si', 'force 200ms', 'force 200ms left', 'force 200ms left net',
    'force 200ms left relative', 'force 200ms left relative net', 'force 200ms net',
    'force 200ms relative', 'force 200ms relative net', 'force 200ms right',
    'force 200ms right net', 'force 200ms right relative', 'force 200ms right relative net',
    'force 200ms si', 'force 250ms', 'force 250ms left', 'force 250ms left net',
    'force 250ms left relative', 'force 250ms left relative net', 'force 250ms net',
    'force 250ms relative', 'force 250ms relative net', 'force 250ms right',
    'force 250ms right net', 'force 250ms right relative', 'force 250ms right relative net',
    'force 250ms si', 'force 50ms', 'force 50ms left', 'force 50ms left net',
    'force 50ms left relative', 'force 50ms left relative net', 'force 50ms net',
    'force 50ms relative', 'force 50ms relative net', 'force 50ms right', 'force 50ms right net',
    'force 50ms right relative', 'force 50ms right relative net', 'force 50ms si',
    'force at max rfd', 'force at max rfd left', 'force at max rfd left net',
    'force at max rfd left relative', 'force at max rfd left relative net', 'force at max rfd net',
    'force at max rfd relative', 'force at max rfd relative net', 'force at max rfd right',
    'force at max rfd right net', 'force at max rfd right relative',
    'force at max rfd right relative net', 'force at max rfd si', 'impulse 100ms', 'impulse 150ms',
    'impulse 200ms', 'impulse 250ms', 'impulse 50ms', 'impulse at max rfd', 'max rfd',
    'onset force', 'onset force net', 'onset force relative', 'onset force relative net',
    'onset time', 'onset torque', 'peak force', 'peak force left', 'peak force left net',
    'peak force left relative', 'peak force left relative net', 'peak force net',
    'peak force relative', 'peak force relative net', 'peak force right', 'peak force right net',
    'peak force right relative', 'peak force right relative net', 'peak force si', 'peak torque',
    'peak torque left', 'peak torque right', 'rfd 100ms', 'rfd 150ms', 'rfd 200ms', 'rfd 250ms',
    'rfd 50ms', 'rfd impulse', 'steadiness rsme force', 'steadiness rsme rfd', 'system weight',
    'time to force end of rise', 'time to peak force', 'time to peak rfd', 'torque 100ms',
    'torque 100ms left', 'torque 100ms right', 'torque 150ms', 'torque 150ms left',
    'torque 150ms right', 'torque 200ms', 'torque 200ms left', 'torque 200ms right',
    'torque 250ms', 'torque 250ms left', 'torque 250ms right', 'torque 50ms', 'torque 50ms left',
    'torque 50ms right', 'torque at max rfd', 'torque at max rfd left', 'torque at max rfd right',
]

# Elenco per jump_type usato dal multiselect di ricerca. sj/cmj condividono
# lo stesso set base; cmrj lo estende con le variabili "rebound ...".
EXTRA_METRICS_CATALOG = {
    "sj": sorted(_BASE_JUMP_METRICS_CATALOG),
    "cmj": sorted(_BASE_JUMP_METRICS_CATALOG),
    "cmrj": sorted(set(_BASE_JUMP_METRICS_CATALOG) | set(_CMRJ_EXTRA_METRICS_CATALOG)),
    "imtp": sorted(_IMTP_METRICS_CATALOG),
}


def extra_metric_label(raw_var):
    """Etichetta leggibile per una metrica extra scelta dall'utente (solo
    la prima lettera maiuscola, il resto è già in minuscolo)."""
    return raw_var[:1].upper() + raw_var[1:]


# ============================================================================
# PARTE 2 — CARICAMENTO FILE (sidebar upload)
# ============================================================================

st.set_page_config(page_title="Force Plate Test Report", page_icon="🏋️", layout="wide")

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

st.sidebar.title("📂 Import dati")

uploaded = st.sidebar.file_uploader(
    "Carica i file XLSX esportati da ForceMate", type=["xlsx"], accept_multiple_files=True
)

st.sidebar.markdown("---")

csv_reload = st.sidebar.file_uploader(
    "🔁 Ricarica CSV esportato (Dettaglio Test)", type=["csv"],
    help="Carica un CSV precedentemente esportato dalla scheda 'Dettaglio Test' per rivedere "
         "grafici e tabelle senza dover ricaricare i file XLSX originali. Se presente, ha "
         "precedenza sui file XLSX caricati sotto.",
)


# Le costanti di popolazione (scheda "⚙️ Costanti") non dipendono dai file
# caricati: inizializziamo subito lo stato in modo che quella scheda sia
# consultabile/modificabile anche prima di caricare qualsiasi file.
if "pop" not in st.session_state:
    st.session_state["pop"] = {k: dict(v) for k, v in DEFAULT_POP.items()}


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


def parse_dettaglio_csv(file_like) -> ParsedFile:
    """Ricostruisce un ParsedFile sintetico a partire da un CSV esportato dalla
    scheda Dettaglio Test (colonne Categoria/Metrica/Prova N/...), per poter
    rivedere grafici e tabelle senza i file XLSX originali. I valori "Rel"
    (dipendenti dal peso corporeo, non presente nel CSV) vengono ricostruiti
    retro-calcolando un peso corporeo equivalente dal valore assoluto e dal
    rapporto già esportato."""
    df = pd.read_csv(file_like, encoding="utf-8-sig")
    if not {"Categoria", "Metrica"}.issubset(df.columns):
        raise ValueError("Il file non sembra un export di un precedente test (colonne mancanti).")

    prova_cols = sorted(
        (c for c in df.columns if re.fullmatch(r"Prova \d+", str(c))),
        key=lambda c: int(c.split(" ")[1]),
    )

    nome = str(df["Nome"].iloc[0]) if "Nome" in df.columns and len(df) else None
    sesso = str(df["Sesso"].iloc[0]) if "Sesso" in df.columns and len(df) else None
    periodo = str(df["Data test"].iloc[0]) if "Data test" in df.columns and len(df) else None

    reps = []
    for cat in CATEGORIES:
        cat_df = df[df["Categoria"] == cat]
        if cat_df.empty:
            continue
        cat_metrics = [m for m in METRICS if m["category"] == cat and m.get("jump_type")]
        if not cat_metrics:
            continue
        jump_type = cat_metrics[0]["jump_type"]
        label_to_metric = {m["label"]: m for m in cat_metrics}

        n_reps_cat = 0
        for _, row in cat_df.iterrows():
            cnt = sum(pd.notna(row.get(pc)) for pc in prova_cols)
            n_reps_cat = max(n_reps_cat, cnt)

        for i in range(n_reps_cat):
            pc = f"Prova {i + 1}"
            rep_vars, derive_targets = {}, {}
            for _, row in cat_df.iterrows():
                m = label_to_metric.get(row.get("Metrica"))
                if m is None:
                    continue
                val = row.get(pc)
                if pd.isna(val):
                    continue
                val = float(val)
                if m.get("raw_var"):
                    rep_vars[m["raw_var"]] = val
                elif m.get("derive"):
                    derive_targets[m["key"]] = val

            # Le metriche "Rel" (per kg) derivano da raw_var / body_mass: non
            # avendo il peso corporeo nel CSV, lo ricaviamo a ritroso dal
            # valore assoluto già noto e dal rapporto già esportato.
            for key, target in derive_targets.items():
                base = None
                if key == "imtp_rel_peak_force":
                    base = rep_vars.get("peak force")
                elif key in ("sj_net_rel_impulse", "cmj_net_rel_impulse"):
                    base = rep_vars.get("net impulse")
                if base is not None and target:
                    rep_vars["body mass"] = base / target

            reps.append({"jump_type": jump_type, "vars": rep_vars, "units": {}})

    metadata = {
        "nome": nome, "sesso": sesso, "altezza_cm": None, "peso_kg_input": None,
        "data_test": None, "device": None, "team": None,
        "test_period": None, "test_type": None, "periodo_override": periodo,
    }
    return ParsedFile(filename="(da CSV)", metadata=metadata, reps=reps)


parsed_files = []
if csv_reload is not None:
    try:
        parsed_files = [parse_dettaglio_csv(io.BytesIO(csv_reload.getvalue()))]
        st.sidebar.success(f"Dati ricaricati da '{csv_reload.name}'.")
        if uploaded:
            st.sidebar.caption("File XLSX caricati sotto ignorati: è attivo il CSV ricaricato sopra.")
    except Exception as e:
        st.sidebar.error(f"Impossibile leggere il CSV: {e}")
elif uploaded:
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


def build_results(files, pop_dict):
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
                pop = pop_dict[metric["pop_key"]]
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

    # EUR = CMJ height / SJ height (rapporto puro), coerente con la tabella
    # costanti fornita dall'utente (~1,11 uomini, ~1,09 donne).
    cmj_h = support.get("cmj_height", {}).get("mean")
    sj_h = support.get("sj_height", {}).get("mean")
    eur_val = (cmj_h / sj_h) if (cmj_h and sj_h) else None

    for key, val in (("dsi", dsi_val), ("eur", eur_val)):
        pop = pop_dict[key]
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
    cmj_re_contact_time = support.get("cmj_re_contact_time", {}).get("mean")

    checks_out = [
        make_check(CHECKS[0], (rebound_h / initial_h) if (initial_h and rebound_h) else None),
        make_check(CHECKS[1], (initial_h / cmj_h_standalone) if (initial_h and cmj_h_standalone) else None),
        make_check(CHECKS[2], (landing_sym / 100.0) if landing_sym is not None else None),
        make_check(CHECKS[3], cmj_re_contact_time if cmj_re_contact_time is not None else None),
    ]

    return dict(results=results, support=support, checks=checks_out, sesso=sesso_raw)


def profilo_forza(results):
    out = {}
    for cat in CATEGORIES:
        ts = [r["t"] for r in results if r["category"] == cat and r["t"] is not None]
        out[cat] = (sum(ts) / len(ts)) if ts else None
    return out


# ============================================================================
# PARTE 4bis — GRAFICI A QUADRANTI (RSQ / EUR quadrant plot)
# ============================================================================
# Incrociano due metriche (es. altezza salto e tempo di contatto, oppure
# altezza SJ e altezza CMJ) mettendo in relazione la media del test rispetto
# alla media di popolazione, così da leggere non solo "quanto" ma "come"
# l'atleta esprime la prestazione. Il crosshair è sempre centrato sulla
# media di popolazione (x, y); il punto mostrato è la media del test con
# barre d'errore pari alla deviazione standard delle ripetizioni incluse.

def quadrant_chart(x_mean, y_mean, x_sd, y_sd, x_label, y_label, quadrant_defs,
                    pop_x=None, pop_y=None, point_color=PRIMARY, diagonal=False,
                    diagonal_ratio=1.0, height=430):
    """quadrant_defs: dict con chiavi 'tl','tr','bl','br' -> (etichetta, colore).
    Il crosshair (linee tratteggiate e confini dei quadranti) è centrato sulla
    media di popolazione (pop_x, pop_y); se non disponibile per una metrica si
    usa la media del test come fallback. Il punto mostrato è la media del test
    con barre d'errore pari alla deviazione standard delle ripetizioni incluse.
    Se diagonal=True, la linea tratteggiata diagonale rappresenta i punti con
    rapporto x/y == diagonal_ratio (es. il valore medio di popolazione
    dell'indice, non necessariamente 1.0).
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

    if diagonal:
        ratio = diagonal_ratio if diagonal_ratio else 1.0
        # Linea dei punti con x/y == ratio (es. rapporto medio di
        # popolazione), tracciata sull'intera larghezza del grafico:
        # y = x / ratio. Plotly ritaglia automaticamente la parte che
        # eccede il range visibile degli assi.
        fig.add_shape(type="line", x0=x0, y0=x0 / ratio, x1=x1, y1=x1 / ratio,
                       line=dict(color=ACCENT, dash="dash", width=2))

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


RSQ_QUADRANTS = dict(
    tl=("Alta Reattività", "#4FC3F7"),
    tr=("Forza-Dominante", "#EF5350"),
    bl=("Elastico-Dominante", "#66BB6A"),
    br=("Bassa Reattività", "#FFEE58"),
)

EUR_QUADRANTS = dict(
    tl=("Atleta potente ma poco esplosivo (EUR < 1.0)", "#FFB74D"),
    tr=("Atleta potente ed esplosivo (EUR ~ 1.0)", "#81C784"),
    bl=("Atleta poco potente e poco esplosivo", "#E57373"),
    br=("Atleta esplosivo ma poco potente (EUR > 1.0)", "#FFF176"),
)

DSI_QUADRANTS = dict(
    tl=("Alta prestaz. IMTP, bassa CMJ (DSI < 1.0)", "#FFB74D"),
    tr=("Alta prestazione in CMJ e IMTP (DSI ~ 1.0)", "#81C784"),
    bl=("Bassa prestazione in CMJ e IMTP", "#E57373"),
    br=("Alta prestaz. CMJ, bassa IMTP (DSI > 1.0)", "#FFF176"),
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


def build_eur_chart(results):
    """Grafico EUR (altezza SJ vs altezza CMJ), se entrambe le medie sono
    disponibili. La diagonale rappresenta il valore medio di popolazione
    dell'EUR (non un fisso 1.0). Ritorna None altrimenti."""
    r_sj_h = next((r for r in results if r["key"] == "sj_height"), None)
    r_cmj_h = next((r for r in results if r["key"] == "cmj_height"), None)
    r_eur = next((r for r in results if r["key"] == "eur"), None)
    if r_sj_h is None or r_cmj_h is None:
        return None
    eur_pop = r_eur["pop_mean"] if r_eur and r_eur["pop_mean"] else 1.0
    return quadrant_chart(
        x_mean=r_cmj_h["mean"], y_mean=r_sj_h["mean"], x_sd=r_cmj_h["sd"], y_sd=r_sj_h["sd"],
        x_label="CMJ Height (cm)", y_label="SJ Height (cm)",
        quadrant_defs=EUR_QUADRANTS, pop_x=r_cmj_h["pop_mean"], pop_y=r_sj_h["pop_mean"],
        point_color=PRIMARY, diagonal=True, diagonal_ratio=eur_pop,
    )


def build_dsi_chart(results):
    """Grafico DSI (CMJ Peak Force vs IMTP Peak Force): mostra la posizione
    dell'atleta nei quattro quadranti senza tracciare una diagonale di
    riferimento (le due forze di picco non sono attese avere la stessa
    scala). Ritorna None se una delle due medie non è disponibile."""
    r_cmj_peak = next((r for r in results if r["key"] == "cmj_peak_force"), None)
    r_imtp_peak = next((r for r in results if r["key"] == "imtp_peak_force"), None)
    if r_cmj_peak is None or r_imtp_peak is None:
        return None
    return quadrant_chart(
        x_mean=r_cmj_peak["mean"], y_mean=r_imtp_peak["mean"], x_sd=r_cmj_peak["sd"], y_sd=r_imtp_peak["sd"],
        x_label="CMJ Peak Force (N)", y_label="IMTP Peak Force (N)",
        quadrant_defs=DSI_QUADRANTS, pop_x=r_cmj_peak["pop_mean"], pop_y=r_imtp_peak["pop_mean"],
        point_color=PRIMARY, diagonal=False,
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


def _wrap_label(text, width=16):
    """Spezza un'etichetta lunga su più righe (per gli assi angolari del
    radar), così non deborda oltre il bordo del grafico invece di essere
    tagliata."""
    return "<br>".join(textwrap.wrap(text, width=width, break_long_words=False))


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


if parsed_files:
    meta0 = parsed_files[0].metadata
    nome = meta0.get("nome") or "Atleta"
    sesso = sesso_da_file(parsed_files) or "-"
    if meta0.get("periodo_override"):
        periodo = meta0["periodo_override"]
    else:
        date_tests = [pf.metadata.get("data_test") for pf in parsed_files if pf.metadata.get("data_test")]
        data_min = min(date_tests).strftime("%d/%m/%Y") if date_tests else "-"
        data_max = max(date_tests).strftime("%d/%m/%Y") if date_tests else "-"
        periodo = data_min if data_min == data_max else f"{data_min} → {data_max}"
else:
    nome, sesso, periodo = "Atleta", "-", "-"


# ============================================================================
# PARTE 5 — REPORT LIVE (interfaccia a schede)
# ============================================================================

if parsed_files:
    st.title(f"🏋️ Force Plate Test Report — {nome}")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Sesso", sesso)
    c2.metric("Data test", periodo)
    c3.metric("File caricati", len(parsed_files))
    c4.metric("Ripetizioni totali", sum(len(pf.reps) for pf in parsed_files))
else:
    st.title("🏋️ Force Plate Test Report")
    st.info(
        "Carica uno o più file XLSX esportati da ForceMate (IMTP, SJ, CMJ, CMJ RE) dalla "
        "barra laterale per vedere il dettaglio dei test, il profilo di forza e generare il "
        "report. Nel frattempo puoi consultare e modificare le costanti di popolazione nella "
        "scheda '⚙️ Costanti'."
    )

tab_costanti, tab_dettaglio, tab_profilo, tab_report = st.tabs(
    ["⚙️ Costanti", "🔍 Dettaglio Test", "📊 Profilo di Forza", "📄 Report"]
)

with tab_costanti:
    st.markdown(
        "Valori di riferimento della popolazione (media e deviazione standard, per uomini e "
        "donne) usati per calcolare i T-score. Modificabili direttamente nella tabella; le "
        "modifiche si applicano subito alle altre schede."
    )

    df_pop = costanti_dataframe(st.session_state["pop"])
    edited_pop = st.data_editor(
        df_pop, use_container_width=True, hide_index=True, key="pop_editor",
        column_config={
            "Categoria": st.column_config.TextColumn(disabled=True),
            "Costante": st.column_config.TextColumn(disabled=True),
            "Unità": st.column_config.TextColumn(disabled=True),
            "Media Uomini": st.column_config.NumberColumn(format="%.4f"),
            "Dev.Std Uomini": st.column_config.NumberColumn(format="%.4f", min_value=0.0),
            "Media Donne": st.column_config.NumberColumn(format="%.4f"),
            "Dev.Std Donne": st.column_config.NumberColumn(format="%.4f", min_value=0.0),
        },
    )
    for key, row in zip(CONST_LABELS.keys(), edited_pop.to_dict("records")):
        mm, sm, mf, sf = row["Media Uomini"], row["Dev.Std Uomini"], row["Media Donne"], row["Dev.Std Donne"]
        if None not in (mm, sm, mf, sf):
            st.session_state["pop"][key] = dict(mean_m=float(mm), sd_m=float(sm), mean_f=float(mf), sd_f=float(sf))

    st.markdown("---")
    col_dl, col_ul, col_reset = st.columns([1, 1, 1])
    with col_dl:
        st.markdown("**Scarica costanti**")
        xlsx_bytes = genera_costanti_xlsx(st.session_state["pop"])
        st.download_button(
            "⬇️ Scarica (.xlsx)", data=xlsx_bytes, file_name="costanti_force_plate.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with col_ul:
        st.markdown("**Ricarica da file**")
        const_file = st.file_uploader(
            "Formato: DATO / UdM / MEDIA / DEV ST / MEDIA2 / DEV ST3", type=["xlsx"], key="const_uploader"
        )
        if const_file is not None:
            updates, unmatched = parse_constants_workbook(io.BytesIO(const_file.getvalue()))
            if updates:
                st.session_state["pop"].update(updates)
                if "pop_editor" in st.session_state:
                    del st.session_state["pop_editor"]
                st.success(f"Aggiornate {len(updates)} costanti da '{const_file.name}'.")
                if unmatched:
                    st.caption("Righe non riconosciute: " + ", ".join(unmatched))
                st.rerun()
            else:
                st.warning("Nessuna costante riconosciuta nel file caricato. Verifica che segua il formato indicato.")
    with col_reset:
        st.markdown("**Ripristina**")
        if st.button("↺ Valori di default"):
            st.session_state["pop"] = {k: dict(v) for k, v in DEFAULT_POP.items()}
            if "pop_editor" in st.session_state:
                del st.session_state["pop_editor"]
            st.rerun()

if parsed_files:
    results_bundle = build_results(parsed_files, st.session_state["pop"])
    results = results_bundle["results"]
    checks = results_bundle["checks"]
    profilo = profilo_forza(results)
else:
    results, checks, profilo = [], [], {}

with tab_dettaglio:
    if not parsed_files:
        st.info("Carica dei file dalla barra laterale per vedere il dettaglio dei test.")
    else:
        export_rows = []
        for cat in CATEGORIES:
            cat_results = [r for r in results if r["category"] == cat and r["mean"] is not None]
            if not cat_results:
                continue
            st.markdown(f"### {cat}")

            # Grafico a barre orizzontali che partono da sinistra (0): la
            # lunghezza della barra è il T-score assoluto. Le bande di
            # valutazione sono mostrate come contesto sullo sfondo (vedi
            # BANDS) e la media di popolazione (T-score = 50) resta segnata
            # da una linea tratteggiata di riferimento.
            tscore_fig = build_tscore_bar_chart(cat_results)
            if tscore_fig:
                st.plotly_chart(tscore_fig, use_container_width=True)

            # --- Salti individuali: inclusione e valori per colonna ---
            cat_metrics = [m for m in METRICS if m["category"] == cat and m.get("jump_type")]
            jump_type = cat_metrics[0]["jump_type"] if cat_metrics else None
            reps_all = collect_reps_all(parsed_files, jump_type) if jump_type else []
            n_reps = len(reps_all)

            if n_reps == 0:
                st.info("Nessuna ripetizione disponibile per questa categoria.")
                st.markdown("---")
                continue

            # --- Metriche extra: ricerca libera su tutte le altre variabili
            # presenti nel file ForceDecks per questo tipo di test, oltre a
            # quelle curate sopra. Non avendo un dato di popolazione non
            # hanno T-score, ma vengono comunque calcolate (media, dev.std,
            # CV%) sulle stesse ripetizioni incluse/escluse qui sotto.
            existing_raw_vars = {m["raw_var"] for m in cat_metrics if m.get("raw_var")}
            catalog_options = [v for v in EXTRA_METRICS_CATALOG.get(jump_type, []) if v not in existing_raw_vars]
            extra_selected = st.multiselect(
                "🔎 Aggiungi metriche extra (cerca per nome)", options=catalog_options,
                key=f"extra_metrics_{jump_type}",
                help="Tutte le altre variabili disponibili nell'export ForceDecks per questo tipo di test. "
                     "Non avendo un dato di popolazione, sono mostrate senza T-score.",
            )
            extra_metric_defs = [
                dict(key=f"extra::{jump_type}::{name}", label=extra_metric_label(name), category=cat,
                     jump_type=jump_type, raw_var=name, unit="", pop_key=None,
                     lower_is_better=False, kind="info")
                for name in extra_selected
            ]

            st.markdown("**Ripetizioni incluse nel calcolo**")
            st.caption("Deseleziona un salto per escluderlo da media, dev.std e T-score di questa categoria.")
            chk_cols = st.columns(n_reps)
            incl_mask = []
            for i in range(n_reps):
                key = f"incl_{jump_type}_{i}"
                checked = chk_cols[i].checkbox(f"Prova {i + 1}", value=st.session_state.get(key, True), key=key)
                incl_mask.append(checked)

            extra_results = []
            for m in extra_metric_defs:
                stats = pooled_stats(compute_metric_values(parsed_files, m))
                extra_results.append(dict(
                    key=m["key"], label=m["label"], category=cat, unit=m["unit"], kind="info",
                    n=stats["n"], mean=stats["mean"], sd=stats["sd"], z=None, t=None, banda=None,
                    colore=None, pop_mean=None, pop_sd=None,
                ))
            results_lookup = results + extra_results
            cat_metrics = cat_metrics + extra_metric_defs

            rows, best_per_row, worst_per_row = [], [], []
            cv_warnings = []
            jump_cols = [f"Prova {i + 1}" for i in range(n_reps)]
            for m in cat_metrics:
                r = next((x for x in results_lookup if x["key"] == m["key"]), None)
                values = per_rep_metric_values(parsed_files, m)
                best_i, worst_i = best_worst_indices(values, incl_mask, m["lower_is_better"])
                best_per_row.append(best_i)
                worst_per_row.append(worst_i)
                row = {"Metrica": m["label"], "Unità": m["unit"]}
                for i, v in enumerate(values):
                    row[jump_cols[i]] = round(v, 3) if isinstance(v, (int, float)) else None
                row["Media"] = round(r["mean"], 3) if r and r["mean"] is not None else None
                row["Media Pop."] = round(r["pop_mean"], 3) if r and r["pop_mean"] is not None else "—"
                row["Dev.Std"] = round(r["sd"], 3) if r and r["sd"] else None
                # CV% = coefficiente di variazione (dev.std / media * 100): misura
                # la variabilità relativa tra le ripetizioni incluse.
                cv = None
                if r and r["mean"] not in (None, 0) and r["sd"] is not None:
                    cv = abs(r["sd"] / r["mean"]) * 100
                row["CV%"] = round(cv, 1) if cv is not None else None
                if cv is not None and cv > 10:
                    cv_warnings.append(m["label"])
                row["T-score"] = round(r["t"], 1) if r and r["t"] is not None else "—"
                row["Valutazione"] = (r["banda"] if r and r["banda"] else "—")
                rows.append(row)

            for row in rows:
                export_rows.append({"Nome": nome, "Sesso": sesso, "Data test": periodo, "Categoria": cat, **row})

            df = pd.DataFrame(rows).reset_index(drop=True)

            def _highlight(row):
                styles = [""] * len(row)
                ridx = row.name
                best_i, worst_i = best_per_row[ridx], worst_per_row[ridx]
                for i, col in enumerate(jump_cols):
                    pos = df.columns.get_loc(col)
                    if not incl_mask[i]:
                        styles[pos] = "color:#adb5bd; text-decoration: line-through;"
                    elif i == best_i:
                        styles[pos] = "background-color:#3b8f3d; font-weight:600;"
                    elif i == worst_i:
                        styles[pos] = "background-color:#d13242; font-weight:600;"
                cv_val = row.get("CV%")
                if isinstance(cv_val, (int, float)) and cv_val > 10:
                    styles[df.columns.get_loc("CV%")] = "background-color:#E4572E; color:white; font-weight:600;"
                return styles

            styled = df.style.apply(_highlight, axis=1)
            st.dataframe(styled, use_container_width=True, hide_index=True)
            if cv_warnings:
                st.warning(
                    "⚠️ Coefficiente di variazione (CV%) superiore al 10% per: " + ", ".join(cv_warnings) +
                    ". Indica un'elevata variabilità tra le ripetizioni incluse: valutarne l'affidabilità."
                )

            # --- Grafico RSQ (Reactive Strength Quadrant) per le categorie
            # che includono un mRSI: mette in relazione la media di altezza
            # del salto (asse Y) e la media del tempo di contatto/contrazione
            # (asse X), con il crosshair centrato sulla media di popolazione.
            if cat in RSQ_METRIC_KEYS:
                rsq_fig = build_rsq_chart(cat, results)
                h_key, t_key = RSQ_METRIC_KEYS[cat]
                r_t = next(r for r in results if r["key"] == t_key)
                if rsq_fig:
                    st.markdown("**RSQ — Reactive Strength Quadrant**")
                    caption = (
                        "Il grafico dà contesto all'mRSI mostrando non solo quanto l'atleta è reattivo, "
                        "ma come esprime quella reattività: le linee tratteggiate sono centrate sulla media "
                        "di popolazione, il punto è la media del test con barre d'errore (± dev.std)."
                    )
                    if r_t["pop_mean"] is None:
                        caption += (
                            " Nota: per il tempo di contatto/contrazione non è disponibile un dato di "
                            "popolazione, quindi la linea verticale usa la media del test."
                        )
                    st.caption(caption)
                    st.plotly_chart(rsq_fig, use_container_width=True)

            # I controlli a soglia riguardano esclusivamente il CMJ Rebound:
            # vengono mostrati qui, subito dopo la tabella della COUNTERMOVEMENT JUMP REBOUND TEST.
            if cat == "COUNTERMOVEMENT JUMP REBOUND TEST":
                st.markdown("#### ✅ Controlli Tecnici (CMJ Rebound)")
                st.caption("Controlli a soglia sul CMJ Rebound, per individuare asimmetrie o esecuzioni tecnicamente scorrette.")
                for c in checks:
                    ccol1, ccol2 = st.columns([3, 1])
                    with ccol1:
                        st.markdown(f"**{c['label']}**")
                        st.caption(c["desc"])
                    with ccol2:
                        if c["value"] is None:
                            st.markdown("—")
                        else:
                            icon = "✅" if c["passed"] else "⚠️"
                            val_disp = c["value"] * c["scale"]
                            th_disp = c["threshold"] * c["scale"]
                            st.markdown(f"##### {icon} {val_disp:.{c['decimals']}f}{c['suffix']}")
                            soglia_lbl = "min" if c["direction"] == "min" else "max"
                            st.caption(f"soglia {soglia_lbl} {th_disp:.{c['decimals']}f}{c['suffix']}")

            st.markdown("---")

        # --- Esportazione CSV di tutti i dati mostrati nelle tabelle di
        # dettaglio (per ripetizione, media, dev.std, CV%, T-score),
        # pensata per raccogliere i dati grezzi da più atleti/test e
        # ricavarne in seguito medie e deviazioni standard di popolazione.
        if export_rows:
            st.markdown("### 📤 Esporta dati del test")
            st.caption(
                "Esporta in CSV tutti i dati delle tabelle sopra (valori per ripetizione, media, dev.std, "
                "CV% e T-score), utile per la successiva determinazione dei dati di popolazione."
            )
            df_export = pd.DataFrame(export_rows)
            csv_bytes = df_export.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "⬇️ Scarica dati test (.csv)", data=csv_bytes,
                file_name=f"forceplate_test_{nome.replace(' ', '_')}_{dt.date.today().isoformat()}.csv",
                mime="text/csv",
            )

with tab_profilo:
    if not parsed_files:
        st.info("Carica dei file dalla barra laterale per vedere il profilo di forza.")
    else:
        cats_valide = [c for c in CATEGORIES if profilo.get(c) is not None]
        if not cats_valide:
            st.warning("Nessuna metrica con confronto di popolazione disponibile: carica almeno un test tra IMTP, SJ, CMJ o CMJ RE.")
        else:
            radar_fig = build_radar_chart(cats_valide, profilo, nome)
            st.plotly_chart(radar_fig, use_container_width=True)

            cols = st.columns(len(cats_valide))
            for col, cat in zip(cols, cats_valide):
                t = profilo[cat]
                banda, colore = banda_da_tscore(t)
                col.markdown(f"**{cat}**")
                col.markdown(f"<span style='font-size:28px;color:{colore}'>{t:.0f}</span>", unsafe_allow_html=True)
                col.caption(banda)

        # --- DSI ed EUR come grafici a quadranti, invece che come solo
        # T-score: mostrano non solo il rapporto (DSI/EUR), ma anche il
        # livello assoluto di prestazione nei due test che lo compongono. Il
        # valore numerico dell'indice è mostrato accanto al proprio grafico
        # (non più in una riga "Indici" separata).
        r_dsi = next((r for r in results if r["key"] == "dsi" and r["mean"] is not None), None)
        dsi_fig = build_dsi_chart(results)
        if dsi_fig:
            st.markdown("#### DSI (Dynamic Strength Index)")
            if r_dsi:
                help_parts = []
                if r_dsi["t"] is not None:
                    help_parts.append(f"T-score: {r_dsi['t']:.0f} ({r_dsi['banda']})")
                if r_dsi["pop_mean"] is not None:
                    help_parts.append(f"Media popolazione: {r_dsi['pop_mean']:.3f}")
                st.metric("DSI", f"{r_dsi['mean']:.3f}", help=" — ".join(help_parts) or None)
            st.caption(
                "CMJ Peak Force vs IMTP Peak Force: le linee tratteggiate sono centrate sulla media di "
                "popolazione (dove disponibile), il punto è la media del test con barre d'errore (± dev.std)."
            )
            st.plotly_chart(dsi_fig, use_container_width=True)

        r_eur = next((r for r in results if r["key"] == "eur" and r["mean"] is not None), None)
        eur_fig = build_eur_chart(results)
        if eur_fig:
            st.markdown("#### EUR (Eccentric Utilisation Ratio)")
            if r_eur:
                help_parts = []
                if r_eur["t"] is not None:
                    help_parts.append(f"T-score: {r_eur['t']:.0f} ({r_eur['banda']})")
                if r_eur["pop_mean"] is not None:
                    help_parts.append(f"Media popolazione: {r_eur['pop_mean']:.3f}")
                st.metric("EUR", f"{r_eur['mean']:.3f}", help=" — ".join(help_parts) or None)
            st.caption(
                "Altezza SJ vs altezza CMJ: le linee tratteggiate sono centrate sulla media di popolazione, "
                "il punto è la media del test con barre d'errore (± dev.std). La linea diagonale rappresenta "
                "il valore medio di popolazione dell'EUR."
            )
            st.plotly_chart(eur_fig, use_container_width=True)
# PARTE 6 — REPORT SCARICABILE (HTML interattivo)
# ============================================================================
# Il report è un unico file HTML autosufficiente: i grafici Plotly restano
# interattivi (hover, zoom) invece di essere "appiattiti" in immagini
# statiche come richiederebbe un PDF. plotly.js viene caricato una sola
# volta via CDN nell'<head>; ogni grafico è poi un semplice <div> generato
# dalle stesse funzioni build_*_chart usate nella UI live (nessuna
# duplicazione della logica di analisi/plotting). Chi riceve il file può
# aprirlo in qualunque browser e, se serve una copia statica, stampare /
# "Salva come PDF" direttamente da lì.

def _fig_div(fig, div_id):
    """Converte una figura Plotly in un <div> HTML incorporabile nel report,
    senza reincludere plotly.js (già caricato una volta nell'head)."""
    if fig is None:
        return ""
    return fig.to_html(full_html=False, include_plotlyjs=False, div_id=div_id,
                        config={"displaylogo": False, "responsive": True})


def _banda_badge(banda, colore):
    if not banda:
        return "—"
    return f'<span class="badge" style="background:{colore}">{banda}</span>'


def _index_value_html(r):
    """Valore numerico di un indice (DSI/EUR), mostrato accanto al proprio
    grafico invece che in una tabella riepilogativa separata."""
    if r is None:
        return ""
    extra = []
    if r["pop_mean"] is not None:
        extra.append(f"Media pop. {r['pop_mean']:.3f}")
    if r["t"] is not None:
        extra.append(f"T-score {r['t']:.0f} ({r['banda']})")
    extra_html = f' <span class="muted">— {" · ".join(extra)}</span>' if extra else ""
    return f'<p class="index-value"><b>{r["label"]}: {r["mean"]:.3f}</b>{extra_html}</p>'


def _metric_table_html(cat_results):
    rows_html = []
    for r in cat_results:
        media = f"{r['mean']:.3f}" if r["mean"] is not None else "N/D"
        media_pop = f"{r['pop_mean']:.3f}" if r["pop_mean"] is not None else "—"
        tscore = f"{r['t']:.0f}" if r["t"] is not None else "—"
        rows_html.append(f"""<tr>
            <td>{r['label']}</td><td>{r['unit'] or ''}</td><td>{media}</td><td>{media_pop}</td>
            <td>{r['n']}</td><td>{tscore}</td><td>{_banda_badge(r['banda'], r['colore'])}</td>
        </tr>""")
    return f"""<table class="report-table">
        <thead><tr><th>Metrica</th><th>Unità</th><th>Media</th><th>Media Pop.</th><th>N</th><th>T-score</th><th>Valutazione</th></tr></thead>
        <tbody>{''.join(rows_html)}</tbody>
    </table>"""


def genera_report_html(nome, sesso, periodo, results, profilo, commento):
    div_counter = {"n": 0}

    def next_id(prefix):
        div_counter["n"] += 1
        return f"{prefix}_{div_counter['n']}"

    sections = []

    # --- Profilo di Forza ---
    cats_valide = [c for c in CATEGORIES if profilo.get(c) is not None]
    radar_fig = build_radar_chart(cats_valide, profilo, nome)
    profilo_cards = "".join(
        f"""<div class="profile-card">
                <div class="profile-card-cat">{cat}</div>
                <div class="profile-card-t" style="color:{banda_da_tscore(profilo[cat])[1]}">{profilo[cat]:.0f}</div>
                <div class="profile-card-banda">{banda_da_tscore(profilo[cat])[0]}</div>
            </div>"""
        for cat in cats_valide
    )
    sections.append(f"""<section>
        <h2>Profilo di Forza</h2>
        {_fig_div(radar_fig, next_id('radar')) if radar_fig else '<p class="muted">Nessuna metrica con confronto di popolazione disponibile.</p>'}
        <div class="profile-cards">{profilo_cards}</div>
    </section>""")

    # --- Categorie di test ---
    # I controlli tecnici a soglia (CMJ Rebound) restano solo nella scheda
    # live "Dettaglio Test": non vengono replicati nel report scaricabile.
    for cat in CATEGORIES:
        cat_results = [r for r in results if r["category"] == cat and r["mean"] is not None]
        if not cat_results:
            continue
        tscore_fig = build_tscore_bar_chart(cat_results)
        rsq_fig = build_rsq_chart(cat, results)

        sections.append(f"""<section>
            <h2>{cat}</h2>
            {_fig_div(tscore_fig, next_id('tscore'))}
            {_metric_table_html(cat_results)}
            {f'<h3>RSQ — Reactive Strength Quadrant</h3>{_fig_div(rsq_fig, next_id("rsq"))}' if rsq_fig else ''}
        </section>""")

    # --- Indici (DSI, EUR): il valore numerico è mostrato accanto al
    # proprio grafico, non più in una tabella riepilogativa separata.
    r_dsi = next((r for r in results if r["key"] == "dsi" and r["mean"] is not None), None)
    r_eur = next((r for r in results if r["key"] == "eur" and r["mean"] is not None), None)
    dsi_fig = build_dsi_chart(results)
    eur_fig = build_eur_chart(results)
    if r_dsi or r_eur or dsi_fig or eur_fig:
        indici_html = []
        if r_dsi or dsi_fig:
            indici_html.append(
                f"<h3>DSI (Dynamic Strength Index)</h3>{_index_value_html(r_dsi)}{_fig_div(dsi_fig, next_id('dsi'))}"
            )
        if r_eur or eur_fig:
            indici_html.append(
                f"<h3>EUR (Eccentric Utilisation Ratio)</h3>{_index_value_html(r_eur)}{_fig_div(eur_fig, next_id('eur'))}"
            )
        sections.append(f"""<section>
            <h2>Indici</h2>
            {''.join(indici_html)}
        </section>""")

    # --- Analisi: testo scritto dal preparatore nella scheda Report
    # dell'app, incluso qui come testo statico (non modificabile nel report).
    commento_html = _html.escape(commento).replace("\n", "<br>") if commento and commento.strip() else (
        "Nessuna analisi inserita."
    )
    sections.append(f"""<section>
        <h2>Analisi</h2>
        <div class="analysis-box">{commento_html}</div>
    </section>""")

    html = f"""<!DOCTYPE html>
<html lang="it">
<head>
<meta charset="utf-8">
<title>Force Plate Test Report — {nome}</title>
<script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
<style>
    :root {{ --primary: {PRIMARY}; --accent: {ACCENT}; --text: {TEXT_COLOR}; --bg: {BG_COLOR}; }}
    * {{ box-sizing: border-box; }}
    body {{
        font-family: "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
        background: var(--bg); color: var(--text); margin: 0; padding: 0 0 60px 0;
    }}
    .report {{ max-width: 960px; margin: 0 auto; padding: 0 24px; }}
    header.report-header {{ background: var(--text); color: #fff; padding: 32px 24px; margin-bottom: 24px; }}
    header.report-header h1 {{ margin: 0 0 12px 0; font-size: 26px; letter-spacing: 0.5px; }}
    .meta-row span {{ margin-right: 24px; font-size: 15px; }}
    .meta-row b {{ color: var(--primary); }}
    section {{ margin-bottom: 36px; padding-bottom: 24px; border-bottom: 1px solid #e0e0e0; }}
    h2 {{ color: var(--text); border-left: 5px solid var(--accent); padding-left: 10px; font-size: 20px; }}
    h3 {{ color: var(--text); font-size: 16px; margin-top: 24px; }}
    .intro-text {{ line-height: 1.5; }}
    table.report-table {{ width: 100%; border-collapse: collapse; margin-top: 12px; font-size: 14px; }}
    table.report-table th {{ background: var(--text); color: #fff; text-align: left; padding: 8px 10px; font-weight: 600; }}
    table.report-table td {{ padding: 8px 10px; border-bottom: 1px solid #eee; }}
    table.report-table tr:nth-child(even) td {{ background: #f7fbfd; }}
    .badge {{ color: #fff; padding: 3px 10px; border-radius: 12px; font-size: 12px; font-weight: 600; display: inline-block; }}
    .profile-cards {{ display: flex; flex-wrap: wrap; gap: 16px; margin-top: 16px; }}
    .profile-card {{ flex: 1 1 200px; background: #f7fbfd; border-radius: 8px; padding: 14px 16px; border: 1px solid #e0e0e0; }}
    .profile-card-cat {{ font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px; color: var(--text); opacity: 0.7; }}
    .profile-card-t {{ font-size: 30px; font-weight: 700; }}
    .profile-card-banda {{ font-size: 13px; }}
    .muted {{ color: #667; font-size: 13px; }}
    .index-value {{ font-size: 16px; margin: 6px 0 14px 0; }}
    .analysis-box {{ min-height: 60px; border: 1px solid #e0e0e0; border-left: 4px solid var(--primary); border-radius: 6px; padding: 14px; line-height: 1.6; background: #f7fbfd; }}
    @media print {{ section {{ break-inside: avoid; }} }}
</style>
</head>
<body>
    <div class="report">
        <header class="report-header">
            <h1>🏋️ FORCE PLATE TEST REPORT</h1>
            <div class="meta-row">
                <span><b>Atleta:</b> {nome}</span>
                <span><b>Sesso:</b> {sesso}</span>
                <span><b>Data test:</b> {periodo}</span>
            </div>
        </header>
        <p class="intro-text">
            Per valutare i risultati del test è stato utilizzato il T-Score, un indice standardizzato
            che confronta la prestazione dell'atleta rispetto a un gruppo di riferimento, esprimendo la
            distanza dalla media in deviazioni standard. Punteggi tra 0 e 50 indicano valori inferiori
            alla media, mentre punteggi tra 50 e 100 indicano valori superiori alla media.
        </p>
        {''.join(sections)}
    </div>
</body>
</html>"""
    return html.encode("utf-8")


with tab_report:
    if not parsed_files:
        st.info("Carica dei file dalla barra laterale per generare il report.")
    else:
        st.markdown(
            "Genera un report HTML autosufficiente (grafici interattivi inclusi) con profilo di forza "
            "e tabelle riassuntive. Si apre in qualunque browser; se serve una copia statica, si può "
            "stampare/salvare come PDF direttamente da lì."
        )
        st.markdown("**Analisi del preparatore**")
        st.caption(
            "Scrivi qui il commento tecnico da includere nel report: punti di forza, aree di "
            "miglioramento e indicazioni di lavoro. Nel report scaricato il testo sarà statico "
            "(non modificabile da chi lo riceve)."
        )
        commento = st.text_area(
            "Analisi del preparatore", key="coach_comment", height=160, label_visibility="collapsed",
            placeholder="Es. Buoni valori di forza isometrica, mRSI sopra media. Da lavorare sulla "
                        "reattività nel CMJ Rebound...",
        )
        if st.button("📄 Genera report HTML", type="primary"):
            html_bytes = genera_report_html(nome, sesso, periodo, results, profilo, commento)
            st.download_button(
                "⬇️ Scarica report (.html)", data=html_bytes,
                file_name=f"Report_{nome.replace(' ', '_')}_{dt.date.today().isoformat()}.html",
                mime="text/html",
            )