"""Comparazione tra sessioni (Hopkins TE/SWC): logica, grafici e scheda 🔀 Comparazione."""

from __future__ import annotations

import datetime as dt
import io
import math
import re

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from core import (
    ACCENT,
    BG_COLOR,
    PRIMARY,
    TEXT_COLOR,
    _wrap_label,
    collect_reps,
    nel_profilo,
    pooled_stats,
    z_t_score,
)
from data import (
    CATEGORIES,
    CATEGORY_QUALITY,
    INDEX_ZONE_COLORS,
    INDEX_ZONE_LABELS,
    JUMP_TYPE_LABELS,
    METRICS,
    banda_da_tscore,
    extra_metric_label,
    zona_da_indice,
)

# ============================================================================
# PARTE 5bis — COMPARAZIONE TRA SESSIONI
# ============================================================================
# Confronta il test attualmente caricato con lo storico dell'atleta, caricato
# come CSV prodotti da "Dettaglio Test → Esporta TUTTE le metriche grezze".
#
# Due confronti con statuto diverso:
#
#  - CONTRO LA MEDIA delle sedute precedenti (esclusa l'attuale): è il
#    confronto statistico. La media di k sedute è più precisa della singola,
#    quindi l'incertezza si stringe: SE = TE * sqrt(1 + 1/k) invece di
#    TE * sqrt(2) del testa a testa.
#
#  - CONTRO IL MIGLIOR VALORE storico: puramente DESCRITTIVO, senza giudizio
#    di significatività. Il massimo di k misure rumorose è per costruzione
#    una sovrastima (il record è quasi sempre il giorno in cui l'atleta era
#    in forma E il rumore ha aiutato), e la distorsione cresce con k:
#    applicarci un test darebbe "peggiorato" in modo sistematico.
#
# L'app non dice MAI se un cambiamento è un bene: mostra direzione (freccia)
# e ampiezza (scala SWC di Hopkins). Il giudizio di merito resta al
# preparatore, che è l'unico a conoscere obiettivo e contesto.
#
#   TE  = errore tipico between-day (rumore della misura), in CV%
#   SE  = TE * sqrt(1 + 1/k)           incertezza della differenza dalla media
#   IC  = delta ± 1.645 * SE           intervallo di confidenza al 90%
#   SWC = 0.2 * SD                     soglia di rilevanza pratica

Z90 = 1.645
SWC_FACTOR = 0.2

# CV% between-day di letteratura, usati finché non ci sono abbastanza
# sessioni per stimare il TE sull'atleta stesso. Il primo pattern che
# matcha vince: ordine dal più specifico al più generico.
DEFAULT_TE_CV = [
    (("asimmetria", "asymmetry", "sym. index", "symmetry index"), 20.0),
    (("rfd",), 15.0),
    (("dsi", "eur"), 10.0),
    (("eccentric", "eccentrica", "braking", "frenata", "unweight"), 9.0),
    (("rsi",), 7.0),
    (("contraction time", "contact time", "time to", "duration", "durata", "time"), 7.0),
    (("power", "potenza"), 5.0),
    (("impulse", "impulso", "work"), 5.0),
    (("velocity", "velocità"), 4.0),
    (("height", "altezza"), 4.0),
    (("force", "forza", "torque"), 4.0),
    (("mass", "weight", "peso"), 1.5),
]
FALLBACK_TE_CV = 8.0

# Indici di simmetria: il segno indica il lato dominante, non la qualità,
# quindi vanno confrontati in valore assoluto (da -2% a +6% è un aumento
# dello sbilanciamento, non una diminuzione).
COMP_ABS_KEYWORDS = ("sym. index", "symmetry index", "asimmetria", "asymmetry")
COMP_ABS_SUFFIXES = (" si",)

# Scala di ampiezza di Hopkins, in multipli di SWC (= 0,2 SD, quindi
# 1x = 0.2 SD, 3x = 0.6, 6x = 1.2, 10x = 2.0 deviazioni standard).
# Il livello "Trascurabile" sotto 1x non è nella scala originale ma serve:
# senza, i cambiamenti minuscoli non avrebbero etichetta.
SWC_SCALE = [
    (1.0, "Trascurabile", "#e4e4dd"),
    (3.0, "Piccolo", "#cfe3f7"),
    (6.0, "Moderato", "#9ecbf0"),
    (10.0, "Grande", "#4FC3F7"),
    (float("inf"), "Molto grande", "#0bb6ff"),
]

# Semaforo sull'attendibilita' del cambiamento, non sulla sua desiderabilita':
# verde = c'e' stato un cambiamento reale; giallo = non si puo' dire; rosso =
# quasi certamente nessun cambiamento rilevante.
ESITO_REALE = "\U0001F7E2 Alta"
ESITO_INCERTO = "\U0001F7E1 Incerta"
ESITO_STABILE = "\U0001F534 Bassa"
ESITO_ND = "\u26A0\uFE0F Non valutabile"
ESITO_COLORI = {
    ESITO_REALE: "#2E7D32",
    ESITO_INCERTO: "#B8860B",
    ESITO_STABILE: "#B00020",
    ESITO_ND: "#8d8d8d",
}

# Colonne di intestazione dell'export completo: tutto il resto è variabile.
COMP_META_COLS = {"Nome", "Sesso", "Data test", "File", "Jump Type", "Indice Ripetizione"}

# Bridge tra variabili grezze e catalogo curato: dove esiste una metrica
# curata si usano la sua etichetta e la sua norma di popolazione.
COMP_CURATED_BY_RAW = {
    (m["jump_type"], m["raw_var"]): m
    for m in METRICS if m.get("jump_type") and m.get("raw_var")
}
COMP_DERIVED = [m for m in METRICS if m.get("jump_type") and m.get("derive")]

COMP_DEFAULT_PRIMARY_KEYS = (
    "imtp_peak_force", "sj_height", "cmj_height", "mrsi_cmj", "cmj_re_rebound_height",
)


def _metric_id(jump_type, var_name):
    """Identità di una metrica tra sessioni: tipo di test + nome variabile.
    Il tipo serve perché la stessa variabile grezza (es. 'peak force')
    esiste in test diversi con significati diversi."""
    return f"{jump_type}||{var_name}"


def comp_id_da_risultato(key):
    """metric_id della Comparazione a partire dalla chiave di un risultato
    di build_results() (curata, derivata o 'extra::tipo::variabile')."""
    if key.startswith("extra::"):
        _, jump_type, var = key.split("::", 2)
        return _metric_id(jump_type, var)
    m = next((x for x in METRICS if x["key"] == key), None)
    if m and m.get("jump_type"):
        return _metric_id(m["jump_type"], m.get("raw_var") or m["key"])
    return None


def _metric_label(jump_type, var_name):
    m = COMP_CURATED_BY_RAW.get((jump_type, var_name))
    if m:
        return m["label"]
    m = next((x for x in COMP_DERIVED
              if x["key"] == var_name and x["jump_type"] == jump_type), None)
    if m:
        return m["label"]
    return extra_metric_label(var_name)


def _metric_unit(jump_type, var_name):
    m = COMP_CURATED_BY_RAW.get((jump_type, var_name))
    if m:
        return m.get("unit") or ""
    m = next((x for x in COMP_DERIVED
              if x["key"] == var_name and x["jump_type"] == jump_type), None)
    return (m.get("unit") or "") if m else ""


def is_symmetry_metric(label):
    lab = str(label).strip().lower()
    return any(k in lab for k in COMP_ABS_KEYWORDS) or lab.endswith(COMP_ABS_SUFFIXES)


def comp_lower_is_better(jump_type, var_name):
    """Solo per scegliere QUALE valore storico è il migliore. Non viene
    usato per giudicare il cambiamento: quello lo fa il preparatore."""
    m = COMP_CURATED_BY_RAW.get((jump_type, var_name))
    if m is None:
        m = next((x for x in COMP_DERIVED
                  if x["key"] == var_name and x["jump_type"] == jump_type), None)
    if m is not None:
        return bool(m["lower_is_better"]), True
    if is_symmetry_metric(var_name):
        return True, True
    return False, False


def te_cv_default(label):
    """CV% between-day di letteratura per una metrica, dal suo nome."""
    lab = str(label).lower()
    for keys, cv in DEFAULT_TE_CV:
        if any(k in lab for k in keys):
            return cv
    return FALLBACK_TE_CV


# ---------------------------------------------------------------------------
# L'incertezza: errore entro seduta
# ---------------------------------------------------------------------------
# L'incertezza di una media di seduta è la SD delle ripetizioni diviso radice
# di n: misura SOLO l'errore strumentale e di esecuzione DENTRO la giornata.
#
# La scelta è deliberata. Fra un giorno e l'altro NON si assume che l'atleta
# sia rimasto uguale: fatica e adattamento sono cambiamenti reali, non rumore.
# Usare la variabilità giorno-per-giorno come soglia renderebbe invisibile
# proprio ciò che il monitoraggio deve vedere — e penalizzerebbe l'atleta che
# risponde di più al carico, perché produrrebbe soglie più larghe.
#
# La soglia di rilevanza resta 0,2 x SD di POPOLAZIONE: una quantità esterna
# all'atleta, che non dipende da come è andata la sua ultima settimana.
# Servono almeno 3 ripetizioni: sotto, la SD non è stimabile e si ricade sul
# CV di letteratura.
MIN_REP_MISURA = 3


def se_entro_seduta(sd, n_rep, media):
    """Errore standard della media di una seduta: SD delle ripetizioni diviso
    radice di n. Nessun pavimento: si usa ciò che si misura.

    Il limite noto è che questa stima NON vede la componente fra giorni
    (riposizionamento sulla pedana, riscaldamento, ora del test), quindi
    sottostima l'errore reale. La correzione arriverà dal CV entro seduta
    calcolato sull'archivio di laboratorio, che sostituirà la SD della singola
    seduta con una costante per metrica, molto più stabile di una stima
    ricavata da 3 salti."""
    try:
        n_rep = int(n_rep)
    except (TypeError, ValueError):
        return None
    if n_rep < MIN_REP_MISURA or media in (None, 0):
        return None
    try:
        if sd is None or math.isnan(float(sd)):
            return None
    except (TypeError, ValueError):
        return None
    se = float(sd) / math.sqrt(n_rep)
    return se if se > 0 else None


def fmt_valore(v, segno=False):
    """Numero leggibile senza notazione scientifica: i decimali si adattano
    alla grandezza. "{:.3g}" trasformerebbe 1815 in "1.82e+03", illeggibile
    per un allenatore che si aspetta dei newton."""
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return "—"
    if not isinstance(v, (int, float)):
        return str(v)
    a = abs(v)
    dec = 0 if a >= 100 else (1 if a >= 10 else (2 if a >= 1 else 3))
    return f"{v:+.{dec}f}" if segno else f"{v:.{dec}f}"


def scala_ampiezza(rapporto):
    """(etichetta, colore) dal rapporto |delta| / SWC."""
    for soglia, etichetta, colore in SWC_SCALE:
        if rapporto < soglia:
            return etichetta, colore
    return SWC_SCALE[-1][1], SWC_SCALE[-1][2]


def comp_pop_sd_map(pop_dict, sesso_atleta):
    """{metric_id: SD di popolazione} dal set di norme attivo (Letteratura o
    Team). Prima le chiavi 'tipo::variabile' (presenti solo nel set Team),
    poi le costanti curate, che hanno la precedenza. Senza SD per quel sesso
    la metrica resta senza soglia (Non valutabile)."""
    is_female = bool(sesso_atleta) and str(sesso_atleta).upper().startswith("F")
    sd_key = "sd_f" if is_female else "sd_m"
    out = {}

    for key, c in pop_dict.items():
        if "::" not in key:
            continue
        jump_type, var = key.split("::", 1)
        sd = c.get(sd_key)
        if sd:
            out[_metric_id(jump_type, var)] = float(sd)

    for m in METRICS:
        if not m.get("pop_key") or m["pop_key"] not in pop_dict:
            continue
        sd = pop_dict[m["pop_key"]][sd_key]
        if not sd:
            continue
        if m.get("jump_type") and m.get("raw_var"):
            out[_metric_id(m["jump_type"], m["raw_var"])] = float(sd)
        elif m.get("jump_type") and m.get("derive"):
            out[_metric_id(m["jump_type"], m["key"])] = float(sd)
        elif m["key"] in ("dsi", "eur"):
            out[_metric_id("indici", m["key"])] = float(sd)
    return out


# ----------------------------------------------------------------------------
# Da ripetizioni a formato lungo
# ----------------------------------------------------------------------------
COMP_RATIO_IDS = {_metric_id("indici", "dsi"), _metric_id("indici", "eur")}


def reps_to_long(reps_by_type, session_label, data_test=None, ordine=0):
    """reps_by_type: {jump_type: [dict_variabili_per_ripetizione, ...]}.

    Aggrega media/dev.std/n per ogni variabile presente, aggiunge le metriche
    derivate del catalogo (rapporti per kg) e infine gli indici cross-test
    DSI ed EUR calcolati dalle medie della sessione."""
    rows = []
    medie = {}

    def _add(jump_type, var, etichetta, stats):
        medie[(jump_type, var)] = stats["mean"]
        rows.append(dict(
            session=session_label, data=data_test, ordine=ordine, jump_type=jump_type,
            test=JUMP_TYPE_LABELS.get(jump_type, jump_type), var=var,
            metric_id=_metric_id(jump_type, var), metrica=etichetta,
            unit=_metric_unit(jump_type, var),
            mean=stats["mean"], sd=stats["sd"], n=stats["n"],
        ))

    for jump_type, reps in reps_by_type.items():
        if not jump_type or not reps:
            continue
        for var in sorted({k for rep in reps for k in rep.keys()}):
            stats = pooled_stats([rep.get(var) for rep in reps
                                  if isinstance(rep.get(var), (int, float))])
            if stats["n"]:
                _add(jump_type, var, _metric_label(jump_type, var), stats)
        for m in COMP_DERIVED:
            if m["jump_type"] != jump_type:
                continue
            stats = pooled_stats([v for v in (m["derive"](rep) for rep in reps)
                                  if isinstance(v, (int, float))])
            if stats["n"]:
                _add(jump_type, m["key"], m["label"], stats)

    # Indici cross-test: stesse formule della PARTE 4, calcolate qui sulle
    # medie di questa sessione. Essendo rapporti fra test diversi non
    # esistono "per ripetizione" e non hanno dev.std.
    cmj_peak = medie.get(("cmj", "peak propulsive force"))
    imtp_peak = medie.get(("imtp", "peak force"))
    cmj_h = medie.get(("cmj", "jump height ft"))
    sj_h = medie.get(("sj", "jump height ft"))
    for key, val in (("dsi", (cmj_peak / imtp_peak) if (cmj_peak and imtp_peak) else None),
                     ("eur", (cmj_h / sj_h) if (cmj_h and sj_h) else None)):
        if val is None:
            continue
        rows.append(dict(
            session=session_label, data=data_test, ordine=ordine, jump_type="indici",
            test="Indici", var=key, metric_id=_metric_id("indici", key),
            metrica=key.upper(), unit="", mean=val, sd=None, n=1,
        ))

    return pd.DataFrame(rows)


def current_reps_by_type(files):
    """Ripetizioni della sessione attuale, già filtrate dalle checkbox
    'Prova N' della scheda Dettaglio Test."""
    out = {}
    for jump_type in ("imtp", "sj", "cmj", "cmrj"):
        reps = collect_reps(files, jump_type)
        if reps:
            out[jump_type] = [rep["vars"] for rep in reps]
    return out


def _parse_data_test(valore):
    """La colonna 'Data test' contiene 'gg/mm/aaaa' oppure un intervallo
    'gg/mm/aaaa → gg/mm/aaaa': si prende la prima data utile."""
    if valore is None:
        return None
    testo = str(valore).strip()
    match = re.search(r"(\d{2})/(\d{2})/(\d{4})", testo)
    if not match:
        return None
    try:
        return dt.date(int(match.group(3)), int(match.group(2)), int(match.group(1)))
    except ValueError:
        return None


def parse_full_export_csv(file_like, fallback_label, ordine=0):
    """Legge un CSV prodotto da 'Scarica export completo (.csv)':
    una riga per ripetizione, una colonna per variabile grezza."""
    df = pd.read_csv(file_like, encoding="utf-8-sig")
    if "Jump Type" not in df.columns:
        raise ValueError(
            "manca la colonna 'Jump Type'. Serve il file prodotto da "
            "\"Scarica export completo (.csv)\", non l'altro export."
        )

    etichetta, data_test = fallback_label, None
    if "Data test" in df.columns and len(df):
        v = str(df["Data test"].iloc[0]).strip()
        if v and v.lower() not in ("-", "nan", "none"):
            etichetta = v
            data_test = _parse_data_test(v)

    var_cols = [c for c in df.columns if c not in COMP_META_COLS]
    if not var_cols:
        raise ValueError("nessuna colonna di variabili trovata.")

    reps_by_type = {}
    for _, row in df.iterrows():
        jump_type = str(row["Jump Type"]).strip().lower()
        if not jump_type or jump_type in ("—", "-", "nan", "none"):
            continue
        rep = {}
        for c in var_cols:
            v = pd.to_numeric(row[c], errors="coerce")
            if pd.notna(v):
                rep[c] = float(v)
        if rep:
            reps_by_type.setdefault(jump_type, []).append(rep)

    if not reps_by_type:
        raise ValueError("nessuna ripetizione con un tipo di test valido.")
    return reps_to_long(reps_by_type, etichetta, data_test, ordine)


def ordina_sessioni(long_frames):
    """Riordina le sessioni per 'Data test'. Serve al TE dell'atleta, che si
    calcola sulle differenze tra sedute CONSECUTIVE: con i file caricati in
    ordine sparso il TE risulterebbe sbagliato. Le sessioni senza data
    leggibile restano in coda, nell'ordine di caricamento."""
    con_data, senza_data = [], []
    for frame in long_frames:
        if frame.empty:
            continue
        data = frame["data"].iloc[0]
        (con_data if data is not None else senza_data).append((data, frame))
    con_data.sort(key=lambda t: t[0])
    ordinati = [f for _, f in con_data] + [f for f in senza_data]
    for i, frame in enumerate(ordinati):
        frame["ordine"] = i
    return ordinati


# ----------------------------------------------------------------------------
# Typical Error dall'atleta
# ----------------------------------------------------------------------------
# ----------------------------------------------------------------------------
# Motore di confronto
# ----------------------------------------------------------------------------
def compare_to_history(cur_long, hist_long, pop_sd_map=None):
    """Confronta la sessione attuale con la MEDIA delle sedute precedenti, e
    riporta accanto il miglior valore storico (descrittivo)."""
    pop_sd_map = pop_sd_map or {}
    if cur_long.empty or hist_long.empty:
        return pd.DataFrame()

    cur = cur_long.drop_duplicates("metric_id").set_index("metric_id")

    rows = []
    for mid, g in hist_long.groupby("metric_id"):
        if mid not in cur.index:
            continue
        storiche = [v for v in g["mean"].tolist()
                    if isinstance(v, (int, float)) and not math.isnan(v)]
        if not storiche:
            continue
        m_cur = float(cur.loc[mid, "mean"])
        if math.isnan(m_cur):
            continue

        etichetta = cur.loc[mid, "metrica"]
        jump_type, var = cur.loc[mid, "jump_type"], cur.loc[mid, "var"]
        if is_symmetry_metric(etichetta) or is_symmetry_metric(var):
            m_cur, storiche = abs(m_cur), [abs(v) for v in storiche]

        k = len(storiche)
        media_rif = sum(storiche) / k
        if media_rif == 0:
            continue

        # Miglior valore storico: direzione nota solo per il catalogo curato,
        # altrove si usa il massimo. È una colonna descrittiva, senza test.
        lower_better, _ = comp_lower_is_better(jump_type, var)
        migliore = min(storiche) if lower_better else max(storiche)
        riga_migliore = g[g["mean"].apply(
            lambda v: isinstance(v, (int, float)) and abs(abs(v) - migliore) < 1e-9
            if (is_symmetry_metric(etichetta) or is_symmetry_metric(var))
            else v == migliore)]
        sess_migliore = riga_migliore["session"].iloc[0] if len(riga_migliore) else "—"

        log_scale = mid in COMP_RATIO_IDS and m_cur > 0 and media_rif > 0
        if log_scale:
            delta = math.log(m_cur / media_rif)
            mostra = lambda x: (math.exp(x) - 1) * 100.0
        else:
            delta = m_cur - media_rif
            mostra = lambda x: x

        # Incertezza della differenza, in unità native. Nel modo "misura" ogni
        # seduta ha il proprio errore: quello della media di k sedute vale
        # sqrt(somma dei quadrati)/k, non se/sqrt(k), perché le sedute possono
        # avere numeri di ripetizioni diversi.
        se_cur = se_entro_seduta(cur.loc[mid, "sd"], cur.loc[mid, "n"], m_cur)
        se_storiche = [x for x in (se_entro_seduta(r["sd"], r["n"], r["mean"])
                                   for _, r in g.iterrows()) if x]
        if se_cur and se_storiche:
            se_rif = math.sqrt(sum(x * x for x in se_storiche)) / len(se_storiche)
            se_diff_abs = math.sqrt(se_cur ** 2 + se_rif ** 2)
            cv, fonte_te = (se_cur / abs(m_cur) * 100.0 if m_cur else None), "misura"
        else:
            # Meno di 3 ripetizioni, oppure indici come DSI/EUR che sono
            # rapporti fra test e non hanno ripetizioni proprie.
            cv, fonte_te = te_cv_default(etichetta), "letteratura"
            se_diff_abs = (cv / 100.0 * abs(media_rif)) * math.sqrt(1.0 + 1.0 / k)
        if log_scale:
            se_diff_abs = se_diff_abs / abs(media_rif)

        # SWC: unicamente 0,2 x SD di popolazione. Niente surrogati ricavati
        # dallo storico dell'atleta: senza norma la metrica resta non
        # valutabile, e lo dice apertamente invece di dare un giudizio
        # costruito su una soglia inventata.
        pop_sd = pop_sd_map.get(mid)
        if pop_sd:
            swc = SWC_FACTOR * (pop_sd / abs(media_rif) if log_scale else pop_sd)
            fonte_swc = "norma"
        else:
            swc, fonte_swc = None, "—"

        margine = Z90 * se_diff_abs
        lo, hi = delta - margine, delta + margine

        if swc is None:
            esito, ampiezza, colore_amp = ESITO_ND, "—", None
            rapporto = None
        else:
            rapporto = abs(delta) / swc if swc else 0.0
            etichetta_amp, colore_amp = scala_ampiezza(rapporto)
            freccia = "🔺" if delta > 0 else ("🔻" if delta < 0 else "▪")
            ampiezza = f"{freccia} {etichetta_amp}"
            if lo > swc or hi < -swc:
                esito = ESITO_REALE
            elif lo > -swc and hi < swc:
                esito = ESITO_STABILE
            else:
                esito = ESITO_INCERTO

        rows.append({
            "Test": cur.loc[mid, "test"],
            "Metrica": etichetta,
            "Attendibilità": esito,
            "Cambiamento": ampiezza,
            "Unità": cur.loc[mid, "unit"],
            "Media rif.": media_rif,
            "Attuale": m_cur,
            "Delta": mostra(delta),
            "Delta %": (m_cur / media_rif - 1) * 100.0,
            "Migliore": migliore,
            # Sempre orientata alla prestazione: 100% = pari al record,
            # oltre 100% = meglio del record. Per le metriche dove meno e'
            # meglio il rapporto va invertito, altrimenti un nuovo primato
            # (valore piu' basso) apparirebbe come un calo.
            "% del migliore": ((migliore / m_cur if lower_better else m_cur / migliore) * 100.0
                               if migliore and m_cur else None),
            "Seduta migliore": sess_migliore,
            "N sedute": k,
            "_mid": mid,
            "_display": f"{cur.loc[mid, 'test']} · {etichetta}",
            "_swc": swc,
            "_margine": margine,
            "_rapporto": rapporto,
            "_colore_amp": colore_amp,
            "_log": log_scale,
            "_fonte_te": fonte_te,
            "_fonte_swc": fonte_swc,
        })

    if not rows:
        return pd.DataFrame()
    return pd.DataFrame(rows).sort_values(
        "_rapporto", ascending=False, na_position="last")


def build_swc_strip(riga, height=134):
    """Striscia a bande per una metrica, nelle sue unità native. Stessa
    geometria del grafico usato per DSI/EUR: fascia sottile, marcatore del
    valore, etichette sopra. Le bande cadono a ±1/3/6/10 SWC attorno alla
    media di riferimento, così l'allenatore legge insieme il valore reale e
    quanto si è spostato."""
    swc, media_rif = riga["_swc"], riga["Media rif."]
    if not swc or media_rif is None:
        return None

    # In scala log il grafico resta in unità native: le soglie sono
    # moltiplicative, quindi si convertono in valori assoluti.
    def confine(mult):
        if riga["_log"]:
            return media_rif * math.exp(mult * swc)
        return media_rif + mult * swc

    valore = riga["Attuale"]
    margine_abs = (abs(confine(riga["_margine"] / swc) - media_rif)
                   if riga["_log"] else riga["_margine"])

    limiti = [1.0, 3.0, 6.0, 10.0]
    estensione = max(limiti[-1], (riga["_rapporto"] or 0) * 1.25 + 1.5)
    lo_ax, hi_ax = confine(-estensione), confine(estensione)
    if lo_ax > hi_ax:
        lo_ax, hi_ax = hi_ax, lo_ax

    fig = go.Figure()
    bordi = [-estensione] + [-m for m in reversed(limiti)] + limiti + [estensione]
    for a, b in zip(bordi, bordi[1:]):
        centro = (abs(a) + abs(b)) / 2
        _, colore = scala_ampiezza(centro if a * b >= 0 else 0.0)
        x0, x1 = confine(a), confine(b)
        fig.add_vrect(x0=min(x0, x1), x1=max(x0, x1), fillcolor=colore,
                      opacity=0.55, line_width=0)

    for mult, etichetta in ((-6.5, "Grande"), (-4.5, "Moderato"), (-2.0, "Piccolo"),
                            (0.0, "Trascurabile"), (2.0, "Piccolo"), (4.5, "Moderato"),
                            (6.5, "Grande")):
        if abs(mult) > estensione:
            continue
        fig.add_annotation(x=confine(mult), y=1.10, xref="x", yref="paper",
                           yanchor="bottom", text=etichetta, showarrow=False,
                           font=dict(size=9, color="#484343"))

    # La media di riferimento e' un rombo VUOTO, non una linea tratteggiata:
    # su una fascia alta ~20px una vline verticale e' praticamente invisibile.
    # Vuoto vs pieno distingue il riferimento dal dato del test a colpo d'occhio.
    fig.add_trace(go.Scatter(
        x=[media_rif], y=[0], mode="markers", name="Media rif.",
        marker=dict(size=15, color=ACCENT, symbol="diamond-open",
                    line=dict(width=3, color=ACCENT)),
        cliponaxis=False,
        hovertemplate=f"Media riferimento: {fmt_valore(media_rif)}<extra></extra>",
    ))
    fig.add_trace(go.Scatter(
        x=[valore], y=[0], mode="markers+text",
        marker=dict(size=13, color=PRIMARY, symbol="diamond",
                    line=dict(width=1.5, color="white")),
        error_x=dict(type="data", array=[margine_abs], color=PRIMARY,
                     thickness=1.5, width=5),
        text=[fmt_valore(valore)], textposition="bottom center",
        textfont=dict(color=TEXT_COLOR, size=13), cliponaxis=False,
        hovertemplate=(f"Attuale: %{{x:.4g}}<br>Media rif.: {fmt_valore(media_rif)}"
                       f"<extra></extra>"),
    ))

    unita = f" ({riga['Unità']})" if riga["Unità"] else ""
    # Il nome della metrica sta SOPRA il grafico come titolo, non sotto come
    # etichetta dell'asse: e' l'informazione che si cerca per prima quando si
    # scorre una sequenza di strisce. Il margine superiore va alzato insieme
    # all'altezza, altrimenti il titolo si sovrappone alle etichette delle
    # fasce (annotazioni in coordinate paper appena sopra l'area di plot).
    fig.update_layout(
        title=dict(text=f"<b>{riga['Metrica']}{unita}</b>", x=0.5, xanchor="center",
                   y=0.90, yanchor="top",
                   font=dict(size=15, color=TEXT_COLOR)),
        xaxis=dict(range=[lo_ax, hi_ax], showgrid=False, automargin=True,
                   tickvals=[media_rif], ticktext=[f"media {fmt_valore(media_rif)}"]),
        yaxis=dict(range=[-1, 1], showgrid=False, showticklabels=False, zeroline=False),
        height=height, margin=dict(t=64, b=50, l=20, r=20), showlegend=False,
        plot_bgcolor=BG_COLOR, paper_bgcolor=BG_COLOR, font=dict(color=TEXT_COLOR),
    )
    return fig


# ----------------------------------------------------------------------------
# Titoli di sezione per test (strisce e andamento storico)
# ----------------------------------------------------------------------------
TITOLI_TEST = {
    "IMTP": "IMTP — Isometric Mid-Thigh Pull",
    "Squat Jump": "SJ — Squat Jump",
    "CMJ": "CMJ — Countermovement Jump",
    "CMJ Rebound": "CMRJ — Countermovement Rebound Jump",
    "Indici": "Indici (DSI, EUR)",
}


def titolo_test(test):
    """Titolo esteso del test, mostrato sopra il gruppo dei suoi grafici."""
    return TITOLI_TEST.get(test, str(test))


ORDINE_TEST = {t: i for i, t in enumerate(["CMJ", "CMJ Rebound", "Squat Jump", "IMTP", "Indici"])}


def ordina_per_test(elementi, test_di, ordine_test):
    """Raggruppa per test (ordine fisso: CMJ, CMRJ, SJ, IMTP, Indici),
    mantenendo l'ordine scelto dall'utente all'interno di ciascun test."""
    return sorted(elementi, key=lambda x: ordine_test.get(test_di(x), len(ordine_test)))


# ----------------------------------------------------------------------------
# Andamento storico di una metrica
# ----------------------------------------------------------------------------
# Grafico di controllo: una linea che unisce le sedute, banda ±SWC attorno al
# riferimento, barra d'errore per seduta. Tre scelte di progetto che vale la
# pena tenere a mente:
#
#  - ASSE X A SPAZIATURA UNIFORME. Le sedute sono equidistanti anche se nella
#    realtà non lo sono: la geometria resta pulita. L'informazione temporale
#    non si perde, sta nella data sotto ogni tick e nella barra dei periodi.
#
#  - BARRA DEI PERIODI sopra l'asse, larga quanto il tempo realmente
#    trascorso. Il colore codifica SOLO il tempo, in grigi: i punti usano già
#    il colore per l'attendibilità, e due scale cromatiche sullo stesso
#    grafico si annullano a vicenda.
#
#  - RIFERIMENTO SCELTO DALL'ALLENATORE (checkbox per seduta). Le spunte
#    definiscono COSA è il riferimento; TE e SWC restano calcolati su TUTTE le
#    sedute caricate, perché misurano quanto oscilla l'atleta — una proprietà
#    sua, non del confronto. Se dipendessero dalle spunte, numeratore e
#    denominatore si muoverebbero insieme e nessuno capirebbe più cosa ha
#    causato cosa.

# Scala CONTINUA per la barra dei periodi: da chiaro (sedute ravvicinate) a
# scuro (lungo intervallo). Toni caldi neutri di proposito — i blu sono già
# usati dalla scala di ampiezza e verde/giallo/rosso dall'attendibilità:
# una terza scala che riusasse quelle tinte renderebbe ambiguo il colore.
# Gli ancoraggi sono in giorni; fra due ancoraggi si interpola linearmente.
GAP_SCALA = [
    (0, "#f2f1ea"),     # stessa settimana
    (14, "#dcd6c2"),    # 2 settimane
    (42, "#bdae8c"),    # 6 settimane
    (90, "#9a8763"),    # 3 mesi
    (180, "#77653f"),   # 6 mesi e oltre
]


def _lerp_hex(c1, c2, t):
    """Interpolazione lineare fra due colori esadecimali."""
    a = tuple(int(c1.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4))
    b = tuple(int(c2.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4))
    return "#" + "".join(f"{round(x + (y - x) * t):02x}" for x, y in zip(a, b))


def _gap_colore(giorni):
    """Colore della barra per un intervallo in giorni, su scala continua."""
    if giorni is None:
        return GAP_SCALA[0][1]
    g = max(0.0, float(giorni))
    if g >= GAP_SCALA[-1][0]:
        return GAP_SCALA[-1][1]
    for (g0, c0), (g1, c1) in zip(GAP_SCALA, GAP_SCALA[1:]):
        if g <= g1:
            t = 0.0 if g1 == g0 else (g - g0) / (g1 - g0)
            return _lerp_hex(c0, c1, t)
    return GAP_SCALA[-1][1]


def _nn(v):
    """None se il valore è mancante. Serve perché pandas trasforma i None
    delle colonne numeriche in NaN, e `is not None` non li intercetta più."""
    if v is None:
        return None
    try:
        return None if pd.isna(v) else v
    except (TypeError, ValueError):
        return v


def serie_storica(mid, all_long):
    """Serie di una metrica su tutte le sedute (storico + attuale), ordinata.
    Ritorna un DataFrame con session, data, mean, ordine."""
    g = all_long[all_long["metric_id"] == mid]
    if g.empty:
        return pd.DataFrame()
    g = (g.drop_duplicates("session")
           .sort_values("ordine")[["session", "data", "mean", "sd", "n",
                                   "metrica", "unit"]]
           .reset_index(drop=True))
    return g[g["mean"].apply(lambda v: isinstance(v, (int, float)) and not math.isnan(v))]


def prepara_storico(serie, sel_mask, swc, te_cv, log_scale=False):
    """Calcola riferimento, incertezza ed esito per ogni seduta.

    sel_mask: lista di bool, quali sedute compongono il riferimento.
    Ritorna (DataFrame arricchito, dict di metadati) oppure (df vuoto, {})."""
    if serie.empty:
        return pd.DataFrame(), {}
    valori = list(serie["mean"])
    scelti = [v for v, s in zip(valori, sel_mask) if s]
    if not scelti:
        return pd.DataFrame(), {}

    k = len(scelti)
    riferimento = sum(scelti) / k
    if riferimento == 0:
        return pd.DataFrame(), {}

    # Incertezza. Nel modo "misura" ogni seduta ha il proprio errore, quindi
    # la barra cambia da punto a punto: una seduta con salti molto simili è
    # nota meglio di una con salti dispersi, e il grafico deve mostrarlo.
    df = serie.copy().reset_index(drop=True)
    se_rif, se_punti = None, None
    se_ogni = [se_entro_seduta(r["sd"], r["n"], r["mean"]) for _, r in df.iterrows()]
    se_scelti = [x for x, sel in zip(se_ogni, sel_mask) if sel and x]
    if all(x for x in se_ogni) and se_scelti:
        se_rif = math.sqrt(sum(x * x for x in se_scelti)) / len(se_scelti)
        se_punti = [math.sqrt(x ** 2 + se_rif ** 2) for x in se_ogni]

    if se_punti is None:
        te_abs = (te_cv / 100.0) if log_scale else (te_cv / 100.0 * abs(riferimento))
        margine_unico = Z90 * te_abs * math.sqrt(1.0 + 1.0 / k)
        margini = [margine_unico] * len(df)
    else:
        fattore = (1.0 / abs(riferimento)) if log_scale else 1.0
        margini = [Z90 * x * fattore for x in se_punti]
    margine = sum(margini) / len(margini)  # valore medio, per le didascalie

    df["riferimento"] = riferimento
    df["nel_riferimento"] = list(sel_mask)
    df["margine"] = margini

    esiti, colori, delta_rif, delta_step, giorni = [], [], [], [], []
    prec_val, prec_data = None, None
    for _, r in df.iterrows():
        v = r["mean"]
        d = math.log(v / riferimento) if (log_scale and v > 0) else (v - riferimento)
        delta_rif.append((math.exp(d) - 1) * 100.0 if log_scale else d)

        marg_i = margini[len(esiti)]
        lo, hi = d - marg_i, d + marg_i
        if swc is None:
            esito = ESITO_ND
        elif lo > swc or hi < -swc:
            esito = ESITO_REALE
        elif lo > -swc and hi < swc:
            esito = ESITO_STABILE
        else:
            esito = ESITO_INCERTO
        esiti.append(esito)
        colori.append(ESITO_COLORI.get(esito, "#8d8d8d"))

        # Passo rispetto alla seduta precedente: qui gli estremi sono due
        # sedute singole, quindi SE = TE*sqrt(2), il confronto più incerto.
        if prec_val is None:
            delta_step.append(None)
        else:
            ds = (math.log(v / prec_val) if (log_scale and v > 0 and prec_val > 0)
                  else v - prec_val)
            delta_step.append((math.exp(ds) - 1) * 100.0 if log_scale else ds)

        data = r["data"]
        giorni.append((data - prec_data).days
                      if (data is not None and prec_data is not None) else None)
        prec_val, prec_data = v, data

    df["esito"] = esiti
    df["colore"] = colori
    df["delta_rif"] = delta_rif
    df["delta_step"] = delta_step
    df["giorni"] = giorni

    meta = dict(riferimento=riferimento, k=k, margine=margine, swc=swc,
                te_cv=te_cv, log=log_scale, per_punto=se_punti is not None,
                etichetta=str(serie["metrica"].iloc[0]),
                unita=str(serie["unit"].iloc[0] or ""),
                sessioni_rif=[s for s, m in zip(serie["session"], sel_mask) if m])
    return df, meta


def build_storico_chart(df, meta, height=480):
    """Grafico di controllo dell'andamento storico."""
    if df.empty or not meta:
        return None

    x = list(range(len(df)))
    rif, swc, margine = meta["riferimento"], meta["swc"], meta["margine"]

    def confine(mult):
        """Bordo della banda in unità native (in scala log le soglie sono
        moltiplicative e vanno convertite)."""
        return rif * math.exp(mult * swc) if meta["log"] else rif + mult * swc

    fig = go.Figure()

    # Banda di rilevanza ±SWC attorno al riferimento
    if swc:
        lo_b, hi_b = confine(-1), confine(1)
        fig.add_hrect(y0=min(lo_b, hi_b), y1=max(lo_b, hi_b),
                      fillcolor="#8d8d8d", opacity=0.16, line_width=0,
                      annotation_text="±SWC", annotation_position="top left",
                      annotation_font=dict(size=9, color="#484343"))
    fig.add_hline(y=rif, line_dash="dash", line_color=ACCENT, line_width=1.5)

    # Linea + punti. Il colore del marcatore è l'attendibilità rispetto al
    # riferimento; la linea resta neutra per non competere.
    hover = []
    for _, r in df.iterrows():
        pezzi = [f"<b>{r['session']}</b>"]
        pezzi.append(f"Valore: {fmt_valore(r['mean'])} {meta['unita']}".strip())
        pezzi.append(f"vs riferimento: {fmt_valore(r['delta_rif'], segno=True)}"
                     + ("%" if meta["log"] else ""))
        ds, gg = _nn(r["delta_step"]), _nn(r["giorni"])
        if ds is not None:
            pezzi.append(f"vs seduta prec.: {fmt_valore(ds, segno=True)}"
                         + ("%" if meta["log"] else ""))
        if gg is not None:
            pezzi.append(f"Giorni dalla precedente: {int(gg)}")
        pezzi.append(f"Attendibilità: {r['esito']}")
        if not r["nel_riferimento"]:
            pezzi.append("<i>esclusa dal riferimento</i>")
        hover.append("<br>".join(pezzi))

    fig.add_trace(go.Scatter(
        x=x, y=list(df["mean"]), mode="lines+markers",
        line=dict(color="#b8c4ca", width=2),
        marker=dict(size=13, color=list(df["colore"]),
                    line=dict(width=2, color="white"),
                    # Vuoto = fa parte del riferimento, pieno = ne è fuori.
                    # Il punto "sotto esame" (di norma la seduta attuale) è
                    # quello escluso, quindi è giusto che sia il più pieno.
                    symbol=["circle-open" if s else "circle"
                            for s in df["nel_riferimento"]]),
        error_y=dict(type="data", array=list(df["margine"]), color="#8d8d8d",
                     thickness=1.4, width=6),
        text=hover, hovertemplate="%{text}<extra></extra>", showlegend=False,
    ))

    # Legenda: due tracce fittizie (nessun punto reale) che spiegano la
    # differenza pieno/vuoto. Il colore qui e' neutro di proposito: nei punti
    # veri il colore codifica l'attendibilita', non l'appartenenza.
    for simbolo, testo in (("circle-open", "Seduta nel riferimento"),
                           ("circle", "Seduta esclusa dal riferimento")):
        fig.add_trace(go.Scatter(
            x=[None], y=[None], mode="markers", name=testo,
            marker=dict(size=11, color="#8d8d8d", symbol=simbolo,
                        line=dict(width=2, color="#8d8d8d")),
            hoverinfo="skip", showlegend=True,
        ))

    # Barra dei periodi: rettangoli in coordinate paper appena sotto l'asse,
    # larghi quanto il tempo realmente trascorso fra due sedute.
    giorni_validi = [g for g in (_nn(v) for v in df["giorni"]) if g is not None]
    max_gap = max(giorni_validi) if giorni_validi else None
    for i in range(1, len(df)):
        g = _nn(df["giorni"].iloc[i])
        colore = _gap_colore(g)
        # La larghezza del rettangolo è proporzionale al gap, centrata nel
        # segmento: un intervallo lungo si vede come una barra lunga.
        frazione = (g / max_gap) if (g is not None and max_gap) else 1.0
        centro = i - 0.5
        semi = max(0.06, 0.5 * frazione)
        fig.add_shape(type="rect", xref="x", yref="paper",
                      x0=centro - semi, x1=centro + semi, y0=-0.33, y1=-0.26,
                      fillcolor=colore, line_width=0)
        if g is not None:
            # Sopra i ~45 giorni il fondo diventa scuro: il testo passa a
            # bianco, altrimenti sparisce.
            fig.add_annotation(x=centro, y=-0.295, xref="x", yref="paper",
                               text=f"{int(g)} g", showarrow=False, yanchor="middle",
                               font=dict(size=9,
                                         color="#ffffff" if g >= 45 else "#3d3d3d"))

    # Il nome della seduta spesso E' gia' la data (l'etichetta viene dalla
    # colonna "Data test" del CSV): in quel caso aggiungerla sotto la
    # duplicherebbe. La seconda riga compare solo se dice qualcosa di nuovo.
    etichette_x = []
    for _, r in df.iterrows():
        nome_s = str(r["session"])
        d = _nn(r["data"])
        if d is None:
            etichette_x.append(nome_s)
            continue
        iso = d.strftime("%d/%m/%Y")
        breve = d.strftime("%d/%m/%y")
        if iso in nome_s or breve in nome_s:
            etichette_x.append(iso)
        else:
            etichette_x.append(f"{nome_s}<br>{breve}")

    titolo = f"<b>{meta['etichetta']}</b>"
    if meta["unita"]:
        titolo += f" ({meta['unita']})"
    sotto = (f"riferimento: media di {meta['k']} sedute"
             if meta["k"] > 1 else f"riferimento: {meta['sessioni_rif'][0]}")

    fig.update_layout(
        title=dict(text=f"{titolo}<br><span style='font-size:12px;opacity:0.7'>{sotto}</span>",
                   x=0.5, xanchor="center", font=dict(size=15, color=TEXT_COLOR)),
        xaxis=dict(tickmode="array", tickvals=x, ticktext=etichette_x,
                   range=[-0.5, len(df) - 0.5], showgrid=False,
                   tickfont=dict(size=10)),
        yaxis=dict(title=dict(text=meta["unita"] or "", standoff=10),
                   showgrid=True, gridcolor="#eef2f4", zeroline=False),
        height=height, margin=dict(t=100, b=125, l=60, r=25),
        plot_bgcolor=BG_COLOR, paper_bgcolor=BG_COLOR, font=dict(color=TEXT_COLOR),
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0,
                    font=dict(size=11)),
    )
    return fig


# ----------------------------------------------------------------------------
# Confronto dei profili di forza (T-score) e degli indici DSI / EUR
# ----------------------------------------------------------------------------
# I T-score sono già normalizzati e direzionali (per le metriche dove meno è
# meglio lo z viene invertito in z_t_score), quindi un T più alto è sempre
# migliore: qui, a differenza delle metriche grezze, "migliore" ha senso.
#
# La scala di ampiezza si traduce senza inventare nulla: T = 50 + 10*z e
# SWC = 0,2 SD, quindi 1 SWC = 2 punti di T-score. Piccolo = 2 punti,
# moderato = 6, grande = 12, molto grande = 20.
SWC_IN_PUNTI_T = 10.0 * SWC_FACTOR


def colore_delta_t(delta):
    """Verde se il T-score sale, rosso se scende. Qui il colore e' legittimo
    (a differenza delle metriche grezze): il T-score e' gia' orientato alla
    prestazione, quindi piu' alto e' sempre migliore."""
    if delta is None or (isinstance(delta, float) and math.isnan(delta)):
        return "#8d8d8d"
    return "#2E7D32" if delta > 0 else ("#B00020" if delta < 0 else "#8d8d8d")


def comp_score_specs(pop_dict, sesso_atleta):
    """{metric_id: (categoria, media_pop, sd_pop, lower_is_better)} per le
    sole metriche curate che hanno un T-score."""
    is_female = bool(sesso_atleta) and str(sesso_atleta).upper().startswith("F")
    specs = {}
    for m in METRICS:
        if (m.get("kind") != "score" or not m.get("pop_key") or not m.get("jump_type")
                or m["pop_key"] not in pop_dict):
            continue
        pop = pop_dict[m["pop_key"]]
        media = pop["mean_f"] if is_female else pop["mean_m"]
        sd = pop["sd_f"] if is_female else pop["sd_m"]
        if not sd:
            continue
        mid = _metric_id(m["jump_type"], m.get("raw_var") or m["key"])
        specs[mid] = (m["category"], float(media), float(sd), bool(m["lower_is_better"]))
    return specs


def profili_per_sessione(long_df, specs):
    """(profili, t_per_metrica): profili è {sessione: {categoria: T medio}},
    t_per_metrica è {sessione: {metric_id: T}}. Stessa media semplice per
    categoria usata da profilo_forza() nella PARTE 4."""
    profili, t_metrica = {}, {}
    for sessione, g in long_df.groupby("session", sort=False):
        per_cat, per_mid = {}, {}
        for _, riga in g.iterrows():
            spec = specs.get(riga["metric_id"])
            if not spec or riga["mean"] is None:
                continue
            categoria, media_pop, sd_pop, lower = spec
            _z, t = z_t_score(riga["mean"], media_pop, sd_pop, lower)
            if t is None:
                continue
            per_mid[riga["metric_id"]] = t
            per_cat.setdefault(categoria, []).append(t)
        if per_cat:
            profili[sessione] = {c: sum(v) / len(v) for c, v in per_cat.items()}
            t_metrica[sessione] = per_mid
    return profili, t_metrica


def confronta_profili(prof_cur, prof_hist, t_hist, hist_long, specs):
    """Confronta il profilo attuale con la media dei profili precedenti.

    L'incertezza in punti T viene ricavata dall'errore entro seduta delle
    metriche che compongono la categoria: SE_T = (SE assoluto / SD popolazione)
    * 10. Non viene divisa per la radice del numero di metriche, perché le
    metriche di una stessa categoria sono correlate fra loro: la stima resta
    prudente."""
    if not prof_cur or not prof_hist:
        return pd.DataFrame()

    medie_rif = {}
    for _sess, per_cat in prof_hist.items():
        for categoria, t in per_cat.items():
            medie_rif.setdefault(categoria, []).append(t)

    # TE in punti T per categoria, dalle metriche che la compongono
    medie_metrica = hist_long.groupby("metric_id")["mean"].mean().to_dict()
    te_per_cat = {}
    for mid, (categoria, _media_pop, sd_pop, _lower) in specs.items():
        media = medie_metrica.get(mid)
        if media is None or not sd_pop:
            continue
        righe = hist_long[hist_long["metric_id"] == mid]
        se_list = [x for x in (se_entro_seduta(r["sd"], r["n"], r["mean"])
                               for _, r in righe.iterrows()) if x]
        se_abs = (sum(se_list) / len(se_list) if se_list
                  else te_cv_default(mid) / 100.0 * abs(media))
        te_per_cat.setdefault(categoria, []).append(se_abs / sd_pop * 10.0)

    rows = []
    for categoria in CATEGORIES:
        if categoria not in prof_cur or categoria not in medie_rif:
            continue
        storici = medie_rif[categoria]
        k = len(storici)
        media_rif = sum(storici) / k
        t_cur = prof_cur[categoria]
        delta = t_cur - media_rif

        te_punti = te_per_cat.get(categoria)
        if te_punti:
            margine = Z90 * (sum(te_punti) / len(te_punti)) * math.sqrt(1.0 + 1.0 / k)
        else:
            margine = None

        if margine is None:
            esito, ampiezza, colore = ESITO_ND, "—", None
            rapporto = None
        else:
            rapporto = abs(delta) / SWC_IN_PUNTI_T
            etichetta, colore = scala_ampiezza(rapporto)
            freccia = "🔺" if delta > 0 else ("🔻" if delta < 0 else "▪")
            ampiezza = f"{freccia} {etichetta}"
            lo, hi = delta - margine, delta + margine
            if lo > SWC_IN_PUNTI_T or hi < -SWC_IN_PUNTI_T:
                esito = ESITO_REALE
            elif lo > -SWC_IN_PUNTI_T and hi < SWC_IN_PUNTI_T:
                esito = ESITO_STABILE
            else:
                esito = ESITO_INCERTO

        migliore = max(storici)
        sess_migliore = next((s for s, p in prof_hist.items()
                              if p.get(categoria) == migliore), "—")
        banda_cur, _colore_banda = banda_da_tscore(t_cur)

        rows.append({
            "Categoria": categoria,
            "Attendibilità": esito,
            "Cambiamento": ampiezza,
            "T medio rif.": media_rif,
            "T attuale": t_cur,
            "Delta T": delta,
            "Valutazione attuale": banda_cur or "—",
            "T migliore": migliore,
            "Seduta migliore": sess_migliore,
            "N sedute": k,
            "_margine": margine,
            "_rapporto": rapporto,
            "_colore": colore,
        })
    return pd.DataFrame(rows)


def build_profili_radar(cats, serie, nome_atleta):
    """Radar con più serie sovrapposte (attuale, media storico, migliore),
    sulla stessa geometria di build_radar_chart della PARTE 4."""
    if not cats:
        return None
    cats_closed = [_wrap_label(c) for c in cats] + [_wrap_label(cats[0])]
    fig = go.Figure()
    fig.add_trace(go.Scatterpolar(
        r=[50] * (len(cats) + 1), theta=cats_closed, mode="lines",
        line=dict(color="rgba(233,74,38,0.6)", dash="dash"),
        name="Media popolazione (T=50)",
    ))
    stili = [
        ("Media sedute precedenti", "#8d8d8d", "dot", None, 2),
        ("Miglior seduta", "#2E7D32", "dash", None, 2),
        (f"{nome_atleta} — attuale", PRIMARY, "solid", "rgba(11,182,255,0.25)", 3),
    ]
    for (etichetta, colore, tratto, riempimento, spessore) in stili:
        valori = serie.get(etichetta)
        if not valori:
            continue
        vals = [valori.get(c) for c in cats]
        if any(v is None for v in vals):
            continue
        fig.add_trace(go.Scatterpolar(
            r=vals + [vals[0]], theta=cats_closed, name=etichetta,
            fill="toself" if riempimento else None, fillcolor=riempimento,
            line=dict(color=colore, width=spessore, dash=tratto),
        ))
    fig.update_layout(
        polar=dict(
            domain=dict(x=[0.12, 0.88], y=[0.08, 0.92]),
            radialaxis=dict(range=[0, 100], showticklabels=True, ticks="",
                            tickfont=dict(color=TEXT_COLOR)),
        ),
        font=dict(color=TEXT_COLOR), paper_bgcolor=BG_COLOR,
        showlegend=True, height=560, margin=dict(t=60, b=60, l=60, r=60),
        legend=dict(orientation="h", yanchor="bottom", y=-0.12, x=0),
    )
    return fig


def build_indice_zone_strip(key, valore_attuale, valore_rif, thr_low, thr_high,
                            decimals=3, height=140):
    """Striscia a zone di profilo per DSI/EUR (stesse bande e stesse etichette
    di ratio_band_strip della PARTE 4), con due marcatori: media delle sedute
    precedenti e test attuale. Serve a vedere se l'atleta ha cambiato ZONA,
    non solo se il numero si è mosso."""
    if valore_attuale is None or not thr_low or not thr_high:
        return None
    valori = [v for v in (valore_attuale, valore_rif) if v is not None]
    lo = min([thr_low - 0.2] + [v - 0.05 for v in valori])
    hi = max([thr_high + 0.2] + [v + 0.05 for v in valori])
    labels = INDEX_ZONE_LABELS[key]

    fig = go.Figure()
    for a, b, i in ((lo, thr_low, 0), (thr_low, thr_high, 1), (thr_high, hi, 2)):
        fig.add_vrect(x0=a, x1=b, fillcolor=INDEX_ZONE_COLORS[i], opacity=0.35, line_width=0)
        fig.add_annotation(x=(a + b) / 2, y=1.10, xref="x", yref="paper", yanchor="bottom",
                           text=_wrap_label(labels[i], 22), showarrow=False,
                           font=dict(size=11, color="#484343"))
    if valore_rif is not None:
        fig.add_trace(go.Scatter(
            x=[valore_rif], y=[0], mode="markers", name="Media precedenti",
            marker=dict(size=12, color="#8d8d8d", symbol="circle-open",
                        line=dict(width=2.5, color="#8d8d8d")),
            cliponaxis=False,
            hovertemplate=f"Media precedenti: %{{x:.{decimals}f}}<extra></extra>",
        ))
    fig.add_trace(go.Scatter(
        x=[valore_attuale], y=[0], mode="markers+text", name="Attuale",
        marker=dict(size=13, color=PRIMARY, symbol="diamond",
                    line=dict(width=1.5, color="white")),
        text=[f"{valore_attuale:.{decimals}f}"], textposition="bottom center",
        textfont=dict(color=TEXT_COLOR, size=13), cliponaxis=False,
        hovertemplate=f"Attuale: %{{x:.{decimals}f}}<extra></extra>",
    ))
    fig.update_layout(
        xaxis=dict(range=[lo, hi], showgrid=False, tickvals=[thr_low, thr_high],
                   ticktext=[f"{thr_low:.2f}", f"{thr_high:.2f}"], automargin=True,
                   title=dict(text=key.upper(), standoff=12)),
        yaxis=dict(range=[-1, 1], showgrid=False, showticklabels=False, zeroline=False),
        height=height, margin=dict(t=64, b=50, l=20, r=20),
        showlegend=True, legend=dict(orientation="h", yanchor="bottom", y=-0.55, x=0),
        plot_bgcolor=BG_COLOR, paper_bgcolor=BG_COLOR, font=dict(color=TEXT_COLOR),
    )
    return fig



# ----------------------------------------------------------------------------
# UI della scheda
# ----------------------------------------------------------------------------
def render_comparazione(parsed_files, nome, sesso, periodo, results):
        if not parsed_files:
            st.info("Carica dei file dalla barra laterale per confrontare questo test con i precedenti.")
        else:
            st.markdown(
                "Confronta il test attuale con lo storico dell'atleta. Carica i CSV prodotti da "
                "**🔍 Dettaglio Test → Esporta TUTTE le metriche grezze**."
            )
            st.caption(
                "Il confronto statistico è contro la MEDIA delle sedute precedenti: la media di più "
                "sedute è più precisa di una singola, quindi l'incertezza si stringe. Il miglior "
                "valore storico è riportato solo come riferimento descrittivo — un record è quasi "
                "sempre il giorno in cui l'atleta era in forma e il rumore ha aiutato, quindi "
                "confrontarcisi statisticamente darebbe 'peggiorato' quasi sempre."
            )

            comp_files = st.file_uploader(
                "CSV delle sedute precedenti (export completo)", type=["csv"],
                accept_multiple_files=True, key="comp_files",
                help="Uno o più file 'Scarica export completo (.csv)' dello STESSO atleta. "
                     "Da 4 sedute in su l'app stima il rumore della misura sull'atleta stesso.",
            )

            if not comp_files:
                st.info("Carica almeno un CSV di riferimento per iniziare il confronto.")
            else:
                storico, errori = [], []
                for i, f in enumerate(comp_files):
                    try:
                        storico.append(parse_full_export_csv(
                            io.BytesIO(f.getvalue()), f.name.rsplit(".", 1)[0], ordine=i))
                    except Exception as e:
                        errori.append(f"**{f.name}**: {e}")
                for msg in errori:
                    st.warning(msg)

                if storico:
                    storico = ordina_sessioni(storico)
                    hist_long = pd.concat(storico, ignore_index=True)
                    cur_long = reps_to_long(current_reps_by_type(parsed_files), "Attuale",
                                            _parse_data_test(periodo), ordine=len(storico))
                    pop_sd_map = comp_pop_sd_map(st.session_state["pop"], sesso)

                    senza_data = [f["session"].iloc[0] for f in storico if f["data"].iloc[0] is None]
                    if senza_data:
                        st.caption(
                            "Sedute senza data leggibile, ordinate come caricate: "
                            + ", ".join(senza_data)
                        )

                    all_long = pd.concat([hist_long, cur_long], ignore_index=True)
                    n_sedute = len(storico)

                    st.caption(
                        f"{n_sedute} sedute di riferimento. L'incertezza di ogni seduta è la "
                        "dispersione fra le sue ripetizioni (dev.std diviso radice di n): misura "
                        "l'errore DENTRO la giornata, non la variabilità fra giorni diversi — "
                        "quella, per un atleta che si allena, è cambiamento reale e non rumore. "
                        "La soglia di rilevanza è sempre 0,2 × dev.std di popolazione."
                    )

                    res = compare_to_history(cur_long, hist_long, pop_sd_map)

                    if res.empty:
                        st.error(
                            "Nessuna metrica in comune tra il test attuale e lo storico. Verifica "
                            "che i CSV appartengano allo stesso atleta e agli stessi tipi di test."
                        )
                    else:
                        st.caption(f"Metriche confrontabili trovate: {len(res)}.")
                        primarie_ids = {
                            _metric_id(m["jump_type"], m.get("raw_var") or m["key"])
                            for m in METRICS
                            if m["key"] in COMP_DEFAULT_PRIMARY_KEYS and m.get("jump_type")
                        }
                        opzioni = list(res["_display"])
                        default_sel = [r["_display"] for _, r in res.iterrows()
                                       if r["_mid"] in primarie_ids]

                        # Le metriche spuntate in "Metriche incluse nel Profilo di
                        # Forza" (Dettaglio Test) entrano sempre nell'analisi: vengono
                        # aggiunte alla selezione corrente a ogni rerun.
                        prof_ids = {comp_id_da_risultato(r["key"]) for r in results
                                    if r["t"] is not None and nel_profilo(r)}
                        prof_display = [r["_display"] for _, r in res.iterrows()
                                        if r["_mid"] in prof_ids]
                        if "comp_scelte" in st.session_state:
                            gia = [d for d in st.session_state["comp_scelte"] if d in opzioni]
                            st.session_state["comp_scelte"] = list(dict.fromkeys(gia + prof_display))
                            default_scelte = None
                        else:
                            default_scelte = (list(dict.fromkeys(prof_display + default_sel))
                                              or opzioni[: min(5, len(opzioni))])

                        scelte = st.multiselect(
                            "Metriche in analisi", opzioni,
                            default=default_scelte,
                            key="comp_scelte",
                            help="L'unico elenco completo delle metriche disponibili. Tutto il "
                                 "resto della scheda — tabella, strisce, andamento storico — "
                                 "lavora su questa selezione. Sceglile PRIMA di guardare i dati: "
                                 "è ciò che distingue una decisione da un inseguimento del rumore.",
                        )
                        vista = res[res["_display"].isin(scelte)] if scelte else res

                        conteggi = vista["Attendibilità"].value_counts()
                        k1, k2, k3, k4 = st.columns(4)
                        k1.metric(ESITO_REALE, int(conteggi.get(ESITO_REALE, 0)))
                        k2.metric(ESITO_STABILE, int(conteggi.get(ESITO_STABILE, 0)))
                        k3.metric(ESITO_INCERTO, int(conteggi.get(ESITO_INCERTO, 0)))
                        k4.metric(ESITO_ND, int(conteggi.get(ESITO_ND, 0)))

                        tab_df = vista[[
                            "Test", "Metrica", "Unità", "Media rif.", "Attuale",
                            "Delta", "Delta %", "Cambiamento", "Attendibilità",
                            "Migliore", "% del migliore", "Seduta migliore", "N sedute",
                        ]].reset_index(drop=True)

                        def _stile_comp(row):
                            stili = [""] * len(row)
                            col = ESITO_COLORI.get(row["Attendibilità"])
                            if col:
                                stili[tab_df.columns.get_loc("Attendibilità")] = (
                                    f"color:{col}; font-weight:600;")
                            return stili

                        st.dataframe(
                            tab_df.style.apply(_stile_comp, axis=1).format({
                                "Media rif.": fmt_valore, "Attuale": fmt_valore,
                                "Delta": lambda v: fmt_valore(v, segno=True),
                                "Delta %": "{:+.1f}%", "Migliore": fmt_valore,
                                "% del migliore": "{:.0f}%",
                            }, na_rep="—"),
                            use_container_width=True, hide_index=True,
                        )

                        nd = [r["_display"] for _, r in vista.iterrows()
                              if r["Attendibilità"] == ESITO_ND]
                        if nd:
                            st.warning(
                                f"⚠️ {len(nd)} metriche non valutabili: manca la norma di "
                                "popolazione, quindi non esiste una soglia di rilevanza. Il valore "
                                "si vede comunque, ma non si può dire se lo scostamento conti."
                            )

                        # --- Strisce per metrica ---
                        st.markdown("#### Scostamento dalla media, metrica per metrica")
                        st.caption(
                            "Il rombo pieno azzurro è il test attuale con la sua barra di "
                            "incertezza; il rombo vuoto arancione è la media delle sedute "
                            "precedenti. Le fasce "
                            "colorate sono l'entità dello scostamento (trascurabile, piccolo, moderato, "
                            "grande), simmetriche sopra e sotto la media."
                        )
                        strip_opzioni = [r["_display"] for _, r in vista.iterrows()
                                         if r["_swc"] is not None]
                        strip_scelte = st.multiselect(
                            "Metriche da graficare", strip_opzioni,
                            default=strip_opzioni[: min(5, len(strip_opzioni))],
                            key="comp_strip",
                            help="Le metriche non valutabili (⚠) non compaiono: senza soglia di "
                                 "rilevanza non c'è scala da disegnare.",
                        )
                        if not strip_opzioni:
                            st.info("Nessuna metrica con una soglia di rilevanza disponibile.")
                        ordine_test = ORDINE_TEST
                        righe_strip = [r for _, r in vista[vista["_display"].isin(strip_scelte)].iterrows()]
                        test_corrente = None
                        for riga in ordina_per_test(righe_strip, lambda r: r["Test"], ordine_test):
                            fig_strip = build_swc_strip(riga)
                            if fig_strip:
                                if riga["Test"] != test_corrente:
                                    test_corrente = riga["Test"]
                                    st.markdown(f"##### {titolo_test(test_corrente)}")
                                st.plotly_chart(fig_strip, use_container_width=True)

                        # --- Andamento storico -----------------------------------
                        storico_export = []
                        st.markdown("---")
                        st.markdown("#### Andamento storico")
                        st.caption(
                            "Ogni punto è una seduta, con la sua barra di incertezza. La banda grigia "
                            "è la zona di rilevanza (±SWC) attorno al riferimento; la linea tratteggiata "
                            "arancione è il riferimento stesso. La barra grigia sotto l'asse è il tempo "
                            "trascorso fra una seduta e l'altra: più è larga, più tempo è passato."
                        )

                        # Le opzioni vengono da `vista`, non da `res`: l'elenco
                        # completo sta solo in "Metriche in analisi" sopra.
                        display_to_mid = {r["_display"]: r["_mid"] for _, r in vista.iterrows()}
                        storico_opzioni = list(display_to_mid)
                        default_storico = [d for d in storico_opzioni
                                           if display_to_mid[d] in primarie_ids]
                        scelte_storico = st.multiselect(
                            "Metriche da seguire nel tempo", storico_opzioni,
                            default=default_storico or storico_opzioni[: min(2, len(storico_opzioni))],
                            key="comp_storico",
                            help="Sottoinsieme delle metriche in analisi: un grafico per ciascuna.",
                        )

                        # Le sedute sono le stesse per tutte le metriche, quindi le
                        # checkbox del riferimento stanno FUORI dal ciclo: una sola
                        # scelta, valida per ogni grafico.
                        sedute_ordinate = (all_long.drop_duplicates("session")
                                           .sort_values("ordine")["session"].tolist())
                        if scelte_storico and sedute_ordinate:
                            st.markdown("**Sedute che compongono il riferimento**")
                            st.caption(
                                "Di default tutte le precedenti, cioè la media dell'atleta. "
                                "Deselezionandone qualcuna cambia solo il riferimento: il rumore "
                                "della misura e la soglia di rilevanza restano calcolati su tutte "
                                "le sedute caricate."
                            )
                            chk_cols = st.columns(min(len(sedute_ordinate), 6))
                            sel_per_seduta = {}
                            for i, nome_seduta in enumerate(sedute_ordinate):
                                e_attuale = nome_seduta == "Attuale"
                                sel_per_seduta[nome_seduta] = chk_cols[i % len(chk_cols)].checkbox(
                                    str(nome_seduta)[:12], value=not e_attuale,
                                    key=f"rif_seduta_{i}",
                                    help="Seduta attuale: normalmente esclusa dal riferimento, "
                                         "altrimenti si confronterebbe con sé stessa."
                                         if e_attuale else None,
                                )

                            if not any(sel_per_seduta.values()):
                                st.warning(
                                    "⚠️ Nessuna seduta selezionata: senza riferimento non c'è "
                                    "confronto da calcolare. Spuntane almeno una."
                                )
                            else:
                                display_to_test = {r["_display"]: r["Test"] for _, r in vista.iterrows()}
                                test_corrente = None
                                for scelta_st in ordina_per_test(scelte_storico, display_to_test.get,
                                                                 ordine_test):
                                    if display_to_test.get(scelta_st) != test_corrente:
                                        test_corrente = display_to_test.get(scelta_st)
                                        st.markdown(f"##### {titolo_test(test_corrente)}")
                                    mid_st = display_to_mid[scelta_st]
                                    serie_st = serie_storica(mid_st, all_long)
                                    if len(serie_st) < 2:
                                        st.info(f"**{scelta_st}** — servono almeno due sedute "
                                                "con questa metrica per disegnare un andamento.")
                                        continue

                                    # La maschera va riallineata alla serie: una
                                    # metrica puo' mancare in qualche seduta.
                                    mask_st = [bool(sel_per_seduta.get(sess, False))
                                               for sess in serie_st["session"]]
                                    if not any(mask_st):
                                        st.info(f"**{scelta_st}** — nessuna delle sedute spuntate "
                                                "contiene questa metrica.")
                                        continue

                                    riga_res = res[res["_mid"] == mid_st]
                                    swc_st = riga_res["_swc"].iloc[0] if len(riga_res) else None
                                    log_st = bool(riga_res["_log"].iloc[0]) if len(riga_res) else False
                                    te_st = te_cv_default(serie_st["metrica"].iloc[0])

                                    df_st, meta_st = prepara_storico(serie_st, mask_st, swc_st,
                                                                     te_st, log_st)
                                    fig_st = build_storico_chart(df_st, meta_st)
                                    if not fig_st:
                                        continue
                                    st.plotly_chart(fig_st, use_container_width=True)
                                    # Per il report salviamo df e meta gia' calcolati:
                                    # la figura viene ricostruita al momento
                                    # dell'export dalla stessa build_storico_chart.
                                    storico_export.append(dict(
                                        display=scelta_st, df=df_st.copy(), meta=meta_st,
                                        senza_swc=swc_st is None,
                                    ))
                                    if swc_st is None:
                                        st.warning(
                                            f"⚠️ **{scelta_st}** non ha una soglia di rilevanza: "
                                            "la linea si vede, ma i punti non possono ricevere "
                                            "un'attendibilità."
                                        )
                                    st.caption(
                                        f"Riferimento: {', '.join(map(str, meta_st['sessioni_rif']))}"
                                        + (" · barra propria per ogni seduta"
                                           if meta_st["per_punto"] else "")
                                        + f" · incertezza media ±{fmt_valore(meta_st['margine'])}"
                                    )

                        # --- Profili di forza (T-score) e indici -----------------
                        st.markdown("---")
                        st.markdown("#### Profilo di forza: attuale vs storico")
                        specs = comp_score_specs(st.session_state["pop"], sesso)
                        prof_hist, _t_hist = profili_per_sessione(hist_long, specs)
                        prof_cur_all, _t_cur = profili_per_sessione(cur_long, specs)
                        prof_cur = prof_cur_all.get("Attuale", {})

                        prof_res, prof_serie, prof_cats = pd.DataFrame(), None, []
                        if not prof_cur or not prof_hist:
                            st.info(
                                "Profilo non calcolabile: servono metriche con norme di popolazione "
                                "(IMTP, SJ, CMJ o CMJ RE) sia nel test attuale sia nello storico."
                            )
                        else:
                            cats_valide = [c for c in CATEGORIES
                                           if c in prof_cur and any(c in p for p in prof_hist.values())]
                            medie_cat, migliori_cat = {}, {}
                            for c in cats_valide:
                                valori = [p[c] for p in prof_hist.values() if c in p]
                                medie_cat[c] = sum(valori) / len(valori)
                                migliori_cat[c] = max(valori)

                            prof_cats = cats_valide
                            prof_serie = {
                                "Media sedute precedenti": medie_cat,
                                "Miglior seduta": migliori_cat,
                                f"{nome} — attuale": prof_cur,
                            }
                            # confronta_profili() viene chiamata PRIMA del radar
                            # perche' le card sotto il grafico usano il Delta T.
                            prof_res = confronta_profili(prof_cur, prof_hist, _t_hist, hist_long,
                                                         specs)
                            delta_per_cat = ({r["Categoria"]: r["Delta T"]
                                              for _, r in prof_res.iterrows()}
                                             if not prof_res.empty else {})

                            radar = build_profili_radar(cats_valide, prof_serie, nome)
                            if radar:
                                st.plotly_chart(radar, use_container_width=True)

                            # Card come nella scheda Profilo di Forza, piu' il solo
                            # scostamento in punti T: entita' e attendibilita' si
                            # leggono in tabella, qui servirebbero solo a togliere
                            # risalto al T-score.
                            cols_prof = st.columns(len(cats_valide))
                            for col, cat in zip(cols_prof, cats_valide):
                                t = prof_cur[cat]
                                banda, colore = banda_da_tscore(t)
                                d = delta_per_cat.get(cat)
                                html_card = (
                                    f"<div style='font-size:11px;letter-spacing:0.5px;"
                                    f"opacity:0.65'>{cat}</div>"
                                    f"<div style='font-size:17px;font-weight:600;"
                                    f"color:{TEXT_COLOR}'>{CATEGORY_QUALITY.get(cat, '')}</div>"
                                    f"<div style='font-size:28px;color:{colore}'>{t:.0f}</div>"
                                    f"<div style='font-size:13px;opacity:0.75'>{banda}</div>"
                                )
                                if d is not None:
                                    col_d = colore_delta_t(d)
                                    html_card += (f"<div style='font-size:15px;font-weight:600;"
                                                  f"color:{col_d}'>{d:+.1f} punti T</div>")
                                col.markdown(html_card, unsafe_allow_html=True)

                            st.caption(
                                "Il T-score è già orientato alla prestazione (per le metriche dove meno "
                                "è meglio lo scarto viene invertito), quindi qui un valore più alto è "
                                "sempre migliore. 1 soglia di rilevanza = 2 punti di T-score, perché "
                                "T = 50 + 10·z e la soglia vale 0,2 deviazioni standard."
                            )

                            if not prof_res.empty:
                                prof_df = prof_res[[
                                    "Categoria", "T medio rif.", "T attuale", "Delta T",
                                    "Cambiamento", "Attendibilità", "Valutazione attuale",
                                    "T migliore", "Seduta migliore", "N sedute",
                                ]].reset_index(drop=True)

                                def _stile_prof(row):
                                    stili = [""] * len(row)
                                    col = ESITO_COLORI.get(row["Attendibilità"])
                                    if col:
                                        stili[prof_df.columns.get_loc("Attendibilità")] = (
                                            f"color:{col}; font-weight:600;")
                                    return stili

                                st.dataframe(
                                    prof_df.style.apply(_stile_prof, axis=1).format({
                                        "T medio rif.": "{:.1f}", "T attuale": "{:.1f}",
                                        "Delta T": "{:+.1f}", "T migliore": "{:.1f}",
                                    }, na_rep="—"),
                                    use_container_width=True, hide_index=True,
                                )

                        # --- Indici di profilo DSI / EUR ------------------------
                        idx_export = []
                        idx_disponibili = [
                            k for k in ("dsi", "eur")
                            if _metric_id("indici", k) in set(cur_long["metric_id"])
                        ]
                        if idx_disponibili:
                            st.markdown("#### Indici di profilo: si è spostato di zona?")
                            st.caption(
                                "DSI ed EUR non hanno un verso migliore: quello che conta è se "
                                "l'atleta ha cambiato zona di profilo. Il rombo è il test attuale, "
                                "il cerchio vuoto la media delle sedute precedenti."
                            )
                            for key in idx_disponibili:
                                mid = _metric_id("indici", key)
                                att = cur_long[cur_long["metric_id"] == mid]["mean"]
                                rif = hist_long[hist_long["metric_id"] == mid]["mean"]
                                v_att = float(att.iloc[0]) if len(att) else None
                                v_rif = float(rif.mean()) if len(rif) else None
                                lo_thr, hi_thr = st.session_state["idx_thr"][key]
                                fig_idx = build_indice_zone_strip(key, v_att, v_rif, lo_thr, hi_thr)
                                if fig_idx is None:
                                    continue
                                zona_att, _c1 = zona_da_indice(key, v_att, lo_thr, hi_thr)
                                zona_rif, _c2 = zona_da_indice(key, v_rif, lo_thr, hi_thr)
                                if zona_att and zona_rif and zona_att != zona_rif:
                                    st.markdown(
                                        f"**{key.upper()}** — cambio di zona: "
                                        f"da *{zona_rif}* a *{zona_att}*."
                                    )
                                elif zona_att:
                                    st.markdown(f"**{key.upper()}** — zona invariata: *{zona_att}*.")
                                st.plotly_chart(fig_idx, use_container_width=True)
                                idx_export.append(dict(
                                    key=key, attuale=v_att, riferimento=v_rif,
                                    zona_att=zona_att, zona_rif=zona_rif,
                                    thr=(lo_thr, hi_thr),
                                ))

                        # Rende i risultati disponibili alla scheda 📄 Report: i
                        # generatori HTML/PDF leggono questa chiave se presente, e
                        # includono la sezione Comparazione. Le figure NON vengono
                        # salvate qui (sono ricostruite al momento dell'export dalle
                        # stesse funzioni), solo i dati.
                        st.session_state["comp_export"] = dict(
                            vista=vista.copy(), n_sedute=n_sedute,
                            prof=prof_res.copy() if not prof_res.empty else None,
                            prof_serie=prof_serie, prof_cats=list(prof_cats),
                            strip=[d for d in strip_scelte],
                            indici=idx_export,
                            storico=storico_export,
                        )

                        st.download_button(
                            "⬇️ Scarica comparazione (.csv)",
                            data=res.drop(columns=[c for c in res.columns if c.startswith("_")])
                                    .to_csv(index=False).encode("utf-8-sig"),
                            file_name=f"comparazione_{nome.replace(' ', '_')}_"
                                      f"{dt.date.today().isoformat()}.csv",
                            mime="text/csv",
                        )

                        with st.expander("ℹ️ Come si leggono questi risultati"):
                            st.markdown(
                                f"""
    **Attendibilità — il cambiamento è reale?**

    - **{ESITO_REALE}** — l'intervallo di incertezza sta interamente oltre la soglia
      di rilevanza: il cambiamento supera il rumore della misura ed è abbastanza
      grande da contare.
    - **{ESITO_STABILE}** — l'intervallo sta interamente dentro la soglia: quasi
      certamente NON c'è stato un cambiamento rilevante. È una conclusione, non un
      dato mancante: sapere che l'atleta è fermo è un'informazione.
    - **{ESITO_INCERTO}** — l'intervallo scavalca la soglia: i dati non bastano per
      decidere. Con poche sedute è l'esito più frequente, ed è la risposta onesta.
    - **{ESITO_ND}** — manca la norma di popolazione, quindi non esiste una soglia
      di rilevanza. Il valore resta leggibile, il giudizio no: sparirà quando le
      norme di laboratorio copriranno anche questa metrica.

    **Cambiamento — quanto si è spostato?** Scala di Hopkins in multipli della soglia:
    trascurabile (sotto 1), piccolo (1-3), moderato (3-6), grande (6-10), molto grande
    (oltre 10). La freccia dice solo la direzione: 🔺 il numero è salito, 🔻 è sceso.
    **L'app non dice se sia un bene**: per un tempo di contatto scendere è un
    progresso, per un'altezza di salto è un calo. Quel giudizio spetta a te.

    **Media rif.** — media delle sedute precedenti, esclusa l'attuale. Escluderla è
    necessario: altrimenti confronteresti il test con qualcosa che contiene sé
    stesso, e lo scostamento risulterebbe artificialmente più piccolo.

    **Migliore** — il miglior valore mai registrato, con la seduta in cui è avvenuto.
    È **descrittivo**: non ha né attendibilità né entità. Il massimo di più misure rumorose
    è per costruzione una sovrastima, quindi un confronto statistico con il record
    segnalerebbe un peggioramento quasi sempre, per un difetto del metodo e non
    dell'atleta. Usalo come riferimento pratico ("è al 92% del suo record").

    **La soglia di rilevanza (SWC)** vale sempre 0,2 × deviazione standard di
    popolazione, presa dalle norme della scheda ⚙️ Costanti. Non viene mai ricavata
    dallo storico dell'atleta: una soglia che dipendesse dalla sua ultima settimana
    si allargherebbe proprio quando lui migliora. Senza norma, l'esito è {ESITO_ND}.

    **L'incertezza** di ogni seduta è la dispersione fra le sue ripetizioni
    (dev.std ÷ radice di n). Misura l'errore dentro la giornata; la variabilità fra
    giorni diversi non ci entra, perché per un atleta che si allena quella è
    cambiamento reale, non rumore. Sotto le 3 ripetizioni — e per DSI ed EUR, che
    sono rapporti fra test — si ricade su valori di letteratura.

    **Indici di simmetria** — confrontati in valore assoluto: il segno indica il lato
    dominante, non la qualità, quindi passare da -2% a +6% è un aumento dello
    sbilanciamento anche se il numero "cresce".

    **DSI ed EUR** — sono rapporti fra test diversi: variazione e soglia sono in
    percentuale, non in unità assolute.
                                """
                            )