# -*- coding: utf-8 -*-
"""
Force Plate Test Report — app Streamlit (entry point)

Carica gli export XLSX di ForceMate/ForceDecks (IMTP, SJ, CMJ, CMJ RE),
calcola medie, T-score rispetto alla popolazione di riferimento e produce
profilo di forza, comparazione storica e report HTML/PDF.

Struttura (6 file):
  streamlit_app.py  — sidebar, stato di sessione, caricamento file, schede + ⚙️ Costanti
  data.py           — dati statici: norme, METRICS, catalogo e descrizioni metriche
  core.py           — tema, parsing, analisi (T-score/profilo), grafici
  tabs.py           — schede 🔍 Dettaglio Test e 📊 Profilo di Forza
  comparazione.py   — logica + scheda 🔀 Comparazione
  report.py         — report HTML/PDF + scheda 📄 Report
"""

from __future__ import annotations

import io
import math

import streamlit as st

from comparazione import (
    render_comparazione,
)
from core import (
    apply_brand_css,
    apply_type_override,
    build_results,
    guess_type_from_filename,
    parse_forcemate_workbook,
    parse_full_export_to_parsed,
    profilo_forza,
    sesso_da_file,
)
from data import (
    CONST_LABELS,
    DEFAULT_INDEX_THRESHOLDS,
    JUMP_TYPE_LABELS,
    NORM_SETS,
    _copia_norme,
    costanti_dataframe,
    genera_costanti_xlsx,
    parse_constants_workbook,
)
from report import (
    render_report,
)
from tabs import (
    render_dettaglio,
    render_profilo,
)

# ============================================================================
# PARTE 2 — CARICAMENTO FILE (sidebar upload)
# ============================================================================

st.set_page_config(page_title="Force Plate Test Report", page_icon="🏋️", layout="wide")

apply_brand_css()

st.sidebar.title("📂 Import dati")

uploaded = st.sidebar.file_uploader(
    "Carica i file XLSX esportati da ForceMate", type=["xlsx"], accept_multiple_files=True
)

st.sidebar.markdown("---")

csv_reload = st.sidebar.file_uploader(
    "🔁 Ricarica export completo (CSV)", type=["csv"],
    help="Carica un CSV precedentemente esportato dalla scheda 'Dettaglio Test' "
         "(forceplate_fulldata_...) per rivedere grafici e tabelle senza i file XLSX "
         "originali. Se presente, ha precedenza sui file XLSX caricati sotto.",
)


# Le costanti di popolazione (scheda "⚙️ Costanti") non dipendono dai file
# caricati: inizializziamo subito lo stato in modo che quella scheda sia
# consultabile/modificabile anche prima di caricare qualsiasi file.
def _on_norm_change():
    """Cambio di set: ricarica le norme e azzera le modifiche manuali della
    tabella (altrimenti il data_editor riapplicherebbe le vecchie)."""
    st.session_state["pop"] = _copia_norme(st.session_state["norm_set"])
    if "pop_editor" in st.session_state:
        del st.session_state["pop_editor"]


st.sidebar.markdown("---")
st.sidebar.radio(
    "📏 Norme di riferimento", list(NORM_SETS), key="norm_set",
    on_change=_on_norm_change,
    help="Popolazione usata per T-score, profilo di forza e soglie di rilevanza "
         "della Comparazione. Cambiare set sovrascrive le modifiche manuali "
         "fatte nella scheda Costanti.",
)

# Ricarica anche se lo stato in sessione è incompleto (es. creato da una
# versione precedente del file con un set di norme vuoto).
if ("pop" not in st.session_state
        or any(k not in st.session_state["pop"] for k in CONST_LABELS)):
    st.session_state["pop"] = _copia_norme(st.session_state["norm_set"])
    if "pop_editor" in st.session_state:
        del st.session_state["pop_editor"]
    st.session_state["pop"] = _copia_norme(st.session_state["norm_set"])

# Soglie degli indici a rapporto (DSI, EUR): modificabili nella scheda
# "⚙️ Costanti", insieme alle costanti di popolazione.
if "idx_thr" not in st.session_state:
    st.session_state["idx_thr"] = {k: list(v) for k, v in DEFAULT_INDEX_THRESHOLDS.items()}


# ############################################################################
# SCHEDA ⚙️ COSTANTI
# ############################################################################
def render_costanti():
    st.markdown(
        "Valori di riferimento della popolazione (media e deviazione standard, per uomini e "
        "donne) usati per calcolare i T-score. Modificabili direttamente nella tabella; le "
        "modifiche si applicano subito alle altre schede."
    )
    st.caption(
        "Nota: per DSI ed EUR la media di popolazione non viene usata per il T-score ma è utile come riferimento e "
        "viene mostrata accanto al valore dell'atleta. Le soglie di lettura di DSI ed EUR si "
        "modificano qui sotto."
    )

    df_pop = costanti_dataframe(st.session_state["pop"])
    edited_pop = st.data_editor(
        df_pop, use_container_width=True, hide_index=True, key="pop_editor",
        height=min(38 + 35 * len(df_pop), 600),
        column_config={
            "Chiave": None,
            "Categoria": st.column_config.TextColumn(disabled=True),
            "Costante": st.column_config.TextColumn(disabled=True),
            "Unità": st.column_config.TextColumn(disabled=True),
            "Media Uomini": st.column_config.NumberColumn(format="%.4f"),
            "Dev.Std Uomini": st.column_config.NumberColumn(format="%.4f", min_value=0.0),
            "Media Donne": st.column_config.NumberColumn(format="%.4f"),
            "Dev.Std Donne": st.column_config.NumberColumn(format="%.4f", min_value=0.0),
        },
    )
    def _num_or_none(v):
        try:
            v = float(v)
        except (TypeError, ValueError):
            return None
        return None if math.isnan(v) else v

    for row in edited_pop.to_dict("records"):
        key = row["Chiave"]
        vals = dict(mean_m=_num_or_none(row["Media Uomini"]), sd_m=_num_or_none(row["Dev.Std Uomini"]),
                    mean_f=_num_or_none(row["Media Donne"]), sd_f=_num_or_none(row["Dev.Std Donne"]))
        # Le costanti curate devono restare complete (servono ai T-score);
        # le extra possono avere celle vuote (es. donne, n=0).
        if key in CONST_LABELS and None in vals.values():
            continue
        st.session_state["pop"][key] = vals

    # --- Soglie degli indici a rapporto (DSI, EUR) ---
    # Stanno qui, e non accanto ai grafici, perché sono costanti di lettura
    # esattamente come le norme di popolazione sopra. Il blocco viene
    # eseguito PRIMA di build_results(), quindi una modifica si applica
    # nello stesso rerun (nessun click aggiuntivo).
    st.markdown("---")
    st.markdown("**Soglie di lettura degli indici (DSI, EUR)**")
    st.caption(
        "Delimitano le tre zone di profilo nei grafici di DSI ed EUR. Default di letteratura: "
        "DSI 0.60 / 0.80, EUR 1.00 / 1.10. Nota: le soglie EUR di letteratura cadono quasi sulla "
        "media di popolazione qui usata (1.108 U / 1.091 D), quindi un atleta 'medio' finisce sul "
        "confine superiore — da valutare una ricalibrazione sui propri dati."
    )
    tc1, tc2, tc3, tc4 = st.columns(4)
    dsi_low = tc1.number_input("DSI — soglia bassa", key="thr_dsi_low", min_value=0.0, step=0.05,
                               format="%.2f", value=float(st.session_state["idx_thr"]["dsi"][0]))
    dsi_high = tc2.number_input("DSI — soglia alta", key="thr_dsi_high", min_value=0.0, step=0.05,
                                format="%.2f", value=float(st.session_state["idx_thr"]["dsi"][1]))
    eur_low = tc3.number_input("EUR — soglia bassa", key="thr_eur_low", min_value=0.0, step=0.05,
                               format="%.2f", value=float(st.session_state["idx_thr"]["eur"][0]))
    eur_high = tc4.number_input("EUR — soglia alta", key="thr_eur_high", min_value=0.0, step=0.05,
                                format="%.2f", value=float(st.session_state["idx_thr"]["eur"][1]))
    for _idx_key, _lo, _hi in (("dsi", dsi_low, dsi_high), ("eur", eur_low, eur_high)):
        if _hi <= _lo:
            st.warning(f"{_idx_key.upper()}: la soglia alta deve essere maggiore della soglia bassa "
                       "(soglie non aggiornate).")
        else:
            st.session_state["idx_thr"][_idx_key] = [_lo, _hi]

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
            st.session_state["pop"] = _copia_norme(st.session_state["norm_set"])
            st.session_state["idx_thr"] = {k: list(v) for k, v in DEFAULT_INDEX_THRESHOLDS.items()}
            for _k in ("pop_editor", "thr_dsi_low", "thr_dsi_high", "thr_eur_low", "thr_eur_high"):
                if _k in st.session_state:
                    del st.session_state[_k]
            st.rerun()

# ============================================================================
# CARICAMENTO / PARSING DEI FILE
# ============================================================================

parsed_files = []
if csv_reload is not None:
    try:
        parsed_files = [parse_full_export_to_parsed(io.BytesIO(csv_reload.getvalue()))]
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

tab_costanti, tab_dettaglio, tab_profilo, tab_comparazione, tab_report = st.tabs(
    ["⚙️ Costanti", "🔍 Dettaglio Test", "📊 Profilo di Forza", "🔀 Comparazione", "📄 Report"]
)
# La scheda Costanti gira PRIMA di build_results, così le modifiche alle
# norme si applicano già in questo rerun.
with tab_costanti:
    render_costanti()

if parsed_files:
    results_bundle = build_results(parsed_files, st.session_state["pop"], st.session_state["idx_thr"])
    results = results_bundle["results"]
    checks = results_bundle["checks"]
    profilo = profilo_forza(results)
else:
    results, checks, profilo = [], [], {}
with tab_dettaglio:
    render_dettaglio(parsed_files, nome, sesso, periodo, results, checks)

with tab_profilo:
    render_profilo(parsed_files, nome, results, profilo)

with tab_comparazione:
    render_comparazione(parsed_files, nome, sesso, periodo, results)

with tab_report:
    render_report(parsed_files, nome, sesso, periodo, results, profilo)
