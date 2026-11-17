"""Report scaricabili (HTML interattivo + PDF statico), sezione Comparazione
nei report e scheda 📄 Report."""

from __future__ import annotations

import datetime as dt
import html as _html
import importlib
import io
import math
import os
import subprocess
import sys

import streamlit as st

from comparazione import (
    ESITO_COLORI,
    ESITO_INCERTO,
    ESITO_ND,
    ESITO_REALE,
    ESITO_STABILE,
    ORDINE_TEST,
    build_indice_zone_strip,
    build_profili_radar,
    build_storico_chart,
    build_swc_strip,
    colore_delta_t,
    fmt_valore,
    ordina_per_test,
    titolo_test,
)
from core import (
    ACCENT,
    BG_COLOR,
    PRIMARY,
    TEXT_COLOR,
    build_dsi_chart,
    build_eur_chart,
    build_radar_chart,
    build_rsq_chart,
    build_tscore_bar_chart,
    ratio_band_strip,
)
from data import (
    CATEGORIES,
    CATEGORY_QUALITY,
    banda_da_tscore,
)


# --- Bootstrap PDF: BROWSER_PATH per kaleido + fpdf2 sideload in /tmp
# (venv di Streamlit Cloud non scrivibile). Deve precedere `import fpdf`.
os.environ.setdefault("BROWSER_PATH", "/usr/bin/chromium")

_DEPS = "/tmp/fpdf2_clean"
_MARK = os.path.join(_DEPS, ".ok")

if not os.path.exists(_MARK):
    r = subprocess.run([sys.executable, "-m", "pip", "install", "--target", _DEPS,
                        "--no-deps", "--upgrade", "fpdf2==2.8.8"],
                       capture_output=True, text=True)
    if r.returncode == 0:
        open(_MARK, "w").close()
    else:
        print("FPDF2 SIDELOAD FAILED:", r.stderr[-2000:], file=sys.stderr, flush=True)

if os.path.isdir(_DEPS):
    sys.path.insert(0, _DEPS)
    for m in [k for k in list(sys.modules) if k == "fpdf" or k.startswith("fpdf.")]:
        del sys.modules[m]
    importlib.invalidate_caches()

import fpdf  # noqa: E402
import inspect  # noqa: E402
print("FPDF FROM:", fpdf.__file__, "| new_x:",
      "new_x" in inspect.signature(fpdf.FPDF.cell).parameters,
      file=sys.stderr, flush=True)
from fpdf import FPDF  # noqa: E402
from fpdf.fonts import FontFace  # noqa: E402

# ============================================================================
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
#
# ORDINE DEL FILE (non modificarlo a caso):
#   PARTE 6    — helper HTML condivisi
#   PARTE 6ter — sezione Comparazione (HTML + PDF)
#   genera_report_html
#   PARTE 6bis — infrastruttura PDF (_pdf_safe, _ReportPDF, ...)
#   genera_report_pdf
#   with tab_report
# La 6ter usa _pdf_safe/_hex_to_rgb/FontFace, definiti più in basso nella
# 6bis: non è un problema, i nomi vengono risolti alla CHIAMATA e non alla
# definizione della funzione.

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
    grafico invece che in una tabella riepilogativa separata. Include la zona
    di profilo, la media di popolazione e le soglie usate, così il report
    resta leggibile anche a distanza di mesi."""
    if r is None:
        return ""
    extra = []
    if r.get("zona"):
        extra.append(r["zona"])
    if r["pop_mean"] is not None:
        extra.append(f"Media pop. {r['pop_mean']:.3f}")
    if r.get("thr_low") is not None:
        extra.append(f"soglie {r['thr_low']:.2f} / {r['thr_high']:.2f}")
    extra_html = f' <span class="muted">— {" · ".join(extra)}</span>' if extra else ""
    return f'<p class="index-value"><b>{r["label"]}: {r["mean"]:.3f}</b>{extra_html}</p>'


def _metric_descriptions_html(items):
    """Blocco a tendina (nativo, <details>/<summary>, nessun JS) con le
    descrizioni delle metriche — stesso pattern dell'expander "ℹ️ Cosa
    significano queste metriche?" della scheda Dettaglio Test. Ritorna
    stringa vuota se nessuna delle metriche passate ha una descrizione."""
    pairs = [(it["label"], it["desc"]) for it in items if it and it.get("desc")]
    if not pairs:
        return ""
    li_html = "".join(
        f"<li><b>{lbl}</b> — {_html.escape(desc)}</li>" for lbl, desc in pairs
    )
    return f"""<details class="metric-help">
        <summary>ℹ️ Cosa significano queste metriche?</summary>
        <ul>{li_html}</ul>
    </details>"""


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

def genera_report_html(nome, sesso, periodo, results, profilo, commento, thresholds,
                       comp=None):
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
                <div class="profile-card-qualita">{CATEGORY_QUALITY.get(cat, '')}</div>
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
        cat_results = [r for r in results if r["category"] == cat and r["mean"] is not None
                       and r["kind"] != "support"]
        if not cat_results:
            continue
        tscore_fig = build_tscore_bar_chart(cat_results)
        rsq_fig = build_rsq_chart(cat, results)

        sections.append(f"""<section>
            <h2>{cat}</h2>
            {_fig_div(tscore_fig, next_id('tscore'))}
            {_metric_table_html(cat_results)}
            {_metric_descriptions_html(cat_results)}
            {f'<h3>RSQ — Reactive Strength Quadrant</h3>{_fig_div(rsq_fig, next_id("rsq"))}' if rsq_fig else ''}
        </section>""")

    # --- Indici (DSI, EUR): barra a bande + grafico a spicchi, con il valore
    # numerico e le soglie usate accanto a ciascun grafico. Nessun T-score:
    # sono indici diagnostici, non metriche "più alto = meglio".
    r_dsi = next((r for r in results if r["key"] == "dsi" and r["mean"] is not None), None)
    r_eur = next((r for r in results if r["key"] == "eur" and r["mean"] is not None), None)
    dsi_fig = build_dsi_chart(results, thresholds)
    eur_fig = build_eur_chart(results, thresholds)
    if r_dsi or r_eur or dsi_fig or eur_fig:
        indici_html = []
        for r_idx, fig_idx, titolo, idx_key in (
            (r_dsi, dsi_fig, "DSI (Dynamic Strength Index)", "dsi"),
            (r_eur, eur_fig, "EUR (Eccentric Utilisation Ratio)", "eur"),
        ):
            if not (r_idx or fig_idx):
                continue
            strip = ratio_band_strip(idx_key, r_idx["mean"], *thresholds[idx_key],
                                     label=idx_key.upper(), decimals=3) if r_idx else None
            indici_html.append(
                f"<h3>{titolo}</h3>{_index_value_html(r_idx)}"
                f"{_fig_div(strip, next_id(idx_key + '_strip'))}"
                f"{_fig_div(fig_idx, next_id(idx_key))}"
            )
        sections.append(f"""<section>
            <h2>Indici</h2>
            {''.join(indici_html)}
            {_metric_descriptions_html([r_dsi, r_eur])}
        </section>""")

    # --- Comparazione con lo storico ---
    # Sezione OPZIONALE: comparazione_sections_html() restituisce una lista
    # vuota se comp è None, cioè se la scheda 🔀 Comparazione non è stata
    # usata in questa sessione. Va prima dell'Analisi, che chiude il report.
    sections.extend(comparazione_sections_html(comp, nome, next_id))

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
    .profile-card-qualita {{ font-size: 17px; font-weight: 600; color: var(--text); margin: 2px 0 4px 0; }}
    .profile-card-t {{ font-size: 30px; font-weight: 700; }}
    .profile-card-banda {{ font-size: 13px; }}
    .profile-card-delta {{ font-size: 15px; font-weight: 600; margin-top: 4px; }}
    .muted {{ color: #667; font-size: 13px; }}
    .index-value {{ font-size: 16px; margin: 6px 0 14px 0; }}
    .metric-help {{ margin: 14px 0 4px 0; font-size: 14px; }}
    .metric-help summary {{ cursor: pointer; color: var(--primary); font-weight: 600; }}
    .metric-help ul {{ margin: 10px 0 0 0; padding-left: 20px; }}
    .metric-help li {{ margin-bottom: 6px; line-height: 1.4; }}
    .analysis-box {{ min-height: 60px; border: 1px solid #e0e0e0; border-left: 4px solid var(--primary); border-radius: 6px; padding: 14px; line-height: 1.6; background: #f7fbfd; }}
    .print-btn {{
        background: var(--accent); color: #fff; border: none; border-radius: 6px;
        padding: 10px 20px; font-size: 15px; font-weight: 600; cursor: pointer;
        margin-top: 18px;
    }}
    .print-btn:hover {{ opacity: 0.9; }}
    @media print {{
        body {{ background: #fff; }}
        .no-print {{ display: none !important; }}
        section {{ break-inside: avoid; page-break-inside: avoid; }}
        .js-plotly-plot .modebar {{ display: none !important; }}
        details.metric-help {{ break-inside: avoid; }}
    }}
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
            <button class="print-btn no-print" onclick="window.print()">🖨️ Stampa / Salva come PDF</button>
        </header>
        <p class="intro-text">
            Per valutare i risultati del test è stato utilizzato il T-Score, un indice standardizzato
            che confronta la prestazione dell'atleta rispetto a un gruppo di riferimento, esprimendo la
            distanza dalla media in deviazioni standard. Punteggi tra 0 e 50 indicano valori inferiori
            alla media, mentre punteggi tra 50 e 100 indicano valori superiori alla media.
            DSI ed EUR fanno eccezione: essendo rapporti tra due test, non vengono letti come
            "più alto = meglio" ma per zona di profilo, confrontando la posizione dell'atleta con
            le soglie di riferimento indicate accanto a ciascun grafico.
        </p>
        {''.join(sections)}
    </div>
</body>
</html>"""
    return html.encode("utf-8")

# ============================================================================
# PARTE 6ter — SEZIONE COMPARAZIONE NEI REPORT
# ============================================================================
# I risultati della scheda 🔀 Comparazione vengono salvati in
# st.session_state["comp_export"] dalla PARTE 5bis. Qui vengono trasformati
# in HTML e in PDF. Se la chiave non esiste (comparazione non eseguita in
# questa sessione) le funzioni restituiscono contenuto vuoto e il report
# resta identico a prima: la sezione è opzionale, non obbligatoria.
#
# Le FIGURE non vengono salvate nello stato ma ricostruite qui dalle stesse
# funzioni build_swc_strip / build_profili_radar / build_indice_zone_strip
# usate nella UI live: un solo posto dove cambiare la grafica.

# Emoji -> testo per il PDF: i font core di fpdf2 sono Latin-1, quindi ogni
# emoji verrebbe scartata da _pdf_safe lasciando l'etichetta mutila
# ("🟢 Alta" -> " Alta"). Le frecce diventano segni ASCII, i pallini del
# semaforo sparisc­ono perché nel PDF il colore lo dà il riempimento cella.
_PDF_COMP_REPLACEMENTS = {
    "\U0001F7E2": "", "\U0001F7E1": "", "\U0001F534": "",
    "\u26A0\uFE0F": "", "\u26A0": "", "\u2753": "",
    "\U0001F53A": "+", "\U0001F53B": "-", "\u25AA": "=",
}


def _pdf_comp(text):
    """_pdf_safe + traduzione delle emoji di Comparazione."""
    t = str(text if text is not None else "")
    for old, new in _PDF_COMP_REPLACEMENTS.items():
        t = t.replace(old, new)
    return _pdf_safe(t.strip())


def _delta_t_html(delta):
    """Riga compatta col solo scostamento in punti T. Entita' e attendibilita'
    restano in tabella: nella card toglierebbero risalto al T-score."""
    if delta is None or (isinstance(delta, float) and math.isnan(delta)):
        return ""
    return (f'<div class="profile-card-delta" style="color:{colore_delta_t(delta)}">'
            f'{delta:+.1f} punti T</div>')


def _comp_esito_colore(esito):
    """Colore associato a un'etichetta di attendibilità, con fallback grigio."""
    return ESITO_COLORI.get(str(esito), "#8d8d8d")


def _fmt(valore, decimali=3, segno=False, suffisso=""):
    """Wrapper su fmt_valore della PARTE 5bis: stessa resa numerica in UI e
    nei report. `decimali` e' rispettato solo per le percentuali, dove serve
    un numero fisso di cifre."""
    if valore is None or (isinstance(valore, float) and math.isnan(valore)):
        return "—"
    if suffisso == "%" and isinstance(valore, (int, float)):
        return (f"{valore:+.{decimali}f}" if segno else f"{valore:.{decimali}f}") + suffisso
    return fmt_valore(valore, segno=segno) + suffisso


# ----------------------------------------------------------------------------
# HTML
# ----------------------------------------------------------------------------
COMP_HTML_COLS = [
    ("Test", 0, False), ("Metrica", 0, False), ("Unità", 0, False),
    ("Media rif.", 3, False), ("Attuale", 3, False),
    ("Delta", 3, True), ("Delta %", 1, True),
    ("Cambiamento", 0, False), ("Attendibilità", 0, False),
    ("Migliore", 3, False), ("% del migliore", 0, False),
    ("Seduta migliore", 0, False), ("N sedute", 0, False),
]


def _comp_table_html(vista):
    """Tabella della Comparazione: stesse colonne e stesso ordine della UI."""
    if vista is None or vista.empty:
        return ""
    thead = "".join(f"<th>{c}</th>" for c, _d, _s in COMP_HTML_COLS)
    righe = []
    for _, r in vista.iterrows():
        celle = []
        for col, dec, segno in COMP_HTML_COLS:
            v = r.get(col)
            if col == "Attendibilità":
                celle.append(f'<td style="color:{_comp_esito_colore(v)};'
                             f'font-weight:600">{v}</td>')
            elif col == "Delta %":
                celle.append(f"<td>{_fmt(v, 1, True, '%')}</td>")
            elif col == "% del migliore":
                celle.append(f"<td>{_fmt(v, 0, False, '%')}</td>")
            elif dec:
                celle.append(f"<td>{_fmt(v, dec, segno)}</td>")
            else:
                celle.append(f"<td>{v if v not in (None, '') else '—'}</td>")
        righe.append("<tr>" + "".join(celle) + "</tr>")
    return (f'<table class="report-table"><thead><tr>{thead}</tr></thead>'
            f'<tbody>{"".join(righe)}</tbody></table>')


PROF_HTML_COLS = [
    ("Categoria", 0), ("T medio rif.", 1), ("T attuale", 1), ("Delta T", 1),
    ("Cambiamento", 0), ("Attendibilità", 0), ("Valutazione attuale", 0),
    ("T migliore", 1), ("Seduta migliore", 0), ("N sedute", 0),
]


def _prof_table_html(prof):
    if prof is None or prof.empty:
        return ""
    thead = "".join(f"<th>{c}</th>" for c, _d in PROF_HTML_COLS)
    righe = []
    for _, r in prof.iterrows():
        celle = []
        for col, dec in PROF_HTML_COLS:
            v = r.get(col)
            if col == "Attendibilità":
                celle.append(f'<td style="color:{_comp_esito_colore(v)};'
                             f'font-weight:600">{v}</td>')
            elif col == "Delta T":
                celle.append(f"<td>{v:+.1f}</td>" if isinstance(v, (int, float)) else "<td>—</td>")
            elif dec:
                celle.append(f"<td>{v:.{dec}f}</td>" if isinstance(v, (int, float)) else "<td>—</td>")
            else:
                celle.append(f"<td>{v if v not in (None, '') else '—'}</td>")
        righe.append("<tr>" + "".join(celle) + "</tr>")
    return (f'<table class="report-table"><thead><tr>{thead}</tr></thead>'
            f'<tbody>{"".join(righe)}</tbody></table>')


def comparazione_sections_html(comp, nome_atleta, next_id):
    """Sezioni HTML della Comparazione. Lista vuota se comp è None."""
    if not comp:
        return []
    vista = comp.get("vista")
    if vista is None or vista.empty:
        return []

    conteggi = vista["Attendibilità"].value_counts()
    cards = "".join(
        f"""<div class="profile-card">
                <div class="profile-card-cat">{etichetta}</div>
                <div class="profile-card-t" style="color:{_comp_esito_colore(etichetta)}">
                    {int(conteggi.get(etichetta, 0))}</div>
            </div>"""
        for etichetta in (ESITO_REALE, ESITO_INCERTO, ESITO_STABILE, ESITO_ND)
    )

    sezioni = [f"""<section>
        <h2>Comparazione con lo storico</h2>
        <p class="intro-text">Il test è confrontato con la <b>media di
        {comp.get('n_sedute', 0)} sedute precedenti</b> (esclusa l'attuale).
        <b>Cambiamento</b> dice direzione ed entità dello scostamento;
        <b>Attendibilità</b> dice quanto è credibile che un cambiamento ci sia
        stato davvero, non se sia un bene: quel giudizio resta al preparatore.</p>
        <p class="muted">Soglia di rilevanza (SWC) = 0,2 × dev.std di popolazione.
        Incertezza di ogni seduta = dispersione fra le sue ripetizioni
        (dev.std ÷ radice di n), intervallo di confidenza al 90%. La variabilità
        fra giorni diversi non entra nell'incertezza: per un atleta che si
        allena è cambiamento reale, non rumore.</p>
        <div class="profile-cards">{cards}</div>
        {_comp_table_html(vista)}
    </section>"""]

    # Strisce per le metriche scelte nella UI
    scelte = comp.get("strip") or []
    strisce = []
    righe = [r for _, r in vista[vista["_display"].isin(scelte)].iterrows()]
    test_corrente = None
    for riga in ordina_per_test(righe, lambda r: r["Test"], ORDINE_TEST):
        fig = build_swc_strip(riga)
        if fig is not None:
            if riga["Test"] != test_corrente:
                test_corrente = riga["Test"]
                strisce.append(f"<h3>{_html.escape(titolo_test(test_corrente))}</h3>")
            strisce.append(_fig_div(fig, next_id("comp_strip")))
    if strisce:
        sezioni.append(f"""<section>
            <h2>Scostamento dalla media, metrica per metrica</h2>
            <p class="intro-text">Il rombo pieno azzurro è il test attuale con la
            sua barra di incertezza; il rombo vuoto arancione è la media delle
            sedute precedenti. Le fasce colorate sono l'entità dello scostamento
            (trascurabile, piccolo, moderato, grande), simmetriche sopra e sotto
            la media.</p>
            {''.join(strisce)}
        </section>""")

    # Andamento storico: una figura per metrica seguita nel tempo
    storico = comp.get("storico") or []
    blocchi_st = []
    display_to_test = {r["_display"]: r["Test"] for _, r in vista.iterrows()}
    test_corrente = None
    for voce in ordina_per_test(storico, lambda v: display_to_test.get(v["display"]), ORDINE_TEST):
        fig = build_storico_chart(voce["df"], voce["meta"])
        if fig is None:
            continue
        if display_to_test.get(voce["display"]) != test_corrente:
            test_corrente = display_to_test.get(voce["display"])
            blocchi_st.append(f"<h3>{_html.escape(titolo_test(test_corrente))}</h3>")
        meta_v = voce["meta"]
        nota = (f"Riferimento: {', '.join(map(str, meta_v['sessioni_rif']))}"
                + (" · barra propria per ogni seduta" if meta_v.get("per_punto")
                   else " · incertezza da valori di letteratura")
                + f" · incertezza media ±{fmt_valore(meta_v['margine'])}"
                + (f" · soglia ±{fmt_valore(meta_v['swc'])}" if meta_v.get("swc") else ""))
        blocchi_st.append(f"{_fig_div(fig, next_id('comp_storico'))}"
                          f'<p class="muted">{nota}</p>')
    if blocchi_st:
        sezioni.append(f"""<section>
            <h2>Andamento storico</h2>
            <p class="intro-text">Ogni punto è una seduta, con la sua barra di
            incertezza. La banda grigia è la zona di rilevanza (±SWC) attorno al
            riferimento; la linea tratteggiata arancione è il riferimento stesso.
            La barra sotto l'asse è il tempo trascorso fra una seduta e l'altra:
            più larga e più scura, più tempo è passato.</p>
            {''.join(blocchi_st)}
        </section>""")

    # Profilo di forza: radar sovrapposto + tabella per categoria
    prof, serie, cats = comp.get("prof"), comp.get("prof_serie"), comp.get("prof_cats")
    if serie and cats:
        radar = build_profili_radar(cats, serie, nome_atleta)
        attuale = serie.get(f"{nome_atleta} — attuale", {})
        delta_map = ({r["Categoria"]: r["Delta T"] for _, r in prof.iterrows()}
                     if prof is not None and not prof.empty else {})
        cards = "".join(
            f"""<div class="profile-card">
                    <div class="profile-card-cat">{c}</div>
                    <div class="profile-card-qualita">{CATEGORY_QUALITY.get(c, '')}</div>
                    <div class="profile-card-t" style="color:{banda_da_tscore(attuale[c])[1]}">{attuale[c]:.0f}</div>
                    <div class="profile-card-banda">{banda_da_tscore(attuale[c])[0]}</div>
                    {_delta_t_html(delta_map.get(c))}
                </div>"""
            for c in cats if c in attuale
        )
        sezioni.append(f"""<section>
            <h2>Profilo di forza: attuale vs storico</h2>
            {_fig_div(radar, next_id('comp_radar')) if radar else ''}
            <div class="profile-cards">{cards}</div>
            <p class="intro-text">Il T-score è già orientato alla prestazione,
            quindi qui un valore più alto è sempre migliore. La soglia di
            rilevanza vale <b>2 punti di T-score</b>: da T = 50 + 10·z e
            SWC = 0,2 deviazioni standard segue che 0,2 z = 2 punti T.</p>
            {_prof_table_html(prof)}
        </section>""")

    # Indici di profilo
    indici = comp.get("indici") or []
    if indici:
        blocchi = []
        for voce in indici:
            fig = build_indice_zone_strip(
                voce["key"], voce["attuale"], voce["riferimento"], *voce["thr"])
            if fig is None:
                continue
            if voce["zona_att"] and voce["zona_rif"] and voce["zona_att"] != voce["zona_rif"]:
                nota = (f"cambio di zona: da <i>{voce['zona_rif']}</i> "
                        f"a <i>{voce['zona_att']}</i>")
            elif voce["zona_att"]:
                nota = f"zona invariata: <i>{voce['zona_att']}</i>"
            else:
                nota = ""
            blocchi.append(
                f"<h3>{voce['key'].upper()}</h3>"
                f'<p class="index-value"><b>{_fmt(voce["attuale"], 3)}</b>'
                f' <span class="muted">— {nota}</span></p>'
                f"{_fig_div(fig, next_id('comp_idx'))}"
            )
        if blocchi:
            sezioni.append(f"""<section>
                <h2>Indici di profilo: si è spostato di zona?</h2>
                <p class="intro-text">DSI ed EUR non hanno un verso migliore:
                quello che conta è se l'atleta ha cambiato zona di profilo.</p>
                {''.join(blocchi)}
            </section>""")

    return sezioni


# ----------------------------------------------------------------------------
# PDF
# ----------------------------------------------------------------------------
# Larghezze in mm: somma 180 = larghezza utile con margini 15/15. "Test" e le
# colonne di coda della tabella HTML sono omesse per stare in A4 verticale.
COMP_PDF_COLS = [
    ("Metrica", 40, lambda r: _pdf_comp(r.get("Metrica"))),
    ("UdM", 11, lambda r: _pdf_comp(r.get("Unità"))),
    ("Media rif.", 17, lambda r: _fmt(r.get("Media rif."), 3)),
    ("Attuale", 17, lambda r: _fmt(r.get("Attuale"), 3)),
    ("Delta", 17, lambda r: _fmt(r.get("Delta"), 3, True)),
    ("Cambiam.", 22, lambda r: _pdf_comp(r.get("Cambiamento"))),
    ("Attendib.", 25, lambda r: _pdf_comp(r.get("Attendibilità"))),
    ("Migliore", 17, lambda r: _fmt(r.get("Migliore"), 3)),
    ("% migl.", 14, lambda r: _fmt(r.get("% del migliore"), 0, False, "%")),
]

PROF_PDF_COLS = [
    ("Categoria", 48, lambda r: _pdf_comp(r.get("Categoria"))),
    ("T rif.", 15, lambda r: f"{r.get('T medio rif.'):.1f}"),
    ("T att.", 15, lambda r: f"{r.get('T attuale'):.1f}"),
    ("Delta T", 17, lambda r: f"{r.get('Delta T'):+.1f}"),
    ("Cambiam.", 22, lambda r: _pdf_comp(r.get("Cambiamento"))),
    ("Attendib.", 25, lambda r: _pdf_comp(r.get("Attendibilità"))),
    ("Valutazione", 38, lambda r: _pdf_comp(r.get("Valutazione attuale"))),
]


def _pdf_comp_table(pdf, df, colonne, col_esito="Attendibilità"):
    """Tabella PDF generica per la Comparazione. La cella di attendibilità
    viene riempita col colore del semaforo, perché nel PDF le emoji non
    esistono e il colore è l'unico segnale visivo che resta."""
    if df is None or df.empty:
        return
    pdf.set_x(pdf.l_margin)
    pdf.set_font("Helvetica", "", 8)
    pdf.set_draw_color(200, 200, 200)
    pdf.set_line_width(0.2)
    pdf.set_fill_color(255, 255, 255)
    pdf.set_text_color(30, 30, 30)
    pdf.ensure_space(12)
    with pdf.table(
        col_widths=[w for _h, w, _f in colonne], text_align="LEFT", line_height=4.6,
        headings_style=FontFace(emphasis="BOLD", color=(255, 255, 255),
                                fill_color=_hex_to_rgb(TEXT_COLOR)),
    ) as table:
        riga = table.row()
        for intestazione, _w, _f in colonne:
            riga.cell(intestazione)
        for _, r in df.iterrows():
            riga = table.row()
            for intestazione, _w, estrai in colonne:
                try:
                    testo = estrai(r)
                except (TypeError, ValueError):
                    testo = "-"
                stile = None
                if intestazione.startswith("Attendib"):
                    stile = FontFace(color=(255, 255, 255),
                                     fill_color=_hex_to_rgb(_comp_esito_colore(r.get(col_esito))))
                riga.cell(testo, style=stile)
    pdf.set_x(pdf.l_margin)


def comparazione_sezione_pdf(pdf, comp, nome_atleta):
    """Aggiunge al PDF la sezione Comparazione. No-op se comp è None."""
    if not comp:
        return
    vista = comp.get("vista")
    if vista is None or vista.empty:
        return

    pdf.add_page()
    pdf.section_title("Comparazione con lo storico")
    pdf.body_text(
        f"Il test e confrontato con la media di {comp.get('n_sedute', 0)} sedute "
        "precedenti (esclusa l'attuale). La colonna Cambiamento indica direzione "
        "(+ in aumento, - in diminuzione) ed entita dello scostamento; "
        "Attendibilita indica quanto e credibile che un cambiamento ci sia stato "
        "davvero, non se sia un bene: quel giudizio resta al preparatore."
    )
    pdf.body_text(
        "Soglia di rilevanza (SWC) = 0,2 x dev.std di popolazione. Incertezza di "
        "ogni seduta = dispersione fra le sue ripetizioni (dev.std diviso radice "
        "di n), intervallo di confidenza al 90%. La variabilita fra giorni diversi "
        "non entra nell'incertezza: per un atleta che si allena e cambiamento "
        "reale, non rumore.", size=8)
    conteggi = vista["Attendibilità"].value_counts()
    pdf.body_text("   ".join(
        f"{_pdf_comp(e)}: {int(conteggi.get(e, 0))}"
        for e in (ESITO_REALE, ESITO_INCERTO, ESITO_STABILE, ESITO_ND)
    ), size=9)
    pdf.ln(1)
    _pdf_comp_table(pdf, vista, COMP_PDF_COLS)

    # Strisce: height_px DEVE coincidere con l'altezza della figura, altrimenti
    # kaleido rende la fascia con uno spessore diverso da quello previsto.
    scelte = comp.get("strip") or []
    strisce = vista[vista["_display"].isin(scelte)]
    if not strisce.empty:
        pdf.ln(2)
        pdf.subsection_title("Scostamento dalla media, metrica per metrica")
        pdf.body_text(
            "Rombo pieno = test attuale con barra di incertezza; rombo vuoto = "
            "media delle sedute precedenti. Le fasce colorate sono l'entita "
            "dello scostamento.", size=9)
        righe = [r for _, r in strisce.iterrows()]
        test_corrente = None
        for riga in ordina_per_test(righe, lambda r: r["Test"], ORDINE_TEST):
            fig = build_swc_strip(riga)
            if fig is not None:
                if riga["Test"] != test_corrente:
                    test_corrente = riga["Test"]
                    pdf.test_title(titolo_test(test_corrente), keep_with_mm=35)
                pdf.chart_image(fig, width_px=900, height_px=fig.layout.height,
                                content_width_mm=165)

    # Andamento storico: una figura per metrica. height_px deve coincidere
    # con l'altezza della figura, altrimenti kaleido sposta la barra dei
    # periodi (che vive in coordinate paper sotto l'area di plot).
    storico = comp.get("storico") or []
    if storico:
        pdf.add_page()
        pdf.section_title("Andamento storico")
        pdf.body_text(
            "Ogni punto e' una seduta con la sua barra di incertezza. La banda "
            "grigia e' la zona di rilevanza (+/- SWC) attorno al riferimento, la "
            "linea tratteggiata e' il riferimento stesso. La barra sotto l'asse "
            "e' il tempo trascorso fra una seduta e l'altra: piu' larga e piu' "
            "scura, piu' tempo e' passato.", size=9)
        display_to_test = {r["_display"]: r["Test"] for _, r in vista.iterrows()}
        test_corrente = None
        for voce in ordina_per_test(storico, lambda v: display_to_test.get(v["display"]),
                                    ORDINE_TEST):
            fig = build_storico_chart(voce["df"], voce["meta"])
            if fig is None:
                continue
            if display_to_test.get(voce["display"]) != test_corrente:
                test_corrente = display_to_test.get(voce["display"])
                pdf.test_title(titolo_test(test_corrente), keep_with_mm=95)
            pdf.chart_image(fig, width_px=1000, height_px=fig.layout.height,
                            content_width_mm=175)
            meta_v = voce["meta"]
            pdf.body_text(
                f"Riferimento: {', '.join(map(str, meta_v['sessioni_rif']))}"
                + ("   barra propria per ogni seduta" if meta_v.get("per_punto")
                   else "   incertezza da valori di letteratura")
                + f"   incertezza media +/-{fmt_valore(meta_v['margine'])}"
                + (f"   soglia +/-{fmt_valore(meta_v['swc'])}"
                   if meta_v.get("swc") else ""), size=8)
            pdf.ln(1)

    prof, serie, cats = comp.get("prof"), comp.get("prof_serie"), comp.get("prof_cats")
    if serie and cats:
        pdf.add_page()
        pdf.section_title("Profilo di forza: attuale vs storico")
        radar = build_profili_radar(cats, serie, nome_atleta)
        if radar:
            pdf.chart_image(radar, width_px=1000, height_px=620, content_width_mm=150)
        attuale = serie.get(f"{nome_atleta} — attuale", {})
        if attuale:
            delta_map = ({r["Categoria"]: r["Delta T"] for _, r in prof.iterrows()}
                         if prof is not None and not prof.empty else {})
            pdf.profile_cards([c for c in cats if c in attuale], attuale, delta_map)
        pdf.body_text(
            "Il T-score e gia orientato alla prestazione, quindi qui un valore "
            "piu alto e sempre migliore. 1 soglia di rilevanza = 2 punti di T-score.",
            size=9)
        _pdf_comp_table(pdf, prof, PROF_PDF_COLS)

    indici = comp.get("indici") or []
    if indici:
        pdf.ln(3)
        pdf.subsection_title("Indici di profilo: si e spostato di zona?")
        for voce in indici:
            fig = build_indice_zone_strip(
                voce["key"], voce["attuale"], voce["riferimento"], *voce["thr"])
            if fig is None:
                continue
            if voce["zona_att"] and voce["zona_rif"] and voce["zona_att"] != voce["zona_rif"]:
                nota = f"cambio di zona: da {voce['zona_rif']} a {voce['zona_att']}"
            elif voce["zona_att"]:
                nota = f"zona invariata: {voce['zona_att']}"
            else:
                nota = ""
            pdf.body_text(f"{voce['key'].upper()}: {_fmt(voce['attuale'], 3)}   {nota}",
                          size=9)
            pdf.chart_image(fig, width_px=900, height_px=fig.layout.height,
                            content_width_mm=150)

# ============================================================================
# PARTE 6bis — REPORT SCARICABILE (PDF statico)
# ============================================================================
# Alternativa "un tap e via" all'HTML per chi ha problemi ad aprire report
# interattivi su iOS (Quick Look non esegue JavaScript). I grafici Plotly
# vengono renderizzati come immagini statiche via kaleido (offline, nessuna
# richiesta di rete) e assemblati in un PDF vero con fpdf2 (puro Python,
# nessuna libreria di sistema richiesta — a differenza di weasyprint).

_PDF_CHAR_REPLACEMENTS = {
    "\u2014": " - ",  # em dash
    "\u2013": "-",    # en dash
    "\u2192": "->",   # freccia destra
    "\u2018": "'", "\u2019": "'",
    "\u201c": '"', "\u201d": '"',
    "\u2022": "-",    # bullet
}


def _pdf_safe(text):
    """I font core di fpdf2 (Helvetica) coprono solo Latin-1: sostituisce i
    caratteri tipografici comuni (em dash, frecce, virgolette curve) con
    equivalenti ASCII e scarta silenziosamente qualsiasi altro carattere
    fuori range (es. emoji), invece di far fallire l'export."""
    if text is None:
        return ""
    text = str(text)
    for old, new in _PDF_CHAR_REPLACEMENTS.items():
        text = text.replace(old, new)
    return "".join(ch if ord(ch) <= 255 else "" for ch in text)


def _hex_to_rgb(hex_color):
    if not hex_color:
        return (230, 230, 230)
    h = hex_color.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def _fig_png_bytes(fig, width=1000, height=560, scale=2):
    """Renderizza una figura Plotly in PNG (bytes) via kaleido, offline.
    Ritorna (BytesIO, aspect_ratio) o None se fig è None."""
    if fig is None:
        return None
    png_bytes = fig.to_image(format="png", width=width, height=height, scale=scale)
    return io.BytesIO(png_bytes), height / width


class _ReportPDF(FPDF):
    def __init__(self):
        super().__init__(format="A4")
        self.set_auto_page_break(auto=True, margin=16)
        self.set_margins(15, 15, 15)

    def ensure_space(self, height_mm):
        """Forza un salto pagina se l'elemento successivo (altezza nota in
        mm) non entra nello spazio rimasto nella pagina corrente."""
        if self.get_y() + height_mm > self.page_break_trigger:
            self.add_page()

    def section_title(self, text):
        self.ensure_space(14)
        self.set_x(self.l_margin)
        self.set_font("Helvetica", "B", 14)
        self.set_text_color(*_hex_to_rgb(TEXT_COLOR))
        self.cell(0, 10, _pdf_safe(text), new_x="LMARGIN", new_y="NEXT")
        self.set_draw_color(*_hex_to_rgb(ACCENT))
        self.set_line_width(0.8)
        self.line(self.l_margin, self.get_y(), self.l_margin + 6, self.get_y())
        self.ln(3)

    def subsection_title(self, text):
        self.ensure_space(10)
        self.set_x(self.l_margin)
        self.set_font("Helvetica", "B", 11)
        self.set_text_color(*_hex_to_rgb(TEXT_COLOR))
        self.cell(0, 8, _pdf_safe(text), new_x="LMARGIN", new_y="NEXT")

    def test_title(self, text, keep_with_mm=30):
        """Piccolo titolo di gruppo (nome del test). keep_with_mm tiene il
        titolo sulla stessa pagina del primo grafico che lo segue."""
        self.ensure_space(8 + keep_with_mm)
        self.ln(1)
        self.set_x(self.l_margin)
        self.set_font("Helvetica", "B", 10)
        self.set_text_color(*_hex_to_rgb(ACCENT))
        self.cell(0, 6, _pdf_safe(text), new_x="LMARGIN", new_y="NEXT")
        self.set_text_color(*_hex_to_rgb(TEXT_COLOR))

    def body_text(self, text, size=10):
        self.set_x(self.l_margin)
        self.set_font("Helvetica", "", size)
        self.set_text_color(60, 60, 70)
        self.multi_cell(0, 5.5, _pdf_safe(text))

    def chart_image(self, fig, width_px=1000, height_px=560, content_width_mm=180):
        result = _fig_png_bytes(fig, width=width_px, height=height_px)
        if result is None:
            return
        img_io, aspect = result
        h_mm = content_width_mm * aspect
        self.ensure_space(h_mm + 4)
        x = self.l_margin + (self.epw - content_width_mm) / 2
        self.image(img_io, x=x, w=content_width_mm)
        self.set_x(self.l_margin)
        self.ln(4)

    def metric_table(self, cat_results):
        headers = ["Metrica", "Unita", "Media", "Media Pop.", "N", "T-score", "Valutazione"]
        widths = [52, 14, 18, 18, 10, 16, 35]  # mm assoluti, sommano ~163mm su 180mm utili
        self.set_x(self.l_margin)
        self.set_font("Helvetica", "", 8.5)
        self.set_draw_color(200, 200, 200)  # i bordi/il fill della tabella non devono
        self.set_line_width(0.2)            # ereditare accent/testo/fill lasciati da altri
        self.set_fill_color(255, 255, 255)  # elementi (es. il rettangolo blu dell'header,
        self.set_text_color(30, 30, 30)     # o il testo blu del blocco Profilo di Forza)
        self.ensure_space(10)
        with self.table(
            col_widths=widths, text_align="LEFT", line_height=5,
            headings_style=FontFace(emphasis="BOLD", color=(255, 255, 255), fill_color=_hex_to_rgb(TEXT_COLOR)),
        ) as table:
            row = table.row()
            for h in headers:
                row.cell(h)
            for r in cat_results:
                row = table.row()
                row.cell(_pdf_safe(r["label"]))
                row.cell(_pdf_safe(r["unit"] or ""))
                row.cell(f"{r['mean']:.3f}" if r["mean"] is not None else "N/D")
                row.cell(f"{r['pop_mean']:.3f}" if r["pop_mean"] is not None else "-")
                row.cell(str(r["n"]))
                row.cell(f"{r['t']:.0f}" if r["t"] is not None else "-")
                style = None
                if r.get("colore"):
                    style = FontFace(fill_color=_hex_to_rgb(r["colore"]))
                row.cell(_pdf_safe(r["banda"] or "-"), style=style)
        # fpdf2 può lasciare il cursore x dopo l'ultima cella dell'ultima riga
        # invece di riportarlo al margine sinistro: senza questo reset, un
        # multi_cell(0, ...) subito dopo la tabella può ricevere una
        # larghezza quasi nulla e sollevare FPDFException.
        self.set_x(self.l_margin)

    def profile_cards(self, cats, profilo, delta_map=None):
        """Riquadri affiancati per il Profilo di Forza, equivalenti alle
        .profile-card del report HTML: nome del protocollo, qualità fisica
        misurata, T-score grande nel colore della banda, lo scostamento in
        punti T (solo nella sezione Comparazione) e infine l'etichetta della
        banda di valutazione.

        delta_map è opzionale: senza, la card è quella del report base e
        l'altezza scende da 36 a 31 mm, altrimenti resterebbe un buco bianco
        sotto la banda. Le coordinate sono assolute, quindi se cambi `h`
        devi spostare anche tutti gli offset y0 + ... qui sotto."""
        if not cats:
            return
        n = len(cats)
        con_delta = bool(delta_map)
        gap = 4
        h = 36 if con_delta else 31
        w = (self.epw - gap * (n - 1)) / n
        self.ensure_space(h + 6)
        y0 = self.get_y()
        for i, cat in enumerate(cats):
            t = profilo[cat]
            banda, colore = banda_da_tscore(t)
            x = self.l_margin + i * (w + gap)

            self.set_fill_color(247, 251, 253)
            self.set_draw_color(224, 224, 224)
            self.set_line_width(0.2)
            self.rect(x, y0, w, h, style="DF")

            # Nome del protocollo: piccolo e grigio, e' il "come"
            self.set_xy(x + 2, y0 + 2.5)
            self.set_font("Helvetica", "B", 6.5)
            self.set_text_color(120, 145, 160)
            self.multi_cell(w - 4, 3.2, _pdf_safe(cat), align="C")

            # Qualita' fisica: e' il "cosa", quindi in evidenza
            self.set_xy(x + 2, y0 + 9.4)
            self.set_font("Helvetica", "B", 9)
            self.set_text_color(*_hex_to_rgb(TEXT_COLOR))
            self.multi_cell(w - 4, 4, _pdf_safe(CATEGORY_QUALITY.get(cat, "")), align="C")

            # T-score nel colore della banda
            self.set_xy(x, y0 + 14.5)
            self.set_font("Helvetica", "B", 20)
            self.set_text_color(*_hex_to_rgb(colore))
            self.cell(w, 9, f"{t:.0f}", align="C")

            # Scostamento in punti T: sta SOPRA la banda, che quindi scende di
            # 5 mm. Senza delta la banda risale al suo posto originale.
            d = (delta_map or {}).get(cat)
            y_banda = y0 + 24
            if d is not None:
                self.set_xy(x + 2, y0 + 23.5)
                self.set_font("Helvetica", "B", 8)
                self.set_text_color(*_hex_to_rgb(colore_delta_t(d)))
                self.multi_cell(w - 4, 3.6, _pdf_safe(f"{d:+.1f}"), align="C")
                y_banda = y0 + 28.5

            self.set_xy(x + 2, y_banda)
            self.set_font("Helvetica", "", 6.5)
            self.set_text_color(90, 90, 100)
            self.multi_cell(w - 4, 3.2, _pdf_safe(banda or "-"), align="C")

        # Reset esplicito: fill/testo/bordo restano impostati sull'ultimo
        # riquadro disegnato e verrebbero ereditati dagli elementi seguenti.
        self.set_xy(self.l_margin, y0 + h)
        self.set_fill_color(255, 255, 255)
        self.set_text_color(30, 30, 30)
        self.set_draw_color(200, 200, 200)
        self.ln(5)

    def descriptions_block(self, items):
        pairs = [(it["label"], it["desc"]) for it in items if it and it.get("desc")]
        if not pairs:
            return
        self.ensure_space(8)
        self.set_x(self.l_margin)
        self.set_font("Helvetica", "I", 8.5)
        self.set_text_color(90, 90, 100)
        for label, desc in pairs:
            self.ensure_space(6)
            self.set_x(self.l_margin)
            self.multi_cell(0, 4.5, _pdf_safe(f"- {label}: {desc}"))
        self.ln(1)


def genera_report_pdf(nome, sesso, periodo, results, profilo, commento, thresholds,
                      comp=None):
    pdf = _ReportPDF()
    pdf.add_page()

    # --- Header ---
    pdf.set_fill_color(*_hex_to_rgb(TEXT_COLOR))
    pdf.rect(0, 0, 210, 26, style="F")
    pdf.set_xy(15, 7)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 17)
    pdf.cell(0, 8, "FORCE PLATE TEST REPORT", new_x="LMARGIN", new_y="NEXT")
    pdf.set_x(15)
    pdf.set_font("Helvetica", "", 10.5)
    pdf.cell(0, 6, _pdf_safe(f"Atleta: {nome}   Sesso: {sesso}   Data test: {periodo}"),
             new_x="LMARGIN", new_y="NEXT")
    pdf.set_y(32)

    pdf.body_text(
        "Per valutare i risultati del test e stato utilizzato il T-Score, un indice standardizzato "
        "che confronta la prestazione dell'atleta rispetto a un gruppo di riferimento, esprimendo la "
        "distanza dalla media in deviazioni standard. Punteggi tra 0 e 50 indicano valori inferiori "
        "alla media, mentre punteggi tra 50 e 100 indicano valori superiori alla media. DSI ed EUR "
        "fanno eccezione: essendo rapporti tra due test, non vengono letti come 'piu alto = meglio' "
        "ma per zona di profilo, rispetto alle soglie indicate accanto a ciascun grafico."
    )
    pdf.ln(3)

    # --- Profilo di Forza ---
    cats_valide = [c for c in CATEGORIES if profilo.get(c) is not None]
    pdf.section_title("Profilo di Forza")
    radar_fig = build_radar_chart(cats_valide, profilo, nome)
    if radar_fig:
        pdf.chart_image(radar_fig, width_px=1000, height_px=620, content_width_mm=150)
    else:
        pdf.body_text("Nessuna metrica con confronto di popolazione disponibile.")
    pdf.profile_cards(cats_valide, profilo)

    pdf.add_page()

    # --- Categorie di test ---
    for cat in CATEGORIES:
        cat_results = [r for r in results if r["category"] == cat and r["mean"] is not None
                       and r["kind"] != "support"]
        if not cat_results:
            continue
        pdf.section_title(cat)
        tscore_fig = build_tscore_bar_chart(cat_results)
        if tscore_fig:
            n_metrics = len([r for r in cat_results if r["t"] is not None])
            h_px = max(360, 70 * n_metrics + 150)
            pdf.chart_image(tscore_fig, width_px=1000, height_px=h_px, content_width_mm=180)
        pdf.metric_table(cat_results)
        pdf.ln(1)
        rsq_fig = build_rsq_chart(cat, results)
        if rsq_fig:
            pdf.subsection_title("RSQ - Reactive Strength Quadrant")
            pdf.chart_image(rsq_fig, width_px=900, height_px=430, content_width_mm=150)
        pdf.descriptions_block(cat_results)
        pdf.add_page()

    # --- Indici (DSI, EUR) ---
    r_dsi = next((r for r in results if r["key"] == "dsi" and r["mean"] is not None), None)
    r_eur = next((r for r in results if r["key"] == "eur" and r["mean"] is not None), None)
    dsi_fig = build_dsi_chart(results, thresholds)
    eur_fig = build_eur_chart(results, thresholds)
    if r_dsi or r_eur or dsi_fig or eur_fig:
        pdf.section_title("Indici")
        for r_idx, fig_idx, titolo, idx_key in (
            (r_dsi, dsi_fig, "DSI (Dynamic Strength Index)", "dsi"),
            (r_eur, eur_fig, "EUR (Eccentric Utilisation Ratio)", "eur"),
        ):
            if not (r_idx or fig_idx):
                continue
            pdf.subsection_title(titolo)
            if r_idx:
                lo, hi = thresholds[idx_key]
                zona = f"  ({r_idx['zona']})" if r_idx.get("zona") else ""
                pop = f"   Media pop. {r_idx['pop_mean']:.3f}" if r_idx["pop_mean"] is not None else ""
                pdf.body_text(f"{idx_key.upper()}: {r_idx['mean']:.3f}{zona}{pop}   "
                              f"soglie {lo:.2f} / {hi:.2f}")
                # height della figura e height_px di kaleido DEVONO coincidere:
                # lo spessore della fascia e' (height - margini) dell'altezza
                # effettivamente renderizzata, non di quella dichiarata nella
                # figura. Se divergono, in PDF la fascia diventa un blocco.
                strip = ratio_band_strip(idx_key, r_idx["mean"], lo, hi,
                                         label=idx_key.upper(), decimals=3)
                if strip:
                    pdf.chart_image(strip, width_px=900, height_px=strip.layout.height,
                                    content_width_mm=150)
            if fig_idx:
                pdf.chart_image(fig_idx, width_px=900, height_px=470, content_width_mm=150)
            pdf.descriptions_block([r_idx])
            pdf.add_page()

    # --- Comparazione con lo storico (sezione opzionale) ---
    comparazione_sezione_pdf(pdf, comp, nome)

    # --- Analisi ---
    # add_page() perche' la sezione Comparazione lascia la pagina a metà.
    pdf.add_page()
    pdf.section_title("Analisi")
    testo = commento.strip() if commento and commento.strip() else "Nessuna analisi inserita."
    pdf.body_text(testo)

    return bytes(pdf.output())



# ############################################################################
# SCHEDA 📄 REPORT
# ############################################################################
def render_report(parsed_files, nome, sesso, periodo, results, profilo):
    if not parsed_files:
        st.info("Carica dei file dalla barra laterale per generare il report.")
    else:
        st.markdown(
            "Scegli in che formato generare il report: **HTML interattivo** (grafici zoomabili, pensato per "
            "desktop/laptop) oppure **PDF statico** (più facilmente visualizzabile da qualsiasi dispositivo)."
        )
        st.caption(
            "Se hai usato la scheda 🔀 Comparazione in questa sessione, il report include anche "
            "il confronto con lo storico: tabella, strisce delle metriche selezionate e profilo "
            "di forza attuale vs storico."
        )
        st.markdown("**Analisi del preparatore**")
        st.caption(
            "Scrivi qui il commento tecnico da includere nel report: punti di forza, aree di "
            "miglioramento e indicazioni di lavoro."
        )
        commento = st.text_area(
            "Analisi del preparatore", key="coach_comment", height=160, label_visibility="collapsed",
            placeholder="Es. Buoni valori di forza isometrica, mRSI sopra media. Da lavorare sulla "
                        "reattività nel CMJ Rebound...",
        )
        col_html, col_pdf = st.columns(2)
        with col_html:
            if st.button("📄 Genera report HTML", type="primary", use_container_width=True):
                html_bytes = genera_report_html(nome, sesso, periodo, results, profilo, commento,
                                                st.session_state["idx_thr"],
                                                comp=st.session_state.get("comp_export"))
                st.download_button(
                    "⬇️ Scarica report (.html)", data=html_bytes,
                    file_name=f"Report_{nome.replace(' ', '_')}_{dt.date.today().isoformat()}.html",
                    mime="text/html", use_container_width=True,
                )
        with col_pdf:
            if st.button("📕 Genera report PDF", use_container_width=True):
                with st.spinner("Rendering grafici e impaginazione PDF..."):
                    pdf_bytes = genera_report_pdf(nome, sesso, periodo, results, profilo, commento,
                                                  st.session_state["idx_thr"],
                                                  comp=st.session_state.get("comp_export"))
                st.download_button(
                    "⬇️ Scarica report (.pdf)", data=pdf_bytes,
                    file_name=f"Report_{nome.replace(' ', '_')}_{dt.date.today().isoformat()}.pdf",
                    mime="application/pdf", use_container_width=True,
                )