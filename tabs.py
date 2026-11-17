"""Schede 🔍 Dettaglio Test e 📊 Profilo di Forza."""

from __future__ import annotations

import pandas as pd
import streamlit as st

from core import (
    RSQ_METRIC_KEYS,
    TEXT_COLOR,
    _data_file_token,
    _slug_nome,
    best_worst_indices,
    build_dsi_chart,
    build_eur_chart,
    build_extra_defs,
    build_full_raw_export,
    build_radar_chart,
    build_rsq_chart,
    build_tscore_bar_chart,
    collect_reps_all,
    nel_profilo,
    per_rep_metric_values,
    ratio_band_strip,
)
from data import (
    CATEGORIES,
    CATEGORY_QUALITY,
    EXTRA_METRICS_CATALOG,
    METRICS,
    banda_da_tscore,
)

# ############################################################################
# SCHEDA 🔍 DETTAGLIO TEST
# ############################################################################
# Stile delle righe "support" (metriche solo per i Controlli Tecnici) nella
# tabella dati del Dettaglio Test: azzurro chiaro della palette brand.
SUPPORT_ROW_STYLE = "background-color:#dff4fd; color:#065678; font-style:italic;"


def render_dettaglio(parsed_files, nome, sesso, periodo, results, checks):
    if not parsed_files:
        st.info("Carica dei file dalla barra laterale per vedere il dettaglio dei test.")
    else:
        for cat in CATEGORIES:
            cat_results = [r for r in results if r["category"] == cat and r["mean"] is not None]
            if not cat_results:
                continue
            st.markdown(f"### {cat}")

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
            extra_metric_defs = build_extra_defs(jump_type, cat)

            st.markdown("**Ripetizioni incluse nel calcolo**")
            st.caption("Deseleziona un salto per escluderlo da media, dev.std e T-score di questa categoria.")
            chk_cols = st.columns(n_reps)
            incl_mask = []
            for i in range(n_reps):
                key = f"incl_{jump_type}_{i}"
                checked = chk_cols[i].checkbox(f"Prova {i + 1}", value=st.session_state.get(key, True), key=key)
                incl_mask.append(checked)

            # Le metriche extra sono già calcolate in build_results().
            extra_keys = {m["key"] for m in extra_metric_defs}
            empty_extra_labels = [r["label"] for r in results
                                  if r["key"] in extra_keys and r["n"] == 0]
            if empty_extra_labels:
                st.warning(
                    "⚠️ Nessun dato trovato per: " + ", ".join(empty_extra_labels) + ". La variabile "
                    "potrebbe non essere presente in questo export ForceDecks (es. campo non calcolato "
                    "dal software per questo tipo di test)."
                )
            results_lookup = results
            cat_metrics = cat_metrics + extra_metric_defs

            rows, best_per_row, worst_per_row = [], [], []
            support_per_row = []  # True = metrica solo per i Controlli Tecnici
            cv_warnings = []
            metric_descriptions = []  # (label, desc) per l'expander sotto la tabella
            jump_cols = [f"Prova {i + 1}" for i in range(n_reps)]
            for m in cat_metrics:
                r = next((x for x in results_lookup if x["key"] == m["key"]), None)
                values = per_rep_metric_values(parsed_files, m)
                best_i, worst_i = best_worst_indices(values, incl_mask, m["lower_is_better"])
                best_per_row.append(best_i)
                worst_per_row.append(worst_i)
                support_per_row.append(m["kind"] == "support")
                desc = m.get("desc")
                if desc:
                    metric_descriptions.append((m["label"], desc))
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
                if m["kind"] == "support":
                    row["Valutazione"] = "Solo controllo"
                rows.append(row)

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
                if support_per_row[ridx]:
                    for col in ("Metrica", "Unità", "Media Pop.", "T-score", "Valutazione"):
                        styles[df.columns.get_loc(col)] = SUPPORT_ROW_STYLE
                cv_val = row.get("CV%")
                if isinstance(cv_val, (int, float)) and cv_val > 10:
                    styles[df.columns.get_loc("CV%")] = "background-color:#E4572E; color:white; font-weight:600;"
                return styles

            styled = df.style.apply(_highlight, axis=1)
            st.dataframe(styled, use_container_width=True, hide_index=True)
            if any(support_per_row):
                st.caption(
                    "🟦 Righe azzurre: metriche usate solo per i Controlli Tecnici. "
                    "Sono calcolate sulle prove incluse ma non hanno T-score e non "
                    "entrano nel Profilo di Forza."
                )

            if cv_warnings:
                st.warning(
                    "⚠️ Coefficiente di variazione (CV%) superiore al 10% per: " + ", ".join(cv_warnings) +
                    ". Indica un'elevata variabilità tra le ripetizioni incluse: valutarne l'affidabilità."
                )
            if metric_descriptions:
                with st.expander("ℹ️ Cosa significano queste metriche?"):
                    for label, desc in metric_descriptions:
                        st.markdown(f"**{label}** — {desc}")

            # --- Profilo di forza: scelta delle metriche, poi il grafico ---
            # Il grafico a barre mostra solo le metriche spuntate: è lo stesso
            # insieme che compone il T-score della categoria nel profilo.
            scored_cat = [r for r in results if r["category"] == cat and r["t"] is not None]
            if scored_cat:
                st.markdown("**Metriche incluse nel Profilo di Forza**")
                st.caption(
                    "Il T-score della categoria nel profilo (radar, card e report) è la media delle "
                    "metriche spuntate. Di default sono tutte spuntate."
                )
                pcols = st.columns(min(len(scored_cat), 4))
                for i, r in enumerate(scored_cat):
                    k = f"prof_{r['key']}"
                    pcols[i % len(pcols)].checkbox(
                        r["label"], value=st.session_state.get(k, True), key=k)

                tscore_fig = build_tscore_bar_chart([r for r in cat_results if nel_profilo(r)])
                if tscore_fig:
                    st.plotly_chart(tscore_fig, use_container_width=True)
                else:
                    st.info("Nessuna metrica spuntata: la categoria non entra nel Profilo di Forza.")

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

        # --- Export unico: TUTTE le variabili presenti nei file XLSX, una
        # riga per ripetizione, escluse quelle deselezionate. È lo stesso
        # file che si ricarica dalla sidebar e che alimenta la Comparazione.
        st.markdown("### 📤 Esporta dati del test")
        st.caption(
            "Esporta OGNI variabile presente nei file caricati, una riga per ripetizione, "
            "escluse le prove deselezionate sopra. È il file da ricaricare nella barra "
            "laterale per rivedere questo test e da usare nella scheda 🔀 Comparazione."
        )
        df_full = build_full_raw_export(parsed_files, nome, sesso, periodo)
        csv_full_bytes = df_full.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "⬇️ Scarica dati completi (.csv)", data=csv_full_bytes,
            file_name=f"forceplate_fulldata_{_slug_nome(nome)}_"
                      f"{_data_file_token(parsed_files, periodo)}.csv",
            mime="text/csv",
        )

# ############################################################################
# SCHEDA 📊 PROFILO DI FORZA
# ############################################################################
def render_profilo(parsed_files, nome, results, profilo):
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
                with col.container(border=True):
                    st.markdown(
                        f"<div style='font-size:11px;letter-spacing:0.5px;opacity:0.65'>{cat}</div>"
                        f"<div style='font-size:17px;font-weight:600;color:{TEXT_COLOR};"
                        f"margin:2px 0 4px 0'>{CATEGORY_QUALITY.get(cat, '')}</div>"
                        f"<div style='font-size:30px;font-weight:700;color:{colore}'>{t:.0f}</div>"
                        f"<div style='font-size:13px;opacity:0.75'>{banda}</div>",
                        unsafe_allow_html=True,
                    )

        # --- Indici a rapporto (DSI, EUR): barra a bande + grafico a spicchi.
        # Le soglie si modificano nella scheda ⚙️ Costanti e si applicano
        # subito anche ai report scaricabili. Le etichette delle zone
        # descrivono il PROFILO dell'atleta, non l'indicazione di
        # allenamento: quella resta una decisione del preparatore, da
        # scrivere nel campo "Analisi" della scheda Report.
        for idx_key, titolo, decimali, caption in (
            ("dsi", "DSI (Dynamic Strength Index)", 3,
             "CMJ Peak Propulsive Force vs IMTP Peak Force. Le linee tratteggiate arancioni sono la "
             "media di popolazione e il punto è la media del test con barre d'errore"),
            ("eur", "EUR (Eccentric Utilisation Ratio)", 3,
             "CMJ Height vs SJ Height. Le linee tratteggiate arancioni sono la media di popolazione."),
        ):
            r_idx = next((r for r in results if r["key"] == idx_key and r["mean"] is not None), None)
            fig_wedge = (build_dsi_chart if idx_key == "dsi" else build_eur_chart)(
                results, st.session_state["idx_thr"])
            if not fig_wedge and not r_idx:
                continue

            st.markdown(f"#### {titolo}")
            if r_idx:
                help_parts = []
                if r_idx.get("zona"):
                    help_parts.append(r_idx["zona"])
                if r_idx["pop_mean"] is not None:
                    help_parts.append(f"Media popolazione: {r_idx['pop_mean']:.3f}")
                lo, hi = st.session_state["idx_thr"][idx_key]
                help_parts.append(f"Soglie {lo:.2f} / {hi:.2f} (modificabili in ⚙️ Costanti)")
                st.metric(idx_key.upper(), f"{r_idx['mean']:.{decimali}f}",
                          help=" — ".join(help_parts))
                strip = ratio_band_strip(idx_key, r_idx["mean"], lo, hi,
                                         label=idx_key.upper(), decimals=decimali)
                if strip:
                    st.plotly_chart(strip, use_container_width=True)
            if fig_wedge:
                st.plotly_chart(fig_wedge, use_container_width=True)
                st.caption(caption)