"""Dati statici: categorie, bande T-score, soglie DSI/EUR, norme di popolazione,
METRICS/CHECKS, catalogo metriche extra e descrizioni (help "?").

NB: in fondo al file METRICS[*]["desc"] viene popolato all'import."""

from __future__ import annotations

import io
import re

import openpyxl
import pandas as pd

# ============================================================================
# PARTE 1 — COSTANTI E DATI DI POPOLAZIONE
# ============================================================================
# Equivalente ai fogli "DATI POP" (norme di popolazione) e "DATI TEST"
# (definizione delle metriche) del Google Sheet originale.

CATEGORIES = ["ISOMETRIC PULL TEST", "SQUAT JUMP TEST", "COUNTERMOVEMENT JUMP TEST", "COUNTERMOVEMENT JUMP REBOUND TEST"]

# Qualità fisica misurata da ciascun test: il nome del protocollo dice COME si
# è misurato, l'etichetta dice COSA significa. Serve all'atleta, che non ha
# motivo di sapere cosa sia un "countermovement jump rebound".
CATEGORY_QUALITY = {
    "ISOMETRIC PULL TEST": "Forza massima",
    "SQUAT JUMP TEST": "Potenza",
    "COUNTERMOVEMENT JUMP TEST": "Esplosività",
    "COUNTERMOVEMENT JUMP REBOUND TEST": "Reattività",
}

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


# Soglie di lettura degli indici-rapporto (DSI, EUR): delimitano le tre zone
# di profilo mostrate nei grafici. Valori di DEFAULT, modificabili nella
# scheda "⚙️ Costanti" (a runtime si usa sempre st.session_state["idx_thr"]).
#   DSI  0.60 / 0.80  — Sheppard & Chapman (2011), via Science for Sport
#   EUR  1.00 / 1.10  — McGuigan et al. (2006), via VALD / Clubb
# NB: le soglie EUR di letteratura cadono praticamente sulla media di
# popolazione qui usata (1.108 U / 1.091 D): con i default un atleta "medio"
# finisce sul confine superiore. Da ricalibrare sui propri dati.
DEFAULT_INDEX_THRESHOLDS = {"dsi": (0.60, 0.80), "eur": (1.00, 1.10)}

# Etichette descrittive (profilo dell'atleta, NON indicazione di allenamento)
# e colori delle tre zone, in ordine basso / medio / alto.
INDEX_ZONE_LABELS = {
    "dsi": ("Atleta più forte che esplosivo", "Atleta bilanciato", "Atleta più esplosivo che forte"),
    "eur": ("Poco contributo elastico", "Contributo elastico medio", "Buon contributo elastico"),
}
INDEX_ZONE_COLORS = ("#FFB74D", "#81C784", "#4FC3F7")


def zona_da_indice(key, value, thr_low, thr_high):
    """Zona di profilo (etichetta, colore) per un indice-rapporto."""
    labels = INDEX_ZONE_LABELS.get(key)
    if value is None or labels is None:
        return None, None
    i = 0 if value < thr_low else (2 if value > thr_high else 1)
    return labels[i], INDEX_ZONE_COLORS[i]


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
LITERATURE_POP = {
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

# Norme del team DU Coaching, calcolate sulle sedute del laboratorio.
# Le chiavi curate (imtp_peak_force, ..., mrsi_cmj_re) sono quelle usate per
# T-score e SWC; le chiavi 'tipo::variabile' sono conservate per usi futuri
# (oggi non lette dall'app). Per le donne (n=0) i valori curati sono quelli
# di letteratura.
TEAM_POP = {
    "imtp_peak_force":                                            dict(mean_m=2162.050497,   sd_m=187.1258932,    mean_f=1575.0,          sd_f=386.145544),  # M n=3, F n=0 default
    "imtp_rel_peak_force":                                        dict(mean_m=27.28493086,   sd_m=0.5046443763,   mean_f=34.56,           sd_f=5.27),  # M n=3, F n=0 default
    "sj_mean_power":                                              dict(mean_m=1025.70804,    sd_m=260.0398012,    mean_f=742.0900714,     sd_f=202.6636875),  # M n=27, F n=7
    "sj_height":                                                  dict(mean_m=21.93888329,   sd_m=4.949861273,    mean_f=15.73722381,     sd_f=3.870400084),  # M n=27, F n=7
    "sj_contraction_time":                                        dict(mean_m=0.5386239352,  sd_m=0.2871596477,   mean_f=0.4117931667,    sd_f=0.09942743726),  # M n=27, F n=7
    "cmj_height":                                                 dict(mean_m=25.84610516,   sd_m=4.423413661,    mean_f=18.37357143,     sd_f=3.892073811),  # M n=29, F n=7
    "mrsi_cmj":                                                   dict(mean_m=0.2423121312,  sd_m=0.07914031445,  mean_f=0.1827452381,    sd_f=0.0580066611),  # M n=29, F n=7
    "dsi":                                                        dict(mean_m=0.7442202762,  sd_m=0.03847254661,  mean_f=0.7,             sd_f=0.074074074),  # M n=3, F n=0 default
    "eur":                                                        dict(mean_m=1.195112981,   sd_m=0.1799391346,   mean_f=1.177256461,     sd_f=0.08579380295),  # M n=27, F n=7
    "cmj_re_rebound_height":                                      dict(mean_m=18.21779109,   sd_m=5.562508602,    mean_f=38.5,            sd_f=5.5996817),  # M n=6, F n=0 default
    "cmj_re_contact_time":                                        dict(mean_m=0.2273584028,  sd_m=0.04543635166,  mean_f=0.25,            sd_f=0.148005087),  # M n=6, F n=0 default
    "cmj_re_rebound_impulse":                                     dict(mean_m=147.2672436,   sd_m=29.24540556,    mean_f=541.5,           sd_f=53.82161458),  # M n=6, F n=0 default
    "mrsi_cmj_re":                                                dict(mean_m=0.9546057997,  sd_m=0.5532670785,   mean_f=1.37,            sd_f=0.331705729),  # M n=6, F n=0 default

    # ======================================================================
    # ALTRE METRICHE — derivate (chiave di METRICS), poi grezze
    # con chiave 'tipo di test::variabile'.
    # None = troppe poche sedute per quel sesso (nessun default).
    # ======================================================================
    "sj_net_rel_impulse":                                         dict(mean_m=2.068984343,   sd_m=0.2052907488,   mean_f=1.747268559,     sd_f=0.289720883),  # M n=27, F n=7
    "cmj_net_rel_impulse":                                        dict(mean_m=2.164466521,   sd_m=0.2187441191,   mean_f=1.83811986,      sd_f=0.2473568718),  # M n=29, F n=7
    "cmj::avg braking force sym. index":                          dict(mean_m=-1.283874338,  sd_m=4.567905442,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::avg braking power sym. index":                          dict(mean_m=-2.014679971,  sd_m=5.147912942,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::avg eccentric force":                                   dict(mean_m=745.4192827,   sd_m=62.67600186,    mean_f=579.6553571,     sd_f=71.64865143),  # M n=29, F n=7
    "cmj::avg eccentric force sym. index":                        dict(mean_m=-2.123787774,  sd_m=3.05450951,     mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::avg eccentric power":                                   dict(mean_m=425.7661879,   sd_m=159.5542961,    mean_f=243.5757143,     sd_f=91.713052),  # M n=29, F n=7
    "cmj::avg eccentric power sym. index":                        dict(mean_m=4.615467338,   sd_m=1.899865328,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::avg eccentric velocity":                                dict(mean_m=-0.5614223677, sd_m=0.1778622353,   mean_f=-0.4223916667,   sd_f=0.1488609183),  # M n=29, F n=7
    "cmj::avg landing rfd":                                       dict(mean_m=75580.07792,   sd_m=80266.55714,    mean_f=18766.88548,     sd_f=7713.656784),  # M n=29, F n=7
    "cmj::avg landing rfd sym. index":                            dict(mean_m=-10.73600121,  sd_m=20.83579186,    mean_f=-8.228892857,    sd_f=21.50964052),  # M n=29, F n=7
    "cmj::avg propulsive force sym. index":                       dict(mean_m=0.9211969641,  sd_m=3.220836622,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::avg propulsive power sym. index":                       dict(mean_m=0.3460774508,  sd_m=3.700549259,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::avg. braking force":                                    dict(mean_m=831.4935555,   sd_m=89.34922972,    mean_f=656.7675,        sd_f=80.99452346),  # M n=29, F n=7
    "cmj::avg. braking power":                                    dict(mean_m=-571.4624859,  sd_m=219.9142427,    mean_f=-343.7313095,    sd_f=137.5949572),  # M n=29, F n=7
    "cmj::avg. braking velocity":                                 dict(mean_m=-0.7215096355, sd_m=0.2317863981,   mean_f=-0.5591919048,   sd_f=0.2011303812),  # M n=29, F n=7
    "cmj::avg. propulsive force":                                 dict(mean_m=1195.192314,   sd_m=112.977886,     mean_f=903.7632143,     sd_f=100.6700524),  # M n=29, F n=7
    "cmj::avg. propulsive power":                                 dict(mean_m=1419.991481,   sd_m=247.8939796,    mean_f=965.7507143,     sd_f=206.2718446),  # M n=29, F n=7
    "cmj::avg. propulsive velocity":                              dict(mean_m=1.249643622,   sd_m=0.1496184117,   mean_f=1.149063214,     sd_f=0.1868537814),  # M n=29, F n=7
    "cmj::avg. rfd":                                              dict(mean_m=1066.100142,   sd_m=519.405147,     mean_f=779.8270238,     sd_f=284.1496252),  # M n=29, F n=7
    "cmj::avg. rfd sym. index":                                   dict(mean_m=6.044548984,   sd_m=14.46858527,    mean_f=-2.961511905,    sd_f=11.65234951),  # M n=29, F n=7
    "cmj::body mass":                                             dict(mean_m=76.01744807,   sd_m=6.386407519,    mean_f=59.10444762,     sd_f=7.307141415),  # M n=29, F n=7
    "cmj::body weight":                                           dict(mean_m=745.4763463,   sd_m=62.62948753,    mean_f=579.6159524,     sd_f=71.65834131),  # M n=29, F n=7
    "cmj::body weight sd":                                        dict(mean_m=8.001687375,   sd_m=4.031599956,    mean_f=6.997619048,     sd_f=0.04597812627),  # M n=29, F n=7
    "cmj::braking duration":                                      dict(mean_m=0.5051323372,  sd_m=0.1823467476,   mean_f=0.4477560714,    sd_f=0.1547879393),  # M n=29, F n=7
    "cmj::braking end time":                                      dict(mean_m=1.238084005,   sd_m=0.1922224867,   mean_f=1.16462019,      sd_f=0.1289832579),  # M n=29, F n=7
    "cmj::braking impulse":                                       dict(mean_m=36.9475585,    sd_m=13.9339882,     mean_f=25.61028571,     sd_f=10.9055815),  # M n=29, F n=7
    "cmj::braking impulse sym. index":                            dict(mean_m=-1.086503598,  sd_m=4.65030897,     mean_f=-2.868,          sd_f=10.43484252),  # M n=29, F n=7
    "cmj::braking rfd":                                           dict(mean_m=2433.042695,   sd_m=1339.749901,    mean_f=1923.350952,     sd_f=1311.921491),  # M n=29, F n=7
    "cmj::braking rfd sym. index":                                dict(mean_m=-3.258399006,  sd_m=11.11789102,    mean_f=-6.610190476,    sd_f=16.58915156),  # M n=29, F n=7
    "cmj::braking work":                                          dict(mean_m=-265.3351088,  sd_m=84.62152654,    mean_f=-137.6044048,    sd_f=49.48938667),  # M n=29, F n=7
    "cmj::braking work sym. index":                               dict(mean_m=-2.014679971,  sd_m=5.147912942,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::contact time":                                          dict(mean_m=1.080237668,   sd_m=0.2548163801,   mean_f=1.011172262,     sd_f=0.1552416033),  # M n=29, F n=7
    "cmj::decel rfd":                                             dict(mean_m=2496.173244,   sd_m=1754.937924,    mean_f=2364.419762,     sd_f=2147.632715),  # M n=29, F n=7
    "cmj::decel rfd sym. index":                                  dict(mean_m=-5.073555755,  sd_m=15.20120724,    mean_f=13.19258333,     sd_f=59.0494613),  # M n=29, F n=7
    "cmj::displacement depth":                                    dict(mean_m=36.590639,     sd_m=8.757035117,    mean_f=26.66908333,     sd_f=6.886689241),  # M n=29, F n=7
    "cmj::eccentric impulse":                                     dict(mean_m=0.3179454925,  sd_m=0.1489970417,   mean_f=0.1834880952,    sd_f=0.08639428925),  # M n=29, F n=7
    "cmj::eccentric impulse sym. index":                          dict(mean_m=4.027171659,   sd_m=1.702269482,    mean_f=6.981833333,     sd_f=4.669336443),  # M n=29, F n=7
    "cmj::eccentric rfd":                                         dict(mean_m=1222.715074,   sd_m=391.8417603,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::eccentric rfd sym. index":                              dict(mean_m=12.08988933,   sd_m=18.55535311,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::eccentric work":                                        dict(mean_m=276.4773863,   sd_m=84.45775619,    mean_f=155.5054762,     sd_f=50.71117946),  # M n=29, F n=7
    "cmj::eccentric work sym. index":                             dict(mean_m=4.615467338,   sd_m=1.899865328,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::flight threshold":                                      dict(mean_m=4.128647027,   sd_m=0.8213404967,   mean_f=3.834642857,     sd_f=0.8806342224),  # M n=29, F n=7
    "cmj::flight time":                                           dict(mean_m=0.4571966504,  sd_m=0.03934019512,  mean_f=0.3852175,       sd_f=0.04072917934),  # M n=29, F n=7
    "cmj::force at min displacement":                             dict(mean_m=1320.630716,   sd_m=247.6553851,    mean_f=996.7366667,     sd_f=183.4907541),  # M n=29, F n=7
    "cmj::force peak power":                                      dict(mean_m=1349.308908,   sd_m=93.16028104,    mean_f=1002.725833,     sd_f=105.5555289),  # M n=29, F n=7
    "cmj::initiation threshold":                                  dict(mean_m=725.4763463,   sd_m=62.62948753,    mean_f=559.6159524,     sd_f=71.65834131),  # M n=29, F n=7
    "cmj::jump height ni":                                        dict(mean_m=24.21058136,   sd_m=4.841453487,    mean_f=17.51415476,     sd_f=4.691357054),  # M n=29, F n=7
    "cmj::jump momentum":                                         dict(mean_m=163.9341166,   sd_m=16.11722175,    mean_f=108.6190952,     sd_f=18.69768302),  # M n=29, F n=7
    "cmj::jump start time":                                       dict(mean_m=0.5351612563,  sd_m=0.09560935146,  mean_f=0.494419,        sd_f=0.02648049526),  # M n=29, F n=7
    "cmj::jump threshold time":                                   dict(mean_m=0.5351612563,  sd_m=0.09560935146,  mean_f=0.494419,        sd_f=0.02648049526),  # M n=29, F n=7
    "cmj::landing peak force time":                               dict(mean_m=2.136692238,   sd_m=0.2365313852,   mean_f=1.977145667,     sd_f=0.1362204257),  # M n=29, F n=7
    "cmj::landing rfd 0-20ms":                                    dict(mean_m=133723.7332,   sd_m=103480.955,     mean_f=None,            sd_f=None),  # M n=11, F n=0
    "cmj::landing rfd 0-20ms sym. index":                         dict(mean_m=-14.66682641,  sd_m=20.95357809,    mean_f=None,            sd_f=None),  # M n=11, F n=0
    "cmj::landing rfd 0-40ms":                                    dict(mean_m=51077.8562,    sd_m=18642.0073,     mean_f=None,            sd_f=None),  # M n=11, F n=0
    "cmj::landing rfd 0-40ms sym. index":                         dict(mean_m=-2.873608939,  sd_m=13.64968472,    mean_f=None,            sd_f=None),  # M n=11, F n=0
    "cmj::landing rfd 0-60ms":                                    dict(mean_m=30092.43387,   sd_m=11456.27427,    mean_f=None,            sd_f=None),  # M n=11, F n=0
    "cmj::landing rfd 0-60ms sym. index":                         dict(mean_m=-8.516970208,  sd_m=15.53101063,    mean_f=None,            sd_f=None),  # M n=11, F n=0
    "cmj::landing rfd 0-80ms":                                    dict(mean_m=21732.44946,   sd_m=5814.751302,    mean_f=None,            sd_f=None),  # M n=11, F n=0
    "cmj::landing rfd 0-80ms sym. index":                         dict(mean_m=10.09340666,   sd_m=18.82619665,    mean_f=None,            sd_f=None),  # M n=11, F n=0
    "cmj::landing time":                                          dict(mean_m=2.072583858,   sd_m=0.2323587047,   mean_f=1.890811738,     sd_f=0.1466932824),  # M n=29, F n=7
    "cmj::left avg braking force":                                dict(mean_m=422.3928235,   sd_m=43.86876305,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::left avg braking power":                                dict(mean_m=-329.2849762,  sd_m=80.20487688,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::left avg eccentric force":                              dict(mean_m=363.4733604,   sd_m=29.62052833,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::left avg eccentric power":                              dict(mean_m=231.5180989,   sd_m=62.88726217,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::left avg landing rfd":                                  dict(mean_m=38278.88664,   sd_m=37006.24225,    mean_f=9940.630238,     sd_f=4949.647224),  # M n=29, F n=7
    "cmj::left avg propulsive force":                             dict(mean_m=608.5094535,   sd_m=35.74745391,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::left avg propulsive power":                             dict(mean_m=726.8364654,   sd_m=59.19256686,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::left avg rfd":                                          dict(mean_m=566.6538118,   sd_m=264.3752168,    mean_f=392.1611905,     sd_f=148.4841529),  # M n=29, F n=7
    "cmj::left braking impulse":                                  dict(mean_m=204.6393793,   sd_m=61.10919352,    mean_f=138.9076905,     sd_f=44.04923923),  # M n=29, F n=7
    "cmj::left braking rfd":                                      dict(mean_m=1218.042251,   sd_m=693.5061294,    mean_f=943.5578571,     sd_f=657.788664),  # M n=29, F n=7
    "cmj::left braking work":                                     dict(mean_m=-136.517673,   sd_m=40.71763253,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::left decel rfd":                                        dict(mean_m=1202.01287,    sd_m=832.6476837,    mean_f=1138.89,         sd_f=1010.802827),  # M n=29, F n=7
    "cmj::left eccentric impulse":                                dict(mean_m=256.413889,    sd_m=75.83192136,    mean_f=189.6994405,     sd_f=45.77027153),  # M n=29, F n=7
    "cmj::left eccentric rfd":                                    dict(mean_m=732.448205,    sd_m=176.4833185,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::left eccentric work":                                   dict(mean_m=142.164055,    sd_m=39.9518942,     mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::left landing rfd 0-20ms":                               dict(mean_m=59956.79218,   sd_m=49075.02806,    mean_f=None,            sd_f=None),  # M n=11, F n=0
    "cmj::left landing rfd 0-40ms":                               dict(mean_m=24475.52437,   sd_m=8848.107799,    mean_f=None,            sd_f=None),  # M n=11, F n=0
    "cmj::left landing rfd 0-60ms":                               dict(mean_m=12897.01427,   sd_m=4229.621983,    mean_f=None,            sd_f=None),  # M n=11, F n=0
    "cmj::left landing rfd 0-80ms":                               dict(mean_m=11907.87022,   sd_m=3560.886867,    mean_f=None,            sd_f=None),  # M n=11, F n=0
    "cmj::left net impulse":                                      dict(mean_m=91.90915858,   sd_m=35.64805302,    mean_f=46.01940476,     sd_f=20.30572621),  # M n=29, F n=7
    "cmj::left p1 impulse":                                       dict(mean_m=107.5969954,   sd_m=17.4082105,     mean_f=65.94152381,     sd_f=17.02713105),  # M n=29, F n=7
    "cmj::left p2 impulse":                                       dict(mean_m=113.0814973,   sd_m=19.6794222,     mean_f=84.78896429,     sd_f=16.85055704),  # M n=29, F n=7
    "cmj::left peak braking force":                               dict(mean_m=727.9164598,   sd_m=77.44833963,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::left peak braking power":                               dict(mean_m=-533.7599268,  sd_m=153.4906374,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::left peak eccentric force":                             dict(mean_m=729.8455074,   sd_m=75.46030974,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::left peak eccentric power":                             dict(mean_m=533.7599268,   sd_m=153.4906374,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::left peak force":                                       dict(mean_m=744.2566333,   sd_m=73.56977433,    mean_f=552.1463095,     sd_f=71.8399128),  # M n=29, F n=7
    "cmj::left peak landing force":                               dict(mean_m=1968.411975,   sd_m=604.0179334,    mean_f=1075.641667,     sd_f=306.6480585),  # M n=29, F n=7
    "cmj::left peak propulsive force":                            dict(mean_m=749.6438267,   sd_m=60.27617447,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::left peak propulsive power":                            dict(mean_m=1390.075236,   sd_m=197.0865648,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::left propulsive impulse":                               dict(mean_m=220.6785302,   sd_m=29.44718033,    mean_f=150.7306667,     sd_f=33.30365148),  # M n=29, F n=7
    "cmj::left propulsive rfd":                                   dict(mean_m=238.998821,    sd_m=154.7555268,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::left propulsive work":                                  dict(mean_m=257.8816468,   sd_m=48.57305488,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::left time to peak braking force":                       dict(mean_m=0.4089640476,  sd_m=0.07794495693,  mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::left time to peak braking power":                       dict(mean_m=0.2180276786,  sd_m=0.02531508121,  mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::left time to peak eccentric force":                     dict(mean_m=0.5851492857,  sd_m=0.09443814468,  mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::left time to peak eccentric power":                     dict(mean_m=0.5001167262,  sd_m=0.2606272855,   mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::left time to peak force":                               dict(mean_m=0.8734300718,  sd_m=0.3182199552,   mean_f=0.7958702381,    sd_f=0.2238504915),  # M n=29, F n=7
    "cmj::left time to peak landing force":                       dict(mean_m=0.07344472126, sd_m=0.04279809263,  mean_f=0.08740047619,   sd_f=0.01643824827),  # M n=29, F n=7
    "cmj::left time to peak power":                               dict(mean_m=0.9788568452,  sd_m=0.3012327389,   mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::left time to peak propulsive force":                    dict(mean_m=0.1101469643,  sd_m=0.1051945626,   mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::left time to peak propulsive power":                    dict(mean_m=0.2819716071,  sd_m=0.04852137445,  mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::min braking velocity":                                  dict(mean_m=-0.9988823336, sd_m=0.3235025962,   mean_f=-0.785587381,    sd_f=0.2695163129),  # M n=29, F n=7
    "cmj::min eccentric velocity":                                dict(mean_m=-0.9988820539, sd_m=0.3235005315,   mean_f=-0.7855845238,   sd_f=0.26951911),  # M n=29, F n=7
    "cmj::min unweight force":                                    dict(mean_m=322.931018,    sd_m=151.127919,     mean_f=375.9175,        sd_f=109.4623308),  # M n=29, F n=7
    "cmj::net impulse":                                           dict(mean_m=163.9344977,   sd_m=16.11704161,    mean_f=108.6190476,     sd_f=18.69761578),  # M n=29, F n=7
    "cmj::net impulse sym. index":                                dict(mean_m=11.72232041,   sd_m=40.45476649,    mean_f=-21.13422619,    sd_f=35.91525297),  # M n=29, F n=7
    "cmj::p1 avg force":                                          dict(mean_m=1274.718555,   sd_m=174.1358605,    mean_f=1012.2425,       sd_f=148.398949),  # M n=29, F n=7
    "cmj::p1 avg power":                                          dict(mean_m=881.2147977,   sd_m=93.09418531,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::p1 avg velocity":                                       dict(mean_m=0.6921596731,  sd_m=0.04709084327,  mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::p1 duration":                                           dict(mean_m=0.1792500556,  sd_m=0.05276516856,  mean_f=0.1361853571,    sd_f=0.03207944089),  # M n=29, F n=7
    "cmj::p1 impulse":                                            dict(mean_m=86.18354574,   sd_m=10.2085248,     mean_f=54.40139286,     sd_f=9.348430138),  # M n=29, F n=7
    "cmj::p1 impulse sym. index":                                 dict(mean_m=-1.762549892,  sd_m=4.828224717,    mean_f=-2.678547619,    sd_f=7.235339588),  # M n=29, F n=7
    "cmj::p1 p2 duration ratio":                                  dict(mean_m=0.9154253315,  sd_m=0.2620035648,   mean_f=0.66325,         sd_f=0.138618197),  # M n=29, F n=7
    "cmj::p1 p2 force ratio":                                     dict(mean_m=1.11856942,    sd_m=0.1292332331,   mean_f=1.202260714,     sd_f=0.1444295683),  # M n=29, F n=7
    "cmj::p1 p2 power ratio":                                     dict(mean_m=0.4417746539,  sd_m=0.0787736173,   mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::p1 p2 velocity ratio":                                  dict(mean_m=0.3670994876,  sd_m=0.02695620909,  mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::p1 peak force":                                         dict(mean_m=1438.476813,   sd_m=145.5961656,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::p1 peak power":                                         dict(mean_m=1586.213613,   sd_m=164.0349537,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::p1 peak velocity":                                      dict(mean_m=1.29016999,    sd_m=0.0707056485,   mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::p2 avg force":                                          dict(mean_m=1140.60181,    sd_m=89.82393816,    mean_f=844.3629167,     sd_f=95.92365901),  # M n=29, F n=7
    "cmj::p2 avg power":                                          dict(mean_m=2019.531938,   sd_m=219.0206761,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::p2 avg velocity":                                       dict(mean_m=1.883972849,   sd_m=0.05420784586,  mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::p2 duration":                                           dict(mean_m=0.1980591485,  sd_m=0.02767332942,  mean_f=0.2047872619,    sd_f=0.01836124363),  # M n=29, F n=7
    "cmj::p2 impulse":                                            dict(mean_m=78.06844589,   sd_m=12.60440429,    mean_f=54.40139286,     sd_f=9.348430138),  # M n=29, F n=7
    "cmj::p2 impulse sym. index":                                 dict(mean_m=-0.3057448739, sd_m=3.54318127,     mean_f=-2.686738095,    sd_f=4.372825482),  # M n=29, F n=7
    "cmj::p2 peak force":                                         dict(mean_m=1388.484526,   sd_m=119.9621809,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::p2 peak power":                                         dict(mean_m=2768.892486,   sd_m=363.2299718,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::p2 peak velocity":                                      dict(mean_m=2.286741268,   sd_m=0.1224723262,   mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::peak braking force":                                    dict(mean_m=1329.187767,   sd_m=255.9187917,    mean_f=1003.98881,      sd_f=180.1962172),  # M n=29, F n=7
    "cmj::peak braking force sym. index":                         dict(mean_m=2.231996502,   sd_m=3.904065343,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::peak braking power":                                    dict(mean_m=-903.2861798,  sd_m=386.0320261,    mean_f=-535.6211905,    sd_f=219.3164297),  # M n=29, F n=7
    "cmj::peak braking power sym. index":                         dict(mean_m=-1.421507443,  sd_m=4.654902724,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::peak eccentric force":                                  dict(mean_m=1329.187767,   sd_m=255.9187917,    mean_f=1003.98881,      sd_f=180.1962172),  # M n=29, F n=7
    "cmj::peak eccentric force sym. index":                       dict(mean_m=1.747425074,   sd_m=3.631232187,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::peak eccentric power":                                  dict(mean_m=903.2861798,   sd_m=386.0320261,    mean_f=535.6211905,     sd_f=219.3164297),  # M n=29, F n=7
    "cmj::peak eccentric power sym. index":                       dict(mean_m=4.616289021,   sd_m=1.735169264,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::peak force":                                            dict(mean_m=1485.870195,   sd_m=147.5038474,    mean_f=1112.489762,     sd_f=133.2046309),  # M n=29, F n=7
    "cmj::peak force sym. index":                                 dict(mean_m=-0.1586381212, sd_m=2.943141557,    mean_f=-1.590857143,    sd_f=4.506607364),  # M n=29, F n=7
    "cmj::peak force time":                                       dict(mean_m=1.479018887,   sd_m=0.2727479506,   mean_f=1.350297452,     sd_f=0.2110012807),  # M n=29, F n=7
    "cmj::peak landing force":                                    dict(mean_m=3762.483539,   sd_m=1104.441963,    mean_f=2065.308929,     sd_f=510.1916488),  # M n=29, F n=7
    "cmj::peak landing force sym. index":                         dict(mean_m=-7.353858571,  sd_m=13.16039451,    mean_f=-4.3045,         sd_f=12.13235497),  # M n=29, F n=7
    "cmj::peak power":                                            dict(mean_m=2860.136931,   sd_m=378.100051,     mean_f=1869.766429,     sd_f=313.9298769),  # M n=29, F n=7
    "cmj::peak propulsive force":                                 dict(mean_m=1478.877237,   sd_m=139.5580459,    mean_f=1111.963929,     sd_f=133.7642298),  # M n=29, F n=7
    "cmj::peak propulsive force sym. index":                      dict(mean_m=1.301297442,   sd_m=2.448162269,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::peak propulsive power":                                 dict(mean_m=2860.136931,   sd_m=378.100051,     mean_f=1869.766429,     sd_f=313.9298769),  # M n=29, F n=7
    "cmj::peak propulsive power sym. index":                      dict(mean_m=0.1874845024,  sd_m=3.454593022,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::peak propulsive velocity":                              dict(mean_m=2.329324245,   sd_m=0.2144278294,   mean_f=2.040116071,     sd_f=0.2238878461),  # M n=29, F n=7
    "cmj::peak velocity":                                         dict(mean_m=2.329324245,   sd_m=0.2144278294,   mean_f=2.040116071,     sd_f=0.2238878461),  # M n=29, F n=7
    "cmj::propulsive duration":                                   dict(mean_m=0.3773091044,  sd_m=0.06489383675,  mean_f=0.3409710714,    sd_f=0.04556881351),  # M n=29, F n=7
    "cmj::propulsive impulse":                                    dict(mean_m=164.2520248,   sd_m=16.13309524,    mean_f=108.802619,      sd_f=18.69672413),  # M n=29, F n=7
    "cmj::propulsive impulse sym. index":                         dict(mean_m=-1.019429934,  sd_m=3.951657842,    mean_f=-2.705904762,    sd_f=5.500288311),  # M n=29, F n=7
    "cmj::propulsive rfd":                                        dict(mean_m=452.0829327,   sd_m=312.7027469,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::propulsive rfd sym. index":                             dict(mean_m=2.646816695,   sd_m=34.12700339,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::propulsive work":                                       dict(mean_m=523.2500146,   sd_m=64.1348539,     mean_f=325.9303571,     sd_f=71.63100674),  # M n=29, F n=7
    "cmj::propulsive work sym. index":                            dict(mean_m=0.3460774508,  sd_m=3.700549259,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::rel. avg. propulsive force":                            dict(mean_m=15.96179277,   sd_m=0.8039504665,   mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::rel. min unweight force":                               dict(mean_m=0.4423243432,  sd_m=0.2200104011,   mean_f=0.6428619048,    sd_f=0.1487256136),  # M n=29, F n=7
    "cmj::rel. propulsive impulse":                               dict(mean_m=2.138224426,   sd_m=0.1285512741,   mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::relative force at min displacement":                    dict(mean_m=1.766406822,   sd_m=0.2736212203,   mean_f=1.740880952,     sd_f=0.3850055683),  # M n=29, F n=7
    "cmj::relative peak force":                                   dict(mean_m=1.964797331,   sd_m=0.1270118111,   mean_f=None,            sd_f=None),  # M n=11, F n=0
    "cmj::relative peak landing force":                           dict(mean_m=5.007013115,   sd_m=1.259638162,    mean_f=3.538022619,     sd_f=0.6785418318),  # M n=29, F n=7
    "cmj::relative peak power":                                   dict(mean_m=37.86612606,   sd_m=5.781970104,    mean_f=31.70033333,     sd_f=4.520388642),  # M n=29, F n=7
    "cmj::right avg braking force":                               dict(mean_m=432.6739303,   sd_m=32.79886802,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::right avg braking power":                               dict(mean_m=-339.9812961,  sd_m=63.4490855,     mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::right avg eccentric force":                             dict(mean_m=379.2680059,   sd_m=31.5693921,     mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::right avg eccentric power":                             dict(mean_m=240.2657731,   sd_m=55.91345859,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::right avg landing rfd":                                 dict(mean_m=47224.02168,   sd_m=49523.1993,     mean_f=11761.14417,     sd_f=5177.011359),  # M n=29, F n=7
    "cmj::right avg propulsive force":                            dict(mean_m=598.2663882,   sd_m=50.03244899,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::right avg propulsive power":                            dict(mean_m=722.0776267,   sd_m=71.2695762,     mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::right avg rfd":                                         dict(mean_m=513.7414821,   sd_m=269.6889969,    mean_f=412.4552381,     sd_f=135.7648812),  # M n=29, F n=7
    "cmj::right braking impulse":                                 dict(mean_m=207.5917216,   sd_m=62.37242955,    mean_f=150.7294762,     sd_f=64.24221291),  # M n=29, F n=7
    "cmj::right braking rfd":                                     dict(mean_m=1236.360142,   sd_m=677.9186199,    mean_f=979.7930952,     sd_f=664.6437094),  # M n=29, F n=7
    "cmj::right braking work":                                    dict(mean_m=-140.7476396,  sd_m=36.91975367,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::right decel rfd":                                       dict(mean_m=1294.159855,   sd_m=939.4255825,    mean_f=1225.529762,     sd_f=1151.122621),  # M n=29, F n=7
    "cmj::right eccentric impulse":                               dict(mean_m=265.7782552,   sd_m=79.03054941,    mean_f=202.8498095,     sd_f=66.17376158),  # M n=29, F n=7
    "cmj::right eccentric rfd":                                   dict(mean_m=569.5640666,   sd_m=227.781265,     mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::right eccentric work":                                  dict(mean_m=147.5314935,   sd_m=37.17526783,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::right landing rfd 0-20ms":                              dict(mean_m=73766.94112,   sd_m=55528.26819,    mean_f=None,            sd_f=None),  # M n=11, F n=0
    "cmj::right landing rfd 0-40ms":                              dict(mean_m=26602.33061,   sd_m=10522.57372,    mean_f=None,            sd_f=None),  # M n=11, F n=0
    "cmj::right landing rfd 0-60ms":                              dict(mean_m=17195.4196,    sd_m=8162.861566,    mean_f=None,            sd_f=None),  # M n=11, F n=0
    "cmj::right landing rfd 0-80ms":                              dict(mean_m=9824.579386,   sd_m=3608.593974,    mean_f=None,            sd_f=None),  # M n=11, F n=0
    "cmj::right net impulse":                                     dict(mean_m=79.1658889,    sd_m=25.31515381,    mean_f=64.57690476,     sd_f=16.57612326),  # M n=29, F n=7
    "cmj::right p1 impulse":                                      dict(mean_m=111.599897,    sd_m=19.52652595,    mean_f=69.02939286,     sd_f=15.45017743),  # M n=29, F n=7
    "cmj::right p2 impulse":                                      dict(mean_m=113.3759189,   sd_m=17.45723277,    mean_f=88.80921429,     sd_f=13.45170714),  # M n=29, F n=7
    "cmj::right peak braking force":                              dict(mean_m=698.613356,    sd_m=93.88160848,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::right peak braking power":                              dict(mean_m=-543.2581634,  sd_m=118.0127841,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::right peak eccentric force":                            dict(mean_m=706.4057369,   sd_m=84.89156276,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::right peak eccentric power":                            dict(mean_m=543.2581634,   sd_m=118.0127841,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::right peak force":                                      dict(mean_m=747.123974,    sd_m=79.79842385,    mean_f=569.2147619,     sd_f=68.48561869),  # M n=29, F n=7
    "cmj::right peak landing force":                              dict(mean_m=2260.323364,   sd_m=721.8077427,    mean_f=1192.170476,     sd_f=351.5441988),  # M n=29, F n=7
    "cmj::right peak propulsive force":                           dict(mean_m=731.8705714,   sd_m=75.53969056,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::right peak propulsive power":                           dict(mean_m=1381.340172,   sd_m=177.8663145,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::right propulsive impulse":                              dict(mean_m=224.9757763,   sd_m=28.1463867,     mean_f=157.838631,      sd_f=27.90587316),  # M n=29, F n=7
    "cmj::right propulsive rfd":                                  dict(mean_m=281.1045619,   sd_m=245.2368558,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::right propulsive work":                                 dict(mean_m=254.3040776,   sd_m=40.05653251,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::right time to peak braking force":                      dict(mean_m=0.4072444643,  sd_m=0.0812577874,   mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::right time to peak braking power":                      dict(mean_m=0.2203660714,  sd_m=0.03596834058,  mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::right time to peak eccentric force":                    dict(mean_m=0.5649282738,  sd_m=0.1140679045,   mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::right time to peak eccentric power":                    dict(mean_m=0.5024555952,  sd_m=0.2339271316,   mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::right time to peak force":                              dict(mean_m=0.8615480517,  sd_m=0.3187569257,   mean_f=0.8048982143,    sd_f=0.1998165974),  # M n=29, F n=7
    "cmj::right time to peak landing force":                      dict(mean_m=0.06422751149, sd_m=0.02579041539,  mean_f=0.0813747619,    sd_f=0.01166980298),  # M n=29, F n=7
    "cmj::right time to peak power":                              dict(mean_m=0.980109881,   sd_m=0.3002808468,   mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::right time to peak propulsive force":                   dict(mean_m=0.1231202381,  sd_m=0.09683243001,  mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::right time to peak propulsive power":                   dict(mean_m=0.2832255952,  sd_m=0.04762405669,  mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::rsi":                                                   dict(mean_m=0.4587134001,  sd_m=0.1154783678,   mean_f=0.4024614286,    sd_f=0.07388949977),  # M n=29, F n=7
    "cmj::takeoff time":                                          dict(mean_m=1.615405077,   sd_m=0.2371859245,   mean_f=1.505593048,     sd_f=0.1570323719),  # M n=29, F n=7
    "cmj::takeoff velocity":                                      dict(mean_m=2.164460402,   sd_m=0.2187432916,   mean_f=1.838122381,     sd_f=0.2473622151),  # M n=29, F n=7
    "cmj::time to peak braking force":                            dict(mean_m=0.4091054167,  sd_m=0.08042203703,  mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::time to peak braking force sym. index":                 dict(mean_m=0.3950759926,  sd_m=1.40065473,     mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::time to peak braking power":                            dict(mean_m=0.2165948214,  sd_m=0.02609744894,  mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::time to peak braking power sym. index":                 dict(mean_m=-0.1829410079, sd_m=8.626403681,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::time to peak eccentric force":                          dict(mean_m=0.6911927976,  sd_m=0.2851664213,   mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::time to peak eccentric force sym. index":               dict(mean_m=4.996714951,   sd_m=12.52392211,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::time to peak eccentric power":                          dict(mean_m=0.4986822024,  sd_m=0.2569363944,   mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::time to peak eccentric power sym. index":               dict(mean_m=2.688453471,   sd_m=2.911707697,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::time to peak force":                                    dict(mean_m=0.8731181667,  sd_m=0.3271564761,   mean_f=0.8156992857,    sd_f=0.2119515243),  # M n=29, F n=7
    "cmj::time to peak landing force":                            dict(mean_m=0.06410156609, sd_m=0.02493723118,  mean_f=0.08633416667,   sd_f=0.01484382506),  # M n=29, F n=7
    "cmj::time to peak landing force sym. index":                 dict(mean_m=2.147433933,   sd_m=7.655545864,    mean_f=3.262166667,     sd_f=8.5870422),  # M n=29, F n=7
    "cmj::time to peak power":                                    dict(mean_m=1.008751739,   sd_m=0.2538497865,   mean_f=0.9358632143,    sd_f=0.1555869409),  # M n=29, F n=7
    "cmj::time to peak power sym. index":                         dict(mean_m=-0.07661756588, sd_m=0.2126857256,   mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::time to peak propulsive force":                         dict(mean_m=0.1207227976,  sd_m=0.1255060838,   mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::time to peak propulsive force sym. index":              dict(mean_m=-6.373193234,  sd_m=48.32726866,    mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::time to peak propulsive power":                         dict(mean_m=0.2819382738,  sd_m=0.04808031112,  mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::time to peak propulsive power sym. index":              dict(mean_m=-0.2852290783, sd_m=0.6530385507,   mean_f=None,            sd_f=None),  # M n=7, F n=0
    "cmj::time to peak si":                                       dict(mean_m=0.7204079334,  sd_m=4.459644435,    mean_f=-0.6537857143,   sd_f=5.27757631),  # M n=29, F n=7
    "cmj::time to takeoff":                                       dict(mean_m=1.080237668,   sd_m=0.2548163801,   mean_f=1.011172262,     sd_f=0.1552416033),  # M n=29, F n=7
    "cmj::unweighted duration":                                   dict(mean_m=0.1977968429,  sd_m=0.1263457414,   mean_f=0.222444881,     sd_f=0.0886230704),  # M n=29, F n=7
    "cmj::velocity peak power":                                   dict(mean_m=2.116554343,   sd_m=0.1862647212,   mean_f=1.858947262,     sd_f=0.2066470676),  # M n=29, F n=7
    "cmrj::avg braking force sym. index":                         dict(mean_m=-0.3088876575, sd_m=3.592502577,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::avg braking power sym. index":                         dict(mean_m=-0.07757401762, sd_m=4.092558211,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::avg eccentric force":                                  dict(mean_m=745.2024462,   sd_m=61.67684596,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::avg eccentric force sym. index":                       dict(mean_m=-1.281749032,  sd_m=2.927714875,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::avg eccentric power":                                  dict(mean_m=424.3482687,   sd_m=155.7213472,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::avg eccentric power sym. index":                       dict(mean_m=4.013452489,   sd_m=1.458027551,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::avg eccentric velocity":                               dict(mean_m=-0.5603169845, sd_m=0.2017579139,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::avg landing rfd":                                      dict(mean_m=46654.36282,   sd_m=15068.37643,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::avg landing rfd sym. index":                           dict(mean_m=-7.830044046,  sd_m=32.80982327,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::avg propulsive force sym. index":                      dict(mean_m=0.03461328643, sd_m=3.101503743,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::avg propulsive power sym. index":                      dict(mean_m=0.02058444946, sd_m=3.54980517,     mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::avg. braking force":                                   dict(mean_m=828.408005,    sd_m=62.74533796,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::avg. braking power":                                   dict(mean_m=-565.1980275,  sd_m=222.4162261,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::avg. braking velocity":                                dict(mean_m=-0.7273061136, sd_m=0.2715899397,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::avg. propulsive force":                                dict(mean_m=1191.803085,   sd_m=79.94723483,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::avg. propulsive power":                                dict(mean_m=1352.269387,   sd_m=162.5897949,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::avg. propulsive velocity":                             dict(mean_m=1.224920086,   sd_m=0.1042580011,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::avg. rfd":                                             dict(mean_m=1422.53845,    sd_m=510.270933,     mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::avg. rfd sym. index":                                  dict(mean_m=8.01165469,    sd_m=9.642430065,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::body mass":                                            dict(mean_m=75.98490433,   sd_m=6.293856142,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::body weight":                                          dict(mean_m=745.1571065,   sd_m=61.72240862,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::body weight sd":                                       dict(mean_m=7.466551327,   sd_m=0.7568757819,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::braking duration":                                     dict(mean_m=1.512789097,   sd_m=2.693782936,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::braking end time":                                     dict(mean_m=2.255371069,   sd_m=2.754702291,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::braking impulse":                                      dict(mean_m=41.85660976,   sd_m=13.64845903,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::braking impulse sym. index":                           dict(mean_m=-0.308346674,  sd_m=3.594853303,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::braking rfd":                                          dict(mean_m=2684.406741,   sd_m=1160.481336,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::braking rfd sym. index":                               dict(mean_m=4.337816333,   sd_m=8.457715527,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::braking work":                                         dict(mean_m=-244.8651879,  sd_m=81.33807114,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::braking work sym. index":                              dict(mean_m=-0.07757401762, sd_m=4.092558211,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::contact time":                                         dict(mean_m=2.067635648,   sd_m=2.838635624,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::decel rfd":                                            dict(mean_m=2824.931551,   sd_m=1538.256151,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::decel rfd sym. index":                                 dict(mean_m=-8.432914167,  sd_m=24.51134879,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::displacement depth":                                   dict(mean_m=42.61999224,   sd_m=15.50569679,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::eccentric impulse":                                    dict(mean_m=0.3153287789,  sd_m=0.1074812478,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::eccentric impulse sym. index":                         dict(mean_m=2.982993124,   sd_m=0.5653186637,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::eccentric rfd":                                        dict(mean_m=1471.424949,   sd_m=489.4998227,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::eccentric rfd sym. index":                             dict(mean_m=5.649901547,   sd_m=11.29597507,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::eccentric work":                                       dict(mean_m=319.4705241,   sd_m=122.9976393,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::eccentric work sym. index":                            dict(mean_m=4.013452489,   sd_m=1.458027551,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::flight threshold":                                     dict(mean_m=4.21625919,    sd_m=0.4831380282,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::flight time":                                          dict(mean_m=0.4323059491,  sd_m=0.02545299451,  mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::force at min displacement":                            dict(mean_m=1416.604223,   sd_m=149.3786123,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::force peak power":                                     dict(mean_m=1359.411868,   sd_m=160.8042825,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::initiation threshold":                                 dict(mean_m=725.1571065,   sd_m=61.72240862,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::jump height ft":                                       dict(mean_m=22.98643503,   sd_m=2.693924456,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::jump height ni":                                       dict(mean_m=19.99513985,   sd_m=3.240849513,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::jump momentum":                                        dict(mean_m=149.8019217,   sd_m=18.30803777,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::jump start time":                                      dict(mean_m=0.5264740787,  sd_m=0.09705668695,  mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::jump threshold time":                                  dict(mean_m=0.5264740787,  sd_m=0.09705668695,  mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::landing peak force time":                              dict(mean_m=3.091936463,   sd_m=2.761826721,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::landing rfd 0-20ms":                                   dict(mean_m=55486.41809,   sd_m=20434.33849,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::landing rfd 0-20ms sym. index":                        dict(mean_m=15.49846063,   sd_m=29.54773695,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::landing rfd 0-40ms":                                   dict(mean_m=61644.1529,    sd_m=34743.20936,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::landing rfd 0-40ms sym. index":                        dict(mean_m=3.528711974,   sd_m=12.13115275,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::landing rfd 0-60ms":                                   dict(mean_m=60910.56241,   sd_m=15153.97766,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::landing rfd 0-60ms sym. index":                        dict(mean_m=-6.62241504,   sd_m=15.93231695,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::landing rfd 0-80ms":                                   dict(mean_m=41867.89466,   sd_m=9963.972234,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::landing rfd 0-80ms sym. index":                        dict(mean_m=0.7284973407,  sd_m=4.614949734,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::landing time":                                         dict(mean_m=3.026392481,   sd_m=2.758971518,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left avg braking force":                               dict(mean_m=413.0823293,   sd_m=37.91177029,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left avg braking power":                               dict(mean_m=-281.4016782,  sd_m=112.9486085,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left avg eccentric force":                             dict(mean_m=367.809926,    sd_m=31.50146763,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left avg eccentric power":                             dict(mean_m=210.5300296,   sd_m=78.42943177,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left avg landing rfd":                                 dict(mean_m=17997.75212,   sd_m=14616.73613,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left avg propulsive force":                            dict(mean_m=595.4456481,   sd_m=30.99294113,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left avg propulsive power":                            dict(mean_m=675.2781111,   sd_m=69.34621266,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left avg rfd":                                         dict(mean_m=813.7434342,   sd_m=377.4107004,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left braking impulse":                                 dict(mean_m=592.7732675,   sd_m=1029.677336,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left braking rfd":                                     dict(mean_m=1731.833603,   sd_m=523.5533047,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left braking work":                                    dict(mean_m=-124.7550571,  sd_m=37.92361462,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left decel rfd":                                       dict(mean_m=1423.374912,   sd_m=822.3545027,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left eccentric impulse":                               dict(mean_m=648.9643228,   sd_m=1071.759565,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left eccentric rfd":                                   dict(mean_m=830.5184566,   sd_m=367.871118,     mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left eccentric work":                                  dict(mean_m=160.0163415,   sd_m=64.92929748,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left landing rfd 0-20ms":                              dict(mean_m=30340.2595,    sd_m=12926.90703,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left landing rfd 0-40ms":                              dict(mean_m=31301.73818,   sd_m=15704.1346,     mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left landing rfd 0-60ms":                              dict(mean_m=28447.69997,   sd_m=9553.703323,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left landing rfd 0-80ms":                              dict(mean_m=21271.9791,    sd_m=5970.677361,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left net impulse":                                     dict(mean_m=104.9158434,   sd_m=65.98433353,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left p1 impulse":                                      dict(mean_m=110.5855862,   sd_m=16.75125216,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left p2 impulse":                                      dict(mean_m=91.92704832,   sd_m=19.80874424,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left peak braking force":                              dict(mean_m=1048.101792,   sd_m=787.0112833,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left peak braking power":                              dict(mean_m=-1148.891276,  sd_m=1440.558702,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left peak eccentric force":                            dict(mean_m=1048.101792,   sd_m=787.0112833,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left peak eccentric power":                            dict(mean_m=1148.891276,   sd_m=1440.558702,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left peak force":                                      dict(mean_m=1080.685941,   sd_m=780.7973544,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left peak landing force":                              dict(mean_m=2654.593554,   sd_m=528.7005598,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left peak propulsive force":                           dict(mean_m=746.5212684,   sd_m=49.48664269,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left peak propulsive power":                           dict(mean_m=1299.878041,   sd_m=190.9595941,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left propulsive impulse":                              dict(mean_m=202.5126484,   sd_m=35.32748091,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left propulsive rfd":                                  dict(mean_m=265.8264671,   sd_m=233.7238466,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left propulsive work":                                 dict(mean_m=230.2508386,   sd_m=46.10162418,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left rebound avg braking force":                       dict(mean_m=1310.858708,   sd_m=175.8684588,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::left rebound avg braking power":                       dict(mean_m=-1419.964293,  sd_m=48.50127382,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::left rebound avg propulsive force":                    dict(mean_m=1127.810469,   sd_m=284.0832616,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::left rebound avg propulsive power":                    dict(mean_m=1455.658224,   sd_m=483.648889,     mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::left rebound braking impulse":                         dict(mean_m=109.6826604,   sd_m=14.26824229,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::left rebound braking rfd":                             dict(mean_m=39256.93596,   sd_m=16342.40656,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::left rebound braking work":                            dict(mean_m=-121.7572288,  sd_m=29.53398492,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::left rebound p1 impulse":                              dict(mean_m=101.7503196,   sd_m=6.305373358,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::left rebound p2 impulse":                              dict(mean_m=33.21146755,   sd_m=6.772156751,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::left rebound peak braking force":                      dict(mean_m=2246.992256,   sd_m=450.2198281,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::left rebound peak braking power":                      dict(mean_m=-2766.270867,  sd_m=342.2030437,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::left rebound peak propulsive force":                   dict(mean_m=1959.797673,   sd_m=655.1282569,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::left rebound peak propulsive power":                   dict(mean_m=2739.59415,    sd_m=1069.601712,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::left rebound propulsive impulse":                      dict(mean_m=134.9617872,   sd_m=5.976759291,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::left rebound propulsive rfd":                          dict(mean_m=4485.66828,    sd_m=4467.586479,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::left rebound propulsive work":                         dict(mean_m=175.6988231,   sd_m=11.95566353,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::left rebound time to peak braking force":              dict(mean_m=0.06195023148, sd_m=0.01722698479,  mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::left rebound time to peak braking power":              dict(mean_m=0.04357638889, sd_m=0.01057603779,  mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::left rebound time to peak propulsive force":           dict(mean_m=0.01212384259, sd_m=0.005091523851, mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::left rebound time to peak propulsive power":           dict(mean_m=0.06021412037, sd_m=0.0138563892,   mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::left time to peak braking force":                      dict(mean_m=0.4958623611,  sd_m=0.2476011052,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left time to peak braking power":                      dict(mean_m=0.3628044444,  sd_m=0.2820358534,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left time to peak eccentric force":                    dict(mean_m=0.7119065046,  sd_m=0.376510135,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left time to peak eccentric power":                    dict(mean_m=0.5788481713,  sd_m=0.4144303895,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left time to peak force":                              dict(mean_m=0.8080000463,  sd_m=0.3876352703,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left time to peak landing force":                      dict(mean_m=0.4018518056,  sd_m=0.199285085,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left time to peak power":                              dict(mean_m=1.347467315,   sd_m=1.253649562,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left time to peak propulsive force":                   dict(mean_m=0.1323500926,  sd_m=0.1121526319,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::left time to peak propulsive power":                   dict(mean_m=0.2665501852,  sd_m=0.04591363757,  mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::min braking velocity":                                 dict(mean_m=-1.285192629,  sd_m=0.3013727812,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::min eccentric velocity":                               dict(mean_m=-1.285193601,  sd_m=0.3013840202,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::min unweight force":                                   dict(mean_m=263.1718471,   sd_m=59.56411356,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::net impulse":                                          dict(mean_m=149.8019772,   sd_m=18.30803479,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::net impulse sym. index":                               dict(mean_m=3.145018442,   sd_m=19.23381938,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::p1 avg force":                                         dict(mean_m=1306.137216,   sd_m=95.90401294,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::p1 avg power":                                         dict(mean_m=840.9273079,   sd_m=158.7796528,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::p1 avg velocity":                                      dict(mean_m=0.657416984,   sd_m=0.093581984,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::p1 duration":                                          dict(mean_m=0.1699083102,  sd_m=0.02467463281,  mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::p1 impulse":                                           dict(mean_m=93.57705347,   sd_m=12.14353699,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::p1 impulse sym. index":                                dict(mean_m=0.02109588385, sd_m=2.961833887,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::p1 p2 duration ratio":                                 dict(mean_m=1.006206652,   sd_m=0.002601308649, mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::p1 p2 force ratio":                                    dict(mean_m=1.227270761,   sd_m=0.1840590859,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::p1 p2 power ratio":                                    dict(mean_m=0.4548450736,  sd_m=0.1109585427,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::p1 p2 velocity ratio":                                 dict(mean_m=0.3647725836,  sd_m=0.04022756272,  mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::p1 peak force":                                        dict(mean_m=1440.234851,   sd_m=143.8799831,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::p1 peak power":                                        dict(mean_m=1543.518464,   sd_m=250.9712008,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::p1 peak velocity":                                     dict(mean_m=1.230887001,   sd_m=0.1631582026,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::p2 avg force":                                         dict(mean_m=1077.139593,   sd_m=119.4971591,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::p2 avg power":                                         dict(mean_m=1868.087144,   sd_m=258.4187019,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::p2 avg velocity":                                      dict(mean_m=1.796055585,   sd_m=0.1376350536,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::p2 duration":                                          dict(mean_m=0.1688945139,  sd_m=0.02489243367,  mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::p2 impulse":                                           dict(mean_m=56.54007199,   sd_m=14.82689349,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::p2 impulse sym. index":                                dict(mean_m=0.06648964846, sd_m=3.602888156,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::p2 peak force":                                        dict(mean_m=1393.852775,   sd_m=132.4000611,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::p2 peak power":                                        dict(mean_m=2612.754626,   sd_m=384.3160936,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::p2 peak velocity":                                     dict(mean_m=2.153177166,   sd_m=0.1534029792,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::peak braking force":                                   dict(mean_m=1848.252701,   sd_m=974.8344228,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::peak braking force sym. index":                        dict(mean_m=1.617360758,   sd_m=5.246393015,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::peak braking power":                                   dict(mean_m=-1815.997038,  sd_m=1744.787877,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::peak braking power sym. index":                        dict(mean_m=3.319144898,   sd_m=7.648195329,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::peak eccentric force":                                 dict(mean_m=1848.252701,   sd_m=974.8344228,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::peak eccentric force sym. index":                      dict(mean_m=1.617360758,   sd_m=5.246393015,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::peak eccentric power":                                 dict(mean_m=1815.997038,   sd_m=1744.787877,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::peak eccentric power sym. index":                      dict(mean_m=6.708693761,   sd_m=5.195771045,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::peak force":                                           dict(mean_m=1905.988382,   sd_m=954.0549824,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::peak force sym. index":                                dict(mean_m=1.979123419,   sd_m=5.128328473,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::peak force time":                                      dict(mean_m=1.468198037,   sd_m=0.4781538224,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::peak landing force":                                   dict(mean_m=4673.736133,   sd_m=707.5679653,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::peak landing force sym. index":                        dict(mean_m=0.824024308,   sd_m=6.041559757,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::peak power":                                           dict(mean_m=2640.148376,   sd_m=390.7757078,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::peak propulsive force":                                dict(mean_m=1489.77021,    sd_m=114.0745574,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::peak propulsive force sym. index":                     dict(mean_m=0.1788659621,  sd_m=2.40485256,     mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::peak propulsive power":                                dict(mean_m=2612.754626,   sd_m=384.3160936,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::peak propulsive power sym. index":                     dict(mean_m=-0.69975264,   sd_m=3.087773021,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::peak propulsive velocity":                             dict(mean_m=2.153179249,   sd_m=0.1534052335,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::peak velocity":                                        dict(mean_m=2.160720499,   sd_m=0.1564670577,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::propulsive duration":                                  dict(mean_m=0.3388024074,  sd_m=0.04956628992,  mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::propulsive impulse":                                   dict(mean_m=150.1172088,   sd_m=18.26653603,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::propulsive impulse sym. index":                        dict(mean_m=0.03517169658, sd_m=3.102075871,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::propulsive rfd":                                       dict(mean_m=438.0649602,   sd_m=298.7413223,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::propulsive rfd sym. index":                            dict(mean_m=-7.035622744,  sd_m=39.24234508,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::propulsive work":                                      dict(mean_m=459.8909142,   sd_m=91.03058,       mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::propulsive work sym. index":                           dict(mean_m=0.02058444946, sd_m=3.54980517,     mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound avg braking force":                            dict(mean_m=2370.054493,   sd_m=473.9753078,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound avg braking force sym. index":                 dict(mean_m=1.519861111,   sd_m=11.88981246,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::rebound avg braking force symmetry index":             dict(mean_m=-1.026325777,  sd_m=7.364919162,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::rebound avg braking power":                            dict(mean_m=-2592.131282,  sd_m=446.9131447,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound avg braking power sym. index":                 dict(mean_m=3.825388889,   sd_m=15.47802655,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::rebound avg braking power symmetry index":             dict(mean_m=-0.7698779897, sd_m=8.94176023,     mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::rebound avg braking velocity":                         dict(mean_m=-1.369896032,  sd_m=0.08829804203,  mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound avg propulsive force":                         dict(mean_m=1920.633016,   sd_m=455.7771747,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound avg propulsive force sym. index":              dict(mean_m=0.8807222222,  sd_m=3.875223658,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::rebound avg propulsive force symmetry index":          dict(mean_m=0.6286094571,  sd_m=3.910505213,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::rebound avg propulsive power":                         dict(mean_m=2311.052898,   sd_m=866.155871,     mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound avg propulsive power sym. index":              dict(mean_m=1.544,         sd_m=3.353302402,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::rebound avg propulsive power symmetry index":          dict(mean_m=-0.2448032573, sd_m=3.927747215,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::rebound avg propulsive velocity":                      dict(mean_m=1.485574495,   sd_m=0.2168124989,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound braking impulse":                              dict(mean_m=149.7432972,   sd_m=18.47228396,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound braking impulse sym. index":                   dict(mean_m=1.523916667,   sd_m=11.94148496,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::rebound braking impulse symmetry index":               dict(mean_m=-1.022768867,  sd_m=7.394159999,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::rebound braking rfd":                                  dict(mean_m=72793.55796,   sd_m=25049.65171,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound braking rfd sym. index":                       dict(mean_m=-2.94125,      sd_m=22.32402322,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::rebound braking rfd symmetry index":                   dict(mean_m=-0.9772253298, sd_m=3.247808523,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::rebound braking work":                                 dict(mean_m=-246.4157803,  sd_m=51.08155526,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound braking work sym. index":                      dict(mean_m=3.825388889,   sd_m=15.47802655,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::rebound braking work symmetry index":                  dict(mean_m=-0.7698779897, sd_m=8.94176023,     mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::rebound flight time":                                  dict(mean_m=0.3810758333,  sd_m=0.059274043,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound force at min displacement":                    dict(mean_m=3302.626473,   sd_m=896.6719619,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound jump momentum":                                dict(mean_m=147.2085496,   sd_m=29.26979238,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound min braking velocity":                         dict(mean_m=-2.047641702,  sd_m=0.1585925736,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound p1 avg force":                                 dict(mean_m=2898.046927,   sd_m=751.2625072,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound p1 avg power":                                 dict(mean_m=2710.814847,   sd_m=1139.222515,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound p1 avg velocity":                              dict(mean_m=0.9624238764,  sd_m=0.1377001754,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound p1 duration":                                  dict(mean_m=0.06547997685, sd_m=0.01101798185,  mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound p1 impulse":                                   dict(mean_m=135.166529,    sd_m=20.53237563,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound p1 impulse sym. index":                        dict(mean_m=0.0145,        sd_m=4.279990975,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::rebound p1 impulse symmetry index":                    dict(mean_m=1.923556612,   sd_m=4.379299738,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::rebound p1 p2 duration ratio":                         dict(mean_m=1.010346302,   sd_m=0.006122078764, mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound p1 p2 force ratio":                            dict(mean_m=3.151494578,   sd_m=0.4923081955,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound p1 p2 power ratio":                            dict(mean_m=1.451602583,   sd_m=0.2574185303,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound p1 p2 velocity ratio":                         dict(mean_m=0.4810399715,  sd_m=0.03032300446,  mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound p1 peak force":                                dict(mean_m=3377.103654,   sd_m=886.1412531,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound p1 peak power":                                dict(mean_m=4180.602865,   sd_m=1733.050685,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound p1 peak velocity":                             dict(mean_m=1.780415303,   sd_m=0.2530199689,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound p2 avg force":                                 dict(mean_m=938.6890889,   sd_m=203.8408639,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound p2 avg power":                                 dict(mean_m=1938.226073,   sd_m=675.640731,     mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound p2 avg velocity":                              dict(mean_m=2.018842221,   sd_m=0.3076368415,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound p2 duration":                                  dict(mean_m=0.06481490741, sd_m=0.01093072323,  mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound p2 impulse":                                   dict(mean_m=12.1006173,    sd_m=9.700630567,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound p2 impulse sym. index":                        dict(mean_m=4.056055556,   sd_m=2.331789193,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::rebound p2 impulse symmetry index":                    dict(mean_m=-3.87998528,   sd_m=3.487900287,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::rebound p2 peak force":                                dict(mean_m=2210.013701,   sd_m=611.459314,     mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound p2 peak power":                                dict(mean_m=4099.331923,   sd_m=1775.40652,     mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound p2 peak velocity":                             dict(mean_m=2.11475195,    sd_m=0.3149991393,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound peak braking force":                           dict(mean_m=4395.990671,   sd_m=665.2804423,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound peak braking force sym. index":                dict(mean_m=-0.8921666667, sd_m=13.2299212,     mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::rebound peak braking force symmetry index":            dict(mean_m=-1.617184007,  sd_m=5.313808935,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::rebound peak braking power":                           dict(mean_m=-5631.105459,  sd_m=920.5732659,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound peak braking power sym. index":                dict(mean_m=1.417555556,   sd_m=24.65876549,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::rebound peak braking power symmetry index":            dict(mean_m=-1.37836184,   sd_m=8.171100596,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::rebound peak propulsive force":                        dict(mean_m=3377.103654,   sd_m=886.1412531,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound peak propulsive force sym. index":             dict(mean_m=-0.02286111111, sd_m=5.105193085,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::rebound peak propulsive force symmetry index":         dict(mean_m=0.3852528581,  sd_m=4.322273639,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::rebound peak propulsive power":                        dict(mean_m=4201.221568,   sd_m=1762.427248,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound peak propulsive power sym. index":             dict(mean_m=0.08588888889, sd_m=3.272947811,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::rebound peak propulsive power symmetry index":         dict(mean_m=0.781054956,   sd_m=4.939031295,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::rebound peak propulsive velocity":                     dict(mean_m=2.114751672,   sd_m=0.3149971334,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound propulsive impulse sym. index":                dict(mean_m=0.8853333333,  sd_m=3.869679132,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::rebound propulsive impulse symmetry index":            dict(mean_m=0.6392929777,  sd_m=3.921175734,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::rebound propulsive rfd":                               dict(mean_m=7920.649788,   sd_m=7162.316139,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound propulsive rfd sym. index":                    dict(mean_m=-13.61311111,  sd_m=34.12883968,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::rebound propulsive rfd symmetry index":                dict(mean_m=22.66401569,   sd_m=11.70821061,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::rebound propulsive work":                              dict(mean_m=293.6717574,   sd_m=68.77765404,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound propulsive work sym. index":                   dict(mean_m=1.544,         sd_m=3.353302402,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::rebound propulsive work symmetry index":               dict(mean_m=-0.2448032573, sd_m=3.927747215,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::rebound relative force at min displacement":           dict(mean_m=4.446161081,   sd_m=1.223957643,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound rsi":                                          dict(mean_m=1.763856333,   sd_m=0.5748974904,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound time to peak braking force":                   dict(mean_m=0.06474222222, sd_m=0.01166905226,  mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound time to peak braking force sym. index":        dict(mean_m=2.376416667,   sd_m=10.11140768,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::rebound time to peak braking force symmetry index":    dict(mean_m=-0.6417016035, sd_m=2.592582952,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::rebound time to peak braking power":                   dict(mean_m=0.05044824074, sd_m=0.01088793097,  mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound time to peak braking power sym. index":        dict(mean_m=1.197277778,   sd_m=16.25835877,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::rebound time to peak braking power symmetry index":    dict(mean_m=4.231369348,   sd_m=3.684171352,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::rebound time to peak propulsive force":                dict(mean_m=0.006640162037, sd_m=0.005371342681, mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound time to peak propulsive force sym. index":     dict(mean_m=-13.61619444,  sd_m=17.96297738,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::rebound time to peak propulsive force symmetry index": dict(mean_m=25.23809524,   sd_m=7.797189899,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::rebound time to peak propulsive power":                dict(mean_m=0.06051837963, sd_m=0.01286216128,  mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rebound time to peak propulsive power sym. index":     dict(mean_m=1.483027778,   sd_m=1.841713956,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::rebound time to peak propulsive power symmetry index": dict(mean_m=-2.45332181,   sd_m=1.249614062,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::rebound time to takeoff":                              dict(mean_m=0.2273584028,  sd_m=0.04543635166,  mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rel. avg. propulsive force":                           dict(mean_m=15.72244346,   sd_m=0.9210749415,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rel. min unweight force":                              dict(mean_m=0.3528211431,  sd_m=0.07594130792,  mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rel. propulsive impulse":                              dict(mean_m=1.974121576,   sd_m=0.1596148872,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::relative force at min displacement":                   dict(mean_m=1.907142944,   sd_m=0.2112626581,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::relative peak force":                                  dict(mean_m=2.565139865,   sd_m=1.28302465,     mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::relative peak landing force":                          dict(mean_m=6.261720101,   sd_m=0.6896387198,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::relative peak power":                                  dict(mean_m=34.65790787,   sd_m=3.556905258,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right avg braking force":                              dict(mean_m=415.3248423,   sd_m=31.08320384,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right avg braking power":                              dict(mean_m=-283.7959326,  sd_m=110.7130694,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right avg eccentric force":                            dict(mean_m=377.3925202,   sd_m=33.82319051,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right avg eccentric power":                            dict(mean_m=213.8176835,   sd_m=78.09577377,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right avg landing rfd":                                dict(mean_m=20344.22064,   sd_m=6371.738834,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right avg propulsive force":                           dict(mean_m=596.356881,    sd_m=54.23683163,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right avg propulsive power":                           dict(mean_m=676.9911372,   sd_m=98.21337801,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right avg rfd":                                        dict(mean_m=669.5096999,   sd_m=280.4044578,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right braking impulse":                                dict(mean_m=575.962928,    sd_m=983.7828199,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right braking rfd":                                    dict(mean_m=1562.146337,   sd_m=436.3974703,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right braking work":                                   dict(mean_m=-120.1097141,  sd_m=44.76622825,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right decel rfd":                                      dict(mean_m=1401.557055,   sd_m=723.4389477,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right eccentric impulse":                              dict(mean_m=638.4208102,   sd_m=1027.072291,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right eccentric rfd":                                  dict(mean_m=709.7919119,   sd_m=254.9171754,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right eccentric work":                                 dict(mean_m=159.4555715,   sd_m=58.32334803,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right landing rfd 0-20ms":                             dict(mean_m=25146.16025,   sd_m=11056.92657,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right landing rfd 0-40ms":                             dict(mean_m=30342.41431,   sd_m=19658.12227,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right landing rfd 0-60ms":                             dict(mean_m=32462.86285,   sd_m=7685.401314,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right landing rfd 0-80ms":                             dict(mean_m=20595.91681,   sd_m=4151.318114,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right net impulse":                                    dict(mean_m=92.59530046,   sd_m=56.37265956,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right p1 impulse":                                     dict(mean_m=110.5092883,   sd_m=16.84602911,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right p2 impulse":                                     dict(mean_m=91.38092986,   sd_m=18.38127427,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right peak braking force":                             dict(mean_m=929.3971287,   sd_m=498.7734556,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right peak braking power":                             dict(mean_m=-853.4845434,  sd_m=751.517089,     mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right peak eccentric force":                           dict(mean_m=929.3971287,   sd_m=498.7734556,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right peak eccentric power":                           dict(mean_m=853.4845434,   sd_m=751.517089,     mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right peak force":                                     dict(mean_m=957.9207995,   sd_m=483.6111449,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right peak landing force":                             dict(mean_m=2564.999906,   sd_m=391.6927293,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right peak propulsive force":                          dict(mean_m=745.0732495,   sd_m=68.67518736,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right peak propulsive power":                          dict(mean_m=1317.108759,   sd_m=203.1686154,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right propulsive impulse":                             dict(mean_m=201.8902042,   sd_m=33.67681232,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right propulsive rfd":                                 dict(mean_m=261.2807556,   sd_m=106.8238315,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right propulsive work":                                dict(mean_m=229.6388255,   sd_m=46.61326393,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right rebound avg braking force":                      dict(mean_m=1350.486112,   sd_m=279.2936515,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::right rebound avg braking power":                      dict(mean_m=-1445.491049,  sd_m=301.2025086,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::right rebound avg propulsive force":                   dict(mean_m=1100.331118,   sd_m=198.2937301,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::right rebound avg propulsive power":                   dict(mean_m=1445.770628,   sd_m=428.7725498,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::right rebound braking impulse":                        dict(mean_m=112.2750626,   sd_m=25.40224337,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::right rebound braking rfd":                            dict(mean_m=41334.52127,   sd_m=19088.13035,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::right rebound braking work":                           dict(mean_m=-123.5121095,  sd_m=42.54978248,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::right rebound p1 impulse":                             dict(mean_m=98.10346702,   sd_m=2.110477428,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::right rebound p2 impulse":                             dict(mean_m=35.27803024,   sd_m=6.088792913,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::right rebound peak braking force":                     dict(mean_m=2326.17088,    sd_m=464.8873591,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::right rebound peak braking power":                     dict(mean_m=-2889.955578,  sd_m=853.1092751,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::right rebound peak propulsive force":                  dict(mean_m=1915.582068,   sd_m=465.1322914,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::right rebound peak propulsive power":                  dict(mean_m=2645.523399,   sd_m=836.1372719,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::right rebound propulsive impulse":                     dict(mean_m=133.3814973,   sd_m=8.105131672,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::right rebound propulsive rfd":                         dict(mean_m=2687.914389,   sd_m=2889.442479,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::right rebound propulsive work":                        dict(mean_m=175.8258029,   sd_m=7.282261169,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::right rebound time to peak braking force":             dict(mean_m=0.06276041667, sd_m=0.01693998796,  mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::right rebound time to peak braking power":             dict(mean_m=0.04056712963, sd_m=0.009004934597, mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::right rebound time to peak propulsive force":          dict(mean_m=0.004282407407, sd_m=0.00417148641,  mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::right rebound time to peak propulsive power":          dict(mean_m=0.06316550926, sd_m=0.01493063967,  mean_f=None,            sd_f=None),  # M n=3, F n=0
    "cmrj::right time to peak braking force":                     dict(mean_m=0.524609838,   sd_m=0.304936573,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right time to peak braking power":                     dict(mean_m=0.4068577315,  sd_m=0.3954306765,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right time to peak eccentric force":                   dict(mean_m=0.7406535648,  sd_m=0.4349209932,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right time to peak eccentric power":                   dict(mean_m=0.6229020139,  sd_m=0.5280614446,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right time to peak force":                             dict(mean_m=0.8618922685,  sd_m=0.4109011711,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right time to peak landing force":                     dict(mean_m=0.3418260648,  sd_m=0.1357890887,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right time to peak power":                             dict(mean_m=1.37884838,    sd_m=1.329133856,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right time to peak propulsive force":                  dict(mean_m=0.1418112963,  sd_m=0.1137559946,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::right time to peak propulsive power":                  dict(mean_m=0.2662041667,  sd_m=0.04816471114,  mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rsi":                                                  dict(mean_m=0.433235387,   sd_m=0.1389351626,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::rsi modified":                                         dict(mean_m=0.1975809749,  sd_m=0.06989192782,  mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::takeoff time":                                         dict(mean_m=2.594132644,   sd_m=2.767455488,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::takeoff velocity":                                     dict(mean_m=1.969976539,   sd_m=0.1607820564,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::time to peak braking force":                           dict(mean_m=0.4997976852,  sd_m=0.2468107092,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::time to peak braking force sym. index":                dict(mean_m=-1.349782343,  sd_m=1.992338061,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::time to peak braking power":                           dict(mean_m=0.363802037,   sd_m=0.2820740002,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::time to peak braking power sym. index":                dict(mean_m=-0.8820012037, sd_m=4.147627428,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::time to peak eccentric force":                         dict(mean_m=0.715841412,   sd_m=0.3756776924,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::time to peak eccentric force sym. index":              dict(mean_m=-1.065135605,  sd_m=1.637224978,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::time to peak eccentric power":                         dict(mean_m=0.5798467361,  sd_m=0.4145913767,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::time to peak eccentric power sym. index":              dict(mean_m=2.517655864,   sd_m=2.45913231,     mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::time to peak force":                                   dict(mean_m=0.8144671065,  sd_m=0.3851712042,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::time to peak landing force":                           dict(mean_m=0.2136136111,  sd_m=0.1756915116,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::time to peak landing force sym. index":                dict(mean_m=8.678298392,   sd_m=29.54750433,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::time to peak power":                                   dict(mean_m=1.347525255,   sd_m=1.252853791,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::time to peak power sym. index":                        dict(mean_m=-1.165092014,  sd_m=2.852776599,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::time to peak propulsive force":                        dict(mean_m=0.1356484028,  sd_m=0.1120923396,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::time to peak propulsive force sym. index":             dict(mean_m=-16.00399681,  sd_m=22.76778748,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::time to peak propulsive power":                        dict(mean_m=0.266694375,   sd_m=0.04757530731,  mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::time to peak propulsive power sym. index":             dict(mean_m=0.1278980357,  sd_m=0.9323514974,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::time to peak si":                                      dict(mean_m=-2.925316729,  sd_m=5.217038816,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::time to takeoff":                                      dict(mean_m=2.067635648,   sd_m=2.838635624,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::unweighted duration":                                  dict(mean_m=0.2160446991,  sd_m=0.1354259755,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "cmrj::velocity peak power":                                  dict(mean_m=1.949283646,   sd_m=0.1469544127,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "dj::avg braking force sym. index":                           dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::avg braking power sym. index":                           dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::avg landing rfd":                                        dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::avg landing rfd sym. index":                             dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::avg propulsive force sym. index":                        dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::avg propulsive power sym. index":                        dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::avg. braking force":                                     dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::avg. braking power":                                     dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::avg. braking velocity":                                  dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::avg. propulsive force":                                  dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::avg. propulsive power":                                  dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::avg. propulsive velocity":                               dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::body mass":                                              dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::body weight":                                            dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::braking end time":                                       dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::braking impulse":                                        dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::braking impulse sym. index":                             dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::braking rfd":                                            dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::braking rfd sym. index":                                 dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::braking work":                                           dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::braking work sym. index":                                dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::contact time":                                           dict(mean_m=0.6795453,     sd_m=0.230528902,    mean_f=0.492447,        sd_f=0.1804647816),  # M n=5, F n=5
    "dj::flight threshold":                                       dict(mean_m=4.100833333,   sd_m=0.8168552078,   mean_f=3.616333333,     sd_f=0.3790945902),  # M n=5, F n=5
    "dj::flight time":                                            dict(mean_m=0.4168710667,  sd_m=0.04747289328,  mean_f=0.3858511667,    sd_f=0.03478029051),  # M n=5, F n=5
    "dj::force peak power":                                       dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::initiation threshold":                                   dict(mean_m=10.0,          sd_m=0.0,            mean_f=10.0,            sd_f=0.0),  # M n=5, F n=5
    "dj::jump height ft":                                         dict(mean_m=21.55295667,   sd_m=4.783238661,    mean_f=18.38935,        sd_f=3.300743096),  # M n=5, F n=5
    "dj::jump momentum":                                          dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::jump start time":                                        dict(mean_m=-0.003470953333, sd_m=0.0008854626265, mean_f=-0.00506965,     sd_f=0.001001158767),  # M n=5, F n=5
    "dj::jump threshold time":                                    dict(mean_m=0.4965290467,  sd_m=0.0008854626265, mean_f=0.49493035,      sd_f=0.001001158767),  # M n=5, F n=5
    "dj::landing peak force time":                                dict(mean_m=1.66877878,    sd_m=0.2590177157,   mean_f=1.46059135,      sd_f=0.1418753962),  # M n=5, F n=5
    "dj::landing rfd 0-20ms":                                     dict(mean_m=32980.5705,    sd_m=14423.45845,    mean_f=26491.43083,     sd_f=3877.954435),  # M n=5, F n=5
    "dj::landing rfd 0-20ms sym. index":                          dict(mean_m=22.9519,       sd_m=31.32000659,    mean_f=15.63631667,     sd_f=43.7860939),  # M n=5, F n=5
    "dj::landing rfd 0-40ms":                                     dict(mean_m=29474.6932,    sd_m=3876.709158,    mean_f=23096.37867,     sd_f=4681.593669),  # M n=5, F n=5
    "dj::landing rfd 0-40ms sym. index":                          dict(mean_m=13.78975667,   sd_m=4.180841933,    mean_f=-1.5859,         sd_f=29.59948834),  # M n=5, F n=5
    "dj::landing rfd 0-60ms":                                     dict(mean_m=43046.3538,    sd_m=16221.33487,    mean_f=23504.04283,     sd_f=5360.608382),  # M n=5, F n=5
    "dj::landing rfd 0-60ms sym. index":                          dict(mean_m=5.614386667,   sd_m=6.635505133,    mean_f=-9.514566667,    sd_f=22.62569296),  # M n=5, F n=5
    "dj::landing rfd 0-80ms":                                     dict(mean_m=24586.60583,   sd_m=2726.101471,    mean_f=18600.7575,      sd_f=7021.892624),  # M n=5, F n=5
    "dj::landing rfd 0-80ms sym. index":                          dict(mean_m=9.89909,       sd_m=3.219072677,    mean_f=-6.621733333,    sd_f=23.26235617),  # M n=5, F n=5
    "dj::landing time":                                           dict(mean_m=1.594558447,   sd_m=0.2524950337,   mean_f=1.37574685,      sd_f=0.1493417895),  # M n=5, F n=5
    "dj::left avg braking force":                                 dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::left avg braking power":                                 dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::left avg landing rfd":                                   dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::left avg propulsive force":                              dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::left avg propulsive power":                              dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::left braking impulse":                                   dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::left braking rfd":                                       dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::left braking work":                                      dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::left landing rfd 0-20ms":                                dict(mean_m=19121.51943,   sd_m=6187.214836,    mean_f=14697.06733,     sd_f=5755.944092),  # M n=5, F n=5
    "dj::left landing rfd 0-40ms":                                dict(mean_m=16746.09407,   sd_m=2113.565583,    mean_f=11668.52367,     sd_f=4835.069431),  # M n=5, F n=5
    "dj::left landing rfd 0-60ms":                                dict(mean_m=22140.24493,   sd_m=7562.912806,    mean_f=10379.7255,      sd_f=3261.815377),  # M n=5, F n=5
    "dj::left landing rfd 0-80ms":                                dict(mean_m=13483.45703,   sd_m=1796.527372,    mean_f=8950.404167,     sd_f=4437.376056),  # M n=5, F n=5
    "dj::left p1 impulse":                                        dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::left p2 impulse":                                        dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::left peak braking force":                                dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::left peak braking power":                                dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::left peak landing force":                                dict(mean_m=2242.284533,   sd_m=857.6879708,    mean_f=1037.088833,     sd_f=395.7337959),  # M n=5, F n=5
    "dj::left peak propulsive force":                             dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::left peak propulsive power":                             dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::left propulsive impulse":                                dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::left propulsive rfd":                                    dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::left propulsive work":                                   dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::left time to peak braking force":                        dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::left time to peak braking power":                        dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::left time to peak landing force":                        dict(mean_m=0.0754064,     sd_m=0.01631822761,  mean_f=0.07718783333,   sd_f=0.006893734212),  # M n=5, F n=5
    "dj::left time to peak propulsive force":                     dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::left time to peak propulsive power":                     dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::min braking velocity":                                   dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::net impulse":                                            dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::p1 avg force":                                           dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::p1 avg power":                                           dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::p1 avg velocity":                                        dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::p1 duration":                                            dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::p1 impulse":                                             dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::p1 impulse sym. index":                                  dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::p1 p2 duration ratio":                                   dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::p1 p2 force ratio":                                      dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::p1 p2 power ratio":                                      dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::p1 p2 velocity ratio":                                   dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::p1 peak force":                                          dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::p1 peak power":                                          dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::p1 peak velocity":                                       dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::p2 avg force":                                           dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::p2 avg power":                                           dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::p2 avg velocity":                                        dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::p2 duration":                                            dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::p2 impulse":                                             dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::p2 impulse sym. index":                                  dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::p2 peak force":                                          dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::p2 peak power":                                          dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::p2 peak velocity":                                       dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::peak braking force":                                     dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::peak braking force sym. index":                          dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::peak braking power":                                     dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::peak braking power sym. index":                          dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::peak force":                                             dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::peak force time":                                        dict(mean_m=0.9528775467,  sd_m=0.3070480167,   mean_f=0.8099991833,    sd_f=0.1942274497),  # M n=5, F n=5
    "dj::peak landing force":                                     dict(mean_m=3819.637633,   sd_m=1117.046231,    mean_f=2023.346667,     sd_f=487.9441822),  # M n=5, F n=5
    "dj::peak landing force sym. index":                          dict(mean_m=0.4380666667,  sd_m=4.860994728,    mean_f=-7.465183333,    sd_f=19.95757423),  # M n=5, F n=5
    "dj::peak power":                                             dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::peak propulsive force":                                  dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::peak propulsive force sym. index":                       dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::peak propulsive power":                                  dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::peak propulsive power sym. index":                       dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::peak propulsive velocity":                               dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::peak velocity":                                          dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::propulsive impulse":                                     dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::propulsive impulse sym. index":                          dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::propulsive rfd":                                         dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::propulsive rfd sym. index":                              dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::propulsive work":                                        dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::propulsive work sym. index":                             dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::rel. avg. propulsive force":                             dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::rel. propulsive impulse":                                dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::relative peak force":                                    dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::relative peak landing force":                            dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::relative peak power":                                    dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::right avg braking force":                                dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::right avg braking power":                                dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::right avg landing rfd":                                  dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::right avg propulsive force":                             dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::right avg propulsive power":                             dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::right braking impulse":                                  dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::right braking rfd":                                      dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::right braking work":                                     dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::right landing rfd 0-20ms":                               dict(mean_m=13859.05007,   sd_m=10388.49505,    mean_f=11794.364,       sd_f=6333.061822),  # M n=5, F n=5
    "dj::right landing rfd 0-40ms":                               dict(mean_m=12728.59847,   sd_m=1904.882171,    mean_f=11427.854,       sd_f=2529.792172),  # M n=5, F n=5
    "dj::right landing rfd 0-60ms":                               dict(mean_m=20906.10897,   sd_m=8760.949915,    mean_f=13124.31733,     sd_f=4264.595737),  # M n=5, F n=5
    "dj::right landing rfd 0-80ms":                               dict(mean_m=11103.14873,   sd_m=1121.993537,    mean_f=9650.353,        sd_f=3280.045945),  # M n=5, F n=5
    "dj::right p1 impulse":                                       dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::right p2 impulse":                                       dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::right peak braking force":                               dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::right peak braking power":                               dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::right peak landing force":                               dict(mean_m=2141.633333,   sd_m=731.7270853,    mean_f=1198.810333,     sd_f=225.8852822),  # M n=5, F n=5
    "dj::right peak propulsive force":                            dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::right peak propulsive power":                            dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::right propulsive impulse":                               dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::right propulsive rfd":                                   dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::right propulsive work":                                  dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::right time to peak braking force":                       dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::right time to peak braking power":                       dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::right time to peak landing force":                       dict(mean_m=0.07678116667, sd_m=0.02166365501,  mean_f=0.09826466667,   sd_f=0.01540453421),  # M n=5, F n=5
    "dj::right time to peak propulsive force":                    dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::right time to peak propulsive power":                    dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::rsi":                                                    dict(mean_m=0.7132011333,  sd_m=0.3422324564,   mean_f=0.9402826667,    sd_f=0.5172975235),  # M n=5, F n=5
    "dj::rsi exponential":                                        dict(mean_m=0.2923863333,  sd_m=0.1171333813,   mean_f=0.3753955,       sd_f=0.2410939153),  # M n=5, F n=5
    "dj::rsi modified":                                           dict(mean_m=0.3584141,     sd_m=0.1435851458,   mean_f=0.4601715,       sd_f=0.2955394924),  # M n=5, F n=5
    "dj::takeoff time":                                           dict(mean_m=1.17762038,    sd_m=0.2301071003,   mean_f=0.9898956833,    sd_f=0.180610468),  # M n=5, F n=5
    "dj::takeoff velocity":                                       dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::time to peak braking force":                             dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::time to peak braking force sym. index":                  dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::time to peak braking power":                             dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::time to peak braking power sym. index":                  dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::time to peak force":                                     dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::time to peak landing force":                             dict(mean_m=0.074267,      sd_m=0.01931402465,  mean_f=0.08484416667,   sd_f=0.009186407569),  # M n=5, F n=5
    "dj::time to peak landing force sym. index":                  dict(mean_m=-0.2294033333, sd_m=5.377922294,    mean_f=-9.773433333,    sd_f=7.486294946),  # M n=5, F n=5
    "dj::time to peak power":                                     dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::time to peak propulsive force":                          dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::time to peak propulsive force sym. index":               dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::time to peak propulsive power":                          dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::time to peak propulsive power sym. index":               dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "dj::time to takeoff":                                        dict(mean_m=1.1811044,     sd_m=0.2307441303,   mean_f=0.994967,        sd_f=0.1807839374),  # M n=5, F n=5
    "dj::velocity peak power":                                    dict(mean_m=None,          sd_m=None,           mean_f=None,            sd_f=None),  # M n=1, F n=0
    "imtp::body mass":                                            dict(mean_m=79.20867491,   sd_m=6.137467679,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::end of rise force":                                    dict(mean_m=1523.961092,   sd_m=136.9321555,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::end of rise force net":                                dict(mean_m=747.1893402,   sd_m=76.91488566,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::end of rise force relative":                           dict(mean_m=196.0730325,   sd_m=2.501719629,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::end of rise force relative net":                       dict(mean_m=96.07303249,   sd_m=2.501719629,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 100ms":                                          dict(mean_m=1210.326155,   sd_m=229.1400671,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 100ms left":                                     dict(mean_m=630.5832969,   sd_m=88.28968974,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 100ms left net":                                 dict(mean_m=242.197421,    sd_m=62.19708479,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 100ms left relative":                            dict(mean_m=80.9740452,    sd_m=6.399078482,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 100ms left relative net":                        dict(mean_m=30.9740452,    sd_m=6.399078482,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 100ms net":                                      dict(mean_m=433.5544029,   sd_m=190.5582987,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 100ms relative":                                 dict(mean_m=155.3415833,   sd_m=23.43465285,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 100ms relative net":                             dict(mean_m=55.34158326,   sd_m=23.43465285,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 100ms right":                                    dict(mean_m=579.7428578,   sd_m=148.8558613,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 100ms right net":                                dict(mean_m=191.3569819,   sd_m=134.1204937,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 100ms right relative":                           dict(mean_m=74.36753806,   sd_m=17.35370678,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 100ms right relative net":                       dict(mean_m=24.36753806,   sd_m=17.35370678,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 100ms si":                                       dict(mean_m=4.903969552,   sd_m=8.338373594,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 150ms":                                          dict(mean_m=1342.725779,   sd_m=214.5459009,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 150ms left":                                     dict(mean_m=699.4465416,   sd_m=78.49055452,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 150ms left net":                                 dict(mean_m=311.0606657,   sd_m=48.88304728,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 150ms left relative":                            dict(mean_m=89.89319794,   sd_m=3.324200087,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 150ms left relative net":                        dict(mean_m=39.89319794,   sd_m=3.324200087,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 150ms net":                                      dict(mean_m=565.9540267,   sd_m=171.8458105,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 150ms relative":                                 dict(mean_m=172.4388907,   sd_m=19.94668788,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 150ms relative net":                             dict(mean_m=72.43889071,   sd_m=19.94668788,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 150ms right":                                    dict(mean_m=643.279237,    sd_m=152.421325,     mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 150ms right net":                                dict(mean_m=254.8933611,   sd_m=136.9452251,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 150ms right relative":                           dict(mean_m=82.54569277,   sd_m=17.45773121,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 150ms right relative net":                       dict(mean_m=32.54569277,   sd_m=17.45773121,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 150ms si":                                       dict(mean_m=4.77802251,    sd_m=9.442558103,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 200ms":                                          dict(mean_m=1423.362375,   sd_m=199.7597338,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 200ms left":                                     dict(mean_m=738.8545681,   sd_m=62.13382313,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 200ms left net":                                 dict(mean_m=350.4686922,   sd_m=32.30464854,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 200ms left relative":                            dict(mean_m=95.08981276,   sd_m=0.9629414045,   mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 200ms left relative net":                        dict(mean_m=45.08981276,   sd_m=0.9629414045,   mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 200ms net":                                      dict(mean_m=646.5906233,   sd_m=155.5093119,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 200ms relative":                                 dict(mean_m=182.9033104,   sd_m=17.33047843,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 200ms relative net":                             dict(mean_m=82.90331036,   sd_m=17.33047843,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 200ms right":                                    dict(mean_m=684.507807,    sd_m=159.0375251,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 200ms right net":                                dict(mean_m=296.1219311,   sd_m=142.3009257,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 200ms right relative":                           dict(mean_m=87.8134976,    sd_m=17.86588759,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 200ms right relative net":                       dict(mean_m=37.8134976,    sd_m=17.86588759,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 200ms si":                                       dict(mean_m=4.724551473,   sd_m=10.5346349,     mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 250ms":                                          dict(mean_m=1515.402173,   sd_m=161.8796202,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 250ms left":                                     dict(mean_m=780.9306405,   sd_m=33.06207303,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 250ms left net":                                 dict(mean_m=392.5447646,   sd_m=10.73601744,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 250ms left relative":                            dict(mean_m=100.7240925,   sd_m=3.948712211,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 250ms left relative net":                        dict(mean_m=50.72409253,   sd_m=3.948712211,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 250ms net":                                      dict(mean_m=738.6304208,   sd_m=124.194812,     mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 250ms relative":                                 dict(mean_m=195.0624717,   sd_m=14.65691568,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 250ms relative net":                             dict(mean_m=95.0624717,    sd_m=14.65691568,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 250ms right":                                    dict(mean_m=734.471532,    sd_m=148.7816583,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 250ms right net":                                dict(mean_m=346.0856561,   sd_m=133.1296,       mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 250ms right relative":                           dict(mean_m=94.33837917,   sd_m=16.67994333,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 250ms right relative net":                       dict(mean_m=44.33837917,   sd_m=16.67994333,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 250ms si":                                       dict(mean_m=3.993564004,   sd_m=10.28921831,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 50ms":                                           dict(mean_m=1101.60925,    sd_m=222.4149525,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 50ms left":                                      dict(mean_m=571.319908,    sd_m=96.90692298,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 50ms left net":                                  dict(mean_m=182.9340321,   sd_m=71.5338827,     mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 50ms left relative":                             dict(mean_m=73.27720039,   sd_m=8.060317038,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 50ms left relative net":                         dict(mean_m=23.27720039,   sd_m=8.060317038,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 50ms net":                                       dict(mean_m=324.8374986,   sd_m=175.8190597,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 50ms relative":                                  dict(mean_m=141.1697588,   sd_m=21.10751705,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 50ms relative net":                              dict(mean_m=41.16975878,   sd_m=21.10751705,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 50ms right":                                     dict(mean_m=530.2893423,   sd_m=126.6728514,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 50ms right net":                                 dict(mean_m=141.9034664,   sd_m=105.0801217,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 50ms right relative":                            dict(mean_m=67.89255839,   sd_m=13.08725874,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 50ms right relative net":                        dict(mean_m=17.89255839,   sd_m=13.08725874,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force 50ms si":                                        dict(mean_m=4.164748002,   sd_m=4.606495155,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force at max rfd":                                     dict(mean_m=1287.16018,    sd_m=127.3614888,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force at max rfd left":                                dict(mean_m=677.5552222,   sd_m=116.4238211,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force at max rfd left net":                            dict(mean_m=289.1693463,   sd_m=136.3976346,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force at max rfd left relative":                       dict(mean_m=88.016684,     sd_m=19.9933343,     mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force at max rfd left relative net":                   dict(mean_m=38.016684,     sd_m=19.9933343,     mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force at max rfd net":                                 dict(mean_m=510.3884281,   sd_m=170.9231272,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force at max rfd relative":                            dict(mean_m=166.8574981,   sd_m=26.28835773,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force at max rfd relative net":                        dict(mean_m=66.85749809,   sd_m=26.28835773,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force at max rfd right":                               dict(mean_m=609.6049577,   sd_m=11.34799041,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force at max rfd right net":                           dict(mean_m=221.2190818,   sd_m=39.69699698,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force at max rfd right relative":                      dict(mean_m=78.84081409,   sd_m=7.078210995,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force at max rfd right relative net":                  dict(mean_m=28.84081409,   sd_m=7.078210995,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::force at max rfd si":                                  dict(mean_m=5.19610329,    sd_m=7.931903825,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::impulse 100ms":                                        dict(mean_m=107.604006,    sd_m=20.48650919,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::impulse 150ms":                                        dict(mean_m=171.1020911,   sd_m=31.7411947,     mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::impulse 200ms":                                        dict(mean_m=240.4120784,   sd_m=42.2886357,     mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::impulse 250ms":                                        dict(mean_m=313.5982278,   sd_m=51.29450942,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::impulse 50ms":                                         dict(mean_m=49.97245195,   sd_m=9.486313578,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::impulse at max rfd":                                   dict(mean_m=385.7619896,   sd_m=509.8016989,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::max rfd":                                              dict(mean_m=7912.470566,   sd_m=1774.841924,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::onset force":                                          dict(mean_m=924.0386702,   sd_m=126.667605,     mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::onset force net":                                      dict(mean_m=147.2669184,   sd_m=74.31056037,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::onset force relative":                                 dict(mean_m=118.6587175,   sd_m=8.685733515,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::onset force relative net":                             dict(mean_m=18.65871751,   sd_m=8.685733515,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::onset time":                                           dict(mean_m=0.99375,       sd_m=0.07825067053,  mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::peak force left":                                      dict(mean_m=1108.790149,   sd_m=85.77347775,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::peak force left net":                                  dict(mean_m=720.4042732,   sd_m=69.73497486,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::peak force left relative":                             dict(mean_m=142.9398225,   sd_m=9.527935491,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::peak force left relative net":                         dict(mean_m=92.93982249,   sd_m=9.527935491,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::peak force net":                                       dict(mean_m=1385.278746,   sd_m=128.4643445,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::peak force relative":                                  dict(mean_m=278.2288637,   sd_m=5.145940523,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::peak force relative net":                              dict(mean_m=178.2288637,   sd_m=5.145940523,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::peak force right":                                     dict(mean_m=1053.260348,   sd_m=161.4747751,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::peak force right net":                                 dict(mean_m=664.8744725,   sd_m=138.9666137,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::peak force right relative":                            dict(mean_m=135.2890412,   sd_m=14.67362607,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::peak force right relative net":                        dict(mean_m=85.28904122,   sd_m=14.67362607,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::peak force si":                                        dict(mean_m=2.806707117,   sd_m=8.835167329,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::rfd 100ms":                                            dict(mean_m=2862.874845,   sd_m=1216.020661,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::rfd 150ms":                                            dict(mean_m=2791.247389,   sd_m=668.9765617,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::rfd 200ms":                                            dict(mean_m=2496.618524,   sd_m=417.1194758,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::rfd 250ms":                                            dict(mean_m=2365.454009,   sd_m=247.1212051,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::rfd 50ms":                                             dict(mean_m=3551.411602,   sd_m=2042.414923,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::rfd impulse":                                          dict(mean_m=536.4724127,   sd_m=389.9685186,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::steadiness rsme force":                                dict(mean_m=263.8841782,   sd_m=83.0368858,     mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::steadiness rsme rfd":                                  dict(mean_m=1320.877514,   sd_m=244.3802331,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::system weight":                                        dict(mean_m=776.7717518,   sd_m=60.18799741,    mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::time to force end of rise":                            dict(mean_m=0.7089699074,  sd_m=0.2926359521,   mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::time to peak force":                                   dict(mean_m=1.715451389,   sd_m=0.1056544504,   mean_f=None,            sd_f=None),  # M n=3, F n=0
    "imtp::time to peak rfd":                                     dict(mean_m=0.2789641204,  sd_m=0.3221964079,   mean_f=None,            sd_f=None),  # M n=3, F n=0
    "sj::avg landing rfd":                                        dict(mean_m=65977.58745,   sd_m=59884.86857,    mean_f=22213.27319,     sd_f=8655.286239),  # M n=27, F n=7
    "sj::avg landing rfd sym. index":                             dict(mean_m=0.02791790886, sd_m=19.26076667,    mean_f=-9.468238095,    sd_f=31.04272154),  # M n=27, F n=7
    "sj::avg propulsive force sym. index":                        dict(mean_m=1.526568028,   sd_m=3.179438036,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::avg propulsive power sym. index":                        dict(mean_m=1.82807541,    sd_m=4.242535756,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::avg. propulsive force":                                  dict(mean_m=1094.548791,   sd_m=117.1205223,    mean_f=848.5694286,     sd_f=130.7908577),  # M n=27, F n=7
    "sj::avg. propulsive velocity":                               dict(mean_m=0.8791385468,  sd_m=0.1775566088,   mean_f=0.8639187381,    sd_f=0.1732585256),  # M n=27, F n=7
    "sj::avg. rfd":                                               dict(mean_m=2628.069404,   sd_m=1103.608254,    mean_f=2554.597595,     sd_f=1582.937917),  # M n=27, F n=7
    "sj::avg. rfd sym. index":                                    dict(mean_m=2.701856139,   sd_m=7.373163581,    mean_f=0.5604666667,    sd_f=9.165245273),  # M n=27, F n=7
    "sj::body mass":                                              dict(mean_m=75.8377476,    sd_m=5.706161302,    mean_f=59.0726681,      sd_f=7.364086646),  # M n=27, F n=7
    "sj::body weight":                                            dict(mean_m=743.7141451,   sd_m=55.95819883,    mean_f=579.3050238,     sd_f=72.21671568),  # M n=27, F n=7
    "sj::body weight sd":                                         dict(mean_m=7.238277076,   sd_m=1.175088731,    mean_f=7.0095,          sd_f=0.02672718827),  # M n=27, F n=7
    "sj::contact time":                                           dict(mean_m=0.5386239352,  sd_m=0.2871596477,   mean_f=0.4117931667,    sd_f=0.09942743726),  # M n=27, F n=7
    "sj::flight threshold":                                       dict(mean_m=4.658666341,   sd_m=1.598653463,    mean_f=4.19802381,      sd_f=1.402824243),  # M n=27, F n=7
    "sj::flight time":                                            dict(mean_m=0.4199480144,  sd_m=0.04995830114,  mean_f=0.3559515714,    sd_f=0.04338801605),  # M n=27, F n=7
    "sj::force peak power":                                       dict(mean_m=1431.020116,   sd_m=118.549825,     mean_f=1100.686405,     sd_f=163.6533862),  # M n=27, F n=7
    "sj::initiation threshold":                                   dict(mean_m=763.7141451,   sd_m=55.95819883,    mean_f=599.3050238,     sd_f=72.21671568),  # M n=27, F n=7
    "sj::jump height ni":                                         dict(mean_m=22.13511541,   sd_m=4.261191488,    mean_f=15.96734048,     sd_f=5.275230264),  # M n=27, F n=7
    "sj::jump momentum":                                          dict(mean_m=156.9383124,   sd_m=19.85484848,    mean_f=103.062681,      sd_f=19.40308509),  # M n=27, F n=7
    "sj::jump start time":                                        dict(mean_m=0.5782893354,  sd_m=0.1008244541,   mean_f=0.5734454643,    sd_f=0.0811708251),  # M n=27, F n=7
    "sj::jump threshold time":                                    dict(mean_m=0.5782893354,  sd_m=0.1008244541,   mean_f=0.5734454643,    sd_f=0.0811708251),  # M n=27, F n=7
    "sj::landing peak force time":                                dict(mean_m=1.601149099,   sd_m=0.2871547376,   mean_f=1.42192125,      sd_f=0.1521357942),  # M n=27, F n=7
    "sj::landing rfd 0-20ms":                                     dict(mean_m=97110.99393,   sd_m=79120.5688,     mean_f=None,            sd_f=None),  # M n=10, F n=0
    "sj::landing rfd 0-20ms sym. index":                          dict(mean_m=8.019013875,   sd_m=28.84350342,    mean_f=None,            sd_f=None),  # M n=10, F n=0
    "sj::landing rfd 0-40ms":                                     dict(mean_m=51879.27313,   sd_m=24666.64024,    mean_f=None,            sd_f=None),  # M n=10, F n=0
    "sj::landing rfd 0-40ms sym. index":                          dict(mean_m=8.263552685,   sd_m=21.33525637,    mean_f=None,            sd_f=None),  # M n=10, F n=0
    "sj::landing rfd 0-60ms":                                     dict(mean_m=33008.63058,   sd_m=13456.45456,    mean_f=None,            sd_f=None),  # M n=10, F n=0
    "sj::landing rfd 0-60ms sym. index":                          dict(mean_m=6.223702474,   sd_m=18.83556753,    mean_f=None,            sd_f=None),  # M n=10, F n=0
    "sj::landing rfd 0-80ms":                                     dict(mean_m=22545.17289,   sd_m=5752.531736,    mean_f=None,            sd_f=None),  # M n=10, F n=0
    "sj::landing rfd 0-80ms sym. index":                          dict(mean_m=7.682662403,   sd_m=21.50436673,    mean_f=None,            sd_f=None),  # M n=10, F n=0
    "sj::landing time":                                           dict(mean_m=1.536869557,   sd_m=0.2769426016,   mean_f=1.341190393,     sd_f=0.1484076113),  # M n=27, F n=7
    "sj::left avg landing rfd":                                   dict(mean_m=38304.89857,   sd_m=30580.90775,    mean_f=11369.25224,     sd_f=5700.585643),  # M n=27, F n=7
    "sj::left avg propulsive force":                              dict(mean_m=551.1077256,   sd_m=30.47162288,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::left avg propulsive power":                              dict(mean_m=532.8730804,   sd_m=58.69969483,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::left avg rfd":                                           dict(mean_m=1347.347516,   sd_m=545.8036993,    mean_f=1295.747524,     sd_f=810.5727266),  # M n=27, F n=7
    "sj::left landing rfd 0-20ms":                                dict(mean_m=44764.83574,   sd_m=33240.46716,    mean_f=None,            sd_f=None),  # M n=10, F n=0
    "sj::left landing rfd 0-40ms":                                dict(mean_m=26955.24464,   sd_m=11150.40386,    mean_f=None,            sd_f=None),  # M n=10, F n=0
    "sj::left landing rfd 0-60ms":                                dict(mean_m=16813.44501,   sd_m=6723.919239,    mean_f=None,            sd_f=None),  # M n=10, F n=0
    "sj::left landing rfd 0-80ms":                                dict(mean_m=11906.9459,    sd_m=3135.181582,    mean_f=None,            sd_f=None),  # M n=10, F n=0
    "sj::left net impulse":                                       dict(mean_m=80.58877186,   sd_m=13.42444355,    mean_f=50.48216667,     sd_f=12.75901017),  # M n=27, F n=7
    "sj::left p1 impulse":                                        dict(mean_m=163.5323024,   sd_m=92.45383488,    mean_f=90.65827857,     sd_f=18.27389845),  # M n=27, F n=7
    "sj::left p2 impulse":                                        dict(mean_m=110.7431948,   sd_m=23.32052179,    mean_f=75.2089,         sd_f=15.08457261),  # M n=27, F n=7
    "sj::left peak force":                                        dict(mean_m=777.4638555,   sd_m=74.77417154,    mean_f=588.0704524,     sd_f=111.9066317),  # M n=27, F n=7
    "sj::left peak landing force":                                dict(mean_m=2184.724637,   sd_m=754.3737521,    mean_f=1159.05581,      sd_f=408.0442183),  # M n=27, F n=7
    "sj::left peak propulsive force":                             dict(mean_m=753.8561658,   sd_m=36.73130497,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::left peak propulsive power":                             dict(mean_m=1384.176658,   sd_m=167.1211741,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::left propulsive impulse":                                dict(mean_m=274.2755064,   sd_m=89.04031071,    mean_f=165.8670762,     sd_f=30.39315934),  # M n=27, F n=7
    "sj::left propulsive rfd":                                    dict(mean_m=1319.157153,   sd_m=434.7278716,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::left propulsive work":                                   dict(mean_m=224.8418923,   sd_m=32.51897793,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::left time to peak force":                                dict(mean_m=0.4142316204,  sd_m=0.2904990313,   mean_f=0.2956688571,    sd_f=0.1142009455),  # M n=27, F n=7
    "sj::left time to peak landing force":                        dict(mean_m=0.067340607,   sd_m=0.02808291717,  mean_f=0.08273995238,   sd_f=0.01287218304),  # M n=27, F n=7
    "sj::left time to peak power":                                dict(mean_m=0.35901625,    sd_m=0.08444537666,  mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::left time to peak propulsive force":                     dict(mean_m=0.3213335417,  sd_m=0.09685659247,  mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::left time to peak propulsive power":                     dict(mean_m=0.35901625,    sd_m=0.08444537666,  mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::net impulse":                                            dict(mean_m=156.9383216,   sd_m=19.85496354,    mean_f=103.0617143,     sd_f=19.40332676),  # M n=27, F n=7
    "sj::net impulse sym. index":                                 dict(mean_m=-0.4143571852, sd_m=8.662198527,    mean_f=-3.116538095,    sd_f=7.663940702),  # M n=27, F n=7
    "sj::p1 avg force":                                           dict(mean_m=1045.457705,   sd_m=133.4284002,    mean_f=836.3189214,     sd_f=142.3672102),  # M n=27, F n=7
    "sj::p1 avg power":                                           dict(mean_m=304.6837354,   sd_m=54.25763057,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::p1 avg velocity":                                        dict(mean_m=0.2645028921,  sd_m=0.03257921234,  mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::p1 duration":                                            dict(mean_m=0.3544230298,  sd_m=0.2950361582,   mean_f=0.2347714524,    sd_f=0.07995558856),  # M n=27, F n=7
    "sj::p1 impulse":                                             dict(mean_m=74.14897776,   sd_m=13.71902194,    mean_f=51.53148571,     sd_f=9.701523646),  # M n=27, F n=7
    "sj::p1 impulse sym. index":                                  dict(mean_m=-1.114079453,  sd_m=4.775043166,    mean_f=-2.091985714,    sd_f=6.530603923),  # M n=27, F n=7
    "sj::p1 p2 duration ratio":                                   dict(mean_m=2.060343801,   sd_m=1.90804627,     mean_f=1.3323,          sd_f=0.3983375874),  # M n=27, F n=7
    "sj::p1 p2 force ratio":                                      dict(mean_m=0.871086982,   sd_m=0.0853360395,   mean_f=0.9533838095,    sd_f=0.06717003219),  # M n=27, F n=7
    "sj::p1 p2 power ratio":                                      dict(mean_m=0.166729886,   sd_m=0.01848091169,  mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::p1 p2 velocity ratio":                                   dict(mean_m=0.1627855325,  sd_m=0.01604597694,  mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::p1 peak force":                                          dict(mean_m=1239.879148,   sd_m=142.9425055,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::p1 peak power":                                          dict(mean_m=964.9962849,   sd_m=207.9171508,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::p1 peak velocity":                                       dict(mean_m=0.7683736022,  sd_m=0.09330140674,  mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::p2 avg force":                                           dict(mean_m=1200.237334,   sd_m=97.23183002,    mean_f=874.9725667,     sd_f=120.2021125),  # M n=27, F n=7
    "sj::p2 avg power":                                           dict(mean_m=1795.620893,   sd_m=227.6492361,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::p2 avg velocity":                                        dict(mean_m=1.59883442,    sd_m=0.1097220102,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::p2 duration":                                            dict(mean_m=0.1842006276,  sd_m=0.03777833973,  mean_f=0.1770217143,    sd_f=0.02989708077),  # M n=27, F n=7
    "sj::p2 impulse":                                             dict(mean_m=82.78939941,   sd_m=12.84633326,    mean_f=51.53148571,     sd_f=9.701523646),  # M n=27, F n=7
    "sj::p2 impulse sym. index":                                  dict(mean_m=0.2589806082,  sd_m=3.871317771,    mean_f=-2.472416667,    sd_f=4.447886051),  # M n=27, F n=7
    "sj::p2 peak force":                                          dict(mean_m=1490.830066,   sd_m=92.503424,      mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::p2 peak power":                                          dict(mean_m=2751.605499,   sd_m=353.7531182,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::p2 peak velocity":                                       dict(mean_m=2.215796426,   sd_m=0.117235742,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::peak force":                                             dict(mean_m=1558.855748,   sd_m=155.6730896,    mean_f=1191.274357,     sd_f=206.132843),  # M n=27, F n=7
    "sj::peak force sym. index":                                  dict(mean_m=-0.3085892747, sd_m=3.435017053,    mean_f=-1.572021429,    sd_f=3.086769741),  # M n=27, F n=7
    "sj::peak force time":                                        dict(mean_m=0.9961489033,  sd_m=0.285608796,    mean_f=0.8698112262,    sd_f=0.1591651596),  # M n=27, F n=7
    "sj::peak landing force":                                     dict(mean_m=3713.033771,   sd_m=989.739567,     mean_f=2239.300524,     sd_f=550.1556684),  # M n=27, F n=7
    "sj::peak landing force sym. index":                          dict(mean_m=0.4680042603,  sd_m=12.12877951,    mean_f=-6.55265,        sd_f=21.3058373),  # M n=27, F n=7
    "sj::peak power":                                             dict(mean_m=2854.956418,   sd_m=426.6425153,    mean_f=1923.715524,     sd_f=392.1503108),  # M n=27, F n=7
    "sj::peak propulsive force":                                  dict(mean_m=1558.855748,   sd_m=155.6730896,    mean_f=1191.274357,     sd_f=206.132843),  # M n=27, F n=7
    "sj::peak propulsive force sym. index":                       dict(mean_m=1.140945307,   sd_m=2.175582471,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::peak propulsive power":                                  dict(mean_m=2854.956418,   sd_m=426.6425153,    mean_f=1923.715524,     sd_f=392.1503108),  # M n=27, F n=7
    "sj::peak propulsive power sym. index":                       dict(mean_m=0.6229356967,  sd_m=3.085900669,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::peak propulsive velocity":                               dict(mean_m=2.238616783,   sd_m=0.1868493258,   mean_f=1.964956595,     sd_f=0.2559748807),  # M n=27, F n=7
    "sj::peak velocity":                                          dict(mean_m=2.238616783,   sd_m=0.1868493258,   mean_f=1.964956595,     sd_f=0.2559748807),  # M n=27, F n=7
    "sj::propulsive duration":                                    dict(mean_m=0.5386239352,  sd_m=0.2871596477,   mean_f=0.4117931667,    sd_f=0.09942743726),  # M n=27, F n=7
    "sj::propulsive impulse":                                     dict(mean_m=156.9383124,   sd_m=19.85484848,    mean_f=103.062681,      sd_f=19.40308509),  # M n=27, F n=7
    "sj::propulsive impulse sym. index":                          dict(mean_m=-0.6438762731, sd_m=4.147847499,    mean_f=-2.327045238,    sd_f=4.978517299),  # M n=27, F n=7
    "sj::propulsive rfd":                                         dict(mean_m=2563.826548,   sd_m=866.3818558,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::propulsive rfd sym. index":                              dict(mean_m=3.440158301,   sd_m=2.680952212,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::propulsive start time":                                  dict(mean_m=47.26620931,   sd_m=29.31089311,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::propulsive work":                                        dict(mean_m=470.7070834,   sd_m=90.39599742,    mean_f=288.552619,      sd_f=74.92955267),  # M n=27, F n=7
    "sj::propulsive work sym. index":                             dict(mean_m=1.82807541,    sd_m=4.242535756,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::rel. avg. propulsive force":                             dict(mean_m=14.7492519,    sd_m=0.848108288,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::rel. propulsive impulse":                                dict(mean_m=2.060311938,   sd_m=0.1222242191,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::relative peak force":                                    dict(mean_m=1.930494163,   sd_m=0.1907426975,   mean_f=None,            sd_f=None),  # M n=10, F n=0
    "sj::relative peak landing force":                            dict(mean_m=4.976291969,   sd_m=1.254889319,    mean_f=3.860813095,     sd_f=0.8224170076),  # M n=27, F n=7
    "sj::relative peak power":                                    dict(mean_m=37.75580789,   sd_m=5.598079098,    mean_f=32.71387381,     sd_f=6.512457139),  # M n=27, F n=7
    "sj::right avg landing rfd":                                  dict(mean_m=36509.39221,   sd_m=27496.14649,    mean_f=14951.3091,      sd_f=8871.744867),  # M n=27, F n=7
    "sj::right avg propulsive force":                             dict(mean_m=535.4948892,   sd_m=50.39325357,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::right avg propulsive power":                             dict(mean_m=516.0240664,   sd_m=84.9666805,     mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::right avg rfd":                                          dict(mean_m=1283.061449,   sd_m=550.4697158,    mean_f=1281.794071,     sd_f=809.8352887),  # M n=27, F n=7
    "sj::right landing rfd 0-20ms":                               dict(mean_m=52346.15786,   sd_m=47324.54843,    mean_f=None,            sd_f=None),  # M n=10, F n=0
    "sj::right landing rfd 0-40ms":                               dict(mean_m=24924.029,     sd_m=14234.10769,    mean_f=None,            sd_f=None),  # M n=10, F n=0
    "sj::right landing rfd 0-60ms":                               dict(mean_m=16195.1859,    sd_m=7950.990251,    mean_f=None,            sd_f=None),  # M n=10, F n=0
    "sj::right landing rfd 0-80ms":                               dict(mean_m=10638.227,     sd_m=4088.625609,    mean_f=None,            sd_f=None),  # M n=10, F n=0
    "sj::right net impulse":                                      dict(mean_m=79.38066087,   sd_m=12.85351318,    mean_f=52.58109524,     sd_f=7.038013416),  # M n=27, F n=7
    "sj::right p1 impulse":                                       dict(mean_m=172.6469502,   sd_m=124.239025,     mean_f=95.20240238,     sd_f=24.19650892),  # M n=27, F n=7
    "sj::right p2 impulse":                                       dict(mean_m=109.8025022,   sd_m=20.50492097,    mean_f=78.77213571,     sd_f=13.86921153),  # M n=27, F n=7
    "sj::right peak force":                                       dict(mean_m=783.7408264,   sd_m=90.00031741,    mean_f=603.9181429,     sd_f=96.62871006),  # M n=27, F n=7
    "sj::right peak landing force":                               dict(mean_m=2106.18407,    sd_m=602.8381636,    mean_f=1338.532286,     sd_f=418.570436),  # M n=27, F n=7
    "sj::right peak propulsive force":                            dict(mean_m=738.1633876,   sd_m=58.90390356,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::right peak propulsive power":                            dict(mean_m=1369.732279,   sd_m=195.9409925,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::right propulsive impulse":                               dict(mean_m=282.4494895,   sd_m=122.1301579,    mean_f=173.9744905,     sd_f=31.95417947),  # M n=27, F n=7
    "sj::right propulsive rfd":                                   dict(mean_m=1239.469951,   sd_m=423.9803721,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::right propulsive work":                                  dict(mean_m=216.4843784,   sd_m=30.23923374,    mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::right time to peak force":                               dict(mean_m=0.4199460031,  sd_m=0.2846569085,   mean_f=0.2954826667,    sd_f=0.1155658074),  # M n=27, F n=7
    "sj::right time to peak landing force":                       dict(mean_m=0.065656857,   sd_m=0.02401273541,  mean_f=0.08068995238,   sd_f=0.01750147983),  # M n=27, F n=7
    "sj::right time to peak power":                               dict(mean_m=0.3583920833,  sd_m=0.08198885468,  mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::right time to peak propulsive force":                    dict(mean_m=0.3248407639,  sd_m=0.09465374744,  mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::right time to peak propulsive power":                    dict(mean_m=0.3583920833,  sd_m=0.08198885468,  mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::rsi":                                                    dict(mean_m=0.9519319968,  sd_m=0.2561298335,   mean_f=0.936050119,     sd_f=0.2982867443),  # M n=27, F n=7
    "sj::rsi modified":                                           dict(mean_m=0.4992246588,  sd_m=0.1620858051,   mean_f=0.4209962381,    sd_f=0.1832986393),  # M n=27, F n=7
    "sj::takeoff time":                                           dict(mean_m=1.116903394,   sd_m=0.2869676072,   mean_f=0.9852371548,    sd_f=0.1439164514),  # M n=27, F n=7
    "sj::takeoff velocity":                                       dict(mean_m=2.068984873,   sd_m=0.2052908925,   mean_f=1.747288452,     sd_f=0.2897216536),  # M n=27, F n=7
    "sj::time to peak force":                                     dict(mean_m=0.4178564506,  sd_m=0.2876142256,   mean_f=0.2963667381,    sd_f=0.1140381587),  # M n=27, F n=7
    "sj::time to peak landing force":                             dict(mean_m=0.0642671965,  sd_m=0.02367809108,  mean_f=0.08072854762,   sd_f=0.01440232233),  # M n=27, F n=7
    "sj::time to peak landing force sym. index":                  dict(mean_m=-0.01997719699, sd_m=9.454646077,    mean_f=1.724354762,     sd_f=6.92335907),  # M n=27, F n=7
    "sj::time to peak power":                                     dict(mean_m=0.4631534465,  sd_m=0.287667575,    mean_f=0.3333034048,    sd_f=0.1005039097),  # M n=27, F n=7
    "sj::time to peak power sym. index":                          dict(mean_m=0.03490947328, sd_m=0.6470212336,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::time to peak propulsive force":                          dict(mean_m=0.3222688889,  sd_m=0.09652100901,  mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::time to peak propulsive force sym. index":               dict(mean_m=-0.6844778174, sd_m=0.6569702332,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::time to peak propulsive power":                          dict(mean_m=0.3585826389,  sd_m=0.08266406749,  mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::time to peak propulsive power sym. index":               dict(mean_m=0.03490947328, sd_m=0.6470212336,   mean_f=None,            sd_f=None),  # M n=6, F n=0
    "sj::time to peak si":                                        dict(mean_m=-1.368052624,  sd_m=4.222750871,    mean_f=0.1534333333,    sd_f=1.156432921),  # M n=27, F n=7
    "sj::velocity peak power":                                    dict(mean_m=1.986256307,   sd_m=0.1582177313,   mean_f=1.744664357,     sd_f=0.2399024865),  # M n=27, F n=7
}
# Set di norme selezionabili dalla barra laterale. Il primo è il default.
NORM_SETS = {
    "Team DU Coaching": TEAM_POP,
    "Letteratura": LITERATURE_POP,   
}


def _copia_norme(nome_set):
    """Copia profonda di un set di norme, così le modifiche nella scheda
    Costanti non alterano mai il dizionario originale. Le chiavi curate
    assenti nel set scelto (o con valori None) vengono prese dalla
    letteratura, così un set incompleto non manda in crash l'app."""
    out = {k: dict(v) for k, v in NORM_SETS[nome_set].items()}
    for k, v in LITERATURE_POP.items():
        c = out.get(k)
        if c is None or None in (c.get("mean_m"), c.get("sd_m"), c.get("mean_f"), c.get("sd_f")):
            out[k] = dict(v)
    return out


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
    """Costruisce il DataFrame mostrato/editato nella scheda Costanti: prima
    le costanti curate, poi le norme extra 'tipo::variabile' (presenti solo
    nel set Team). La colonna 'Chiave' è nascosta nella UI e serve a
    riscrivere ogni riga nella chiave giusta."""
    rows = []
    for key, (label, unit, categoria) in CONST_LABELS.items():
        c = pop_dict[key]
        rows.append({
            "Chiave": key, "Categoria": categoria, "Costante": label, "Unità": unit,
            "Media Uomini": c["mean_m"], "Dev.Std Uomini": c["sd_m"],
            "Media Donne": c["mean_f"], "Dev.Std Donne": c["sd_f"],
        })
    for key in sorted(k for k in pop_dict if "::" in k):
        jump_type, var = key.split("::", 1)
        c = pop_dict[key]
        rows.append({
            "Chiave": key,
            "Categoria": f"{JUMP_TYPE_LABELS.get(jump_type, jump_type)} (extra)",
            "Costante": extra_metric_label(var), "Unità": "",
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
#         (indice aggregato cross-test, es. DSI/EUR — letto per zona, non
#         per T-score), "info" (valore mostrato senza confronto)
METRICS = [
    dict(key="imtp_peak_force", label="IMTP Peak Force", category="ISOMETRIC PULL TEST",
         jump_type="imtp", raw_var="peak force", unit="N",
         pop_key="imtp_peak_force", lower_is_better=False, kind="score"),
    dict(key="imtp_rel_peak_force", label="IMTP Rel Peak Force", category="ISOMETRIC PULL TEST",
         jump_type="imtp", raw_var=None,
         derive=lambda rep: _safe_div(rep.get("peak force"), rep.get("body mass")),
         unit="N/kg", pop_key="imtp_rel_peak_force", lower_is_better=False, kind="score"),

    # Set di default per SJ / CMJ / CMJ RE, nello stesso ordine per ogni test:
    # OUTCOME, FORCE, TIME, IMPULSE, POWER, MOVEMENT STRATEGY.
    # SQUAT JUMP (niente TIME e MOVEMENT STRATEGY)
    dict(key="sj_height", label="SJ Height", category="SQUAT JUMP TEST",
         jump_type="sj", raw_var="jump height ft", unit="cm",
         pop_key="sj_height", lower_is_better=False, kind="score"),
    dict(key="sj_peak_force", label="SJ Peak Force", category="SQUAT JUMP TEST",
         jump_type="sj", raw_var="peak propulsive force", unit="N",
         pop_key=None, lower_is_better=False, kind="info"),
    dict(key="sj_net_rel_impulse", label="SJ Net Rel Impulse", category="SQUAT JUMP TEST",
         jump_type="sj", raw_var=None,
         derive=lambda rep: _safe_div(rep.get("net impulse"), rep.get("body mass")),
         unit="N\u00b7s/kg", pop_key="sj_net_rel_impulse", lower_is_better=False, kind="info"),
    dict(key="sj_mean_power", label="SJ Mean Power", category="SQUAT JUMP TEST",
         jump_type="sj", raw_var="avg. propulsive power", unit="W",
         pop_key="sj_mean_power", lower_is_better=False, kind="score"),

    # COUNTERMOVEMENT JUMP
    dict(key="cmj_height", label="CMJ Height", category="COUNTERMOVEMENT JUMP TEST",
         jump_type="cmj", raw_var="jump height ft", unit="cm",
         pop_key="cmj_height", lower_is_better=False, kind="score"),
    dict(key="cmj_peak_force", label="CMJ Peak Force", category="COUNTERMOVEMENT JUMP TEST",
         jump_type="cmj", raw_var="peak propulsive force", unit="N",
         pop_key=None, lower_is_better=False, kind="info"),
    dict(key="cmj_contraction_time", label="CMJ Contraction Time", category="COUNTERMOVEMENT JUMP TEST",
         jump_type="cmj", raw_var="time to takeoff", unit="s",
         pop_key=None, lower_is_better=True, kind="info"),
    dict(key="cmj_net_rel_impulse", label="CMJ Net Rel Impulse", category="COUNTERMOVEMENT JUMP TEST",
         jump_type="cmj", raw_var=None,
         derive=lambda rep: _safe_div(rep.get("net impulse"), rep.get("body mass")),
         unit="N\u00b7s/kg", pop_key="cmj_net_rel_impulse", lower_is_better=False, kind="info"),
    dict(key="cmj_mean_power", label="CMJ Mean Power", category="COUNTERMOVEMENT JUMP TEST",
         jump_type="cmj", raw_var="avg. propulsive power", unit="W",
         pop_key=None, lower_is_better=False, kind="info"),
    dict(key="mrsi_cmj", label="mRSI-CMJ", category="COUNTERMOVEMENT JUMP TEST",
         jump_type="cmj", raw_var="rsi modified", unit="m/s",
         pop_key="mrsi_cmj", lower_is_better=False, kind="score"),

    # COUNTERMOVEMENT JUMP REBOUND (niente FORCE e POWER)
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

    # kind="support": servono ai controlli a soglia del CMJ RE (CHECKS).
    # Compaiono nella tabella dati del Dettaglio Test (riga evidenziata) ma
    # senza T-score: non entrano nel profilo, nei grafici T-score né nei report.
    dict(key="cmj_re_initial_height", label="CMJ RE Jump Height (iniziale)", category="COUNTERMOVEMENT JUMP REBOUND TEST",
         jump_type="cmrj", raw_var="jump height ft", unit="cm",
         pop_key=None, lower_is_better=False, kind="support"),
    dict(key="unbalanced_landing_raw", label="Braking Impulse Sym. Index (CMJ RE)", category="COUNTERMOVEMENT JUMP REBOUND TEST",
         jump_type="cmrj", raw_var="braking impulse sym. index", unit="%",
         pop_key=None, lower_is_better=False, kind="support"),

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
# ripetizione, con le stesse esclusioni già impostate sui salti.
# Nomi già lowercase per combaciare con le chiavi salvate in rep["vars"]
# (il parsing normalizza sempre con .strip().lower()).
_SJ_METRICS_CATALOG = [
    'avg landing rfd', 'avg landing rfd sym. index', 'avg propulsive force sym. index',
    'avg propulsive power sym. index', 'avg. propulsive force', 'avg. propulsive power',
    'avg. propulsive velocity', 'avg. rfd', 'avg. rfd sym. index', 'body mass', 'body weight',
    'body weight sd', 'contact time', 'flight threshold', 'flight time', 'force peak power',
    'initiation threshold', 'jump height ft', 'jump height ni', 'jump momentum', 'jump start time',
    'jump threshold time', 'landing peak force time', 'landing rfd 0-20ms',
    'landing rfd 0-20ms sym. index', 'landing rfd 0-40ms', 'landing rfd 0-40ms sym. index',
    'landing rfd 0-60ms', 'landing rfd 0-60ms sym. index', 'landing rfd 0-80ms',
    'landing rfd 0-80ms sym. index', 'landing time', 'left avg landing rfd',
    'left avg propulsive force', 'left avg propulsive power', 'left avg rfd',
    'left landing rfd 0-20ms', 'left landing rfd 0-40ms', 'left landing rfd 0-60ms',
    'left landing rfd 0-80ms', 'left net impulse', 'left p1 impulse', 'left p2 impulse',
    'left peak force', 'left peak landing force', 'left peak propulsive force',
    'left peak propulsive power', 'left propulsive impulse', 'left propulsive rfd',
    'left propulsive work', 'left time to peak force', 'left time to peak landing force',
    'left time to peak power', 'left time to peak propulsive force',
    'left time to peak propulsive power', 'net impulse', 'net impulse sym. index', 'p1 avg force',
    'p1 avg power', 'p1 avg velocity', 'p1 duration', 'p1 impulse', 'p1 impulse sym. index',
    'p1 p2 duration ratio', 'p1 p2 force ratio', 'p1 p2 power ratio', 'p1 p2 velocity ratio',
    'p1 peak force', 'p1 peak power', 'p1 peak velocity', 'p2 avg force', 'p2 avg power',
    'p2 avg velocity', 'p2 duration', 'p2 impulse', 'p2 impulse sym. index', 'p2 peak force',
    'p2 peak power', 'p2 peak velocity', 'peak force', 'peak force sym. index', 'peak force time',
    'peak landing force', 'peak landing force sym. index', 'peak power', 'peak propulsive force',
    'peak propulsive force sym. index', 'peak propulsive power',
    'peak propulsive power sym. index', 'peak propulsive velocity', 'peak velocity',
    'propulsive duration', 'propulsive impulse', 'propulsive impulse sym. index', 'propulsive rfd',
    'propulsive rfd sym. index', 'propulsive start time', 'propulsive work',
    'propulsive work sym. index', 'rel. avg. propulsive force', 'rel. propulsive impulse',
    'relative peak force', 'relative peak landing force', 'relative peak power',
    'right avg landing rfd', 'right avg propulsive force', 'right avg propulsive power',
    'right avg rfd', 'right landing rfd 0-20ms', 'right landing rfd 0-40ms',
    'right landing rfd 0-60ms', 'right landing rfd 0-80ms', 'right net impulse',
    'right p1 impulse', 'right p2 impulse', 'right peak force', 'right peak landing force',
    'right peak propulsive force', 'right peak propulsive power', 'right propulsive impulse',
    'right propulsive rfd', 'right propulsive work', 'right time to peak force',
    'right time to peak landing force', 'right time to peak power',
    'right time to peak propulsive force', 'right time to peak propulsive power', 'rsi',
    'rsi modified', 'takeoff time', 'takeoff velocity', 'time to peak force',
    'time to peak landing force', 'time to peak landing force sym. index', 'time to peak power',
    'time to peak power sym. index', 'time to peak propulsive force',
    'time to peak propulsive force sym. index', 'time to peak propulsive power',
    'time to peak propulsive power sym. index', 'time to peak si', 'time to takeoff',
    'velocity peak power',
]

_CMJ_EXTRA_METRICS_CATALOG = [
    'avg braking force sym. index', 'avg braking power sym. index', 'avg eccentric force',
    'avg eccentric force sym. index', 'avg eccentric power', 'avg eccentric power sym. index',
    'avg eccentric velocity', 'avg. braking force', 'avg. braking power', 'avg. braking velocity',
    'braking duration', 'braking end time', 'braking impulse', 'braking impulse sym. index',
    'braking rfd', 'braking rfd sym. index', 'braking work', 'braking work sym. index',
    'decel rfd', 'decel rfd sym. index', 'displacement depth', 'eccentric impulse',
    'eccentric impulse sym. index', 'eccentric rfd', 'eccentric rfd sym. index', 'eccentric work',
    'eccentric work sym. index', 'force at min displacement', 'left avg braking force',
    'left avg braking power', 'left avg eccentric force', 'left avg eccentric power',
    'left braking impulse', 'left braking rfd', 'left braking work', 'left decel rfd',
    'left eccentric impulse', 'left eccentric rfd', 'left eccentric work',
    'left peak braking force', 'left peak braking power', 'left peak eccentric force',
    'left peak eccentric power', 'left time to peak braking force',
    'left time to peak braking power', 'left time to peak eccentric force',
    'left time to peak eccentric power', 'min braking velocity', 'min eccentric velocity',
    'min unweight force', 'peak braking force', 'peak braking force sym. index',
    'peak braking power', 'peak braking power sym. index', 'peak eccentric force',
    'peak eccentric force sym. index', 'peak eccentric power', 'peak eccentric power sym. index',
    'rel. min unweight force', 'relative force at min displacement', 'right avg braking force',
    'right avg braking power', 'right avg eccentric force', 'right avg eccentric power',
    'right braking impulse', 'right braking rfd', 'right braking work', 'right decel rfd',
    'right eccentric impulse', 'right eccentric rfd', 'right eccentric work',
    'right peak braking force', 'right peak braking power', 'right peak eccentric force',
    'right peak eccentric power', 'right time to peak braking force',
    'right time to peak braking power', 'right time to peak eccentric force',
    'right time to peak eccentric power', 'rsi exponential', 'time to peak braking force',
    'time to peak braking force sym. index', 'time to peak braking power',
    'time to peak braking power sym. index', 'time to peak eccentric force',
    'time to peak eccentric force sym. index', 'time to peak eccentric power',
    'time to peak eccentric power sym. index', 'unweighted duration', 'with armswing',
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

# Elenco per jump_type usato dal multiselect di ricerca. SJ non include le
# metriche della fase eccentrica/di contromovimento (non presenti in un
# export SJ reale); CMJ le aggiunge; CMJ RE aggiunge anche le variabili
# "rebound ...".
EXTRA_METRICS_CATALOG = {
    "sj": sorted(_SJ_METRICS_CATALOG),
    "cmj": sorted(set(_SJ_METRICS_CATALOG) | set(_CMJ_EXTRA_METRICS_CATALOG)),
    "cmrj": sorted(set(_SJ_METRICS_CATALOG) | set(_CMJ_EXTRA_METRICS_CATALOG) | set(_CMRJ_EXTRA_METRICS_CATALOG)),
    "imtp": sorted(_IMTP_METRICS_CATALOG),
}


def extra_metric_label(raw_var):
    """Etichetta leggibile per una metrica extra scelta dall'utente (solo
    la prima lettera maiuscola, il resto è già in minuscolo)."""
    return raw_var[:1].upper() + raw_var[1:]


# ============================================================================
# PARTE 1ter — DESCRIZIONI DELLE METRICHE (help "?")
# ============================================================================
# Testi tratti da ccathletics.dk/what-we-measure/ (glossario ufficiale
# ForceMate/ForceDecks), usati per l'help "?" mostrato accanto a ogni
# metrica in Dettaglio Test e nel report. Se una metrica non ha una
# descrizione nota, semplicemente non mostra alcun help (nessuna icona).

# Metriche generiche di salto (SJ/CMJ/CMJ RE): stessa fase, stesso nome,
# stessa definizione a prescindere dal tipo di test.
_CORE_METRIC_DESCRIPTIONS = {
    "jump height ft": "Altezza del salto calcolata dal tempo di volo (Flight Time).",
    "jump height ni": "Altezza del salto calcolata dall'impulso netto (Net Impulse).",
    "takeoff velocity": "Velocità al distacco da terra.",
    "rsi": "Tempo di volo / tempo di contatto (DJ). (Time-to-Takeoff per CMJ/SJ).",
    "rsi modified": "Altezza del salto / tempo di contatto (DJ). (Time-to-Takeoff per CMJ/SJ). Usa l'impulso netto quando disponibile.",
    "rsi exponential": "Formula proprietaria CC Athletics per l'RSI nei Drop Jump: Tempo di volo\u00b2 / Tempo di contatto.",
    "jump momentum": "Quantità di moto generata durante il salto, calcolata come prodotto tra massa corporea e velocità al distacco.",
    "displacement depth": "Profondità massima raggiunta durante la fase di contromovimento, cioè quanto scende il centro di massa prima dell'inizio della fase propulsiva.",
    "net impulse": "Impulso totale (al netto del peso corporeo) applicato durante le fasi di frenata+propulsione.",
    "eccentric impulse": "Impulso netto sull'intera fase eccentrica (scarico + frenata). Poiché l'atleta parte e finisce questa fase da fermo, le due sotto-fasi tendono ad annullarsi a vicenda: valori vicini a zero sono normali.",
    "propulsive impulse": "Impulso netto (forza meno peso corporeo) generato durante la fase propulsiva (dal punto più basso al distacco). Rappresenta l'integrale forza-tempo che accelera il corpo verso l'alto. Equivalente al 'Concentric Impulse' di altri sistemi.",
    "rel. propulsive impulse": "Impulso durante la fase propulsiva relativo alla massa corporea.",
    "braking impulse": "Impulso netto (forza meno peso corporeo) generato solo durante la fase di frenata/decelerazione (dal picco di velocità eccentrica al punto più basso). Quantifica quanto efficacemente l'atleta decelera dopo aver raggiunto la velocità massima verso il basso.",
    "peak force": "Picco di forza istantanea durante l'intero salto.",
    "min unweight force": "Forza minima misurata durante la fase di scarico (unweighting).",
    "rel. min unweight force": "Forza minima durante lo scarico, espressa come multiplo del peso corporeo.",
    "peak braking force": "Picco di forza istantanea durante la sola fase di frenata/decelerazione (dal picco di velocità eccentrica al punto più basso). Per l'intera fase eccentrica, vedi 'Peak Eccentric Force'.",
    "avg. braking force": "Forza media durante la sola fase di frenata/decelerazione (dal picco di velocità eccentrica al punto più basso). Per l'intera fase eccentrica, vedi 'Avg Eccentric Force'.",
    "avg braking force": "Forza media durante la sola fase di frenata/decelerazione (dal picco di velocità eccentrica al punto più basso). Per l'intera fase eccentrica, vedi 'Avg Eccentric Force'.",
    "avg eccentric force": "Forza media durante l'intera fase eccentrica (scarico + frenata). Rappresenta la forza media dall'inizio del movimento al punto più basso del contromovimento.",
    "peak eccentric force": "Picco di forza istantanea durante l'intera fase eccentrica (scarico + frenata). Rappresenta la forza massima dall'inizio del movimento al punto più basso del contromovimento.",
    "avg. propulsive force": "Forza media durante la fase propulsiva.",
    "avg propulsive force": "Forza media durante la fase propulsiva.",
    "rel. avg. propulsive force": "Forza media durante la fase propulsiva, relativa alla massa corporea.",
    "peak propulsive force": "Picco di forza istantanea durante la fase propulsiva.",
    "relative peak force": "Picco di forza espresso come multiplo del peso corporeo.",
    "rel. peak force": "Picco di forza espresso come multiplo del peso corporeo.",
    "force peak power": "Forza istantanea nel momento di picco potenza (sull'intero dataset).",
    "force at min displacement": "Forza totale (di reazione al suolo) misurata nel punto di massima profondità di spostamento durante il contromovimento.",
    "relative force at min displacement": "Forza al punto di minimo spostamento, espressa come multiplo del peso corporeo.",
    "peak braking power": "Picco di potenza istantanea durante la sola fase di frenata/decelerazione (valore negativo). Per l'intera fase eccentrica, vedi 'Peak Eccentric Power'.",
    "avg. braking power": "Potenza meccanica media durante la sola fase di frenata/decelerazione (valore negativo). Per l'intera fase eccentrica, vedi 'Avg Eccentric Power'.",
    "avg braking power": "Potenza meccanica media durante la sola fase di frenata/decelerazione (valore negativo). Per l'intera fase eccentrica, vedi 'Avg Eccentric Power'.",
    "braking work": "Lavoro totale svolto durante la sola fase di frenata/decelerazione (valore negativo). Per l'intera fase eccentrica, vedi 'Eccentric Work'.",
    "peak eccentric power": "Picco di potenza istantanea durante l'intera fase eccentrica (scarico + frenata). Riportato come valore assoluto. Valori alti indicano una buona capacità di potenza eccentrica.",
    "avg eccentric power": "Potenza meccanica media durante l'intera fase eccentrica (scarico + frenata). Riportata come valore assoluto. Rappresenta il tasso medio di assorbimento di energia nella fase discendente.",
    "eccentric work": "Lavoro totale svolto durante l'intera fase eccentrica (scarico + frenata). Riportato come valore assoluto. Quantifica l'energia totale assorbita dall'inizio del movimento al punto più basso.",
    "peak propulsive power": "Potenza massima durante la fase propulsiva.",
    "avg. propulsive power": "Potenza media durante la fase propulsiva.",
    "avg propulsive power": "Potenza media durante la fase propulsiva.",
    "propulsive work": "Lavoro totale svolto durante la fase propulsiva.",
    "peak power": "Picco di potenza istantanea durante l'intero salto.",
    "relative peak power": "Picco di potenza espresso in watt per chilogrammo di peso corporeo.",
    "min braking velocity": "Velocità minima durante la fase di frenata.",
    "avg. braking velocity": "Velocità media durante la fase di frenata.",
    "avg braking velocity": "Velocità media durante la fase di frenata.",
    "min eccentric velocity": "Velocità minima (più negativa) durante l'intera fase eccentrica \u2014 il picco di velocità verso il basso (Eccentric Peak Velocity, EPV), che segna il passaggio dallo scarico alla frenata.",
    "avg eccentric velocity": "Velocità media durante l'intera fase eccentrica (scarico + frenata).",
    "peak propulsive velocity": "Velocità massima raggiunta durante la fase propulsiva.",
    "avg. propulsive velocity": "Velocità media durante la fase propulsiva.",
    "avg propulsive velocity": "Velocità media durante la fase propulsiva.",
    "peak velocity": "Picco di velocità istantanea durante la fase di propulsione.",
    "velocity peak power": "Velocità istantanea nel momento di picco potenza (sull'intero dataset).",
    "avg. rfd": "Tasso medio di sviluppo della forza (Rate of Force Development, RFD).",
    "avg rfd": "Tasso medio di sviluppo della forza (Rate of Force Development, RFD).",
    "braking rfd": "Pendenza media della forza durante la fase di frenata.",
    "decel rfd": "Tasso di sviluppo della forza nella porzione di decelerazione della fase di frenata, misurato dalla velocità minima alla fine della frenata. Simile a 'Braking RFD' o 'Load' di altri sistemi, ma focalizzato sulla fase tardiva della frenata.",
    "eccentric rfd": "Tasso di sviluppo della forza durante la fase eccentrica, calcolato dal valore iniziale al picco di forza.",
    "propulsive rfd": "Tasso di sviluppo della forza durante la fase propulsiva, calcolato dal valore iniziale al picco di forza.",
    "p1 impulse": "Impulso generato durante la prima metà della fase propulsiva (posizione iniziale a triplice flessione). Rappresenta il primo 50% dell'impulso propulsivo.",
    "p2 impulse": "Impulso generato durante la seconda metà della fase propulsiva (fino alla triplice estensione). Rappresenta l'ultimo 50% dell'impulso propulsivo.",
    "p1 duration": "Durata della prima metà della fase propulsiva.",
    "p2 duration": "Durata della seconda metà della fase propulsiva.",
    "p1 avg force": "Forza media durante la prima metà della fase propulsiva (posizione di squat profondo).",
    "p2 avg force": "Forza media durante la seconda metà della fase propulsiva (estensione fino al distacco).",
    "p1 p2 force ratio": "Rapporto tra la forza media in P1 e quella in P2. Sopra 1.0 = produzione di forza maggiore nella posizione più profonda; sotto 1.0 = maggiore durante l'estensione.",
    "p1 p2 duration ratio": "Rapporto tra la durata di P1 e quella di P2. Sopra 1.0 = più tempo nella prima metà dell'impulso; sotto 1.0 = fase iniziale più rapida.",
    "p1 peak force": "Picco di forza durante la fase propulsiva iniziale (P1).",
    "p1 peak velocity": "Picco di velocità durante la fase propulsiva iniziale (P1).",
    "p1 avg velocity": "Velocità media durante la fase propulsiva iniziale (P1).",
    "p1 peak power": "Picco di potenza durante la fase propulsiva iniziale (P1).",
    "p1 avg power": "Potenza media durante la fase propulsiva iniziale (P1).",
    "p2 peak force": "Picco di forza durante la fase propulsiva finale (P2).",
    "p2 peak velocity": "Picco di velocità durante la fase propulsiva finale (P2).",
    "p2 avg velocity": "Velocità media durante la fase propulsiva finale (P2).",
    "p2 peak power": "Picco di potenza durante la fase propulsiva finale (P2).",
    "p2 avg power": "Potenza media durante la fase propulsiva finale (P2).",
    "p1 p2 velocity ratio": "Rapporto tra la velocità media di P1 e P2, indica le caratteristiche di velocità tra fase iniziale e finale.",
    "p1 p2 power ratio": "Rapporto tra la potenza media di P1 e P2, indica le caratteristiche di potenza tra fase iniziale e finale.",
    "peak landing force": "Forza massima registrata durante la fase di atterraggio del salto. Misura la capacità di assorbire le forze in atterraggio.",
    "relative peak landing force": "Picco di forza in atterraggio, espresso come multiplo del peso corporeo.",
    "time to peak landing force": "Tempo dal contatto in atterraggio al picco di forza in atterraggio. Indica quanto rapidamente si raggiungono i picchi di forza durante l'atterraggio.",
    "landing peak force time": "Tempo dal contatto in atterraggio al picco di forza in atterraggio. Indica quanto rapidamente si raggiungono i picchi di forza durante l'atterraggio.",
    "avg landing rfd": "Tasso medio di sviluppo della forza in atterraggio, dal peso corporeo al picco di forza in atterraggio. Indica la rigidità in atterraggio.",
    "landing rfd 0-20ms": "Tasso di sviluppo della forza nei primi 20ms dopo l'atterraggio.",
    "landing rfd 0-40ms": "Tasso di sviluppo della forza nei primi 40ms dopo l'atterraggio.",
    "landing rfd 0-60ms": "Tasso di sviluppo della forza nei primi 60ms dopo l'atterraggio.",
    "landing rfd 0-80ms": "Tasso di sviluppo della forza nei primi 80ms dopo l'atterraggio.",
    "time to peak force": "Tempo impiegato per raggiungere il picco di forza istantanea (dall'inizio del salto).",
    "peak force time": "Tempo impiegato per raggiungere il picco di forza istantanea (dall'inizio del salto).",
    "unweighted duration": "Tempo impiegato per completare la fase di scarico (unweighting).",
    "braking duration": "Durata della sola fase di frenata/decelerazione, misurata dal picco di velocità eccentrica (EPV) al punto più basso del contromovimento.",
    "propulsive duration": "Durata dalla fine della fase di frenata (quando la velocità diventa positiva) fino al distacco. A volte chiamato 'Concentric Time'.",
    "flight time": "Tempo trascorso in fase di volo.",
    "time to peak power": "Tempo dall'inizio del salto al momento di picco potenza.",
    "contact time": "Tempo di contatto con il suolo tra atterraggio e distacco.",
    "contraction time": "Tempo totale di contrazione muscolare dall'inizio del movimento al distacco. Per il CMJ: fasi eccentrica + concentrica. Per lo SJ: solo la fase propulsiva (concentrica).",
    "time to takeoff": "Tempo dall'inizio del movimento al distacco. Indica il tempo totale di contatto col suolo prima del decollo.",
    "takeoff time": "Tempo dall'inizio del movimento al distacco. Indica il tempo totale di contatto col suolo prima del decollo.",
    "time to peak braking force": "Tempo per raggiungere il picco di forza durante la fase di frenata, misurato dall'inizio della frenata.",
    "time to peak braking power": "Tempo per raggiungere il picco di potenza durante la fase di frenata, misurato dall'inizio della frenata.",
    "time to peak eccentric force": "Tempo per raggiungere il picco di forza durante la fase eccentrica, misurato dall'inizio del salto.",
    "time to peak eccentric power": "Tempo per raggiungere il picco di potenza durante la fase eccentrica, misurato dall'inizio del salto.",
    "time to peak propulsive force": "Tempo per raggiungere il picco di forza durante la fase propulsiva, misurato dall'inizio della propulsione.",
    "time to peak propulsive power": "Tempo per raggiungere il picco di potenza durante la fase propulsiva, misurato dall'inizio della propulsione.",
    "system weight": "Peso misurato dell'atleta durante la fase di pesata, appena prima del salto.",
    "body mass": "Massa corporea dell'atleta, misurata prima del test.",
    "body weight": "Peso corporeo dell'atleta registrato per questo test.",
    "body weight sd": "Deviazione standard del peso corporeo durante la fase di pesata (indica quanto l'atleta è rimasto fermo).",
    "propulsive start time": "Momento in cui inizia la fase propulsiva (quando la forza passa da frenata a propulsiva).",
}

# Descrizioni specifiche IMTP (Isometric Mid-Thigh Pull): nomi e definizioni
# non coincidono con quelli dei test di salto, quindi hanno priorità quando
# jump_type == "imtp".
_IMTP_METRIC_DESCRIPTIONS = {
    "peak force": "Picco di forza (valore di forza più alto durante la contrazione volontaria massimale).",
    "peak force net": "Picco di forza netta (valore di forza più alto durante la contrazione volontaria massimale).",
    "peak force relative": "Picco di forza espresso come percentuale del peso corporeo.",
    "peak force relative net": "Picco di forza netta espresso come percentuale del peso corporeo.",
    "onset force": "Forza al momento dell'inizio (onset) della contrazione.",
    "onset force net": "Forza netta al momento dell'inizio (onset) della contrazione.",
    "onset force relative": "Forza all'inizio della contrazione, espressa come percentuale del peso corporeo.",
    "onset force relative net": "Forza netta all'inizio della contrazione, espressa come percentuale del peso corporeo.",
    "onset torque": "Coppia (torque) al momento dell'inizio della contrazione.",
    "onset time": "Istante di inizio della contrazione (dall'inizio del segnale).",
    "end of rise force": "Forza raggiunta nel punto in cui la derivata forza-tempo si stabilizza per la prima volta (fine della fase di salita).",
    "end of rise force net": "Forza netta raggiunta nel punto in cui la derivata forza-tempo si stabilizza per la prima volta.",
    "end of rise force relative": "Forza a fine salita, espressa come percentuale del peso corporeo.",
    "end of rise force relative net": "Forza netta a fine salita, espressa come percentuale del peso corporeo.",
    "end of rise torque": "Coppia raggiunta nel punto in cui la derivata coppia-tempo si stabilizza per la prima volta.",
    "force at max rfd": "Forza nel punto in cui si raggiunge l'RFD massimo.",
    "force at max rfd net": "Forza netta nel punto in cui si raggiunge l'RFD massimo.",
    "force at max rfd relative": "Forza nel punto di RFD massimo, espressa come percentuale del peso corporeo.",
    "force at max rfd relative net": "Forza netta nel punto di RFD massimo, espressa come percentuale del peso corporeo.",
    "max rfd": "Picco del tasso di sviluppo della forza (Rate of Force Development) durante la contrazione.",
    "impulse at max rfd": "Impulso dall'inizio della contrazione al punto di RFD massimo.",
    "rfd impulse": "Impulso accumulato fino al punto di massimo tasso di sviluppo della forza.",
    "steadiness rsme force": "Errore quadratico medio (RMSE) della forza durante la fase di stazionarietà (steady state).",
    "steadiness rsme rfd": "Errore quadratico medio (RMSE) dell'RFD durante la fase di stazionarietà (steady state).",
    "time to force end of rise": "Tempo dall'inizio della contrazione alla fine della fase di salita (End Of Rise).",
    "time to peak force": "Tempo dall'inizio della contrazione al picco di forza durante la contrazione volontaria massimale.",
    "time to peak rfd": "Tempo dall'inizio della contrazione al picco di RFD.",
    "peak torque": "Picco di coppia (valore di coppia più alto durante la contrazione volontaria massimale).",
    "torque at max rfd": "Coppia nel punto in cui si raggiunge l'RFD massimo.",
    "system weight": "Peso corporeo misurato sulla pedana, appena prima della trazione.",
}
for _ms in (50, 100, 150, 200, 250):
    _IMTP_METRIC_DESCRIPTIONS[f"force {_ms}ms"] = f"Forza a {_ms}ms dall'inizio della contrazione."
    _IMTP_METRIC_DESCRIPTIONS[f"force {_ms}ms net"] = f"Forza netta a {_ms}ms dall'inizio della contrazione."
    _IMTP_METRIC_DESCRIPTIONS[f"force {_ms}ms relative"] = f"Forza a {_ms}ms dall'inizio della contrazione, come percentuale del peso corporeo."
    _IMTP_METRIC_DESCRIPTIONS[f"force {_ms}ms relative net"] = f"Forza netta a {_ms}ms dall'inizio della contrazione, come percentuale del peso corporeo."
    _IMTP_METRIC_DESCRIPTIONS[f"rfd {_ms}ms"] = f"Tasso di sviluppo della forza (RFD) a {_ms}ms dall'inizio della contrazione."
    _IMTP_METRIC_DESCRIPTIONS[f"impulse {_ms}ms"] = f"Impulso dall'inizio della contrazione a {_ms}ms."
    _IMTP_METRIC_DESCRIPTIONS[f"torque {_ms}ms"] = f"Coppia a {_ms}ms dall'inizio della contrazione."

# Metriche "rebound ..." (CMJ RE) documentate esplicitamente sul sito con
# testo proprio (non solo "stessa metrica in fase di rimbalzo").
_REBOUND_METRIC_OVERRIDES = {
    "rebound jump height ft": "Altezza del salto di rimbalzo, calcolata dal tempo di volo.",
    "rebound contact time": "Durata del contatto a terra tra l'atterraggio dal CMJ e il distacco del salto di rimbalzo.",
    "rebound rsi": "Reactive Strength Index del salto di rimbalzo (altezza del salto \u00f7 tempo di contatto).",
    "rebound rsi modified": "Reactive Strength Index modificato del salto di rimbalzo (altezza del salto \u00f7 tempo di contatto, formulazione modificata).",
    "rebound jump momentum": "Quantità di moto verticale al distacco del salto di rimbalzo (massa corporea \u00d7 velocità al distacco).",
    "rebound braking impulse": "Impulso totale generato durante la fase di frenata del rimbalzo.",
    "rebound propulsive impulse": "Impulso totale generato durante la fase propulsiva del rimbalzo.",
    "rebound force at min displacement": "Forza verticale nel punto più basso di spostamento del centro di massa durante il contatto di rimbalzo.",
    "rebound relative force at min displacement": "Forza al punto di minimo spostamento durante il rimbalzo, normalizzata sul peso corporeo.",
    "rebound avg braking force": "Forza verticale media durante la fase di frenata del rimbalzo.",
    "rebound peak braking force": "Picco di forza verticale durante la fase di frenata del rimbalzo.",
    "rebound avg propulsive force": "Forza verticale media durante la fase propulsiva del rimbalzo.",
    "rebound peak propulsive force": "Picco di forza verticale durante la fase propulsiva del rimbalzo.",
    "rebound avg braking power": "Potenza media durante la fase di frenata del rimbalzo.",
    "rebound peak braking power": "Picco di potenza durante la fase di frenata del rimbalzo.",
    "rebound avg propulsive power": "Potenza media durante la fase propulsiva del rimbalzo.",
    "rebound peak propulsive power": "Picco di potenza durante la fase propulsiva del rimbalzo.",
    "rebound braking work": "Lavoro totale svolto durante la fase di frenata del rimbalzo.",
    "rebound propulsive work": "Lavoro totale svolto durante la fase propulsiva del rimbalzo.",
    "rebound min braking velocity": "Velocità minima durante la fase di frenata del rimbalzo.",
    "rebound avg braking velocity": "Velocità media durante la fase di frenata del rimbalzo.",
    "rebound peak propulsive velocity": "Velocità massima raggiunta durante la fase propulsiva del rimbalzo.",
    "rebound avg propulsive velocity": "Velocità media durante la fase propulsiva del rimbalzo.",
    "rebound avg braking rfd": "Pendenza media della forza durante la fase di frenata del rimbalzo.",
    "rebound propulsive rfd": "Tasso di sviluppo della forza durante la fase propulsiva del rimbalzo, calcolato dal valore iniziale al picco di forza.",
    "rebound p1 impulse": "Impulso durante la prima metà della fase propulsiva del rimbalzo.",
    "rebound p2 impulse": "Impulso durante la seconda metà della fase propulsiva del rimbalzo.",
    "rebound p1 duration": "Durata della prima metà della fase propulsiva del rimbalzo.",
    "rebound p1 peak force": "Picco di forza durante la fase propulsiva iniziale del rimbalzo (P1).",
    "rebound p1 avg force": "Forza media durante la prima metà della fase propulsiva del rimbalzo.",
    "rebound p1 peak velocity": "Picco di velocità durante la fase propulsiva iniziale del rimbalzo (P1).",
    "rebound p1 avg velocity": "Velocità media durante la fase propulsiva iniziale del rimbalzo (P1).",
    "rebound p1 peak power": "Picco di potenza durante la fase propulsiva iniziale del rimbalzo (P1).",
    "rebound p1 avg power": "Potenza media durante la fase propulsiva iniziale del rimbalzo (P1).",
    "rebound p2 duration": "Durata della seconda metà della fase propulsiva del rimbalzo.",
    "rebound p2 peak force": "Picco di forza durante la fase propulsiva finale del rimbalzo (P2).",
    "rebound p2 avg force": "Forza media durante la seconda metà della fase propulsiva del rimbalzo.",
    "rebound p2 peak velocity": "Picco di velocità durante la fase propulsiva finale del rimbalzo (P2).",
    "rebound p2 avg velocity": "Velocità media durante la fase propulsiva finale del rimbalzo (P2).",
    "rebound p2 peak power": "Picco di potenza durante la fase propulsiva finale del rimbalzo (P2).",
    "rebound p2 avg power": "Potenza media durante la fase propulsiva finale del rimbalzo (P2).",
    "rebound p1 p2 force ratio": "Rapporto tra la forza media in P1 e in P2 del rimbalzo.",
    "rebound p1 p2 velocity ratio": "Rapporto tra la velocità media di P1 e P2 del rimbalzo.",
    "rebound p1 p2 power ratio": "Rapporto tra la potenza media di P1 e P2 del rimbalzo.",
    "rebound p1 p2 duration ratio": "Rapporto tra la durata di P1 e la durata di P2 del rimbalzo.",
    "rebound flight time": "Durata della fase di volo del salto di rimbalzo.",
    "rebound time to peak braking force": "Tempo dall'inizio del contatto di rimbalzo al picco di forza in frenata.",
    "rebound time to takeoff": "Tempo dall'inizio del contatto di rimbalzo al distacco del rimbalzo.",
    "rebound time to peak braking power": "Tempo per raggiungere il picco di potenza durante la fase di frenata del rimbalzo, misurato dall'inizio della frenata.",
    "rebound time to peak propulsive force": "Tempo per raggiungere il picco di forza durante la fase propulsiva del rimbalzo, misurato dall'inizio della propulsione.",
    "rebound time to peak propulsive power": "Tempo per raggiungere il picco di potenza durante la fase propulsiva del rimbalzo, misurato dall'inizio della propulsione.",
}

# Descrizioni manuali per le metriche derivate dall'app stessa (nessun
# raw_var diretto, quindi non ricavabili dal sito).
_DERIVED_METRIC_DESCRIPTIONS = {
    "imtp_rel_peak_force": "IMTP Peak Force diviso per la massa corporea (N/kg): permette di confrontare atleti di taglia diversa.",
    "sj_net_rel_impulse": "SJ Net Impulse diviso per la massa corporea (N\u00b7s/kg).",
    "cmj_net_rel_impulse": "CMJ Net Impulse diviso per la massa corporea (N\u00b7s/kg).",
    "dsi": "Dynamic Strength Index: rapporto tra CMJ Peak Propulsive Force e IMTP Peak Force. Indica quanta della forza massima isometrica viene espressa durante un gesto balistico come il CMJ. Non è un indice 'più alto = meglio': entrambi gli estremi descrivono un profilo diverso, non uno migliore.",
    "eur": "Eccentric Utilisation Ratio: rapporto tra CMJ Height e SJ Height. Indica quanto la fase eccentrica (contromovimento) contribuisce alla prestazione rispetto a un salto puramente concentrico. Da leggere sempre insieme ai valori assoluti: due atleti con lo stesso EUR possono avere prestazioni molto diverse.",
}


def get_metric_description(jump_type, raw_var):
    """Cerca la descrizione di una metrica (per l'help "?"), con fallback
    su prefissi lato sinistro/destro, fase di rimbalzo (rebound) e indice
    di simmetria. Ritorna None se non trovata (nessun help, per design)."""
    if not raw_var:
        return None
    key = raw_var.strip().lower()

    if jump_type == "imtp" and key in _IMTP_METRIC_DESCRIPTIONS:
        return _IMTP_METRIC_DESCRIPTIONS[key]
    if key in _REBOUND_METRIC_OVERRIDES:
        return _REBOUND_METRIC_OVERRIDES[key]
    if key in _CORE_METRIC_DESCRIPTIONS:
        return _CORE_METRIC_DESCRIPTIONS[key]

    # Nell'IMTP "left"/"right" compare come token in mezzo al nome (es.
    # "force 100ms left net"), non come prefisso: lo rimuoviamo e cerchiamo
    # la base risultante.
    if jump_type == "imtp":
        for side_label, token in (("lato sinistro", "left"), ("lato destro", "right")):
            if re.search(rf"\b{token}\b", key):
                candidate = re.sub(rf"\b{token}\b", "", key)
                candidate = re.sub(r"\s+", " ", candidate).strip()
                if candidate in _IMTP_METRIC_DESCRIPTIONS:
                    return f"{_IMTP_METRIC_DESCRIPTIONS[candidate]} ({side_label})"

    for side_label, prefix in (("lato sinistro", "left "), ("lato destro", "right ")):
        if key.startswith(prefix):
            base_desc = get_metric_description(jump_type, key[len(prefix):])
            if base_desc:
                return f"{base_desc} ({side_label})"

    if key.startswith("rebound "):
        base_desc = get_metric_description(jump_type, key[len("rebound "):])
        if base_desc:
            return f"{base_desc} (fase di rimbalzo / rebound)"

    for suffix in (" sym. index", " symmetry index", " si"):
        if key.endswith(suffix):
            base_desc = get_metric_description(jump_type, key[: -len(suffix)].strip())
            if base_desc:
                return f"Indice di simmetria sinistra-destra per: {base_desc}"

    return None


# Attacca la descrizione a ciascuna metrica curata (usata per l'help "?"
# nella scheda Dettaglio Test e nel report). Fatto qui, dopo la definizione
# di get_metric_description, per evitare di duplicare il testo nel
# dizionario METRICS stesso.
for _m in METRICS:
    if _m.get("raw_var"):
        _m["desc"] = get_metric_description(_m.get("jump_type"), _m["raw_var"])
    else:
        _m["desc"] = _DERIVED_METRIC_DESCRIPTIONS.get(_m["key"])