"""
Générateur de spécifications de graphiques — Agence VU.

Chaque slide de la section Performance Globale a son propre graphique,
construit DÉTERMINISTIQUEMENT depuis les KPIs et données images.
Retourne None si les données sont insuffisantes (→ slide sans graphique).

Spec retournée :
{
  "type":         "bar" | "column" | "line" | "pie",
  "title":        str,
  "categories":   [str, ...],
  "series": [
    {"name": str, "values": [float, ...], "color": "#RRGGBB"}
  ],
  "value_format": "%" | "€" | "j" | "k€" | None
}
"""
from typing import Optional

# ── Palette Agence VU ─────────────────────────────────────────────────────────
VU_BLUE   = "#008BD2"
VU_NAVY   = "#1A2E4A"
VU_GREEN  = "#27AE60"
VU_ORANGE = "#F39C12"
VU_RED    = "#E74C3C"
VU_GREY   = "#95A5A6"
VU_LIGHT  = "#BDC3C7"

# ── Benchmarks marché 2025 (méthodologie Agence VU) ───────────────────────────
BM = {
    "frequentation_j":     180.0,
    "panier_total":         40.8,
    "panier_ordos":         58.3,
    "panier_conseil":       13.89,
    "marge_etp_faible":     90_000,
    "marge_etp_correct":   105_000,
    "marge_etp_perf":      120_000,
    "rotation_ordos":       35.0,   # milieu de 30-40j
    "rotation_hors":        90.0,   # milieu de 80-100j
}


def _v(kpi_dict: dict, *keys, default=None):
    """Retourne la première valeur KPI non-nulle parmi les clés candidates."""
    for k in keys:
        entry = kpi_dict.get(k, {})
        val = entry.get("valeur") if isinstance(entry, dict) else None
        if val is not None:
            return val
    return default


def _label(kpi_dict: dict, key: str, fallback: str) -> str:
    entry = kpi_dict.get(key, {})
    return entry.get("label_fr", fallback) if isinstance(entry, dict) else fallback


# ── PG_01_SECTION — aucun graphique ──────────────────────────────────────────

# ── PG_02_PROFIL_TYPE — camembert TVA / hors ordos ───────────────────────────

def _chart_profil_type(kpi_dict: dict, **_) -> Optional[dict]:
    """Pie : part ordonnances vs hors ordonnances dans le CA."""
    pct = _v(kpi_dict, "part_ordonnances_pct", "part_tva_2_1_pct")
    if pct is None:
        return None
    # Normalise : si valeur > 1 c'est déjà un %, sinon multiplie par 100
    if pct <= 1:
        pct = pct * 100
    hors = round(100 - pct, 1)
    pct  = round(pct, 1)
    return {
        "type":         "pie",
        "title":        "Répartition CA — Ordonnances vs Hors ordonnances",
        "categories":   ["Ordonnances (TVA 2,1%)", "Hors ordonnances"],
        "series": [{"name": "CA", "values": [pct, hors], "color": VU_NAVY}],
        "value_format": "%",
    }


# ── PG_03_PROFIL_PATIENTS — pyramide des âges ────────────────────────────────

def _chart_profil_patients(kpi_dict: dict, image_results: list = None, **_) -> Optional[dict]:
    """
    Horizontal bar : % par tranche d'âge officine vs référence.
    Construit depuis les valeurs extraites des images si disponibles.
    """
    # Tentative extraction pyramide depuis images
    if image_results:
        result = _extract_age_pyramid(image_results)
        if result:
            return result
    # Fallback — pas de données âge disponibles
    return None


def _extract_age_pyramid(image_results: list) -> Optional[dict]:
    AGE_PATTERNS = ["0-4", "5-14", "5-9", "10-14", "15-24", "15-19", "20-24",
                    "25-34", "35-44", "45-54", "55-64", "65-74", "75-84",
                    "80+", "85+", "65+", "60+", "ans", "age", "âge"]
    officine, ref = {}, {}
    for img in (image_results or []):
        for entry in img.get("valeurs_extraites", []):
            label = str(entry.get("label", "")).lower().strip()
            val   = entry.get("valeur")
            if val is None:
                continue
            is_age = any(p in label for p in AGE_PATTERNS)
            if not is_age:
                continue
            is_ref = any(r in label for r in ["ref", "ville", "national", "france", "national", "moyen", "bm", "benchmark"])
            if is_ref:
                ref[label] = float(val)
            else:
                officine[label] = float(val)

    if len(officine) < 3:
        return None

    cats   = sorted(officine.keys())
    o_vals = [officine[c] for c in cats]
    series = [{"name": "Votre officine (%)", "values": o_vals, "color": VU_BLUE}]
    if len(ref) >= 3:
        r_vals = [ref.get(c, 0) for c in cats]
        series.append({"name": "Référence locale (%)", "values": r_vals, "color": VU_GREY})

    return {
        "type":         "bar",
        "title":        "Profil patients — Répartition par tranche d'âge (%)",
        "categories":   cats,
        "series":       series,
        "value_format": "%",
    }


# ── PG_04_FINANCIERS — column CA/Marge N-1 vs N ──────────────────────────────

def _chart_financiers(kpi_dict: dict, **_) -> Optional[dict]:
    """Column : CA HT et Marge Brute en k€ — N-1 reconstitué depuis évolution."""
    ca_n    = _v(kpi_dict, "ca_total", "ca_ht")
    marge_n = _v(kpi_dict, "marge_brute", "marge_globale")
    evo_ca  = _v(kpi_dict, "evolution_ca_pct", "evo_ca")
    evo_mg  = _v(kpi_dict, "evolution_marge_pct", "evo_marge")

    if ca_n is None and marge_n is None:
        return None

    def to_k(v): return round(v / 1000, 1) if v and v > 1000 else v

    if evo_ca is not None and ca_n is not None:
        ca_n1 = round(ca_n / (1 + evo_ca / 100) / 1000, 1)
        ca_n_k = round(ca_n / 1000, 1)
        ca_series = [ca_n1, ca_n_k]
        cats = ["N-1", "N"]
    elif ca_n is not None:
        ca_series = [round(ca_n / 1000, 1)]
        cats = ["N"]
    else:
        ca_series = None
        cats = ["N"]

    if evo_mg is not None and marge_n is not None:
        mg_n1 = round(marge_n / (1 + evo_mg / 100) / 1000, 1)
        mg_n_k = round(marge_n / 1000, 1)
        mg_series = [mg_n1, mg_n_k] if len(cats) == 2 else [mg_n_k]
    elif marge_n is not None:
        mg_series = [round(marge_n / 1000, 1)] * len(cats)
    else:
        mg_series = None

    series = []
    if ca_series:
        series.append({"name": "CA HT (k€)",       "values": ca_series, "color": VU_NAVY})
    if mg_series:
        series.append({"name": "Marge Brute (k€)", "values": mg_series, "color": VU_BLUE})

    if not series:
        return None

    return {
        "type":         "column",
        "title":        "CA et Marge Brute HT (k€)",
        "categories":   cats,
        "series":       series,
        "value_format": "k€",
    }


# ── PG_05_COMMERCIAUX — grouped bar indicateurs vs benchmarks ────────────────

def _chart_commerciaux(kpi_dict: dict, **_) -> Optional[dict]:
    """Grouped bar : indicateurs commerciaux officine vs benchmark marché."""
    freq    = _v(kpi_dict, "frequentation_j", "clients_j", "nb_clients_jour")
    panier  = _v(kpi_dict, "panier_moyen", "panier_total")
    conseil = _v(kpi_dict, "panier_conseil")

    cats, off_vals, bm_vals = [], [], []

    if freq is not None:
        cats.append("Fréquentation\n(clients/j)")
        off_vals.append(round(freq, 0))
        bm_vals.append(BM["frequentation_j"])

    if panier is not None:
        cats.append("Panier moyen\ntotal (€)")
        off_vals.append(round(panier, 2))
        bm_vals.append(BM["panier_total"])

    if conseil is not None:
        cats.append("Panier\nconseil (€)")
        off_vals.append(round(conseil, 2))
        bm_vals.append(BM["panier_conseil"])

    if not cats:
        return None

    return {
        "type":         "bar",
        "title":        "Indicateurs commerciaux vs Benchmark marché 2025",
        "categories":   cats,
        "series": [
            {"name": "Votre officine",        "values": off_vals, "color": VU_BLUE},
            {"name": "Benchmark marché 2025", "values": bm_vals,  "color": VU_GREY},
        ],
        "value_format": None,
    }


# ── PG_06_UNIVERS_CA — horizontal bar répartition univers ────────────────────

def _chart_univers_ca(kpi_dict: dict, **_) -> Optional[dict]:
    """Horizontal bar : CA% et Marge% par univers hors ordonnances."""
    UNIVERS = ["senior", "sénior", "jambes", "nature", "libre", "hygiene",
               "hygiène", "bébé", "bebe", "beauté", "beaute", "veto", "véto", "ortho"]

    ca_u, mg_u = {}, {}
    for kid, kpi in kpi_dict.items():
        if not isinstance(kpi, dict):
            continue
        kid_l = kid.lower()
        val   = kpi.get("valeur")
        if val is None:
            continue
        for u in UNIVERS:
            if u in kid_l:
                canon = u.replace("é", "e").capitalize()
                if any(x in kid_l for x in ["ca", "chiffre", "vente"]):
                    ca_u[canon] = val
                elif "marge" in kid_l:
                    mg_u[canon] = val

    all_u = sorted(set(list(ca_u) + list(mg_u)))
    if not all_u:
        return None

    series = []
    if ca_u:
        series.append({"name": "CA hors ordos (%)",    "values": [ca_u.get(u, 0) for u in all_u], "color": VU_BLUE})
    if mg_u:
        series.append({"name": "Marge hors ordos (%)", "values": [mg_u.get(u, 0) for u in all_u], "color": VU_NAVY})

    if not series:
        return None

    return {
        "type":         "bar",
        "title":        "Répartition CA et Marge — Univers hors ordonnances",
        "categories":   all_u,
        "series":       series,
        "value_format": "%",
    }


# ── PG_07_UNIVERS_EVO — column chart évolution % par univers ─────────────────

def _chart_univers_evo(kpi_dict: dict, **_) -> Optional[dict]:
    """Column : évolution CA% par univers N vs N-1."""
    UNIVERS = ["senior", "sénior", "jambes", "nature", "libre", "hygiene",
               "hygiène", "bébé", "bebe", "beauté", "beaute", "veto", "véto"]

    evo_u = {}
    for kid, kpi in kpi_dict.items():
        if not isinstance(kpi, dict):
            continue
        kid_l = kid.lower()
        val   = kpi.get("valeur")
        if val is None:
            continue
        if not any(x in kid_l for x in ["evo", "evolution", "croissance", "variation"]):
            continue
        for u in UNIVERS:
            if u in kid_l:
                canon = u.replace("é", "e").capitalize()
                evo_u[canon] = val

    if len(evo_u) < 2:
        return None

    cats = sorted(evo_u.keys(), key=lambda x: evo_u[x], reverse=True)
    vals = [evo_u[c] for c in cats]
    colors = [VU_GREEN if v >= 0 else VU_RED for v in vals]

    # python-pptx ne supporte pas les couleurs par barre nativement →
    # on sépare en 2 séries positives/négatives
    pos_vals = [v if v >= 0 else 0 for v in vals]
    neg_vals = [v if v <  0 else 0 for v in vals]

    series = [{"name": "Croissance (%)",  "values": pos_vals, "color": VU_GREEN}]
    if any(v < 0 for v in neg_vals):
        series.append({"name": "Recul (%)", "values": neg_vals, "color": VU_RED})

    return {
        "type":         "column",
        "title":        "Évolution CA par univers (N vs N-1, %)",
        "categories":   cats,
        "series":       series,
        "value_format": "%",
    }


# ── PG_08_TOP_MARQUES — horizontal bar top marques ───────────────────────────

def _chart_top_marques(kpi_dict: dict, **_) -> Optional[dict]:
    """Horizontal bar : top marques par marge (si données disponibles)."""
    MARQUE_KEYS = ["marque", "brand", "top_", "top10", "top 10"]
    marques = {}
    for kid, kpi in kpi_dict.items():
        if not isinstance(kpi, dict):
            continue
        kid_l = kid.lower()
        val   = kpi.get("valeur")
        label = kpi.get("label_fr", kid)
        if val is None:
            continue
        if any(m in kid_l for m in MARQUE_KEYS) and "marge" in kid_l:
            marques[label] = val

    if len(marques) < 3:
        return None

    sorted_m = sorted(marques.items(), key=lambda x: x[1], reverse=True)[:10]
    cats, vals = zip(*sorted_m)

    return {
        "type":         "bar",
        "title":        "Top marques — Marge HT hors ordonnances",
        "categories":   list(cats),
        "series": [{"name": "Marge HT", "values": list(vals), "color": VU_BLUE}],
        "value_format": "€",
    }


# ── PG_09_MERCH_EXPO — grouped bar exposition% vs marge% ─────────────────────

def _chart_merch_expo(kpi_dict: dict, **_) -> Optional[dict]:
    """Grouped bar : exposition linéaire% vs marge% par univers."""
    UNIVERS = ["senior", "sénior", "jambes", "nature", "libre", "hygiene",
               "hygiène", "bébé", "bebe", "beauté", "beaute", "veto", "véto"]

    expo_u, marge_u = {}, {}
    for kid, kpi in kpi_dict.items():
        if not isinstance(kpi, dict):
            continue
        kid_l = kid.lower()
        val   = kpi.get("valeur")
        if val is None:
            continue
        for u in UNIVERS:
            if u not in kid_l:
                continue
            canon = u.replace("é", "e").capitalize()
            if any(x in kid_l for x in ["expo", "lineaire", "linéaire", "rayon"]):
                expo_u[canon] = val
            elif "marge" in kid_l:
                marge_u[canon] = val

    all_u = sorted(set(list(expo_u) + list(marge_u)))
    if len(all_u) < 2:
        return None

    series = []
    if expo_u:
        series.append({"name": "Exposition linéaire (%)", "values": [expo_u.get(u, 0) for u in all_u], "color": VU_ORANGE})
    if marge_u:
        series.append({"name": "Marge (%)",               "values": [marge_u.get(u, 0) for u in all_u], "color": VU_NAVY})

    if len(series) < 1:
        return None

    return {
        "type":         "bar",
        "title":        "Exposition linéaire vs Marge — par univers (%)",
        "categories":   all_u,
        "series":       series,
        "value_format": "%",
    }


# ── PG_10_MERCH_STOCKS — horizontal bar rotation stocks ──────────────────────

def _chart_merch_stocks(kpi_dict: dict, **_) -> Optional[dict]:
    """Horizontal bar : jours de stock par univers vs benchmark."""
    STOCK_KW = ["rotation", "stock", "jours", "couverture"]
    stocks = {}
    for kid, kpi in kpi_dict.items():
        if not isinstance(kpi, dict):
            continue
        kid_l = kid.lower()
        val   = kpi.get("valeur")
        label = kpi.get("label_fr", kid)
        if val is None:
            continue
        if any(kw in kid_l for kw in STOCK_KW):
            stocks[label] = round(float(val), 0)

    if len(stocks) < 2:
        return None

    cats = sorted(stocks.keys(), key=lambda x: stocks[x], reverse=True)
    vals = [stocks[c] for c in cats]
    bm   = [BM["rotation_hors"]] * len(cats)

    return {
        "type":         "bar",
        "title":        "Rotation des stocks (jours) vs Benchmark optimal",
        "categories":   cats,
        "series": [
            {"name": "Votre officine (j)",           "values": vals, "color": VU_BLUE},
            {"name": f"Benchmark hors ordos ({int(BM['rotation_hors'])}j)", "values": bm, "color": VU_GREY},
        ],
        "value_format": "j",
    }


# ── Dispatch ─────────────────────────────────────────────────────────────────

CHART_BUILDERS = {
    "PG_00_CONTEXTE":      None,
    "PG_01_SECTION":       None,
    "PG_02_PROFIL_TYPE":   _chart_profil_type,
    "PG_03_PROFIL_PATIENTS": _chart_profil_patients,
    "PG_04_FINANCIERS":    _chart_financiers,
    "PG_05_COMMERCIAUX":   _chart_commerciaux,
    "PG_06_UNIVERS_CA":    _chart_univers_ca,
    "PG_07_UNIVERS_EVO":   _chart_univers_evo,
    "PG_08_TOP_MARQUES":   _chart_top_marques,
    "PG_09_MERCH_EXPO":    _chart_merch_expo,
    "PG_10_MERCH_STOCKS":  _chart_merch_stocks,
    "PG_11_MERCH_SIGNA":   None,
    "PG_12_SYNTHESE":      None,
}


def build_chart_for_slide(
    slide_id: str,
    kpi_dict: dict,
    image_results: list = None,
) -> Optional[dict]:
    """
    Retourne la spec graphique pour un slide, ou None si non applicable / données insuffisantes.
    """
    builder = CHART_BUILDERS.get(slide_id)
    if builder is None:
        return None
    try:
        return builder(kpi_dict, image_results=image_results)
    except Exception:
        return None
