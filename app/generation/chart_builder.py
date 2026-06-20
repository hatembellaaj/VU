"""
Générateur de spécifications de graphiques pour Agence VU.

Chaque slide peut recevoir un graphique construit DÉTERMINISTIQUEMENT
depuis les KPIs calculés et les données images extraites.
Les graphiques sont définis dans la méthodologie :

  Slide Profil patients   → bar chart pyramide des âges officine vs référence
  Slide Financiers        → column chart CA/marge N-1 vs N + position Marge/ETP
  Slide Commerciaux       → grouped bar fréquentation + paniers vs benchmarks
  Slide Univers CA/marge  → horizontal bar répartition par univers
  Slide Univers évolution → column chart évolution % par univers
  Slide Merchandising expo→ grouped bar exposition% vs marge% par univers
  Slide Merchandising stock→ horizontal bar jours de stock par univers

Spec retournée :
{
  "type":       "bar" | "column" | "line" | "pie",
  "title":      str,
  "categories": [str, ...],
  "series": [
    {"name": str, "values": [float, ...], "color": "#RRGGBB" (optionnel)}
  ],
  "value_format": "%" | "€" | "j" | None,
  "show_benchmark_line": float | None   # valeur de la ligne de référence marché
}
"""

from typing import Optional

# ── Couleurs Agence VU ────────────────────────────────────────────────────────
VU_BLUE   = "#008BD2"
VU_NAVY   = "#1A2E4A"
VU_GREEN  = "#27AE60"
VU_ORANGE = "#F39C12"
VU_RED    = "#E74C3C"
VU_GREY   = "#BDC3C7"
VU_LIGHT  = "#D6EAF8"

# Benchmarks marché (source: méthodologie Agence VU 2025)
BENCHMARK = {
    "frequentation_j":    180.0,
    "panier_total":        40.8,
    "panier_ordos":        58.3,
    "panier_conseil":      13.89,
    "marge_etp_faible":    90_000,
    "marge_etp_correct":  105_000,
    "marge_etp_performant":120_000,
    "rotation_ordos_min":  30,
    "rotation_ordos_max":  40,
    "rotation_hors_min":   80,
    "rotation_hors_max":  100,
}


def _kv(kpi_dict: dict, key: str, default=None):
    """Récupère la valeur d'un KPI."""
    entry = kpi_dict.get(key, {})
    if isinstance(entry, dict):
        return entry.get("valeur", default)
    return default


# ── Builders par slide ────────────────────────────────────────────────────────

def _chart_financiers(kpi_dict: dict) -> Optional[dict]:
    """
    Column chart : CA et Marge Brute — Évolution N-1 → N.
    Nécessite : ca_total + evolution_ca_pct OU ca_n1 + marge_brute + marge_n1.
    """
    ca_n    = _kv(kpi_dict, "ca_total")
    marge_n = _kv(kpi_dict, "marge_brute")
    evo_ca  = _kv(kpi_dict, "evolution_ca_pct")
    evo_mg  = _kv(kpi_dict, "evolution_marge_pct")

    if ca_n is None or marge_n is None:
        return None

    # Reconstitue N-1 si évolution disponible
    if evo_ca is not None:
        ca_n1 = round(ca_n / (1 + evo_ca / 100), 0)
    else:
        ca_n1 = None

    if evo_mg is not None:
        marge_n1 = round(marge_n / (1 + evo_mg / 100), 0)
    else:
        marge_n1 = None

    if ca_n1 is not None and marge_n1 is not None:
        categories = ["N-1", "N"]
        ca_vals    = [ca_n1, ca_n]
        marge_vals = [marge_n1, marge_n]
    else:
        categories = ["N"]
        ca_vals    = [ca_n]
        marge_vals = [marge_n]

    return {
        "type":         "column",
        "title":        "CA et Marge Brute HT",
        "categories":   categories,
        "series": [
            {"name": "Chiffre d'Affaires HT", "values": ca_vals,    "color": VU_NAVY},
            {"name": "Marge Brute HT",        "values": marge_vals, "color": VU_BLUE},
        ],
        "value_format": "€",
        "show_benchmark_line": None,
    }


def _chart_commerciaux(kpi_dict: dict) -> Optional[dict]:
    """
    Grouped bar : officine vs benchmark — Fréquentation et paniers moyens.
    """
    freq = _kv(kpi_dict, "frequentation_j") or _kv(kpi_dict, "nb_clients_actifs")
    panier = _kv(kpi_dict, "panier_moyen")

    if freq is None and panier is None:
        return None

    categories = []
    officine_vals = []
    benchmark_vals = []

    if freq is not None:
        categories.append("Fréquentation\n(clients/jour)")
        officine_vals.append(round(freq, 1))
        benchmark_vals.append(BENCHMARK["frequentation_j"])

    if panier is not None:
        categories.append("Panier moyen\ntotal (€)")
        officine_vals.append(round(panier, 2))
        benchmark_vals.append(BENCHMARK["panier_total"])

    panier_conseil = _kv(kpi_dict, "panier_conseil")
    if panier_conseil is not None:
        categories.append("Panier moyen\nconseil (€)")
        officine_vals.append(round(panier_conseil, 2))
        benchmark_vals.append(BENCHMARK["panier_conseil"])

    return {
        "type":       "bar",
        "title":      "Indicateurs commerciaux vs marché",
        "categories": categories,
        "series": [
            {"name": "Votre officine", "values": officine_vals,  "color": VU_BLUE},
            {"name": "Benchmark marché 2025", "values": benchmark_vals, "color": VU_GREY},
        ],
        "value_format": None,
        "show_benchmark_line": None,
    }


def _chart_clientele(kpi_dict: dict, image_results: list = None) -> Optional[dict]:
    """
    Bar chart horizontal : Profil patients par tranche d'âge.
    Si les données image contiennent une pyramide des âges, on l'utilise.
    Sinon on affiche fréquentation + taux de fidélisation vs benchmarks.
    """
    # Tentative de reconstruction depuis les données images extraites
    if image_results:
        age_data = _extract_age_pyramid(image_results)
        if age_data:
            return age_data

    # Fallback : fréquentation + paniers si disponibles
    return _chart_commerciaux(kpi_dict)


def _extract_age_pyramid(image_results: list) -> Optional[dict]:
    """
    Tente d'extraire une pyramide des âges depuis les valeurs images.
    Cherche des labels contenant des tranches d'âge (ex: "0-4", "5-14", "65+").
    """
    age_keywords = ["ans", "age", "âge", "tranche", "-", "+"]
    officine_vals = {}
    ref_vals = {}

    for img_result in (image_results or []):
        for entry in img_result.get("valeurs_extraites", []):
            label = str(entry.get("label", "")).lower()
            valeur = entry.get("valeur")
            if valeur is None:
                continue
            # Détecte les labels de tranche d'âge
            is_age = any(kw in label for kw in age_keywords) or (
                any(c.isdigit() for c in label) and (
                    "+" in label or "-" in label or "ans" in label
                )
            )
            if not is_age:
                continue
            # Distingue officine vs référence dans le label
            if any(ref in label for ref in ["ref", "ville", "national", "france", "moyen"]):
                ref_vals[label] = float(valeur)
            else:
                officine_vals[label] = float(valeur)

    if len(officine_vals) < 3:
        return None

    categories = sorted(officine_vals.keys())
    ofi_series  = [officine_vals.get(c, 0) for c in categories]
    ref_series  = [ref_vals.get(c, 0) for c in categories] if ref_vals else None

    series = [{"name": "Votre officine", "values": ofi_series, "color": VU_BLUE}]
    if ref_series and any(v > 0 for v in ref_series):
        series.append({"name": "Référence locale", "values": ref_series, "color": VU_GREY})

    return {
        "type":         "bar",
        "title":        "Profil patients — Répartition par tranche d'âge (%)",
        "categories":   categories,
        "series":       series,
        "value_format": "%",
        "show_benchmark_line": None,
    }


def _chart_univers_repartition(kpi_dict: dict) -> Optional[dict]:
    """
    Horizontal bar : CA% et Marge% par univers hors ordonnances.
    Cherche des KPIs dont la clé contient un nom d'univers connu.
    """
    univers_keywords = [
        "senior", "sénior", "jambes", "nature", "bébé", "bebe",
        "beauté", "beaute", "hygiene", "hygiène", "libre_acces",
        "libre accès", "veto", "véto",
    ]

    ca_by_univers    = {}
    marge_by_univers = {}

    for kpi_id, kpi in kpi_dict.items():
        kpi_id_low = kpi_id.lower()
        val = kpi.get("valeur")
        if val is None:
            continue
        for uname in univers_keywords:
            if uname in kpi_id_low:
                canon = uname.capitalize()
                if "ca" in kpi_id_low or "chiffre" in kpi_id_low:
                    ca_by_univers[canon] = val
                elif "marge" in kpi_id_low:
                    marge_by_univers[canon] = val

    if not ca_by_univers and not marge_by_univers:
        return None

    all_univers = sorted(set(list(ca_by_univers.keys()) + list(marge_by_univers.keys())))
    series = []
    if ca_by_univers:
        series.append({
            "name": "CA Hors ordos (%)",
            "values": [ca_by_univers.get(u, 0) for u in all_univers],
            "color": VU_BLUE,
        })
    if marge_by_univers:
        series.append({
            "name": "Marge Hors ordos (%)",
            "values": [marge_by_univers.get(u, 0) for u in all_univers],
            "color": VU_NAVY,
        })

    return {
        "type":         "bar",
        "title":        "Répartition CA et Marge — Univers hors ordonnances",
        "categories":   all_univers,
        "series":       series,
        "value_format": "%",
        "show_benchmark_line": None,
    }


def _chart_stocks(kpi_dict: dict) -> Optional[dict]:
    """
    Horizontal bar : Rotation stocks (jours) par univers vs benchmark.
    """
    rotation_keys = {k: v for k, v in kpi_dict.items() if "rotation" in k.lower() or "stock" in k.lower()}
    if not rotation_keys:
        return None

    categories, vals = [], []
    for kpi_id, kpi in rotation_keys.items():
        val = kpi.get("valeur")
        if val is not None:
            label = kpi.get("label_fr", kpi_id)
            categories.append(label)
            vals.append(round(val, 0))

    if not categories:
        return None

    return {
        "type":       "bar",
        "title":      "Rotation des stocks (jours)",
        "categories": categories,
        "series": [
            {"name": "Votre officine",              "values": vals,
             "color": VU_BLUE},
            {"name": "Benchmark optimal (hors ordos)",
             "values": [BENCHMARK["rotation_hors_max"]] * len(categories),
             "color": VU_GREY},
        ],
        "value_format": "j",
        "show_benchmark_line": BENCHMARK["rotation_hors_max"],
    }


# ── Dispatch principal ────────────────────────────────────────────────────────

# Mapping slide_id → builder
CHART_BUILDERS = {
    "PG_01_INTRO":        _chart_financiers,
    "PG_02_CA":           _chart_financiers,
    "PG_03_CLIENTELE":    _chart_clientele,
    "PG_04_PANIER":       _chart_commerciaux,
    "PG_05_SAISONNALITE": _chart_commerciaux,
    "PG_06_SYNTHESE":     None,   # SWOT — pas de graphique
}


def build_chart_for_slide(
    slide_id: str,
    kpi_dict: dict,
    image_results: list = None,
) -> Optional[dict]:
    """
    Retourne la spec de graphique pour un slide donné, ou None si non applicable.

    Args:
        slide_id:      Identifiant du slide (ex: "PG_03_CLIENTELE")
        kpi_dict:      Dict des KPIs calculés
        image_results: Liste des résultats image (lot1 ou lot2) pour les données
                       pyramide des âges

    Returns:
        Dict spec graphique ou None
    """
    builder = CHART_BUILDERS.get(slide_id)
    if builder is None:
        return None

    try:
        if slide_id == "PG_03_CLIENTELE":
            return builder(kpi_dict, image_results)
        return builder(kpi_dict)
    except Exception:
        return None
