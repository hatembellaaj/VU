"""
data_checker.py — Agence VU

Vérifie la disponibilité des données requises par chaque slide
avant et après la génération LLM.

Le module lit optionnellement le mapping depuis la méthodologie :
  ## Mapping des données
  <!-- VU_MAPPING -->
  ```json
  {
    "ca_total": {"colonnes": ["CA Total", "CA HT"], "onglets": ["CA", "Financier"]},
    ...
  }
  ```
  <!-- /VU_MAPPING -->

Si ce bloc est absent, les hints par défaut du KPIEngine sont utilisés.
"""

from __future__ import annotations

import json
import re
from typing import Optional

# ── Import paresseux pour éviter les imports circulaires ─────────────────────
def _get_default_rules():
    from engine.kpi_engine import DEFAULT_RULES
    return DEFAULT_RULES


def _get_derived_rules():
    from engine.kpi_engine import DERIVED_RULES
    return DERIVED_RULES


def _get_slides_def():
    from generation.llm_generator import PERFORMANCE_GLOBALE_SLIDES
    return PERFORMANCE_GLOBALE_SLIDES


# ── Constantes ────────────────────────────────────────────────────────────────

STATUS_OK        = "ok"           # Toutes les données requises trouvées
STATUS_PARTIAL   = "partial"      # Certaines données manquantes
STATUS_MISSING   = "missing"      # Aucune donnée trouvée (mais requises)
STATUS_NO_DATA   = "no_data"      # Slide ne nécessite pas de données
STATUS_SKIPPED   = "skipped"      # Slide conditionnel non généré (ex: contexte sans PDF)

STATUS_LABELS = {
    STATUS_OK:      "✅ Données complètes",
    STATUS_PARTIAL: "⚠️ Données partielles",
    STATUS_MISSING: "❌ Données manquantes",
    STATUS_NO_DATA: "ℹ️ Aucune donnée requise",
    STATUS_SKIPPED: "⏭️ Non généré",
}

STATUS_COLORS = {
    STATUS_OK:      "#EAFAF1",
    STATUS_PARTIAL: "#FEF9E7",
    STATUS_MISSING: "#FDEDEC",
    STATUS_NO_DATA: "#EBF5FB",
    STATUS_SKIPPED: "#F8F9FA",
}


# ── Parsing du mapping depuis la méthodologie ─────────────────────────────────

def parse_mapping_from_methodology(methodology_text: str) -> dict:
    """
    Extrait le bloc JSON de mapping données depuis la méthodologie.

    Format attendu dans le Markdown :
      <!-- VU_MAPPING -->
      ```json
      { "ca_total": {"colonnes": [...], "onglets": [...]}, ... }
      ```
      <!-- /VU_MAPPING -->

    Retourne un dict {kpi_id: {"colonnes": [...], "onglets": [...]}}
    ou un dict vide si le bloc n'est pas trouvé ou invalide.
    """
    if not methodology_text:
        return {}

    # Cherche le bloc entre les balises VU_MAPPING
    pattern = r'<!--\s*VU_MAPPING\s*-->(.*?)<!--\s*/VU_MAPPING\s*-->'
    match = re.search(pattern, methodology_text, re.DOTALL | re.IGNORECASE)
    if not match:
        return {}

    block = match.group(1).strip()

    # Retire les balises ```json ... ```
    json_match = re.search(r'```(?:json)?\s*(\{.*?\})\s*```', block, re.DOTALL)
    if json_match:
        block = json_match.group(1)

    try:
        data = json.loads(block)
        # Valider la structure minimale
        if isinstance(data, dict):
            return {
                kpi_id: {
                    "colonnes": v.get("colonnes", []) if isinstance(v, dict) else [],
                    "onglets":  v.get("onglets", [])  if isinstance(v, dict) else [],
                }
                for kpi_id, v in data.items()
            }
    except (json.JSONDecodeError, AttributeError):
        pass

    return {}


def build_mapping_markdown(mapping: dict) -> str:
    """
    Sérialise un mapping dict en bloc Markdown insérable dans la méthodologie.
    """
    json_str = json.dumps(mapping, ensure_ascii=False, indent=2)
    return (
        "\n\n## Mapping des données\n\n"
        "Ce bloc est utilisé par le pipeline pour localiser les données dans les fichiers Excel.\n"
        "Modifiez les listes `colonnes` et `onglets` selon vos fichiers.\n\n"
        "<!-- VU_MAPPING -->\n"
        f"```json\n{json_str}\n```\n"
        "<!-- /VU_MAPPING -->\n"
    )


# ── Construction des hints enrichis pour le KPIEngine ────────────────────────

def enrich_kpi_rules(base_rules: dict, methodology_mapping: dict) -> dict:
    """
    Fusionne les hints du mapping méthodologie dans les règles KPI de base.
    Les noms de colonnes/onglets issus de la méthodologie sont ajoutés EN TÊTE
    des listes (priorité maximale).

    Retourne une copie enrichie des règles — n'altère pas l'original.
    """
    import copy
    enriched = copy.deepcopy(base_rules)

    for kpi_id, mapping_entry in methodology_mapping.items():
        if kpi_id not in enriched:
            # KPI inconnu du moteur — on l'ignore
            continue

        extra_cols   = mapping_entry.get("colonnes", [])
        extra_sheets = mapping_entry.get("onglets", [])

        if extra_cols:
            # Ajouter en tête sans doublon
            existing = [h.lower() for h in enriched[kpi_id].get("header_hints", [])]
            new_hints = [c for c in extra_cols if c.lower() not in existing]
            enriched[kpi_id]["header_hints"] = new_hints + enriched[kpi_id].get("header_hints", [])

        if extra_sheets:
            existing = [s.lower() for s in enriched[kpi_id].get("sheet_hints", [])]
            new_hints = [s for s in extra_sheets if s.lower() not in existing]
            enriched[kpi_id]["sheet_hints"] = new_hints + enriched[kpi_id].get("sheet_hints", [])

    return enriched


# ── Vérification de disponibilité par slide ───────────────────────────────────

def check_slides_data(
    kpi_dict: dict,
    methodology_mapping: Optional[dict] = None,
    context_text: str = "",
    slides_def: Optional[list] = None,
) -> dict:
    """
    Vérifie la disponibilité des données requises pour chaque slide.

    Args:
        kpi_dict:           KPIs calculés par le KPIEngine (kpi_id → entry).
        methodology_mapping:Mapping issu de parse_mapping_from_methodology().
        context_text:       Texte du PDF de contexte (pour PG_00_CONTEXTE).
        slides_def:         Définitions des slides (defaut : PERFORMANCE_GLOBALE_SLIDES).

    Returns:
        {
          "slides": [
            {
              "slide_id": str,
              "titre":    str,
              "status":   "ok" | "partial" | "missing" | "no_data" | "skipped",
              "status_label": str,
              "kpis_found":   [kpi_id, ...],
              "kpis_missing": [kpi_id, ...],
              "missing_details": [
                {
                  "kpi_id":       str,
                  "label_fr":     str,
                  "colonnes_attendues": [...],
                  "onglets_attendus":   [...],
                }
              ],
              "found_details": [
                {
                  "kpi_id": str,
                  "label_fr": str,
                  "valeur": float,
                  "unite": str,
                  "onglet": str,
                  "cellule": str,
                }
              ],
            }
          ],
          "summary": {"ok": int, "partial": int, "missing": int, "no_data": int, "skipped": int},
          "global_status": "ok" | "partial" | "missing",
        }
    """
    if slides_def is None:
        slides_def = _get_slides_def()

    mapping = methodology_mapping or {}
    default_rules = _get_default_rules()
    derived_rules = _get_derived_rules()

    result_slides = []
    summary = {STATUS_OK: 0, STATUS_PARTIAL: 0, STATUS_MISSING: 0,
               STATUS_NO_DATA: 0, STATUS_SKIPPED: 0}

    for slide in slides_def:
        slide_id       = slide["slide_id"]
        titre          = slide["titre_defaut"]
        kpis_requis    = slide.get("kpis_requis", [])
        requires_ctx   = slide.get("requires_context", False)

        # ── Slide contexte : dépend du PDF ────────────────────────────────
        if requires_ctx:
            if context_text and context_text.strip():
                status = STATUS_OK
                entry = {
                    "slide_id": slide_id,
                    "titre": titre,
                    "status": status,
                    "status_label": STATUS_LABELS[status],
                    "kpis_found": [],
                    "kpis_missing": [],
                    "missing_details": [],
                    "found_details": [{"kpi_id": "_context", "label_fr": "PDF contexte fourni",
                                       "valeur": len(context_text), "unite": "chars",
                                       "onglet": "—", "cellule": "—"}],
                    "note": f"PDF de contexte — {len(context_text):,} caractères extraits",
                }
            else:
                status = STATUS_SKIPPED
                entry = {
                    "slide_id": slide_id,
                    "titre": titre,
                    "status": status,
                    "status_label": STATUS_LABELS[status],
                    "kpis_found": [],
                    "kpis_missing": [],
                    "missing_details": [],
                    "found_details": [],
                    "note": "Aucun PDF de contexte fourni — slide non généré",
                }
            summary[status] += 1
            result_slides.append(entry)
            continue

        # ── Slides sans données requises ─────────────────────────────────
        if not kpis_requis:
            status = STATUS_NO_DATA
            result_slides.append({
                "slide_id": slide_id,
                "titre": titre,
                "status": status,
                "status_label": STATUS_LABELS[status],
                "kpis_found": [],
                "kpis_missing": [],
                "missing_details": [],
                "found_details": [],
                "note": "Ce slide ne nécessite pas de données calculées",
            })
            summary[status] += 1
            continue

        # ── Slides avec KPIs requis ───────────────────────────────────────
        kpis_found   = []
        kpis_missing = []
        missing_details = []
        found_details   = []

        for kpi_id in kpis_requis:
            kpi_entry = kpi_dict.get(kpi_id)

            if kpi_entry and kpi_entry.get("valeur") is not None:
                kpis_found.append(kpi_id)
                found_details.append({
                    "kpi_id":   kpi_id,
                    "label_fr": kpi_entry.get("label_fr", kpi_id),
                    "valeur":   kpi_entry.get("valeur"),
                    "unite":    kpi_entry.get("unite", ""),
                    "onglet":   kpi_entry.get("onglet") or "—",
                    "cellule":  kpi_entry.get("cellule") or "—",
                })
            else:
                kpis_missing.append(kpi_id)

                # Récupérer les noms de colonnes/onglets attendus
                colonnes_attendues = []
                onglets_attendus   = []

                # Priorité : mapping méthodologie
                if kpi_id in mapping:
                    colonnes_attendues = mapping[kpi_id].get("colonnes", [])
                    onglets_attendus   = mapping[kpi_id].get("onglets", [])
                # Fallback : règles par défaut du KPIEngine (raw puis dérivé)
                if not colonnes_attendues:
                    if kpi_id in default_rules:
                        colonnes_attendues = default_rules[kpi_id].get("header_hints", [])
                    elif kpi_id in derived_rules:
                        dh = derived_rules[kpi_id].get("direct_hints", {})
                        colonnes_attendues = dh.get("header_hints", [])
                        formula_src = derived_rules[kpi_id].get("formula_source", "")
                        if formula_src:
                            colonnes_attendues = [f"Calculé: {formula_src}"] + colonnes_attendues
                if not onglets_attendus:
                    if kpi_id in default_rules:
                        onglets_attendus = default_rules[kpi_id].get("sheet_hints", [])
                    elif kpi_id in derived_rules:
                        dh = derived_rules[kpi_id].get("direct_hints", {})
                        onglets_attendus = dh.get("sheet_hints", [])

                label_fr = (
                    (kpi_entry or {}).get("label_fr")
                    or default_rules.get(kpi_id, {}).get("label_fr")
                    or derived_rules.get(kpi_id, {}).get("label_fr", kpi_id)
                )

                missing_details.append({
                    "kpi_id":            kpi_id,
                    "label_fr":          label_fr,
                    "colonnes_attendues": colonnes_attendues,
                    "onglets_attendus":   onglets_attendus,
                })

        # Déterminer le statut
        if not kpis_missing:
            status = STATUS_OK
        elif not kpis_found:
            status = STATUS_MISSING
        else:
            status = STATUS_PARTIAL

        result_slides.append({
            "slide_id":       slide_id,
            "titre":          titre,
            "status":         status,
            "status_label":   STATUS_LABELS[status],
            "kpis_found":     kpis_found,
            "kpis_missing":   kpis_missing,
            "missing_details": missing_details,
            "found_details":   found_details,
            "note": (
                f"{len(kpis_found)}/{len(kpis_requis)} KPI(s) disponible(s)"
                if kpis_requis else ""
            ),
        })
        summary[status] += 1

    # Statut global
    if summary[STATUS_MISSING] > 0 or summary[STATUS_PARTIAL] > 0:
        global_status = STATUS_PARTIAL if summary[STATUS_OK] > 0 else STATUS_MISSING
    else:
        global_status = STATUS_OK

    return {
        "slides":        result_slides,
        "summary":       summary,
        "global_status": global_status,
    }


# ── Mapping par défaut (utile pour initialiser une méthodologie) ──────────────

DEFAULT_METHODOLOGY_MAPPING = {
    # ── KPIs bruts directs ────────────────────────────────────────────────────
    "ca_total": {
        "colonnes": ["CA Total", "CA HT Total", "Chiffre d'affaires HT", "CA HT", "Total CA"],
        "onglets":  ["CA", "Financier", "Synthèse", "Résultats", "Ventes par TVA"]
    },
    # KPIs intermédiaires nécessaires aux calculs dérivés
    "ca_tva_21": {
        "colonnes": ["TVA 2,1%", "2,1%", "CA TVA 2,1", "Remboursable", "CA ordonnances"],
        "onglets":  ["TVA", "Ventes par TVA", "CA", "Remboursé"]
    },
    "ca_hors_ordos": {
        "colonnes": ["CA hors ordo", "TVA 10%", "TVA 20%", "CA conseil", "Libre accès"],
        "onglets":  ["TVA", "Ventes par TVA", "CA"]
    },
    "nb_transactions": {
        "colonnes": ["Nb transactions", "Nb actes", "Nb tickets", "Nb ventes", "Passages"],
        "onglets":  ["Ventes", "Commercial", "Activité", "Transactions"]
    },
    "nb_transactions_ordos": {
        "colonnes": ["Nb actes ordo", "Nb ordonnances", "Actes ordonnances", "Tickets ordo"],
        "onglets":  ["Ventes", "Commercial", "Ordonnances"]
    },
    "marge_brute": {
        "colonnes": ["Marge brute", "Marge HT", "MB", "Marge totale"],
        "onglets":  ["Marge", "CA", "Financier"]
    },
    "evolution_ca_pct": {
        "colonnes": ["Évolution CA", "Evolution CA (%)", "Var. CA", "Croissance CA"],
        "onglets":  ["CA", "Evolution", "Financier"]
    },
    "evolution_marge_pct": {
        "colonnes": ["Évolution marge", "Evolution marge (%)", "Var. marge"],
        "onglets":  ["Marge", "CA", "Evolution"]
    },
    "nb_etp": {
        "colonnes": ["ETP", "Nb ETP", "Nombre ETP", "Équivalent temps plein", "Total ETP"],
        "onglets":  ["ETP", "RH", "Ressources humaines", "Personnel", "Effectifs"]
    },
    "ca_par_etp": {
        "colonnes": ["CA/ETP", "CA par ETP", "Calculé: ca_total / nb_etp"],
        "onglets":  ["ETP", "RH", "Financier", "CA"]
    },
    "marge_par_etp": {
        "colonnes": ["Marge/ETP", "Marge par ETP", "Calculé: marge_brute / nb_etp"],
        "onglets":  ["ETP", "RH", "Marge"]
    },
    # ── KPIs dérivés (calculés depuis les intermédiaires) ─────────────────────
    "frequentation_j": {
        "colonnes": ["Fréquentation/jour", "Clients/jour", "Passages/jour",
                     "Calculé: nb_transactions / 300 jours ouvrés"],
        "onglets":  ["Fréquentation", "Clients", "Activité"]
    },
    "panier_moyen": {
        "colonnes": ["Panier moyen", "Ticket moyen", "Calculé: ca_total / nb_transactions"],
        "onglets":  ["Panier", "CA", "Commercial"]
    },
    "panier_ordonnances": {
        "colonnes": ["Panier ordonnances", "Panier ordo", "Ticket ordo", "Panier Rx",
                     "Calculé: ca_tva_21 / nb_transactions_ordos"],
        "onglets":  ["Panier", "Ordonnances", "Commercial"]
    },
    "panier_conseil": {
        "colonnes": ["Panier conseil", "Panier hors ordo", "Ticket conseil", "Panier OTC",
                     "Calculé: ca_hors_ordos / nb_transactions_conseil"],
        "onglets":  ["Panier", "Conseil", "Commercial"]
    },
    "part_ordonnances_pct": {
        "colonnes": ["Part ordonnances (%)", "% ordonnances", "Taux ordo", "Part Rx",
                     "Calculé: ca_tva_21 / ca_total × 100"],
        "onglets":  ["Ordonnances", "CA", "Répartition", "TVA"]
    },
}
