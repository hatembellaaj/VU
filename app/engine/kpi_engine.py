"""
KPI Engine — Agence VU

Calcule les KPIs à partir des données Excel/images parsées.
Architecture en deux couches :
  1. RAW_RULES  : recherche directe d'une valeur dans une cellule Excel
  2. DERIVED_RULES : calcul à partir d'autres KPIs bruts (formule)

Règles de matching améliorées :
  - prefer_row: "first" | "last" | "max" — quelle ligne prendre dans une colonne
  - min_value / max_value : filtrage des valeurs hors plage raisonnable
  - AUCUN fallback heuristique "prend le max d'une feuille" → retourne None si pas trouvé
"""

from typing import Optional
import pandas as pd


# ── KPIs bruts — recherche directe dans les cellules ─────────────────────────

RAW_RULES: dict = {
    # ── Financiers ──────────────────────────────────────────────────────────
    "ca_total": {
        "label_fr": "Chiffre d'affaires total",
        "unite": "€",
        "seuil_bas": 800_000,
        "seuil_haut": 2_000_000,
        # Cherche d'abord dans des feuilles "synthèse" avec une ligne total
        "sheet_hints": ["Synthese", "Synthèse", "Annuel", "Total", "Bilan", "CA global", "CA total"],
        "header_hints": [
            "ca total", "ca ht total", "ca ht annuel", "chiffre d'affaires total",
            "total ca ht", "ca global", "ventes totales ht", "total ht",
            "ca annuel", "total ventes", "ca net",
        ],
        "prefer_row": "max",
        "min_value": 200_000,     # un CA total pharmacie est toujours > 200k€
        "preferred_sheets_only": False,
        # Stratégie alternative : sommer les colonnes mensuelles si feuille TVA
        "_aggregate_fallback": True,
    },
    "ca_total_from_tva": {
        "label_fr": "CA total (somme mensuelle TVA)",
        "unite": "€",
        "_raw_only": True,
        "seuil_bas": 0,
        "seuil_haut": 99_999_999,
        "sheet_hints": ["TVA", "Ventes par TVA", "Ventes mois", "Mensuel"],
        "header_hints": [
            "total", "ca total", "total ca", "ca ht", "total ht",
            "ca ttc", "total ttc", "montant total",
        ],
        "prefer_row": "sum",     # additionne toutes les lignes (12 mois → total annuel)
        "min_value": 1_000,      # ignore les valeurs négligeables
        "preferred_sheets_only": True,
    },
    "ca_tva_21": {
        "label_fr": "CA TVA 2,1% (ordonnances remboursables)",
        "unite": "€",
        "_raw_only": True,       # utilisé pour calculs dérivés, pas affiché seul
        "seuil_bas": 0,
        "seuil_haut": 99_999_999,
        "sheet_hints": ["TVA", "Ventes par TVA", "Remboursé", "Remboursable"],
        "header_hints": ["ca ht", "ca ht n", "montant ht", "ca"],
        # Filtre : uniquement la ligne où Taux de TVA = 2,1%
        "row_filter": {"col": "Taux de TVA", "val": "2,1"},
        "prefer_row": "first",
        "min_value": 1_000,
        "preferred_sheets_only": True,
    },
    "ca_hors_ordos": {
        "label_fr": "CA hors ordonnances (libre accès)",
        "unite": "€",
        "_raw_only": True,
        "seuil_bas": 0,
        "seuil_haut": 99_999_999,
        "sheet_hints": ["TVA", "Ventes", "Conseil", "Para", "CA"],
        "header_hints": [
            "ca hors ordo", "ca hors ordonnances", "ca conseil",
            "ca libre", "hors remboursé", "hors rembourse",
            "ventes libres", "tva 10", "tva 20", "taux 10", "taux 20",
            "non remboursable", "para",
        ],
        "prefer_row": "max",
        "min_value": 1_000,
    },
    "nb_transactions": {
        "label_fr": "Nombre de transactions total",
        "unite": "transactions",
        "_raw_only": True,
        "seuil_bas": 0,
        "seuil_haut": 99_999_999,
        "sheet_hints": ["Ventes", "Commercial", "Activité", "Transactions", "Actes", "TVA"],
        "header_hints": [
            "nb transactions", "nb actes", "nb tickets", "nb ventes",
            "nombre transactions", "nombre actes", "nombre ventes",
            "transactions totales", "actes totaux", "total transactions",
            "passages", "nb passages",
        ],
        # Sommer les lignes mensuelles
        "prefer_row": "sum",
        "min_value": 10,
        "preferred_sheets_only": True,  # ne pas chercher dans les transactions brutes
    },
    "nb_transactions_ordos": {
        "label_fr": "Nombre de transactions ordonnances",
        "unite": "transactions",
        "_raw_only": True,
        "seuil_bas": 0,
        "seuil_haut": 99_999_999,
        "sheet_hints": ["Ventes", "Commercial", "Ordonnances", "Actes"],
        "header_hints": [
            "nb actes ordo", "nb transactions ordo", "nb ordonnances",
            "actes ordonnances", "tickets ordo", "nb ventes ordo",
            "transactions ordo", "actes remboursables",
        ],
        "prefer_row": "max",
        "min_value": 10,
    },
    "evolution_ca_pct": {
        "label_fr": "Évolution CA (%)",
        "unite": "%",
        "seuil_bas": -2.0,
        "seuil_haut": 5.0,
        # Cherche exclusivement dans des feuilles d'évolution/synthèse
        # (jamais dans les transactions brutes "Toutes les ventes")
        "sheet_hints": ["Evolution", "Évolution", "Synthese", "Synthèse", "Bilan", "Comparatif"],
        "header_hints": [
            "evolution ca", "évolution ca", "variation ca", "croissance ca",
            "var ca", "delta ca", "evol ca", "% évolution ca",
            "evol ca %", "ca evol", "% variation ca",
        ],
        "prefer_row": "last",
        "min_value": -100,
        "max_value": 200,
        "preferred_sheets_only": True,   # JAMAIS dans les transactions brutes
    },
    "evolution_marge_pct": {
        "label_fr": "Évolution marge brute (%)",
        "unite": "%",
        "seuil_bas": -2.0,
        "seuil_haut": 5.0,
        "sheet_hints": ["Marge", "Evolution", "Évolution", "Synthese", "Comparatif"],
        "header_hints": [
            "evolution marge", "évolution marge", "variation marge",
            "var marge", "delta marge", "evol marge", "% évolution marge",
        ],
        "prefer_row": "last",
        "min_value": -100,
        "max_value": 200,
        "preferred_sheets_only": True,
    },
    "marge_brute": {
        "label_fr": "Marge brute",
        "unite": "€",
        "seuil_bas": 200_000,
        "seuil_haut": 600_000,
        "sheet_hints": ["Marge", "Financier", "Résultats", "Synthese", "Synthèse", "Bilan"],
        "header_hints": [
            "marge brute", "marge brute ht", "marge ht totale",
            "mb total", "marge totale", "total marge", "marge globale",
            "marge ht", "mb ht",
        ],
        "prefer_row": "max",
        "min_value": 50_000,
        "preferred_sheets_only": True,
    },
    # ── ETP (nb équivalents temps plein) — utilisé pour calculer CA/ETP et marge/ETP ──
    "nb_etp": {
        "label_fr": "Nombre d'ETP",
        "unite": "ETP",
        "_raw_only": True,   # intermédiaire, utilisé par les formules dérivées
        "seuil_bas": 0,
        "seuil_haut": 99_999,
        "sheet_hints": ["ETP", "RH", "Ressources humaines", "Personnel", "Effectifs"],
        "header_hints": [
            "etp", "nb etp", "nombre etp", "effectif etp",
            "equivalent temps plein", "équivalent temps plein",
            "etp total", "total etp", "etp annuel",
        ],
        "prefer_row": "last",
        "min_value": 0.5,
        "max_value": 200,
    },
    # ── Paniers (lookup direct si présents comme colonne calculée) ───────────
    "panier_moyen_direct": {
        "label_fr": "Panier moyen (colonne directe)",
        "unite": "€",
        "_raw_only": True,
        "seuil_bas": 20.0,
        "seuil_haut": 80.0,
        "sheet_hints": ["Panier", "Ticket", "CA", "Commercial"],
        "header_hints": [
            "panier moyen", "ticket moyen", "panier total moyen",
            "panier moyen total", "ticket moyen total",
        ],
        "prefer_row": "last",
        "min_value": 5,
        "max_value": 500,
    },
    "panier_ordonnances_direct": {
        "label_fr": "Panier moyen ordonnances (colonne directe)",
        "unite": "€",
        "_raw_only": True,
        "seuil_bas": 30.0,
        "seuil_haut": 120.0,
        "sheet_hints": ["Panier", "Ticket", "Ordonnances", "Commercial"],
        "header_hints": [
            "panier ordonnances", "panier ordo", "ticket ordo",
            "panier rx", "panier moyen ordo", "ticket moyen ordo",
            "panier moyen ordonnances",
        ],
        "prefer_row": "last",
        "min_value": 5,
        "max_value": 500,
    },
    "panier_conseil_direct": {
        "label_fr": "Panier moyen conseil (colonne directe)",
        "unite": "€",
        "_raw_only": True,
        "seuil_bas": 8.0,
        "seuil_haut": 40.0,
        "sheet_hints": ["Panier", "Ticket", "Conseil", "Commercial"],
        "header_hints": [
            "panier conseil", "panier hors ordo", "ticket conseil",
            "panier otc", "panier hors ordonnances",
            "panier moyen conseil", "ticket moyen conseil",
        ],
        "prefer_row": "last",
        "min_value": 1,
        "max_value": 200,
    },
    # ── Autres ──────────────────────────────────────────────────────────────
    "nb_clients_actifs": {
        "label_fr": "Nombre de clients actifs",
        "unite": "clients",
        "seuil_bas": 2_000,
        "seuil_haut": 8_000,
        "sheet_hints": ["Clients", "Fidélisation", "Fidelisation", "Activité"],
        "header_hints": [
            "clients actifs", "nb clients actifs", "nombre clients actifs",
            "patients actifs", "nb patients actifs",
        ],
        "prefer_row": "max",
        "min_value": 100,
    },
    "taux_fidelisation": {
        "label_fr": "Taux de fidélisation (%)",
        "unite": "%",
        "seuil_bas": 40.0,
        "seuil_haut": 70.0,
        "sheet_hints": ["Fidélisation", "Fidelisation", "Clients"],
        "header_hints": [
            "fidélisation", "fidelisation", "taux fidel",
            "taux de fidelisation", "retention", "rétention",
        ],
        "prefer_row": "last",
        "min_value": 0,
        "max_value": 100,
    },
    "indice_saisonnalite": {
        "label_fr": "Indice de saisonnalité",
        "unite": "indice",
        "seuil_bas": 0.7,
        "seuil_haut": 1.3,
        "sheet_hints": ["Saisonnalité", "Saisonnalite", "Mois"],
        "header_hints": [
            "saisonnalité", "saisonnalite", "indice saison", "seasonal",
            "indice mensuel",
        ],
        "prefer_row": "last",
        "min_value": 0.1,
        "max_value": 5.0,
    },
}

# Alias : DEFAULT_RULES pointe sur RAW_RULES pour compatibilité
DEFAULT_RULES = RAW_RULES


# ── KPIs dérivés — calculés à partir d'autres KPIs ───────────────────────────

DERIVED_RULES: dict = {
    # part_ordonnances_pct peut venir d'une colonne directe OU être calculé
    "part_ordonnances_pct": {
        "label_fr": "Part des ordonnances (%)",
        "unite": "%",
        "seuil_bas": 50.0,
        "seuil_haut": 80.0,
        # Cherche d'abord une colonne directe (dans feuilles dédiées seulement)
        "direct_hints": {
            "sheet_hints": ["Répartition", "Ordonnances", "Synthese", "Synthèse", "Panier", "IMG::"],
            "header_hints": [
                "part ordonnances", "% ordonnances", "taux ordo",
                "part ordo", "% ordo", "part rx", "part remboursable",
                "% remboursable",
                "repartition ca ttc ordonnances", "repartition ordonnances ca",
                # NOTE: "% tva 2,1" retiré — correspond au taux TVA (2.1), pas à la part
            ],
            "prefer_row": "last",
            "min_value": 20,
            "max_value": 100,
            "preferred_sheets_only": False,
        },
        # Si colonne directe absente → calcul
        "formula": lambda kpis: (
            (kpis["ca_tva_21"]["valeur"] / kpis["ca_total"]["valeur"] * 100)
            if (kpis.get("ca_tva_21", {}).get("valeur") is not None
                and kpis.get("ca_total", {}).get("valeur") not in (None, 0))
            else None
        ),
        "formula_deps": ["ca_tva_21", "ca_total"],
        "formula_source": "ca_tva_21 / ca_total × 100",
    },
    "frequentation_j": {
        "label_fr": "Fréquentation journalière (clients/jour)",
        "unite": "clients/j",
        "seuil_bas": 100,
        "seuil_haut": 250,
        "direct_hints": {
            "sheet_hints": ["Fréquentation", "Frequentation", "Clients", "Activité", "IMG::"],
            "header_hints": [
                "frequentation/jour", "frequentation journaliere",
                "clients/jour", "clients par jour", "passages/jour",
                "fréquentation journalière", "nb visites/jour",
                "fréquentation j", "freq/j",
                "ventes par jour", "moyenne ventes jour",
                "en moyenne vous avez",
            ],
            "prefer_row": "last",
            "min_value": 10,
            "max_value": 2000,
        },
        "formula": lambda kpis: (
            round(kpis["nb_transactions"]["valeur"] / 300, 1)
            if kpis.get("nb_transactions", {}).get("valeur") is not None
            else None
        ),
        "formula_deps": ["nb_transactions"],
        "formula_source": "nb_transactions / 300 jours ouvrés",
    },
    "panier_moyen": {
        "label_fr": "Panier moyen",
        "unite": "€",
        "seuil_bas": 20.0,
        "seuil_haut": 45.0,
        "direct_hints": {
            "sheet_hints": ["Panier", "Ticket", "CA", "Commercial", "IMG::"],
            "header_hints": [
                "panier moyen", "ticket moyen", "panier moyen total",
                "global", "panier global",
            ],
            "prefer_row": "last",
            "min_value": 5,
            "max_value": 500,
        },
        "formula": lambda kpis: (
            round(kpis["ca_total"]["valeur"] / kpis["nb_transactions"]["valeur"], 2)
            if (kpis.get("ca_total", {}).get("valeur") is not None
                and kpis.get("nb_transactions", {}).get("valeur") not in (None, 0))
            else None
        ),
        "formula_deps": ["ca_total", "nb_transactions"],
        "formula_source": "ca_total / nb_transactions",
    },
    "panier_ordonnances": {
        "label_fr": "Panier moyen ordonnances",
        "unite": "€",
        "seuil_bas": 45.0,
        "seuil_haut": 75.0,
        "direct_hints": {
            "sheet_hints": ["Panier", "Ticket", "Ordonnances", "Commercial", "IMG::"],
            "header_hints": [
                "panier ordonnances", "panier ordo", "ticket ordo",
                "panier rx", "panier moyen ordo", "panier moyen ordonnances",
                "ordonnance",
            ],
            "prefer_row": "last",
            "min_value": 5,
            "max_value": 500,
        },
        "formula": lambda kpis: (
            round(kpis["ca_tva_21"]["valeur"] / kpis["nb_transactions_ordos"]["valeur"], 2)
            if (kpis.get("ca_tva_21", {}).get("valeur") is not None
                and kpis.get("nb_transactions_ordos", {}).get("valeur") not in (None, 0))
            else None
        ),
        "formula_deps": ["ca_tva_21", "nb_transactions_ordos"],
        "formula_source": "ca_tva_21 / nb_transactions_ordos",
    },
    # ── CA et Marge par ETP (ratio ca_total ou marge_brute / nb_etp) ─────────
    "ca_par_etp": {
        "label_fr": "CA par ETP",
        "unite": "€",
        "seuil_bas": 300_000,
        "seuil_haut": 420_000,
        "direct_hints": {
            "sheet_hints": ["ETP", "RH", "Financier"],
            "header_hints": ["ca/etp", "ca par etp", "ca etp", "ca / etp"],
            "prefer_row": "last",
            "min_value": 10_000,
            "max_value": 2_000_000,
        },
        "formula": lambda kpis: (
            round(kpis["ca_total"]["valeur"] / kpis["nb_etp"]["valeur"], 2)
            if (kpis.get("ca_total", {}).get("valeur") is not None
                and kpis.get("nb_etp", {}).get("valeur") not in (None, 0))
            else None
        ),
        "formula_deps": ["ca_total", "nb_etp"],
        "formula_source": "ca_total / nb_etp",
    },
    "marge_par_etp": {
        "label_fr": "Marge par ETP",
        "unite": "€",
        "seuil_bas": 80_000,
        "seuil_haut": 130_000,
        "direct_hints": {
            "sheet_hints": ["ETP", "RH", "Marge"],
            "header_hints": ["marge/etp", "marge par etp", "marge etp", "mb par etp", "mb / etp"],
            "prefer_row": "last",
            "min_value": 5_000,
            "max_value": 1_000_000,
        },
        "formula": lambda kpis: (
            round(kpis["marge_brute"]["valeur"] / kpis["nb_etp"]["valeur"], 2)
            if (kpis.get("marge_brute", {}).get("valeur") is not None
                and kpis.get("nb_etp", {}).get("valeur") not in (None, 0))
            else None
        ),
        "formula_deps": ["marge_brute", "nb_etp"],
        "formula_source": "marge_brute / nb_etp",
    },
    "panier_conseil": {
        "label_fr": "Panier moyen conseil (hors ordos)",
        "unite": "€",
        "seuil_bas": 10.0,
        "seuil_haut": 18.0,
        "direct_hints": {
            "sheet_hints": ["Panier", "Ticket", "Conseil", "Commercial"],
            "header_hints": [
                "panier conseil", "panier hors ordo", "ticket conseil",
                "panier otc", "panier hors ordonnances",
            ],
            "prefer_row": "last",
            "min_value": 1,
            "max_value": 200,
        },
        "formula": lambda kpis: (
            round(
                kpis["ca_hors_ordos"]["valeur"]
                / max(kpis["nb_transactions"]["valeur"] - (kpis.get("nb_transactions_ordos", {}).get("valeur") or 0), 1),
                2,
            )
            if (kpis.get("ca_hors_ordos", {}).get("valeur") is not None
                and kpis.get("nb_transactions", {}).get("valeur") not in (None, 0))
            else None
        ),
        "formula_deps": ["ca_hors_ordos", "nb_transactions"],
        "formula_source": "ca_hors_ordos / (nb_transactions - nb_transactions_ordos)",
    },
}


# ── Moteur KPI ────────────────────────────────────────────────────────────────

class KPIEngine:
    """
    Computes pharmacy KPIs from parsed raw data.

    Two-pass architecture:
      Pass 1 — RAW_RULES  : lookup KPIs directly in Excel cells
      Pass 2 — DERIVED_RULES: first try direct column lookup, then formula from raw KPIs
    """

    def __init__(self, raw_data: dict, rules: dict = None):
        self.raw_data = raw_data
        self.rules = rules if rules is not None else RAW_RULES
        self._kpis: dict = {}

    # ── String normalization ─────────────────────────────────────────────────

    def _normalize_str(self, s: str) -> str:
        return (
            str(s).lower().strip()
            .replace("é", "e").replace("è", "e").replace("ê", "e")
            .replace("à", "a").replace("â", "a").replace("î", "i")
            .replace("ô", "o").replace("û", "u").replace("ç", "c")
            .replace("\xa0", " ")
        )

    # ── Core lookup ─────────────────────────────────────────────────────────

    def _find_in_image_sheet(
        self,
        sheet_name: str,
        sheet_data: dict,
        header_hints: list,
        min_value: Optional[float] = None,
        max_value: Optional[float] = None,
    ) -> tuple:
        """
        Recherche spéciale pour les feuilles images (IMG::...).
        Structure : headers=["label","valeur","unite","confiance"], rows=[["Panier moyen", 43.02, ...]]
        Cherche dans la colonne "label" (valeurs texte) plutôt que dans les en-têtes.
        Returns (value, sheet_name, cell_ref, label_matched) or (None, None, None, None).
        """
        headers = sheet_data.get("headers", [])
        rows    = sheet_data.get("rows", [])

        # Trouve les indices des colonnes "label" et "valeur"
        norm_headers = [self._normalize_str(h) for h in headers]
        try:
            label_idx = norm_headers.index("label")
            valeur_idx = norm_headers.index("valeur")
        except ValueError:
            return None, None, None, None

        for row in rows:
            if len(row) <= max(label_idx, valeur_idx):
                continue
            row_label = self._normalize_str(str(row[label_idx]))
            for hint in header_hints:
                norm_hint = self._normalize_str(hint)
                if norm_hint in row_label or row_label in norm_hint:
                    val = row[valeur_idx]
                    if not isinstance(val, (int, float)) or isinstance(val, bool):
                        continue
                    fval = float(val)
                    if min_value is not None and fval < min_value:
                        continue
                    if max_value is not None and fval > max_value:
                        continue
                    return fval, sheet_name, f"label={row[label_idx]}", str(row[label_idx])

        return None, None, None, None

    def _find_value_in_sheets(
        self,
        header_hints: list,
        sheet_hints: list,
        prefer_row: str = "first",
        min_value: Optional[float] = None,
        max_value: Optional[float] = None,
        preferred_sheets_only: bool = False,
        aggregate: Optional[str] = None,
        row_filter: Optional[dict] = None,
    ) -> tuple:
        """
        Search for a KPI value in parsed sheets.

        prefer_row:
          "first" — première ligne numérique de la colonne
          "last"  — dernière ligne numérique (souvent le total)
          "max"   — valeur absolue maximum dans la colonne
          "sum"   — somme de toutes les valeurs de la colonne (équivalent aggregate="sum")

        preferred_sheets_only:
          True  — ne cherche QUE dans les feuilles dont le nom matche sheet_hints
          False — cherche d'abord dans les feuilles préférées, puis dans les autres (défaut)

        aggregate:
          "sum" — additionne toutes les valeurs de la colonne (pour CA mensuel → annuel)

        row_filter:
          {"col": "Taux de TVA", "val": "2,1%"} — ne garde que les lignes où col == val
          Utile pour isoler une ligne spécifique dans un tableau multi-lignes (ex: TVA 2,1%).

        Returns (value, sheet_name, cell_ref, matched_col) or (None, None, None, None).
        AUCUN fallback heuristique — retourne None si aucun header ne correspond.
        """
        sheets = self.raw_data.get("sheets", {})

        # Tri : feuilles préférées en premier
        preferred, others = [], []
        for sname in sheets:
            norm = self._normalize_str(sname)
            if any(self._normalize_str(h) in norm for h in sheet_hints):
                preferred.append(sname)
            else:
                others.append(sname)

        # Si preferred_sheets_only → on ignore les feuilles non préférées
        search_order = preferred if preferred_sheets_only else preferred + others

        do_sum = (aggregate == "sum" or prefer_row == "sum")

        for sheet_name in search_order:
            sheet_data = sheets[sheet_name]

            # ── Feuilles images : logique label/valeur ───────────────────────
            if sheet_name.startswith("IMG::"):
                val, sn, cr, col = self._find_in_image_sheet(
                    sheet_name, sheet_data, header_hints, min_value, max_value
                )
                if val is not None:
                    return val, sn, cr, col
                continue  # ne pas appliquer la logique Excel aux images

            headers = sheet_data.get("headers", [])
            rows    = sheet_data.get("rows", [])
            numeric_cells = sheet_data.get("numeric_cells", [])

            # ── Résolution de row_filter : trouve l'index de la colonne filtre ──
            filter_col_idx = None
            filter_val_norm = None
            if row_filter:
                filter_col_norm = self._normalize_str(row_filter.get("col", ""))
                filter_val_norm = self._normalize_str(str(row_filter.get("val", "")))
                for hi, hdr in enumerate(headers):
                    if self._normalize_str(hdr) == filter_col_norm:
                        filter_col_idx = hi
                        break

            for col_idx, header in enumerate(headers):
                norm_header = self._normalize_str(header)
                for hint in header_hints:
                    norm_hint = self._normalize_str(hint)
                    if norm_hint in norm_header or norm_header in norm_hint:
                        # Collecte toutes les valeurs numériques de cette colonne
                        candidates = []
                        for row_idx, row in enumerate(rows):
                            # Applique le filtre de ligne si défini
                            if filter_col_idx is not None and filter_val_norm is not None:
                                if filter_col_idx >= len(row):
                                    continue
                                cell_norm = self._normalize_str(str(row[filter_col_idx]))
                                if filter_val_norm not in cell_norm and cell_norm not in filter_val_norm:
                                    continue
                            if col_idx < len(row):
                                val = row[col_idx]
                                if isinstance(val, (int, float)) and not isinstance(val, bool):
                                    fval = float(val)
                                    if min_value is not None and fval < min_value:
                                        continue
                                    if max_value is not None and fval > max_value:
                                        continue
                                    candidates.append((fval, row_idx))

                        if not candidates:
                            continue

                        # Agrégation ou sélection selon prefer_row
                        if do_sum:
                            chosen_val = round(sum(v for v, _ in candidates), 2)
                            # cell_ref = plage approximative
                            cell_ref = f"SUM({len(candidates)} lignes)"
                        elif prefer_row == "last":
                            chosen_val, chosen_row = candidates[-1]
                            cell_ref = None
                            for nc in numeric_cells:
                                if abs(nc["valeur"] - chosen_val) < 0.001 and nc["col"] == col_idx + 1:
                                    cell_ref = nc["ref"]
                                    break
                        elif prefer_row == "max":
                            chosen_val, chosen_row = max(candidates, key=lambda x: abs(x[0]))
                            cell_ref = None
                            for nc in numeric_cells:
                                if abs(nc["valeur"] - chosen_val) < 0.001 and nc["col"] == col_idx + 1:
                                    cell_ref = nc["ref"]
                                    break
                        else:  # "first"
                            chosen_val, chosen_row = candidates[0]
                            cell_ref = None
                            for nc in numeric_cells:
                                if abs(nc["valeur"] - chosen_val) < 0.001 and nc["col"] == col_idx + 1:
                                    cell_ref = nc["ref"]
                                    break

                        # Retourne aussi le nom exact de la colonne matchée
                        return chosen_val, sheet_name, cell_ref, str(header)

        # ── PAS DE FALLBACK HEURISTIQUE ─────────────────────────────────────
        # Retourner None proprement plutôt que de deviner une valeur incorrecte
        return None, None, None, None

    def _find_exact(
        self,
        sheet_key: str,
        col_name: str,
        prefer_row: str = "sum",
        min_value: Optional[float] = None,
        max_value: Optional[float] = None,
    ) -> tuple:
        """
        Recherche exacte par nom d'onglet et nom de colonne (fournis par l'utilisateur).
        Comparaison insensible à la casse et aux accents.

        Returns (value, sheet_key, cell_ref, col_name) or (None, None, None, None).
        """
        sheets = self.raw_data.get("sheets", {})
        norm_sheet = self._normalize_str(sheet_key)
        norm_col   = self._normalize_str(col_name)

        for sname, sdata in sheets.items():
            if self._normalize_str(sname) != norm_sheet:
                continue
            headers = sdata.get("headers", [])
            rows    = sdata.get("rows", [])
            numeric_cells = sdata.get("numeric_cells", [])

            for col_idx, header in enumerate(headers):
                if self._normalize_str(header) != norm_col:
                    continue

                candidates = []
                for row_idx, row in enumerate(rows):
                    if col_idx < len(row):
                        val = row[col_idx]
                        if isinstance(val, (int, float)) and not isinstance(val, bool):
                            fval = float(val)
                            if min_value is not None and fval < min_value:
                                continue
                            if max_value is not None and fval > max_value:
                                continue
                            candidates.append((fval, row_idx))

                if not candidates:
                    return None, None, None, None

                do_sum = (prefer_row == "sum")
                if do_sum:
                    chosen_val = round(sum(v for v, _ in candidates), 2)
                    cell_ref = f"SUM({len(candidates)} lignes)"
                elif prefer_row == "last":
                    chosen_val, _ = candidates[-1]
                    cell_ref = None
                    for nc in numeric_cells:
                        if abs(nc["valeur"] - chosen_val) < 0.001 and nc["col"] == col_idx + 1:
                            cell_ref = nc["ref"]
                            break
                elif prefer_row == "max":
                    chosen_val, _ = max(candidates, key=lambda x: abs(x[0]))
                    cell_ref = None
                    for nc in numeric_cells:
                        if abs(nc["valeur"] - chosen_val) < 0.001 and nc["col"] == col_idx + 1:
                            cell_ref = nc["ref"]
                            break
                else:  # first
                    chosen_val, _ = candidates[0]
                    cell_ref = None
                    for nc in numeric_cells:
                        if abs(nc["valeur"] - chosen_val) < 0.001 and nc["col"] == col_idx + 1:
                            cell_ref = nc["ref"]
                            break

                return chosen_val, sname, cell_ref, col_name

        return None, None, None, None

    def list_all_headers(self) -> dict:
        """
        Retourne tous les en-têtes de colonnes trouvés dans chaque feuille.
        Returns: {sheet_name: [header, ...]}
        """
        result = {}
        for sname, sdata in self.raw_data.get("sheets", {}).items():
            headers = [
                str(h).strip() for h in sdata.get("headers", [])
                if h and str(h).strip() and str(h).strip().lower() != "nan"
            ]
            if headers:
                result[sname] = headers
        return result

    def detect_mapping(self) -> list:
        """
        Dry-run : détecte quelle colonne/onglet correspond à chaque KPI
        sans modifier l'état interne. Retourne une liste de dicts pour
        alimenter le tableau de mapping utilisateur.

        Returns list of {
            kpi_id, label_fr, unite,
            found,           # bool — détection auto réussie
            source_type,     # "lookup" | "derived" | "not_found"
            onglet_auto,     # onglet détecté (ou "")
            colonne_auto,    # colonne détectée (ou "")
            value_preview,   # valeur numérique ou None
            _raw_only,       # bool — KPI intermédiaire non affiché normalement
        }
        """
        rows = []

        # Pass 1 : raw rules
        for kpi_id, rule in self.rules.items():
            val, sheet, _, col = self._find_value_in_sheets(
                rule.get("header_hints", []),
                rule.get("sheet_hints", []),
                prefer_row=rule.get("prefer_row", "first"),
                min_value=rule.get("min_value"),
                max_value=rule.get("max_value"),
                preferred_sheets_only=rule.get("preferred_sheets_only", False),
                aggregate=rule.get("aggregate"),
                row_filter=rule.get("row_filter"),
            )
            rows.append({
                "kpi_id":       kpi_id,
                "label_fr":     rule["label_fr"],
                "unite":        rule["unite"],
                "found":        val is not None,
                "source_type":  "lookup" if val is not None else "not_found",
                "onglet_auto":  sheet or "",
                "colonne_auto": col or "",
                "value_preview": val,
                "_raw_only":    rule.get("_raw_only", False),
            })

        # Pass 2 : derived rules (direct hints only for detection)
        for kpi_id, rule in DERIVED_RULES.items():
            direct = rule.get("direct_hints", {})
            val, sheet, _, col = (None, None, None, None)
            if direct:
                val, sheet, _, col = self._find_value_in_sheets(
                    direct.get("header_hints", []),
                    direct.get("sheet_hints", []),
                    prefer_row=direct.get("prefer_row", "last"),
                    min_value=direct.get("min_value"),
                    max_value=direct.get("max_value"),
                    preferred_sheets_only=direct.get("preferred_sheets_only", False),
                )

            # Check formula deps availability
            is_computable = False
            if val is None and "formula_deps" in rule:
                # Check if all deps would be found (rough check)
                is_computable = True  # optimistic — will be confirmed at compute time

            rows.append({
                "kpi_id":       kpi_id,
                "label_fr":     rule["label_fr"],
                "unite":        rule["unite"],
                "found":        val is not None or is_computable,
                "source_type":  "lookup" if val is not None else ("derived" if is_computable else "not_found"),
                "onglet_auto":  sheet or "",
                "colonne_auto": col or "",
                "value_preview": val,
                "_raw_only":    False,
            })

        return rows

    # ── Status helper ────────────────────────────────────────────────────────

    def compute_statut(self, valeur: Optional[float], seuil_bas: float, seuil_haut: float) -> str:
        if valeur is None:
            return "inconnu"
        if valeur >= seuil_haut:
            return "bon"
        if valeur >= seuil_bas:
            return "moyen"
        return "faible"

    # ── Main computation ─────────────────────────────────────────────────────

    def compute_all(self, overrides: Optional[dict] = None) -> dict:
        """
        Compute all KPIs (raw + derived).

        Args:
            overrides: dict {kpi_id: {"onglet": str, "colonne": str, "prefer_row": str}}
                       Spécifications exactes fournies par l'utilisateur. Prennent la
                       priorité sur la détection automatique par hints.

        Returns dict of kpi_id -> KPI entry.
        """
        self._kpis = {}
        overrides = overrides or {}
        source_fichier = self.raw_data.get("source", "inconnu")

        def _lookup_raw(kpi_id: str, rule: dict) -> tuple:
            """Retourne (valeur, onglet, cellule, matched_col) pour un KPI brut."""
            ov = overrides.get(kpi_id, {})
            if ov.get("onglet") and ov.get("colonne"):
                # Override utilisateur : recherche exacte
                v, s, c, col = self._find_exact(
                    ov["onglet"], ov["colonne"],
                    prefer_row=ov.get("prefer_row", rule.get("prefer_row", "sum")),
                    min_value=rule.get("min_value"),
                    max_value=rule.get("max_value"),
                )
                return v, s, c, col, "override"
            # Détection automatique
            v, s, c, col = self._find_value_in_sheets(
                rule.get("header_hints", []),
                rule.get("sheet_hints", []),
                prefer_row=rule.get("prefer_row", "first"),
                min_value=rule.get("min_value"),
                max_value=rule.get("max_value"),
                preferred_sheets_only=rule.get("preferred_sheets_only", False),
                aggregate=rule.get("aggregate"),
                row_filter=rule.get("row_filter"),
            )
            return v, s, c, col, "lookup"

        # ── PASS 1 : KPIs bruts ──────────────────────────────────────────────
        for kpi_id, rule in self.rules.items():
            valeur, onglet, cellule, matched_col, src_type = _lookup_raw(kpi_id, rule)

            statut = self.compute_statut(
                valeur,
                rule.get("seuil_bas", 0),
                rule.get("seuil_haut", 999_999_999),
            )

            self._kpis[kpi_id] = {
                "kpi_id":         kpi_id,
                "label_fr":       rule["label_fr"],
                "valeur":         valeur,
                "unite":          rule["unite"],
                "statut":         statut,
                "seuil_bas":      rule.get("seuil_bas", 0),
                "seuil_haut":     rule.get("seuil_haut", 999_999_999),
                "source_fichier": source_fichier,
                "onglet":         onglet,
                "cellule":        cellule,
                "matched_col":    matched_col,
                "_raw_only":      rule.get("_raw_only", False),
                "source_type":    src_type,
            }

        # ── PASS 1b : fallback ca_total depuis somme TVA mensuelles ──────────
        if self._kpis.get("ca_total", {}).get("valeur") is None:
            tva_entry = self._kpis.get("ca_total_from_tva", {})
            if tva_entry.get("valeur") is not None:
                self._kpis["ca_total"]["valeur"]      = tva_entry["valeur"]
                self._kpis["ca_total"]["onglet"]      = tva_entry.get("onglet")
                self._kpis["ca_total"]["cellule"]     = tva_entry.get("cellule")
                self._kpis["ca_total"]["matched_col"] = tva_entry.get("matched_col")
                self._kpis["ca_total"]["source_type"] = "sum_mensuel"
                self._kpis["ca_total"]["statut"] = self.compute_statut(
                    tva_entry["valeur"], 800_000, 2_000_000,
                )

        # ── PASS 2 : KPIs dérivés (lookup direct OU formule) ─────────────────
        for kpi_id, rule in DERIVED_RULES.items():
            ov = overrides.get(kpi_id, {})
            valeur, onglet, cellule, matched_col, source_type = None, None, None, None, None

            # 2a : override utilisateur → recherche exacte
            if ov.get("onglet") and ov.get("colonne"):
                all_rules_merged = {**self.rules, **DERIVED_RULES}
                ref_rule = all_rules_merged.get(kpi_id, rule)
                valeur, onglet, cellule, matched_col = self._find_exact(
                    ov["onglet"], ov["colonne"],
                    prefer_row=ov.get("prefer_row", "last"),
                    min_value=rule.get("seuil_bas"),
                    max_value=rule.get("seuil_haut"),
                )
                if valeur is not None:
                    source_type = "override"

            # 2b : lookup par hints directs
            if valeur is None:
                direct = rule.get("direct_hints", {})
                if direct:
                    valeur, onglet, cellule, matched_col = self._find_value_in_sheets(
                        direct.get("header_hints", []),
                        direct.get("sheet_hints", []),
                        prefer_row=direct.get("prefer_row", "last"),
                        min_value=direct.get("min_value"),
                        max_value=direct.get("max_value"),
                        preferred_sheets_only=direct.get("preferred_sheets_only", False),
                        aggregate=direct.get("aggregate"),
                    )
                    if valeur is not None:
                        source_type = "lookup"

            # 2c : formule dérivée
            if valeur is None and "formula" in rule:
                try:
                    computed = rule["formula"](self._kpis)
                    if computed is not None:
                        valeur = computed
                        source_type = "computed"
                        deps = rule.get("formula_deps", [])
                        onglets_src = [
                            self._kpis[d]["onglet"]
                            for d in deps
                            if d in self._kpis and self._kpis[d].get("onglet")
                        ]
                        onglet   = " + ".join(set(filter(None, onglets_src))) or "calculé"
                        cellule  = rule.get("formula_source", "formule")
                        matched_col = rule.get("formula_source", "")
                except Exception:
                    valeur = None

            statut = self.compute_statut(
                valeur,
                rule.get("seuil_bas", 0),
                rule.get("seuil_haut", 999_999_999),
            )

            self._kpis[kpi_id] = {
                "kpi_id":         kpi_id,
                "label_fr":       rule["label_fr"],
                "valeur":         valeur,
                "unite":          rule["unite"],
                "statut":         statut,
                "seuil_bas":      rule.get("seuil_bas", 0),
                "seuil_haut":     rule.get("seuil_haut", 999_999_999),
                "source_fichier": source_fichier,
                "onglet":         onglet,
                "cellule":        cellule,
                "matched_col":    matched_col,
                "_raw_only":      False,
                "source_type":    source_type or "inconnu",
            }

        return self._kpis

    # ── DataFrame export ─────────────────────────────────────────────────────

    def get_as_dataframe(self) -> pd.DataFrame:
        """Return computed KPIs as a pandas DataFrame (hides _raw_only entries)."""
        if not self._kpis:
            self.compute_all()

        records = [
            v for v in self._kpis.values()
            if not v.get("_raw_only", False)
        ]
        if not records:
            return pd.DataFrame(columns=[
                "kpi_id", "label_fr", "valeur", "unite", "statut",
                "seuil_bas", "seuil_haut", "source_fichier", "onglet", "cellule",
                "matched_col", "source_type",
            ])

        df = pd.DataFrame(records)
        cols = ["kpi_id", "label_fr", "valeur", "unite", "statut",
                "seuil_bas", "seuil_haut", "source_fichier", "onglet", "cellule",
                "matched_col", "source_type"]
        return df[[c for c in cols if c in df.columns]]
