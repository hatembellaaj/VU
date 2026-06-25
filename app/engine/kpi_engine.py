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
        "header_hints": [
            "tva 2,1", "tva 2.1", "2,1%", "2.1%",
            "remboursable", "remboursé", "rembourse",
            "ca ordo", "ca ordonnances", "ca remboursable",
            "ventes remboursables",
        ],
        # Sommer les 12 mois pour obtenir le total annuel
        "prefer_row": "sum",
        "min_value": 1_000,
        "preferred_sheets_only": True,   # ne PAS chercher dans les transactions brutes
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
        "sheet_hints": ["Evolution", "Évolution", "Synthese", "Synthèse", "Bilan"],
        "header_hints": [
            "evolution ca", "évolution ca", "variation ca", "croissance ca",
            "var ca", "delta ca", "evol ca", "% évolution ca",
            "evol ca %", "ca evol",
        ],
        "prefer_row": "last",
        "min_value": -100,
        "max_value": 200,
        # Ne cherche PAS dans les fichiers de transactions brutes
        "preferred_sheets_only": True,
    },
    "evolution_marge_pct": {
        "label_fr": "Évolution marge brute (%)",
        "unite": "%",
        "seuil_bas": -2.0,
        "seuil_haut": 5.0,
        "sheet_hints": ["Marge", "Evolution", "Évolution", "Synthese"],
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
        "sheet_hints": ["Marge", "CA", "Financier", "Résultats", "Synthese"],
        "header_hints": [
            "marge brute", "marge brute ht", "marge ht totale",
            "mb total", "marge totale", "total marge", "marge globale",
        ],
        "prefer_row": "max",
        "min_value": 50_000,
        "preferred_sheets_only": True,
    },
    "ca_par_etp": {
        "label_fr": "CA par ETP",
        "unite": "€",
        "seuil_bas": 300_000,
        "seuil_haut": 420_000,
        "sheet_hints": ["ETP", "RH", "Financier", "CA"],
        "header_hints": [
            "ca/etp", "ca par etp", "ca etp", "chiffre par etp",
            "ca par equivalent", "ca / etp",
        ],
        "prefer_row": "last",
        "min_value": 10_000,
    },
    "marge_par_etp": {
        "label_fr": "Marge par ETP",
        "unite": "€",
        "seuil_bas": 80_000,
        "seuil_haut": 130_000,
        "sheet_hints": ["ETP", "RH", "Marge"],
        "header_hints": [
            "marge/etp", "marge par etp", "marge etp",
            "mb par etp", "mb / etp",
        ],
        "prefer_row": "last",
        "min_value": 5_000,
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
            "sheet_hints": ["Répartition", "Ordonnances", "Synthese", "Synthèse"],
            "header_hints": [
                "part ordonnances", "% ordonnances", "taux ordo",
                "part ordo", "% ordo", "part rx", "part remboursable",
                "% remboursable",
                # NOTE: "% tva 2,1" retiré — correspond au taux TVA (2.1), pas à la part
            ],
            "prefer_row": "last",
            "min_value": 20,     # part ordo < 20% serait aberrant pour une pharmacie française
            "max_value": 100,
            "preferred_sheets_only": True,
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
            "sheet_hints": ["Fréquentation", "Frequentation", "Clients", "Activité"],
            "header_hints": [
                "frequentation/jour", "frequentation journaliere",
                "clients/jour", "clients par jour", "passages/jour",
                "fréquentation journalière", "nb visites/jour",
                "fréquentation j", "freq/j",
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
            "sheet_hints": ["Panier", "Ticket", "CA", "Commercial"],
            "header_hints": [
                "panier moyen", "ticket moyen", "panier moyen total",
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
            "sheet_hints": ["Panier", "Ticket", "Ordonnances", "Commercial"],
            "header_hints": [
                "panier ordonnances", "panier ordo", "ticket ordo",
                "panier rx", "panier moyen ordo", "panier moyen ordonnances",
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

    def _find_value_in_sheets(
        self,
        header_hints: list,
        sheet_hints: list,
        prefer_row: str = "first",
        min_value: Optional[float] = None,
        max_value: Optional[float] = None,
        preferred_sheets_only: bool = False,
        aggregate: Optional[str] = None,
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

        Returns (value, sheet_name, cell_ref) or (None, None, None).
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
            headers = sheet_data.get("headers", [])
            rows    = sheet_data.get("rows", [])
            numeric_cells = sheet_data.get("numeric_cells", [])

            for col_idx, header in enumerate(headers):
                norm_header = self._normalize_str(header)
                for hint in header_hints:
                    norm_hint = self._normalize_str(hint)
                    if norm_hint in norm_header or norm_header in norm_hint:
                        # Collecte toutes les valeurs numériques de cette colonne
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

                        return chosen_val, sheet_name, cell_ref

        # ── PAS DE FALLBACK HEURISTIQUE ─────────────────────────────────────
        # Retourner None proprement plutôt que de deviner une valeur incorrecte
        return None, None, None

    def list_all_headers(self) -> dict:
        """
        Retourne tous les en-têtes de colonnes trouvés dans chaque feuille.
        Utile pour le diagnostic : permet d'identifier les vrais noms de colonnes Excel.

        Returns: {sheet_name: [header, ...]}
        """
        result = {}
        for sname, sdata in self.raw_data.get("sheets", {}).items():
            headers = [h for h in sdata.get("headers", []) if h and str(h).strip()]
            if headers:
                result[sname] = headers
        return result

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

    def compute_all(self) -> dict:
        """
        Compute all KPIs (raw + derived).

        Returns dict of kpi_id -> KPI entry.
        """
        self._kpis = {}
        source_fichier = self.raw_data.get("source", "inconnu")

        # ── PASS 1 : KPIs bruts (lookup direct) ─────────────────────────────
        for kpi_id, rule in self.rules.items():
            header_hints           = rule.get("header_hints", [])
            sheet_hints            = rule.get("sheet_hints", [])
            prefer_row             = rule.get("prefer_row", "first")
            min_value              = rule.get("min_value")
            max_value              = rule.get("max_value")
            preferred_sheets_only  = rule.get("preferred_sheets_only", False)
            aggregate              = rule.get("aggregate")

            valeur, onglet, cellule = self._find_value_in_sheets(
                header_hints, sheet_hints,
                prefer_row=prefer_row,
                min_value=min_value,
                max_value=max_value,
                preferred_sheets_only=preferred_sheets_only,
                aggregate=aggregate,
            )

            statut = self.compute_statut(
                valeur,
                rule.get("seuil_bas", 0),
                rule.get("seuil_haut", 999_999_999),
            )

            self._kpis[kpi_id] = {
                "kpi_id":        kpi_id,
                "label_fr":      rule["label_fr"],
                "valeur":        valeur,
                "unite":         rule["unite"],
                "statut":        statut,
                "seuil_bas":     rule.get("seuil_bas", 0),
                "seuil_haut":    rule.get("seuil_haut", 999_999_999),
                "source_fichier": source_fichier,
                "onglet":        onglet,
                "cellule":       cellule,
                "_raw_only":     rule.get("_raw_only", False),
                "source_type":   "lookup",
            }

        # ── PASS 1b : fallback ca_total depuis somme TVA ────────────────────────
        # Si ca_total n'a pas été trouvé comme ligne de synthèse, utilise la somme
        # des valeurs mensuelles TVA (ca_total_from_tva)
        if self._kpis.get("ca_total", {}).get("valeur") is None:
            tva_entry = self._kpis.get("ca_total_from_tva", {})
            if tva_entry.get("valeur") is not None:
                self._kpis["ca_total"]["valeur"]   = tva_entry["valeur"]
                self._kpis["ca_total"]["onglet"]   = tva_entry.get("onglet")
                self._kpis["ca_total"]["cellule"]  = tva_entry.get("cellule")
                self._kpis["ca_total"]["source_type"] = "sum_mensuel"
                self._kpis["ca_total"]["statut"] = self.compute_statut(
                    tva_entry["valeur"], 800_000, 2_000_000,
                )

        # Même logique pour ca_tva_21 si non trouvé : essai avec hints plus larges
        # (le module intermédiaire ca_tva_21 doit être trouvé pour les dérivés)

        # ── PASS 2 : KPIs dérivés (lookup direct OU formule) ─────────────────
        for kpi_id, rule in DERIVED_RULES.items():
            # 2a : essaie d'abord le lookup direct
            direct = rule.get("direct_hints", {})
            valeur, onglet, cellule, source_type = None, None, None, None

            if direct:
                valeur, onglet, cellule = self._find_value_in_sheets(
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

            # 2b : si lookup raté → formule
            if valeur is None and "formula" in rule:
                try:
                    computed = rule["formula"](self._kpis)
                    if computed is not None:
                        valeur = computed
                        source_type = "computed"
                        # Récupère les onglets sources
                        deps = rule.get("formula_deps", [])
                        onglets_src = [
                            self._kpis[d]["onglet"]
                            for d in deps
                            if d in self._kpis and self._kpis[d].get("onglet")
                        ]
                        onglet  = " + ".join(set(onglets_src)) if onglets_src else "calculé"
                        cellule = rule.get("formula_source", "formule")
                except Exception:
                    valeur = None

            statut = self.compute_statut(
                valeur,
                rule.get("seuil_bas", 0),
                rule.get("seuil_haut", 999_999_999),
            )

            self._kpis[kpi_id] = {
                "kpi_id":        kpi_id,
                "label_fr":      rule["label_fr"],
                "valeur":        valeur,
                "unite":         rule["unite"],
                "statut":        statut,
                "seuil_bas":     rule.get("seuil_bas", 0),
                "seuil_haut":    rule.get("seuil_haut", 999_999_999),
                "source_fichier": source_fichier,
                "onglet":        onglet,
                "cellule":       cellule,
                "_raw_only":     False,
                "source_type":   source_type or "inconnu",
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
                "seuil_bas", "seuil_haut", "source_fichier", "onglet", "cellule", "source_type",
            ])

        df = pd.DataFrame(records)
        cols = ["kpi_id", "label_fr", "valeur", "unite", "statut",
                "seuil_bas", "seuil_haut", "source_fichier", "onglet", "cellule", "source_type"]
        return df[[c for c in cols if c in df.columns]]
