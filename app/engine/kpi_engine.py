from typing import Optional
import pandas as pd


# Default industry benchmark thresholds for French pharmacies
DEFAULT_RULES = {
    "ca_total": {
        "label_fr": "Chiffre d'affaires total",
        "unite": "€",
        "seuil_bas": 800_000,
        "seuil_haut": 2_000_000,
        "sheet_hints": ["CA", "Chiffre", "Ventes", "Total"],
        "header_hints": ["ca total", "chiffre d'affaires", "chiffre daffaires", "ventes totales"],
    },
    "evolution_ca_pct": {
        "label_fr": "Évolution CA (%)",
        "unite": "%",
        "seuil_bas": -2.0,
        "seuil_haut": 5.0,
        "sheet_hints": ["Evolution", "Évolution", "CA"],
        "header_hints": ["evolution ca", "évolution ca", "variation ca", "croissance"],
    },
    "panier_moyen": {
        "label_fr": "Panier moyen",
        "unite": "€",
        "seuil_bas": 20.0,
        "seuil_haut": 45.0,
        "sheet_hints": ["Panier", "Ticket", "CA"],
        "header_hints": ["panier moyen", "ticket moyen", "panier"],
    },
    "nb_clients_actifs": {
        "label_fr": "Nombre de clients actifs",
        "unite": "clients",
        "seuil_bas": 2_000,
        "seuil_haut": 8_000,
        "sheet_hints": ["Clients", "Fidélisation", "Fidelisation"],
        "header_hints": ["clients actifs", "nb clients", "nombre clients"],
    },
    "frequentation_mensuelle": {
        "label_fr": "Fréquentation mensuelle moyenne",
        "unite": "visites/mois",
        "seuil_bas": 500,
        "seuil_haut": 3_000,
        "sheet_hints": ["Fréquentation", "Frequentation", "Clients", "Visites"],
        "header_hints": ["frequentation", "fréquentation", "visites", "passages"],
    },
    "part_ordonnances_pct": {
        "label_fr": "Part des ordonnances (%)",
        "unite": "%",
        "seuil_bas": 50.0,
        "seuil_haut": 80.0,
        "sheet_hints": ["Ordonnances", "Prescriptions", "CA"],
        "header_hints": ["ordonnances", "prescriptions", "part ordo", "rx"],
    },
    "part_parapharmacie_pct": {
        "label_fr": "Part de la parapharmacie (%)",
        "unite": "%",
        "seuil_bas": 5.0,
        "seuil_haut": 25.0,
        "sheet_hints": ["Parapharmacie", "Para", "CA"],
        "header_hints": ["parapharmacie", "para", "otc", "conseil"],
    },
    "taux_fidelisation": {
        "label_fr": "Taux de fidélisation (%)",
        "unite": "%",
        "seuil_bas": 40.0,
        "seuil_haut": 70.0,
        "sheet_hints": ["Fidélisation", "Fidelisation", "Clients"],
        "header_hints": ["fidélisation", "fidelisation", "taux fidel", "retention"],
    },
    "evolution_panier_pct": {
        "label_fr": "Évolution panier moyen (%)",
        "unite": "%",
        "seuil_bas": -1.0,
        "seuil_haut": 4.0,
        "sheet_hints": ["Panier", "Evolution", "CA"],
        "header_hints": ["evolution panier", "évolution panier", "variation panier"],
    },
    "indice_saisonnalite": {
        "label_fr": "Indice de saisonnalité",
        "unite": "indice",
        "seuil_bas": 0.7,
        "seuil_haut": 1.3,
        "sheet_hints": ["Saisonnalité", "Saisonnalite", "Mois"],
        "header_hints": ["saisonnalité", "saisonnalite", "indice saison", "seasonal"],
    },
}


class KPIEngine:
    """Computes pharmacy KPIs from parsed raw data."""

    def __init__(self, raw_data: dict, rules: dict = None):
        """
        Initialize KPIEngine.

        Args:
            raw_data: Output from ExcelParser or combined parser results.
                      Expected structure: {"sheets": {sheet_name: {"headers": [], "rows": [[]], "numeric_cells": []}}}
            rules: Optional custom KPI rules (overrides defaults).
        """
        self.raw_data = raw_data
        self.rules = rules if rules is not None else DEFAULT_RULES
        self._kpis: dict = {}

    def _normalize_str(self, s: str) -> str:
        """Lowercase and strip for fuzzy matching."""
        return (
            str(s)
            .lower()
            .strip()
            .replace("é", "e")
            .replace("è", "e")
            .replace("ê", "e")
            .replace("à", "a")
            .replace("â", "a")
            .replace("î", "i")
            .replace("ô", "o")
            .replace("û", "u")
            .replace("ç", "c")
        )

    def _find_value_in_sheets(
        self, header_hints: list, sheet_hints: list
    ) -> tuple:
        """
        Search for a KPI value in parsed sheets using header and sheet hints.

        Returns:
            (value, sheet_name, cell_ref) or (None, None, None)
        """
        sheets = self.raw_data.get("sheets", {})

        # First pass: preferred sheets
        preferred_sheets = []
        all_sheets = []
        for sname in sheets:
            norm_sname = self._normalize_str(sname)
            is_preferred = any(
                self._normalize_str(hint) in norm_sname for hint in sheet_hints
            )
            if is_preferred:
                preferred_sheets.append(sname)
            else:
                all_sheets.append(sname)

        search_order = preferred_sheets + all_sheets

        for sheet_name in search_order:
            sheet_data = sheets[sheet_name]
            headers = sheet_data.get("headers", [])
            rows = sheet_data.get("rows", [])
            numeric_cells = sheet_data.get("numeric_cells", [])

            # Look for matching header
            for col_idx, header in enumerate(headers):
                norm_header = self._normalize_str(header)
                for hint in header_hints:
                    norm_hint = self._normalize_str(hint)
                    if norm_hint in norm_header or norm_header in norm_hint:
                        # Found matching column — look for first numeric value
                        for row in rows:
                            if col_idx < len(row):
                                val = row[col_idx]
                                if isinstance(val, (int, float)) and not isinstance(
                                    val, bool
                                ):
                                    # Find cell ref from numeric_cells
                                    cell_ref = None
                                    for nc in numeric_cells:
                                        if (
                                            abs(nc["valeur"] - float(val)) < 0.001
                                            and nc["col"] == col_idx + 1
                                        ):
                                            cell_ref = nc["ref"]
                                            break
                                    return float(val), sheet_name, cell_ref

        # Second pass: scan all numeric cells for any that might match
        for sheet_name, sheet_data in sheets.items():
            numeric_cells = sheet_data.get("numeric_cells", [])
            if numeric_cells:
                # Return largest numeric value in preferred sheets as fallback
                norm_sname = self._normalize_str(sheet_name)
                is_preferred = any(
                    self._normalize_str(hint) in norm_sname for hint in sheet_hints
                )
                if is_preferred and numeric_cells:
                    # Take the maximum value from preferred sheet as heuristic
                    best = max(numeric_cells, key=lambda x: abs(x["valeur"]))
                    return best["valeur"], sheet_name, best["ref"]

        return None, None, None

    def compute_statut(
        self,
        valeur: Optional[float],
        seuil_bas: float,
        seuil_haut: float,
    ) -> str:
        """Determine performance status based on thresholds."""
        if valeur is None:
            return "inconnu"
        if valeur >= seuil_haut:
            return "bon"
        if valeur >= seuil_bas:
            return "moyen"
        return "faible"

    def compute_all(self) -> dict:
        """
        Compute all KPIs from raw data.

        Returns:
            Dict of kpi_id -> KPI entry:
            {
                "kpi_id": str,
                "label_fr": str,
                "valeur": float | None,
                "unite": str,
                "statut": "bon" | "moyen" | "faible" | "inconnu",
                "seuil_bas": float,
                "seuil_haut": float,
                "source_fichier": str,
                "onglet": str | None,
                "cellule": str | None
            }
        """
        self._kpis = {}
        source_fichier = self.raw_data.get("source", "inconnu")

        for kpi_id, rule in self.rules.items():
            header_hints = rule.get("header_hints", [])
            sheet_hints = rule.get("sheet_hints", [])

            valeur, onglet, cellule = self._find_value_in_sheets(
                header_hints, sheet_hints
            )

            statut = self.compute_statut(
                valeur, rule["seuil_bas"], rule["seuil_haut"]
            )

            self._kpis[kpi_id] = {
                "kpi_id": kpi_id,
                "label_fr": rule["label_fr"],
                "valeur": valeur,
                "unite": rule["unite"],
                "statut": statut,
                "seuil_bas": rule["seuil_bas"],
                "seuil_haut": rule["seuil_haut"],
                "source_fichier": source_fichier,
                "onglet": onglet,
                "cellule": cellule,
            }

        return self._kpis

    def get_as_dataframe(self) -> pd.DataFrame:
        """Return computed KPIs as a pandas DataFrame."""
        if not self._kpis:
            self.compute_all()

        records = list(self._kpis.values())
        if not records:
            return pd.DataFrame(
                columns=[
                    "kpi_id",
                    "label_fr",
                    "valeur",
                    "unite",
                    "statut",
                    "seuil_bas",
                    "seuil_haut",
                    "source_fichier",
                    "onglet",
                    "cellule",
                ]
            )

        df = pd.DataFrame(records)
        df = df[
            [
                "kpi_id",
                "label_fr",
                "valeur",
                "unite",
                "statut",
                "seuil_bas",
                "seuil_haut",
                "source_fichier",
                "onglet",
                "cellule",
            ]
        ]
        return df
