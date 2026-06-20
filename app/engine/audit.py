import re
from typing import Optional


class AuditEngine:
    """
    Audits generated narrative content against computed KPI values.
    Ensures no hallucinated numbers appear in LLM output.
    """

    def _extract_numbers_with_context(self, text: str) -> list:
        """
        Extract all numbers from text with surrounding context.

        Returns list of (number_float, context_str) tuples.
        """
        # Pattern: optional sign, digits with optional thousands separator, optional decimal
        pattern = r"[-+]?\d[\d\s ]*(?:[.,]\d+)?"
        results = []
        for match in re.finditer(pattern, text):
            raw = match.group().strip()
            # Normalize: remove thousands separators (spaces, non-breaking spaces)
            normalized = (
                raw.replace(" ", "")
                .replace("\xa0", "")
                .replace(" ", "")
                .replace(",", ".")
            )
            try:
                value = float(normalized)
                # Ignore trivially small integers used as ordinals/years/counts (1, 2, 3 etc.)
                # We keep them because even small numbers must be validated
                start = max(0, match.start() - 50)
                end = min(len(text), match.end() + 50)
                context = text[start:end].replace("\n", " ")
                results.append((value, context))
            except ValueError:
                continue
        return results

    def _is_number_valid(self, number: float, kpi_dict: dict, tolerance: float = 0.01) -> bool:
        """
        Check if a number exists in any KPI value within tolerance.

        Checks:
        - Direct KPI values
        - Seuil_bas and seuil_haut thresholds
        - Rounded versions (0 decimal, 1 decimal, 2 decimals)
        """
        for kpi in kpi_dict.values():
            kpi_value = kpi.get("valeur")
            if kpi_value is None:
                continue

            # Direct comparison with tolerance
            if abs(number - kpi_value) <= tolerance:
                return True

            # Check rounded variants
            for decimals in (0, 1, 2):
                rounded = round(kpi_value, decimals)
                if abs(number - rounded) <= tolerance:
                    return True

            # Check percentage equivalents (e.g., 0.45 vs 45%)
            if abs(number - kpi_value * 100) <= tolerance:
                return True
            if abs(number * 100 - kpi_value) <= tolerance:
                return True

            # Check seuils
            for seuil_key in ("seuil_bas", "seuil_haut"):
                seuil = kpi.get(seuil_key)
                if seuil is not None and abs(number - seuil) <= tolerance:
                    return True

        return False

    def audit(self, generated_content: str, kpi_dict: dict) -> dict:
        """
        Audit generated narrative against computed KPI values.

        Args:
            generated_content: LLM-generated text to audit.
            kpi_dict: Dict of kpi_id -> KPI entry (from KPIEngine.compute_all()).

        Returns:
            {
                "passed": bool,
                "score_pct": float,
                "total_numbers_found": int,
                "validated": int,
                "rejected": [
                    {
                        "number": float,
                        "context": str
                    }
                ],
                "message": str
            }
        """
        if not generated_content or not generated_content.strip():
            return {
                "passed": True,
                "score_pct": 100.0,
                "total_numbers_found": 0,
                "validated": 0,
                "rejected": [],
                "message": "Aucun contenu à auditer.",
            }

        extracted = self._extract_numbers_with_context(generated_content)

        if not extracted:
            return {
                "passed": True,
                "score_pct": 100.0,
                "total_numbers_found": 0,
                "validated": 0,
                "rejected": [],
                "message": "Aucun nombre trouvé dans le contenu généré.",
            }

        # Filter out years (1900-2100) and small ordinals (1-31) since they
        # are legitimate in narrative text without being KPI values
        def is_trivial(n: float) -> bool:
            """Years and small integers used as ordinals are not KPI numbers."""
            if n == int(n):
                ni = int(n)
                if 1900 <= ni <= 2100:  # years
                    return True
                if 1 <= ni <= 12:  # months
                    return True
            return False

        non_trivial = [(n, ctx) for n, ctx in extracted if not is_trivial(n)]

        if not non_trivial:
            return {
                "passed": True,
                "score_pct": 100.0,
                "total_numbers_found": len(extracted),
                "validated": len(extracted),
                "rejected": [],
                "message": "Tous les nombres trouvés sont des valeurs triviales (années, ordinals).",
            }

        validated_count = 0
        rejected_list = []

        for number, context in non_trivial:
            if self._is_number_valid(number, kpi_dict):
                validated_count += 1
            else:
                rejected_list.append({"number": number, "context": context})

        total = len(non_trivial)
        score_pct = (validated_count / total * 100) if total > 0 else 100.0
        passed = len(rejected_list) == 0

        if passed:
            message = (
                f"Audit réussi: {validated_count}/{total} nombres validés (100%). "
                "Aucun nombre halluciné détecté."
            )
        else:
            message = (
                f"Audit échoué: {validated_count}/{total} nombres validés "
                f"({score_pct:.1f}%). "
                f"{len(rejected_list)} nombre(s) non trouvé(s) dans les KPIs calculés."
            )

        return {
            "passed": passed,
            "score_pct": round(score_pct, 2),
            "total_numbers_found": len(extracted),
            "validated": validated_count,
            "rejected": rejected_list,
            "message": message,
        }
