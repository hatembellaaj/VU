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

        Exclusions :
        - "N-1", "N-2" (comptabilité française) : précédé d'une lettre → ignoré
        - Intervalles du type "19–21" : seul le premier nombre est gardé (pas le second)
        """
        # (?<![A-Za-z/]) — ne pas matcher si précédé d'une lettre (évite "N-1" → -1)
        pattern = r"(?<![A-Za-z/])[-+]?\d[\d\s\xa0]*(?:[.,]\d+)?"
        results = []
        for match in re.finditer(pattern, text):
            raw = match.group().strip()
            # Normalize: remove thousands separators (spaces, non-breaking spaces)
            normalized = (
                raw.replace(" ", "")
                .replace("\xa0", "")
                .replace(" ", "")
                .replace(",", ".")
            )
            try:
                value = float(normalized)
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

    def _extract_allowed_from_sources(self, *source_texts: str) -> set:
        """
        Extrait tous les nombres présents dans les textes sources (PDF contexte,
        questionnaire brut, méthodologie, descriptions slides) et les retourne
        comme ensemble de valeurs autorisées.
        Ces nombres ne sont pas des KPIs mais sont des faits légitimes cités dans le
        document d'entrée — ils ne doivent pas être marqués comme hallucinations.
        """
        allowed = set()
        for text in source_texts:
            if not text:
                continue
            for value, _ in self._extract_numbers_with_context(text):
                allowed.add(round(value, 2))
                allowed.add(round(value, 0))
                try:
                    allowed.add(int(value) if value == int(value) else value)
                except (OverflowError, ValueError):
                    pass
        return allowed

    def _build_slide_descriptions_text(self) -> str:
        """
        Retourne les descriptions hardcodées des slides (PERFORMANCE_GLOBALE_SLIDES)
        comme texte source pour le whitelist — les benchmarks marché y sont définis.
        """
        try:
            from generation.llm_generator import PERFORMANCE_GLOBALE_SLIDES
            return "\n".join(
                slide.get("description", "") for slide in PERFORMANCE_GLOBALE_SLIDES
            )
        except Exception:
            return ""

    def audit(
        self,
        generated_content: str,
        kpi_dict: dict,
        context_text: str = "",
        questionnaire_raw_text: str = "",
        methodology_text: str = "",
    ) -> dict:
        """
        Audit generated narrative against computed KPI values.

        Args:
            generated_content:      LLM-generated text to audit.
            kpi_dict:               Dict of kpi_id -> KPI entry (from KPIEngine).
            context_text:           Raw text from the context PDF (optional).
                                    Numbers found here are whitelisted — they are
                                    legitimate source facts, not hallucinations.
            questionnaire_raw_text: Raw questionnaire text (optional, same logic).
            methodology_text:       Methodology markdown (benchmarks whitelisted).

        Returns:
            {
                "passed": bool,
                "score_pct": float,
                "total_numbers_found": int,
                "validated": int,
                "rejected": [...],
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

        # Nombres présents dans les documents sources → autorisés sans vérification KPI
        # Inclut :
        #   - PDF contexte (données pharmacie spécifiques)
        #   - Questionnaire brut
        #   - Méthodologie (benchmarks marché 40,8€, 180 clients/j…)
        #   - Descriptions slides (benchmarks hardcodés 350-380k€ CA/ETP, 13,89€ panier…)
        slide_descriptions_text = self._build_slide_descriptions_text()
        source_whitelist = self._extract_allowed_from_sources(
            context_text,
            questionnaire_raw_text,
            methodology_text,
            slide_descriptions_text,
        )

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

        # Nombres suivis de ° dans le texte original (ex: "360°", "45°")
        # → termes de marque ou degrés, pas des KPIs
        degree_numbers: set = set()
        for m in re.finditer(r"(\d[\d\s.,]*)\s*°", generated_content):
            raw = m.group(1).strip().replace(" ", "").replace(",", ".")
            try:
                degree_numbers.add(round(float(raw), 2))
            except ValueError:
                pass

        def is_trivial(n: float) -> bool:
            """Years, ordinals, and degree-suffixed numbers are not KPI hallucinations."""
            if round(n, 2) in degree_numbers:
                return True
            if n == int(n):
                ni = int(n)
                if 1900 <= ni <= 2100:  # années
                    return True
                if 1 <= ni <= 12:       # mois
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
            # 1. Vérifie d'abord si le nombre vient d'une source légitime
            in_whitelist = (
                round(number, 2) in source_whitelist
                or round(number, 0) in source_whitelist
                or (int(number) if number == int(number) else number) in source_whitelist
            )
            if in_whitelist or self._is_number_valid(number, kpi_dict):
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
