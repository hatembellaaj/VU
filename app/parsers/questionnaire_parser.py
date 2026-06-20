from typing import Any


class QuestionnaireParser:
    """Parses questionnaire/survey data and categorizes each answer by type."""

    # Known categorical enumerations for pharmacy surveys
    KNOWN_CATEGORIES = {
        "type_officine",
        "zone_geographique",
        "logiciel_gestion",
        "specialisation",
        "segment_clientele",
        "niveau_formation",
        "statut_juridique",
    }

    def _classify_value(self, key: str, value: Any) -> str:
        """Determine the category of a value."""
        if isinstance(value, bool):
            return "booleen"
        if isinstance(value, (int, float)):
            return "numerique"
        if isinstance(value, list):
            return "categoriel"
        if isinstance(value, str):
            # Check for boolean-like strings
            lower = value.strip().lower()
            if lower in ("oui", "non", "yes", "no", "true", "false", "vrai", "faux"):
                return "booleen"
            # Check for numeric string
            try:
                float(value.replace(",", ".").replace(" ", "").replace(" ", ""))
                return "numerique"
            except ValueError:
                pass
            # Check for known categorical keys
            if key.lower() in self.KNOWN_CATEGORIES:
                return "categoriel"
            # Short strings with no spaces are likely categorical
            if len(value) <= 50 and "\n" not in value:
                return "categoriel"
            return "texte_libre"
        # For any other type (None, dict, etc.)
        return "texte_libre"

    def _normalize_bool(self, value: Any) -> bool:
        """Normalize a boolean-like value to Python bool."""
        if isinstance(value, bool):
            return value
        if isinstance(value, str):
            lower = value.strip().lower()
            return lower in ("oui", "yes", "true", "vrai", "1")
        return bool(value)

    def _normalize_numeric(self, value: Any) -> float:
        """Normalize a numeric-like value to float."""
        if isinstance(value, (int, float)):
            return float(value)
        if isinstance(value, str):
            cleaned = (
                value.strip()
                .replace(",", ".")
                .replace(" ", "")
                .replace(" ", "")
                .replace("\xa0", "")
            )
            return float(cleaned)
        return float(value)

    def parse(self, data: dict) -> dict:
        """
        Parse questionnaire answers and categorize them by type.

        Args:
            data: Flat dictionary of question_key -> answer_value.

        Returns:
            {
                "numerique": {
                    key: {"valeur": float, "source": "questionnaire"}
                },
                "booleen": {
                    key: {"valeur": bool, "source": "questionnaire"}
                },
                "categoriel": {
                    key: {"valeur": str | list, "source": "questionnaire"}
                },
                "texte_libre": {
                    key: {"valeur": str, "source": "questionnaire"}
                }
            }
        """
        result = {
            "numerique": {},
            "booleen": {},
            "categoriel": {},
            "texte_libre": {},
        }

        for key, value in data.items():
            category = self._classify_value(key, value)

            if category == "numerique":
                try:
                    normalized = self._normalize_numeric(value)
                    result["numerique"][key] = {
                        "valeur": normalized,
                        "source": "questionnaire",
                    }
                except (ValueError, TypeError):
                    # Fallback to texte_libre if conversion fails
                    result["texte_libre"][key] = {
                        "valeur": str(value),
                        "source": "questionnaire",
                    }

            elif category == "booleen":
                result["booleen"][key] = {
                    "valeur": self._normalize_bool(value),
                    "source": "questionnaire",
                }

            elif category == "categoriel":
                normalized_val = value if isinstance(value, list) else str(value)
                result["categoriel"][key] = {
                    "valeur": normalized_val,
                    "source": "questionnaire",
                }

            else:  # texte_libre
                result["texte_libre"][key] = {
                    "valeur": str(value) if value is not None else "",
                    "source": "questionnaire",
                }

        return result
