import json
import re
from typing import Optional

import anthropic


# Slide definitions for the "Performance Globale" section (Part 2)
# PG_00_CONTEXTE est généré uniquement si un PDF de contexte est fourni
PERFORMANCE_GLOBALE_SLIDES = [
    {
        "slide_id": "PG_00_CONTEXTE",
        "titre_defaut": "Contexte & Présentation de l'officine",
        "description": (
            "Slide d'introduction qui résume le contexte de la pharmacie : "
            "localisation, historique, positionnement, enjeux identifiés. "
            "Basé exclusivement sur le document de contexte fourni."
        ),
        "kpis_requis": [],           # pas de KPIs — basé sur le PDF contexte
        "requires_context": True,    # slide ignorée si pas de PDF contexte
    },
    {
        "slide_id": "PG_01_INTRO",
        "titre_defaut": "Performance Globale — Vue d'ensemble",
        "description": (
            "Introduction synthétique de la performance de la pharmacie. "
            "Présente les indicateurs clés et le positionnement général."
        ),
        "kpis_requis": ["ca_total", "evolution_ca_pct", "panier_moyen"],
    },
    {
        "slide_id": "PG_02_CA",
        "titre_defaut": "Analyse du Chiffre d'Affaires",
        "description": (
            "Analyse détaillée du chiffre d'affaires: total, évolution, "
            "répartition par famille de produits."
        ),
        "kpis_requis": [
            "ca_total",
            "evolution_ca_pct",
            "part_ordonnances_pct",
            "part_parapharmacie_pct",
        ],
    },
    {
        "slide_id": "PG_03_CLIENTELE",
        "titre_defaut": "Analyse de la Clientèle",
        "description": (
            "Profil et comportement de la clientèle: nombre de clients actifs, "
            "fréquentation, fidélisation."
        ),
        "kpis_requis": [
            "nb_clients_actifs",
            "frequentation_mensuelle",
            "taux_fidelisation",
            "panier_moyen",
        ],
    },
    {
        "slide_id": "PG_04_PANIER",
        "titre_defaut": "Évolution du Panier Moyen",
        "description": (
            "Analyse du panier moyen et de son évolution. "
            "Opportunités d'amélioration du panier."
        ),
        "kpis_requis": ["panier_moyen", "evolution_panier_pct", "nb_clients_actifs"],
    },
    {
        "slide_id": "PG_05_SAISONNALITE",
        "titre_defaut": "Saisonnalité et Tendances",
        "description": (
            "Analyse de la saisonnalité des ventes et identification "
            "des périodes clés de l'année."
        ),
        "kpis_requis": ["indice_saisonnalite", "frequentation_mensuelle", "ca_total"],
    },
    {
        "slide_id": "PG_06_SYNTHESE",
        "titre_defaut": "Synthèse et Recommandations",
        "description": (
            "Synthèse des forces et axes d'amélioration identifiés. "
            "Recommandations prioritaires pour la pharmacie."
        ),
        "kpis_requis": [
            "ca_total",
            "evolution_ca_pct",
            "taux_fidelisation",
            "panier_moyen",
        ],
    },
]


class LLMGenerator:
    """Generates narrative content for pharmacy performance slides using Claude."""

    def __init__(self, api_key: str, model: str):
        self.api_key = api_key
        self.model = model
        self.client = anthropic.Anthropic(api_key=api_key)

    def _format_kpi_block(self, kpi_dict: dict, required_kpi_ids: Optional[list] = None) -> str:
        """
        Format KPI values as a structured JSON block for injection into prompts.
        If required_kpi_ids is provided, only include those KPIs.
        """
        if required_kpi_ids:
            filtered = {k: v for k, v in kpi_dict.items() if k in required_kpi_ids}
        else:
            filtered = kpi_dict

        # Build a clean representation
        kpi_lines = []
        for kpi_id, kpi in filtered.items():
            valeur = kpi.get("valeur")
            unite = kpi.get("unite", "")
            label = kpi.get("label_fr", kpi_id)
            statut = kpi.get("statut", "inconnu")
            if valeur is not None:
                kpi_lines.append(
                    f'  "{label}": {{"valeur": {valeur}, "unite": "{unite}", "statut": "{statut}"}}'
                )
            else:
                kpi_lines.append(
                    f'  "{label}": {{"valeur": null, "unite": "{unite}", "statut": "inconnu"}}'
                )

        return "{\n" + ",\n".join(kpi_lines) + "\n}"

    def _build_context_prompt(
        self,
        context_text: str,
        methodology: str,
        pharmacy_name: str = "",
    ) -> str:
        """
        Prompt dédié pour le slide de contexte (PG_00_CONTEXTE).
        Basé uniquement sur le document de contexte — aucun KPI.
        """
        pharmacy_label = pharmacy_name or "la pharmacie"
        return f"""Tu es un expert en marketing pharmaceutique travaillant pour l'Agence VU.
Tu dois rédiger le contenu du **premier slide** de l'audit stratégique 360° de {pharmacy_label}.
Ce slide de contexte présente la pharmacie avant toute analyse de performance.

## DOCUMENT DE CONTEXTE (source exclusive)
```
{context_text[:6000]}
```

## MÉTHODOLOGIE AGENCE VU (ton et style)
{methodology if methodology else "Adopter un ton professionnel, factuel et orienté action."}

## CONSIGNE
Rédige un slide d'introduction qui synthétise :
- La présentation de la pharmacie (localisation, type, historique si mentionné)
- Le contexte stratégique (enjeux, projets en cours, points saillants)
- Les 2-3 axes d'analyse qui vont suivre dans le rapport

Utilise UNIQUEMENT les informations présentes dans le document de contexte.
N'invente aucun chiffre ni fait non mentionné.

## FORMAT DE RÉPONSE (JSON strict, aucun texte avant ou après)
{{
  "titre": "Titre du slide (ex: Pharmacie du Marché — Contexte & Enjeux)",
  "contenu": "Texte du slide en français (4-6 bullets maximum, style Agence VU)",
  "chiffres_cites": [],
  "sources": ["Document de contexte PDF"]
}}"""

    def _build_grounded_prompt(
        self,
        slide_id: str,
        kpi_dict: dict,
        methodology: str,
        slide_description: str = "",
        slide_title: str = "",
        pharmacy_name: str = "",
        required_kpi_ids: Optional[list] = None,
    ) -> str:
        """
        Build a strictly grounded prompt that injects all KPI values and
        instructs Claude to only use those values.

        Args:
            slide_id: Identifier of the slide to generate.
            kpi_dict: All computed KPIs.
            methodology: Methodology text from Lot 1 analysis.
            slide_description: What this slide should cover.
            slide_title: Suggested title.
            pharmacy_name: Name of the pharmacy.
            required_kpi_ids: KPI IDs specifically needed for this slide.

        Returns:
            Full prompt string.
        """
        kpi_block = self._format_kpi_block(kpi_dict, required_kpi_ids)
        pharmacy_label = pharmacy_name or "la pharmacie"

        prompt = f"""Tu es un expert en marketing pharmaceutique travaillant pour l'Agence VU.
Tu dois rédiger le contenu narratif d'une diapositive PowerPoint pour l'audit stratégique 360° de {pharmacy_label}.

## RÈGLE ABSOLUE — ANTI-HALLUCINATION
Tu ne peux utiliser QUE les valeurs numériques fournies dans le bloc KPI ci-dessous.
N'invente AUCUN chiffre. N'arrondis PAS sans utiliser le chiffre exact fourni.
Si une valeur KPI est null, ne mentionne PAS de chiffre pour cet indicateur.

## KPIs CALCULÉS (source de vérité)
```json
{kpi_block}
```

## MÉTHODOLOGIE AGENCE VU
{methodology if methodology else "Adopter un ton professionnel, factuel et orienté action. Structurer en constats puis recommandations."}

## DIAPOSITIVE À RÉDIGER
- Identifiant: {slide_id}
- Titre suggéré: {slide_title}
- Contenu attendu: {slide_description}

## FORMAT DE RÉPONSE (JSON strict, aucun texte avant ou après)
{{
  "titre": "Titre exact de la diapositive (court, impactant)",
  "contenu": "Texte de la diapositive en français (3-5 phrases maximum). Utilise UNIQUEMENT les chiffres du bloc KPI.",
  "chiffres_cites": [liste des valeurs numériques exactes utilisées dans le contenu],
  "sources": [liste des labels KPI utilisés]
}}"""

        return prompt

    def generate_slide(
        self,
        slide_id: str,
        kpi_dict: dict,
        methodology: str,
        slide_description: str = "",
        slide_title: str = "",
        pharmacy_name: str = "",
        required_kpi_ids: Optional[list] = None,
    ) -> dict:
        """
        Generate content for a single slide.

        Returns:
            {
                "slide_id": str,
                "titre": str,
                "contenu": str,
                "chiffres_cites": list,
                "sources": list,
                "erreur": None | str
            }
        """
        prompt = self._build_grounded_prompt(
            slide_id=slide_id,
            kpi_dict=kpi_dict,
            methodology=methodology,
            slide_description=slide_description,
            slide_title=slide_title,
            pharmacy_name=pharmacy_name,
            required_kpi_ids=required_kpi_ids,
        )

        try:
            response = self.client.messages.create(
                model=self.model,
                max_tokens=1024,
                messages=[{"role": "user", "content": prompt}],
            )

            raw_text = response.content[0].text.strip()

            # Extract JSON from markdown code blocks if present
            json_match = re.search(r"```(?:json)?\s*([\s\S]*?)```", raw_text)
            if json_match:
                raw_text = json_match.group(1).strip()

            parsed = json.loads(raw_text)

            return {
                "slide_id": slide_id,
                "titre": parsed.get("titre", slide_title),
                "contenu": parsed.get("contenu", ""),
                "chiffres_cites": parsed.get("chiffres_cites", []),
                "sources": parsed.get("sources", []),
                "erreur": None,
            }

        except json.JSONDecodeError as exc:
            return {
                "slide_id": slide_id,
                "titre": slide_title,
                "contenu": raw_text if "raw_text" in dir() else "",
                "chiffres_cites": [],
                "sources": [],
                "erreur": f"Réponse non-JSON: {str(exc)}",
            }
        except anthropic.APIError as exc:
            return {
                "slide_id": slide_id,
                "titre": slide_title,
                "contenu": "",
                "chiffres_cites": [],
                "sources": [],
                "erreur": f"Erreur API: {str(exc)}",
            }
        except Exception as exc:
            return {
                "slide_id": slide_id,
                "titre": slide_title,
                "contenu": "",
                "chiffres_cites": [],
                "sources": [],
                "erreur": f"Erreur inattendue: {str(exc)}",
            }

    def generate_performance_globale(
        self,
        kpi_dict: dict,
        methodology: str,
        pharmacy_name: str = "",
        context_text: str = "",
    ) -> list:
        """
        Generate all slides for the "Performance Globale" section (Part 2).

        Args:
            kpi_dict: Computed KPIs from KPIEngine.
            methodology: Methodology text from Lot 1.
            pharmacy_name: Name of the pharmacy.
            context_text: Raw text extracted from the context PDF (optional).
                          If provided, generates an extra intro slide PG_00_CONTEXTE.

        Returns:
            List of slide content dicts (PG_00 first if context provided).
        """
        slides = []
        for slide_def in PERFORMANCE_GLOBALE_SLIDES:

            # ── Slide contexte — traitement spécial ──────────────────────────
            if slide_def.get("requires_context"):
                if not context_text:
                    continue   # ignore PG_00 si pas de PDF contexte fourni
                try:
                    prompt = self._build_context_prompt(
                        context_text=context_text,
                        methodology=methodology,
                        pharmacy_name=pharmacy_name,
                    )
                    response = self.client.messages.create(
                        model=self.model,
                        max_tokens=1024,
                        messages=[{"role": "user", "content": prompt}],
                    )
                    raw = response.content[0].text.strip()
                    json_match = re.search(r"```(?:json)?\s*([\s\S]*?)```", raw)
                    if json_match:
                        raw = json_match.group(1).strip()
                    parsed = json.loads(raw)
                    slides.append({
                        "slide_id":      slide_def["slide_id"],
                        "titre":         parsed.get("titre", slide_def["titre_defaut"]),
                        "contenu":       parsed.get("contenu", ""),
                        "chiffres_cites": [],
                        "sources":       parsed.get("sources", ["Document de contexte PDF"]),
                        "erreur":        None,
                    })
                except Exception as exc:
                    slides.append({
                        "slide_id":      slide_def["slide_id"],
                        "titre":         slide_def["titre_defaut"],
                        "contenu":       "",
                        "chiffres_cites": [],
                        "sources":       [],
                        "erreur":        f"Erreur slide contexte: {exc}",
                    })
                continue

            # ── Slides KPI standards ──────────────────────────────────────────
            result = self.generate_slide(
                slide_id=slide_def["slide_id"],
                kpi_dict=kpi_dict,
                methodology=methodology,
                slide_description=slide_def["description"],
                slide_title=slide_def["titre_defaut"],
                pharmacy_name=pharmacy_name,
                required_kpi_ids=slide_def.get("kpis_requis"),
            )
            slides.append(result)

        return slides
