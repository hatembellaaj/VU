import json
import re
from typing import Optional

import anthropic


# ── 12 slides définis par la méthodologie Agence VU ──────────────────────────
# PG_00_CONTEXTE : généré uniquement si un PDF de contexte est fourni
PERFORMANCE_GLOBALE_SLIDES = [
    {
        "slide_id": "PG_00_CONTEXTE",
        "titre_defaut": "Contexte & Présentation de l'officine",
        "description": (
            "Slide d'introduction : localisation, historique, positionnement, enjeux. "
            "Basé exclusivement sur le document de contexte PDF fourni. "
            "Ne cite aucun KPI financier — uniquement des faits contextuels."
        ),
        "kpis_requis": [],
        "requires_context": True,
    },
    {
        "slide_id": "PG_01_SECTION",
        "titre_defaut": "Performances Globales",
        "description": (
            "Page de section introductive. Annonce que les analyses s'appuient sur des "
            "données marché consolidées : IQVIA, GERS data, OFFISANTÉ, Fiducial Conseil, "
            "KPMG Santé. Slide de transition — pas de chiffres officine, ton institutionnel."
        ),
        "kpis_requis": [],
    },
    {
        "slide_id": "PG_02_PROFIL_TYPE",
        "titre_defaut": "PROFIL-TYPE & Profil de votre officine",
        "description": (
            "Profil-type de l'officine : "
            "(1) type officine (urbain/périurbain/rural) d'après localisation, "
            "(2) part des ordonnances (TVA 2,1%) — si ≥ 70% → 'part ordonnances très élevée', "
            "(3) top 2-3 univers hors ordonnances par marge, "
            "(4) expertises potentielles détectées (ortho, MAD, naturo…). "
            "Structure : 4 points courts, 1 phrase chacun."
        ),
        "kpis_requis": ["part_ordonnances_pct", "ca_total"],
    },
    {
        "slide_id": "PG_03_PROFIL_PATIENTS",
        "titre_defaut": "Le PROFIL de vos patients",
        "description": (
            "Pyramide des âges officine vs benchmark local/national. "
            "Calculer la part des 60+ (cumul tranches) — si > 55% → 'profil très séniorisé'. "
            "Identifier la tranche sur-représentée et la tranche sous-représentée. "
            "Conclure par 1 implication stratégique (ex: 'peu de familles, peu d'actifs'). "
            "Si pas de données pyramide disponibles, décrire le profil depuis les données commerciales."
        ),
        "kpis_requis": [],
    },
    {
        "slide_id": "PG_04_FINANCIERS",
        "titre_defaut": "Les indicateurs FINANCIERS",
        "description": (
            "CA HT et marge brute avec évolution N-1 vs N. CA/ETP et Marge/ETP positionnés : "
            "< 90k€ faible | 90-105k€ correct | 105-120k€ performant | > 120k€ excellent. "
            "Alerte si médicaments chers (TVA 2,1% élevée) masquent la marge réelle. "
            "Benchmark CA/ETP marché : 350-380k€ | Marge brute : 400-540k€ / 27%. "
            "Utilise UNIQUEMENT les KPIs fournis."
        ),
        "kpis_requis": [
            "ca_total", "evolution_ca_pct",
            "marge_brute", "evolution_marge_pct",
            "ca_par_etp", "marge_par_etp",
        ],
    },
    {
        "slide_id": "PG_05_COMMERCIAUX",
        "titre_defaut": "Les indicateurs COMMERCIAUX",
        "description": (
            "Fréquentation (clients/jour) vs benchmark marché 180 clients/j. "
            "Panier moyen total vs benchmark 40,8€ — panier ordonnances vs 58,3€ — "
            "panier conseil vs 13,89€. "
            "Si fréquentation < 60% de la moyenne → 'fort levier de croissance par le flux'. "
            "Si panier conseil sous la moyenne → 'conseil associé non systématisé'. "
            "Identifier 1-2 leviers prioritaires."
        ),
        "kpis_requis": [
            "frequentation_j", "panier_moyen",
            "panier_ordonnances", "panier_conseil",
        ],
    },
    {
        "slide_id": "PG_06_UNIVERS_CA",
        "titre_defaut": "Les UNIVERS — CA et marge HT hors ordos",
        "description": (
            "Répartition CA% et Marge% par univers hors ordonnances "
            "(JAMBES, SÉNIOR, NATURE, LIBRE ACCÈS, HYGIÈNE, BÉBÉ, BEAUTÉ, VÉTO). "
            "Identifier piliers de marge (top 2 ≥ 15% marge chacun), relais (10-15%), "
            "univers para sur-exposés/sous-rentables. "
            "Mentionner la part ordonnances vs hors ordos dans le CA total."
        ),
        "kpis_requis": ["part_ordonnances_pct", "ca_total"],
    },
    {
        "slide_id": "PG_07_UNIVERS_EVO",
        "titre_defaut": "Les UNIVERS — Évolutions CA et marge",
        "description": (
            "Évolution CA% et marge% par univers hors ordos N vs N-1. "
            "Moteur de croissance principal (ordonnances ou hors ordos + chiffre %). "
            "Univers en croissance hors ordos (liste + %). "
            "Univers qui décrochent (CA < -10%) → qualifier 'levier manqué'. "
            "Conclusion : le hors ordos est-il un levier ou un frein ?"
        ),
        "kpis_requis": ["evolution_ca_pct", "evolution_marge_pct"],
    },
    {
        "slide_id": "PG_08_TOP_MARQUES",
        "titre_defaut": "Les UNIVERS — TOP 10 des marques par marge HT",
        "description": (
            "Top 10 marques hors ordonnances par marge HT. "
            "Part cumulée du top 10 dans la marge totale hors ordos. "
            "Si top 10 > 35% → 'portefeuille à concentrer pour gagner en puissance'. "
            "Source : ventes OSPHARM × catégorisation BCB VU. "
            "Recommandation 1 phrase."
        ),
        "kpis_requis": [],
    },
    {
        "slide_id": "PG_09_MERCH_EXPO",
        "titre_defaut": "Les indicateurs MERCHANDISING — Exposition",
        "description": (
            "Ratio exposition (linéaire%) vs marge% par univers. "
            "Sous-exposé (rouge) : marge >> exposition. "
            "Sur-exposé (orange) : exposition >> marge → immobilisent stocks/trésorerie. "
            "Signal spécifique si univers signature dispersé (NATURE, herboristerie…). "
            "3 bullets max : piliers sous-exposés | univers sur-exposés | cas particulier."
        ),
        "kpis_requis": [],
    },
    {
        "slide_id": "PG_10_MERCH_STOCKS",
        "titre_defaut": "Les indicateurs MERCHANDISING — Stocks",
        "description": (
            "Rotation stocks (jours) par univers. "
            "Benchmark : ordonnances 30-40j | hors ordos 80-100j. "
            "Global hors ordos > 100j → 'rotation élevée' (alerte). "
            "Identifier univers cohérents (45-80j) et univers > 150j → 'trésorerie dormante'. "
            "Conclure : 'pilotez vos stocks par univers'."
        ),
        "kpis_requis": [],
    },
    {
        "slide_id": "PG_11_MERCH_SIGNA",
        "titre_defaut": "Les indicateurs MERCHANDISING — Signalétique",
        "description": (
            "Identité visuelle : logo existant et différenciant → valoriser et recommander de décliner. "
            "Ce qui manque : balisage univers, cohérence charte graphique. "
            "Recommandation liée au projet de transfert/aménagement si mentionné. "
            "3 bullets : constat identité | ce qui manque | recommandation projet."
        ),
        "kpis_requis": [],
    },
    {
        "slide_id": "PG_12_SYNTHESE",
        "titre_defaut": "En Synthèse — Votre SWOT",
        "description": (
            "SWOT en 4 quadrants, 2-3 bullets max par quadrant. "
            "Forces : ce qui est au-dessus des benchmarks marché. "
            "Faiblesses : ce qui est en dessous ou déséquilibré. "
            "Opportunités : leviers identifiés (flux, conseil, univers sous-exposés). "
            "Menaces : dépendances (médicaments chers, séniorisation, univers décrochants). "
            "Utilise les KPIs et les constats des slides précédents."
        ),
        "kpis_requis": [
            "ca_total", "evolution_ca_pct",
            "marge_brute", "panier_moyen", "frequentation_j",
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
        image_results: list = None,
    ) -> list:
        """
        Generate all slides for the "Performance Globale" section (Part 2).

        Args:
            kpi_dict:      Computed KPIs from KPIEngine.
            methodology:   Methodology text from Lot 1.
            pharmacy_name: Name of the pharmacy.
            context_text:  Raw text extracted from the context PDF (optional).
                           If provided, generates an extra intro slide PG_00_CONTEXTE.
            image_results: List of image extraction results (for age pyramid chart).

        Returns:
            List of slide content dicts (PG_00 first if context provided).
        """
        from generation.chart_builder import build_chart_for_slide

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
                        "chart_data":    None,   # slide contexte → pas de graphique
                        "erreur":        None,
                    })
                except Exception as exc:
                    slides.append({
                        "slide_id":      slide_def["slide_id"],
                        "titre":         slide_def["titre_defaut"],
                        "contenu":       "",
                        "chiffres_cites": [],
                        "sources":       [],
                        "chart_data":    None,
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
            # Ajoute la spec de graphique (déterministique depuis les KPIs)
            result["chart_data"] = build_chart_for_slide(
                slide_id=slide_def["slide_id"],
                kpi_dict=kpi_dict,
                image_results=image_results,
            )
            slides.append(result)

        return slides
