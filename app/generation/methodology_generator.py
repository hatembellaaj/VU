"""
Générateur de méthodologie Agence VU.

À partir des textes extraits d'un PPTX exemple (lot1_pptx_texts),
Claude analyse les slides de la section "Performance Globale" et
induit les règles de transformation : sources, benchmarks, logique
narrative, seuils, ton.

La méthodologie générée est un document Markdown structuré,
sauvegardable et réutilisable pour tous les projets de l'agence.
"""

import anthropic

# ── Prompt système ────────────────────────────────────────────────────────────

SYSTEM_PROMPT = """Tu es un expert en analyse de présentations marketing pour les pharmacies françaises.
Tu travailles pour Agence VU, cabinet de conseil spécialisé dans les audits 360° officine.

Ton rôle : analyser le contenu extrait d'un PowerPoint exemple (textes bruts de chaque diapositive)
pour en déduire la méthodologie de transformation des données brutes en narratif de conseil.

Tu dois produire un document Markdown structuré et exhaustif qui servira d'instruction
réutilisable pour générer automatiquement de nouvelles présentations à partir de données financières."""

# ── Prompt utilisateur ────────────────────────────────────────────────────────

INDUCTION_PROMPT_TEMPLATE = """Voici le contenu extrait d'un PowerPoint exemple Agence VU (section "Performance Globale").
Chaque clé est le titre d'une diapositive, chaque valeur est la liste des textes présents sur cette slide.

```json
{slides_json}
```

---

Analyse ce contenu et génère une **MÉTHODOLOGIE COMPLÈTE** en Markdown avec les sections suivantes :

## 1. Structure générale
- Liste des slides de la section Performance Globale avec leur titre et thème
- Pour chaque slide : source de données principale (Excel, PNG image, questionnaire)

## 2. Règles de transformation par slide
Pour CHAQUE slide identifiée, décris précisément :
- **Input** : quelles données sont nécessaires (quel onglet Excel, quel type d'image, quels champs questionnaire)
- **Calculs** : quelles formules ou ratios sont appliqués
- **Benchmarks** : valeurs de référence marché utilisées (avec sources si détectables)
- **Seuils** : thresholds de qualification (bon/moyen/faible)
- **Règle narrative** : comment les données deviennent du texte (structure du message, conditions)

## 3. Ton et style rédactionnel
- Voix (tu/vous/troisième personne)
- Structure type par slide (Constat → Analyse → Recommandation ?)
- Adjectifs qualificatifs calibrés sur les performances
- Format des chiffres (unités, décimales, comparatifs)
- Longueur des bullets

## 4. Données de référence marché
Tableau récapitulatif de tous les benchmarks détectés (indicateur, valeur, source présumée).

## 5. Mapping des données requises

**SECTION CRITIQUE** — Cette section est parsée automatiquement par le pipeline.
Génère un bloc JSON EXACTEMENT dans ce format (respecter la syntaxe strictement) :

<!-- VU_MAPPING -->
```json
{{
  "ca_total":           {{"colonnes": ["CA Total", "Chiffre d'affaires HT", "CA HT"], "onglets": ["CA", "Financier", "Synthèse"]}},
  "evolution_ca_pct":   {{"colonnes": ["Évolution CA (%)", "Variation CA", "Croissance CA"], "onglets": ["CA", "Evolution"]}},
  "marge_brute":        {{"colonnes": ["Marge brute", "Marge HT", "MB"], "onglets": ["Marge", "CA", "Financier"]}},
  "evolution_marge_pct":{{"colonnes": ["Évolution marge (%)", "Variation marge"], "onglets": ["Marge", "Evolution"]}},
  "ca_par_etp":         {{"colonnes": ["CA/ETP", "CA par ETP"], "onglets": ["ETP", "RH", "CA"]}},
  "marge_par_etp":      {{"colonnes": ["Marge/ETP", "Marge par ETP"], "onglets": ["ETP", "RH", "Marge"]}},
  "frequentation_j":    {{"colonnes": ["Fréquentation/jour", "Clients/jour", "Passages/jour"], "onglets": ["Fréquentation", "Clients"]}},
  "panier_moyen":       {{"colonnes": ["Panier moyen", "Ticket moyen"], "onglets": ["Panier", "CA", "Commercial"]}},
  "panier_ordonnances": {{"colonnes": ["Panier ordonnances", "Panier ordo", "Ticket ordo"], "onglets": ["Panier", "Ordonnances"]}},
  "panier_conseil":     {{"colonnes": ["Panier conseil", "Panier hors ordo", "Ticket conseil"], "onglets": ["Panier", "Conseil"]}},
  "part_ordonnances_pct":{{"colonnes": ["Part ordonnances (%)", "% ordonnances", "Taux ordo"], "onglets": ["Ordonnances", "CA"]}},
  "nb_clients_actifs":  {{"colonnes": ["Clients actifs", "Nb clients", "Nombre clients"], "onglets": ["Clients", "Fidélisation"]}},
  "taux_fidelisation":  {{"colonnes": ["Taux fidélisation (%)", "Fidélisation", "Rétention"], "onglets": ["Fidélisation", "Clients"]}}
}}
```
<!-- /VU_MAPPING -->

**Adapte les noms de colonnes** si tu as détecté dans le PPTX des références à des champs Excel spécifiques.
Ajoute ou modifie les entrées selon ce que tu as observé dans les slides analysées.

---

Sois le plus précis et opérationnel possible. Ce document sera utilisé comme instruction
système pour une IA qui doit reproduire exactement le même style de présentation.
Ne laisse aucune règle implicite — tout doit être explicite et chiffré."""


# ── Générateur ────────────────────────────────────────────────────────────────

class MethodologyGenerator:
    """Génère une méthodologie structurée en Markdown depuis les slides d'un PPTX exemple."""

    MAX_SLIDES_CHARS = 80_000   # limite caractères pour ne pas exploser le contexte
    OUTPUT_MAX_TOKENS = 8_192   # méthodologie peut être longue

    def __init__(self, api_key: str, model: str = "claude-sonnet-4-6"):
        self.client = anthropic.Anthropic(api_key=api_key)
        self.model = model

    # ── Helpers ──────────────────────────────────────────────────────────────

    def _filter_slides(self, pptx_texts: dict) -> dict:
        """
        Garde uniquement les slides pertinentes pour la méthodologie.
        Priorité : slides contenant des mots-clés "Performance Globale".
        Si trop de contenu, tronque intelligemment.
        """
        import json

        # Essaie de sélectionner les slides de la section cible
        keywords = [
            "performance", "financier", "commercial", "univers", "merchandising",
            "profil", "patient", "panier", "fréquentation", "marge", "ca ",
            "rotation", "stock", "synthèse", "swot", "signalétique", "exposition",
        ]

        scored = {}
        for slide_name, lines in pptx_texts.items():
            text = " ".join(lines).lower()
            score = sum(1 for kw in keywords if kw in text)
            scored[slide_name] = (score, lines)

        # Trie par score décroissant, prend les plus pertinentes
        sorted_slides = sorted(scored.items(), key=lambda x: x[1][0], reverse=True)

        # Sélectionne jusqu'à la limite de caractères
        selected = {}
        total_chars = 0
        for slide_name, (score, lines) in sorted_slides:
            chunk = json.dumps({slide_name: lines}, ensure_ascii=False)
            if total_chars + len(chunk) > self.MAX_SLIDES_CHARS:
                break
            selected[slide_name] = lines
            total_chars += len(chunk)

        # Si rien sélectionné par keywords, prend toutes les slides dans la limite
        if not selected:
            for slide_name, lines in pptx_texts.items():
                chunk = json.dumps({slide_name: lines}, ensure_ascii=False)
                if total_chars + len(chunk) > self.MAX_SLIDES_CHARS:
                    break
                selected[slide_name] = lines
                total_chars += len(chunk)

        return selected

    # ── Estimation coût ──────────────────────────────────────────────────────

    def estimate_cost(self, pptx_texts: dict) -> dict:
        """
        Estime le coût API avant génération.
        """
        import json
        filtered = self._filter_slides(pptx_texts)
        slides_json = json.dumps(filtered, ensure_ascii=False, indent=1)

        prompt_text = INDUCTION_PROMPT_TEMPLATE.format(slides_json=slides_json)
        # ~4 caractères/token (estimation grossière)
        input_tokens = (len(SYSTEM_PROMPT) + len(prompt_text)) // 4
        output_tokens = self.OUTPUT_MAX_TOKENS

        input_cost  = (input_tokens  / 1_000_000) * 3.00
        output_cost = (output_tokens / 1_000_000) * 15.00

        return {
            "slides_selected": len(filtered),
            "slides_total":    len(pptx_texts),
            "input_tokens":    input_tokens,
            "output_tokens":   output_tokens,
            "input_cost_usd":  input_cost,
            "output_cost_usd": output_cost,
            "total_cost_usd":  input_cost + output_cost,
        }

    # ── Génération ────────────────────────────────────────────────────────────

    def generate(self, pptx_texts: dict) -> str:
        """
        Génère la méthodologie Markdown à partir des textes extraits du PPTX.

        Args:
            pptx_texts: dict {slide_name: [lines]} issu de lot1_pptx_texts

        Returns:
            Méthodologie complète en Markdown.

        Raises:
            ValueError: si pptx_texts est vide.
            anthropic.APIError: en cas d'erreur API.
        """
        import json

        if not pptx_texts:
            raise ValueError("Aucun texte PPTX fourni. Chargez d'abord un PPTX exemple dans Lot 1.")

        filtered = self._filter_slides(pptx_texts)
        slides_json = json.dumps(filtered, ensure_ascii=False, indent=1)

        prompt = INDUCTION_PROMPT_TEMPLATE.format(slides_json=slides_json)

        response = self.client.messages.create(
            model=self.model,
            max_tokens=self.OUTPUT_MAX_TOKENS,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": prompt}],
        )

        text = response.content[0].text if response.content else ""

        # Ajoute un header si Claude n'en a pas mis
        if text and not text.startswith("#"):
            text = "# Méthodologie Agence VU — Performance Globale\n\n" + text

        return text
