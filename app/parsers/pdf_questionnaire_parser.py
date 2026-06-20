"""
Parseur de questionnaire PDF pour Agence VU.

Flux :
  1. Extrait le texte brut du PDF avec pdfplumber
  2. Envoie le texte à Claude qui le structure en dict {question: réponse}
  3. Retourne un dict compatible avec QuestionnaireParser.parse()

Utilisé dans Lot 2 pour ingérer les questionnaires terrain envoyés en PDF.
"""

import io
import json
import re
from typing import Optional

import anthropic

# ── Prompt Claude ─────────────────────────────────────────────────────────────

SYSTEM_PROMPT = """Tu es un assistant spécialisé dans l'analyse de questionnaires de pharmacies françaises.
Tu extrais les informations d'un texte brut issu d'un PDF et tu les structures en JSON."""

EXTRACTION_PROMPT = """Voici le texte brut extrait d'un questionnaire PDF rempli par une pharmacie.

```
{raw_text}
```

Extrait toutes les paires question/réponse et retourne un objet JSON **uniquement**, sans explication.
Format attendu :
{{
  "Nom de la pharmacie": "...",
  "Localisation (ville)": "...",
  "Nombre d'habitants": "...",
  "Type d'officine": "...",
  "Nombre d'ETP": "...",
  ...
}}

Règles :
- Inclure TOUTES les réponses présentes, même partielles
- Pour les cases à cocher : indiquer la valeur cochée
- Pour les champs vides : ne pas les inclure
- Les nombres restent en format numérique (pas de guillemets autour des chiffres)
- Retourne UNIQUEMENT le JSON, sans markdown, sans explication"""


# ── Parser ────────────────────────────────────────────────────────────────────

class PDFQuestionnaireParser:
    """Extrait et structure un questionnaire depuis un fichier PDF."""

    MAX_TEXT_CHARS = 40_000   # limite avant troncature

    def __init__(self, api_key: Optional[str] = None, model: str = "claude-sonnet-4-6"):
        self.api_key = api_key
        self.model   = model

    # ── Extraction texte ──────────────────────────────────────────────────────

    def extract_text(self, pdf_bytes: bytes) -> str:
        """Extrait le texte brut d'un PDF avec pdfplumber."""
        try:
            import pdfplumber
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                pages_text = []
                for i, page in enumerate(pdf.pages, start=1):
                    text = page.extract_text() or ""
                    if text.strip():
                        pages_text.append(f"--- Page {i} ---\n{text}")
                return "\n\n".join(pages_text)
        except ImportError:
            # Fallback pypdf si pdfplumber pas installé
            return self._extract_with_pypdf(pdf_bytes)
        except Exception as exc:
            raise RuntimeError(f"Impossible d'extraire le texte du PDF : {exc}") from exc

    def _extract_with_pypdf(self, pdf_bytes: bytes) -> str:
        try:
            import pypdf
            reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
            pages_text = []
            for i, page in enumerate(reader.pages, start=1):
                text = page.extract_text() or ""
                if text.strip():
                    pages_text.append(f"--- Page {i} ---\n{text}")
            return "\n\n".join(pages_text)
        except ImportError:
            raise RuntimeError(
                "Ni pdfplumber ni pypdf ne sont installés. "
                "Ajoutez 'pdfplumber' dans requirements.txt."
            )

    # ── Structuration via Claude ──────────────────────────────────────────────

    def _structure_with_llm(self, raw_text: str) -> dict:
        """Envoie le texte brut à Claude pour obtenir un dict structuré."""
        if not self.api_key:
            raise ValueError("Clé API Anthropic requise pour structurer le PDF questionnaire.")

        # Troncature si trop long
        if len(raw_text) > self.MAX_TEXT_CHARS:
            raw_text = raw_text[:self.MAX_TEXT_CHARS] + "\n[... tronqué ...]"

        client = anthropic.Anthropic(api_key=self.api_key)
        prompt = EXTRACTION_PROMPT.format(raw_text=raw_text)

        response = client.messages.create(
            model=self.model,
            max_tokens=4096,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": prompt}],
        )

        text = response.content[0].text.strip() if response.content else "{}"

        # Nettoyage des backticks éventuels
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text)

        try:
            return json.loads(text)
        except json.JSONDecodeError:
            # Dernière tentative : extraire le premier bloc JSON
            match = re.search(r"\{.*\}", text, re.DOTALL)
            if match:
                try:
                    return json.loads(match.group())
                except json.JSONDecodeError:
                    pass
            return {"_raw_text": raw_text[:2000], "_parse_error": "JSON invalide retourné par Claude"}

    # ── Interface principale ──────────────────────────────────────────────────

    def parse(self, pdf_bytes: bytes, use_llm: bool = True) -> dict:
        """
        Parse un PDF questionnaire.

        Args:
            pdf_bytes: contenu binaire du PDF
            use_llm: si True, utilise Claude pour structurer (recommandé)
                     si False, retourne le texte brut dans _raw_text

        Returns:
            dict {question: réponse} compatible avec QuestionnaireParser
        """
        raw_text = self.extract_text(pdf_bytes)

        if not raw_text.strip():
            raise ValueError(
                "Le PDF ne contient pas de texte extractible. "
                "S'il est scanné, une version numérique est nécessaire."
            )

        if use_llm and self.api_key:
            return self._structure_with_llm(raw_text)
        else:
            # Mode sans LLM : parsing heuristique basique (ligne: valeur)
            result = {}
            for line in raw_text.splitlines():
                line = line.strip()
                if ":" in line:
                    parts = line.split(":", 1)
                    key   = parts[0].strip()
                    value = parts[1].strip()
                    if key and value:
                        result[key] = value
            if not result:
                result["_raw_text"] = raw_text
            return result

    def estimate_cost(self, pdf_bytes: bytes) -> dict:
        """Estime le coût API avant structuration."""
        try:
            raw_text = self.extract_text(pdf_bytes)
        except Exception:
            raw_text = ""

        chars = min(len(raw_text), self.MAX_TEXT_CHARS)
        input_tokens  = (chars + len(EXTRACTION_PROMPT) + len(SYSTEM_PROMPT)) // 4
        output_tokens = 1_000  # JSON structuré estimé

        return {
            "pages_extracted": raw_text.count("--- Page "),
            "chars_extracted": len(raw_text),
            "input_tokens":   input_tokens,
            "output_tokens":  output_tokens,
            "input_cost_usd":  (input_tokens  / 1_000_000) * 3.00,
            "output_cost_usd": (output_tokens / 1_000_000) * 15.00,
            "total_cost_usd":  (input_tokens  / 1_000_000) * 3.00 + (output_tokens / 1_000_000) * 15.00,
        }
