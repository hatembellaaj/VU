import base64
import io
import json
import re
from typing import Union

import anthropic


class ImageParser:
    """Parses pharmacy dashboard images using Claude Vision to extract numeric values."""

    def __init__(self, api_key: str, model: str = "claude-sonnet-4-6"):
        self.api_key = api_key
        self.model = model
        self.client = anthropic.Anthropic(api_key=api_key)

    def _encode_image(self, image_path_or_bytes: Union[str, bytes, io.BytesIO]) -> tuple:
        """Encode image to base64 and determine media type."""
        if isinstance(image_path_or_bytes, (bytes, io.BytesIO)):
            if isinstance(image_path_or_bytes, bytes):
                raw = image_path_or_bytes
            else:
                raw = image_path_or_bytes.read()
            # Detect format from magic bytes
            if raw[:4] == b'\x89PNG':
                media_type = "image/png"
            elif raw[:2] in (b'\xff\xd8', b'FF'):
                media_type = "image/jpeg"
            elif raw[:4] == b'GIF8':
                media_type = "image/gif"
            elif raw[:4] == b'RIFF':
                media_type = "image/webp"
            else:
                media_type = "image/png"  # default
            encoded = base64.standard_b64encode(raw).decode("utf-8")
        else:
            file_path = str(image_path_or_bytes)
            lower = file_path.lower()
            if lower.endswith(".jpg") or lower.endswith(".jpeg"):
                media_type = "image/jpeg"
            elif lower.endswith(".gif"):
                media_type = "image/gif"
            elif lower.endswith(".webp"):
                media_type = "image/webp"
            else:
                media_type = "image/png"
            with open(file_path, "rb") as f:
                raw = f.read()
            encoded = base64.standard_b64encode(raw).decode("utf-8")

        return encoded, media_type

    def parse(
        self,
        image_path_or_bytes: Union[str, bytes, io.BytesIO],
        filename: str = "",
    ) -> dict:
        """
        Parse a pharmacy dashboard image and extract all visible numeric values.

        Args:
            image_path_or_bytes: Path to image file, raw bytes, or BytesIO.
            filename: Original filename for reference.

        Returns:
            {
                "source": filename,
                "valeurs_extraites": [
                    {
                        "label": str,
                        "valeur": float | None,
                        "unite": str | None,
                        "confiance": "haute" | "moyenne" | "faible"
                    },
                    ...
                ],
                "erreur": None | str
            }
        """
        source = filename or (
            str(image_path_or_bytes)
            if isinstance(image_path_or_bytes, str)
            else "uploaded_image"
        )

        try:
            encoded, media_type = self._encode_image(image_path_or_bytes)
        except Exception as exc:
            return {
                "source": source,
                "valeurs_extraites": [],
                "erreur": f"Erreur d'encodage de l'image: {str(exc)}",
            }

        prompt = (
            "Tu es un expert en analyse d'images de tableaux de bord pharmaceutiques.\n\n"
            "INSTRUCTIONS STRICTES:\n"
            "1. Extrais UNIQUEMENT les valeurs numériques VISIBLES dans cette image.\n"
            "2. Pour chaque valeur numérique visible, retourne: le libellé (label), "
            "la valeur numérique exacte, l'unité si visible (€, %, nb, etc.).\n"
            "3. NE CALCULE RIEN. Ne déduis aucune valeur. N'invente rien.\n"
            "4. Si une valeur est floue, partiellement visible ou incertaine, "
            'mets valeur=null et confiance="faible".\n'
            "5. Niveau de confiance: "
            '"haute" = clairement visible, '
            '"moyenne" = lisible mais petite taille ou légère ambiguïté, '
            '"faible" = flou ou incertain.\n\n'
            "FORMAT DE RÉPONSE (JSON uniquement, aucun texte avant ou après):\n"
            "{\n"
            '  "valeurs_extraites": [\n'
            "    {\n"
            '      "label": "Chiffre d\'affaires total",\n'
            '      "valeur": 1234567.89,\n'
            '      "unite": "€",\n'
            '      "confiance": "haute"\n'
            "    },\n"
            "    {\n"
            '      "label": "Taux de fidélisation",\n'
            '      "valeur": null,\n'
            '      "unite": "%",\n'
            '      "confiance": "faible"\n'
            "    }\n"
            "  ]\n"
            "}"
        )

        try:
            response = self.client.messages.create(
                model=self.model,
                max_tokens=4096,
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "image",
                                "source": {
                                    "type": "base64",
                                    "media_type": media_type,
                                    "data": encoded,
                                },
                            },
                            {
                                "type": "text",
                                "text": prompt,
                            },
                        ],
                    }
                ],
            )

            raw_text = response.content[0].text.strip()

            # Strategy 1: complete markdown block  ```json ... ```
            json_match = re.search(r"```(?:json)?\s*([\s\S]*?)```", raw_text)
            if json_match:
                raw_text = json_match.group(1).strip()
            else:
                # Strategy 2: opening ``` without closing (truncated response)
                raw_text = re.sub(r"^```(?:json)?\s*", "", raw_text).strip()

            # Strategy 3: find the JSON object boundaries as last resort
            if not raw_text.startswith("{"):
                start = raw_text.find("{")
                end = raw_text.rfind("}")
                if start != -1 and end > start:
                    raw_text = raw_text[start : end + 1]

            parsed = json.loads(raw_text)
            valeurs = parsed.get("valeurs_extraites", [])

            # Normalize and validate entries
            normalized = []
            for entry in valeurs:
                label = str(entry.get("label", "")).strip()
                valeur = entry.get("valeur")
                unite = entry.get("unite")
                confiance = entry.get("confiance", "moyenne")

                if valeur is not None:
                    try:
                        valeur = float(valeur)
                    except (TypeError, ValueError):
                        valeur = None
                        confiance = "faible"

                if confiance not in ("haute", "moyenne", "faible"):
                    confiance = "moyenne"

                normalized.append(
                    {
                        "label": label,
                        "valeur": valeur,
                        "unite": str(unite).strip() if unite is not None else None,
                        "confiance": confiance,
                    }
                )

            return {
                "source": source,
                "valeurs_extraites": normalized,
                "erreur": None,
            }

        except json.JSONDecodeError as exc:
            return {
                "source": source,
                "valeurs_extraites": [],
                "erreur": f"Réponse Claude non parseable en JSON: {str(exc)}. "
                f"Réponse brute: {raw_text[:200]}",
            }
        except anthropic.APIError as exc:
            return {
                "source": source,
                "valeurs_extraites": [],
                "erreur": f"Erreur API Anthropic: {str(exc)}",
            }
        except Exception as exc:
            return {
                "source": source,
                "valeurs_extraites": [],
                "erreur": f"Erreur inattendue: {str(exc)}",
            }
