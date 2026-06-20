"""
Estimation du coût API Anthropic avant traitement.

Tarifs claude-sonnet-4-6 (source: docs.anthropic.com) :
  - Tokens en entrée  : $3.00 / 1M tokens
  - Tokens en sortie  : $15.00 / 1M tokens

Pour les images, Anthropic redimensionne d'abord à 1568×1568 max,
puis découpe en tuiles de 512×512 (~1 750 tokens/tuile + 85 de base).
"""

import io
import math
from typing import Union

# ── Tarifs ──────────────────────────────────────────────────────────────────
INPUT_PRICE_PER_MTK  = 3.00   # $ par million de tokens en entrée
OUTPUT_PRICE_PER_MTK = 15.00  # $ par million de tokens en sortie

# Estimation tokens de sortie pour le parser image (valeurs extraites en JSON)
IMAGE_OUTPUT_TOKENS_ESTIMATE = 2_000
# Overhead du prompt texte (instructions + format JSON)
IMAGE_PROMPT_TOKENS = 350


def _resize_dims(width: int, height: int, max_dim: int = 1568) -> tuple[int, int]:
    """Redimensionne en gardant le ratio, borne à max_dim."""
    if width <= max_dim and height <= max_dim:
        return width, height
    ratio = min(max_dim / width, max_dim / height)
    return int(width * ratio), int(height * ratio)


def estimate_image_tokens(image_bytes: bytes) -> dict:
    """
    Calcule le nombre estimé de tokens pour une image.

    Returns:
        {
            "width": int, "height": int,
            "width_resized": int, "height_resized": int,
            "tiles": int,
            "input_tokens": int,   # image + prompt
            "output_tokens": int,
        }
    """
    try:
        from PIL import Image
        img = Image.open(io.BytesIO(image_bytes))
        width, height = img.size
    except Exception:
        # Fallback si PIL échoue : on suppose 1000×800
        width, height = 1000, 800

    rw, rh = _resize_dims(width, height)
    width_tiles  = math.ceil(rw / 512)
    height_tiles = math.ceil(rh / 512)
    tiles = width_tiles * height_tiles
    image_tokens = tiles * 1_750 + 85

    return {
        "width": width,
        "height": height,
        "width_resized": rw,
        "height_resized": rh,
        "tiles": tiles,
        "image_tokens": image_tokens,
        "input_tokens": image_tokens + IMAGE_PROMPT_TOKENS,
        "output_tokens": IMAGE_OUTPUT_TOKENS_ESTIMATE,
    }


def estimate_image_cost(image_bytes: bytes) -> dict:
    """
    Coût complet (entrée + sortie) pour parser une image.

    Returns:
        {
            ...(champs de estimate_image_tokens),
            "input_cost_usd": float,
            "output_cost_usd": float,
            "total_cost_usd": float,
        }
    """
    tokens = estimate_image_tokens(image_bytes)
    input_cost  = (tokens["input_tokens"]  / 1_000_000) * INPUT_PRICE_PER_MTK
    output_cost = (tokens["output_tokens"] / 1_000_000) * OUTPUT_PRICE_PER_MTK
    return {
        **tokens,
        "input_cost_usd":  input_cost,
        "output_cost_usd": output_cost,
        "total_cost_usd":  input_cost + output_cost,
    }


def format_cost(usd: float) -> str:
    """Formate un coût en $ avec suffisamment de décimales pour les petits montants."""
    if usd < 0.001:
        return f"${usd * 1000:.4f} m$"   # en millièmes de dollar
    if usd < 0.01:
        return f"${usd:.5f}"
    return f"${usd:.4f}"


def estimate_excel_info(file_bytes: bytes, filename: str = "") -> dict:
    """
    Résumé d'un fichier Excel (pas d'appel API, coût = $0).

    Returns:
        { "filename": str, "size_kb": float, "sheets": list[str], "cost_usd": 0.0 }
    """
    try:
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
        sheets = wb.sheetnames
        wb.close()
    except Exception:
        sheets = []

    return {
        "filename": filename,
        "size_kb": len(file_bytes) / 1024,
        "sheets": sheets,
        "cost_usd": 0.0,
    }
