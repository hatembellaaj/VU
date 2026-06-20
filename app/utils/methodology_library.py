"""
Bibliothèque de méthodologies Agence VU.

Chaque méthodologie est un fichier JSON dans /data/methodologies/.
Plusieurs méthodologies peuvent coexister (ex: une par type d'officine,
une par version de benchmark, etc.).

Structure d'un fichier :
  /data/methodologies/{id}.json
  {
    "id":                str   (slug court, ex: "alesienne-2025"),
    "nom":               str   (nom affiché, ex: "Alésienne 2025"),
    "description":       str   (optionnel),
    "source_pptx":       str   (nom du fichier PPTX source, optionnel),
    "date_creation":     str   (ISO),
    "date_modification": str   (ISO),
    "content":           str   (texte Markdown complet)
  }
"""

import json
import re
import shutil
import uuid
from datetime import datetime
from pathlib import Path
from typing import Optional

METHODO_DIR = Path("/data/methodologies")
GLOBAL_FALLBACK = Path("/data/methodology.txt")   # compatibilité ancienne version


# ── Helpers ───────────────────────────────────────────────────────────────────

def _ts() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _slugify(name: str) -> str:
    """Convertit un nom en slug alphanumérique court."""
    slug = name.lower().strip()
    slug = re.sub(r"[àáâãäå]", "a", slug)
    slug = re.sub(r"[èéêë]", "e", slug)
    slug = re.sub(r"[ìíîï]", "i", slug)
    slug = re.sub(r"[òóôõö]", "o", slug)
    slug = re.sub(r"[ùúûü]", "u", slug)
    slug = re.sub(r"[ç]", "c", slug)
    slug = re.sub(r"[^a-z0-9]+", "-", slug)
    slug = slug.strip("-")[:40]
    return slug or str(uuid.uuid4())[:8]


def _unique_id(name: str) -> str:
    """Génère un ID unique basé sur le slug du nom."""
    base = _slugify(name)
    METHODO_DIR.mkdir(parents=True, exist_ok=True)
    candidate = base
    i = 2
    while (METHODO_DIR / f"{candidate}.json").exists():
        candidate = f"{base}-{i}"
        i += 1
    return candidate


def _path(methodo_id: str) -> Path:
    return METHODO_DIR / f"{methodo_id}.json"


def _read(methodo_id: str) -> Optional[dict]:
    p = _path(methodo_id)
    if p.exists():
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            return None
    return None


def _write(data: dict) -> None:
    METHODO_DIR.mkdir(parents=True, exist_ok=True)
    _path(data["id"]).write_text(
        json.dumps(data, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


# ── CRUD ──────────────────────────────────────────────────────────────────────

def list_methodologies() -> list[dict]:
    """
    Retourne toutes les méthodologies triées par date de modification décroissante.
    Chaque entrée contient toutes les métadonnées SAUF le contenu (pour la perf).
    """
    METHODO_DIR.mkdir(parents=True, exist_ok=True)
    result = []

    for p in METHODO_DIR.glob("*.json"):
        try:
            data = json.loads(p.read_text(encoding="utf-8"))
            # On exclut le contenu pour alléger la liste
            meta = {k: v for k, v in data.items() if k != "content"}
            meta["char_count"] = len(data.get("content", ""))
            meta["word_count"] = len(data.get("content", "").split())
            result.append(meta)
        except Exception:
            pass

    # Fallback : si aucune méthodologie sauvegardée mais fichier global existant
    if not result and GLOBAL_FALLBACK.exists():
        try:
            content = GLOBAL_FALLBACK.read_text(encoding="utf-8")
            if content.strip():
                # Migre automatiquement le fichier global dans la bibliothèque
                save_methodology(
                    nom="Méthodologie par défaut",
                    content=content,
                    description="Migrée automatiquement depuis methodology.txt",
                )
                return list_methodologies()
        except Exception:
            pass

    return sorted(result, key=lambda x: x.get("date_modification", ""), reverse=True)


def save_methodology(
    nom: str,
    content: str,
    description: str = "",
    source_pptx: str = "",
    methodo_id: Optional[str] = None,
) -> dict:
    """
    Crée ou met à jour une méthodologie dans la bibliothèque.

    Si methodo_id est fourni → mise à jour de la méthodologie existante.
    Sinon → création avec un nouvel ID dérivé du nom.

    Retourne les métadonnées complètes (sans le contenu).
    """
    if methodo_id:
        existing = _read(methodo_id)
        if existing:
            existing.update({
                "nom":               nom,
                "description":       description or existing.get("description", ""),
                "source_pptx":       source_pptx or existing.get("source_pptx", ""),
                "date_modification": _ts(),
                "content":           content,
            })
            _write(existing)
            return {k: v for k, v in existing.items() if k != "content"}

    # Nouvelle méthodologie
    new_id = _unique_id(nom)
    data = {
        "id":                new_id,
        "nom":               nom,
        "description":       description,
        "source_pptx":       source_pptx,
        "date_creation":     _ts(),
        "date_modification": _ts(),
        "content":           content,
    }
    _write(data)
    return {k: v for k, v in data.items() if k != "content"}


def load_methodology(methodo_id: str) -> Optional[dict]:
    """Charge une méthodologie complète (métadonnées + contenu)."""
    return _read(methodo_id)


def get_content(methodo_id: str) -> str:
    """Retourne uniquement le contenu Markdown d'une méthodologie."""
    data = _read(methodo_id)
    return data.get("content", "") if data else ""


def rename_methodology(methodo_id: str, new_nom: str) -> bool:
    """Renomme une méthodologie (ne change pas son ID)."""
    data = _read(methodo_id)
    if not data:
        return False
    data["nom"] = new_nom
    data["date_modification"] = _ts()
    _write(data)
    return True


def delete_methodology(methodo_id: str) -> bool:
    """Supprime une méthodologie de la bibliothèque."""
    p = _path(methodo_id)
    if p.exists():
        p.unlink()
        return True
    return False


def duplicate_methodology(methodo_id: str, new_nom: str) -> Optional[dict]:
    """Duplique une méthodologie existante sous un nouveau nom."""
    data = _read(methodo_id)
    if not data:
        return None
    return save_methodology(
        nom=new_nom,
        content=data.get("content", ""),
        description=f"Copie de « {data['nom']} »",
        source_pptx=data.get("source_pptx", ""),
    )


# ── Utilitaires d'affichage ───────────────────────────────────────────────────

def summary_line(meta: dict) -> str:
    """Ligne courte pour un sélecteur dropdown."""
    date = meta.get("date_modification", "")[:10]
    words = meta.get("word_count", 0)
    return f"{meta['nom']} — {words} mots — {date}"
