"""
Gestionnaire de projets Agence VU.

Chaque projet correspond à une pharmacie et regroupe toutes les données
du pipeline : extraction Lot 1, méthodologie, KPIs, slides générées, audit.

Structure sur disque :
  /data/projects/
    {project_id}/
      meta.json       ← nom, pharmacie, dates, statut
      lot1_data.json  ← données extraites (excel, pptx, images)
      methodology.txt ← texte de la méthodologie
      kpis.json       ← KPIs calculés
      slides.json     ← slides générées
      audit.json      ← rapport d'audit
      pptx_output.pptx ← PowerPoint final (si généré)
"""

import json
import shutil
import uuid
from datetime import datetime
from pathlib import Path
from typing import Optional

PROJECTS_DIR = Path("/data/projects")


# ── Helpers ─────────────────────────────────────────────────────────────────

def _ts() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _project_dir(project_id: str) -> Path:
    return PROJECTS_DIR / project_id


def _meta_path(project_id: str) -> Path:
    return _project_dir(project_id) / "meta.json"


def _read_json(path: Path) -> Optional[dict | list]:
    if path.exists():
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            return None
    return None


def _write_json(path: Path, data) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


# ── CRUD projets ─────────────────────────────────────────────────────────────

def list_projects() -> list[dict]:
    """
    Retourne la liste de tous les projets triés par date de modification décroissante.
    """
    PROJECTS_DIR.mkdir(parents=True, exist_ok=True)
    projects = []
    for meta_file in PROJECTS_DIR.glob("*/meta.json"):
        try:
            meta = json.loads(meta_file.read_text(encoding="utf-8"))
            projects.append(meta)
        except Exception:
            pass
    return sorted(projects, key=lambda p: p.get("date_modification", ""), reverse=True)


def create_project(nom: str, pharmacie: str) -> dict:
    """
    Crée un nouveau projet et retourne ses métadonnées.
    """
    project_id = str(uuid.uuid4())[:8]
    meta = {
        "id": project_id,
        "nom": nom,
        "pharmacie": pharmacie,
        "date_creation": _ts(),
        "date_modification": _ts(),
        "statut": "nouveau",          # nouveau | lot1 | methodologie | lot2 | termine
        "lot1_complet": False,
        "methodologie_complete": False,
        "lot2_complet": False,
    }
    _write_json(_meta_path(project_id), meta)
    return meta


def load_project(project_id: str) -> Optional[dict]:
    """
    Charge les métadonnées d'un projet.
    """
    return _read_json(_meta_path(project_id))


def update_project_meta(project_id: str, **kwargs) -> None:
    """
    Met à jour les métadonnées d'un projet (statut, flags, etc.).
    """
    meta = load_project(project_id) or {}
    meta.update(kwargs)
    meta["date_modification"] = _ts()
    _write_json(_meta_path(project_id), meta)


def delete_project(project_id: str) -> bool:
    """
    Supprime un projet et tous ses fichiers.
    """
    d = _project_dir(project_id)
    if d.exists():
        shutil.rmtree(d)
        return True
    return False


# ── Sauvegarde des données par couche ────────────────────────────────────────

def save_lot1_data(project_id: str, data: dict) -> None:
    _write_json(_project_dir(project_id) / "lot1_data.json", data)
    update_project_meta(project_id, lot1_complet=True, statut="lot1")


def load_lot1_data(project_id: str) -> Optional[dict]:
    return _read_json(_project_dir(project_id) / "lot1_data.json")


def save_methodology(project_id: str, text: str) -> None:
    path = _project_dir(project_id) / "methodology.txt"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")
    update_project_meta(project_id, methodologie_complete=bool(text.strip()), statut="methodologie")


def load_methodology(project_id: str) -> str:
    path = _project_dir(project_id) / "methodology.txt"
    if path.exists():
        return path.read_text(encoding="utf-8")
    # Fallback sur le fichier global (compatibilité ancienne version)
    global_path = Path("/data/methodology.txt")
    if global_path.exists():
        return global_path.read_text(encoding="utf-8")
    return ""


def save_kpis(project_id: str, kpis: dict) -> None:
    _write_json(_project_dir(project_id) / "kpis.json", kpis)


def load_kpis(project_id: str) -> Optional[dict]:
    return _read_json(_project_dir(project_id) / "kpis.json")


def save_slides(project_id: str, slides: list) -> None:
    _write_json(_project_dir(project_id) / "slides.json", slides)


def load_slides(project_id: str) -> Optional[list]:
    return _read_json(_project_dir(project_id) / "slides.json")


def save_audit(project_id: str, audit: dict) -> None:
    _write_json(_project_dir(project_id) / "audit.json", audit)


def load_audit(project_id: str) -> Optional[dict]:
    return _read_json(_project_dir(project_id) / "audit.json")


def save_pptx(project_id: str, pptx_bytes: bytes) -> Path:
    path = _project_dir(project_id) / "pptx_output.pptx"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(pptx_bytes)
    update_project_meta(project_id, lot2_complet=True, statut="termine")
    return path


def load_pptx(project_id: str) -> Optional[bytes]:
    path = _project_dir(project_id) / "pptx_output.pptx"
    if path.exists():
        return path.read_bytes()
    return None


# ── Session state ↔ projet ───────────────────────────────────────────────────

def session_to_project(project_id: str, session: dict) -> None:
    """
    Sauvegarde toutes les données de la session Streamlit dans le projet.
    """
    lot1 = {
        "excel":  session.get("lot1_excel_results", []),
        "pptx":   session.get("lot1_pptx_texts", {}),
        "images": session.get("lot1_image_results", []),
    }
    has_lot1 = any([lot1["excel"], lot1["pptx"], lot1["images"]])
    if has_lot1:
        save_lot1_data(project_id, lot1)

    methodo = session.get("methodology_text", "")
    if methodo:
        save_methodology(project_id, methodo)

    kpis = session.get("lot2_kpis")
    if kpis:
        save_kpis(project_id, kpis)

    slides = session.get("lot2_slides")
    if slides:
        save_slides(project_id, slides)

    audit = session.get("lot2_audit")
    if audit:
        save_audit(project_id, audit)

    pptx = session.get("lot2_pptx_bytes")
    if pptx:
        save_pptx(project_id, pptx)


def project_to_session(project_id: str, session: dict) -> dict:
    """
    Charge toutes les données d'un projet dans la session Streamlit.
    Retourne un dict des clés mises à jour.
    """
    updated = {}

    lot1 = load_lot1_data(project_id)
    if lot1:
        session["lot1_excel_results"]  = lot1.get("excel", [])
        session["lot1_pptx_texts"]     = lot1.get("pptx", {})
        session["lot1_image_results"]  = lot1.get("images", [])
        updated["lot1"] = True

    methodo = load_methodology(project_id)
    if methodo:
        session["methodology_text"] = methodo
        updated["methodology"] = True

    kpis = load_kpis(project_id)
    if kpis:
        session["lot2_kpis"] = kpis
        updated["kpis"] = True

    slides = load_slides(project_id)
    if slides:
        session["lot2_slides"] = slides
        updated["slides"] = True

    audit = load_audit(project_id)
    if audit:
        session["lot2_audit"] = audit
        updated["audit"] = True

    pptx = load_pptx(project_id)
    if pptx:
        session["lot2_pptx_bytes"] = pptx
        updated["pptx"] = True

    meta = load_project(project_id)
    if meta:
        session["lot2_pharmacy_name"] = meta.get("pharmacie", "")
        updated["meta"] = True

    return updated


# ── Utilitaires d'affichage ───────────────────────────────────────────────────

STATUT_LABELS = {
    "nouveau":       ("🆕", "Nouveau"),
    "lot1":          ("📊", "Lot 1 extrait"),
    "methodologie":  ("📋", "Méthodologie"),
    "lot2":          ("⚙️",  "Lot 2 en cours"),
    "termine":       ("✅", "Terminé"),
}


def statut_badge(statut: str) -> str:
    icon, label = STATUT_LABELS.get(statut, ("❓", statut))
    return f"{icon} {label}"


def project_summary(meta: dict) -> str:
    badge = statut_badge(meta.get("statut", "nouveau"))
    date  = meta.get("date_modification", "")[:10]
    return f"{meta['nom']} — {badge} — {date}"
