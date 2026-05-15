"""
annotation_store.py
-------------------
Reads and writes the sidecar JSON file that persists annotation state
between sessions.

File location: same folder as the V1/V2 Excel files.
File name:     qa_run_{YYYY-MM-DD}_{v1_label}_{v2_label}.json

The store is append-only for annotations — writing a disposition for a
site+column key overwrites any previous entry for that key.
"""

from __future__ import annotations

import json
import re
import unicodedata
from datetime import datetime
from pathlib import Path


def _safe_label(label: str) -> str:
    """Convert an analyst-entered version label into a safe filename fragment."""
    label = unicodedata.normalize("NFKD", label)
    label = re.sub(r"[^\w\s-]", "", label)
    label = re.sub(r"[\s]+", "_", label.strip())
    return label[:30]  # Truncate to keep filenames manageable


def sidecar_path(
    project_folder: str | Path,
    run_date: str,          # "YYYY-MM-DD"
    v1_label: str,
    v2_label: str,
) -> Path:
    """Return the canonical path for a sidecar JSON file."""
    name = f"qa_run_{run_date}_{_safe_label(v1_label)}_vs_{_safe_label(v2_label)}.json"
    return Path(project_folder) / name


class AnnotationStore:
    """
    Manages loading, saving, and updating annotation state for one QA run.

    Usage
    -----
    store = AnnotationStore.load_or_create(path, metadata)
    store.set_annotation("Riverbend", "EF Natural Gas", "Error — fixing", "Wrong year")
    store.save()
    """

    DISPOSITION_EXPECTED = "Expected"
    DISPOSITION_ERROR = "Error — fixing"
    DISPOSITION_REVIEW = "Needs review"
    VALID_DISPOSITIONS = {DISPOSITION_EXPECTED, DISPOSITION_ERROR, DISPOSITION_REVIEW}

    def __init__(self, path: Path, data: dict):
        self._path = path
        self._data = data

    # ------------------------------------------------------------------
    # Construction
    # ------------------------------------------------------------------

    @classmethod
    def load_or_create(
        cls,
        path: Path,
        *,
        analyst: str,
        project: str,
        v1_path: str,
        v1_label: str,
        v2_path: str,
        v2_label: str,
        id_column: str,
        materiality_threshold: float,
        column_mapping: list[dict],
        structural_changes: dict,
    ) -> "AnnotationStore":
        """
        Load an existing sidecar file (restoring annotations) or create a
        new one with the provided run metadata.
        """
        if path.exists():
            with open(path, encoding="utf-8") as f:
                data = json.load(f)
            # Update mutable fields in case the run was resumed
            data["column_mapping"] = column_mapping
            data["structural_changes"] = structural_changes
        else:
            data = {
                "run_id": path.stem,
                "created": datetime.now().isoformat(timespec="seconds"),
                "analyst": analyst,
                "project": project,
                "v1": {"path": str(v1_path), "label": v1_label},
                "v2": {"path": str(v2_path), "label": v2_label},
                "id_column": id_column,
                "materiality_threshold": materiality_threshold,
                "column_mapping": column_mapping,
                "structural_changes": structural_changes,
                "annotations": {},
            }
        return cls(path, data)

    # ------------------------------------------------------------------
    # Annotation access
    # ------------------------------------------------------------------

    @staticmethod
    def annotation_key(site_id: str, column: str) -> str:
        """Canonical key format: 'SiteID::ColumnName'"""
        return f"{site_id}::{column}"

    def set_annotation(
        self,
        site_id: str,
        column: str,
        disposition: str,
        note: str = "",
    ) -> None:
        """
        Record (or overwrite) an annotation for a site+column flag.
        Auto-saves to disk after every call.

        Parameters
        ----------
        site_id : str
            The site identifier (from the ID column).
        column : str
            The column name as it appears in V1.
        disposition : str
            One of the DISPOSITION_* constants.
        note : str
            Optional analyst free-text note.
        """
        if disposition not in self.VALID_DISPOSITIONS:
            raise ValueError(
                f"Invalid disposition '{disposition}'. "
                f"Must be one of: {self.VALID_DISPOSITIONS}"
            )
        key = self.annotation_key(site_id, column)
        self._data["annotations"][key] = {
            "disposition": disposition,
            "note": note,
            "timestamp": datetime.now().isoformat(timespec="seconds"),
            "analyst": self._data.get("analyst", ""),
        }
        self.save()

    def get_annotation(self, site_id: str, column: str) -> dict | None:
        """Return the stored annotation for a site+column, or None."""
        key = self.annotation_key(site_id, column)
        return self._data["annotations"].get(key)

    def get_site_disposition(self, site_id: str) -> str | None:
        """
        Return the disposition for an entire site if ALL its columns share
        the same disposition; otherwise None.
        Useful for site-level annotation shortcuts.
        """
        site_annotations = {
            v["disposition"]
            for k, v in self._data["annotations"].items()
            if k.startswith(f"{site_id}::")
        }
        if len(site_annotations) == 1:
            return site_annotations.pop()
        return None

    def unanswered_high_impact_keys(self, high_impact_keys: list[str]) -> list[str]:
        """
        Given a list of site::column keys that are HIGH severity,
        return those that have no annotation yet.
        """
        answered = set(self._data["annotations"].keys())
        return [k for k in high_impact_keys if k not in answered]

    def all_annotations(self) -> dict:
        """Return a copy of the full annotations dict."""
        return dict(self._data["annotations"])

    # ------------------------------------------------------------------
    # Persistence
    # ------------------------------------------------------------------

    def save(self) -> None:
        """Write current state to disk (pretty-printed JSON)."""
        self._path.parent.mkdir(parents=True, exist_ok=True)
        with open(self._path, "w", encoding="utf-8") as f:
            json.dump(self._data, f, indent=2, ensure_ascii=False)

    def to_dict(self) -> dict:
        """Return a copy of the full sidecar data."""
        return dict(self._data)

    # ------------------------------------------------------------------
    # Convenience properties for the UI
    # ------------------------------------------------------------------

    @property
    def analyst(self) -> str:
        return self._data.get("analyst", "")

    @property
    def project(self) -> str:
        return self._data.get("project", "")

    @property
    def v1_label(self) -> str:
        return self._data["v1"]["label"]

    @property
    def v2_label(self) -> str:
        return self._data["v2"]["label"]

    @property
    def materiality_threshold(self) -> float:
        return float(self._data.get("materiality_threshold", 5.0))

    @property
    def id_column(self) -> str:
        return self._data.get("id_column", "")
