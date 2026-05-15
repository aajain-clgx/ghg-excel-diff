"""
tests/unit/test_annotation_store.py
-------------------------------------
Unit tests for annotation_store.py — complete coverage of:
  _safe_label, sidecar_path, AnnotationStore construction, all methods,
  all properties, persistence, and edge cases.
"""

import json
import pytest
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent.parent / "app"))

from annotation_store import (
    _safe_label,
    sidecar_path,
    AnnotationStore,
)


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def _make_store(tmp_path, *, analyst="Alice", project="GHG Test",
                v1_label="Draft 1", v2_label="Draft 2",
                path_override=None):
    sp = path_override or sidecar_path(tmp_path, "2026-01-01", v1_label, v2_label)
    return AnnotationStore.load_or_create(
        sp,
        analyst=analyst,
        project=project,
        v1_path="/v1.xlsx",
        v1_label=v1_label,
        v2_path="/v2.xlsx",
        v2_label=v2_label,
        id_column="Site ID",
        materiality_threshold=5.0,
        column_mapping=[],
        structural_changes={},
    )


# ─────────────────────────────────────────────────────────────────────────────
# _safe_label
# ─────────────────────────────────────────────────────────────────────────────

class TestSafeLabel:
    def test_spaces_become_underscores(self):
        assert _safe_label("Draft 1") == "Draft_1"

    def test_special_chars_removed(self):
        # Characters like / and : should be stripped
        assert "/" not in _safe_label("Q1/2025")
        assert ":" not in _safe_label("12:00 PM")

    def test_truncated_to_30(self):
        long = "A" * 50
        assert len(_safe_label(long)) <= 30

    def test_empty_string(self):
        assert _safe_label("") == ""

    def test_already_safe(self):
        assert _safe_label("Draft_1") == "Draft_1"

    def test_hyphen_preserved(self):
        # Hyphens are allowed ([\w\s-] pattern)
        result = _safe_label("Draft-1")
        assert "-" in result

    def test_unicode_normalized(self):
        # Accented characters decomposed and stripped
        result = _safe_label("Café")
        assert isinstance(result, str)


# ─────────────────────────────────────────────────────────────────────────────
# sidecar_path
# ─────────────────────────────────────────────────────────────────────────────

class TestSidecarPath:
    def test_returns_path(self, tmp_path):
        p = sidecar_path(tmp_path, "2026-01-01", "V1", "V2")
        assert isinstance(p, Path)

    def test_inside_project_folder(self, tmp_path):
        p = sidecar_path(tmp_path, "2026-01-01", "V1", "V2")
        assert p.parent == tmp_path

    def test_filename_contains_date(self, tmp_path):
        p = sidecar_path(tmp_path, "2026-01-01", "V1", "V2")
        assert "2026-01-01" in p.name

    def test_filename_has_json_extension(self, tmp_path):
        p = sidecar_path(tmp_path, "2026-01-01", "V1", "V2")
        assert p.suffix == ".json"

    def test_filename_contains_both_labels(self, tmp_path):
        p = sidecar_path(tmp_path, "2026-01-01", "Draft_1", "Draft_2")
        assert "Draft_1" in p.name
        assert "Draft_2" in p.name

    def test_string_path_accepted(self, tmp_path):
        p = sidecar_path(str(tmp_path), "2026-01-01", "A", "B")
        assert isinstance(p, Path)


# ─────────────────────────────────────────────────────────────────────────────
# AnnotationStore construction
# ─────────────────────────────────────────────────────────────────────────────

class TestAnnotationStoreConstruct:
    def test_creates_new_file(self, tmp_path):
        store = _make_store(tmp_path)
        store.save()  # load_or_create is in-memory; save persists to disk
        sp = sidecar_path(tmp_path, "2026-01-01", "Draft 1", "Draft 2")
        assert sp.exists()

    def test_new_store_has_empty_annotations(self, tmp_path):
        store = _make_store(tmp_path)
        assert store.all_annotations() == {}

    def test_metadata_stored_correctly(self, tmp_path):
        store = _make_store(tmp_path, analyst="Bob", project="Test Project")
        assert store.analyst == "Bob"
        assert store.project == "Test Project"

    def test_v1_v2_labels_stored(self, tmp_path):
        store = _make_store(tmp_path, v1_label="Prior Year", v2_label="Current Year")
        assert store.v1_label == "Prior Year"
        assert store.v2_label == "Current Year"

    def test_materiality_property(self, tmp_path):
        store = _make_store(tmp_path)
        assert store.materiality_threshold == 5.0

    def test_id_column_property(self, tmp_path):
        store = _make_store(tmp_path)
        assert store.id_column == "Site ID"

    def test_load_existing_file_restores_annotations(self, tmp_path):
        sp = sidecar_path(tmp_path, "2026-01-01", "V1", "V2")
        s1 = AnnotationStore.load_or_create(
            sp, analyst="A", project="P",
            v1_path="/a", v1_label="V1", v2_path="/b", v2_label="V2",
            id_column="ID", materiality_threshold=5.0,
            column_mapping=[], structural_changes={},
        )
        s1.set_annotation("Site-A", "EF", AnnotationStore.DISPOSITION_EXPECTED, "ok")
        # Load again — should see the annotation
        s2 = AnnotationStore.load_or_create(
            sp, analyst="A", project="P",
            v1_path="/a", v1_label="V1", v2_path="/b", v2_label="V2",
            id_column="ID", materiality_threshold=5.0,
            column_mapping=[], structural_changes={},
        )
        assert s2.get_annotation("Site-A", "EF") is not None

    def test_load_updates_column_mapping(self, tmp_path):
        sp = sidecar_path(tmp_path, "2026-01-01", "V1", "V2")
        mapping_v1 = [{"v1": "A", "v2": "A", "status": "matched", "confirmed": True, "manual": False}]
        mapping_v2 = [{"v1": "B", "v2": "B", "status": "matched", "confirmed": True, "manual": False}]
        AnnotationStore.load_or_create(
            sp, analyst="A", project="P",
            v1_path="/a", v1_label="V1", v2_path="/b", v2_label="V2",
            id_column="ID", materiality_threshold=5.0,
            column_mapping=mapping_v1, structural_changes={},
        )
        s2 = AnnotationStore.load_or_create(
            sp, analyst="A", project="P",
            v1_path="/a", v1_label="V1", v2_path="/b", v2_label="V2",
            id_column="ID", materiality_threshold=5.0,
            column_mapping=mapping_v2, structural_changes={},
        )
        assert s2.to_dict()["column_mapping"] == mapping_v2


# ─────────────────────────────────────────────────────────────────────────────
# set_annotation / get_annotation
# ─────────────────────────────────────────────────────────────────────────────

class TestSetGetAnnotation:
    def test_set_and_get_round_trip(self, tmp_path):
        s = _make_store(tmp_path)
        s.set_annotation("SiteA", "EF", AnnotationStore.DISPOSITION_EXPECTED, "looks good")
        ann = s.get_annotation("SiteA", "EF")
        assert ann is not None
        assert ann["disposition"] == AnnotationStore.DISPOSITION_EXPECTED
        assert ann["note"] == "looks good"

    def test_get_missing_returns_none(self, tmp_path):
        s = _make_store(tmp_path)
        assert s.get_annotation("NoSite", "NoCol") is None

    def test_overwrite_annotation(self, tmp_path):
        s = _make_store(tmp_path)
        s.set_annotation("S", "C", AnnotationStore.DISPOSITION_EXPECTED, "first")
        s.set_annotation("S", "C", AnnotationStore.DISPOSITION_ERROR, "second")
        ann = s.get_annotation("S", "C")
        assert ann["disposition"] == AnnotationStore.DISPOSITION_ERROR
        assert ann["note"] == "second"

    def test_annotation_has_timestamp(self, tmp_path):
        s = _make_store(tmp_path)
        s.set_annotation("S", "C", AnnotationStore.DISPOSITION_REVIEW, "")
        ann = s.get_annotation("S", "C")
        assert "timestamp" in ann

    def test_annotation_has_analyst(self, tmp_path):
        s = _make_store(tmp_path, analyst="Charlie")
        s.set_annotation("S", "C", AnnotationStore.DISPOSITION_EXPECTED, "")
        ann = s.get_annotation("S", "C")
        assert ann["analyst"] == "Charlie"

    def test_all_valid_dispositions(self, tmp_path):
        s = _make_store(tmp_path)
        for i, disp in enumerate(AnnotationStore.VALID_DISPOSITIONS):
            s.set_annotation(f"Site{i}", "Col", disp, "")
            assert s.get_annotation(f"Site{i}", "Col")["disposition"] == disp

    def test_invalid_disposition_raises(self, tmp_path):
        s = _make_store(tmp_path)
        with pytest.raises(ValueError, match="Invalid disposition"):
            s.set_annotation("S", "C", "NotADisposition", "")

    def test_empty_note_allowed(self, tmp_path):
        s = _make_store(tmp_path)
        s.set_annotation("S", "C", AnnotationStore.DISPOSITION_EXPECTED, "")
        assert s.get_annotation("S", "C")["note"] == ""

    def test_multiple_sites_independent(self, tmp_path):
        s = _make_store(tmp_path)
        s.set_annotation("S1", "EF", AnnotationStore.DISPOSITION_EXPECTED, "fine")
        s.set_annotation("S2", "EF", AnnotationStore.DISPOSITION_ERROR, "wrong")
        assert s.get_annotation("S1", "EF")["disposition"] == AnnotationStore.DISPOSITION_EXPECTED
        assert s.get_annotation("S2", "EF")["disposition"] == AnnotationStore.DISPOSITION_ERROR

    def test_auto_saves_to_disk(self, tmp_path):
        # _make_store uses "Draft 1"/"Draft 2" labels — use the same path
        s = _make_store(tmp_path)
        s.set_annotation("S", "C", AnnotationStore.DISPOSITION_EXPECTED, "x")
        sp = sidecar_path(tmp_path, "2026-01-01", "Draft 1", "Draft 2")
        raw = json.loads(sp.read_text())
        key = AnnotationStore.annotation_key("S", "C")
        assert key in raw["annotations"]

    def test_annotation_key_format(self):
        assert AnnotationStore.annotation_key("Site-A", "EF") == "Site-A::EF"
        assert AnnotationStore.annotation_key("X", "Y") == "X::Y"


# ─────────────────────────────────────────────────────────────────────────────
# get_site_disposition
# ─────────────────────────────────────────────────────────────────────────────

class TestGetSiteDisposition:
    def test_returns_none_when_no_annotations(self, tmp_path):
        s = _make_store(tmp_path)
        assert s.get_site_disposition("S1") is None

    def test_returns_disposition_when_all_same(self, tmp_path):
        s = _make_store(tmp_path)
        s.set_annotation("S1", "EF", AnnotationStore.DISPOSITION_EXPECTED, "")
        s.set_annotation("S1", "Activity", AnnotationStore.DISPOSITION_EXPECTED, "")
        assert s.get_site_disposition("S1") == AnnotationStore.DISPOSITION_EXPECTED

    def test_returns_none_when_mixed(self, tmp_path):
        s = _make_store(tmp_path)
        s.set_annotation("S1", "EF", AnnotationStore.DISPOSITION_EXPECTED, "")
        s.set_annotation("S1", "Activity", AnnotationStore.DISPOSITION_ERROR, "")
        assert s.get_site_disposition("S1") is None

    def test_single_annotation_returns_its_disposition(self, tmp_path):
        s = _make_store(tmp_path)
        s.set_annotation("S1", "EF", AnnotationStore.DISPOSITION_REVIEW, "")
        assert s.get_site_disposition("S1") == AnnotationStore.DISPOSITION_REVIEW

    def test_other_sites_not_affected(self, tmp_path):
        s = _make_store(tmp_path)
        s.set_annotation("S1", "EF", AnnotationStore.DISPOSITION_EXPECTED, "")
        assert s.get_site_disposition("S2") is None


# ─────────────────────────────────────────────────────────────────────────────
# unanswered_high_impact_keys
# ─────────────────────────────────────────────────────────────────────────────

class TestUnansweredHighImpact:
    def test_all_unanswered_when_empty(self, tmp_path):
        s = _make_store(tmp_path)
        keys = ["S1::EF", "S2::Activity"]
        assert s.unanswered_high_impact_keys(keys) == keys

    def test_answered_keys_removed(self, tmp_path):
        s = _make_store(tmp_path)
        s.set_annotation("S1", "EF", AnnotationStore.DISPOSITION_EXPECTED, "")
        keys = ["S1::EF", "S2::Activity"]
        unanswered = s.unanswered_high_impact_keys(keys)
        assert "S1::EF" not in unanswered
        assert "S2::Activity" in unanswered

    def test_empty_input_returns_empty(self, tmp_path):
        s = _make_store(tmp_path)
        assert s.unanswered_high_impact_keys([]) == []

    def test_all_answered(self, tmp_path):
        s = _make_store(tmp_path)
        s.set_annotation("S1", "EF", AnnotationStore.DISPOSITION_EXPECTED, "")
        s.set_annotation("S2", "Activity", AnnotationStore.DISPOSITION_ERROR, "")
        assert s.unanswered_high_impact_keys(["S1::EF", "S2::Activity"]) == []


# ─────────────────────────────────────────────────────────────────────────────
# all_annotations / to_dict / save
# ─────────────────────────────────────────────────────────────────────────────

class TestPersistence:
    def test_all_annotations_returns_copy(self, tmp_path):
        s = _make_store(tmp_path)
        s.set_annotation("S", "C", AnnotationStore.DISPOSITION_EXPECTED, "")
        anns = s.all_annotations()
        assert "S::C" in anns
        # Modifying the returned dict does not affect the store
        anns["extra"] = {}
        assert "extra" not in s.all_annotations()

    def test_to_dict_contains_metadata(self, tmp_path):
        s = _make_store(tmp_path, analyst="Dave", project="MyProject")
        d = s.to_dict()
        assert d["analyst"] == "Dave"
        assert d["project"] == "MyProject"

    def test_to_dict_contains_annotations(self, tmp_path):
        s = _make_store(tmp_path)
        s.set_annotation("S", "C", AnnotationStore.DISPOSITION_EXPECTED, "note")
        d = s.to_dict()
        assert "S::C" in d["annotations"]

    def test_save_creates_parent_dirs(self, tmp_path):
        nested = tmp_path / "a" / "b" / "c"
        sp = nested / "test.json"
        s = AnnotationStore.load_or_create(
            sp, analyst="A", project="P",
            v1_path="/a", v1_label="V1", v2_path="/b", v2_label="V2",
            id_column="ID", materiality_threshold=5.0,
            column_mapping=[], structural_changes={},
        )
        s.save()  # explicitly persist; save() creates parent dirs
        assert sp.exists()

    def test_json_is_valid_on_disk(self, tmp_path):
        s = _make_store(tmp_path)
        s.set_annotation("S", "C", AnnotationStore.DISPOSITION_EXPECTED, "note")
        sp = sidecar_path(tmp_path, "2026-01-01", "Draft 1", "Draft 2")
        data = json.loads(sp.read_text(encoding="utf-8"))
        assert "annotations" in data
        assert "created" in data

    def test_reload_round_trip(self, tmp_path):
        sp = sidecar_path(tmp_path, "2026-01-01", "V1", "V2")
        s1 = AnnotationStore.load_or_create(
            sp, analyst="E", project="P",
            v1_path="/a", v1_label="V1", v2_path="/b", v2_label="V2",
            id_column="ID", materiality_threshold=5.0,
            column_mapping=[], structural_changes={},
        )
        s1.set_annotation("Site-X", "EF", AnnotationStore.DISPOSITION_REVIEW, "check this")
        s2 = AnnotationStore.load_or_create(
            sp, analyst="E", project="P",
            v1_path="/a", v1_label="V1", v2_path="/b", v2_label="V2",
            id_column="ID", materiality_threshold=5.0,
            column_mapping=[], structural_changes={},
        )
        ann = s2.get_annotation("Site-X", "EF")
        assert ann is not None
        assert ann["note"] == "check this"
        assert ann["disposition"] == AnnotationStore.DISPOSITION_REVIEW
