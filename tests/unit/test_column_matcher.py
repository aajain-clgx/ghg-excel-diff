"""
tests/unit/test_column_matcher.py
----------------------------------
Unit tests for column_matcher.py — complete coverage of:
  _normalize, SITE_ID_CANDIDATES, ColumnMatch, ColumnMatchResult (all properties),
  match_columns, apply_manual_pair
"""

import pytest
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent.parent / "app"))

from column_matcher import (
    _normalize,
    match_columns,
    apply_manual_pair,
    ColumnMatch,
    ColumnMatchResult,
    SITE_ID_CANDIDATES,
)


# ─────────────────────────────────────────────────────────────────────────────
# _normalize
# ─────────────────────────────────────────────────────────────────────────────

class TestNormalize:
    def test_lowercases(self):
        assert _normalize("Hello") == "hello"

    def test_strips_leading_whitespace(self):
        assert _normalize("  foo") == "foo"

    def test_strips_trailing_whitespace(self):
        assert _normalize("foo  ") == "foo"

    def test_strips_both_sides(self):
        assert _normalize("  Foo Bar  ") == "foo bar"

    def test_empty_string(self):
        assert _normalize("") == ""

    def test_already_normalized(self):
        assert _normalize("site id") == "site id"

    def test_mixed_case_internal(self):
        assert _normalize("EF Natural Gas") == "ef natural gas"


# ─────────────────────────────────────────────────────────────────────────────
# ColumnMatch dataclass
# ─────────────────────────────────────────────────────────────────────────────

class TestColumnMatch:
    def test_is_structural_v1_only(self):
        m = ColumnMatch(v1_header="A", v2_header=None, status="v1_only")
        assert m.is_structural_change is True

    def test_is_structural_v2_only(self):
        m = ColumnMatch(v1_header=None, v2_header="B", status="v2_only")
        assert m.is_structural_change is True

    def test_not_structural_matched(self):
        m = ColumnMatch(v1_header="A", v2_header="A", status="matched")
        assert m.is_structural_change is False

    def test_not_structural_manual(self):
        m = ColumnMatch(v1_header="A", v2_header="B", status="manual")
        assert m.is_structural_change is False

    def test_defaults(self):
        m = ColumnMatch(v1_header="X", v2_header="Y", status="matched")
        assert m.confirmed is False
        assert m.manual is False


# ─────────────────────────────────────────────────────────────────────────────
# ColumnMatchResult properties
# ─────────────────────────────────────────────────────────────────────────────

class TestColumnMatchResult:
    def _make(self):
        r = ColumnMatchResult()
        r.matches = [
            ColumnMatch("A", "A", "matched", confirmed=True),
            ColumnMatch("B", "B2", "manual", confirmed=True, manual=True),
            ColumnMatch("C", None, "v1_only"),
            ColumnMatch(None, "D", "v2_only"),
            ColumnMatch("E", "E", "matched", confirmed=False),
        ]
        return r

    def test_matched_pairs_includes_matched_and_manual(self):
        r = self._make()
        ids = {(m.v1_header, m.v2_header) for m in r.matched_pairs}
        assert ("A", "A") in ids
        assert ("B", "B2") in ids
        assert len(r.matched_pairs) == 3  # A, B, E

    def test_v1_only(self):
        r = self._make()
        assert len(r.v1_only) == 1
        assert r.v1_only[0].v1_header == "C"

    def test_v2_only(self):
        r = self._make()
        assert len(r.v2_only) == 1
        assert r.v2_only[0].v2_header == "D"

    def test_needs_confirmation(self):
        r = self._make()
        # E is matched but confirmed=False
        unconfirmed = r.needs_confirmation
        assert len(unconfirmed) == 1
        assert unconfirmed[0].v1_header == "E"

    def test_to_serializable(self):
        r = self._make()
        s = r.to_serializable()
        assert isinstance(s, list)
        assert all(isinstance(d, dict) for d in s)
        keys = {"v1", "v2", "status", "confirmed", "manual"}
        for d in s:
            assert keys == set(d.keys())

    def test_to_serializable_values(self):
        r = ColumnMatchResult()
        r.matches = [ColumnMatch("X", "Y", "matched", confirmed=True, manual=False)]
        s = r.to_serializable()
        assert s[0] == {"v1": "X", "v2": "Y", "status": "matched",
                        "confirmed": True, "manual": False}

    def test_empty_result(self):
        r = ColumnMatchResult()
        assert r.matched_pairs == []
        assert r.v1_only == []
        assert r.v2_only == []
        assert r.needs_confirmation == []
        assert r.to_serializable() == []


# ─────────────────────────────────────────────────────────────────────────────
# match_columns
# ─────────────────────────────────────────────────────────────────────────────

class TestMatchColumns:
    def test_identical_headers_all_match(self):
        h = ["Site ID", "EF", "Activity"]
        r = match_columns(h, h)
        assert len(r.matched_pairs) == 3
        assert len(r.v1_only) == 0
        assert len(r.v2_only) == 0

    def test_case_insensitive_match(self):
        r = match_columns(["EF Natural Gas"], ["ef natural gas"])
        assert len(r.matched_pairs) == 1

    def test_whitespace_match(self):
        r = match_columns(["  EF  "], ["EF"])
        assert len(r.matched_pairs) == 1

    def test_v1_only_column(self):
        r = match_columns(["A", "B"], ["A"])
        assert len(r.v1_only) == 1
        assert r.v1_only[0].v1_header == "B"

    def test_v2_only_column(self):
        r = match_columns(["A"], ["A", "C"])
        assert len(r.v2_only) == 1
        assert r.v2_only[0].v2_header == "C"

    def test_no_overlap(self):
        r = match_columns(["A", "B"], ["X", "Y"])
        assert len(r.matched_pairs) == 0
        assert len(r.v1_only) == 2
        assert len(r.v2_only) == 2

    def test_empty_headers(self):
        r = match_columns([], [])
        assert r.matched_pairs == []
        assert r.auto_id_column is None

    def test_site_id_auto_detect_exact(self):
        r = match_columns(["Site ID", "EF"], ["Site ID", "EF"])
        assert r.auto_id_column == "Site ID"

    def test_site_id_auto_detect_case_insensitive(self):
        r = match_columns(["SITE ID", "EF"], ["SITE ID", "EF"])
        assert r.auto_id_column == "SITE ID"

    def test_site_id_auto_detect_facility(self):
        r = match_columns(["Facility", "EF"], ["Facility", "EF"])
        assert r.auto_id_column == "Facility"

    def test_site_id_multiple_candidates_no_auto_select(self):
        # Two candidates → ambiguous → auto_id_column stays None
        r = match_columns(["Site ID", "Facility ID", "EF"], ["Site ID", "Facility ID", "EF"])
        assert r.auto_id_column is None
        assert len(r.id_candidates) == 2

    def test_no_site_id_candidate(self):
        r = match_columns(["EF", "Activity"], ["EF", "Activity"])
        assert r.auto_id_column is None
        assert r.id_candidates == []

    def test_matched_columns_confirmed(self):
        r = match_columns(["A"], ["A"])
        assert r.matched_pairs[0].confirmed is True

    def test_v1_only_not_confirmed(self):
        r = match_columns(["A"], [])
        assert r.v1_only[0].confirmed is False

    # ── Template validation ──────────────────────────────────────────────

    def test_template_no_warnings_when_all_present(self):
        cols = ["Site ID", "EF"]
        r = match_columns(cols, cols, template_headers=cols)
        # Template columns in V1 extra columns may still warn but v1 cols ARE in template
        assert not any("missing from V1" in w for w in r.template_warnings)
        assert not any("missing from V2" in w for w in r.template_warnings)

    def test_template_warns_missing_from_v1(self):
        r = match_columns(["EF"], ["EF"], template_headers=["EF", "Activity"])
        assert any("missing from V1" in w for w in r.template_warnings)

    def test_template_warns_missing_from_v2(self):
        r = match_columns(["EF", "Activity"], ["EF"], template_headers=["EF", "Activity"])
        assert any("missing from V2" in w for w in r.template_warnings)

    def test_template_warns_extra_in_v1(self):
        r = match_columns(["EF", "Extra"], ["EF"], template_headers=["EF"])
        assert any("not in the template" in w for w in r.template_warnings)

    def test_template_none_no_warnings(self):
        r = match_columns(["EF"], ["EF"], template_headers=None)
        assert r.template_warnings == []

    def test_preserves_original_casing_in_result(self):
        r = match_columns(["EF Natural Gas"], ["ef natural gas"])
        pair = r.matched_pairs[0]
        assert pair.v1_header == "EF Natural Gas"
        assert pair.v2_header == "ef natural gas"


# ─────────────────────────────────────────────────────────────────────────────
# apply_manual_pair
# ─────────────────────────────────────────────────────────────────────────────

class TestApplyManualPair:
    def _result_with_unmatched(self):
        r = match_columns(["A", "OldName"], ["A", "NewName"])
        return r  # OldName is v1_only, NewName is v2_only

    def test_removes_v1_only(self):
        r = self._result_with_unmatched()
        apply_manual_pair(r, "OldName", "NewName")
        assert not any(m.v1_header == "OldName" and m.status == "v1_only"
                       for m in r.matches)

    def test_removes_v2_only(self):
        r = self._result_with_unmatched()
        apply_manual_pair(r, "OldName", "NewName")
        assert not any(m.v2_header == "NewName" and m.status == "v2_only"
                       for m in r.matches)

    def test_adds_manual_pair(self):
        r = self._result_with_unmatched()
        apply_manual_pair(r, "OldName", "NewName")
        manual = [m for m in r.matches if m.status == "manual"]
        assert len(manual) == 1
        assert manual[0].v1_header == "OldName"
        assert manual[0].v2_header == "NewName"
        assert manual[0].manual is True
        assert manual[0].confirmed is True

    def test_returns_same_result_object(self):
        r = self._result_with_unmatched()
        ret = apply_manual_pair(r, "OldName", "NewName")
        assert ret is r

    def test_auto_match_unaffected(self):
        r = self._result_with_unmatched()
        before_a = next(m for m in r.matches if m.v1_header == "A")
        apply_manual_pair(r, "OldName", "NewName")
        after_a = next(m for m in r.matches if m.v1_header == "A")
        assert before_a.status == after_a.status

    def test_manual_pair_in_matched_pairs(self):
        r = self._result_with_unmatched()
        apply_manual_pair(r, "OldName", "NewName")
        headers = [(m.v1_header, m.v2_header) for m in r.matched_pairs]
        assert ("OldName", "NewName") in headers

    def test_successive_manual_pairs(self):
        r = match_columns(["A", "B", "X"], ["A", "B2", "X2"])
        apply_manual_pair(r, "B", "B2")
        apply_manual_pair(r, "X", "X2")
        assert len(r.v1_only) == 0
        assert len(r.v2_only) == 0
        assert len([m for m in r.matches if m.status == "manual"]) == 2
