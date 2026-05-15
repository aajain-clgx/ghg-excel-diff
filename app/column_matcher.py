"""
column_matcher.py
-----------------
Matches column headers between two Excel sheets and optionally validates
against a reference template.

All matching is case-insensitive with leading/trailing whitespace stripped.
No fuzzy matching — a renamed column is treated as deleted + added until
the analyst manually pairs it. This preserves trust by never silently
making a wrong match.
"""

from __future__ import annotations
from dataclasses import dataclass, field


# ---------------------------------------------------------------------------
# Known site-identifier column name candidates (auto-detection hints)
# ---------------------------------------------------------------------------
SITE_ID_CANDIDATES = {"site", "site id", "site_id", "id", "facility", "location",
                      "facility id", "facility_id", "site name", "sitename"}


@dataclass
class ColumnMatch:
    """Represents one matched or unmatched column pair."""
    v1_header: str | None       # Original header from V1 (None if V2-only)
    v2_header: str | None       # Original header from V2 (None if V1-only)
    status: str                 # "matched" | "v1_only" | "v2_only" | "manual"
    confirmed: bool = False     # Analyst has explicitly confirmed this pair
    manual: bool = False        # Analyst manually paired these (not auto-matched)

    @property
    def is_structural_change(self) -> bool:
        return self.status in ("v1_only", "v2_only")


@dataclass
class ColumnMatchResult:
    """Full output of a column matching run."""
    matches: list[ColumnMatch] = field(default_factory=list)
    auto_id_column: str | None = None          # Best-guess site ID column (original casing)
    id_candidates: list[str] = field(default_factory=list)  # All plausible ID columns
    template_warnings: list[str] = field(default_factory=list)  # Advisory template issues

    @property
    def matched_pairs(self) -> list[ColumnMatch]:
        return [m for m in self.matches if m.status in ("matched", "manual")]

    @property
    def v1_only(self) -> list[ColumnMatch]:
        return [m for m in self.matches if m.status == "v1_only"]

    @property
    def v2_only(self) -> list[ColumnMatch]:
        return [m for m in self.matches if m.status == "v2_only"]

    @property
    def needs_confirmation(self) -> list[ColumnMatch]:
        """Matched pairs that have not been confirmed by the analyst."""
        return [m for m in self.matched_pairs if not m.confirmed]

    def to_serializable(self) -> list[dict]:
        """Convert to a JSON-serializable list for the sidecar store."""
        return [
            {
                "v1": m.v1_header,
                "v2": m.v2_header,
                "status": m.status,
                "confirmed": m.confirmed,
                "manual": m.manual,
            }
            for m in self.matches
        ]


def _normalize(header: str) -> str:
    """Normalize a header for comparison: lowercase + strip whitespace."""
    return header.strip().lower()


def match_columns(
    headers_v1: list[str],
    headers_v2: list[str],
    template_headers: list[str] | None = None,
) -> ColumnMatchResult:
    """
    Match columns between two sheets.

    Parameters
    ----------
    headers_v1 : list[str]
        Column headers read from the V1 sheet (row 1).
    headers_v2 : list[str]
        Column headers read from the V2 sheet (row 1).
    template_headers : list[str] | None
        If provided, column headers from the reference template. Used only
        for advisory warnings — never blocks the diff.

    Returns
    -------
    ColumnMatchResult
        Matched pairs, unmatched columns, ID candidates, and template warnings.
    """
    result = ColumnMatchResult()

    norm_v2_map: dict[str, str] = {_normalize(h): h for h in headers_v2}
    matched_v2_norm: set[str] = set()

    for h1 in headers_v1:
        norm = _normalize(h1)
        if norm in norm_v2_map:
            h2 = norm_v2_map[norm]
            result.matches.append(
                ColumnMatch(v1_header=h1, v2_header=h2, status="matched", confirmed=True)
            )
            matched_v2_norm.add(norm)
        else:
            result.matches.append(
                ColumnMatch(v1_header=h1, v2_header=None, status="v1_only")
            )

    # V2 columns that had no V1 counterpart
    for h2 in headers_v2:
        if _normalize(h2) not in matched_v2_norm:
            result.matches.append(
                ColumnMatch(v1_header=None, v2_header=h2, status="v2_only")
            )

    # Auto-detect site ID column candidates from V1 headers
    for h in headers_v1:
        if _normalize(h) in SITE_ID_CANDIDATES:
            result.id_candidates.append(h)

    if len(result.id_candidates) == 1:
        result.auto_id_column = result.id_candidates[0]

    # Optional template validation (advisory only)
    if template_headers is not None:
        norm_v1_set = {_normalize(h) for h in headers_v1}
        norm_v2_set = {_normalize(h) for h in headers_v2}
        for th in template_headers:
            norm_th = _normalize(th)
            if norm_th not in norm_v1_set:
                result.template_warnings.append(
                    f"'{th}' is in the template but missing from V1"
                )
            if norm_th not in norm_v2_set:
                result.template_warnings.append(
                    f"'{th}' is in the template but missing from V2"
                )
        # Columns in either file not in template
        norm_template_set = {_normalize(th) for th in template_headers}
        for h in headers_v1:
            if _normalize(h) not in norm_template_set:
                result.template_warnings.append(
                    f"'{h}' (V1) is not in the template"
                )

    return result


def apply_manual_pair(
    result: ColumnMatchResult,
    v1_header: str,
    v2_header: str,
) -> ColumnMatchResult:
    """
    Analyst manually pairs a V1-only column with a V2-only column.
    Removes both from their unmatched lists and adds a confirmed manual pair.

    Parameters
    ----------
    result : ColumnMatchResult
        Existing match result to update (mutated in place, also returned).
    v1_header : str
        Original (un-normalized) V1 column header.
    v2_header : str
        Original (un-normalized) V2 column header.
    """
    # Remove the individual unmatched entries
    result.matches = [
        m for m in result.matches
        if not (m.status == "v1_only" and m.v1_header == v1_header)
        and not (m.status == "v2_only" and m.v2_header == v2_header)
    ]
    result.matches.append(
        ColumnMatch(
            v1_header=v1_header,
            v2_header=v2_header,
            status="manual",
            confirmed=True,
            manual=True,
        )
    )
    return result
