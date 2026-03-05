"""
Comparison engine.

Stages (in strict order):
  1. Tag validation  — duplicates in either sheet, tags present in only one sheet.
  2. Header matching — exact-string match of column names (excluding tag columns).
  3. Data comparison — cell-level diff for every (tag, header) pair where
                       both the tag and the header exist in both sheets.
"""

from __future__ import annotations

from dataclasses import dataclass, field

import pandas as pd

from .normalizer import normalize_value, normalize_series, values_match


# ---------------------------------------------------------------------------
# Data containers
# ---------------------------------------------------------------------------

@dataclass
class TagIssue:
    tag: str
    sheet: str          # "Client" | "INFORM"
    reason: str


@dataclass
class TagValidationResult:
    issues: list[TagIssue]
    comparable_tags: set          # tags safe to compare (in both, no duplicates)
    client_only_tags: set         # in client but not INFORM (non-duplicates)
    inform_only_tags: set         # in INFORM but not client (non-duplicates)

    @property
    def has_issues(self) -> bool:
        return bool(self.issues)

    def issues_df(self) -> pd.DataFrame:
        if not self.issues:
            return pd.DataFrame(columns=["Tag", "Sheet", "Reason"])
        return pd.DataFrame(
            [{"Tag": i.tag, "Sheet": i.sheet, "Reason": i.reason} for i in self.issues]
        )


@dataclass
class HeaderMatchResult:
    matched: list[str]
    client_only: list[str]
    inform_only: list[str]

    @property
    def all_matched(self) -> bool:
        return not self.client_only and not self.inform_only


@dataclass
class CellDiff:
    tag: str
    header: str
    client_raw: str
    inform_raw: str
    client_norm: str
    inform_norm: str
    match: bool


@dataclass
class ComparisonResult:
    tag_validation: TagValidationResult
    header_match: HeaderMatchResult
    diffs: list[CellDiff] = field(default_factory=list)

    def diffs_df(self) -> pd.DataFrame:
        if not self.diffs:
            return pd.DataFrame(
                columns=[
                    "Tag", "Header",
                    "Client Value", "INFORM Value",
                    "Client (Normalised)", "INFORM (Normalised)",
                    "Match",
                ]
            )
        return pd.DataFrame(
            [
                {
                    "Tag": d.tag,
                    "Header": d.header,
                    "Client Value": d.client_raw,
                    "INFORM Value": d.inform_raw,
                    "Client (Normalised)": d.client_norm,
                    "INFORM (Normalised)": d.inform_norm,
                    "Match": d.match,
                }
                for d in self.diffs
            ]
        )

    @property
    def mismatch_count(self) -> int:
        return sum(1 for d in self.diffs if not d.match)

    @property
    def match_count(self) -> int:
        return sum(1 for d in self.diffs if d.match)


# ---------------------------------------------------------------------------
# Stage 1: Tag validation
# ---------------------------------------------------------------------------

def validate_tags(
    client_df: pd.DataFrame,
    inform_df: pd.DataFrame,
    client_tag_col: str,
    inform_tag_col: str,
) -> TagValidationResult:
    """
    Check for:
      - Empty / null tags in either sheet
      - Duplicate tags within each sheet
      - Tags present in only one sheet

    Returns a TagValidationResult.  The comparable_tags set contains only tags
    that are unique in both sheets AND present in both sheets.
    """
    issues: list[TagIssue] = []

    client_norm = client_df[client_tag_col].apply(normalize_value)
    inform_norm = inform_df[inform_tag_col].apply(normalize_value)

    # ---- Null tags ----
    cn_null = int(client_norm.isna().sum())
    in_null = int(inform_norm.isna().sum())
    if cn_null:
        issues.append(TagIssue(
            tag="(empty)",
            sheet="Client",
            reason=f"{cn_null} row(s) have empty / null tag values and will be skipped",
        ))
    if in_null:
        issues.append(TagIssue(
            tag="(empty)",
            sheet="INFORM",
            reason=f"{in_null} row(s) have empty / null tag values and will be skipped",
        ))

    client_clean = client_norm.dropna()
    inform_clean = inform_norm.dropna()

    # ---- Duplicates ----
    client_dup_mask = client_clean.duplicated(keep=False)
    for tag in sorted(client_clean[client_dup_mask].unique(), key=str):
        count = int((client_clean == tag).sum())
        issues.append(TagIssue(
            tag=str(tag),
            sheet="Client",
            reason=f"Duplicate — appears {count} times in client sheet; excluded from comparison",
        ))

    inform_dup_mask = inform_clean.duplicated(keep=False)
    for tag in sorted(inform_clean[inform_dup_mask].unique(), key=str):
        count = int((inform_clean == tag).sum())
        issues.append(TagIssue(
            tag=str(tag),
            sheet="INFORM",
            reason=f"Duplicate — appears {count} times in INFORM sheet; excluded from comparison",
        ))

    # Non-duplicate tag sets
    client_unique = set(client_clean[~client_dup_mask])
    inform_unique = set(inform_clean[~inform_dup_mask])

    comparable_tags = client_unique & inform_unique
    client_only = client_unique - inform_unique
    inform_only = inform_unique - client_unique

    # ---- One-sided tags ----
    for tag in sorted(client_only, key=str):
        issues.append(TagIssue(
            tag=str(tag),
            sheet="Client",
            reason="Tag exists in client sheet only — no INFORM row to compare against",
        ))
    for tag in sorted(inform_only, key=str):
        issues.append(TagIssue(
            tag=str(tag),
            sheet="INFORM",
            reason="Tag exists in INFORM sheet only — no client row to compare against",
        ))

    return TagValidationResult(
        issues=issues,
        comparable_tags=comparable_tags,
        client_only_tags=client_only,
        inform_only_tags=inform_only,
    )


# ---------------------------------------------------------------------------
# Stage 2: Header matching
# ---------------------------------------------------------------------------

def match_headers(
    client_headers: list[str],
    inform_headers: list[str],
    client_tag_col: str,
    inform_tag_col: str,
) -> HeaderMatchResult:
    """
    Exact-string comparison of column headers, excluding the tag columns.
    Returns matched, client-only, and INFORM-only header lists.
    """
    client_set = set(h for h in client_headers if h != client_tag_col)
    inform_set = set(h for h in inform_headers if h != inform_tag_col)

    return HeaderMatchResult(
        matched=sorted(client_set & inform_set),
        client_only=sorted(client_set - inform_set),
        inform_only=sorted(inform_set - client_set),
    )


# ---------------------------------------------------------------------------
# Stage 3: Cell-level data comparison
# ---------------------------------------------------------------------------

def _fmt(value) -> str:
    """Human-readable representation of a normalised value."""
    if value is None:
        return ""
    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d")
    return str(value)


def compare_data(
    client_df: pd.DataFrame,
    inform_df: pd.DataFrame,
    client_tag_col: str,
    inform_tag_col: str,
    matched_headers: list[str],
    comparable_tags: set,
) -> list[CellDiff]:
    """
    For every (tag, header) pair where the tag is in comparable_tags and the
    header is in matched_headers, compare the client and INFORM values after
    normalisation.

    Returns a list of CellDiff records (one per cell).
    """
    if not comparable_tags or not matched_headers:
        return []

    # Normalise tag columns for fast lookup
    client_df = client_df.copy()
    inform_df = inform_df.copy()
    client_df["_tag_norm"] = client_df[client_tag_col].apply(normalize_value)
    inform_df["_tag_norm"] = inform_df[inform_tag_col].apply(normalize_value)

    # Pre-normalise comparison columns (vectorised for speed)
    for col in matched_headers:
        client_df[f"_n_{col}"] = normalize_series(client_df[col])
        inform_df[f"_n_{col}"] = normalize_series(inform_df[col])

    # Index by normalised tag for O(1) lookup
    client_indexed = client_df[client_df["_tag_norm"].isin(comparable_tags)].set_index("_tag_norm")
    inform_indexed = inform_df[inform_df["_tag_norm"].isin(comparable_tags)].set_index("_tag_norm")

    diffs: list[CellDiff] = []

    for tag in sorted(comparable_tags, key=str):
        try:
            c_row = client_indexed.loc[tag]
            i_row = inform_indexed.loc[tag]
        except KeyError:
            continue

        # If somehow duplicates survived (shouldn't happen), take first row
        if isinstance(c_row, pd.DataFrame):
            c_row = c_row.iloc[0]
        if isinstance(i_row, pd.DataFrame):
            i_row = i_row.iloc[0]

        for col in matched_headers:
            c_raw = c_row[col]
            i_raw = i_row[col]
            c_norm = c_row[f"_n_{col}"]
            i_norm = i_row[f"_n_{col}"]

            diffs.append(CellDiff(
                tag=str(tag),
                header=col,
                client_raw=str(c_raw) if c_raw is not None else "",
                inform_raw=str(i_raw) if i_raw is not None else "",
                client_norm=_fmt(c_norm),
                inform_norm=_fmt(i_norm),
                match=values_match(c_norm, i_norm),
            ))

    return diffs


# ---------------------------------------------------------------------------
# Convenience: run all three stages
# ---------------------------------------------------------------------------

def run_full_comparison(
    client_df: pd.DataFrame,
    inform_df: pd.DataFrame,
    client_tag_col: str,
    inform_tag_col: str,
    client_headers: list[str],
    inform_headers: list[str],
) -> ComparisonResult:
    """Execute all three comparison stages and return a ComparisonResult."""
    tag_result = validate_tags(client_df, inform_df, client_tag_col, inform_tag_col)
    header_result = match_headers(client_headers, inform_headers, client_tag_col, inform_tag_col)

    diffs = compare_data(
        client_df=client_df,
        inform_df=inform_df,
        client_tag_col=client_tag_col,
        inform_tag_col=inform_tag_col,
        matched_headers=header_result.matched,
        comparable_tags=tag_result.comparable_tags,
    )

    return ComparisonResult(
        tag_validation=tag_result,
        header_match=header_result,
        diffs=diffs,
    )
