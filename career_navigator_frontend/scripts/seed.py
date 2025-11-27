#!/usr/bin/env python3
"""
Python Seeding Utility for Supabase (Career Navigator MVP)

Features:
- Reads provided Excel attachments to seed:
  roles, competencies, role_competencies, role_descriptions, role_adjacency, learning_resources
- Idempotent upserts via Supabase REST (PostgREST) using natural keys:
  - roles: on_conflict=code
  - competencies: on_conflict=code
  - role_competencies: on_conflict=role_id,competency_id
  - role_adjacency: on_conflict=source_role_id,target_role_id
  - role_descriptions: on_conflict=role_id,source
  - learning_resources: on_conflict=competency_id,url
- Robust header normalization and heuristics for mixed Excel layouts
- Retry logic with exponential backoff for transient failures
- Logging with progress summaries
- Selective seeding via CLI flags

Environment:
- SUPABASE_URL
- SUPABASE_SERVICE_ROLE_KEY

Optional fallbacks (if above are not set):
- REACT_APP_SUPABASE_URL
- REACT_APP_SUPABASE_SERVICE_ROLE_KEY or REACT_APP_SUPABASE_KEY (requires RLS disabled; service role is preferred)

Usage:
  python3 seed.py --help

Default source files use direct attachment paths under /home/kavia/workspace/code-generation/attachments/.
"""
import argparse
import os
import sys
import json
import time
import math
import glob
import logging
from typing import Dict, List, Any, Optional, Tuple, Iterable

import pandas as pd
import requests


# ---------- Logging configuration ----------
LOG = logging.getLogger("seed")
handler = logging.StreamHandler()
fmt = logging.Formatter("[%(levelname)s] %(message)s")
handler.setFormatter(fmt)
LOG.addHandler(handler)
LOG.setLevel(logging.INFO)


# ---------- Defaults: use direct attachment paths ----------
DEFAULTS = {
    "competency_xlsx": "/home/kavia/workspace/code-generation/attachments/20251127_090403_Competency_mapping.xlsx",
    "adjacency_xlsx": "/home/kavia/workspace/code-generation/attachments/20251127_090402_CA_Role_Adjacency29.xlsx",
    "adjacency_xlsx_alt": "/home/kavia/workspace/code-generation/attachments/20251127_090401_CA_Role_Adjacency.xlsx",
    "roles_xlsx": "/home/kavia/workspace/code-generation/attachments/20251127_090406_Role_Navigator_Worksheet.xlsx",
    "resources_xlsx": None,  # if not given, we will try roles_xlsx to find a resources sheet
    "cards_glob": "/home/kavia/workspace/code-generation/attachments/2025*Role_Card_*",
}

# ---------- Utility functions ----------
# PUBLIC_INTERFACE
def slugify(value: Any) -> str:
    """Generate a safe slug from an input string (lowercase, hyphenated alphanumerics)."""
    s = "" if value is None else str(value)
    s = s.strip().lower()
    # Remove parentheses
    s = s.replace("(", "").replace(")", "")
    # Replace non-alphanumerics with hyphens
    out = []
    prev_dash = False
    for ch in s:
        if ch.isalnum():
            out.append(ch)
            prev_dash = False
        else:
            if not prev_dash:
                out.append("-")
                prev_dash = True
    # Trim leading/trailing hyphens
    sl = "".join(out).strip("-")
    return sl


def _norm(s: Any) -> str:
    return "" if s is None else str(s).strip().lower()


# PUBLIC_INTERFACE
def map_level(raw: Any) -> Optional[int]:
    """Map various competency level representations to 1..4 or None."""
    if raw is None:
        return None
    s = str(raw).strip().lower()
    if s in {"", "0", "-", "n/a", "na"}:
        return None

    mapping = {
        "b": 1, "basic": 1, "beginner": 1, "foundation": 1, "foundational": 1, "novice": 1, "baseline": 1,
        "p": 2, "proficient": 2, "intermediate": 2,
        "a": 3, "advanced": 3,
        "m": 4, "master": 4, "authority": 4, "expert": 4, "e": 4,
    }
    if s in mapping:
        return mapping[s]
    try:
        n = float(s.replace("%", ""))
        if n <= 1:
            return 1
        if n <= 2:
            return 2
        if n <= 3:
            return 3
        return 4
    except Exception:
        # Fallback: first letter mapping
        if s and s[0] in mapping:
            return mapping[s[0]]
    return None


def safe_number(raw: Any) -> Optional[float]:
    if raw is None:
        return None
    if isinstance(raw, (int, float)):
        return float(raw)
    try:
        s = str(raw).strip().replace("%", "")
        if s == "":
            return None
        return float(s)
    except Exception:
        return None


def unique_by(seq: Iterable[dict], key_fn) -> List[dict]:
    seen = set()
    out = []
    for item in seq:
        k = key_fn(item)
        if k not in seen:
            seen.add(k)
            out.append(item)
    return out


# ---------- Supabase REST client ----------
class SupabaseRestClient:
    """Lightweight client for Supabase PostgREST endpoints."""

    def __init__(self, url: str, key: str, timeout: int = 60):
        base = url.rstrip("/")
        self.rest_url = f"{base}/rest/v1"
        self.headers = {
            "apikey": key,
            "Authorization": f"Bearer {key}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        }
        self.timeout = timeout

    def _request(self, method: str, path: str, *, params: dict = None, data: Any = None,
                 prefer: Optional[str] = None) -> Tuple[int, dict, Any]:
        headers = dict(self.headers)
        if prefer:
            headers["Prefer"] = prefer
        url = f"{self.rest_url}/{path.lstrip('/')}"
        resp = requests.request(method, url, headers=headers, params=params, data=data, timeout=self.timeout)
        try:
            js = resp.json()
        except Exception:
            js = None
        return resp.status_code, dict(resp.headers), js

    def get_select(self, table: str, select: str, limit: int = 1000, extra_params: dict = None) -> List[dict]:
        params = {"select": select, "limit": limit}
        if extra_params:
            params.update(extra_params)
        code, _, js = self._request("GET", table, params=params)
        if code >= 400:
            raise RuntimeError(f"GET {table} failed: HTTP {code} {js}")
        return js or []

    def upsert(self, table: str, rows: List[dict], on_conflict: str, return_minimal: bool = True) -> None:
        # Use Prefer with upsert and return=minimal (faster)
        prefer = "resolution=merge-duplicates"
        if return_minimal:
            prefer += ", return=minimal"
        params = {"on_conflict": on_conflict}
        payload = json.dumps(rows)
        code, _, js = self._request("POST", table, params=params, data=payload, prefer=prefer)
        if code >= 400:
            raise RuntimeError(f"Upsert {table} failed: HTTP {code} {js}")


# ---------- Retry wrapper ----------
async_mode = False  # we are not using asyncio, but keep retry logic generic


# PUBLIC_INTERFACE
def with_retries(func, label: str, retries: int = 4, base_delay_ms: int = 500):
    """Retry a synchronous function for transient failures with exponential backoff."""
    last_err = None
    for attempt in range(retries + 1):
        try:
            return func()
        except Exception as e:
            last_err = e
            if attempt == retries:
                break
            delay = base_delay_ms * (2 ** attempt) + int(50 * math.sin(time.time()))
            LOG.warning(f"Retry {attempt + 1}/{retries} on {label}: {e}. Waiting {delay}ms...")
            time.sleep(delay / 1000.0)
    raise last_err


# ---------- Parsing logic ----------
META_COMP_KEYS = {"competency", "competency name", "competency_name", "competency code", "competency_code", "category", "description", "notes", "summary"}


def _is_role_column(header: str) -> bool:
    h = _norm(header)
    if h in META_COMP_KEYS:
        return False
    return True


def _read_excel(file_path: str) -> Dict[str, pd.DataFrame]:
    try:
        xls = pd.ExcelFile(file_path)
        sheets = {}
        for name in xls.sheet_names:
            try:
                df = pd.read_excel(xls, sheet_name=name, dtype=object)
                sheets[name] = df
            except Exception as e:
                LOG.warning(f"Skipping sheet '{name}' in {file_path}: {e}")
        return sheets
    except Exception as e:
        LOG.error(f"Failed to read workbook: {file_path} -> {e}")
        return {}


# PUBLIC_INTERFACE
def parse_competency_mapping(file_path: str) -> Tuple[Dict[str, dict], Dict[str, dict], List[dict]]:
    """
    Parse competency mapping matrix:
    Returns:
      competencies: code -> {code, name, category}
      roles: code -> {code, name, description}
      role_competencies: [{role_code, competency_code, target_level}]
    """
    competencies: Dict[str, dict] = {}
    roles: Dict[str, dict] = {}
    role_competencies: List[dict] = []

    sheets = _read_excel(file_path)
    if not sheets:
        LOG.warning(f"Competency mapping workbook empty or unreadable: {file_path}")
        return competencies, roles, role_competencies

    # Pick a sheet with "competency" or "mapping" first
    sheet_names = sorted(sheets.keys(), key=lambda n: (0 if "competency" in _norm(n) or "mapping" in _norm(n) else 1, n))
    df = sheets[sheet_names[0]].fillna("")

    headers = list(df.columns)
    role_headers = [h for h in headers if _is_role_column(h)]
    # Validate presence of competency descriptor column
    comp_name_cols = [c for c in headers if _norm(c) in {"competency", "competency name", "competency_name"}]
    comp_code_cols = [c for c in headers if _norm(c) in {"competency code", "competency_code"}]
    category_cols = [c for c in headers if _norm(c) == "category"]

    if not comp_name_cols and not comp_code_cols:
        LOG.warning(f"Competency mapping sheet lacks clear competency columns; headers: {headers}")

    for _, row in df.iterrows():
        comp_name = None
        comp_code_raw = None
        category = None

        for c in comp_name_cols:
            if comp_name is None:
                val = str(row.get(c, "")).strip()
                if val:
                    comp_name = val
        for c in comp_code_cols:
            if comp_code_raw is None:
                val = str(row.get(c, "")).strip()
                if val:
                    comp_code_raw = val
        for c in category_cols:
            if category is None:
                val = str(row.get(c, "")).strip()
                if val:
                    category = val

        if not comp_name and not comp_code_raw:
            continue

        comp_code = slugify(comp_code_raw or comp_name)
        if comp_code not in competencies:
            competencies[comp_code] = {
                "code": comp_code,
                "name": comp_name or comp_code_raw or comp_code,
                "category": category or None,
            }

        # Role columns (matrix): map level in each role col
        for hdr in role_headers:
            val = row.get(hdr, "")
            if val in ("", None):
                continue
            level = map_level(val)
            if level is None:
                continue
            role_code = slugify(hdr)
            if role_code not in roles:
                roles[role_code] = {"code": role_code, "name": str(hdr).strip(), "description": None}
            role_competencies.append({
                "role_code": role_code,
                "competency_code": comp_code,
                "target_level": int(level),
            })

    return competencies, roles, role_competencies


# PUBLIC_INTERFACE
def parse_additional_roles_from_worksheet(file_path: str) -> List[dict]:
    """Scan all sheets for role-like columns and return additional roles."""
    roles: List[dict] = []
    sheets = _read_excel(file_path)
    for sname, df in sheets.items():
        if "resource" in _norm(sname):
            continue
        df = df.fillna("")
        headers = list(df.columns)
        role_like = [h for h in headers if _norm(h) in {
            "role", "role name", "title", "current role", "target role", "name"
        }]
        if not role_like:
            continue
        for _, row in df.iterrows():
            found = None
            for h in role_like:
                v = str(row.get(h, "")).strip()
                if v:
                    found = v
                    break
            if found:
                code = slugify(found)
                roles.append({"code": code, "name": found, "description": None})
    # Dedup by code
    dedup = {}
    for r in roles:
        dedup[r["code"]] = r
    return list(dedup.values())


# PUBLIC_INTERFACE
def parse_learning_resources(resources_path: Optional[str], fallback_roles_path: Optional[str],
                             competencies: Dict[str, dict]) -> List[dict]:
    """
    Parse learning resources from a file that either has a sheet named like 'resource'
    or columns containing URL/Link. Maps optional 'Competency Code' or 'Competency' to competency_code.
    Returns [{competency_code, title, url, provider}]
    """
    paths = [p for p in [resources_path, fallback_roles_path] if p]
    for p in paths:
        sheets = _read_excel(p)
        if not sheets:
            continue
        # prioritize sheets with 'resource' in name
        sheet_names = sorted(sheets.keys(), key=lambda n: (0 if "resource" in _norm(n) else 1, n))
        for sname in sheet_names:
            df = sheets[sname].fillna("")
            if df.empty:
                continue
            # detect URL-like columns
            url_like = [c for c in df.columns if _norm(c) in {"url", "link"}]
            if not url_like and "resource" not in _norm(sname):
                continue
            out: List[dict] = []
            for _, row in df.iterrows():
                comp_code_or_name = None
                for cand in ["Competency Code", "competency code", "Competency", "competency"]:
                    if cand in df.columns:
                        v = str(row.get(cand, "")).strip()
                        if v:
                            comp_code_or_name = v
                            break
                title = None
                for cand in ["Title", "title", "Name", "name"]:
                    if cand in df.columns:
                        v = str(row.get(cand, "")).strip()
                        if v:
                            title = v
                            break
                url = None
                for cand in ["URL", "Url", "url", "Link", "link"]:
                    if cand in df.columns:
                        v = str(row.get(cand, "")).strip()
                        if v:
                            url = v
                            break
                provider = None
                for cand in ["Provider", "provider", "Source", "source"]:
                    if cand in df.columns:
                        v = str(row.get(cand, "")).strip()
                        if v:
                            provider = v
                            break
                if not title or not url:
                    continue
                comp_code = slugify(comp_code_or_name) if comp_code_or_name else None
                if comp_code and comp_code not in competencies:
                    competencies[comp_code] = {"code": comp_code, "name": comp_code_or_name, "category": None}
                out.append({"competency_code": comp_code, "title": title, "url": url, "provider": provider})
            if out:
                return out
    return []


# PUBLIC_INTERFACE
def parse_role_adjacency(file_paths: List[str]) -> List[dict]:
    """
    Parse adjacency rows supporting both row-wise and matrix styles.
    Returns [{source_role_code, target_role_code, score}]
    """
    for p in file_paths:
        if not p or not os.path.exists(p):
            continue
        sheets = _read_excel(p)
        if not sheets:
            continue

        # Heuristic: try row-wise first across sheets, else try matrix on the first sheet
        # Row-wise: look for source/target/score patterns
        aggregated: List[dict] = []
        rowwise_found = False
        for sname, df in sheets.items():
            df = df.fillna("")
            headers = list(df.columns)
            lower = _norm(sname)
            # candidates
            src_cols = [h for h in headers if _norm(h) in {
                "source", "from", "role a", "role 1", "source role", "sourcerole"
            }]
            dst_cols = [h for h in headers if _norm(h) in {
                "target", "to", "role b", "role 2", "target role", "targetrole"
            }]
            score_cols = [h for h in headers if _norm(h) in {
                "score", "weight", "percent", "percentage", "adjacency"
            }]
            if src_cols and dst_cols:
                rowwise_found = True
                score_col = score_cols[0] if score_cols else None
                for _, row in df.iterrows():
                    src = None
                    dst = None
                    for sc in src_cols:
                        if not src:
                            v = str(row.get(sc, "")).strip()
                            if v:
                                src = v
                    for dc in dst_cols:
                        if not dst:
                            v = str(row.get(dc, "")).strip()
                            if v:
                                dst = v
                    if not src or not dst:
                        continue
                    numeric = safe_number(row.get(score_col)) if score_col else None
                    aggregated.append({
                        "source_role_code": slugify(src),
                        "target_role_code": slugify(dst),
                        "score": numeric
                    })
                if aggregated:
                    return aggregated
            # If sheet name suggests adjacency, we'll try matrix fallback after
            if "adjacency" in lower:
                break

        # Matrix fallback: first row = targets, first col = sources
        s_first = sheets.get(next(iter(sheets)))  # first available sheet
        if s_first is None:
            continue
        df = s_first.fillna("")
        if df.empty:
            continue
        # Transform to array of arrays
        try:
            values = df.values.tolist()
        except Exception:
            continue
        if len(values) < 2:
            continue
        header_row = [str(v).strip() for v in values[0]]
        # assume first col contains source role labels
        targets = [t for t in header_row[1:] if t]
        if not targets:
            continue
        out = []
        for r in range(1, len(values)):
            row = values[r]
            src = str(row[0]).strip() if len(row) > 0 else ""
            if not src:
                continue
            src_code = slugify(src)
            for c in range(1, len(header_row)):
                tgt_name = header_row[c] if c < len(header_row) else ""
                if not tgt_name:
                    continue
                val = row[c] if c < len(row) else None
                sc = safe_number(val)
                if sc is None:
                    continue
                out.append({
                    "source_role_code": src_code,
                    "target_role_code": slugify(tgt_name),
                    "score": sc
                })
        if out:
            return out
    return []


def _read_text(file_path: str) -> str:
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            return f.read()
    except Exception:
        return ""


# PUBLIC_INTERFACE
def parse_role_cards(cards_glob: str) -> List[dict]:
    """
    Parse role cards from .txt files.
    Returns list of {title, content, summary}
    - title: first line (role name)
    - content: entire file content (trim to 15000 chars to be safe)
    - summary: first non-empty paragraph limited to ~600 chars
    """
    files = glob.glob(cards_glob)
    out: List[dict] = []
    for fp in files:
        content = _read_text(fp)
        if not content:
            continue
        lines = content.splitlines()
        title = lines[0].strip() if lines else ""
        # summary = first non-empty paragraph
        paras = [p.strip() for p in content.split("\n\n") if p.strip()]
        summary = (paras[0][:600] if paras else content[:600])
        out.append({
            "title": title,
            "content": content[:15000],
            "summary": summary
        })
    return out


# ---------- Seeding helpers ----------
def chunked(seq: List[dict], chunk_size: int = 500) -> Iterable[List[dict]]:
    for i in range(0, len(seq), chunk_size):
        yield seq[i:i + chunk_size]


# PUBLIC_INTERFACE
def ensure_tables_exist(client: SupabaseRestClient, tables: List[str]) -> None:
    """Quick existence check for required tables."""
    missing = []
    for t in tables:
        try:
            _ = client.get_select(t, "id", limit=1)
        except Exception as e:
            msg = str(e).lower()
            if "does not exist" in msg or "not find the table" in msg or "schema cache" in msg:
                missing.append(t)
            else:
                # Unexpected error - still consider missing to avoid noisy loops
                missing.append(t)
    if missing:
        raise RuntimeError(f"Required tables missing: {', '.join(missing)}. Apply schema first.")


# PUBLIC_INTERFACE
def upsert_roles(client: SupabaseRestClient, roles: List[dict]) -> int:
    """Upsert roles on (code)."""
    total = 0
    for part in chunked(roles, 500):
        def _do():
            client.upsert("roles", part, on_conflict="code")
        with_retries(_do, f"upsert roles [{total+1}-{total+len(part)}]")
        total += len(part)
    return total


# PUBLIC_INTERFACE
def upsert_competencies(client: SupabaseRestClient, comps: List[dict]) -> int:
    """Upsert competencies on (code)."""
    total = 0
    for part in chunked(comps, 500):
        def _do():
            client.upsert("competencies", part, on_conflict="code")
        with_retries(_do, f"upsert competencies [{total+1}-{total+len(part)}]")
        total += len(part)
    return total


# PUBLIC_INTERFACE
def fetch_id_map(client: SupabaseRestClient, table: str, code_column: str = "code") -> Dict[str, str]:
    """Return a dict mapping natural code -> id for a table (roles or competencies)."""
    rows = with_retries(lambda: client.get_select(table, f"id,{code_column}", limit=20000), f"fetch {table} id map")
    out = {}
    for r in rows:
        k = r.get(code_column)
        if k:
            out[str(k)] = r.get("id")
    return out


# PUBLIC_INTERFACE
def upsert_role_competencies(client: SupabaseRestClient, rc_rows: List[dict]) -> int:
    """Upsert role_competencies on (role_id,competency_id)."""
    total = 0
    for part in chunked(rc_rows, 500):
        def _do():
            client.upsert("role_competencies", part, on_conflict="role_id,competency_id")
        with_retries(_do, f"upsert role_competencies [{total+1}-{total+len(part)}]")
        total += len(part)
    return total


# PUBLIC_INTERFACE
def upsert_role_adjacency(client: SupabaseRestClient, adj_rows: List[dict]) -> int:
    """Upsert role_adjacency on (source_role_id,target_role_id)."""
    total = 0
    for part in chunked(adj_rows, 500):
        def _do():
            client.upsert("role_adjacency", part, on_conflict="source_role_id,target_role_id")
        with_retries(_do, f"upsert role_adjacency [{total+1}-{total+len(part)}]")
        total += len(part)
    return total


# PUBLIC_INTERFACE
def upsert_learning_resources(client: SupabaseRestClient, lr_rows: List[dict]) -> int:
    """Upsert learning_resources on (competency_id,url)."""
    total = 0
    for part in chunked(lr_rows, 500):
        def _do():
            client.upsert("learning_resources", part, on_conflict="competency_id,url")
        with_retries(_do, f"upsert learning_resources [{total+1}-{total+len(part)}]")
        total += len(part)
    return total


# PUBLIC_INTERFACE
def upsert_role_descriptions(client: SupabaseRestClient, rd_rows: List[dict]) -> int:
    """Upsert role_descriptions on (role_id,source)."""
    total = 0
    for part in chunked(rd_rows, 500):
        def _do():
            client.upsert("role_descriptions", part, on_conflict="role_id,source")
        with_retries(_do, f"upsert role_descriptions [{total+1}-{total+len(part)}]")
        total += len(part)
    return total


# ---------- CLI and main ----------
def _build_arg_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Seed Supabase from Excel attachments using REST and idempotent upserts.")
    # Optional Supabase overrides (environment is preferred)
    p.add_argument("--supabase-url", dest="supabase_url", type=str, default=None,
                   help="Override Supabase URL (fallback to SUPABASE_URL/REACT_APP_SUPABASE_URL)")
    p.add_argument("--supabase-key", dest="supabase_key", type=str, default=None,
                   help="Override Supabase Service Role Key (fallback to SUPABASE_SERVICE_ROLE_KEY/REACT_APP_SUPABASE_SERVICE_ROLE_KEY/REACT_APP_SUPABASE_KEY)")

    p.add_argument("--competency", dest="competency_xlsx", type=str, default=DEFAULTS["competency_xlsx"],
                   help="Path to Competency_mapping.xlsx")
    p.add_argument("--adjacency", dest="adjacency_xlsx", type=str, default=DEFAULTS["adjacency_xlsx"],
                   help="Path to role adjacency workbook (row-wise or matrix)")
    p.add_argument("--roles", dest="roles_xlsx", type=str, default=DEFAULTS["roles_xlsx"],
                   help="Path to Role Navigator Worksheet workbook")
    p.add_argument("--resources", dest="resources_xlsx", type=str, default=DEFAULTS["resources_xlsx"],
                   help="Optional path to Learning Resources workbook (otherwise try roles workbook)")
    p.add_argument("--cardsGlob", dest="cards_glob", type=str, default=DEFAULTS["cards_glob"],
                   help="Glob for role card .txt files to generate role_descriptions")
    p.add_argument("--only", dest="entities", type=str, default="all",
                   help="Comma-separated list of entities to seed: roles,competencies,role_competencies,role_descriptions,role_adjacency,learning_resources. Default: all")
    p.add_argument("--log-level", dest="log_level", type=str, default="INFO",
                   help="Logging level (DEBUG, INFO, WARNING, ERROR)")
    return p


def _resolve_env() -> Tuple[Optional[str], Optional[str]]:
    url = os.getenv("SUPABASE_URL") or os.getenv("REACT_APP_SUPABASE_URL")
    # Accept both REACT_APP_SUPABASE_SERVICE_ROLE_KEY and REACT_APP_SUPABASE_KEY fallbacks
    key = (
        os.getenv("SUPABASE_SERVICE_ROLE_KEY")
        or os.getenv("REACT_APP_SUPABASE_SERVICE_ROLE_KEY")
        or os.getenv("REACT_APP_SUPABASE_KEY")
    )
    return url, key


# PUBLIC_INTERFACE
def main():
    """Entry point for the Python seeding CLI."""
    parser = _build_arg_parser()
    args = parser.parse_args()

    LOG.setLevel(getattr(logging, args.log_level.upper(), logging.INFO))

    supabase_url, supabase_key = _resolve_env()
    # Allow explicit CLI override
    if getattr(args, "supabase_url", None):
        supabase_url = args.supabase_url
    if getattr(args, "supabase_key", None):
        supabase_key = args.supabase_key

    if not supabase_url or not supabase_key:
        LOG.error("Missing SUPABASE_URL and/or SUPABASE_SERVICE_ROLE_KEY in environment (or --supabase-url/--supabase-key overrides). Aborting.")
        LOG.error("Note: REACT_APP_* fallbacks may be used if present, but service role key is preferred.")
        sys.exit(1)

    client = SupabaseRestClient(supabase_url, supabase_key, timeout=120)

    # Which entities to seed
    ent_raw = args.entities.strip().lower()
    if ent_raw == "all" or ent_raw == "":
        entities = {"roles", "competencies", "role_competencies", "role_descriptions", "role_adjacency", "learning_resources"}
    else:
        entities = set([e.strip() for e in ent_raw.split(",") if e.strip()])

    required_tables = [
        "roles", "competencies", "role_competencies", "role_adjacency", "learning_resources", "role_descriptions"
    ]
    try:
        ensure_tables_exist(client, required_tables)
    except Exception as e:
        LOG.error(f"Table existence check failed: {e}")
        sys.exit(1)

    # Aggregates
    roles_map: Dict[str, dict] = {}
    comps_map: Dict[str, dict] = {}
    role_comp: List[dict] = []
    adjacency: List[dict] = []
    learning: List[dict] = []
    role_descs: List[dict] = []

    # 1) Parse competency mapping (primary source for competencies, roles, role_competencies)
    if args.competency_xlsx and os.path.exists(args.competency_xlsx):
        LOG.info(f"Parsing competency mapping: {args.competency_xlsx}")
        comps, roles, rc = parse_competency_mapping(args.competency_xlsx)
        comps_map.update(comps)
        roles_map.update(roles)
        role_comp.extend(rc)
        LOG.info(f"Parsed: competencies={len(comps_map)}, roles={len(roles_map)}, role_competencies={len(role_comp)}")
    else:
        LOG.warning(f"Competency mapping file not found: {args.competency_xlsx}")

    # 1b) Supplement roles from Role Navigator Worksheet
    if args.roles_xlsx and os.path.exists(args.roles_xlsx):
        extras = parse_additional_roles_from_worksheet(args.roles_xlsx)
        before = len(roles_map)
        for r in extras:
            if r["code"] not in roles_map:
                roles_map[r["code"]] = r
        LOG.info(f"Added {len(roles_map) - before} additional roles from worksheet.")
    else:
        if args.roles_xlsx:
            LOG.warning(f"Roles worksheet not found: {args.roles_xlsx}")

    # 2) Parse role adjacency (with fallback alternate file)
    candidates = [args.adjacency_xlsx, DEFAULTS["adjacency_xlsx_alt"]]
    adj_rows = parse_role_adjacency([p for p in candidates if p and os.path.exists(p)])
    if adj_rows:
        adjacency.extend(adj_rows)
        # ensure roles exist
        for a in adj_rows:
            for code in [a["source_role_code"], a["target_role_code"]]:
                if code and code not in roles_map:
                    roles_map[code] = {"code": code, "name": code, "description": None}
        LOG.info(f"Parsed adjacency rows: {len(adj_rows)}")
    else:
        LOG.warning("No adjacency rows parsed from provided files.")

    # 3) Parse learning resources
    lr = parse_learning_resources(args.resources_xlsx, args.roles_xlsx, comps_map)
    if lr:
        learning.extend(lr)
        LOG.info(f"Parsed learning resources: {len(lr)}")

    # 4) Role descriptions from cards
    cards = parse_role_cards(args.cards_glob) if args.cards_glob else []
    if cards:
        LOG.info(f"Preparing role_descriptions from {len(cards)} role card text files...")
        # Try to map title to existing role codes
        # Build name->code lookup using both role name and code
        name_to_code = {}
        for r in roles_map.values():
            name_to_code[_norm(r["name"])] = r["code"]
            name_to_code[_norm(r["code"])] = r["code"]

        # Stage: fill roles.description with summary as a convenience, and prepare role_descriptions rows
        # Role IDs not known yet, so we stage content keyed by role_code; later we map to ids.
        staged_role_descs: List[dict] = []
        for c in cards:
            title = c["title"].strip()
            if not title:
                continue
            code_guess = slugify(title)
            role_code = name_to_code.get(_norm(title)) or name_to_code.get(_norm(code_guess)) or code_guess
            # Enrich roles map with summary if present
            if role_code not in roles_map:
                roles_map[role_code] = {"code": role_code, "name": title, "description": c["summary"]}
            else:
                roles_map[role_code]["description"] = roles_map[role_code].get("description") or c["summary"]
            staged_role_descs.append({
                "role_code": role_code,
                "source": "card",
                "summary": c["summary"],
                "content": c["content"]
            })
    else:
        staged_role_descs = []

    # Deduplicate arrays
    roles_arr = list({k: v for k, v in roles_map.items()}.values())
    comps_arr = list({k: v for k, v in comps_map.items()}.values())
    role_comp_arr = unique_by(role_comp, lambda x: f"{x['role_code']}|{x['competency_code']}")
    adjacency_arr = unique_by(adjacency, lambda x: f"{x['source_role_code']}|{x['target_role_code']}")
    learning_arr = unique_by(learning, lambda x: f"{x.get('competency_code') or 'none'}|{x['url']}")
    staged_role_descs = unique_by(staged_role_descs, lambda x: f"{x['role_code']}|{x['source']}")

    # Seeding order: roles -> competencies -> role_competencies -> role_adjacency -> learning_resources -> role_descriptions
    seeded_counts = {
        "roles": 0, "competencies": 0, "role_competencies": 0, "role_adjacency": 0, "learning_resources": 0, "role_descriptions": 0
    }

    # roles
    if "roles" in entities:
        LOG.info(f"Upserting roles: {len(roles_arr)}")
        if roles_arr:
            seeded_counts["roles"] = upsert_roles(client, roles_arr)

    # competencies
    if "competencies" in entities:
        LOG.info(f"Upserting competencies: {len(comps_arr)}")
        if comps_arr:
            seeded_counts["competencies"] = upsert_competencies(client, comps_arr)

    # id maps
    role_id_map = fetch_id_map(client, "roles", "code")
    comp_id_map = fetch_id_map(client, "competencies", "code")

    # role_competencies
    if "role_competencies" in entities:
        rc_rows = []
        for rc in role_comp_arr:
            rid = role_id_map.get(rc["role_code"])
            cid = comp_id_map.get(rc["competency_code"])
            if not rid or not cid:
                continue
            rc_rows.append({"role_id": rid, "competency_id": cid, "target_level": int(rc["target_level"])})
        LOG.info(f"Upserting role_competencies: {len(rc_rows)}")
        if rc_rows:
            seeded_counts["role_competencies"] = upsert_role_competencies(client, rc_rows)

    # role_adjacency
    if "role_adjacency" in entities:
        adj_rows = []
        for a in adjacency_arr:
            sid = role_id_map.get(a["source_role_code"])
            tid = role_id_map.get(a["target_role_code"])
            if not sid or not tid:
                continue
            row = {"source_role_id": sid, "target_role_id": tid}
            sc = safe_number(a.get("score"))
            if sc is not None:
                row["score"] = sc
            adj_rows.append(row)
        LOG.info(f"Upserting role_adjacency: {len(adj_rows)}")
        if adj_rows:
            seeded_counts["role_adjacency"] = upsert_role_adjacency(client, adj_rows)

    # learning_resources
    if "learning_resources" in entities:
        lr_rows = []
        for r in learning_arr:
            if not r.get("competency_code"):
                continue
            cid = comp_id_map.get(r["competency_code"])
            if not cid:
                continue
            row = {"competency_id": cid, "title": r["title"], "url": r["url"]}
            if r.get("provider"):
                row["provider"] = r["provider"]
            lr_rows.append(row)
        LOG.info(f"Upserting learning_resources: {len(lr_rows)}")
        if lr_rows:
            seeded_counts["learning_resources"] = upsert_learning_resources(client, lr_rows)

    # role_descriptions (from cards)
    if "role_descriptions" in entities:
        rd_rows = []
        for rd in staged_role_descs:
            rid = role_id_map.get(rd["role_code"])
            if not rid:
                continue
            row = {"role_id": rid, "source": rd["source"], "content": rd["content"]}
            if rd.get("summary"):
                row["summary"] = rd["summary"]
            rd_rows.append(row)
        LOG.info(f"Upserting role_descriptions: {len(rd_rows)}")
        if rd_rows:
            seeded_counts["role_descriptions"] = upsert_role_descriptions(client, rd_rows)

    LOG.info("Seeding complete.")
    LOG.info(json.dumps({
        "summary": seeded_counts,
        "entities": sorted(list(entities)),
    }, indent=2))


if __name__ == "__main__":
    main()
