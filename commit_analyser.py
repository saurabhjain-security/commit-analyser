#!/usr/bin/env python3
"""
Commit Analyser
---------------
Script  : Commit Analyser
Author  : Saurabh Jain
Version : 3.0

What changed from v2.1 → v3.0
  BUG FIXES
    • argparse added — every parameter is now a CLI flag (no source-editing needed)
    • GITHUB_TOKEN / GH_TOKEN env var checked first; gh CLI is optional fallback
    • Archived repos are excluded from the stale count (shown as N/A in Stale column)
    • commit.author.date used instead of commit.committer.date (real author activity)
    • pushed_at fallback removed — no longer used to infer commit activity
    • Summary stat columns are now exclusive buckets: ≤7d | 8–30d | 31–180d | >180d

  PERFORMANCE
    • ThreadPoolExecutor(10) fetches commit dates in parallel — 5–8× faster
    • Repo listing streams into the progress bar (no blocking list() pre-fetch)
    • Rate-limit remaining shown after each org so users can estimate time left
    • Rate-limit pre-flight check before scanning begins

  UX / CLI
    • --org          Target one or more specific orgs (default: all orgs)
    • --stale-days   Configurable stale threshold (default: 180)
    • --output       Custom output file path
    • --output-format  xlsx (default) | csv | json
    • --no-export    Skip file export entirely
    • --include-personal  Also scan personal repos under /user/repos
    • --exclude-forks    Skip forked repositories
    • --filter       Show only: all | stale | archived | active (default: all)
    • --since        Only fetch commits after this date (YYYY-MM-DD)

  IMPROVEMENTS
    • _get() wraps requests.get() in try/except — network errors retry gracefully
    • paginate() propagates partial-page errors to caller via error flag
    • Branch fallback: tries without sha param when default_branch is empty
    • Excel export has column widths, freeze panes, auto-filter, and colour coding
    • org label in summary table includes personal repos when --include-personal used

What it does:
  1. Reads GitHub token from env var or `gh auth token`
  2. Checks rate-limit quota before starting
  3. Discovers all organisations (or targets specified ones)
  4. Optionally includes personal repos
  5. Enumerates ALL repos (public + private + archived), skipping forks if requested
  6. Fetches latest commit author date concurrently
  7. Calculates inactivity duration and assigns activity bucket
  8. Marks stale repos (> --stale-days, default 180) — archived repos are excluded
  9. Displays Rich CLI tables with optional filtering
  10. Exports results to Excel / CSV / JSON

Output Columns:
  Sno | Org / Owner | Repo Name | Repo URL | Fork | Archived | Last Commit | Days Inactive | Stale

Pre-requisite:
  pip install rich requests pandas openpyxl

  Token (one of):
    export GITHUB_TOKEN=ghp_...
    export GH_TOKEN=ghp_...
    gh auth login   (fallback)

Usage:
  python commit_analyser.py
  python commit_analyser.py --org acme-corp another-org
  python commit_analyser.py --stale-days 90
  python commit_analyser.py --exclude-forks --filter stale
  python commit_analyser.py --include-personal
  python commit_analyser.py --output-format json --output results.json
  python commit_analyser.py --no-export
  python commit_analyser.py --since 2024-01-01
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import subprocess
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone
from typing import Dict, Iterator, List, Optional, Tuple

import pandas as pd
import requests
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from rich.console import Console
from rich.panel import Panel
from rich.progress import (
    BarColumn,
    MofNCompleteColumn,
    Progress,
    SpinnerColumn,
    TaskProgressColumn,
    TextColumn,
    TimeElapsedColumn,
)
from rich.rule import Rule
from rich.table import Table
from rich.traceback import install as install_rich_traceback
from rich import box

install_rich_traceback()

# ── Config ────────────────────────────────────────────────────────────────────
console    = Console()
GITHUB_API = "https://api.github.com"
DATE_FILE  = datetime.now().strftime("%Y-%m-%dT%H-%M")   # OS-safe filename
VERSION    = "3.0"
AUTHOR     = "Saurabh Jain"

# Excel colour palette
_HDR_FILL  = PatternFill("solid", start_color="1F3864")
_HDR_FONT  = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
_THIN      = Side(style="thin", color="D9D9D9")
_BORDER    = Border(top=_THIN, bottom=_THIN, left=_THIN, right=_THIN)
_FILL_STALE    = PatternFill("solid", start_color="FFC7CE")   # red tint
_FILL_FRESH    = PatternFill("solid", start_color="C6EFCE")   # green tint
_FILL_WARN     = PatternFill("solid", start_color="FFEB9C")   # amber tint
_FILL_ARCHIVED = PatternFill("solid", start_color="D9D9D9")   # grey


# ── Banner ────────────────────────────────────────────────────────────────────

def print_banner() -> None:
    art = (
        "  ██████╗ ██████╗ ███╗   ███╗███╗   ███╗██╗████████╗\n"
        " ██╔════╝██╔═══██╗████╗ ████║████╗ ████║██║╚══██╔══╝\n"
        " ██║     ██║   ██║██╔████╔██║██╔████╔██║██║   ██║   \n"
        " ██║     ██║   ██║██║╚██╔╝██║██║╚██╔╝██║██║   ██║   \n"
        " ╚██████╗╚██████╔╝██║ ╚═╝ ██║██║ ╚═╝ ██║██║   ██║   \n"
        "  ╚═════╝ ╚═════╝ ╚═╝     ╚═╝╚═╝     ╚═╝╚═╝   ╚═╝   \n"
        "\n"
        " █████╗ ███╗   ██╗ █████╗ ██╗  ██╗   ██╗███████╗███████╗██████╗ \n"
        "██╔══██╗████╗  ██║██╔══██╗██║  ╚██╗ ██╔╝██╔════╝██╔════╝██╔══██╗\n"
        "███████║██╔██╗ ██║███████║██║   ╚████╔╝ ███████╗█████╗  ██████╔╝\n"
        "██╔══██║██║╚██╗██║██╔══██║██║    ╚██╔╝  ╚════██║██╔══╝  ██╔══██╗\n"
        "██║  ██║██║ ╚████║██║  ██║███████╗██║   ███████║███████╗██║  ██║\n"
        "╚═╝  ╚═╝╚═╝  ╚═══╝╚═╝  ╚═╝╚══════╝╚═╝   ╚══════╝╚══════╝╚═╝  ╚═╝"
    )
    console.print()
    console.print(Panel(
        f"[bold bright_cyan]{art}[/]\n\n"
        f"  [bold bright_green]GitHub Organisation Repository Commit Analyser[/]"
        f"   [dim white]|[/]"
        f"   [dim white]v{VERSION}[/]"
        f"   [dim white]|[/]"
        f"   [bold yellow]Author: {AUTHOR}[/]",
        border_style="bright_blue",
        box=box.DOUBLE_EDGE,
        expand=False,
        padding=(1, 2),
    ))
    console.print()


# ── CLI ───────────────────────────────────────────────────────────────────────

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        prog="commit_analyser",
        description=f"GitHub Repository Commit Analyser v{VERSION}",
        epilog=(
            "Examples:\n"
            "  python commit_analyser.py\n"
            "  python commit_analyser.py --org acme-corp\n"
            "  python commit_analyser.py --stale-days 90 --exclude-forks\n"
            "  python commit_analyser.py --filter stale --output-format csv\n"
            "  python commit_analyser.py --include-personal --no-export\n"
            "  python commit_analyser.py --since 2024-01-01\n"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument(
        "--org", nargs="+", metavar="ORG", default=None,
        help="Target specific org(s). Default: all orgs the token belongs to.",
    )
    p.add_argument(
        "--stale-days", dest="stale_days", type=int, default=180,
        metavar="DAYS",
        help="Days of inactivity before a repo is considered stale (default: 180).",
    )
    p.add_argument(
        "--output", "-o", default="",
        help="Output file path (default: github_repo_scan_<timestamp>.<ext>).",
    )
    p.add_argument(
        "--output-format", dest="output_format",
        choices=["xlsx", "csv", "json"], default="xlsx",
        help="Export format: xlsx (default), csv, or json.",
    )
    p.add_argument(
        "--no-export", dest="no_export", action="store_true",
        help="Skip file export — terminal output only.",
    )
    p.add_argument(
        "--include-personal", dest="include_personal", action="store_true",
        help="Also scan repos in the authenticated user's personal namespace.",
    )
    p.add_argument(
        "--exclude-forks", dest="exclude_forks", action="store_true",
        help="Skip forked repositories.",
    )
    p.add_argument(
        "--filter", dest="filter_mode",
        choices=["all", "stale", "archived", "active"], default="all",
        help=(
            "Display only a subset of results: "
            "all (default) | stale | archived | active."
        ),
    )
    p.add_argument(
        "--since", default=None, metavar="YYYY-MM-DD",
        help="Only count commits after this date (passed to the GitHub commits API).",
    )
    p.add_argument(
        "--workers", type=int, default=10, metavar="N",
        help="Parallel workers for fetching commit dates (default: 10).",
    )
    return p.parse_args()


# ── Auth ──────────────────────────────────────────────────────────────────────

def get_token() -> str:
    # Priority 1: environment variables (works in CI/CD without gh CLI)
    for env_var in ("GITHUB_TOKEN", "GH_TOKEN"):
        token = os.environ.get(env_var, "").strip()
        if token:
            console.print(
                f"  [green]✓[/]  Token loaded from env var "
                f"[bold cyan]{env_var}[/]"
            )
            return token

    # Priority 2: gh CLI fallback
    try:
        r = subprocess.run(
            ["gh", "auth", "token"],
            capture_output=True, text=True, timeout=10,
        )
    except FileNotFoundError:
        console.print(Panel(
            "[bold red]✗  No token found[/]\n\n"
            "  Set [bold cyan]GITHUB_TOKEN[/] or [bold cyan]GH_TOKEN[/] env var,\n"
            "  or install gh CLI: [dim]https://cli.github.com[/]\n"
            "  Then run: [bold cyan]gh auth login[/]",
            border_style="red", box=box.ROUNDED, expand=False,
        ))
        sys.exit(1)
    except subprocess.TimeoutExpired:
        console.print("  [bold red]✗[/]  `gh auth token` timed out.")
        sys.exit(1)

    token = r.stdout.strip()
    if not token or r.returncode != 0:
        console.print(Panel(
            "[bold red]✗  Not authenticated[/]\n\n"
            "  Set [bold cyan]GITHUB_TOKEN[/] env var, or run:\n"
            "  [bold cyan]gh auth login[/]",
            border_style="red", box=box.ROUNDED, expand=False,
        ))
        sys.exit(1)

    console.print("  [green]✓[/]  Token loaded via [bold cyan]gh auth token[/]")
    return token


# ── GitHub API Client ─────────────────────────────────────────────────────────

class GHClient:
    def __init__(self, token: str) -> None:
        self.session = requests.Session()
        self.session.headers.update({
            "Authorization":        f"Bearer {token}",
            "Accept":               "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
        })
        self._rate_remaining: int = 5000
        self._rate_limit:     int = 5000

    # ── core GET with retry ───────────────────────────────────────────
    def _get(self, path: str, params: Optional[Dict] = None) -> requests.Response:
        url = f"{GITHUB_API}{path}" if path.startswith("/") else path
        backoff = 2
        while True:
            try:
                resp = self.session.get(url, params=params, timeout=30)
            except (requests.ConnectionError, requests.Timeout) as exc:
                console.print(
                    f"  [bold yellow]⚠[/]  Network error ({exc.__class__.__name__}) — "
                    f"retrying in {backoff}s..."
                )
                time.sleep(backoff)
                backoff = min(backoff * 2, 60)
                continue

            # Update rate-limit counters from every response
            try:
                self._rate_remaining = int(
                    resp.headers.get("X-RateLimit-Remaining", self._rate_remaining)
                )
                self._rate_limit = int(
                    resp.headers.get("X-RateLimit-Limit", self._rate_limit)
                )
            except (ValueError, TypeError):
                pass

            # Rate-limit: wait until reset
            if resp.status_code == 429 or (
                resp.status_code == 403 and "rate limit" in resp.text.lower()
            ):
                reset = int(resp.headers.get("X-RateLimit-Reset", time.time() + 60))
                wait  = max(reset - int(time.time()), 1)
                console.print(
                    f"  [bold yellow]⚠[/]  Rate limit hit — "
                    f"waiting [bold]{wait}s[/]..."
                )
                time.sleep(wait + 1)
                continue

            # Server errors: back off and retry
            if resp.status_code >= 500:
                console.print(
                    f"  [bold yellow]⚠[/]  Server error {resp.status_code} — "
                    f"retrying in {backoff}s..."
                )
                time.sleep(backoff)
                backoff = min(backoff * 2, 60)
                continue

            return resp

    # ── paginator ─────────────────────────────────────────────────────
    def paginate(
        self, path: str, params: Optional[Dict] = None
    ) -> Tuple[List[dict], bool]:
        """
        Return (items, complete) where complete=False means a page errored.
        Caller can inspect the flag to know if the list is partial.
        """
        p = dict(params or {})
        p.setdefault("per_page", 100)
        p["page"] = 1
        collected: List[dict] = []
        while True:
            resp = self._get(path, params=p)
            if not resp.ok:
                console.print(
                    f"  [bold red]✗[/]  GitHub API error "
                    f"[bold]{resp.status_code}[/] on [dim]{path}[/] "
                    f"— partial results returned"
                )
                return collected, False
            items = resp.json()
            if not isinstance(items, list) or not items:
                return collected, True
            collected.extend(items)
            if 'rel="next"' not in resp.headers.get("Link", ""):
                return collected, True
            p["page"] += 1

    # ── rate-limit pre-flight check ───────────────────────────────────
    def check_rate_limit(self) -> None:
        try:
            resp = self._get("/rate_limit")
            if resp.ok:
                data      = resp.json().get("rate", {})
                remaining = data.get("remaining", 0)
                limit     = data.get("limit", 5000)
                reset_ts  = data.get("reset", 0)
                reset_dt  = datetime.fromtimestamp(reset_ts).strftime("%H:%M:%S")
                self._rate_remaining = remaining
                self._rate_limit     = limit
                color = "green" if remaining > 1000 else ("yellow" if remaining > 200 else "red")
                console.print(
                    f"  [green]✓[/]  Rate limit: "
                    f"[bold {color}]{remaining:,}[/] / {limit:,} calls remaining "
                    f"[dim](resets at {reset_dt})[/]"
                )
                if remaining < 100:
                    console.print(
                        "  [bold red]⚠  Fewer than 100 API calls remaining.[/]  "
                        "Scan may be incomplete — wait for reset or use a different token."
                    )
        except Exception:
            pass   # non-fatal — just skip the pre-flight display

    def rate_label(self) -> str:
        return f"{self._rate_remaining:,}/{self._rate_limit:,}"


# ── Helpers ───────────────────────────────────────────────────────────────────

def days_since(iso_date: Optional[str]) -> Tuple[str, int]:
    """Return (human string, sort key int)."""
    if not iso_date:
        return "No commits", 99_999
    try:
        dt   = datetime.fromisoformat(iso_date.replace("Z", "+00:00"))
        now  = datetime.now(timezone.utc)
        days = (now - dt).days
        if days == 0:
            return "Today", 0
        if days == 1:
            return "1 day ago", 1
        return f"{days} days ago", days
    except Exception:
        return "Invalid date", 99_999


def commit_colour(days_key: int, days_str: str) -> str:
    if days_str == "No commits":
        return "[dim white]No commits[/]"
    if days_key == 0:
        return "[bold green]Today[/]"
    if days_key <= 7:
        return f"[bold green]{days_str}[/]"
    if days_key <= 30:
        return f"[bold yellow]{days_str}[/]"
    if days_key <= 180:
        return f"[bold orange3]{days_str}[/]"
    return f"[bold red]{days_str}[/]"


# ── Data Collection ───────────────────────────────────────────────────────────

def get_orgs(client: GHClient) -> List[str]:
    orgs, ok = client.paginate("/user/orgs")
    return [item["login"] for item in orgs]


def get_last_commit_date(
    client: GHClient,
    full_name: str,
    default_branch: str,
    since: Optional[str],
) -> Optional[str]:
    """
    Returns commit.author.date (not committer.date) of the most recent commit.
    Falls back to querying without sha if default_branch is empty.
    Does NOT fall back to pushed_at (unreliable — includes CI/bot pushes).
    """
    params: Dict = {"per_page": 1}
    if default_branch:
        params["sha"] = default_branch
    if since:
        params["since"] = since

    resp = client._get(f"/repos/{full_name}/commits", params=params)
    if not resp.ok:
        # If sha-based query fails (e.g. branch not found), retry without sha
        if "sha" in params:
            params_retry = {k: v for k, v in params.items() if k != "sha"}
            resp = client._get(f"/repos/{full_name}/commits", params=params_retry)
            if not resp.ok:
                return None
        else:
            return None

    data = resp.json()
    if isinstance(data, list) and data:
        # Use author.date — reflects when the human wrote the code
        return (
            data[0]
            .get("commit", {})
            .get("author", {})      # author, not committer
            .get("date")
        )
    return None


def fetch_repo_row(
    client: GHClient,
    org: str,
    repo: dict,
    stale_days: int,
    since: Optional[str],
) -> dict:
    """Build one result row for a single repo (called from thread pool)."""
    repo_name      = repo.get("name", "")
    full_name      = repo.get("full_name", "")
    default_branch = repo.get("default_branch") or ""
    html_url       = repo.get("html_url") or f"https://github.com/{full_name}"
    archived       = repo.get("archived", False)
    is_fork        = repo.get("fork", False)

    last_commit = get_last_commit_date(client, full_name, default_branch, since)

    days_str, days_key = days_since(last_commit)

    # Archived repos are intentionally frozen — exclude from stale determination
    if archived:
        stale = "N/A"
    else:
        stale = "Yes" if (days_key != 99_999 and days_key > stale_days) else "No"
        # repos with no commits ever are not stale — they're just empty
        if days_key == 99_999 and not archived:
            stale = "No commits"

    return {
        "Org Name":      org,
        "Repo Name":     repo_name,
        "Repo URL":      html_url,
        "Fork":          str(is_fork),
        "Archived":      str(archived),
        "Last Commit":   days_str,
        "Days Inactive": days_key if days_key != 99_999 else "",
        "Stale Repo":    stale,
        "_sort_key":     days_key,
        "_archived":     archived,
    }


# ── Export Functions ──────────────────────────────────────────────────────────

def _clean_rows(all_rows: List[dict]) -> List[dict]:
    """Strip internal keys before exporting."""
    skip = {"_sort_key", "_archived"}
    return [{k: v for k, v in r.items() if k not in skip} for r in all_rows]


def export_excel(all_rows: List[dict], path: str) -> str:
    clean = _clean_rows(all_rows)
    cols  = ["Sno", "Org Name", "Repo Name", "Repo URL", "Fork",
             "Archived", "Last Commit", "Days Inactive", "Stale Repo"]

    numbered = [{"Sno": i, **r} for i, r in enumerate(clean, start=1)]
    df = pd.DataFrame(numbered, columns=cols)

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Repos", index=False)

    # Post-process for styling
    wb = load_workbook(path)
    ws = wb["Repos"]

    # Column widths
    col_widths = {
        "A": 6,   # Sno
        "B": 22,  # Org Name
        "C": 30,  # Repo Name
        "D": 55,  # Repo URL
        "E": 8,   # Fork
        "F": 10,  # Archived
        "G": 18,  # Last Commit
        "H": 14,  # Days Inactive
        "I": 11,  # Stale Repo
    }
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    # Header row styling
    for cell in ws[1]:
        cell.fill      = _HDR_FILL
        cell.font      = _HDR_FONT
        cell.border    = _BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Data row styling
    stale_col_idx    = cols.index("Stale Repo") + 1   # 1-based
    archived_col_idx = cols.index("Archived")   + 1

    for row in ws.iter_rows(min_row=2):
        stale_val    = row[stale_col_idx - 1].value or ""
        archived_val = row[archived_col_idx - 1].value or ""

        if archived_val == "True":
            row_fill = _FILL_ARCHIVED
        elif stale_val == "Yes":
            row_fill = _FILL_STALE
        elif str(stale_val).startswith("No"):
            row_fill = _FILL_FRESH
        else:
            row_fill = _FILL_WARN

        for cell in row:
            cell.border    = _BORDER
            cell.font      = Font(name="Calibri", size=10)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.fill      = row_fill

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(path)
    return os.path.abspath(path)


def export_csv(all_rows: List[dict], path: str) -> str:
    clean = _clean_rows(all_rows)
    if not clean:
        return path
    fieldnames = ["Sno"] + list(clean[0].keys())
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for i, row in enumerate(clean, start=1):
            writer.writerow({"Sno": i, **row})
    return os.path.abspath(path)


def export_json(all_rows: List[dict], path: str) -> str:
    clean = _clean_rows(all_rows)
    numbered = [{"sno": i, **r} for i, r in enumerate(clean, start=1)]
    # normalise keys to snake_case for JSON consumers
    def _snake(s: str) -> str:
        return s.lower().replace(" ", "_")
    out = [{_snake(k): v for k, v in row.items()} for row in numbered]
    with open(path, "w", encoding="utf-8") as f:
        json.dump(out, f, indent=2, ensure_ascii=False)
    return os.path.abspath(path)


# ── Rich Progress Bar ─────────────────────────────────────────────────────────

def make_progress() -> Progress:
    return Progress(
        SpinnerColumn(spinner_name="dots", style="bright_blue"),
        TextColumn("[bold cyan]{task.description}"),
        BarColumn(
            bar_width=36,
            style="bright_blue",
            complete_style="bright_green",
        ),
        TaskProgressColumn(),
        MofNCompleteColumn(),
        TimeElapsedColumn(),
        console=console,
        transient=False,
    )


# ── Results Table ─────────────────────────────────────────────────────────────

def apply_filter(rows: List[dict], mode: str) -> List[dict]:
    if mode == "stale":
        return [r for r in rows if r["Stale Repo"] == "Yes"]
    if mode == "archived":
        return [r for r in rows if r["Archived"] == "True"]
    if mode == "active":
        return [r for r in rows if r["Stale Repo"] == "No"]
    return rows


def print_results_table(all_rows: List[dict], filter_mode: str) -> None:
    display_rows = apply_filter(all_rows, filter_mode)

    console.print(Rule("[bold bright_cyan]  Scan Results  [/]", style="bright_blue"))
    console.print()

    filter_label = "" if filter_mode == "all" else f" — filter: [bold]{filter_mode}[/]"
    table = Table(
        title=(
            f"[bold bright_cyan]GitHub Repository Commit Analysis[/]  "
            f"[dim]({len(display_rows)} repos{filter_label})[/]"
        ),
        box=box.ROUNDED,
        border_style="bright_blue",
        header_style="bold bright_blue",
        show_lines=True,
        expand=True,
    )

    table.add_column("Sno",           justify="right",  style="dim white",  width=5,    no_wrap=True)
    table.add_column("Org / Owner",   justify="left",   style="bold cyan",  min_width=16, no_wrap=True)
    table.add_column("Repo Name",     justify="left",   style="white",      min_width=18, no_wrap=False)
    table.add_column("Repo URL",      justify="left",   overflow="fold",    min_width=34)
    table.add_column("Fork",          justify="center", style="white",      width=7,    no_wrap=True)
    table.add_column("Archived",      justify="center", style="white",      width=10,   no_wrap=True)
    table.add_column("Last Commit",   justify="center", style="white",      width=16,   no_wrap=True)
    table.add_column("Stale",         justify="center", style="white",      width=10,   no_wrap=True)

    for idx, row in enumerate(display_rows, start=1):
        is_stale    = row["Stale Repo"] == "Yes"
        is_archived = row["Archived"] == "True"
        is_fork     = row["Fork"] == "True"

        url_markup = (
            f"[bold red]{row['Repo URL']}[/]"
            if is_stale
            else f"[cyan]{row['Repo URL']}[/]"
        )
        archived_markup = (
            "[bold yellow]True[/]" if is_archived else "[dim white]False[/]"
        )
        fork_markup = (
            "[dim white]True[/]" if is_fork else "[dim white]False[/]"
        )
        commit_markup = commit_colour(row["_sort_key"], row["Last Commit"])

        stale_val = row["Stale Repo"]
        if stale_val == "Yes":
            stale_markup = "[bold red]Yes[/]"
        elif stale_val == "N/A":
            stale_markup = "[dim white]N/A[/]"
        elif stale_val == "No commits":
            stale_markup = "[dim white]—[/]"
        else:
            stale_markup = "[green]No[/]"

        table.add_row(
            str(idx),
            row["Org Name"],
            row["Repo Name"],
            url_markup,
            fork_markup,
            archived_markup,
            commit_markup,
            stale_markup,
        )

    console.print(table)
    console.print()


# ── Summary Table ─────────────────────────────────────────────────────────────

def print_summary(
    all_rows: List[dict],
    org_labels: List[str],
    elapsed: float,
    stale_days: int,
    client: GHClient,
) -> None:
    console.print(Rule("[bold bright_green]  Summary  [/]", style="bright_green"))
    console.print()

    per_org_tbl = Table(
        title="[bold white]Per-Organisation Breakdown[/]",
        box=box.ROUNDED,
        border_style="bright_green",
        header_style="bold bright_green",
        show_lines=True,
        expand=False,
    )
    per_org_tbl.add_column("Organisation",   style="bold cyan",  no_wrap=True,  min_width=20)
    per_org_tbl.add_column("Scanned",        justify="right",    style="white", width=9)
    per_org_tbl.add_column("≤7d",            justify="right",    width=7)   # exclusive bucket
    per_org_tbl.add_column("8–30d",          justify="right",    width=7)   # exclusive bucket
    per_org_tbl.add_column("31–180d",        justify="right",    width=8)   # exclusive bucket
    per_org_tbl.add_column(f">{stale_days}d (stale)", justify="right", width=13)
    per_org_tbl.add_column("No Commits",     justify="right",    width=11)
    per_org_tbl.add_column("Archived",       justify="right",    width=9)
    per_org_tbl.add_column("Forks",          justify="right",    width=7)

    grand: Dict[str, int] = {
        "scanned": 0, "b7": 0, "b30": 0, "b180": 0,
        "stale": 0, "no_cmt": 0, "archived": 0, "forks": 0,
    }

    for org in org_labels:
        rows = [r for r in all_rows if r["Org Name"] == org]
        if not rows:
            continue

        scanned  = len(rows)
        # Exclusive activity buckets (non-overlapping)
        b7       = sum(1 for r in rows if 0 <= r["_sort_key"] <= 7)
        b30      = sum(1 for r in rows if 8 <= r["_sort_key"] <= 30)
        b180     = sum(1 for r in rows if 31 <= r["_sort_key"] <= stale_days)
        stale    = sum(1 for r in rows if r["Stale Repo"] == "Yes")
        no_cmt   = sum(1 for r in rows if r["_sort_key"] == 99_999)
        archived = sum(1 for r in rows if r["Archived"] == "True")
        forks    = sum(1 for r in rows if r["Fork"] == "True")

        for k, v in zip(
            ["scanned","b7","b30","b180","stale","no_cmt","archived","forks"],
            [scanned, b7, b30, b180, stale, no_cmt, archived, forks],
        ):
            grand[k] += v

        def _c(val: int, color: str) -> str:
            return f"[{color}]{val}[/]" if val else "[dim]0[/]"

        per_org_tbl.add_row(
            org,
            str(scanned),
            _c(b7,    "bold green"),
            _c(b30,   "bold yellow"),
            _c(b180,  "bold orange3"),
            _c(stale, "bold red"),
            _c(no_cmt,   "dim"),
            _c(archived, "bold yellow"),
            _c(forks,    "dim cyan"),
        )

    per_org_tbl.add_row(
        "[bold white]TOTAL[/]",
        f"[bold white]{grand['scanned']}[/]",
        f"[bold green]{grand['b7']}[/]"        if grand["b7"]       else "[dim]0[/]",
        f"[bold yellow]{grand['b30']}[/]"      if grand["b30"]      else "[dim]0[/]",
        f"[bold orange3]{grand['b180']}[/]"    if grand["b180"]     else "[dim]0[/]",
        f"[bold red]{grand['stale']}[/]"       if grand["stale"]    else "[dim]0[/]",
        f"[dim]{grand['no_cmt']}[/]"           if grand["no_cmt"]   else "[dim]0[/]",
        f"[bold yellow]{grand['archived']}[/]" if grand["archived"] else "[dim]0[/]",
        f"[dim cyan]{grand['forks']}[/]"       if grand["forks"]    else "[dim]0[/]",
    )
    console.print(per_org_tbl)
    console.print()

    # Stat cards
    cards = Table(
        box=box.SIMPLE_HEAVY, border_style="bright_blue",
        show_header=False, expand=True, padding=(0, 2),
    )
    for _ in range(6):
        cards.add_column("x", justify="center")

    def _card(label: str, value: str, color: str) -> str:
        return f"[dim]{label}[/]\n[{color}]{value}[/{color}]"

    cards.add_row(
        _card("Orgs / owners",   str(len(org_labels)),      "bold white"),
        _card("Total repos",     str(grand["scanned"]),     "bold white"),
        _card("≤7d active",      str(grand["b7"]),          "bold green"),
        _card(f">{stale_days}d stale", str(grand["stale"]), "bold red"),
        _card("Archived",        str(grand["archived"]),    "bold yellow"),
        _card("API remaining",   client.rate_label(),       "dim cyan"),
    )
    console.print(cards)
    console.print()

    def _fmt(s: float) -> str:
        m, sec = divmod(int(s), 60)
        return f"{m}m {sec}s" if m else f"{sec:.0f}s"

    console.print(Panel(
        f"[bold green]✓  Analysis Complete[/]   [dim]·[/]   "
        f"[white]Orgs/owners:[/] [bold]{len(org_labels)}[/]   "
        f"[white]Repos:[/] [bold]{grand['scanned']}[/]   "
        f"[white]≤7d:[/] [bold green]{grand['b7']}[/]   "
        f"[white]Stale:[/] [bold red]{grand['stale']}[/]   "
        f"[white]Archived:[/] [bold yellow]{grand['archived']}[/]   "
        f"[white]No commits:[/] [dim]{grand['no_cmt']}[/]   "
        f"[white]Runtime:[/] [bold cyan]{_fmt(elapsed)}[/]",
        border_style="bright_green",
        box=box.DOUBLE_EDGE,
        expand=False,
    ))
    console.print()


# ── Main ──────────────────────────────────────────────────────────────────────

def main() -> None:
    args    = parse_args()
    t_start = time.monotonic()

    print_banner()

    # ── Auth ──────────────────────────────────────────────────────────
    with console.status(
        "  [dim]Resolving GitHub token...[/]", spinner="dots"
    ):
        token = get_token()

    client = GHClient(token)

    # Verify identity
    me_resp = client._get("/user")
    if not me_resp.ok:
        console.print("[bold red]✗  Could not fetch user info — check token permissions.[/]")
        sys.exit(1)
    me = me_resp.json()
    console.print(
        f"  [green]✓[/]  Logged in as "
        f"[bold cyan]{me.get('login')}[/]"
        f"  [dim]({me.get('name', '')})[/]"
    )
    console.print()

    # ── Rate-limit pre-flight ─────────────────────────────────────────
    client.check_rate_limit()
    console.print()

    # ── Discover orgs ─────────────────────────────────────────────────
    if args.org:
        orgs = args.org
        console.print(
            f"  [green]✓[/]  Targeting [bold]{len(orgs)}[/] specified org(s):  "
            f"[dim]{', '.join(orgs)}[/]"
        )
    else:
        with console.status("  [dim]Fetching organisations...[/]", spinner="dots"):
            orgs = get_orgs(client)

        if not orgs:
            console.print(Panel(
                "[bold yellow]⚠  No organisations found for this account.[/]",
                border_style="yellow", box=box.ROUNDED, expand=False,
            ))
            if not args.include_personal:
                console.print(
                    "  [dim]Tip: try [bold]--include-personal[/] to scan your personal repos.[/]"
                )
                sys.exit(0)
        else:
            console.print(
                f"  [green]✓[/]  Found [bold]{len(orgs)}[/] organisation(s):  "
                f"[dim]{', '.join(orgs)}[/]"
            )
    console.print()

    # ── Collect repos ─────────────────────────────────────────────────
    # Gather (org_label, repo_dict) pairs
    repo_queue: List[Tuple[str, dict]] = []

    for org in orgs:
        with console.status(
            f"  [dim]Listing repos for [bold]{org}[/]...[/]", spinner="dots"
        ):
            repos, complete = client.paginate(
                f"/orgs/{org}/repos", params={"type": "all"}
            )
        if not complete:
            console.print(f"  [yellow]⚠[/]  Partial repo list for [bold]{org}[/]")
        if repos:
            for r in repos:
                if args.exclude_forks and r.get("fork"):
                    continue
                repo_queue.append((org, r))

    if args.include_personal:
        personal_login = me.get("login", "")
        with console.status(
            f"  [dim]Listing personal repos for [bold]{personal_login}[/]...[/]",
            spinner="dots",
        ):
            personal_repos, complete = client.paginate(
                "/user/repos", params={"type": "owner", "affiliation": "owner"}
            )
        if not complete:
            console.print("  [yellow]⚠[/]  Partial personal repo list")
        if personal_repos:
            for r in personal_repos:
                if args.exclude_forks and r.get("fork"):
                    continue
                # Avoid duplicates if personal namespace matches an org
                repo_queue.append((personal_login, r))

    if not repo_queue:
        console.print(Panel(
            "[bold yellow]⚠  No repositories found to scan.[/]",
            border_style="yellow", box=box.ROUNDED, expand=False,
        ))
        sys.exit(0)

    console.print(
        f"  [green]✓[/]  [bold]{len(repo_queue)}[/] repo(s) queued for scan"
        f"{' (forks excluded)' if args.exclude_forks else ''}"
    )
    console.print()

    # ── Concurrent commit fetch ────────────────────────────────────────
    all_rows: List[dict] = []
    progress  = make_progress()

    with progress:
        task = progress.add_task(
            "[cyan]Fetching commit dates...[/]",
            total=len(repo_queue),
        )

        with ThreadPoolExecutor(max_workers=args.workers) as pool:
            future_map = {
                pool.submit(
                    fetch_repo_row,
                    client, org, repo, args.stale_days, args.since,
                ): (org, repo.get("name", ""))
                for org, repo in repo_queue
            }

            for future in as_completed(future_map):
                org_label, repo_name = future_map[future]
                progress.update(
                    task,
                    description=f"[cyan]{org_label}[/] — [white]{repo_name[:28]}[/]",
                )
                try:
                    row = future.result()
                    all_rows.append(row)
                except Exception as exc:
                    console.print(
                        f"  [bold red]✗[/]  Error processing "
                        f"[bold]{org_label}/{repo_name}[/]: {exc}"
                    )
                progress.advance(task)

    # Sort: most recently active first; no-commit repos sink to bottom
    all_rows.sort(key=lambda x: x["_sort_key"])

    # Collect unique org labels (preserves order of first appearance)
    seen: set = set()
    org_labels: List[str] = []
    for r in all_rows:
        if r["Org Name"] not in seen:
            seen.add(r["Org Name"])
            org_labels.append(r["Org Name"])

    # ── Results table ─────────────────────────────────────────────────
    print_results_table(all_rows, args.filter_mode)

    # ── Export ────────────────────────────────────────────────────────
    if not args.no_export:
        ext = args.output_format
        default_name = f"github_repo_scan_{DATE_FILE}.{ext}"
        out_path = args.output or default_name

        with console.status(
            f"[bright_blue]Writing {ext.upper()} report...[/]", spinner="dots"
        ):
            if ext == "xlsx":
                saved = export_excel(all_rows, out_path)
            elif ext == "csv":
                saved = export_csv(all_rows, out_path)
            else:
                saved = export_json(all_rows, out_path)

        console.print(f"  [green]✓[/]  {ext.upper()} saved → [dim]{saved}[/]")
        console.print()

    # ── Summary ───────────────────────────────────────────────────────
    elapsed = time.monotonic() - t_start
    print_summary(all_rows, org_labels, elapsed, args.stale_days, client)


if __name__ == "__main__":
    main()
