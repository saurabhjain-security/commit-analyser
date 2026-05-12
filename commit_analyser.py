#!/usr/bin/env python3
"""
Commit Analyser
---------------
Script  : Commit Analyser
Author  : Saurabh Jain
Version : 2.1

What it does:
  1. Reads GitHub token from `gh auth token`
  2. Discovers all organisations the authenticated user belongs to
  3. Enumerates ALL repos (public + private + archived)
  4. Fetches latest commit date
  5. Calculates inactivity duration
  6. Marks stale repos (>180 days)
  7. Detects archived repos
  8. Displays Rich CLI tables
  9. Exports results to Excel

Output Columns:
  Sno | Org Name | Repo Name | Repo URL | Archived | Last Commit | Stale Repo

Pre-requisite:
  gh auth login
  pip install rich requests pandas openpyxl
"""

from __future__ import annotations

import os
import subprocess
import sys
import time
from datetime import datetime, timezone
from typing import Dict, Iterator, List, Optional

import pandas as pd
import requests
from rich.console import Console
from rich.live import Live
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
from rich.text import Text
from rich.traceback import install as install_rich_traceback
from rich import box

install_rich_traceback()

# ── Config ────────────────────────────────────────────────────────────────────
console    = Console()
GITHUB_API = "https://api.github.com"
STALE_DAYS = 180
DATE_STR   = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
VERSION    = "2.1"
AUTHOR     = "Saurabh Jain"


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


# ── Auth ──────────────────────────────────────────────────────────────────────

def get_token() -> str:
    try:
        r = subprocess.run(
            ["gh", "auth", "token"],
            capture_output=True, text=True, timeout=10,
        )
    except FileNotFoundError:
        console.print(Panel(
            "[bold red]✗  gh CLI not found[/]\n\n"
            "  Install from: [dim]https://cli.github.com[/]\n"
            "  Then run:     [bold cyan]gh auth login[/]",
            border_style="red", box=box.ROUNDED, expand=False,
        ))
        sys.exit(1)
    except subprocess.TimeoutExpired:
        console.print("  [bold red]✗[/] `gh auth token` timed out.")
        sys.exit(1)

    token = r.stdout.strip()
    if not token or r.returncode != 0:
        console.print(Panel(
            "[bold red]✗  Not logged in to GitHub CLI[/]\n\n"
            "  Run: [bold cyan]gh auth login[/]",
            border_style="red", box=box.ROUNDED, expand=False,
        ))
        sys.exit(1)

    return token


# ── GitHub API Client ─────────────────────────────────────────────────────────

class GHClient:
    def __init__(self, token: str):
        self.session = requests.Session()
        self.session.headers.update({
            "Authorization":        f"Bearer {token}",
            "Accept":               "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
        })

    def _get(self, path: str, params: Optional[Dict] = None) -> requests.Response:
        url = f"{GITHUB_API}{path}" if path.startswith("/") else path
        while True:
            resp = self.session.get(url, params=params, timeout=30)
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
            return resp

    def paginate(self, path: str, params: Optional[Dict] = None) -> Iterator[dict]:
        """Yield all items across all pages (100 per page)."""
        p = dict(params or {})
        p.setdefault("per_page", 100)
        p["page"] = 1
        while True:
            resp = self._get(path, params=p)
            if not resp.ok:
                console.print(
                    f"  [bold red]✗[/]  GitHub API Error "
                    f"[bold]{resp.status_code}[/] on {path}"
                )
                return
            items = resp.json()
            if not isinstance(items, list) or not items:
                return
            yield from items
            if 'rel="next"' not in resp.headers.get("Link", ""):
                return
            p["page"] += 1


# ── Helpers ───────────────────────────────────────────────────────────────────

def days_since(iso_date: Optional[str]) -> str:
    if not iso_date:
        return "No commits"
    try:
        dt   = datetime.fromisoformat(iso_date.replace("Z", "+00:00"))
        now  = datetime.now(timezone.utc)
        days = (now - dt).days
        if days == 0:
            return "Today"
        if days == 1:
            return "1 day ago"
        return f"{days} days ago"
    except Exception:
        return "Invalid date"


def days_sort_key(days_str: str) -> int:
    if days_str == "Today":
        return 0
    if days_str in ("No commits", "Invalid date"):
        return 99_999
    try:
        return int(days_str.split()[0])
    except Exception:
        return 99_999


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
    return [item["login"] for item in client.paginate("/user/orgs")]


def get_last_commit_date(
    client: GHClient, full_name: str, default_branch: str
) -> Optional[str]:
    branch = default_branch or "main"
    resp   = client._get(
        f"/repos/{full_name}/commits",
        params={"per_page": 1, "sha": branch},
    )
    if not resp.ok:
        return None
    data = resp.json()
    if isinstance(data, list) and data:
        return (
            data[0]
            .get("commit", {})
            .get("committer", {})
            .get("date")
        )
    return None


# ── Excel Export ──────────────────────────────────────────────────────────────

def export_excel(rows: list) -> str:
    filename = f"github_repo_scan_{DATE_STR}.xlsx"
    df = pd.DataFrame(rows)
    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Repos", index=False)
    return os.path.abspath(filename)


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

def print_results_table(all_rows: list) -> None:
    console.print(Rule(
        "[bold bright_cyan]  Scan Results  [/]",
        style="bright_blue",
    ))
    console.print()

    table = Table(
        title=(
            f"[bold bright_cyan]GitHub Repository Commit Analysis[/]  "
            f"[dim]({len(all_rows)} repos)[/]"
        ),
        box=box.ROUNDED,
        border_style="bright_blue",
        header_style="bold bright_blue",
        show_lines=True,
        expand=True,
    )

    table.add_column("Sno",         justify="right",  style="dim white",   width=5,   no_wrap=True)
    table.add_column("Org Name",    justify="left",   style="bold cyan",   min_width=18, no_wrap=True)
    table.add_column("Repo Name",   justify="left",   style="white",       min_width=20, no_wrap=False)
    table.add_column("Repo URL",    justify="left",   overflow="fold",     min_width=36)
    table.add_column("Archived",    justify="center", style="white",       width=10,  no_wrap=True)
    table.add_column("Last Commit", justify="center", style="white",       width=16,  no_wrap=True)
    table.add_column("Stale Repo",  justify="center", style="white",       width=11,  no_wrap=True)

    for idx, row in enumerate(all_rows, start=1):
        is_stale    = row["Stale Repo"] == "Yes"
        is_archived = row["Archived"] == "True"

        # Repo URL — red if stale, cyan if active
        url_markup = (
            f"[bold red]{row['Repo URL']}[/]"
            if is_stale
            else f"[cyan]{row['Repo URL']}[/]"
        )

        archived_markup = (
            "[bold yellow]True[/]"
            if is_archived
            else "[dim white]False[/]"
        )

        commit_markup = commit_colour(row["_sort_key"], row["Last Commit"])

        stale_markup = (
            "[bold red]Yes[/]"
            if is_stale
            else "[green]No[/]"
        )

        table.add_row(
            str(idx),
            row["Org Name"],
            row["Repo Name"],
            url_markup,
            archived_markup,
            commit_markup,
            stale_markup,
        )

    console.print(table)
    console.print()


# ── Summary Table ─────────────────────────────────────────────────────────────

def print_summary(all_rows: list, orgs: List[str], elapsed: float) -> None:
    console.print(Rule(
        "[bold bright_green]  Summary  [/]",
        style="bright_green",
    ))
    console.print()

    # ── Per-org breakdown ──────────────────────────────────────────────
    per_org_tbl = Table(
        title="[bold white]Per-Organisation Breakdown[/]",
        box=box.ROUNDED,
        border_style="bright_green",
        header_style="bold bright_green",
        show_lines=True,
        expand=False,
    )
    per_org_tbl.add_column("Organisation",  style="bold cyan",     no_wrap=True,  min_width=20)
    per_org_tbl.add_column("Scanned",       justify="right",       style="white", width=9)
    per_org_tbl.add_column("Active (≤7d)",  justify="right",       width=12)
    per_org_tbl.add_column("Active (≤30d)", justify="right",       width=13)
    per_org_tbl.add_column("Stale (>180d)", justify="right",       width=13)
    per_org_tbl.add_column("No Commits",    justify="right",       width=11)
    per_org_tbl.add_column("Archived",      justify="right",       width=9)

    grand_scanned  = 0
    grand_fresh7   = 0
    grand_fresh30  = 0
    grand_stale    = 0
    grand_no_cmt   = 0
    grand_archived = 0

    for org in orgs:
        rows       = [r for r in all_rows if r["Org Name"] == org]
        if not rows:
            continue

        scanned    = len(rows)
        fresh7     = sum(1 for r in rows if r["_sort_key"] <= 7)
        fresh30    = sum(1 for r in rows if r["_sort_key"] <= 30)
        stale      = sum(1 for r in rows if r["Stale Repo"] == "Yes")
        no_cmt     = sum(1 for r in rows if r["_sort_key"] == 99_999)
        archived   = sum(1 for r in rows if r["Archived"] == "True")

        grand_scanned  += scanned
        grand_fresh7   += fresh7
        grand_fresh30  += fresh30
        grand_stale    += stale
        grand_no_cmt   += no_cmt
        grand_archived += archived

        per_org_tbl.add_row(
            org,
            str(scanned),
            f"[bold green]{fresh7}[/]"    if fresh7    else "[dim]0[/]",
            f"[bold yellow]{fresh30}[/]"  if fresh30   else "[dim]0[/]",
            f"[bold red]{stale}[/]"       if stale     else "[dim]0[/]",
            f"[dim]{no_cmt}[/]"           if no_cmt    else "[dim]0[/]",
            f"[bold yellow]{archived}[/]" if archived  else "[dim]0[/]",
        )

    # Totals row
    per_org_tbl.add_row(
        "[bold white]TOTAL[/]",
        f"[bold white]{grand_scanned}[/]",
        f"[bold green]{grand_fresh7}[/]"    if grand_fresh7   else "[dim]0[/]",
        f"[bold yellow]{grand_fresh30}[/]"  if grand_fresh30  else "[dim]0[/]",
        f"[bold red]{grand_stale}[/]"       if grand_stale    else "[dim]0[/]",
        f"[dim]{grand_no_cmt}[/]"           if grand_no_cmt   else "[dim]0[/]",
        f"[bold yellow]{grand_archived}[/]" if grand_archived else "[dim]0[/]",
    )

    console.print(per_org_tbl)
    console.print()

    # ── Stat cards row ─────────────────────────────────────────────────
    # Build a single-row table that acts as stat cards
    cards = Table(
        box=box.SIMPLE_HEAVY,
        border_style="bright_blue",
        show_header=False,
        expand=True,
        padding=(0, 2),
    )
    cards.add_column("A", justify="center")
    cards.add_column("B", justify="center")
    cards.add_column("C", justify="center")
    cards.add_column("D", justify="center")
    cards.add_column("E", justify="center")

    def _card(label: str, value: str, color: str) -> str:
        return f"[dim]{label}[/]\n[{color}]{value}[/{color}]"

    cards.add_row(
        _card("Orgs Scanned",   str(len(orgs)),         "bold white"),
        _card("Total Repos",    str(grand_scanned),     "bold white"),
        _card("Active (≤7d)",   str(grand_fresh7),      "bold green"),
        _card("Stale (>180d)",  str(grand_stale),       "bold red"),
        _card("Archived",       str(grand_archived),    "bold yellow"),
    )
    console.print(cards)
    console.print()

    # ── Final done panel ───────────────────────────────────────────────
    def _fmt(s: float) -> str:
        m, sec = divmod(int(s), 60)
        return f"{m}m {sec}s" if m else f"{sec:.0f}s"

    console.print(Panel(
        f"[bold green]✓  Analysis Complete[/]   [dim]·[/]   "
        f"[white]Orgs:[/] [bold]{len(orgs)}[/]   "
        f"[white]Repos:[/] [bold]{grand_scanned}[/]   "
        f"[white]Active (≤7d):[/] [bold green]{grand_fresh7}[/]   "
        f"[white]Stale:[/] [bold red]{grand_stale}[/]   "
        f"[white]Archived:[/] [bold yellow]{grand_archived}[/]   "
        f"[white]No Commits:[/] [dim]{grand_no_cmt}[/]   "
        f"[white]Runtime:[/] [bold cyan]{_fmt(elapsed)}[/]",
        border_style="bright_green",
        box=box.DOUBLE_EDGE,
        expand=False,
    ))
    console.print()


# ── Main ──────────────────────────────────────────────────────────────────────

def main() -> None:
    t_start = time.monotonic()

    print_banner()

    # ── Auth ──────────────────────────────────────────────────────────
    with console.status(
        "  [dim]Reading token via [bold]gh auth token[/]...[/]",
        spinner="dots",
    ):
        token = get_token()

    client = GHClient(token)
    me     = client._get("/user").json()

    console.print(
        f"  [green]✓[/]  Logged in as "
        f"[bold cyan]{me.get('login')}[/]"
        f"  [dim]({me.get('name', '')})[/]"
    )
    console.print()

    # ── Discover orgs ─────────────────────────────────────────────────
    with console.status(
        "  [dim]Fetching organisations...[/]", spinner="dots"
    ):
        orgs = get_orgs(client)

    if not orgs:
        console.print(Panel(
            "[bold yellow]⚠  No organisations found for this account.[/]",
            border_style="yellow", box=box.ROUNDED, expand=False,
        ))
        sys.exit(0)

    console.print(
        f"  [green]✓[/]  Found [bold]{len(orgs)}[/] organisation(s):  "
        f"[dim]{', '.join(orgs)}[/]"
    )
    console.print()

    # ── Scan repos ────────────────────────────────────────────────────
    all_rows: List[dict] = []
    progress = make_progress()

    with progress:
        for org in orgs:
            with console.status(
                f"  [dim]Listing repos for [bold]{org}[/]...[/]",
                spinner="dots",
            ):
                repos = list(
                    client.paginate(
                        f"/orgs/{org}/repos",
                        params={"type": "all"},
                    )
                )

            if not repos:
                continue

            task = progress.add_task(
                f"[cyan]{org}[/] — fetching commits",
                total=len(repos),
            )

            for repo in repos:
                repo_name      = repo.get("name", "")
                full_name      = repo.get("full_name", "")
                default_branch = repo.get("default_branch", "main")
                html_url       = repo.get("html_url") or f"https://github.com/{full_name}"
                archived       = repo.get("archived", False)

                progress.update(
                    task,
                    description=f"[cyan]{org}[/] — [white]{repo_name[:26]}[/]",
                )

                last_commit = get_last_commit_date(client, full_name, default_branch)
                if not last_commit:
                    last_commit = repo.get("pushed_at")

                days_str = days_since(last_commit)
                days_key = days_sort_key(days_str)
                stale    = "Yes" if days_key > STALE_DAYS else "No"

                all_rows.append({
                    "Org Name":   org,
                    "Repo Name":  repo_name,
                    "Repo URL":   html_url,
                    "Archived":   str(archived),
                    "Last Commit":days_str,
                    "Stale Repo": stale,
                    "_sort_key":  days_key,
                })

                progress.advance(task)

    # Sort by inactivity (most recent first)
    all_rows.sort(key=lambda x: x["_sort_key"])

    # ── Print results table ───────────────────────────────────────────
    print_results_table(all_rows)

    # ── Export to Excel ───────────────────────────────────────────────
    export_rows = [
        {
            "Sno":        idx,
            "Org Name":   r["Org Name"],
            "Repo Name":  r["Repo Name"],
            "Repo URL":   r["Repo URL"],
            "Archived":   r["Archived"],
            "Last Commit":r["Last Commit"],
            "Stale Repo": r["Stale Repo"],
        }
        for idx, r in enumerate(all_rows, start=1)
    ]

    with console.status(
        "[bright_blue]Writing Excel report...", spinner="dots"
    ):
        excel_path = export_excel(export_rows)

    console.print(
        f"  [green]✓[/]  Excel saved → [dim]{excel_path}[/]"
    )
    console.print()

    # ── Summary ───────────────────────────────────────────────────────
    elapsed = time.monotonic() - t_start
    print_summary(all_rows, orgs, elapsed)


if __name__ == "__main__":
    main()