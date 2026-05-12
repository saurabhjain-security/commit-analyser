# Commit Analyser

<div align="center">

```
  ██████╗ ██████╗ ███╗   ███╗███╗   ███╗██╗████████╗
 ██╔════╝██╔═══██╗████╗ ████║████╗ ████║██║╚══██╔══╝
 ██║     ██║   ██║██╔████╔██║██╔████╔██║██║   ██║   
 ██║     ██║   ██║██║╚██╔╝██║██║╚██╔╝██║██║   ██║   
 ╚██████╗╚██████╔╝██║ ╚═╝ ██║██║ ╚═╝ ██║██║   ██║   
  ╚═════╝ ╚═════╝ ╚═╝     ╚═╝╚═╝     ╚═╝╚═╝   ╚═╝   

 █████╗ ███╗   ██╗ █████╗ ██╗  ██╗   ██╗███████╗███████╗██████╗ 
██╔══██╗████╗  ██║██╔══██╗██║  ╚██╗ ██╔╝██╔════╝██╔════╝██╔══██╗
███████║██╔██╗ ██║███████║██║   ╚████╔╝ ███████╗█████╗  ██████╔╝
██╔══██║██║╚██╗██║██╔══██║██║    ╚██╔╝  ╚════██║██╔══╝  ██╔══██╗
██║  ██║██║ ╚████║██║  ██║███████╗██║   ███████║███████╗██║  ██║
╚═╝  ╚═╝╚═╝  ╚═══╝╚═╝  ╚═╝╚══════╝╚═╝   ╚══════╝╚══════╝╚═╝  ╚═╝
```

**GitHub Organisation Repository Commit Analyser**

![Python](https://img.shields.io/badge/Python-3.9%2B-blue?style=flat-square&logo=python)
![Rich](https://img.shields.io/badge/Rich-CLI-blueviolet?style=flat-square)
![GitHub CLI](https://img.shields.io/badge/GitHub_CLI-required-black?style=flat-square&logo=github)
![Version](https://img.shields.io/badge/Version-2.1-brightgreen?style=flat-square)
![Author](https://img.shields.io/badge/Author-Saurabh_Jain-orange?style=flat-square)

*Scan every repo across all your GitHub organisations — find stale code, archived projects, and commit activity at a glance.*

</div>

---

## What it does

Commit Analyser connects to GitHub via the `gh` CLI, discovers every organisation your account belongs to, and scans all their repositories. For each repo it fetches the most recent commit date, calculates how many days ago it was, and flags it as stale if no commit has happened in the last 180 days.

Results are displayed as a colour-coded Rich table in the terminal and exported to a timestamped Excel file — all in a single command.

---

## Capabilities

| Feature | Details |
|---|---|
| **Auto org discovery** | Finds all GitHub organisations your account belongs to — no config needed |
| **Full repo coverage** | Scans public repos, private repos, and archived repos in one pass |
| **Accurate commit dates** | Fetches the latest commit from the default branch via the Commits API, with `pushed_at` as fallback for empty repos |
| **Stale detection** | Flags any repo with no commit in the last 180 days as stale |
| **Archived detection** | Identifies archived repos and marks them clearly |
| **Colour-coded output** | Commit age is colour graded: green (≤7d) → yellow (≤30d) → orange (≤180d) → red (>180d). Stale repo URLs turn red. Archived repos turn yellow |
| **Rate limit handling** | Automatically detects GitHub API rate limits and sleeps until the reset window, then continues without intervention |
| **Per-org breakdown** | Summary table showing scanned / active / stale / archived counts per organisation, with a TOTAL row |
| **Stat cards** | Quick-glance stat panel: Orgs Scanned, Total Repos, Active ≤7d, Stale, Archived |
| **Runtime tracking** | Total scan duration shown in the final summary panel |
| **Excel export** | Single-sheet `.xlsx` file with all columns including full repo URLs |
| **Rich terminal UI** | ASCII art banner, animated progress bar, rule separators, styled panels — all via the Rich library |

---

## Output columns

| Column | Description |
|---|---|
| `Sno` | Row number |
| `Org Name` | GitHub organisation the repo belongs to |
| `Repo Name` | Repository name |
| `Repo URL` | Full `https://github.com/org/repo` link — cyan if active, red if stale |
| `Archived` | `True` (yellow) or `False` (dim) |
| `Last Commit` | Human-readable days ago — colour graded by age |
| `Stale Repo` | `Yes` (red) if last commit > 180 days ago, `No` (green) otherwise |

---

## Basic requirements

- **Python** 3.9 or later
- **GitHub CLI** (`gh`) installed and authenticated
- A GitHub account that belongs to at least one organisation
- Internet access to reach `api.github.com`

---

## Installation

### Step 1 — Install GitHub CLI

```bash
# macOS
brew install gh

# Ubuntu / Debian
sudo apt install gh

# Windows (winget)
winget install --id GitHub.cli
```

Verify:

```bash
gh --version
```

---

### Step 2 — Authenticate with GitHub

Run once and follow the prompts:

```bash
gh auth login
```

Select **GitHub.com → HTTPS → Login with a web browser**.

Verify you are logged in:

```bash
gh auth status
```

---

### Step 3 — Get the script

```bash
# Clone
git clone https://github.com/your-username/commit-analyser.git
cd commit-analyser

# Or download the single file
curl -O https://raw.githubusercontent.com/your-username/commit-analyser/main/commit_analyser.py
```

---

### Step 4 — Install Python dependencies

```bash
pip install -r requirements.txt
```

---

## requirements.txt

```text
# Commit Analyser — Python dependencies
# pip install -r requirements.txt

rich>=13.7.0        # Terminal UI — tables, panels, progress bars, colours
requests>=2.31.0    # GitHub REST API calls
pandas>=2.0.0       # DataFrame construction and Excel export
openpyxl>=3.1.0     # Excel .xlsx writer engine (used by pandas)
```

---

## How to run

```bash
python commit_analyser.py
```

That is the only command. The script discovers organisations automatically from your `gh` session — no flags, no config file, no token to paste.

---

### What happens step by step

```
1.  Banner is printed with version and author
2.  Token is read silently from `gh auth token`
3.  Your GitHub identity is verified and displayed
4.  All organisations are auto-discovered
5.  For each org:
      → All repos fetched (public + private + archived)
      → Latest commit date fetched per repo
      → Days since last commit calculated
      → Stale flag set if > 180 days with no commit
      → Archived flag read from API response
6.  Results table printed — colour-coded by commit age
7.  Excel file written to current directory
8.  Per-org summary breakdown table printed
9.  Stat cards + final done panel with total runtime
```

---

### Example terminal output (abbreviated)

```
╔══════════════════════════════════════════════════════════╗
║         COMMIT ANALYSER    v2.1  |  Saurabh Jain        ║
╚══════════════════════════════════════════════════════════╝

  ✓  Logged in as saurabh-jain  (Saurabh Jain)
  ✓  Found 2 organisation(s):  mycompany, mycompany-devops

─────────────────── Scan Results ──────────────────────────

  GitHub Repository Commit Analysis  (68 repos)
 ┌────┬──────────────────┬──────────────────────┬──────────────┬──────────┬──────────────┬────────────┐
 │ #  │ Org Name         │ Repo URL             │ Archived     │ Last Commit   │ Stale      │
 ├────┼──────────────────┼──────────────────────┼──────────────┼───────────────┼────────────┤
 │  1 │ mycompany        │ https://github.com/… │ False        │ Today         │ No         │
 │  2 │ mycompany        │ https://github.com/… │ False        │ 3 days ago    │ No         │
 │  3 │ mycompany-devops │ https://github.com/… │ False        │ 45 days ago   │ No         │
 │  4 │ mycompany        │ https://github.com/… │ True         │ 247 days ago  │ Yes        │
 └────┴──────────────────┴──────────────────────┴──────────────┴───────────────┴────────────┘

  ✓  Excel saved → /Users/you/github_repo_scan_2025-05-12_10-30-00.xlsx

─────────────────── Summary ────────────────────────────────

  Per-Organisation Breakdown
 ┌──────────────────┬─────────┬────────────┬─────────────┬─────────────┬────────────┬──────────┐
 │ Organisation     │ Scanned │ Active ≤7d │ Active ≤30d │ Stale >180d │ No Commits │ Archived │
 ├──────────────────┼─────────┼────────────┼─────────────┼─────────────┼────────────┼──────────┤
 │ mycompany        │      45 │         12 │          28 │           8 │          2 │        3 │
 │ mycompany-devops │      23 │          5 │          11 │           4 │          1 │        0 │
 │ TOTAL            │      68 │         17 │          39 │          12 │          3 │        3 │
 └──────────────────┴─────────┴────────────┴─────────────┴─────────────┴────────────┴──────────┘

  Orgs Scanned    Total Repos    Active ≤7d    Stale >180d    Archived
       2               68            17              12            3

╔══════════════════════════════════════════════════════════════════╗
║ ✓ Analysis Complete  · Orgs: 2  Repos: 68  Stale: 12  Time: 4m ║
╚══════════════════════════════════════════════════════════════════╝
```

---

## Colour reference

| Colour | Meaning |
|---|---|
| 🟢 Bold Green | Committed today or within the last 7 days |
| 🟡 Bold Yellow | Committed 8–30 days ago / Archived repo |
| 🟠 Bold Orange | Committed 31–180 days ago |
| 🔴 Bold Red | Last commit > 180 days ago (stale) — row URL also turns red |
| ⬜ Dim White | No commits found / repo is not archived / not stale |

---

## Excel output

The script writes a file named:

```
github_repo_scan_YYYY-MM-DD_HH-MM-SS.xlsx
```

Saved in the current working directory. One sheet named **Repos**:

| Sno | Org Name | Repo Name | Repo URL | Archived | Last Commit | Stale Repo |
|---|---|---|---|---|---|---|
| 1 | mycompany | api-service | https://github.com/mycompany/api-service | False | 3 days ago | No |
| 2 | mycompany | legacy-tool | https://github.com/mycompany/legacy-tool | False | 247 days ago | Yes |
| 3 | mycompany | old-archive | https://github.com/mycompany/old-archive | True | 512 days ago | Yes |

---

## Adjusting the stale threshold

The 180-day threshold is defined at the top of the script:

```python
STALE_DAYS = 180
```

Change it to any number. For example, to flag repos inactive for 90 days:

```python
STALE_DAYS = 90
```

---

## Recording a GIF demo

Install [asciinema](https://asciinema.org) (free, no account needed to record locally):

```bash
# macOS
brew install asciinema

# Ubuntu
sudo apt install asciinema
```

Record a session:

```bash
asciinema rec demo.cast
python commit_analyser.py
# Press Ctrl+D when done
```

Convert to GIF using [agg](https://github.com/asciinema/agg):

```bash
# Install agg
cargo install --git https://github.com/asciinema/agg

# Render
agg demo.cast demo.gif
```

Then add to this README:

```markdown
![Commit Analyser Demo](demo.gif)
```

Alternatively use [Terminalizer](https://github.com/faressoft/terminalizer) which records and renders to GIF directly:

```bash
npm install -g terminalizer
terminalizer record demo
python commit_analyser.py
# Ctrl+D to stop
terminalizer render demo
```

---

## Troubleshooting

**`gh CLI not found`**
Install from [cli.github.com](https://cli.github.com) and run `gh auth login`.

**`Not logged into GitHub CLI`**
Run `gh auth login` and follow the prompts.

**`No organisations found`**
Your account must be a member of at least one GitHub organisation. Personal-only accounts will return no orgs.

**Rate limit pauses**
Normal behaviour for large orgs. The script reads the `X-RateLimit-Reset` header and waits automatically. Authenticated requests allow 5,000 calls/hour.

**`ModuleNotFoundError`**
Run `pip install -r requirements.txt`.

**`openpyxl` not found**
Run `pip install openpyxl` separately — sometimes pip skips optional engine installs.

---

## Project structure

```
commit-analyser/
├── commit_analyser.py   ← Main script (single file, no package needed)
├── requirements.txt     ← Python dependencies
└── README.md            ← This file
```

---

## Author

**Saurabh Jain** — Version 2.1

---

<div align="center">
<sub>Built with Python · Rich · GitHub REST API v3</sub>
</div>
