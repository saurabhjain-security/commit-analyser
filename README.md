<div align="center">

# 🔍 Commit Analyser

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

### **GitHub Organisation Repository Commit Analyser**
*Find stale code, archived repos, and commit activity — across every org, in one command.*

<br>

![Python](https://img.shields.io/badge/Python-3.9%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)
![GitHub CLI](https://img.shields.io/badge/GitHub_CLI-Required-181717?style=for-the-badge&logo=github&logoColor=white)
![Rich](https://img.shields.io/badge/Rich-Terminal_UI-7B2FBE?style=for-the-badge)
![Excel](https://img.shields.io/badge/Excel-Export-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![Version](https://img.shields.io/badge/Version-2.1-00C853?style=for-the-badge)
![Author](https://img.shields.io/badge/Author-Saurabh_Jain-FF6D00?style=for-the-badge)

<br>

</div>

---

## ✨ What is this?

> **Commit Analyser** is a single-command Python CLI tool that connects to GitHub via the `gh` CLI, auto-discovers every organisation you belong to, and scans **all their repositories** — public, private, and archived.
>
> For every repo it fetches the latest commit date, calculates exactly how many days ago it was, and flags it as **stale** if inactive for more than 180 days. Results are printed as a beautiful colour-coded Rich table and exported to a timestamped Excel file.
>
> No config files. No tokens to paste. No flags to remember. Just:

```bash
python commit_analyser.py
```

---

## 🚀 Capabilities

<table>
<tr>
<td width="50%">

### 🔎 Discovery
- 🏢 **Auto org detection** — finds every org you are in
- 📦 **Full repo coverage** — public + private + archived
- 🌿 **Default branch aware** — commits fetched per branch
- 🔁 **Fallback logic** — uses `pushed_at` for empty repos

</td>
<td width="50%">

### 📊 Analysis
- ⏰ **Stale detection** — flags repos inactive > 180 days
- 📁 **Archived detection** — identifies archived repos
- 📅 **Human-readable dates** — "3 days ago", "Today"
- ⚡ **Rate limit handling** — auto-sleeps and continues

</td>
</tr>
<tr>
<td width="50%">

### 🎨 Terminal UI
- 🎭 **ASCII art banner** with version and author
- 📊 **Live progress bar** per org with repo names
- 🌈 **Colour-coded results** by commit age
- 📋 **Per-org breakdown** table with totals row
- 🃏 **Stat cards** — orgs / repos / stale / archived
- ⏱️ **Runtime tracking** in the final panel

</td>
<td width="50%">

### 💾 Export
- 📄 **Single Excel sheet** — all orgs in one file
- 🔗 **Full repo URLs** — clickable links in Excel
- 🕐 **Timestamped filenames** — never overwrites
- 🗂️ **7 columns** — Sno, Org, Repo, URL, Archived, Commit, Stale

</td>
</tr>
</table>

---

## 📋 Basic Requirements

| | Requirement | Version | Notes |
|---|---|---|---|
| 🐍 | Python | 3.9+ | [python.org](https://python.org) |
| 🐙 | GitHub CLI (`gh`) | Any | Needs `gh auth login` |
| 🌐 | Internet access | — | Reaches `api.github.com` |
| 🏢 | GitHub Org membership | At least one | Personal-only accounts won't work |

---

## ⚙️ Installation

### 1️⃣ Install GitHub CLI

```bash
# 🍎 macOS
brew install gh

# 🐧 Ubuntu / Debian
sudo apt install gh

# 🪟 Windows
winget install --id GitHub.cli
```

```bash
gh --version   # verify it's working
```

---

### 2️⃣ Authenticate with GitHub

```bash
gh auth login
```

> 💡 **Select:** GitHub.com → HTTPS → Login with a web browser
> A browser window opens — log in and you are done.

```bash
gh auth status   # verify you're logged in
```

---

### 3️⃣ Get the script

```bash
# Option A — Clone
git clone https://github.com/your-username/commit-analyser.git
cd commit-analyser

# Option B — Download just the script
curl -O https://raw.githubusercontent.com/your-username/commit-analyser/main/commit_analyser.py
```

---

### 4️⃣ Install Python dependencies

```bash
pip install -r requirements.txt
```

---

## 📦 requirements.txt

```text
# ┌─────────────────────────────────────────────────┐
# │     Commit Analyser — Python Dependencies       │
# │     pip install -r requirements.txt             │
# └─────────────────────────────────────────────────┘

rich>=13.7.0       # 🎨 Terminal UI — tables, panels, progress, colours
requests>=2.31.0   # 🌐 GitHub REST API v3 calls
pandas>=2.0.0      # 📊 DataFrame building and Excel export
openpyxl>=3.1.0    # 📄 Excel .xlsx writer engine (used by pandas)
```

---

## ▶️ How to Run

```bash
python commit_analyser.py
```

> ✅ **That's the only command.** No arguments. No config. The script handles everything.

---

### 🔄 What happens step by step

```
┌─────────────────────────────────────────────────────────────┐
│                                                             │
│  1  🎨  ASCII art banner printed (v2.1 · Saurabh Jain)      │
│  2  🔑  Token read silently via gh auth token               │
│  3  ✅  GitHub identity verified and displayed              │
│  4  🏢  All organisations auto-discovered                   │
│                                                             │
│  5  🔁  For each organisation:                              │
│         📦  All repos fetched (public + private + archived) │
│         📅  Latest commit fetched per repo                  │
│         ⏰  Days since last commit calculated               │
│         🔴  Stale flag set if inactive > 180 days           │
│         📁  Archived flag read from GitHub API              │
│                                                             │
│  6  🌈  Colour-coded results table printed                  │
│  7  💾  Excel file written to current directory             │
│  8  📊  Per-org summary breakdown table printed             │
│  9  🃏  Stat cards + final panel with runtime               │
│                                                             │
└─────────────────────────────────────────────────────────────┘
```

---

### 🖥️ Terminal Output Preview

```
╔══════════════════════════════════════════════════════════════╗
║      COMMIT ANALYSER    v2.1   |   Author: Saurabh Jain     ║
╚══════════════════════════════════════════════════════════════╝

  ✓  Logged in as saurabh-jain  (Saurabh Jain)
  ✓  Found 2 organisation(s):  mycompany, mycompany-devops

  ⠋ mycompany — api-service  ████████████░░░░  12/23  00:38

──────────────────── Scan Results ─────────────────────────────

  ┌────┬──────────────────┬───────────────────────────────┬──────────┬───────────────┬────────────┐
  │  # │ Org Name         │ Repo URL                      │ Archived │ Last Commit   │ Stale Repo │
  ├────┼──────────────────┼───────────────────────────────┼──────────┼───────────────┼────────────┤
  │  1 │ mycompany        │ https://github.com/myco/api   │ False    │ Today         │ No         │
  │  2 │ mycompany        │ https://github.com/myco/web   │ False    │ 3 days ago    │ No         │
  │  3 │ mycompany-devops │ https://github.com/mcd/infra  │ False    │ 45 days ago   │ No         │
  │  4 │ mycompany        │ https://github.com/myco/old   │ True     │ 247 days ago  │ Yes        │
  └────┴──────────────────┴───────────────────────────────┴──────────┴───────────────┴────────────┘

  ✓  Excel saved → github_repo_scan_2025-05-12_10-30-00.xlsx

──────────────────── Summary ───────────────────────────────────

  ┌──────────────────┬─────────┬────────────┬─────────────┬─────────────┬────────────┬──────────┐
  │ Organisation     │ Scanned │ Active ≤7d │ Active ≤30d │ Stale >180d │ No Commits │ Archived │
  ├──────────────────┼─────────┼────────────┼─────────────┼─────────────┼────────────┼──────────┤
  │ mycompany        │      45 │         12 │          28 │           8 │          2 │        3 │
  │ mycompany-devops │      23 │          5 │          11 │           4 │          1 │        0 │
  │ TOTAL            │      68 │         17 │          39 │          12 │          3 │        3 │
  └──────────────────┴─────────┴────────────┴─────────────┴─────────────┴────────────┴──────────┘

   Orgs: 2   │   Total Repos: 68   │   Active ≤7d: 17   │   Stale: 12   │   Archived: 3

╔══════════════════════════════════════════════════════════════════════╗
║  ✓ Analysis Complete · Repos: 68 · Stale: 12 · Runtime: 4m 23s    ║
╚══════════════════════════════════════════════════════════════════════╝
```

---

## 🌈 Colour Guide

| Colour | Applies to | Meaning |
|---|---|---|
| 🟢 **Bold Green** | Last Commit, Stale = No | Committed today or ≤ 7 days — actively maintained |
| 🟡 **Bold Yellow** | Last Commit, Archived = True | Committed 8–30 days ago, or repo is archived |
| 🟠 **Bold Orange** | Last Commit | Committed 31–180 days ago — slowing down |
| 🔴 **Bold Red** | Last Commit, Repo URL, Stale = Yes | Last commit > 180 days — stale, URL turns red |
| 🔵 **Cyan** | Repo URL | Active repo — clickable link |
| ⬜ **Dim White** | No Commits, Archived = False | Empty repo or confirmed not archived |

---

## 📄 Excel Output

Every run creates a **timestamped Excel file** so old results are never overwritten:

```
github_repo_scan_2025-05-12_10-30-00.xlsx
```

**One sheet — `Repos` — with 7 columns:**

| # | Column | Example value |
|---|---|---|
| 1 | `Sno` | 1, 2, 3 … |
| 2 | `Org Name` | mycompany |
| 3 | `Repo Name` | api-service |
| 4 | `Repo URL` | https://github.com/mycompany/api-service |
| 5 | `Archived` | True / False |
| 6 | `Last Commit` | 3 days ago |
| 7 | `Stale Repo` | Yes / No |

---

## 🎛️ Configuration

The only thing you might want to tweak — the stale threshold:

```python
# commit_analyser.py  ·  line 62
STALE_DAYS = 180   # ← change this to suit your team
```

| `STALE_DAYS` | What it means |
|---|---|
| `30` | Flag repos with nothing merged in a month |
| `90` | Flag repos dormant for a quarter |
| `180` | ✅ Default — 6 months of inactivity |
| `365` | Only flag repos dead for a full year |

---

## 🎬 Record a GIF

> Show this off! Here's how to record a GIF of the tool running.

**Option A — asciinema + agg** *(lightweight, recommended)*

```bash
# Install
brew install asciinema
cargo install --git https://github.com/asciinema/agg

# Record
asciinema rec demo.cast
python commit_analyser.py
# Ctrl+D to stop

# Render to GIF
agg demo.cast demo.gif
```

**Option B — Terminalizer** *(all-in-one, needs Node.js)*

```bash
npm install -g terminalizer
terminalizer record demo
python commit_analyser.py
# Ctrl+D to stop
terminalizer render demo
```

Then drop it right into this README:

```markdown
![Commit Analyser Demo](demo.gif)
```

---

## 🛠️ Troubleshooting

<details>
<summary>❌ &nbsp;<strong>gh CLI not found</strong></summary>
<br>

Install the GitHub CLI from [cli.github.com](https://cli.github.com) then run `gh auth login`.

```bash
brew install gh      # macOS
sudo apt install gh  # Ubuntu
```

</details>

<details>
<summary>❌ &nbsp;<strong>Not logged into GitHub CLI</strong></summary>
<br>

```bash
gh auth login
```

Select: **GitHub.com → HTTPS → Login with a web browser**

</details>

<details>
<summary>⚠️ &nbsp;<strong>No organisations found</strong></summary>
<br>

Your account must be a **member of at least one GitHub organisation**.

Check here: [github.com/settings/organizations](https://github.com/settings/organizations)

</details>

<details>
<summary>⏳ &nbsp;<strong>Script pauses with "Rate limit hit"</strong></summary>
<br>

Completely normal for large orgs. GitHub allows **5,000 API requests/hour** for authenticated users. The script reads the reset time from the API header and waits automatically — no action needed.

</details>

<details>
<summary>❌ &nbsp;<strong>ModuleNotFoundError</strong></summary>
<br>

```bash
pip install -r requirements.txt

# If openpyxl is still missing:
pip install openpyxl
```

</details>

---

## 📁 Project Structure

```
commit-analyser/
│
├── 🐍  commit_analyser.py   ← Main script (single file, zero config)
├── 📦  requirements.txt     ← Python dependencies
└── 📖  README.md            ← You are here
```

---

## 📜 Changelog

| Version | What changed |
|---|---|
| **2.1** | 🎨 Rich UI overhaul — ASCII banner, stat cards, per-org summary table, runtime tracking |
| **2.0** | 📊 Excel export, stale detection, archived flag, per-org breakdown |
| **1.0** | 🚀 Initial release — basic repo listing |

---

<div align="center">

---

*Made with &nbsp;🐍 Python &nbsp;·&nbsp; 🎨 Rich &nbsp;·&nbsp; 🐙 GitHub REST API v3*

**Author: Saurabh Jain &nbsp;·&nbsp; v2.1**

---

</div>
