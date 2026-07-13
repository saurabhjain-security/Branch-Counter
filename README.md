# 🌿 Branch Counter

> **Scan every repo. Count every branch. Spot the mess instantly.**

A Python CLI tool that counts branches across all repositories in one or more GitHub organisations — with a gorgeous terminal dashboard, colour-coded health ratings, and one-click exports to Excel, CSV, or JSON.

![Python](https://img.shields.io/badge/Python-3.9%2B-blue?style=flat-square&logo=python)
![License](https://img.shields.io/badge/License-MIT-green?style=flat-square)
![Version](https://img.shields.io/badge/Version-1.0-orange?style=flat-square)

---

## ✨ Features

| | Feature | What it does |
|---|---|---|
| 🎨 | **Rich terminal dashboard** | Colour-coded tables, live progress bars, per-org summary & stat cards |
| 🏥 | **Health classification** | Repos rated automatically: Healthy · Moderate · Crowded · Bloated |
| ⚡ | **Parallel fetching** | Concurrent repo listing + branch counting via thread pool |
| 📦 | **Multiple export formats** | Excel (multi-sheet), CSV, JSON |
| 🔍 | **Flexible filtering** | Show only bloated, crowded, healthy, or archived repos |
| 🔃 | **Sorting** | By branch count, name, org, or health tier |
| 🛡️ | **Rate-limit aware** | Auto-throttle, waits on reset, shows UTC reset time |
| 🛑 | **Graceful Ctrl+C** | Partial results are saved before exit |

---

## 🚀 Quick Start

```bash
# 1. Clone or download the script
git clone <your-repo-url>
cd branch-counter

# 2. Install dependencies
pip install -r requirements.txt

# 3. Set your GitHub token
export GITHUB_TOKEN=ghp_your_token_here

# 4. Run!
python branch_counter.py
```

---

## 📦 Installation

### Prerequisites

- **Python 3.9+** — [Download here](https://www.python.org/downloads/)
- **GitHub Personal Access Token** with `repo` and `read:org` scopes
  → [Create one here](https://github.com/settings/tokens)

### Step 1 — Install dependencies

Using `requirements.txt` (recommended):

```bash
pip install -r requirements.txt
```

Or install manually:

```bash
pip install rich requests openpyxl
```

> 💡 Using a virtual environment is recommended:
> ```bash
> python -m venv venv
> source venv/bin/activate      # macOS / Linux
> venv\Scripts\activate         # Windows
> pip install -r requirements.txt
> ```

### Step 2 — Set up authentication

The tool resolves a token in this order:

```
GITHUB_TOKEN  →  GH_TOKEN  →  gh auth token  (GitHub CLI fallback)
```

**Option A — Environment variable (recommended)**
```bash
# macOS / Linux
export GITHUB_TOKEN=ghp_your_token_here

# Windows (Command Prompt)
set GITHUB_TOKEN=ghp_your_token_here

# Windows (PowerShell)
$env:GITHUB_TOKEN="ghp_your_token_here"
```

**Option B — GitHub CLI**
```bash
gh auth login
```

### Step 3 — Run

```bash
python branch_counter.py
```

---

## 🎮 Usage

```bash
python branch_counter.py [OPTIONS]
```

### ⚙️ Options

| Option | Default | Description |
|--------|---------|-------------|
| `--org ORG [ORG ...]` | `carter-rmn shyftlabs` | Target one or more organisations |
| `--min-branches N` | `0` | Only show repos with at least N branches |
| `--filter MODE` | `all` | `all` · `healthy` · `moderate` · `crowded` · `bloated` · `archived` |
| `--sort BY` | `count` | `count` · `name` · `org` · `health` |
| `--output-format FMT` | `xlsx` | `xlsx` · `csv` · `json` |
| `--output PATH` | auto-generated | Custom output file path |
| `--no-export` | — | Terminal output only, skip file export |
| `--exclude-forks` | — | Skip forked repositories |
| `--exclude-archived` | — | Skip archived repositories |
| `--workers N` | `10` | Number of parallel worker threads |
| `--verbose / -v` | — | Show per-API-call logging and timings |
| `--help / -h` | — | Show help and exit |

### 💡 Examples

```bash
# Scan default orgs
python branch_counter.py

# Scan a specific org
python branch_counter.py --org acme-corp

# Show only bloated repos, export as CSV
python branch_counter.py --filter bloated --output-format csv

# Repos with more than 10 branches, sorted by name
python branch_counter.py --min-branches 10 --sort name

# Skip forks and archived, export to a custom path
python branch_counter.py --exclude-forks --exclude-archived --output report.xlsx

# Terminal only, no export file
python branch_counter.py --no-export

# Verbose mode — shows every API call, timing, and rate-limit state
python branch_counter.py --verbose
```

---

## 🏥 Health Classification

Repos are automatically classified based on branch count:

| Tier | Branch Count | Terminal Colour | What it means |
|------|-------------|-----------------|---------------|
| 🟢 **Healthy** | ≤ 15 | `Green` | Clean, well-maintained |
| 🟡 **Moderate** | 16 – 50 | `Yellow` | Getting busy, keep an eye on it |
| 🟠 **Crowded** | 51 – 70 | `Orange` | Needs pruning soon |
| 🔴 **Bloated** | > 70 | `Red` | Overdue for a cleanup |
| ⚪ **Archived** | — | `Grey` | Read-only, skipped in counts |

---

## 📤 Output

### 🖥️ Terminal

A full-width table with per-repo health, followed by a per-organisation breakdown and stat cards:

```
┌─────────────────────────────────────────────────────────────────────┐
│   GitHub Repository Branch Analysis   (42 repos)                    │
├──────┬──────────────┬──────────────────┬───────────┬────────┬───────┤
│  Sno │ Org / Owner  │ Repo Name        │ Branches  │ Health │  ...  │
├──────┼──────────────┼──────────────────┼───────────┼────────┼───────┤
│    1 │ acme-corp    │ legacy-monolith  │    87 🔴  │Bloated │  ...  │
│    2 │ acme-corp    │ api-gateway      │    22 🟠  │Crowded │  ...  │
│    3 │ acme-corp    │ frontend         │     4 🟢  │Healthy │  ...  │
└──────┴──────────────┴──────────────────┴───────────┴────────┴───────┘
```

### 📊 Excel (default)

- **"All Repos"** sheet with every repository
- One sheet per organisation
- Health column is colour-coded; headers include auto-filter and freeze panes

### 📄 CSV

Standard comma-separated with headers:

```
Sno, Org / Owner, Repo Name, Repo URL, Fork, Archived, Branch Count, Health
```

### 🗂️ JSON

Array of objects with snake_case keys:

```json
[
  {
    "sno": 1,
    "org_name": "acme-corp",
    "repo_name": "legacy-monolith",
    "repo_url": "https://github.com/acme-corp/legacy-monolith",
    "is_fork": false,
    "is_archived": false,
    "branch_count": 87,
    "health": "Bloated"
  }
]
```

---

## 🛠️ Troubleshooting

| Problem | Fix |
|---------|-----|
| `ModuleNotFoundError: No module named 'rich'` | Run `pip install -r requirements.txt` |
| `✗ Not authenticated` | Set `GITHUB_TOKEN` or run `gh auth login` |
| `403 permission denied` | Token is missing `repo` or `read:org` scope |
| Scan is slow | Increase `--workers` (e.g. `--workers 20`) |
| Results look incomplete | Check API rate limit; re-run after reset |

---

## 👤 Author

**Saurabh Jain** — v1.0