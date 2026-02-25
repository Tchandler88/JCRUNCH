# JCRUNCH
**JCR Content Repository Unifier and Node-to-Column Harvester**

A Python CLI tool that parses Adobe Experience Manager (AEM) Package Manager exports,
audits your content repository, and populates a structured Excel workbook for migration
planning — with an optional Excel ribbon UI for one-click operation.

---

## Table of Contents

1. [What JCRUNCH Does](#what-jcrunch-does)
2. [Prerequisites](#prerequisites)
3. [Installation](#installation)
4. [Folder Structure](#folder-structure)
5. [Excel Workbook Setup](#excel-workbook-setup)
6. [Running from the Command Line](#running-from-the-command-line)
7. [Running from Excel (Ribbon UI)](#running-from-excel-ribbon-ui)
   - [Step A — Import the VBA Module](#step-a--import-the-vba-module)
   - [Step B — Install the Ribbon XML](#step-b--install-the-ribbon-xml)
8. [Phase Reference](#phase-reference)
9. [AI Bot (Optional)](#ai-bot-optional)
10. [Troubleshooting](#troubleshooting)

---

## What JCRUNCH Does

JCRUNCH takes an AEM Package Manager `.zip` export and runs a five-phase audit pipeline:

| Phase | Sheet Name | What It Produces |
|-------|-----------|-----------------|
| 1 | Phase 1 — Taxonomy Audit | Every CQ tag with status, depth, hierarchy (L1–L4), usage count, and cloud migration recommendation |
| 2 | Phase 2 — Metadata Schema | Every JCR property with data type, namespace, usage frequency, and anomaly flags |
| 3 | Phase 3 — Workflow Extraction | Workflow step inventory (populated from package if present) |
| 4 | Phase 4 — Folder Redesign | Full folder tree with child/asset counts and metadata-pattern detection |
| 5 | Phase 5 — Namespace Validation | Every namespace URI with cloud support classification, migration strategy, effort, and timeline |

All processing is **in-memory** — no database required. Output is written directly into the
Excel workbook (`AEM_Migration_Analysis_Tool_v3.xlsx`).

---

## Prerequisites

| Requirement | Version | Notes |
|-------------|---------|-------|
| Python | 3.8 or higher | Must be on your system PATH |
| Microsoft Excel | 2016 or higher | Required for the ribbon UI |
| AEM Package Manager export | `.zip` format | Exported from AEM Package Manager |
| `AEM_Migration_Analysis_Tool_v3.xlsx` | Provided separately | Must be placed next to the `jcrunch/` folder |

> **Check your Python version:**
> ```bash
> python --version
> ```

---

## Installation

### 1. Clone or download the repository

```bash
git clone https://github.com/Tchandler88/JCRUNCH.git
```

Or download the ZIP from GitHub and extract it.

### 2. Place the Excel workbook next to the `jcrunch/` folder

Your folder layout must look like this:

```
(any parent folder)/
├── AEM_Migration_Analysis_Tool_v3.xlsx   ← your workbook goes here
└── jcrunch/
    ├── jcrunch.py
    ├── requirements.txt
    ├── parser/
    ├── audit/
    ├── export/
    ├── ai/
    ├── db/
    ├── config/
    ├── vba/
    └── tests/
```

> The VBA ribbon UI expects the `jcrunch/` folder to be a sibling of the workbook.
> If you move one, move the other.

### 3. Install Python dependencies

Open a terminal, navigate into the `jcrunch/` folder, and run:

```bash
cd jcrunch
pip install -r requirements.txt
```

This installs:

| Package | Purpose |
|---------|---------|
| `click` | CLI argument parsing |
| `openpyxl` | Reading and writing the Excel workbook |
| `python-dotenv` | Loading the `.env` file for the AI Bot |
| `anthropic` | Claude API client (AI Bot only — optional) |

### 4. (Optional) Configure the AI Bot

Only needed if you intend to use `--run-ai`.

```bash
cp .env.example .env
```

Open `.env` and set your Anthropic API key:

```
ANTHROPIC_API_KEY=sk-ant-your-key-here
```

> Without this key the core pipeline (parsing + auditing + writing) runs fine.
> Only the `--run-ai` flag requires it.

---

## Folder Structure

```
jcrunch/
├── jcrunch.py                  # CLI entry point — run this
├── requirements.txt            # pip dependencies
├── .env.example                # Template for Anthropic API key
├── README.md                   # This file
│
├── parser/
│   ├── package_reader.py       # Unzips the AEM package, walks every .content.xml
│   ├── xml_parser.py           # Parses a single .content.xml → structured dict
│   └── tag_resolver.py         # Tag hierarchy helpers (L1–L4, depth, parent)
│
├── audit/
│   ├── tag_auditor.py          # Phase 1 — enriches tags with status + cloud notes
│   ├── metadata_auditor.py     # Phase 2 — aggregates properties into field summary
│   ├── folder_auditor.py       # Phase 4 — enriches folders with counts + patterns
│   └── namespace_auditor.py    # Phase 5 — classifies namespaces + migration strategy
│
├── export/
│   └── workbook_writer.py      # Writes all 5 phase sheets into the Excel workbook
│
├── ai/
│   └── bot.py                  # AI Bot stub (future: fills AI columns via Claude)
│
├── config/
│   ├── namespace_map.json      # Known namespace URI → prefix mappings
│   ├── required_fields.json    # Required metadata fields for validation
│   └── ai_prompts.json         # Prompt templates for the AI Bot
│
├── vba/
│   ├── JCRUNCH_Ribbon.bas      # VBA module — ribbon button logic
│   └── JCRUNCH_RibbonUI.xml    # Custom ribbon XML — adds the JCRUNCH tab
│
├── tests/
│   └── test_parser.py          # Unit tests (in progress)
│
└── verify_workbook_writer.py   # Standalone sanity-check script for the export module
```

---

## Excel Workbook Setup

The workbook (`AEM_Migration_Analysis_Tool_v3.xlsx`) must have these five sheets
with exact names (em-dashes, not hyphens):

- `Phase 1 — Taxonomy Audit`
- `Phase 2 — Metadata Schema`
- `Phase 3 — Workflow Extraction`
- `Phase 4 — Folder Redesign`
- `Phase 5 — Namespace Validation`

Each sheet uses:
- **Rows 1–3** — Title, column headers, source labels (do not modify)
- **Row 4 onward** — Data written by JCRUNCH (cleared and rewritten on each run)

> JCRUNCH never touches AI Bot columns or manually entered columns.
> Only the data columns it owns are written.

---

## Running from the Command Line

Open a terminal, navigate into the `jcrunch/` folder, and run:

### Basic run (all phases)

```bash
cd jcrunch
python jcrunch.py --package "/path/to/your-package.zip" --workbook "/path/to/AEM_Migration_Analysis_Tool_v3.xlsx"
```

### Run a specific phase only

```bash
python jcrunch.py --package "package.zip" --workbook "workbook.xlsx" --phase 1
```

Valid values for `--phase`: `1`, `2`, `3`, `4`, `5`, or `all` (default).

### Run with the AI Bot

```bash
python jcrunch.py --package "package.zip" --workbook "workbook.xlsx" --run-ai
```

Requires `ANTHROPIC_API_KEY` set in your `.env` file.

### Run AI Bot only (skip parsing, use existing workbook data)

```bash
python jcrunch.py --workbook "workbook.xlsx" --ai-only
```

### All options

```
Options:
  --package PATH    AEM Package Manager .zip file
  --workbook PATH   Path to AEM_Migration_Analysis_Tool_v3.xlsx  [required]
  --run-ai          Run AI Bot fills after parsing
  --ai-only         Skip parsing, only run AI fills on existing workbook
  --phase TEXT      Run specific phase: 1, 2, 3, 4, 5, or all  [default: all]
  --help            Show this message and exit.
```

### Expected output

```
JCRUNCH -- It's GR-R-REAT for metadata audits
Reading package: your-package.zip
   Phase 1 tag audit complete
   Phase 2 metadata audit complete
   Phase 4 folder audit complete
   Phase 5 namespace audit complete
Writing to workbook: AEM_Migration_Analysis_Tool_v3.xlsx
   Workbook populated
JCRUNCH done. Open your workbook.
```

---

## Running from Excel (Ribbon UI)

The ribbon UI adds a **JCRUNCH** tab directly to Excel with two buttons:

- **Browse Package** — opens a file dialog to pick your `.zip`
- **Run JCRUNCH** — runs the full pipeline and refreshes the workbook

Setup requires two steps: importing the VBA module and installing the ribbon XML.

---

### Step A — Import the VBA Module

1. Open `AEM_Migration_Analysis_Tool_v3.xlsx` in Excel
2. Press `Alt + F11` to open the Visual Basic Editor
3. In the menu bar: **File → Import File...**
4. Browse to `jcrunch/vba/JCRUNCH_Ribbon.bas`
5. Click **Open**
6. Close the VBA Editor (`Alt + F4` or the X button)
7. Save the workbook as **Macro-Enabled** format: **File → Save As → Excel Macro-Enabled Workbook (`.xlsm`)**

> If you see a security warning about macros when reopening, click **Enable Content**.

---

### Step B — Install the Ribbon XML

Run the injector script — it handles everything automatically and is safe to re-run:

```bash
python jcrunch/vba/inject_ribbon.py --workbook "AEM_Migration_Analysis_Tool_v3.xlsx"
```

This produces a new file alongside the original:

```
AEM_Migration_Analysis_Tool_v3_ribbon.xlsm
```

1. **Open `AEM_Migration_Analysis_Tool_v3_ribbon.xlsm`** in Excel
2. Click **Enable Content** if prompted (allows macros)
3. The **JCRUNCH** tab will appear in the ribbon
4. If the VBA module is not yet imported, complete Step A on this new file

> The original workbook is never modified. The `_ribbon.xlsm` file is your working copy going forward.

> **What the script does:** reads the workbook as a zip, adds `customUI/customUI14.xml`,
> patches `_rels/.rels` and `[Content_Types].xml`, and writes a brand-new valid zip.
> No manual editing — no corruption risk.

---

### Using the Ribbon Buttons

Once the ribbon is installed:

1. Click **Browse Package** — select your AEM `.zip` file
2. Click **Run JCRUNCH** — the pipeline runs in a terminal window
3. When complete, a dialog confirms success and the workbook refreshes automatically

The selected package path is stored in a hidden sheet (`_Config`) and remembered
between sessions.

---

## Phase Reference

### Phase 1 — Taxonomy Audit

Audits every CQ tag found in the package.

**Tag Status values (priority order):**

| Status | Condition | Recommendation |
|--------|-----------|---------------|
| DEPRECATE - Missing Title | Tag has no title | Do not migrate |
| DEPRECATE - Bad Naming | Uppercase or space in tag ID leaf | Rename before migrating |
| DEPRECATE - Obsolete | Name contains: test, temp, mock, old, delete, backup, draft | Review and remove |
| CONSOLIDATE - Duplicate Title | Multiple tags share the same display title | Merge into one |
| REVIEW - Zero Usage | Tag is defined but used on zero assets | Evaluate relevance |
| REVIEW - High Usage | Used on more than 100 assets | Validate mapping carefully |
| REVIEW - Too Deep | Hierarchy depth exceeds 4 levels | Flatten before migrating |
| KEEP - Standard | Passes all checks | Migrate as-is |

**Hierarchy columns:** L1 through L4 (ID, title, description) are extracted automatically.

---

### Phase 2 — Metadata Schema

Aggregates all JCR properties into a unique field inventory.

**Data types detected automatically:**

| Type | Detection Rule |
|------|---------------|
| Boolean | Value is `true` or `false` |
| Long | Value contains only digits |
| Date | Value matches `YYYY-MM-DD` pattern |
| Path Reference | Value starts with `/content/` |
| String | Anything else |

**System-managed namespaces** (marked "Yes"): `jcr`, `oak`, `sling`, `granite`, `rep`, `nt`, `mix`, `vlt`, `cq`

---

### Phase 4 — Folder Redesign

Analyzes the DAM folder tree.

**Metadata-like folder names** (flagged for review):

| Category | Examples |
|----------|---------|
| Date/Year | `2024`, `2023`, `1` |
| Quarter | `q1`, `q2`, `q3`, `q4` |
| Orientation | `landscape`, `portrait`, `square` |
| State | `approved`, `archive`, `archived` |
| Region | `apac`, `emea`, `nam`, `latam`, `north`, `south` |
| Color | `color`, `colour` |

---

### Phase 5 — Namespace Validation

Classifies every XML namespace found in the package.

**Namespace types:**

| Type | URI Pattern | Cloud Support | Effort |
|------|------------|--------------|--------|
| System (Repo) | `jcr`, `oak`, `sling`, `granite`, `day.com/jcr` | Native Core (Restricted) | Low |
| Standard | `w3.org`, `purl.org`, `iptc`, `cipa`, `prism` | Native Supported | Low |
| Vendor | Adobe, Microsoft, Apple URIs | Requires CND | Medium |
| Custom | Everything else | Requires CND | High |

---

## AI Bot (Optional)

The AI Bot (`ai/bot.py`) is reserved for a future phase that will use the
Claude API to automatically fill AI-labeled columns in the workbook
(e.g., recommended display names, migration notes, consolidation suggestions).

To enable it:

1. Set `ANTHROPIC_API_KEY` in your `.env` file
2. Add `--run-ai` to your CLI command

The core pipeline works fully without the AI Bot.

---

## Troubleshooting

### "Python not found"
Make sure Python is on your system PATH:
```bash
python --version
```
If not found, reinstall Python and check **"Add Python to PATH"** during setup.

### "jcrunch.py not found" (from VBA)
The `jcrunch/` folder must be in the same directory as the workbook.
Check that your layout matches the [required structure](#2-place-the-excel-workbook-next-to-the-jcrunch-folder).

### "Package file not found"
The path to the `.zip` must be a full path with no special characters.
Avoid paths that contain brackets `[]` or other unusual characters.

### Workbook is not refreshing after a run
The VBA calls `ThisWorkbook.RefreshAll` on completion. If your workbook has
no data connections this is a no-op — simply close and reopen the workbook
to see the updated data.

### Macro security warning on open
Go to **File → Options → Trust Center → Trust Center Settings → Macro Settings**
and select **"Disable all macros with notification"**, then click **Enable Content**
when prompted.

### JCRUNCH tab does not appear after ribbon installation
- Confirm `customUI/customUI14.xml` exists inside the `.xlsm` (as a zip)
- Confirm the `rId99` relationship was added to `_rels/.rels`
- Ensure you saved the VBA module as a `.xlsm` (macro-enabled), not `.xlsx`

### Running the workbook writer verification script

To verify the export module is working independently:

```bash
cd jcrunch
python verify_workbook_writer.py
```

This creates a minimal harvest, writes it to the workbook, reads it back,
and prints PASS or FAIL for each of the 5 phases.

---

## License

Internal tool. Not for redistribution.
