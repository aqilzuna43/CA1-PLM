# CA1-PLM: Lightweight Product Lifecycle Management System

> A Google Apps Script-based PLM tool for managing Bill of Materials (BOM) and Engineering Change Requests (ECR) at Celestica CA1

[![License: Internal](https://img.shields.io/badge/License-Internal-blue.svg)]()
[![Platform: Google Apps Script](https://img.shields.io/badge/Platform-Google%20Apps%20Script-green.svg)]()
[![Runtime: V8](https://img.shields.io/badge/Runtime-V8-yellow.svg)]()

## Overview

CA1-PLM implements a **hub-and-spoke architecture** where a centralised Google Sheet serves as the Source of Truth for all product data, with satellite ECR forms acting as controlled transaction clients. This eliminates direct manual editing of the Master BOM while maintaining data integrity and change traceability.

### Key Design Principles
- **Single Source of Truth** — All product data centralised in one master spreadsheet
- **Controlled Modifications** — Changes only via validated ECR workflow
- **Hierarchical BOM Structure** — Level-based parent/child relationships
- **Modular Codebase** — 13 focused modules organised by domain
- **Real-Time Validation** — onEdit/onChange triggers guard data integrity as you type
- **Batch-First Performance** — Single API calls replace row-by-row operations

---

## System Architecture

### Hub-and-Spoke Model
```
┌─────────────────────────────────────────────────────────┐
│          12016271_CA1 Mini BoM (THE HUB)                │
│                  Source of Truth                          │
├─────────────────────────────────────────────────────────┤
│  ┌─────────┐    ┌─────────┐    ┌──────────┐            │
│  │  ITEMS  │    │   AML   │    │  MASTER  │            │
│  │   Tab   │    │   Tab   │    │   Tab    │            │
│  └─────────┘    └─────────┘    └──────────┘            │
│       │              │                │                  │
│   Part Data    Manufacturer      BOM Structure          │
│  Dictionary        List          (Hierarchical)         │
└─────────────────────────────────────────────────────────┘
                         ▲
                         │  Read/Write via API
                         │
┌────────────────────────┴─────────────────────────────────┐
│                   ECR Form (THE SPOKE)                    │
│                  Transaction Client                       │
├───────────────────────────────────────────────────────────┤
│  • Validates change requests against Master BOM           │
│  • Prevents invalid Parent/Child relationships            │
│  • Executes controlled modifications via commitToMaster() │
│  • Change Types: ADD | REMOVE | QTY CHANGE | REV ROLL    │
│                  DESC CHANGE | AVL ADD/REMOVE/REPLACE     │
└───────────────────────────────────────────────────────────┘
```

---

## Data Model

### The Hub: `12016271_CA1 Mini BoM`

#### 1. ITEMS Tab (Part Dictionary)
Defines **what** each part is.

| Column | Description |
|--------|-------------|
| Item Number | Unique identifier |
| Part Description | Part description |
| Item Rev | Current revision level |
| Lifecycle | Status (Active, Obsolete, EOL, NRND, etc.) |

#### 2. AML Tab (Approved Manufacturer List)
Defines **where** parts can be sourced.

| Column | Description |
|--------|-------------|
| Item Number | Reference to ITEMS |
| Mfr. Name | Approved manufacturer |
| Mfr. Part Number | Manufacturer's part number |

#### 3. MASTER Tab (Bill of Materials)
Defines **how** parts are assembled — uses hierarchical level structure.

| Column | Description | Used By |
|--------|-------------|---------|
| **Level** | Hierarchy depth (1, 2, 3 or 1.1, 1.1.1) | Parent/Child detection |
| Item Number | Reference to ITEMS | Part identification |
| Part Description | Auto-managed from ITEMS | Display |
| Item Rev | Auto-managed from ITEMS | Revision tracking |
| Qty | Assembly quantity | BOM calculations |
| Lifecycle | Auto-managed from ITEMS | Audit tools |
| Mfr. Name | Auto-managed from AML | Sourcing |
| Mfr. Part Number | Auto-managed from AML | Sourcing |
| Reference Notes | MAKE / BUY / REF status | Fabricator, Audit tools |

**Hierarchy Example:**
```
Level    Item Number    Description
1        ASSY-001       Main Assembly
  2      PCB-100        PCB Board            <- Child of ASSY-001
    3    R-001          Resistor             <- Child of PCB-100
    3    C-001          Capacitor            <- Child of PCB-100
  2      CABLE-200      Cable Assembly       <- Child of ASSY-001
1        ASSY-002       Second Assembly      <- New top-level
```

---

## BOM Tools Menu

The Google Sheets custom menu organises 18 tools into categorised submenus:

```
BOM Tools
├── Comparison
│   ├── Generate Detailed Comparison (ECO)
│   └── Compare Master vs. PDM BOM
├── PDM Integration
│   └── Import Children from PDM (Graft)
├── Fabrication
│   └── Generate Fabricator BOMs
├── Data Integrity
│   ├── Reconcile Master Data (Full Sync)
│   ├── Validate BOM (9-Check Audit)
│   ├── Protect Master Sheets
│   └── Install Change Watchdog
├── Audit & Quality
│   ├── Audit BOM Lifecycle Status
│   ├── Audit BOM Structural Integrity
│   └── List 'BUY' Items with 'REF' Children
├── Analysis & Reports
│   ├── Where-Used Analysis (Full Chain)
│   ├── Generate BOM Dashboard
│   └── Generate Master Lists from BOM
├── ─────────────
├── Prepare Rows for AML
├── Set BOM Effectivity Dates
└── Finalize and Release New BOM
```

### Tool Reference

#### Comparison
| Tool | Function | Description |
|------|----------|-------------|
| Detailed Comparison | `runDetailedComparison()` | Compares two BOM sheet versions, links changes to ECR/ECO numbers, generates colour-coded report with change-impact markers |
| PDM Comparison | `runCompareWithExternalBOM()` | Compares a Master BOM subassembly against an external PDM export with different headers |

#### PDM Integration
| Tool | Function | Description |
|------|----------|-------------|
| Import Children (Graft) | `runImportPdmChildren()` | Imports child components from a PDM export sheet, auto-adjusting hierarchy levels relative to a selected parent |

#### Fabrication
| Tool | Function | Description |
|------|----------|-------------|
| Fabricator BOMs | `runGenerateFabricatorBOMs()` | Generates flat BOM sheets for manufacturing — filters to BUY/REF items only, one sheet per assembly |

#### Data Integrity
| Tool | Function | Description |
|------|----------|-------------|
| Reconcile Master Data | `runReconcileMasterData()` | Batch sync — overwrites all managed columns from ITEMS/AML, replaces VLOOKUP formulas with plain values, flags orphans and missing AML |
| Validate BOM (9-Check) | `runValidateBOM()` | Comprehensive audit: orphans, missing AML, level gaps, structural mismatches, stale values, BUY/REF conflicts, lifecycle risk, circular deps, blank PNs |
| Protect Master Sheets | `runProtectMasterSheets()` | Applies warning-level protection on ITEMS/AML sheets and data-validation triangles on managed MASTER columns |
| Install Change Watchdog | `installChangeTrigger_()` | Installs the onChange trigger for automatic row-insert population and row-delete gap detection |

#### Audit & Quality
| Tool | Function | Description |
|------|----------|-------------|
| Lifecycle Audit | `runAuditBOMLifecycle()` | Flags components with OBSOLETE, EOL, NRND, or NOT RECOMMENDED status |
| Structural Integrity | `runAuditDuplicatePartNumbers()` | Detects assemblies reused in multiple locations that have inconsistent child structures |
| BUY Item Screen | `runScreenBuyItems()` | Finds BUY items that incorrectly contain REF children |

#### Analysis & Reports
| Tool | Function | Description |
|------|----------|-------------|
| Where-Used | `runWhereUsedAnalysis()` | Traces full ancestor chain for a part (Part > Parent > ... > Top Level) |
| Dashboard | `runGenerateDashboard()` | Creates summary sheet with metrics: unique parts, depth, lifecycle/status distribution, AML coverage |
| Master Lists | `runGenerateMasterLists()` | Extracts deduplicated ITEMS and AML lists from a BOM sheet |

#### Utilities
| Tool | Function | Description |
|------|----------|-------------|
| Prepare AML Rows | `runPrepareAMLRows()` | Inserts blank rows in a BOM sheet to accommodate multiple AML entries per part |
| Effectivity Dates | `runSetEffectivityDates()` | Adds/stamps "Effective From" and "Effective Until" columns on a BOM sheet |
| Release BOM | `runReleaseNewBOM()` | Copies WIP sheet, deletes REF rows, removes change-tracking columns, protects the sheet |

---

## ECR Workflow (Spoke)

The ECR Form (`ECR_FORM.gs`) runs in a **separate** Google Sheets spreadsheet linked to the Master Hub.

### Core Functions

| Function | Purpose |
|----------|---------|
| `populateCurrentData()` | Auto-fills "Curr *" columns from Master, preserves user edits in "New *" columns |
| `submitComprehensiveECR()` | Pushes ECR changes to the Master's `ECR_Affected_Items` log |
| `commitToMaster()` | **Admin-only** — Executes approved changes on MASTER / AML / ITEMS sheets |

### Supported Change Types

| Type | Target Sheet | Action |
|------|-------------|--------|
| ADDED | MASTER BOM | Insert new component (VLOOKUP formulas replaced by Reconcile sync) |
| REMOVED | MASTER BOM | Delete component and cascade children |
| QTY CHANGE | MASTER BOM | Update assembly quantity |
| REV ROLL | ITEMS | Update part revision |
| DESC CHANGE | ITEMS | Update description |
| MODIFIED | Multiple | Combined qty + rev + desc changes |
| AVL_ADD | AML | Add new manufacturer entry |
| AVL_REMOVE | AML | Delete manufacturer entry |
| AVL_REPLACE | AML | Update existing manufacturer |

---

## Project Structure

```
CA1-PLM/
├── appsscript.json               # GAS manifest (V8 runtime, OAuth scopes)
├── .clasp.json                   # Clasp project link
├── .claspignore                  # Deploy whitelist (src/bom + src/utils)
├── package.json                  # Node.js dependencies (clasp)
│
├── src/
│   ├── bom/                      # ── BOM Tool Modules (pushed to GAS) ──
│   │   ├── Config.gs             #   Constants, column mappings, validation config
│   │   ├── Menu.gs               #   onOpen() — custom menu with submenus
│   │   ├── Validation.gs         #   Real-time onEdit/onChange triggers & handlers
│   │   ├── Reconcile.gs          #   Batch data sync & sheet protection
│   │   ├── Comparison.gs         #   Detailed & PDM comparison tools
│   │   ├── BomMap.gs             #   Core BOM hierarchy parser
│   │   ├── PdmImport.gs          #   PDM children import (grafting)
│   │   ├── Fabricator.gs         #   Fabricator BOM generation
│   │   ├── Audit.gs              #   Lifecycle, structural, BUY-item, 9-check audits
│   │   ├── Analysis.gs           #   Where-Used, Dashboard, Master Lists
│   │   ├── Release.gs            #   Finalize/Release + Effectivity Dates
│   │   ├── History.gs            #   ECO & Revision change logging
│   │   └── Helpers.gs            #   Shared utilities (prompts, indexing)
│   │
│   ├── utils/                    # ── Shared Utility Modules ──
│   │   ├── CacheService.gs       #   Script-level caching (JSON, TTL)
│   │   ├── Constants.gs          #   UTIL_CONFIG (cache keys, batch size)
│   │   └── SheetService.gs       #   Sheet API wrapper with batching
│   │
│   └── legacy/                   # ── Archived / Separate Projects ──
│       ├── CA1_MINI_BOM.gs       #   (archived — replaced by src/bom/)
│       └── ECR_FORM.gs           #   ECR automation (separate GAS project)
│
├── scripts/                      # Deployment & migration utilities
│   ├── backup.sh
│   ├── deploy.sh
│   └── migrate.sh
│
└── README.md
```

### Module Dependency Map
```
Config.gs ← (no deps — loaded first)
Helpers.gs ← Config.gs (uses COL, BOM_CONFIG)
BomMap.gs ← Config.gs, Helpers.gs (uses getColumnIndexes, calculatePdmLevel)
    ↑
    Used by: Comparison.gs, Fabricator.gs, Audit.gs

Validation.gs ← Config.gs, Helpers.gs (onEdit/onChange triggers, lookup builders)
    ↑
    Exports: buildItemsLookup_(), buildAmlLookup_() — used by Reconcile.gs, Audit.gs

Reconcile.gs ← Config.gs, Helpers.gs, Validation.gs (batch sync, sheet protection)
History.gs ← Config.gs (uses BOM_CONFIG, REV_HISTORY_SHEET_NAME)
    ↑
    Used by: Comparison.gs (logs revision changes & ECO comparisons)

Menu.gs ← References all run*() entry-point functions by name
```

All `.gs` files share a single global scope in Google Apps Script — no imports required.

---

## Getting Started

### Prerequisites
- Google Account with access to the Master BOM sheet
- Node.js & npm (for local development with clasp)
- Git

### Installation

#### Option 1: Direct Google Apps Script
1. Open `12016271_CA1 Mini BoM` in Google Sheets
2. Go to **Extensions > Apps Script**
3. Create files matching the `src/bom/` structure in the script editor
4. Save and authorise permissions

#### Option 2: Local Development with Clasp (Recommended)
```bash
# Clone the repository
git clone https://github.com/aqilzuna43/CA1-PLM.git
cd CA1-PLM

# Install dependencies
npm install

# Login to Google
npx clasp login

# Push code to Apps Script
npx clasp push

# Open in browser to test
npx clasp open
```

### Finding Your Script ID
1. Open your Google Sheet
2. **Extensions > Apps Script**
3. Click **Project Settings** (gear icon)
4. Copy the **Script ID** into `.clasp.json`

---

## Development Workflow

### Local Development Cycle
```bash
# Pull latest from Google Apps Script
npx clasp pull

# Edit files in src/bom/ and src/utils/
# Use your preferred IDE (VS Code, etc.)

# Push changes to Google Apps Script
npx clasp push

# Test in Google Sheets — reload the sheet to trigger onOpen()

# Commit to Git
git add .
git commit -m "feat: add validation for negative quantities"
git push origin main
```

### Clasp Deployment Notes
- `.claspignore` uses a **deny-all + whitelist** pattern
- Only `src/bom/*.gs`, `src/utils/*.gs`, and `appsscript.json` are pushed
- `ECR_FORM.gs` belongs to a **separate** GAS project and is excluded
- `rootDir` in `.clasp.json` is `.` (project root)

### Testing Strategy
1. **Unit Testing** — Test individual functions in the Apps Script debugger
2. **Integration Testing** — Use a copy of the Master BOM sheet
3. **UAT** — Engineers test ECR workflow on non-critical assemblies
4. **Production** — Deploy to live `12016271_CA1 Mini BoM`

---

## Key Algorithms

### BOM Hierarchy Parsing (`buildBOMMap`)
The core parser converts flat sheet rows into a keyed Map structure:

```
Sheet Row:  Level=2, PN=PCB-100, under ASSY-001
    ↓
Map Key:    "ASSY-001/PCB-100"
Map Value:  { startRow, endRow, parent, mainRow{...}, aml[{...}] }
```

- Supports both integer levels (1, 2, 3) and dot-notation (1.1, 1.1.1)
- Tracks AML continuation rows (blank Item Number, populated Mfr. fields)
- Used by: Comparison, PDM Comparison, Fabricator, Audit tools

### Change Impact Propagation
During BOM comparison, changes are tracked bidirectionally:
- **Direct changes** (added/modified/removed items) are marked with `●`
- **Parent impact** (assemblies containing changed children) are marked with `▼`
- Propagation walks the location-key path upward (`A/B/C` flags `A/B`, then `A`)

### Batch Operations
All tools use batch reads (`getDataRange().getValues()`) and batch writes (`setValues()`) instead of row-by-row API calls, providing 4-15x performance improvements on large BOMs.

---

## Data Integrity & Safety

### 3-Layer Validation Architecture

| Layer | Trigger | Scope | Latency |
|-------|---------|-------|---------|
| **L1 — onEdit** | Simple trigger (automatic) | Cell-level: auto-populate on PN change, restore overwritten managed cols, validate Level gaps and Qty | Instant (<1s) |
| **L2 — onChange** | Installable trigger (via menu) | Row-level: populate managed cols on row insert, detect hierarchy gaps on row delete | Instant (<2s) |
| **L3 — Batch** | Manual (via menu) | Sheet-level: full reconcile of all managed columns, 9-check comprehensive audit, sheet protection | On-demand |

**Managed Columns** — Description, Item Rev, Lifecycle (from ITEMS) and Mfr. Name, Mfr. Part Number (from AML) are auto-populated by script using plain values instead of VLOOKUP formulas. This eliminates formula breakage from row insertions, column shifts, or accidental overwrites.

**Visual Feedback** — Errors and warnings are communicated via cell background colours and cell notes (prefixed `[BOM Validation]`), never via popup dialogs during real-time editing.

| Colour | Meaning |
|--------|---------|
| Red (`#f4cccc`) | Critical error — orphan PN, circular dependency, level gap |
| Yellow (`#fff2cc`) | Warning — missing AML, stale value |
| Orange (`#fce5cd`) | Stale — managed value overwritten by user |
| Green (`#d9ead3`) | Restored — value auto-corrected from source |

### Additional Protection Mechanisms
1. **ECR Validation** — `populateCurrentData()` prevents invalid ECR relationships
2. **Transaction Model** — All changes via ECR (full audit trail)
3. **Pre-Release Audit** — `runReleaseNewBOM()` checks for OBSOLETE/EOL components before release
4. **Sheet Protection** — Released BOM sheets locked with `sheet.protect()`; ITEMS/AML sheets get warning-level protection via `runProtectMasterSheets()`
5. **Admin Password** — `commitToMaster()` requires password verification

### Automatic Logging
| Log Sheet | Populated By | Tracks |
|-----------|-------------|--------|
| `Rev_History` | `logRevisionChange()` | Per-item revision changes with timestamp and source |
| `ECO History` | `logECOComparison()` | ECO comparisons with counts and related ECRs |
| `ECR_Affected_Items` | `submitComprehensiveECR()` | All ECR submissions with change details |

### Backup Strategy
- **Version History** — Enable on Google Sheets (File > Version History)
- **Weekly Export** — Export Master BOM to backup location
- **Test First** — Always test ECR changes on a copy

---

## Contributing

### Commit Message Convention
```
feat: add validation for negative quantities
fix: prevent duplicate entries in AML tab
refactor: split monolithic BOM file into modules
docs: update README for modular architecture
```

### Pull Request Process
1. Create feature branch from `main`
2. Make changes and test thoroughly
3. Update documentation if API changes
4. Submit PR with description of changes
5. Code review by maintainer

---

## Known Issues

1. **Large BOMs** — Sheets with >1000 items may hit GAS execution time limits during comparison
   - **Mitigation**: Batch operations already implemented; further optimisation via `CacheService` available

2. **Concurrency** — Multiple users submitting ECRs simultaneously may cause conflicts
   - **Workaround**: Coordinate ECR submissions
   - **Planned**: Implement locking mechanism

---

## Additional Resources

- [Google Apps Script Documentation](https://developers.google.com/apps-script)
- [Apps Script Best Practices](https://developers.google.com/apps-script/guides/support/best-practices)
- [Clasp Documentation](https://github.com/google/clasp)

---

**Engineering Team Contact**: Aqil
**Department**: Celestica CA1 Engineering
**Internal Use Only** — Celestica Corporation

---

**Last Updated**: February 2026
**Version**: 2.1.0
**Maintainer**: aqilzuna43
