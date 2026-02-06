# CA1-PLM: Lightweight Product Lifecycle Management System

> A Google Apps Script-based PLM tool for managing Bill of Materials (BOM) and Engineering Change Requests (ECR) at Celestica CA1

[![License: Internal](https://img.shields.io/badge/License-Internal-blue.svg)]()
[![Platform: Google Apps Script](https://img.shields.io/badge/Platform-Google%20Apps%20Script-green.svg)]()

## 🎯 Overview

CA1-PLM implements a **hub-and-spoke architecture** where a centralized Google Sheet serves as the "Source of Truth" for all product data, with satellite ECR forms acting as controlled transaction clients. This eliminates direct manual editing of the Master BOM while maintaining data integrity and change traceability.

### Key Design Principles
- **Single Source of Truth**: All product data centralized in one master spreadsheet
- **Controlled Modifications**: Changes only via validated ECR workflow
- **Hierarchical BOM Structure**: Level-based parent/child relationships
- **Data Validation**: Pre-commit validation against existing BOM structure

---

## 🏗️ System Architecture

### Hub-and-Spoke Model
```
┌─────────────────────────────────────────────────────────┐
│          12016271_CA1 Mini BoM (THE HUB)               │
│                  Source of Truth                         │
├─────────────────────────────────────────────────────────┤
│  ┌─────────┐    ┌─────────┐    ┌──────────┐           │
│  │  ITEMS  │    │   AML   │    │  MASTER  │           │
│  │   Tab   │    │   Tab   │    │   Tab    │           │
│  └─────────┘    └─────────┘    └──────────┘           │
│       │              │                │                 │
│   Part Data    Manufacturer      BOM Structure         │
│  Dictionary        List          (Hierarchical)        │
└─────────────────────────────────────────────────────────┘
                         ▲
                         │
                         │ Read/Write via API
                         │
┌────────────────────────┴─────────────────────────────────┐
│                   ECR Form (THE SPOKE)                    │
│                  Transaction Client                       │
├───────────────────────────────────────────────────────────┤
│  • Validates change requests against Master BOM           │
│  • Prevents invalid Parent/Child relationships            │
│  • Executes controlled modifications via commitToMaster() │
│  • Change Types: ADD | REMOVE | QTY CHANGE | REV ROLL    │
└───────────────────────────────────────────────────────────┘
```

---

## 📊 Data Model

### The Hub: `12016271_CA1 Mini BoM`

#### 1. ITEMS Tab (Part Dictionary)
Defines **what** each part is.

| Column | Description |
|--------|-------------|
| Part Number | Unique identifier |
| Description | Part description |
| Revision | Current revision level |
| Lifecycle | Status (Active, Obsolete, etc.) |

#### 2. AML Tab (Approved Manufacturer List)
Defines **where** parts can be sourced.

| Column | Description |
|--------|-------------|
| Part Number | Reference to ITEMS |
| Manufacturer Name | Approved supplier |
| Manufacturer Part Number | Supplier's part number |

#### 3. MASTER Tab (Bill of Materials)
Defines **how** parts are assembled - uses hierarchical level structure.

| Column | Description | Critical for |
|--------|-------------|-------------|
| **Level** | Hierarchy depth (e.g., "1", "1.1", "1.1.1") | Parent/Child detection |
| Part Number | Reference to ITEMS | Part identification |
| Quantity | Assembly quantity | BOM calculations |
| Reference Designator | PCB location | Manufacturing |

**Hierarchy Logic Example:**
```
Level    Part Number    Description
1        ASSY-001       Main Assembly
1.1      PCB-100        PCB Board         <- Child of ASSY-001
1.1.1    R-001          Resistor          <- Child of PCB-100
1.1.2    C-001          Capacitor         <- Child of PCB-100
1.2      CABLE-200      Cable Assembly    <- Child of ASSY-001
2        ASSY-002       Second Assembly   <- New top-level parent
```

---

## 🔧 Core Functionality

### File: `CA1 MINI BOM.gs`

**BOM Management & Analysis**

| Function | Purpose | Use Case |
|----------|---------|----------|
| `runDetailedComparison()` | Compare two BOM versions | Engineering change impact analysis |
| `graftPDMData()` | Import external PDM data | Sync with upstream CAD/PLM systems |
| `generateFabricatorList()` | Generate flat BOM | Manufacturing handoff |
| `findChildren()` | Parse hierarchy | Navigate parent/child relationships |

**Key Logic:**
- **Level Parsing**: Scans the "Level" column to determine assembly structure
- **Child Detection**: Finds all items under a parent until the level resets to a new parent
- **Version Comparison**: Identifies added/removed/modified items between BOM snapshots

### File: `ECR-FORM.gs`

**Engineering Change Request Workflow**

| Function | Purpose | Critical Operation |
|----------|---------|-------------------|
| `populateCurrentData()` | Validate user input | Prevents invalid relationships |
| `commitToMaster()` | Execute changes | **THE ENGINE** - Modifies MASTER BOM |
| `validateChange()` | Pre-commit checks | Data integrity enforcement |

#### Change Types Supported
```javascript
// ADD: Insert new child component
{ 
  action: "ADD",
  parent: "ASSY-001",
  child: "R-NEW-001",
  qty: 10,
  refDes: "R100-R109"
}

// REMOVE: Delete component from assembly
{ 
  action: "REMOVE",
  parent: "ASSY-001",
  child: "C-OLD-001"
}

// QTY CHANGE: Update component quantity
{ 
  action: "QTY CHANGE",
  parent: "ASSY-001",
  child: "R-001",
  newQty: 20
}

// REV ROLL: Update component revision
{ 
  action: "REV ROLL",
  item: "PCB-100",
  oldRev: "A",
  newRev: "B"
}
```

#### `commitToMaster()` Algorithm

**Critical Logic Flow:**
```javascript
1. Locate Parent Item in MASTER array
   └─> Search for matching Part Number at correct Level

2. Scan Children under Parent
   └─> Continue until Level resets (new parent) or Level decreases

3. Execute Action:
   ADD:
     └─> Find correct insertion point (maintain level hierarchy)
     └─> Insert row with calculated Level value
   
   REMOVE:
     └─> Locate exact Child match
     └─> Delete row
   
   QTY CHANGE:
     └─> Locate Child
     └─> Update Quantity column
   
   REV ROLL:
     └─> Update Revision in ITEMS tab
     └─> Propagate to all BOM instances
```

---

## 🚀 Getting Started

### Prerequisites
- Google Account with access to Google Sheets
- Node.js & npm (for local development with clasp)
- Git

### Installation

#### Option 1: Direct Google Apps Script (Simple)
1. Open your Google Sheet: `12016271_CA1 Mini BoM`
2. Go to **Extensions > Apps Script**
3. Copy content from `CA1 MINI BOM.gs` into the script editor
4. Save and authorize permissions

#### Option 2: Local Development with Clasp (Recommended)
```bash
# Clone this repository
git clone https://github.com/aqilzuna43/CA1-PLM.git
cd CA1-PLM

# Install clasp globally
npm install -g @google/clasp

# Login to Google
clasp login

# Link to your existing Apps Script project
clasp clone <YOUR_SCRIPT_ID>

# Or create a new project
clasp create --type sheets --title "CA1 PLM Tool"
```

### Finding Your Script ID
1. Open your Google Sheet
2. **Extensions > Apps Script**
3. Click **Project Settings** (gear icon)
4. Copy the **Script ID**

---

## 💻 Development Workflow

### Local Development Cycle
```bash
# Pull latest from Google Apps Script
clasp pull

# Make your changes in .gs files locally
# Use your preferred IDE (VS Code, etc.)

# Push changes to Google Apps Script
clasp push

# Test in your Google Sheet
# Open the sheet and run functions manually or via custom menu

# Once tested, commit to Git
git add .
git commit -m "feat: add validation for negative quantities"
git push origin main
```

### Testing Strategy
1. **Unit Testing**: Test individual functions in Apps Script debugger
2. **Integration Testing**: Use a TEST copy of the Master BOM
3. **UAT**: Engineers test ECR workflow on real (but non-critical) assemblies
4. **Production**: Deploy to live `12016271_CA1 Mini BoM`

---

## 📈 Performance Optimization (Roadmap)

### Current State
- **Codebase**: 1,600+ lines across 2 main files
- **Performance**: Row-by-row processing causing bottlenecks
- **Target**: 4-15x speed improvement via batch operations

### Identified Bottlenecks

#### 1. Row-by-Row API Calls
```javascript
// ❌ SLOW: Individual writes
for (let i = 0; i < items.length; i++) {
  sheet.getRange(i+1, 1).setValue(items[i]);
}

// ✅ FAST: Batch write (4-15x faster)
let values = items.map(item => [item]);
sheet.getRange(1, 1, items.length, 1).setValues(values);
```

#### 2. Redundant Data Fetching
```javascript
// ❌ SLOW: Multiple identical API calls
function getItems() {
  return itemsSheet.getDataRange().getValues();
}
// Called 100+ times in one execution = 100+ API calls

// ✅ FAST: Cache data
let itemsCache = null;
function getItems() {
  if (!itemsCache) {
    itemsCache = itemsSheet.getDataRange().getValues();
  }
  return itemsCache;
}
```

### Optimization Phases
- [ ] **Phase 1**: Audit all `getRange()` and `setValue()` calls
- [ ] **Phase 2**: Convert to batch operations (`getValues()`, `setValues()`)
- [ ] **Phase 3**: Implement script-level caching
- [ ] **Phase 4**: Add `CacheService` for multi-execution persistence
- [ ] **Phase 5**: Performance benchmarking and validation

**See [OPTIMIZATION.md](./docs/OPTIMIZATION.md) for detailed strategies**

---

## 📁 Project Structure
```
CA1-PLM/
├── CA1 MINI BOM.gs           # Hub: BOM management functions
├── ECR-FORM.gs               # Spoke: ECR transaction client
├── appsscript.json           # Apps Script configuration
├── .gitignore                # Git exclusions
├── README.md                 # This file
└── docs/
    ├── OPTIMIZATION.md       # Performance improvement guide
    ├── API_REFERENCE.md      # Function documentation
    └── CHANGELOG.md          # Version history
```

---

## 🛡️ Data Integrity & Safety

### Protection Mechanisms
1. **Validation Layer**: `populateCurrentData()` prevents invalid relationships
2. **Transaction Model**: All changes via ECR (audit trail)
3. **Read-Before-Write**: `commitToMaster()` validates target exists before modification
4. **Manual Override Prevention**: Users cannot directly edit MASTER tab

### Backup Strategy
- **Recommended**: Version history enabled on Google Sheets (File > Version History)
- **Critical**: Export Master BOM weekly to backup location
- **Testing**: Always test ECR changes on a COPY first

---

## 🤝 Contributing

### Branching Strategy
```
main                     # Production code
├── develop              # Integration branch
│   ├── feature/xxx      # New features
│   ├── fix/xxx          # Bug fixes
│   └── optimize/xxx     # Performance improvements
```

### Commit Message Convention
```
feat: add batch processing for commitToMaster
fix: prevent duplicate entries in AML tab
optimize: cache ITEMS data to reduce API calls
docs: update API reference for findChildren function
```

### Pull Request Process
1. Create feature branch from `develop`
2. Make changes and test thoroughly
3. Update documentation if API changes
4. Submit PR with description of changes
5. Code review by maintainer
6. Merge to `develop`, then to `main` for release

---

## 📝 Roadmap

### Short Term (Q1 2026)
- [ ] Performance optimization (batch operations)
- [ ] Enhanced error handling in `commitToMaster()`
- [ ] Add logging for ECR audit trail

### Medium Term (Q2 2026)
- [ ] Automated BOM validation rules
- [ ] Integration with upstream PDM system
- [ ] Dashboard for ECR status tracking

### Long Term (Q3+ 2026)
- [ ] Multi-user concurrent ECR support
- [ ] BOM cost rollup calculations
- [ ] Export to ERP system (SAP/Oracle)

---

## 🐛 Known Issues

1. **Performance**: Large BOMs (>1000 items) experience slow `commitToMaster()` execution
   - **Workaround**: Batch multiple ECRs before committing
   - **Fix Planned**: Phase 2 optimization (batch operations)

2. **Concurrency**: Multiple users submitting ECRs simultaneously may cause conflicts
   - **Workaround**: Coordinate ECR submissions
   - **Fix Planned**: Implement locking mechanism

---

## 📚 Additional Resources

- [Google Apps Script Documentation](https://developers.google.com/apps-script)
- [Apps Script Best Practices](https://developers.google.com/apps-script/guides/support/best-practices)
- [Clasp Documentation](https://github.com/google/clasp)

---

## 📞 Support

**Engineering Team Contact**: Aqil  
**Department**: Celestica CA1 Engineering  
**Internal Use**: This tool is for internal Celestica use only

---

## 📄 License

**Internal Use Only** - Celestica Corporation  
Not for external distribution

---

## ⚡ Quick Start Commands
```bash
# Setup
git clone https://github.com/aqilzuna43/CA1-PLM.git
cd CA1-PLM
clasp login
clasp clone <SCRIPT_ID>

# Development
clasp pull          # Download from Google
# ... make changes ...
clasp push          # Upload to Google
clasp open          # Open in browser

# Version Control
git add .
git commit -m "your message"
git push origin main
```

---

**Last Updated**: February 2026  
**Version**: 1.0.0  
**Maintainer**: aqilzuna43
