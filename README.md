# TateByWater Document Management - Database Assessment Tools

**Comprehensive MS Access database extraction and analysis tools**

[![Platform](https://img.shields.io/badge/platform-Windows-blue.svg)](https://www.microsoft.com/windows)
[![Python](https://img.shields.io/badge/python-3.6+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/license-Proprietary-red.svg)]()

---

## 📋 Overview

This project provides Python tools to extract complete information from MS Access databases, including:
- ✅ **Tables** - Schema, fields, indexes, record counts
- ✅ **Queries** - SQL, parameters, output fields
- ✅ **VBA Code** - All modules, classes, forms, and reports

**Database:** TB CMS.SQL.accdb (Case Management System)
- 91 tables | 211 queries | 158 VBA components | 19,367 lines of code

---

## 🚀 Quick Start

### Prerequisites
- **Windows** (COM automation requirement)
- **Python 3.6+**
- **MS Access** installed
- **pywin32** library

### Installation

```bash
# Install dependencies
pip install -r requirements.txt
```

### Usage

#### Comprehensive Assessment (Tables + Queries + VBA)
```bash
py assess_access_db.py "path/to/database.accdb" "output_directory"
```

**Output:**
- `{DatabaseName}_assessment.json` - Complete structured data
- `{DatabaseName}_assessment.md` - Formatted markdown report
- `{DatabaseName}_assessment.txt` - Human-readable text
- `queries/` - Individual SQL files (211 files)
- `vba_code/` - Individual VBA files (158 files)

#### VBA Only Extraction
```bash
py extract_vba.py "path/to/database.accdb" "output_directory"
```

**Output:**
- `.bas` - Standard modules
- `.cls` - Class modules
- `.form.bas` - Forms with code
- `.report.bas` - Reports with code

---

## 📊 Project Structure

```
TateByWater-DocumentManagement/
├── assess_access_db.py          # Comprehensive assessment tool
├── extract_vba.py                # VBA extraction only
├── requirements.txt              # Python dependencies
├── README.md                     # This file
│
├── database_assessment/          # Assessment outputs
│   └── TB_CMS_SQL/              # TB CMS database
│       ├── TB_CMS_SQL_assessment.json (37,915 lines)
│       ├── TB_CMS_SQL_assessment.md
│       ├── TB_CMS_SQL_assessment.txt
│       ├── queries/ (211 SQL files)
│       └── vba_code/ (158 VBA files)
│
├── msaccess/                     # Source databases
│   ├── TB CMS.SQL.accdb         # Main CMS database
│   └── DropboxPOC.accdb         # Dropbox integration POC
│
└── docs/                         # Project documentation
    ├── PROJECT-SUMMARY.md        # Project overview
    ├── project-plan.md           # Current work & roadmap
    ├── tech-debt.md              # Known issues
    ├── architecture-decisions.md # Design decisions (8 ADRs)
    ├── vba-extraction-notes.md   # VBA domain knowledge
    ├── document-management-analysis.md # Document system analysis
    ├── dropbox-migration-plan.md # Dropbox migration (10-14 weeks)
    └── DROPBOX-POC-FINAL.md      # Dropbox POC results
```

---

## 🎯 Key Features

### `assess_access_db.py` - Comprehensive Assessment ⭐

**Extracts:**
- **Table Definitions** - Fields (name, type, size, constraints), indexes, record counts
- **Query Definitions** - Full SQL, query types, parameters
- **VBA Code** - All components with categorization

**Multi-Format Output:**
- JSON (machine-readable)
- Markdown (formatted report)
- Text (human-readable)
- Individual files (SQL & VBA)

**Advanced Features:**
- Field type mapping (Access codes → readable names)
- Graceful error handling (continues on failures)
- Metadata preservation
- Organized output structure

### `extract_vba.py` - VBA Extraction

**Features:**
- Focused VBA-only extraction
- Component type detection
- Metadata headers
- Proper COM lifecycle management
- UTF-8 encoding with error replacement
- Retry logic for locked databases

---

## 📈 TB CMS Database Stats

| Category | Count |
|----------|-------|
| **Tables** | 91 |
| **Queries** | 211 |
| **VBA Components** | 158 |
| **VBA Lines** | 19,367 |

### Key Tables
- `tblCase` - Case management
- `tblMatter` - Matter tracking
- `Billing` - Billing (1,783 records)
- `Bankruptcy` - Bankruptcy cases (211 records)
- `Trust Account` - Trust transactions
- `TB Time Keeping` - Time tracking

### Key Modules
- `DocumentManagement` - 841 lines (document operations)
- `Authentication` - User login
- `CaseGeneratorModule` - File number generation
- `FormUtils` - Form utilities

### Forms
- 48 forms with code (case, billing, time tracking)

### Reports
- 90 reports with code (invoices, trust, case reports)

---

## 🔧 Configuration

### Database Paths (Default)
```python
# Main database
accdb_path = "msaccess/TB CMS.SQL.accdb"

# Output directory
output_dir = "database_assessment/TB_CMS_SQL"
```

### Command Line Override
```bash
# Custom paths
py assess_access_db.py "C:/path/to/database.accdb" "C:/output/dir"
```

---

## ⚠️ Platform Requirements

**Windows Only**
- Project uses COM automation (`win32com.client`)
- Requires MS Access installation
- Cannot run on Linux/Mac

**Python Version**
- Python 3.6+ required
- Type hints use modern syntax (`dict | None`)

**Dependencies**
- `pywin32>=306` - COM automation

---

## 📚 Documentation

| Document | Purpose |
|----------|---------|
| **README.md** | Quick start & usage (this file) |
| **PROJECT-SUMMARY.md** | Comprehensive project overview |
| **project-plan.md** | Current work & roadmap |
| **tech-debt.md** | Known issues & improvements |
| **architecture-decisions.md** | Design decisions (8 ADRs) |
| **dropbox-migration-plan.md** | Dropbox migration plan |
| **DROPBOX-POC-FINAL.md** | Dropbox POC results |

---

## 🚧 Current Status

### ✅ Completed
- Comprehensive database assessment tool
- VBA extraction tool
- TB CMS.SQL.accdb complete assessment
- Dropbox API POC (successful)
- Documentation framework

### 🔨 In Progress
- Type hints addition
- Logging implementation
- Code quality improvements

### 📋 Planned
- README completion (this document)
- Dropbox migration Phase 1
- Testing framework (TBD)

---

## 🎓 Code Quality

### Cursor Rules (8 rules enforced)
- `python-standards.mdc` - Type hints, PEP 8
- `windows-com-automation.mdc` - COM lifecycle
- `file-io-patterns.mdc` - Path handling
- `error-handling.mdc` - Exception patterns
- `project-structure.mdc` - Organization
- `vba-extraction-workflow.mdc` - Domain knowledge
- `dependency-management.mdc` - Requirements
- `documentation-practices.mdc` - Docs framework

### Best Practices
- ✅ Proper COM cleanup (finally blocks)
- ✅ UTF-8 encoding with error handling
- ✅ Graceful degradation (continues on errors)
- ✅ Metadata preservation
- ✅ Comprehensive error messages

---

## 🔍 Troubleshooting

### Database Locked Error
```
Error: Database is locked by another process.
```
**Solution:** Close MS Access and all applications accessing the database.

### COM Automation Error
```
ImportError: No module named 'win32com'
```
**Solution:** Install pywin32: `pip install pywin32>=306`

### Access Not Installed
```
Error: Access.Application COM object not found
```
**Solution:** Install MS Access on your system.

---

## 📞 Support

**Project Location:** `C:\GitHub\TateByWater-DocumentManagement`  
**Documentation:** See `docs/` directory  
**Issues:** See `docs/tech-debt.md`

---

## 📄 Project Information

**Created:** 2024  
**Updated:** January 13, 2026  
**Python:** 2 scripts (~1,051 lines)  
**Documentation:** 8 markdown files  
**Status:** Active Development

---

## 🎯 Next Steps

1. **Add type hints** to Python code
2. **Implement logging** module
3. **Begin Dropbox migration** Phase 1
4. **Create test databases** for development

---

**For detailed project information, see [`docs/PROJECT-SUMMARY.md`](docs/PROJECT-SUMMARY.md)**
