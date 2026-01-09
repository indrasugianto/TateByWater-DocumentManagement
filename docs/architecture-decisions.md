# Architecture Decision Records

## ADR-001: Python for VBA Extraction

**Date**: 2024  
**Status**: Accepted

### Context
Need to extract VBA code from MS Access databases for version control and analysis.

### Decision
Use Python with `win32com.client` for COM automation to access MS Access databases.

### Consequences
- ✅ Python is cross-platform, but COM automation is Windows-only
- ✅ `win32com` provides direct access to Access VBA project
- ❌ Project is limited to Windows platform
- ⚠️ Requires MS Access to be installed on system

---

## ADR-002: Single Script Architecture

**Date**: 2024  
**Status**: Accepted

### Context
Project is a utility tool for extracting VBA code from Access databases.

### Decision
Use a single Python script (`extract_vba.py`) rather than a multi-module package.

### Consequences
- ✅ Simple structure, easy to understand
- ✅ No complex imports or package setup
- ⚠️ May need refactoring if project grows
- ✅ Appropriate for current scope

---

## ADR-003: Dictionary-Based Statistics

**Date**: 2024  
**Status**: Accepted

### Context
Need to track extraction statistics (counts, line numbers, component details).

### Decision
Use a dictionary to store statistics and return from function.

### Consequences
- ✅ Simple and flexible
- ✅ Easy to extend with new fields
- ✅ Can be easily serialized (JSON) if needed
- ✅ Clear structure

---

## ADR-004: UTF-8 Encoding with Error Replacement

**Date**: 2024  
**Status**: Accepted

### Context
VBA code may contain legacy encodings or invalid UTF-8 characters.

### Decision
Use `encoding='utf-8'` with `errors='replace'` when writing extracted files.

### Consequences
- ✅ Prevents crashes from encoding errors
- ✅ Invalid bytes replaced with replacement character
- ⚠️ Some characters may be lost (but code remains readable)
- ✅ More robust than strict encoding

---

## ADR-005: File Extension Naming Convention

**Date**: 2024  
**Status**: Accepted

### Context
Need to distinguish between different VBA component types in extracted files.

### Decision
Use extensions: `.bas` (modules), `.cls` (classes), `.form.bas` (forms), `.report.bas` (reports).

### Consequences
- ✅ Clear distinction between component types
- ✅ Standard VBA extensions for modules/classes
- ✅ Descriptive extensions for forms/reports
- ✅ Easy to filter by type

---

## ADR-006: Metadata Headers in Extracted Files

**Date**: 2024  
**Status**: Accepted

### Context
Extracted VBA files should include metadata about their source.

### Decision
Add header comments to each extracted file with component name, type, and line count.

### Consequences
- ✅ Preserves context about extracted code
- ✅ Easy to identify source component
- ✅ Helps with debugging and analysis
- ⚠️ Adds a few lines to each file (minimal impact)

---

## ADR-007: Documentation in Root docs/ Directory

**Date**: 2025-01-09  
**Status**: Accepted

### Context
Need location for project documentation and knowledge framework. Documentation practices rule originally referenced `docs/cursor/` but actual structure has docs at root.

### Decision
Place all documentation markdown files in `docs/` directory at project root (not nested in `docs/cursor/`).

### Consequences
- ✅ Simpler directory structure
- ✅ Easier to access documentation
- ✅ Standard convention for many projects
- ✅ All documentation co-located

---

## ADR-008: Cursor Rules Ecosystem

**Date**: 2025-01-09  
**Status**: Accepted

### Context
Need to enforce consistent coding standards and preserve domain knowledge across AI sessions.

### Decision
Create comprehensive Cursor rules ecosystem with:
- Auto-attached rules for Python files (standards, COM automation, file I/O)
- Always-active rules for project structure and error handling
- Manual reference rules for domain knowledge and dependencies
- Documentation practices rule to maintain context

### Consequences
- ✅ Consistent code quality enforced automatically
- ✅ Domain knowledge preserved across sessions
- ✅ Best practices documented and enforced
- ✅ AI maintains context through documentation framework
- ⚠️ Requires keeping rules updated as project evolves
