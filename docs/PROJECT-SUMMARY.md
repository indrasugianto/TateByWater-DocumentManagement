# Project Summary - TateByWater Document Management

## Overview

This repository contains tooling and extracted artifacts for MS Access analysis, plus planning for migrating TBCMS document management from shared folders to Dropbox Business.

## Current Focus

- Primary initiative: Dropbox Business migration for TBCMS document workflows.
- Approved direction: API-native Dropbox integration with per-user OAuth in local Access frontends.
- Canonical implementation plan: `docs/dropbox-migration-plan.md`.

## Key Assets

- `msaccess/TBCMS/extract/` - Extracted forms, reports, modules, queries, lineage.
- `database_assessment/DropboxPOC/vba_code/DropboxAPI_POC.bas` - Dropbox API proof-of-concept module.
- `docs/document-management-analysis.md` - Detailed analysis of current document/file handling in TBCMS.
- `docs/dropbox-migration-plan.md` - Current migration plan and phased rollout approach.

## Documentation Guide

- `docs/project-plan.md` - Current execution-oriented priorities.
- `docs/tech-debt.md` - Open risks and technical debt.
- `docs/architecture-decisions.md` - Architecture decision records.
- `docs/DROPBOX-POC-FINAL.md` - Historical POC summary and lessons learned.
- `docs/vba-extraction-notes.md` - VBA/COM extraction domain notes.

## Notes

- This summary intentionally avoids static counts/timelines that become stale quickly.
- For migration specifics, always refer to `docs/dropbox-migration-plan.md`.
