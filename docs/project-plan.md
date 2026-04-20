# Project Plan

## Project Overview

**Project**: TateByWater Document Management  
**Focus**: Migrate TBCMS document/file management from shared folders to Dropbox Business  
**Status**: Planning approved, implementation pending

## Active Workstream

### In Scope Now
1. Finalize migration design and implementation sequencing.
2. Validate data model and metadata mapping for Dropbox references.
3. Prepare pilot rollout and rollback readiness.

### Canonical Migration Plan
- `docs/dropbox-migration-plan.md`

## Current Priorities

1. **Document workflow mapping**
   - Confirm all form and VBA entry points that perform open/save/move/copy operations.
2. **Metadata contract updates**
   - Define how existing local path fields/procedures map to Dropbox path/file-id references.
3. **OAuth user model**
   - Finalize per-user authentication and token lifecycle in local Access frontends.
4. **Pilot readiness**
   - Define pilot users, success criteria, and rollback trigger conditions.

## Completed Milestones

- Dropbox API POC completed and validated in `msaccess/DropboxPOC.accdb`.
- Current TBCMS document-management flow analyzed from extracted VBA.
- Migration strategy agreed: API-native Dropbox + per-user local token storage.
- Consolidated migration plan published in `docs/dropbox-migration-plan.md`.

## Risks to Track

- Metadata migration quality for existing document references.
- Dropbox permission alignment for all legal workflows.
- User onboarding friction during OAuth rollout.
- Runtime behavior under API throttling and intermittent network issues.

## Notes

- Keep this file concise and execution-oriented.
- Treat `docs/dropbox-migration-plan.md` as the source of truth for phases and design details.