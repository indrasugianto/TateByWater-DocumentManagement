# Technical Debt

## Critical Items

### Multi-user Dropbox token isolation
- **Priority**: HIGH
- **Impact**: Incorrect user identity could be used for document operations.
- **Context**: Existing POC token model is not production-safe for shared use.
- **Action**: Enforce per-user token ownership and startup identity checks.

### Production OAuth hardening
- **Priority**: HIGH
- **Impact**: Increased auth and security risk during rollout.
- **Context**: POC flow needs production controls.
- **Action**: Add OAuth state handling, robust token-expiry handling, and hardened error paths.

### Metadata migration quality
- **Priority**: HIGH
- **Impact**: Broken file links during/after migration from local paths to Dropbox references.
- **Action**: Build migration verification scripts and reconciliation reporting before cutover.

## Important (Non-blocking) Debt

### API throttling/network resilience
- **Priority**: MEDIUM
- **Impact**: Intermittent failures under load/network instability.
- **Action**: Central retry policy with `Retry-After` support and operator-facing diagnostics.

### Auditability and observability
- **Priority**: MEDIUM
- **Impact**: Harder incident response and compliance tracing.
- **Action**: Standardized structured logging for upload/move/delete/open-link operations.

### Legacy workflow compatibility
- **Priority**: MEDIUM
- **Impact**: User-facing regressions if migration is all-at-once.
- **Action**: Provider abstraction + feature flags for phased rollout and rollback.

## Ongoing Constraints

### Windows-only COM workflow
- **Priority**: LOW
- **Impact**: Tooling remains Windows-bound.
- **Action**: Documented limitation; accepted by design.

## Reference

- Canonical migration plan: `docs/dropbox-migration-plan.md`
- Current execution priorities: `docs/project-plan.md`
