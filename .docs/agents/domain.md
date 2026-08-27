# Domain Docs

How the engineering skills should consume this repo's domain documentation when exploring the codebase.

This repo keeps its docs under `.docs/` (dot-prefixed), not `docs/` — that's a deliberate repo convention (see the master plan and analysis docs already there), and the paths below follow it.

## Before exploring, read these

- **`.docs/CONTEXT.md`** at the repo root, or
- **`.docs/CONTEXT-MAP.md`** at the repo root if it exists — it points at one `CONTEXT.md` per context. Read each one relevant to the topic.
- **`.docs/adr/`** — read ADRs that touch the area you're about to work in. In multi-context repos, also check `src/<context>/.docs/adr/` for context-scoped decisions.

Also worth reading for this repo specifically, even though they predate this ADR/CONTEXT convention:

- **`.docs/dropbox-migration-plan.md`** — the master plan for the TBCMS → Dropbox migration; start at its `▶ NEXT SESSION: START HERE` block.
- **`.docs/document-management-analysis.md`** — grounded current-state analysis of the legacy document-management subsystem.
- **`.docs/dropbox-bridge-plan.md`** and **`.docs/bridge-deployment-runbook.md`** — the Dropbox bridge component's plan and deployment runbook.

If `.docs/CONTEXT.md` / `.docs/CONTEXT-MAP.md` / `.docs/adr/` don't exist yet, **proceed silently**. Don't flag their absence; don't suggest creating them upfront. The `/domain-modeling` skill (reached via `/grill-with-docs` and `/improve-codebase-architecture`) creates them lazily when terms or decisions actually get resolved.

## File structure

Single-context repo (this repo):

```
/
├── .docs/
│   ├── CONTEXT.md
│   ├── adr/
│   │   ├── 0001-....md
│   │   └── 0002-....md
│   ├── dropbox-migration-plan.md
│   ├── document-management-analysis.md
│   ├── dropbox-bridge-plan.md
│   └── bridge-deployment-runbook.md
└── Dropbox-Migration/
```

Multi-context repo (presence of `.docs/CONTEXT-MAP.md` at the root) — not this repo's current shape, but the pattern to switch to if it ever splits into multiple bounded contexts:

```
/
├── .docs/
│   ├── CONTEXT-MAP.md
│   └── adr/                           ← system-wide decisions
└── src/
    ├── ordering/
    │   ├── .docs/CONTEXT.md
    │   └── .docs/adr/                 ← context-specific decisions
    └── billing/
        ├── .docs/CONTEXT.md
        └── .docs/adr/
```

## Use the glossary's vocabulary

When your output names a domain concept (in an issue title, a refactor proposal, a hypothesis, a test name), use the term as defined in `.docs/CONTEXT.md`. Don't drift to synonyms the glossary explicitly avoids.

If the concept you need isn't in the glossary yet, that's a signal — either you're inventing language the project doesn't use (reconsider) or there's a real gap (note it for `/domain-modeling`).

## Flag ADR conflicts

If your output contradicts an existing ADR, surface it explicitly rather than silently overriding:

> _Contradicts ADR-0007 (event-sourced orders) — but worth reopening because…_
