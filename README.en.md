# Agent Memory Management

A **production-validated** multi-layer memory management scheme for LLM agents. It was distilled from long-term real-world production use (Hermes Agent) into a framework-agnostic, copyable protocol.

**Core idea: memory is not a single cache — it is a three-layer division of labor.**

```
┌─────────────────────────────────────────────────┐
│  L1 Cache Memory (injection layer · every turn)  │
│  MEMORY.md (agent notes) + USER.md (user profile) │
│  Budget-capped (char/token), tagged, injected     │
│  into the system prompt every round               │
├─────────────────────────────────────────────────┤
│  L2 Storage Memory (index layer · on-demand read) │
│  memory-kb/ dual-card structure (topic index      │
│  cards + session summary cards), auto-maintained  │
│  catalog, watchdog two-way sync, unbounded        │
│  Light start: a single MEMORY-KB.md works too     │
├─────────────────────────────────────────────────┤
│  L3 Session Logs (stream layer · auditable)       │
│  Raw conversation + tool calls, searchable,       │
│  periodically archived                            │
└─────────────────────────────────────────────────┘
```

---

## Why three layers

| Layer | Question it answers | Cost of not having it |
|---|---|---|
| L1 Cache Memory | What must the agent know every round? | Stuff everything in → context explosion, token burn every turn |
| L2 Storage Memory | Where do full topic details live? How to find them across sessions? | Cache only → budget overflow; sessions only → can't search |
| L3 Session Logs | What actually happened? How to review? | Index only, no raw source → distorted details, no traceability |

**One-line principle: what can live in L3 should not live in L1; L1 holds only hard rules needed every turn; L2 holds knowledge you can find when needed.**

---

## Quick start (30-minute onboarding)

1. **Create the directory structure** (see `templates/` and `docs/03-storage-memory-protocol.md`)
2. **Initialize L1**: write MEMORY.md (agent notes) + USER.md (user profile); tag each entry, separate with `§`
3. **Initialize L2**: create memory-kb/ and the master catalog, add the first topic index card (**light start**: a single MEMORY-KB.md file works until it grows)
4. **Wire into the agent**: inject the two L1 files into the system prompt; declare the two trigger words — "archive" and "store memory"
5. **Optional sync**: deploy the md↔xlsx watchdog so humans can manage L2 in Excel

## Repository layout

```
agent-memory-management/
├── README.md                    # this doc (Chinese) / README.en.md (English)
├── docs/
│   ├── 01-three-layer-architecture.md
│   ├── 02-cache-memory-protocol.md
│   ├── 03-storage-memory-protocol.md
│   ├── 04-session-archive-protocol.md
│   ├── 05-cross-agent-onboarding.md
│   └── 06-tag-system.md
├── templates/
│   ├── MEMORY.template.md       # L1 agent notes template
│   ├── USER.template.md         # L1 user profile template
│   ├── topic-index-card.template.md
│   ├── session-summary-card.template.md
│   ├── handoff-checklist.template.md
│   └── master-catalog.template.md
├── scripts/
│   ├── memorykb_sync.py         # md↔xlsx two-way sync + catalog maintenance (watchdog)
│   └── archive_reminder.py      # periodic archive reminders
└── requirements.txt             # openpyxl / watchdog
```

## Concepts at a glance

- **Trigger words**: `archive` = wrap up the current session (L3 → archive store + session summary card into L2 + handoff checklist); `store memory` = update L2 index. Don't mix them up.
- **Memory review**: background auto-writes can be disabled; major writes require human confirmation first; periodically scan all entries and adjudicate keep/delete/edit by tag.
- **Cross-agent**: files are plain Markdown + tag conventions — any LLM can read them; sync scripts are reusable.

---

> Protocol details live in `docs/`. The scheme does not depend on any specific agent framework — the L1/L2 file formats and the L3 archive workflow port directly to any LLM application.

Licensed under MIT.
