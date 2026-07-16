# xlblueprint

[![CI](https://github.com/ShunsukeTamura06/xlblueprint/actions/workflows/ci.yml/badge.svg)](https://github.com/ShunsukeTamura06/xlblueprint/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](./LICENSE)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![Nuxt 3](https://img.shields.io/badge/Nuxt-3-00DC82.svg?logo=nuxt.js&logoColor=white)](https://nuxt.com/)

> **Understand and safely modernize legacy Excel workbooks (.xlsm / .xlsx)
> with an LLM that knows where every formula and VBA reference leads.**

📖 日本語版: [README.ja.md](./README.ja.md)

`xlblueprint` is built around two tasks for people who inherit an unfamiliar
business workbook:

1. **Investigate this Excel file** — get a plain-language diagnosis of its
   likely purpose, entry points, inputs, outputs, features, external
   dependencies, and risks. Every claim carries evidence IDs and is labelled
   as extracted fact, inference, or unknown.
2. **I need to change this Excel file** — select the affected feature, describe
   the desired business outcome, and turn it into a change brief containing
   current behavior, likely impact areas, clarification questions, evidence,
   and acceptance criteria before any edit begins.

Under those workflows, `xlblueprint` extracts every sheet, formula, named
range, conditional format, chart, pivot, VBA module, form control, and external
link, and also provides:

1. A **structured technical report** (Markdown + Mermaid)
2. An **interactive dependency graph** (sheets & VBA call graph)
3. A **grounded change consultation** that answers *"if I change X, what breaks?"*
   by calling tools over the extracted index (cells / refs / VBA /
   risks / external functions)
4. An honest **"what we can't statically analyze" report** — INDIRECT,
   OFFSET, dynamic VBA refs, event macros, external links. The tool
   refuses to say "no impact" where static analysis can't prove it.

It's aimed at the person who *inherited* a 10-year-old workbook and was
told "make this small change next Monday."

## Why

LLMs are great at writing new Excel formulas. They are terrible at
*understanding* a 50-sheet .xlsm full of `VLOOKUP` chains, dynamic VBA
ranges, and ten-year-old comments — because they have no map. `xlblueprint`
builds the map first, then lets the LLM walk it with grounded tool calls.

Existing tools either focus on writing new spreadsheets (Office Scripts,
"AI Excel" SaaS) or are libraries without an interactive layer
(openpyxl, formulas). `xlblueprint` sits in the gap: **a UI + LLM workbench
for legacy `.xlsm` modernization.**

## Screenshots

> _Screenshots below are placeholders until captured from a live deployment.
> See [Phase 2-C in the OSS launch plan](./docs/OSS_LAUNCH_PLAN.md)._

| Spec dashboard | Dependency graph | Chat with grounded tools |
|---|---|---|
| ![Spec overview](./docs/images/spec-overview.png) | ![Dependency graph](./docs/images/diagram.png) | ![Chat](./docs/images/chat.png) |

## Features

### Extraction (no LLM required)
- **Sheets**: dimensions, formulas (tokenized references), named ranges,
  conditional formats, Excel tables, data validations, merged cells,
  form controls
- **VBA**: modules and procedures via `oletools` (olevba), call-graph
  construction
- **OOXML internals**: charts (series references), pivot tables (source +
  fields), Power Query / external connections (connection strings with
  password fields auto-masked)
- **External link** detection

### Analysis
- **Evidence-backed workbook diagnosis** — headline, overview, feature map,
  inputs, outputs, external dependencies and warnings. It remains useful with
  no LLM configured and never promotes an unsupported guess to a fact
- **Business change brief** — desired outcome, current behavior, likely impact
  areas, follow-up questions, acceptance criteria and automation boundary,
  ready to hand into the grounded consultation flow
- **Reverse reference index** with range-intersection lookup
  (formula + VBA + chart series + pivot source)
- **Risk analyzer** — surfaces the things static analysis *cannot* track:
  `INDIRECT` / `OFFSET` / dynamic VBA refs (`Range(var)`,
  `Worksheets(name).Range(var)`), event macros (`Workbook_Open`,
  `Worksheet_Change`), runtime state (`ActiveSheet`, `Selection`),
  external dependencies
- **External function registry** — built-in Bloomberg BDH/BDP/BDS
  definitions; pluggable for in-house add-ins

### LLM (any OpenAI-compatible endpoint)
- **Spec annotation**: each sheet / procedure gets a short purpose +
  inputs/outputs/main-calculations summary (fast model, batched)
- **Chat with function calling** (10+ tools, 16-iteration loop):
  `get_cells_range`, `find_cells`, `lookup_references`,
  `list_vba_modules`, `get_vba_procedure`, `list_sheet_formulas`,
  `list_workbook_objects`, `list_analysis_risks`,
  `lookup_external_function`, `list_external_functions_used`

### Safety gate & safe auto-fix (roadmap S1/S2, shipping incrementally)
- **Structural diff** between two workbook versions (cells / named ranges /
  conditional formats / validations / charts / pivots / VBA modules) with
  **blast-radius** analysis — re-save noise is ignored by design
  (normalized-extraction comparison, validated against real Excel re-saves)
- **Safe auto-fix loop** for three narrow patterns: named-range
  redefinition, fixed-reference replacement, formula range expansion.
  The LLM can only *propose* (read-only impact estimate); applying
  requires an explicit button click, writes to a **new job** (the
  original file is never modified), and passes the result through an
  **exact structural-diff policy gate**. Every execution stores the
  provider, mutation plan, expected diff, observed diff, and verdict
  (`passed` / `needs_review` / `failed`) as an audit record
- **Pluggable mutation providers** separate the component that edits a
  workbook from xlblueprint's understanding and verification. The built-in
  openpyxl provider supports all three patterns. An optional OfficeCLI
  process adapter currently supports `.xlsx` named-range updates; OfficeCLI
  is not a required dependency and is never treated as the proof of safety

### Frontend (Nuxt 3 SPA, dark theme)
- **Two primary journeys**: "investigate this Excel" and "change this Excel";
  provider names and implementation details stay out of the normal workflow
- **Grounded diagnosis**: purpose/structure summary, feature cards,
  input/output candidates, dependencies, grouped warnings, evidence drill-down
- **Guided change request**: feature selection, business outcome, impact areas,
  questions and acceptance criteria, then pre-filled grounded consultation
- **Interactive diagrams** (Vue Flow + dagre, themed minimap)
- **Sheet explorer** with formula↔reference chips, A/B/C column
  headers and 1/2/3 row numbers in previews
- **Reverse reference search**
- **Chat** with progress streaming, multi-conversation, thumbs feedback

### Deployment
- **Air-gap friendly**: fonts and icons bundled, no runtime CDN calls
- **No data leaves your network**: LLM goes through whatever
  OpenAI-compatible endpoint you configure (Ollama, vLLM,
  self-hosted gateway, OpenAI, etc.)
- **600+ tests** (pytest), `mypy --strict` on `core/`, `ruff` clean

## Quick start

### Requirements
- Python 3.10+ with [uv](https://docs.astral.sh/uv/)
- Node 20+ with pnpm (via corepack)

### Install
```bash
# Backend
uv sync --group dev

# Frontend
corepack enable
cd frontend && pnpm install
```

### Run (two terminals)
```bash
# Terminal 1 — Backend on http://localhost:8001
uv run uvicorn backend.main:app --reload --port 8001

# Terminal 2 — Frontend on http://localhost:3001
cd frontend && cp .env.example .env
pnpm dev
```

### Try the sample
Open <http://localhost:3001>, click **Download sample**, and re-upload
`retail_monthly_ops.xlsx`. The sample is an 8-sheet retail-operations
workbook with 170 formulas, charts, pivots, INDIRECT/OFFSET, and
21 "unresolvable risk" items — a small but realistic showcase of what
the tool surfaces.

Without an LLM configured (default), the diagnosis, feature map, change brief,
technical report, diagrams, search, and risk analysis all work. The consultation
screen tells users that an administrator must configure the assistant before it
can create a grounded change plan.

## Architecture

```
Frontend (Nuxt 3 SPA, TypeScript)        Backend (FastAPI)
├── Nuxt UI v3 + Tailwind v4        <──> ├── /extract, /analyze
├── Vue Flow + dagre (diagrams)          ├── /diagnosis, /change-request
├── @nuxtjs/mdc (Markdown rendering)     ├── /spec, /workbook
└── Pinia                                ├── /references, /diagrams
                                         ├── /chat (function calling SSE)
                                         ├── /system/llm-status
                                         └── /jobs

                                         Core (pure Python)
                                         ├── olevba / openpyxl extraction
                                         ├── reference_index
                                         ├── risk_analyzer
                                         ├── diagnosis / change brief
                                         ├── external_functions registry
                                         ├── mutation plan / provider boundary
                                         ├── exact structural-diff verification
                                         ├── spec_generator (Markdown + Mermaid)
                                         └── diagrams (sheet / VBA graph)
```

Dependency direction is strictly `frontend → backend → core`. The core
is a pure Python library that you can use standalone:

```python
from pathlib import Path
from core.extractors.vba import extract_vba
from core.extractors.workbook import extract_workbook
from core.reference_index import build_reference_index
from core.risk_analyzer import detect_analysis_risks

path = Path("your_workbook.xlsm")
wb = extract_workbook(path)
wb.vba_modules = extract_vba(path)
wb.analysis_risks = detect_analysis_risks(wb)
idx = build_reference_index(wb)
# inspect wb / idx programmatically
```

More detail: [docs/architecture.md](./docs/architecture.md) (English summary)
and [docs/SPEC.ja.md](./docs/SPEC.ja.md) (full Japanese spec).

## Verification boundary

On macOS and Linux, xlblueprint can extract, diagnose, build a reference map,
create change briefs, apply the currently supported narrow OOXML edits, and
verify exact normalized structural diffs. That is **static and structural
verification**, not proof that Excel executed the workbook correctly.

Recalculation in Microsoft Excel, VBA compilation/execution, event macros,
COM/VBIDE behavior, protected projects, and environment-specific external
connections require a Windows machine with Microsoft Excel. OfficeCLI is an
optional editing adapter, not a safety oracle; its process success is never
accepted as verification without xlblueprint re-extraction and policy checks.

## Configuration

| Variable | Default | Description |
|---|---|---|
| `JOBS_DIR` | `./jobs` | Where uploads and extraction outputs are stored |
| `LLM_BASE_URL` | (unset) | OpenAI-compatible LLM base URL (Ollama, vLLM, OpenAI, self-hosted, …) |
| `LLM_API_KEY` | (unset) | LLM API key (any string works for local Ollama) |
| `LLM_MODEL` | (unset) | Default model name |
| `LLM_MODEL_PRO` | = `LLM_MODEL` | Model for chat (caching-friendly) |
| `LLM_MODEL_FAST` | = `LLM_MODEL` | Model for batch annotation (cheap/fast) |
| `CORS_ALLOW_ORIGINS` | localhost:3001,3000 | Comma-separated origin list |
| `CHAT_HISTORY_LIMIT_PAIRS` | `10` | How many recent user/assistant pairs to keep in LLM context |
| `MAX_UPLOAD_BYTES` | `52428800` (50 MB) | Upload size cap |
| `OFFICECLI_BIN` | PATH lookup | Optional OfficeCLI executable used by the mutation-provider adapter |
| `NUXT_PUBLIC_BACKEND_URL` | `http://localhost:8001` | Backend URL the SPA hits |
| `NUXT_PORT` | `3001` | Frontend dev server port |

If LLM variables are unset, `xlblueprint` falls back to `MockLLMClient` —
spec generation / diagrams / search still work; chat returns mock
responses with an in-app onboarding card pointing here.

### Ollama example
```bash
LLM_BASE_URL=http://localhost:11434/v1 \
LLM_API_KEY=ollama \
LLM_MODEL=llama3.1:8b \
uv run uvicorn backend.main:app --port 8001
```

## Limitations

- `.xls` (legacy binary): VBA is extracted but `openpyxl` cannot read
  sheet structure
- VBA projects with a project password cannot be opened (olevba limit)
- Power Query: connection / output inventory is captured; M-code body
  and embedded credentials are not deeply parsed (and credentials are
  masked when seen)
- Designed for workbooks up to ~50 MB / ~5000 rows. Larger files work
  but parsing time and memory grow accordingly
- Pivot table *creation* (vs. extraction) is out of scope; the sample
  workbook uses `SUMIFS` as a pivot stand-in

## Project status

- **0.1.x — Beta.** Core extraction is stable and well tested
  (637 tests, `mypy --strict` on `core/`). LLM-side prompts and UI
  layout are still iterating.
- **Maintained by**: Shunsuke Tamura ([@ShunsukeTamura06](https://github.com/ShunsukeTamura06)),
  solo, in open development. Issues and PRs are welcome (Japanese or English).
- **Not** trying to be: a multi-tenant SaaS, an authentication
  framework, or an Office Scripts replacement.

## Roadmap (next)

See [docs/OSS_LAUNCH_PLAN.md](./docs/OSS_LAUNCH_PLAN.md) for the full plan.

- [ ] Bundle a pre-built `.xlsm` sample with real VBA (currently only
  `.xlsx` ships; `.xlsm` requires running `scripts/inject_vba.ps1` on
  a machine with Excel)
- [ ] Live demo (Hugging Face Space or Render)
- [ ] Docker Compose for one-command setup
- [ ] CI (GitHub Actions: pytest + ruff + mypy + frontend typecheck)
- [ ] UI brand alignment (current sidebar shows the Japanese product
  title; will be unified with `xlblueprint`)
- [ ] Frontend i18n (Japanese UI strings)

## Contributing

The contribution guide lives in [CONTRIBUTING.md](./CONTRIBUTING.md)
(coming soon — Phase 3 of the OSS launch). For now:

- Bug reports and feature requests are welcome via Issues
- For larger changes, please open an Issue first to discuss the
  approach
- Commits follow [Conventional Commits](https://www.conventionalcommits.org/)
- See [CLAUDE.md](./CLAUDE.md) for the agent / collaborator working
  rules used in this repository

## License

[MIT](./LICENSE) — Copyright (c) 2026 Shunsuke Tamura and `xlblueprint`
contributors.

## Acknowledgments

- [openpyxl](https://openpyxl.readthedocs.io/) for spreadsheet parsing
- [oletools](https://github.com/decalage2/oletools) for VBA extraction
- [Nuxt UI](https://ui.nuxt.com/) and [Vue Flow](https://vueflow.dev/)
  for the frontend
- [Anthropic Claude Code](https://claude.com/claude-code) was the
  primary development assistant on this repository — see commit
  history for co-author trailers
