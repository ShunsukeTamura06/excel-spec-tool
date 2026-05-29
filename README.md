# xlblueprint

[![CI](https://github.com/ShunsukeTamura06/xlblueprint/actions/workflows/ci.yml/badge.svg)](https://github.com/ShunsukeTamura06/xlblueprint/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](./LICENSE)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![Nuxt 3](https://img.shields.io/badge/Nuxt-3-00DC82.svg?logo=nuxt.js&logoColor=white)](https://nuxt.com/)

> **Understand and safely modernize legacy Excel workbooks (.xlsm / .xlsx)
> with an LLM that knows where every formula and VBA reference leads.**

📖 日本語版: [README.ja.md](./README.ja.md)

`xlblueprint` extracts the structure of a complex business workbook — every
sheet, formula, named range, conditional format, chart, pivot, VBA module,
form control, external link — and turns it into:

1. A **structured spec** (Markdown + Mermaid)
2. An **interactive dependency graph** (sheets & VBA call graph)
3. A **chat assistant** that answers *"if I change X, what breaks?"*
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

### Frontend (Nuxt 3 SPA, dark theme)
- **Insight dashboard**: TL;DR + input/output sheet candidates +
  ranking (most calc / most read / most reader) + risk summary
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
- **518 tests** (pytest), `mypy --strict` on `core/`, `ruff` clean

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

Without an LLM configured (default), the spec / diagrams / search /
risk analysis all work. The chat falls back to mock responses; an
in-app onboarding card explains how to wire up Ollama / OpenAI / etc.

## Architecture

```
Frontend (Nuxt 3 SPA, TypeScript)        Backend (FastAPI)
├── Nuxt UI v3 + Tailwind v4        <──> ├── /extract, /analyze
├── Vue Flow + dagre (diagrams)          ├── /spec, /workbook
├── @nuxtjs/mdc (Markdown rendering)     ├── /references, /diagrams
└── Pinia                                ├── /chat (function calling SSE)
                                         ├── /system/llm-status
                                         └── /jobs

                                         Core (pure Python)
                                         ├── olevba / openpyxl extraction
                                         ├── reference_index
                                         ├── risk_analyzer
                                         ├── external_functions registry
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
  (518 tests, `mypy --strict` on `core/`). LLM-side prompts and UI
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
