# xlblueprint — Architecture summary

This is a short, English design summary intended for contributors. The
full product spec (much longer, in Japanese) is in
[SPEC.ja.md](./SPEC.ja.md). When the two disagree, the Japanese spec is
authoritative.

## Top-level shape

Three layers with a strict one-way dependency:

```
┌──────────────────────────┐    HTTP    ┌──────────────────────────┐
│  Frontend (Nuxt 3 SPA)   │ ─────────> │  Backend (FastAPI)       │
│  - Upload UI             │ <───────── │  - REST + SSE            │
│  - Spec / diagrams view  │            │  - Storage (filesystem)  │
│  - Chat                  │            │  - LLM client            │
└──────────────────────────┘            └──────────────────────────┘
                                                      │
                                                      ▼
                                          ┌──────────────────────┐
                                          │  Core (pure Python)  │
                                          │  - Extraction        │
                                          │  - Reference index   │
                                          │  - Risk analysis     │
                                          │  - Spec generator    │
                                          │  - Diagrams          │
                                          └──────────────────────┘
```

| Layer | Responsibility | Must NOT |
|---|---|---|
| `core` | Workbook decomposition, reference / risk analysis, spec text generation. Importable as a standalone library. | Touch HTTP, UI, or LLM clients |
| `backend` | Wrap `core` over HTTP, persist jobs, drive the LLM tool-use loop, manage chat history | Contain Excel parsing logic |
| `frontend` | Upload files, render specs / diagrams / chat. Talk only to `backend`. | Parse `.xlsm` itself |

## Module map (`core/`)

| Module | Purpose |
|---|---|
| `extractors/vba.py` | Wrap `olevba` to extract VBA modules and procedures from `.xls{m,b}` |
| `extractors/workbook.py` | Wrap `openpyxl` + read OOXML parts directly for charts, pivots, Power Query, form controls |
| `extractors/cells.py` | Build the SQLite-backed cell value store (`cells.db`) for chat range lookups |
| `models.py` | Pydantic models shared across all layers (`Workbook`, `SheetInfo`, `CellFormula`, `VbaModule`, `Reference`, `AnalysisRisk`, …) |
| `reference_index.py` | Build the inverse index `target_range → list[Reference]`. Combines formula, VBA, chart-series, and pivot-source edges with range-intersection lookup |
| `risk_analyzer.py` | Detect anything static analysis cannot resolve: `INDIRECT` / `OFFSET`, dynamic VBA refs (`Range(var)`, `Worksheets(name).Range(var)`), runtime state (`ActiveSheet`), event macros, external links |
| `external_functions/` | Plugin-style registry of vendor Excel add-in functions. Ships with Bloomberg BDH/BDP/BDS |
| `spec_generator.py` | Render the extracted `Workbook` to Markdown (with Mermaid diagrams), LLM-independent |
| `diagrams.py` | Build sheet-dependency and VBA call-graph models for the frontend (Vue Flow) |
| `exceptions.py` | `CoreError`, `ExtractionError`, `UnsupportedFormatError` |

The deliberate scope choice across `reference_index` / `risk_analyzer`:
**catch the common cases statically; surface the rest as risks**. This
lets the chat layer confidently say "X is impacted" *or* "I cannot
prove X is unaffected — please check manually," instead of silently
guessing.

## Module map (`backend/`)

| Module | Purpose |
|---|---|
| `main.py` | FastAPI app factory, CORS, logging middleware, route registration |
| `storage.py` | Per-job filesystem layout under `JOBS_DIR/<uuid>/` (original / extracted.json / spec.md / references.json / chat_history.jsonl / meta.json / cells.db) |
| `llm_client.py` | `LLMClient` Protocol + `MockLLMClient` + `OpenAICompatibleLLMClient` (via `openai` SDK). `get_default_client()` returns mock if env vars are missing |
| `llm_tools.py` | OpenAI function-calling definitions and dispatcher for 10+ tools (cells / find / refs / VBA / risks / external functions) |
| `annotators.py` | Batched LLM annotation pass: each sheet / procedure gets structured fields (`purpose`, `inputs`, `outputs`, …). Tolerant of parse failures |
| `dependencies.py` | FastAPI DI helpers (`get_storage`, `get_llm_client`) |
| `logging_config.py` | Structured logging with job-id correlation |
| `routes/` | One file per endpoint group (see below) |

## Backend API

| Method | Path | Purpose |
|---|---|---|
| POST | `/extract` | Multipart upload → extraction → `{job_id}` |
| POST | `/analyze/{job_id}` | LLM annotation pass + spec generation |
| GET | `/spec/{job_id}` | Generated Markdown + meta |
| GET | `/workbook/{job_id}` | Structured `Workbook` (sheets / VBA / risks / …) |
| GET | `/diagrams/{job_id}` | Sheet-dependency + VBA call-graph nodes/edges |
| GET | `/references/{job_id}?target=...` | Reverse reference lookup |
| GET | `/cells/{job_id}/...` | Cell value range / search (SQLite-backed) |
| GET | `/external_functions/...` | Registry browse |
| POST | `/chat/{job_id}` (SSE) | Chat with LLM + tool loop, streamed |
| GET | `/chat/{job_id}/history` | Persisted history |
| POST | `/feedback` | Inline thumbs / FAB feedback |
| GET | `/system/llm-status` | Whether the server has real LLM credentials |
| GET | `/health` | Reachability probe |
| GET, DELETE | `/jobs[/{id}]` | List / delete jobs |

## Frontend shape

Nuxt 3 SPA (`ssr: false`) with Pinia for the current-job state.

| Page | Component highlights |
|---|---|
| `/` | `FileDropzone`, `JobCard` list, `LlmOnboardingCard` (when LLM unset) |
| `/spec/[jobId]` | `SpecMetrics`, `SpecInsightDashboard`, tabs: overview / sheets / VBA / external functions / refs / diagrams |
| `/chat/[jobId]` | Multi-conversation chat with progress streaming and `ChatMessageBubble` |

Key composables:
- `useBackend` — typed fetch client (env: `NUXT_PUBLIC_BACKEND_URL`)
- `useDiagramLayout` — wraps `dagre` to lay out the Vue Flow graph
- `useJobStore` (Pinia) — current job id, persisted in `localStorage`

## Tool-use loop

```
chat request ──> backend
                  │
                  ├─ build system prompt (spec.md + job meta)
                  ├─ load recent history (paired user/assistant)
                  ├─ append user message
                  └─ loop up to MAX_TOOL_ITERATIONS (16):
                       │
                       ├─ call LLM with tools
                       ├─ if no tool_calls → emit assistant message, done
                       └─ else: run each tool (cells / refs / VBA / risks / ...)
                                with truncation cap, append tool messages
```

The whole loop runs inside `asyncio.to_thread()` because the
underlying `openai` SDK is synchronous. This keeps the FastAPI event
loop responsive while the LLM is thinking.

## Storage layout

```
JOBS_DIR/
└── <job_id>/                  # job_id = UUIDv4, path-traversal validated
    ├── original.xlsm          # uploaded file
    ├── extracted.json         # Workbook (core.models)
    ├── spec.md                # generated spec
    ├── references.json        # ReferenceIndex
    ├── cells.db               # SQLite of literal cell values
    ├── chat_history.jsonl     # one message per line, append-only
    └── meta.json              # JobMeta
```

Directory permissions are set to 0700 on POSIX. Job cleanup is left to
operators (a cron + `delete` route would be a small add-on).

## What's deliberately out of scope

These are **not** failures to implement — they would be the wrong shape
for a public OSS workbench:

- Auth, RBAC, multi-tenancy
- Generating modified `.xlsm` files (the tool advises; the user edits)
- Multi-file diff
- A managed SaaS

Adopters who want any of those can layer them on top of the FastAPI
backend without modifying `core`.
