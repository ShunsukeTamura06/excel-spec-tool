# CLAUDE.md — Agent working rules

This file tells Claude Code (and other coding assistants) how to work on
this repository. The product spec is in [SPEC.md](./SPEC.md), the OSS
launch plan in [docs/OSS_LAUNCH_PLAN.md](./docs/OSS_LAUNCH_PLAN.md).
This file describes the *process*, not the product.

## 0. Core principles

- **SPEC.md is the source of truth.** Do not implement anything that
  contradicts it. If you need to change behavior in SPEC.md, update
  SPEC.md first and call it out.
- **Respect the layering:** `frontend → backend → core` is the only
  allowed direction.
  - `core/` must not import `backend/` or `frontend/`
  - `backend/` must not import `frontend/`
- **Work in small, reviewable steps.** Pause and confirm with the user
  before large or unexpected changes (see §7).
- **When in doubt, ask.** Do not silently reinterpret ambiguous spec.

## 1. Toolchain

- Python 3.10+
- Package manager: `uv` (`uv sync` / `uv add` / `uv run`)
- Formatter: `ruff format`
- Linter: `ruff check`
- Type checker: `mypy --strict` against `core/`
- Tests: `pytest`
- Frontend: Node 20+, `pnpm` (via corepack), Nuxt 3

Setup:
```bash
uv sync --group dev
corepack enable
cd frontend && pnpm install
```

Common commands:
```bash
uv run pytest               # all backend + core tests
uv run ruff check
uv run mypy core
cd frontend && pnpm dev     # Nuxt dev server
```

## 2. Repository layout (current)

```
xlblueprint/
├── core/             # Pure Python: extraction / references / spec / diagrams
├── backend/          # FastAPI routes + storage + LLM client
├── frontend/         # Nuxt 3 SPA (TypeScript)
├── tests/            # pytest (core / backend)
├── scripts/          # Utilities (sample generator, VBA injection)
├── docs/             # OSS launch plan and design docs
├── SPEC.md           # Product spec (source of truth)
├── CLAUDE.md         # This file
├── LICENSE           # MIT
└── pyproject.toml
```

Detailed module responsibilities live in [SPEC.md](./SPEC.md) §2 and §4–6.

## 3. Allowed / disallowed

### Allowed
- Add small test fixtures (`tests/fixtures/`). Generated dummy data only.
- Make reasonable implementation-level decisions not covered by SPEC.md.
  Note the decision in the commit message.
- Refactor within the scope of the current task.

### Disallowed
- Add **features** that are not in SPEC.md without asking
- Implement authentication, multi-tenancy, or RBAC (out of scope for
  the public release; can be added by adopters downstream)
- Connect to external APIs that send workbook content outside of the
  configured LLM endpoint
- Land oversized PRs that span many unrelated changes
- Rewrite large existing files without flagging the impact first

## 4. Tests

- Every `core/*.py` has a matching `tests/core/test_*.py`
- Backend tests use `TestClient` with `tmp_path` for `JOBS_DIR`
- LLM calls are mocked via `MockLLMClient` (`backend/llm_client.py`)
- Target ~80% coverage (soft target, not enforced)

Tests must pass before any commit. Pre-existing flaky test
`test_directory_permissions_700` is known to fail on Windows (POSIX
mode checks); ignore unless explicitly fixing.

## 5. Coding conventions

- Type hints on every function
- Google-style docstrings on public functions
- Use the `logging` module — `print` is forbidden in library code
- Wrap exceptions at layer boundaries: `core` raises `CoreError`
  subclasses (`core/exceptions.py`); `backend` translates to
  `HTTPException`
- Read configuration from environment variables (with defaults), never
  hardcode

## 6. Commits

- **1 commit = 1 logical change**
- Conventional Commits style in Japanese or English:
  - `feat: add VBA extractor`
  - `fix(ui): preview row headers were off by one`
  - `test: add tests for reference index`
  - `docs: clarify deployment notes`
- Co-author trailer encouraged when authoring with Claude:
  ```
  Co-Authored-By: Claude <noreply@anthropic.com>
  ```

## 7. When to stop and ask

Pause and ask the user before:

- Doing something that contradicts SPEC.md
- Implementing a feature not in SPEC.md
- Making large design decisions (DB choice, swapping a major dependency)
- Anything irreversible (force push, history rewrite, dropping a feature)

Question template:
```
[Question] {title}
Situation: {what you were doing, where it got stuck}
Related SPEC.md section: {if any}
Options:
  A) ...
  B) ...
  C) ...
Recommendation: {} ({why})
```

## 8. Known gotchas

| Area | Gotcha | Workaround |
|---|---|---|
| `extract_vba` | password-protected VBA projects | skip, leave TODO |
| `extract_workbook` | `.xls` is not openpyxl-compatible | return empty `sheets=[]`, log warning |
| `extract_workbook` | external link extraction uses `wb._external_links` (private API) | wrap in `try/except`, version-pin awareness |
| `reference_index` | VBA regex never catches everything | aim for coverage of common patterns; risk_analyzer surfaces the gap |
| Frontend `File` → backend | binary upload | `FormData()` with `append("file", file)`, pass as Nuxt `$fetch` body |
| Large `.xlsm` | memory / time | default upload cap is 50 MB / ~5000 rows; warn beyond that, override via `MAX_UPLOAD_BYTES` |

## 9. OSS launch context

- The project is preparing for public release on GitHub.
- Default language for new documentation is English. A Japanese
  counterpart (`*.ja.md`) can be added in parallel.
- Avoid embedding company-specific terminology, customer names, or
  internal URLs. The codebase should read as a self-contained OSS
  project.
- If you spot residual internal references during a task, flag them
  in the report (don't silently leave them).

## 10. Definition of done

- [ ] Implementation file(s) committed
- [ ] Matching tests committed; `pytest` is green
- [ ] `ruff check` is green
- [ ] `mypy --strict` is green for `core/`
- [ ] Manual smoke check noted in the commit message when relevant
- [ ] SPEC.md remains consistent
- [ ] Status reported to the user; next step confirmed
