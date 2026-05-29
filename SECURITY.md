# Security Policy

## Reporting a vulnerability

**Please do not open a public GitHub Issue for security problems.**

Use [GitHub Security Advisories](https://github.com/ShunsukeTamura06/xlblueprint/security/advisories/new)
to report privately. If for any reason you cannot use that flow, email
the maintainer directly: `shunsuke.tamura06@gmail.com`.

Acknowledgement: within 7 days. A fix or mitigation plan: best effort,
typically within 30 days for high-severity issues. This is a
solo-maintained OSS project — turnaround is not guaranteed but
reports will be taken seriously.

## What's in scope

This project parses untrusted Excel files and exposes a local HTTP API
that an LLM drives. The following are explicitly in scope for security
reports:

### File parsing (highest concern)

- **Malicious `.xlsx` / `.xlsm` / `.xls`**: anything that causes
  `xlblueprint` to crash unbounded, consume excessive memory or CPU,
  or escape the parser (e.g. XXE, zip-slip, path traversal via cell
  content or VBA strings)
- **`oletools` / `openpyxl` upstream issues** triggered by our usage
  patterns (please report to those projects too; we'll mirror the
  advisory)

### HTTP backend (`backend/`)

- **Path traversal** via `job_id` or file uploads (validated as UUIDv4,
  but report any bypass)
- **Server-side request forgery** in URL handling
- **Information disclosure** through error messages or logs that
  weren't supposed to surface (e.g. masked Power Query connection
  strings leaking)
- **Resource exhaustion** (large uploads, long-running tool loops not
  bounded). `MAX_UPLOAD_BYTES` and `MAX_TOOL_ITERATIONS` exist as
  guardrails — report bypasses

### LLM tool execution (`backend/llm_tools.py`)

- **Prompt injection** in workbook content (sheet names, formulas,
  cell text) that successfully causes the chat assistant to call
  tools beyond the registered safe set, or that exfiltrates cell
  values the user did not request
- Note: prompt injection that only changes the *wording* of the
  assistant's response (without unauthorized tool use) is a UX issue,
  not a security one

### Frontend (`frontend/`)

- **XSS** through workbook content rendered in the UI without escaping
- **CSRF** against the local backend (we currently rely on `CORSMiddleware`
  + same-origin assumptions — report bypasses)

## What's NOT in scope

- **The Excel application itself**. We do not execute VBA; we only
  parse its source. If a workbook is malicious *when opened in Excel*,
  that's a Microsoft Excel concern, not ours.
- **Your LLM provider**. If your configured LLM endpoint leaks data,
  rate-limits you, or hallucinates dangerous advice, that's between
  you and the provider.
- **Production deployment hardening** that isn't part of the default
  configuration (TLS termination, reverse proxy, authentication —
  these are deliberately out of scope; see `docs/architecture.md`).
- **Issues only reachable with admin-level access to the host**
  (writing to `JOBS_DIR`, modifying environment variables, etc.).

## Disclosure

We follow coordinated disclosure: please don't post details publicly
until a fix is released or 90 days have passed (whichever first).
After a fix lands, the advisory will be published and you'll be
credited in the changelog if you'd like.

## Supported versions

The project is pre-1.0 and only the latest `main` is supported.
Security fixes will be released as new patch versions.
