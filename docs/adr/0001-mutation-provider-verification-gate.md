# ADR 0001: Separate mutation providers from verification policy

- Status: Accepted
- Date: 2026-07-16

## Context

xlblueprint already had narrow openpyxl-based fixes and a structural diff,
but the success of a workbook writer was not contractually separated from the
decision that its output matched the intended change. Integrating more capable
writers such as OfficeCLI makes that distinction essential: a successful
process exit is not evidence that a workbook is safe.

## Decision

- Represent intent as a provider-independent `MutationPlan`.
- Put workbook writers behind the `MutationProvider` contract and require them
  to produce a separate artifact instead of modifying the uploaded original.
- Compute an expected normalized structural diff before mutation, fully
  re-extract the result, and compare the observed diff in xlblueprint's policy
  gate.
- Persist the plan, provider identity and version, source and result hashes,
  expected and observed diffs, and the verdict as one audit record.
- Keep OfficeCLI optional and process-isolated. Publish and enforce only the
  operations confirmed by the adapter's capability contract.

## Consequences

- Additional workbook writers can be added without duplicating the meaning of
  verification or audit evidence.
- openpyxl and OfficeCLI implementations can be exercised against the same
  provider contract.
- Structural agreement cannot prove Excel's dynamic behavior. Changes with a
  blast radius or unresolved high-risk items remain `needs_review`, and COM
  recalculation evidence must later join the same policy chain.
- OfficeCLI's licensing and distribution do not become mandatory xlblueprint
  dependencies.
