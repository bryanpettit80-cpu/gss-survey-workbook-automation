# Deferred updates and verification

These items remain outside v1.2.1. They are requirements for later work, not
features delivered or checks completed by the documentation cleanup. Reporting
continues through the manual launcher; no scheduling or mail activation is
included.

## Zero-response results

Current behavior preserves vendor values for source identification but stops
publication if any retained population row has zero responses. Blank scores and
vendor numeric zeros both reach this gate. The workbook currently cannot safely
represent those rows as unscored across its formulas and reports.

A future change must meet all of these criteria:

- Keep response count zero, the source values, and source identity intact. Show
  the measured score as unavailable or unscored; never present zero responses as
  a measured zero-percent result. A measured zero with a positive count remains
  valid.
- Carry the unscored state through raw-data helpers, dashboards, comparison
  tables, charts, PDFs, and package analysis. Suppress comparisons, rankings,
  and statistical interpretations that require a missing score. Do not fill
  the gap by averaging other restaurants or earlier observations.
- Check each metric's available denominator. A positive overall count does not
  establish a response to a conditional question; do not invent missing
  question-level denominators.
- Cover a current zero-response row, historical zero-response rows, one
  restaurant with zero responses, and scored-to-unscored and unscored-to-scored
  comparisons. Include blank vendor scores, numeric vendor zeros, and a genuine
  measured zero with a positive count.
- Run independent value assertions and isolated desktop Excel recalculation,
  then inspect all affected workbook views and PDFs for false zeros, broken
  formulas, misleading trends, and missing labels. Confirm correction,
  duplicate detection, recovery, and package compatibility still behave as
  expected before lifting the existing publication gate.

## Current-format snapshot restoration

`gss-run-snapshot/v1` stores a selected run inventory with original relative paths
and byte hashes. `VerifyRun` verifies those files. Transaction recovery can
restore the previous report or retry independent backup. These operations do
not provide an automated full-folder restore. The existing quarterly restore
drill applies to the older commissioned backup format.

Future restore tooling and its drill must meet these criteria:

- Declare the recovery scope before copying. Restore every manifest-listed
  file and report absent prerequisites. A run snapshot excludes some GSS
  content, including accumulated packages and the current program tree; do not
  call it a complete folder backup. Define any supplemental inventory needed
  for full operational recovery, including correction assignments, transaction
  state, required source history, the program release, and immutable package
  evidence.
- Verify manifest identity, every size and hash, required roles, and both
  stored and restored path containment. Reject missing files, corrupt files,
  unsafe paths, and colliding destinations before changing an existing target.
- Default to a new isolated destination with a reviewable restore plan. Never
  overwrite the live folder during a drill. Preserve original snapshots,
  package evidence, and receipts; record any path rebinding or recovery action
  in a new receipt rather than rewriting historical evidence.
- Make interruption and retry safe. Record restored and pending files, verify
  completed copies, and stop if a destination changed independently. Require
  a separate reviewed cutover before a restored folder becomes operational.
- Exercise complete, incomplete, corrupt, duplicate-path, interrupted, and
  repeated restores. Restore under a different root, verify every included
  file, and run isolated desktop Excel checks for dates, all four entities,
  counts, scores, comparisons, formulas, and PDF presentation.
- Check correction choices, detail assignments, pending-run reconciliation,
  and existing draft-consumer compatibility in the restored environment.
  Do not infer that old absolute package paths work under a new root; any
  migration needs its own compatibility evidence and must preserve immutable
  originals. No drill may create drafts or send mail.
- Produce a receipt with the snapshot and program identity, coverage and
  exclusions, copied-file counts and hashes, Excel/PDF results, unresolved
  holds, and whether any live files changed. State local byte verification and
  remote cloud-upload verification separately. Completion requires the stated
  scope to pass; a partial restore must remain explicitly partial.

## Second-computer verification

This is a separate acceptance check, not a claim established by portable-path
unit tests or a second folder on the original computer. It remains unverified.

On another Windows computer with desktop Excel, download the versioned program
and required files, verify installation without Git or old test receipts, and
configure that computer's backup destination. Use an isolated copy first and
record the computer, Windows/PowerShell/Excel versions, program release, setup
result, and file-availability checks.

Validate a normal update, an already-current repeat, correction-choice reuse,
backup failure and retry, and interrupted-run recovery. Inspect the resulting
workbook and PDF. For sequential handoff, finish the first computer's run and
cloud synchronization before starting the second; never run both writers
together. Record any unresolved package-path or cloud-sync limitation. This
check must not enable a schedule or mail workflow.
