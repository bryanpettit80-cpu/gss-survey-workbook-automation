# Future coordinator interface

This document specifies a future integration. It does not install a scheduler,
scan mail, download attachments, send notifications, or create mail drafts.

## Responsibility boundary

GSS owns the reporting period, expected restaurant/entity set, source validation,
duplicate detection, corrections, calculations, rendering, and report packaging.
The shared coordinator owns attachment delivery, notification routing, and
management drafts. AOL is authoritative for read/unread state; Gmail mirrors AOL
and must never update AOL state in reverse. Gift-card Micros intake remains with
its existing automation.

The current external draft consumer reads `_automation_runs/email_outbox` and
`gss-email-package/v2`. Keep those paths and the immutable package contract until
a coordinated consumer migration has passed compatibility tests. READY means a
package is ready for recipient/content review, not permission to send it.

## Proposed operations

| Operation | Inputs | Result |
| --- | --- | --- |
| Inspect | Target Sunday; delivered attachment inventory with source identity and explicit report-period evidence | Missing, invalid, conflicting, ready, or already-current inputs; separate report and draft readiness |
| Run | Target Sunday and previously resolved correction choices | The same local validation and transaction path used by the manual launcher |
| Resume | Existing run identifier | Complete pending backup or publication without reapplying a committed report |

Inspect must not write business files or infer completeness from the newest file
modification time. An attachment's latest guest visit date does not establish its
report assignment. The coordinator supplies delivery evidence; GSS determines
whether the assigned reporting set is complete.

## Result contract

A versioned result will contain:

- project, reporting period, logical business revision, and run identifier;
- independent report, analysis, package, and backup statuses;
- stable actionable reason codes and the missing/conflicting input identities;
- output paths relative to the configured data root;
- the immutable package-manifest reference, if a package is ready;
- previous-result and supersession references for genuine corrections.

Identical business inputs reuse the existing result. Transport byte hashes protect
the actual attachment payloads; they do not define a new business revision merely
because Excel saved a file or an operator renamed it. A changed business revision
is distinct and must identify the revision it supersedes.

## Required scenarios before activation

Contract fixtures must cover complete arrival, missing restaurant/period/detail,
partial download, equivalent duplicate, conflicting correction, already-current
report, interrupted publication, Drive unavailable, and sequential computer
handoff. Repeated coordinator requests must not duplicate reports or drafts.

Notifications may say that a report is complete while its independent backup is
pending. A draft must remain withheld when its own required evidence is missing.
Never label local DriveFS verification as confirmed remote upload.

Activation is a later task. Current reporting is manually launched, and no GSS
schedule is enabled. The obsolete weekly-update task is a removal candidate;
the AOL draft-writer task can remain disabled for possible future use. See the
[task guidance](operations.md#disabled-scheduled-tasks). Neither task's presence
authorizes activation, and a future coordinator must not create a second writer.
