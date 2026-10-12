# Regression scenarios

The tests compare business values, dates, entities, and report presentation.
Hashes establish that staged, published, and recovery copies are intact; they do
not define whether two vendor workbooks mean the same thing.

| Scenario | Expected result | Coverage |
| --- | --- | --- |
| Richmond overall 70, previous 75, prior year 80 | Current 70; changes -5 and -10 percentage points | Reporting unit suite and desktop Excel scenario |
| Richmond dissatisfaction 5, previous 8, prior year 9 | Current 5; improvements +3 and +4 points | Reporting unit suite and desktop Excel scenario |
| Corrected overall 72 and dissatisfaction 4 | Changes -3/-8 and +4/+5; exactly four rows replaced | Desktop Excel scenario |
| Vendor company aggregate differs from restaurant average | Retain the supplied company result | Reporting unit suite |
| Two missed report weeks and missing prior-year counterparts | Import four complete periods chronologically: 16 rows | Desktop Excel scenario and required-period unit cases |
| Source or main workbook renamed; identical source copied | Recognize content and return no business changes | Path, source-choice, and desktop Excel scenarios |
| Different versions of one reporting period | Ask for a specific revision; retain the prior report until chosen | Reporting and desktop Excel scenarios |
| Previously accepted correction appears again | Reuse its content decision; an unknown revision needs a new choice | Source-choice and desktop Excel scenarios |
| Missing entity, duplicate entity, incorrect period, invalid count or percentage | Block publication with a reason | Logic and reporting suites |
| Zero-response row with blank scores | Keep the source unscored; stop workbook publication because the legacy model can display false zeros; preserve the previous report | Reporting compatibility-gate regression; zero-response workbook rendering is explicitly unsupported |
| Excel error inspection itself fails | Fail validation rather than report zero errors | Reporting unit suite |
| Interruption before, during, or after workbook/PDF promotion | Recover from the journal, or stop on unrelated changes | Transactions and operations suites |
| Local recovery snapshot fails | Retain recovery state before allowing another publication | Operations suite |
| Independent backup is unavailable | Keep the valid report; expose pending backup and permit retry | Operations suite |
| Partial copy, repeated retry, or corrupted committed backup | Resume safe copies; remain idempotent; reject corruption | Operations and legacy backup suites |
| Another data root or extracted program without Git/history | Resolve relative evidence; validate installation without Git | Operations, paths, and release-integrity suites |
| Guest detail lacks a reliable report assignment | Leave draft preparation waiting without invalidating aggregates | Email-package suite |
| Corrected guest detail overlaps an original | Require explicit supersession; avoid duplicate surveys | Email-package suite |
| Frozen package's original files are renamed or resaved | Preserve the package's own evidence and identity | Email-package suite; optional existing AOL parser |

The desktop Excel runner is deliberately separate from GitHub-hosted CI:

```powershell
.\tests\Test-GSS-ReportingIntegration.ps1 `
  -TemplatePath 'D:\Isolated QA\GSS Score Trends - Main.xlsx' `
  -EvidenceRoot 'D:\Isolated QA\Results'
```

It creates a new isolated folder, replaces the copied template's raw observations
with synthetic values, and promotes only within that folder. The template is not
included in this repository. It also checks that catch-up and corrections leave
the model's Raw_Data formula references unchanged.

Inspect both generated comparison PDFs for correct periods, both restaurants,
visible values and signs, readable labels, and clipping. A second physical
computer remains a separate acceptance test. A successful local DriveFS copy
does not establish completed remote cloud synchronization.
