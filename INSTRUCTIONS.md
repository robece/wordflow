# INSTRUCTIONS

## Objective

Maintain `wordflow` as the reusable document-automation toolchain for Word workflows that need:

- deterministic behavior
- staging in `C:\Temp`
- fail-fast publish rules
- explicit validation
- explicit migration handling
- explicit fidelity checks for inherited formatting

This project is not just a low-level XML editor. It is the beginning of a workflow runner for safe Word document generation and update.

## What this workspace owns

This workspace owns:

- the Rust CLI
- the OpenXML editing engine
- staged publish behavior
- staged migration behavior
- source-fidelity checks
- example specs
- test coverage for the supported workflow contract

This workspace does **not** own:

- customer deliverables
- business document content
- long-lived temporary artifacts
- document-family rules that belong in downstream project instructions

## Source-of-truth files

| File | Owns |
|---|---|
| `README.md` | user-facing explanation, command usage, scenarios, benefits |
| `INSTRUCTIONS.md` | engineering intent, maintenance rules, workflow guardrails |
| `src\main.rs` | CLI surface |
| `src\lib.rs` | engine, workflow logic, validation logic, tests |
| `examples\*.json` | example operation specs |

## Documentation contract

Keep these three surfaces aligned whenever capability changes:

1. `README.md` feature support matrix
2. `examples\*.json` sample specs
3. tests in `src\lib.rs`

If a feature is only partial, the README matrix must say so explicitly instead of implying full support.

## Feature support contract

| Feature area | Supported | Partial | Notes |
|---|---|---|---|
| Paragraph editing and styles | [x] | [ ] | insertion, replacement, deletion, headings, bullets, numbering, quote, visible markup |
| Simple tables | [x] | [ ] | row/cell insertion with style and highlight |
| Complex tables | [ ] | [x] | merged cells, exact widths, nested tables, and inherited layout surgery remain partial |
| Hyperlinks, images, section breaks | [x] | [ ] | relationship-bearing structural edits are part of the supported contract |
| Comments and review markup | [x] | [ ] | add/list/update/delete comments plus tracked insert/delete |
| Notes, content controls, fields | [x] | [ ] | footnotes/endnotes, content controls, and complex fields |
| Header/footer and core properties | [x] | [ ] | reusable document-structure updates |
| Validation, inspection, sessions, publish | [x] | [ ] | normalization, validation, fidelity checks, staged publish, reusable sessions |
| Protected-source migration | [ ] | [x] | supported workflow, but still intentionally depends on a guarded Word-backed conversion path |

## Why this tool matters

The main value of this project is not only that it can modify XML.

The real value is that it turns a fragile Word workflow into a stricter pipeline:

1. stage source outside OneDrive friction
2. use a single candidate output
3. validate candidate before publish
4. reject bad candidates
5. keep the final destination untouched until success

Without this tool, the workflow becomes much more manual and less reliable:

- more Word COM scripting
- more one-off PowerShell
- more file-lock surprises
- more cloud/hydration issues
- more chances to lose formatting while keeping the text
- more operator involvement in routine steps

## With vs without the tool

| Dimension | With `wordflow` | Without `wordflow` |
|---|---|---|
| Temp handling | staged in `C:\Temp` by default | ad hoc and manual |
| Candidate control | one candidate, explicit validation | often multiple trial outputs |
| Publish behavior | destination touched only at the end | easy to touch destination too early |
| Migration path | explicit `migrate` command | Word COM or manual conversion |
| Format preservation checks | explicit `check-fidelity` support | usually implicit or manual |
| Failure handling | fail-fast, preserve temp for debugging | retry-heavy, harder to diagnose |
| Reuse | same CLI and same rules across docs | custom scripts and operator memory |

## Core workflow model

There are five major workflow layers in this tool:

### 1. Source normalization layer

Commands:

- `inspect`
- `normalize`

Use when:

- the source may come from mixed origins
- you need a quick readiness check before OpenXML editing
- you want a deterministic decision about whether the file is already normalized

### 2. Direct editing layer

Commands:

- `insert`
- `insert --dry-run`

Use when:

- the input is already trusted OOXML
- you are operating at a low level
- you do not need full publish orchestration

### 3. Safe publish layer

Commands:

- `publish`
- `publish-next`

Use when:

- you are creating a final output document
- the destination must not be touched until the candidate is valid
- you want temp staging and fail-fast behavior

### 4. Safe migration layer

Command:

- `migrate`

Use when:

- the source may not be OOXML
- Word is needed for conversion only
- you need text export comparison and optional fidelity comparison before trusting the migrated base

### 5. Reusable session layer

Commands:

- `prepare-session`
- `publish-next --session ...`
- comment review commands

Use when:

- a document family is versioned
- the published output may be protection-sensitive or rewrapped
- you want the cache to live under `C:\Temp`, not inside the `.docx`
- multiple edits should reuse the same normalized working copy

## Command reference

### `insert`

Low-level editing command.

```powershell
cargo +stable-x86_64-pc-windows-msvc run -- insert `
  --input "C:\path\source.docx" `
  --output "C:\path\updated.docx" `
  --spec ".\examples\document-update.json"
```

### `insert --dry-run`

Low-risk validation of a spec without writing the output.

```powershell
cargo +stable-x86_64-pc-windows-msvc run -- insert `
  --input "C:\path\source.docx" `
  --output "C:\path\ignored.docx" `
  --spec ".\examples\document-update.json" `
  --dry-run
```

### `normalize`

Normalization preflight command.

```powershell
cargo +stable-x86_64-pc-windows-msvc run -- normalize `
  --input "C:\path\source.docx" `
  --output "C:\path\normalized.docx"
```

Current expectation:

- if the source is already a real OOXML zip package, `normalize` validates it and publishes a clean OOXML-safe copy
- if the source is legacy, protected, or otherwise non-normalized, `normalize` must stop with an explicit explanation unless a reliable normalization backend is available
- a guarded Word COM exception may be allowed only for detected `ole-encrypted-package` sources, and only inside `normalize`
- that COM exception must never become the default path for general editing or publishing

### `publish`

Preferred command for final document generation.

```powershell
cargo +stable-x86_64-pc-windows-msvc run -- publish `
  --input "C:\path\source.docx" `
  --output "C:\path\final.docx" `
  --spec ".\examples\document-update.json"
```

### `publish-next`

Version-aware publish command.

Direct input:

```powershell
cargo +stable-x86_64-pc-windows-msvc run -- publish-next `
  --input "C:\path\document-v015.docx" `
  --spec ".\examples\document-update.json"
```

Reusable session:

```powershell
cargo +stable-x86_64-pc-windows-msvc run -- publish-next `
  --session "my-doc-session" `
  --spec ".\examples\document-update.json"
```

Current expectation:

- the current filename must end in a version marker like `v001`
- the next output should be derived from the highest existing matching version in the output directory
- the default target mode should keep creating a new published version
- an explicit `latest` target mode may update the latest published version in place when the caller intentionally wants to continue working without incrementing
- when a session is used, the cached normalized working copy should advance to the newly validated content

### `prepare-session`

Reusable working-session command.

```powershell
cargo +stable-x86_64-pc-windows-msvc run -- prepare-session `
  --input "C:\path\source.docx" `
  --session "my-doc-session"
```

Current expectation:

- the cache belongs to the temp session, not to the document
- the normalized working copy should be reused when the source path and source hash match
- if the source changed, the session should refresh the normalized working copy

### `validate-spec`

```powershell
cargo +stable-x86_64-pc-windows-msvc run -- validate-spec `
  --spec ".\examples\document-update.json"
```

Current expectation:

- invalid highlight values must fail before document editing starts
- unsupported part targets must fail before document editing starts
- empty anchors and missing assets must fail before document editing starts

### Comment review commands

Commands:

- `add-comment`
- `list-comments`
- `update-comment`
- `delete-comment`

Current expectation:

- comments should support a reviewer workflow, including GitHub Copilot CLI as a reasonable default reviewer identity
- delete must remove both the comment entry and its in-document references
- update must preserve a valid OOXML package after mutation
- list must expose enough information for a caller to understand comment ids, authors, and where the comment is attached

### `migrate`

Preferred command for protected or ambiguous source documents.

```powershell
cargo +stable-x86_64-pc-windows-msvc run -- migrate `
  --input "C:\path\source.docx" `
  --output "C:\path\migrated.docx"
```

With trusted OOXML reference:

```powershell
cargo +stable-x86_64-pc-windows-msvc run -- migrate `
  --input "C:\path\source.docx" `
  --output "C:\path\migrated.docx" `
  --trusted-ooxml "C:\path\trusted-base.docx"
```

### `validate`

```powershell
cargo +stable-x86_64-pc-windows-msvc run -- validate --input "C:\path\source.docx"
```

### `check-fidelity`

```powershell
cargo +stable-x86_64-pc-windows-msvc run -- check-fidelity `
  --before "C:\path\trusted-base.docx" `
  --after "C:\path\candidate.docx"
```

With spec:

```powershell
cargo +stable-x86_64-pc-windows-msvc run -- check-fidelity `
  --before "C:\path\trusted-base.docx" `
  --after "C:\path\candidate.docx" `
  --spec ".\examples\document-update.json"
```

### Inspection/debug commands

```powershell
cargo +stable-x86_64-pc-windows-msvc run -- inspect --input "C:\path\source.docx"
cargo +stable-x86_64-pc-windows-msvc run -- list-parts --input "C:\path\source.docx"
cargo +stable-x86_64-pc-windows-msvc run -- find-anchors --input "C:\path\source.docx" --text "Strategic principles" --mode equals --occurrence 2
cargo +stable-x86_64-pc-windows-msvc run -- diff-docx --before "C:\path\old.docx" --after "C:\path\new.docx"
```

## Default operating assumptions

The tool should assume all of the following unless proven otherwise:

1. the source document may already be open in Word
2. the source may live under OneDrive
3. the source may be hydration-sensitive
4. the source may be protected, IRM-wrapped, or ambiguous
5. a technically valid `.docx` may still be unacceptable if inherited formatting was lost
6. version churn is worse than a fail-fast stop

## Default runtime workspace

Use:

`C:\Temp\wordflow`

as the default runtime workspace for:

- publish
- publish-next
- migrate
- session metadata and normalized working copies
- intermediate candidate generation
- temporary validation artifacts

If the folder does not exist, the workflow should create it.

## Default build workspace

Keep Cargo build artifacts and the runnable executable out of the source tree.

Use:

`C:\Temp\wordflow-target`

as the default Cargo target directory so the repository under OneDrive remains source-only during normal `cargo build`, `cargo run`, and `cargo test` usage.

## Required engineering behaviors

### Fail-fast

The tool must fail fast for structural problems such as:

- invalid OOXML package
- missing anchor
- failed migration
- text export mismatch
- fidelity mismatch
- inability to publish the validated output
- inability to normalize a non-ready source with a reliable backend

It must **not** try to solve these with repeated automatic retries that can churn versions or repeatedly touch the final destination.

### Single candidate rule

For a publish-style workflow:

1. stage the source
2. build one candidate
3. validate it
4. publish it only if it passes

Do not materialize multiple numbered outputs while attempting to recover from failures.

For `publish-next`, version incrementing is acceptable only after the candidate is already validated and the command has chosen exactly one destination path.

### Destination safety

The destination path should be treated as the last step, not an intermediate workspace.

The tool should never write to the final destination before:

- technical validation passes
- workflow-specific checks pass
- fidelity checks pass where required

### Temp cleanup policy

- on success: clean temp workspace
- on failure: preserve temp workspace for debugging

This is intentional behavior, not accidental residue.

Reusable sessions are different:

- session metadata and the normalized working copy should persist across successful runs
- ephemeral candidate workspaces created during a session-backed publish should still be cleaned on success

## Format fidelity rule

Text preservation alone is not enough.

The workflow must also guard against losing inherited formatting from untouched paragraphs, especially:

- highlight colors
- paragraph styles
- approved visual markup that should survive into the next version

That is why `check-fidelity` exists, and why `publish` uses fidelity checks before final publish.

## Source readiness rule

Every document workflow should begin by checking whether the source is already normalized for OpenXML automation.

- `inspect` should report the detected packaging
- `normalize` should pass through already-normalized OOXML sources
- `normalize` may use Word COM only as an explicit exception for `ole-encrypted-package` sources
- editing and publish flows should stop early on non-normalized sources instead of failing later with low-level zip errors
- non-normalized inputs should be treated as a workflow state, not as an unexpected crash

## Session cache rule

The cache belongs to the work session, not to the document.

Cache examples:

- normalized working copy path
- source hash for reuse decisions
- current published version path for `publish-next`

Do not embed session state into the `.docx` package just to accelerate the workflow.

## Protected / IRM workflow rule

If the source is not a true OOXML zip package:

1. stage it in `C:\Temp`
2. use Word only for conversion
3. export plain text from source
4. convert through `Word -> RTF -> OOXML`
5. export plain text from migrated candidate
6. export Word-observed paragraph signatures from source and migrated candidate
7. require those text exports to match
8. require the Word-observed paragraph signatures to preserve inherited formatting that should survive conversion
9. validate the migrated OOXML
10. if available, compare fidelity against a trusted OOXML reference
11. only then trust the migrated file as a candidate technical base

Do not run OpenXML editing directly against a protected container.

## Build and test rules

Use:

```powershell
cargo +stable-x86_64-pc-windows-msvc test
```

Cargo should already be configured to place `target` output under `C:\Temp\wordflow-target`.

If a caller must override that location explicitly:

```powershell
$env:CARGO_TARGET_DIR = 'C:\Temp\wordflow-target'
cargo +stable-x86_64-pc-windows-msvc test
```

When adding or changing workflow behavior:

1. add or update tests
2. rerun the crate tests
3. update `README.md`
4. update this file if workflow expectations changed

## Change rules

When modifying this project:

1. keep the tool generic where possible
2. keep document-family rules out of the engine unless they are truly reusable
3. preserve CLI clarity; do not add commands with overlapping or ambiguous responsibilities
4. prefer explicit errors over silent fallback
5. preserve namespaced OOXML attributes exactly
6. do not add retry loops that undermine fail-fast behavior
7. document new workflow assumptions immediately
8. keep version-orchestration logic generic and convention-based, not hardcoded to one document family

## What to optimize for

Optimize for:

- deterministic workflows
- low operator involvement
- explicit validation
- reusable orchestration
- safer final publish behavior

Do not optimize for:

- clever retries
- silent recovery
- touching the final destination early
- hiding failure details

## Success criteria

This tool is in a good state when:

1. a collaborator can understand the commands and when to use each one
2. the final destination is touched only after candidate validation
3. migration is treated as a controlled workflow, not a one-off manual trick
4. inherited formatting loss is caught before publish
5. the tool reduces operator intervention compared with an ad hoc Word/PowerShell workflow
6. the README matrix and sample specs accurately reflect the real tested capability set
