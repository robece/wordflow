# INSTRUCTIONS

## Objective

Maintain `wordflow` as the reusable document-automation toolchain for Word workflows that need:

- deterministic behavior
- staging in `.wordflow/` adjacent to the document
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
| Protection round-trip | [x] | [ ] | `prepare-session` captures protection metadata; `protect` re-applies it — Windows only |

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
| Temp handling | staged in `.wordflow/` adjacent to the document | ad hoc and manual |
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
- `protect --session ...`
- comment review commands

Use when:

- a document family is versioned
- the published output may be protection-sensitive or rewrapped
- you want the cache to live under `.wordflow/`, not inside the `.docx`
- multiple edits should reuse the same normalized working copy
- the source was protected and the final delivery must be protected again

## Command reference

See `README.md` for full command examples, usage patterns, and the feature support matrix. This file focuses on engineering contracts and maintenance rules, not on usage documentation.

Command-specific engineering expectations:

- `normalize`: may use Word COM only for detected `ole-encrypted-package` sources; that exception must never become the default path for general editing or publishing
- `publish-next`: the filename must end in a version marker like `v001`; version incrementing happens only after the candidate is already validated and exactly one destination path has been chosen
- `prepare-session`: the cache belongs to the work session, not to the document; reuse only when source path and source hash match
- `validate-spec`: invalid highlights, unsupported part targets, empty anchors, and missing assets must all fail before any document edit runs
- comment commands: `delete` must remove both the comment entry and its in-document references; `update` must preserve a valid OOXML package after mutation
- `protect`: requires a session created from an encrypted or restricted source; must fail with a descriptive error if no protection metadata exists, if `--password` is missing when required, or if the original had no protection at all; uses `current_version_path` from session metadata as input — never requires `--input`

## Default operating assumptions

The tool should assume all of the following unless proven otherwise:

1. the source document may already be open in Word
2. the source may live under OneDrive
3. the source may be hydration-sensitive
4. the source may be protected, IRM-wrapped, or ambiguous
5. a technically valid `.docx` may still be unacceptable if inherited formatting was lost
6. version churn is worse than a fail-fast stop

## Default runtime workspace

The default runtime workspace is `.wordflow/` created adjacent to the `--input` document:

```
<input-document-directory>/
└── .wordflow/
    ├── wordflow.log                ← always-on structured log (INFO level)
    ├── <stem>.session.json         ← publish audit trail (shared across all versions)
    ├── run-<pid>-<stamp>/          ← ephemeral staging (cleaned on success)
    └── sessions/
        └── <session-id>/           ← persistent normalized working copy
```

If the input has no parent directory (bare filename), `.wordflow/` is created in the current working directory.

The caller may override this with `--temp-root`. The directory is created automatically if it does not exist. `.wordflow/` is excluded from version control via `.gitignore`.

## Default build workspace

Cargo build artifacts are placed in the default `target/` directory within the project. No external path configuration is required. The `.cargo/config.toml` does not override the target directory.

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

## Session concepts

There are two distinct session concepts inside `.wordflow/` — do not confuse them:

**1. Publish audit log** — `<stem>.session.json`

A timestamped record of every successful `publish` and `publish-next` call. Written by `track_session` in `session.rs`. Shared across all versions of the same document family. Survives across tool runs indefinitely.

**2. Work session cache** — `sessions/<session-id>/`

A normalized working copy managed by `prepare-session`. Used by `publish-next --session` to avoid re-normalizing the source on every run. Keyed by session ID. Reused when the source path and hash match; refreshed when the source changes.

The cache belongs to the work session, not to the document:

- normalized working copy path
- source hash for reuse decisions
- current published version path for `publish-next`

Do not embed either type of session state into the `.docx` package.

## Protected / IRM workflow rule

If the source is not a true OOXML zip package:

1. stage it in `.wordflow/` adjacent to the source
2. use Word only for conversion
3. export plain text from source
4. convert through `Word -> RTF -> OOXML`
5. export plain text from migrated candidate
6. export Word-observed paragraph signatures from source and migrated candidate
7. **capture protection metadata** (`ProtectionType`, `HasPassword`, IRM status) while the document is open in Word
8. require those text exports to match
9. require the Word-observed paragraph signatures to preserve inherited formatting that should survive conversion
10. validate the migrated OOXML
11. if available, compare fidelity against a trusted OOXML reference
12. only then trust the migrated file as a candidate technical base

Do not run OpenXML editing directly against a protected container.

## Protection round-trip rule

When the source required Word COM to open (encrypted or protection-restricted), the workflow must offer a way to re-apply the original protection after N editing iterations.

Protection metadata is captured once during `prepare-session` and stored in `.wordflow/sessions/<id>/protection.json`. It is never re-captured during editing iterations.

Re-protection rules:

- `protect` must use the `current_version_path` from session metadata — never require the caller to specify `--input`
- if the original had a password, `--password` is required; fail with a clear error if omitted
- if the original had only editing restrictions (no encryption), `--password` is optional
- if the original had IRM, attempt re-application through Word COM; document in the error if the RMS service is unreachable
- `protect` is Windows-only — it must never be called from a Linux/macOS environment

## Build and test rules

The 64 unit tests cover the cross-platform surface (OOXML editing, validation, sessions, publish workflows). They run on any OS:

```bash
cargo test
```

On Windows with the MSVC toolchain:

```powershell
cargo +stable-x86_64-pc-windows-msvc test
```

The `migrate` command and `--allow-word-com-encrypted-package` require Windows with Word installed and are not covered by the unit test suite.

When adding or changing workflow behavior:

1. add or update tests in `src/lib.rs` or `src/session.rs`
2. rerun the crate tests
3. update `README.md` — feature support matrix, command descriptions, and runtime workspace docs
4. update this file if engineering contracts or maintenance rules changed

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
