# wordflow

Reliable `.docx` automation through OpenXML, with a staged workflow for Word-sensitive documents.

---

## Table of contents

- [The idea behind this](#the-idea-behind-this)
- [What this solves in practice](#what-this-solves-in-practice)
- [wordflow vs pandoc](#wordflow-vs-pandoc)
- [Configuring agents to use wordflow](#configuring-agents-to-use-wordflow)
  - [Priority rule](#priority-rule)
  - [Base instruction block](#base-instruction-block)
  - [Protected document block](#protected-document-block)
  - [System prompt pattern — any AI assistant](#system-prompt-pattern-any-ai-assistant)
  - [Agent skill definition pattern](#agent-skill-definition-pattern)
  - [What to tell the agent explicitly](#what-to-tell-the-agent-explicitly)
- [Why this tool exists](#why-this-tool-exists)
- [What problem it solves](#what-problem-it-solves)
- [What you get by having this tool](#what-you-get-by-having-this-tool)
  - [Operational benefits](#operational-benefits)
  - [Engineering benefits](#engineering-benefits)
  - [Product/process benefits](#productprocess-benefits)
- [How this behaves if we do not have the tool](#how-this-behaves-if-we-do-not-have-the-tool)
- [Feature support matrix](#feature-support-matrix)
- [Project layout](#project-layout)
- [Build and test](#build-and-test)
- [Runtime workspace](#runtime-workspace)
- [Command overview](#command-overview)
- [Command details](#command-details)
  - [1. `insert`](#1-insert)
  - [2. `insert --dry-run`](#2-insert---dry-run)
  - [3. `normalize`](#3-normalize)
  - [4. `publish`](#4-publish)
  - [5. `prepare-session`](#5-prepare-session)
  - [6. `publish-next`](#6-publish-next)
  - [7. `validate-spec`](#7-validate-spec)
  - [8. `migrate`](#8-migrate)
  - [9. `validate`](#9-validate)
  - [10. `check-fidelity`](#10-check-fidelity)
  - [11. Comment review commands](#11-comment-review-commands)
  - [12. Inspection and debugging commands](#12-inspection-and-debugging-commands)
- [Recommended usage patterns](#recommended-usage-patterns)
  - [Safe final update for a normal OOXML source](#safe-final-update-for-a-normal-ooxml-source)
  - [Fast iterative update for a versioned document family](#fast-iterative-update-for-a-versioned-document-family)
  - [Safe migration for a protected source](#safe-migration-for-a-protected-source)
  - [Investigate why a candidate is bad](#investigate-why-a-candidate-is-bad)
- [Spec model](#spec-model)
- [Example spec](#example-spec)
- [Highlighting behavior](#highlighting-behavior)
- [Practical rules](#practical-rules)
- [Using with AI agents](#using-with-ai-agents)
  - [Prompt patterns](#prompt-patterns)
  - [Practical guidance](#practical-guidance)
- [Session history](#session-history)
  - [What the session file contains](#what-the-session-file-contains)
  - [Reading the session — `show-session`](#reading-the-session-show-session)
  - [Rebuilding history from existing files — `reconstruct-session`](#rebuilding-history-from-existing-files-reconstruct-session)
- [Current limitations](#current-limitations)

---

## The idea behind this

Most Word automation works in one direction: you describe what you want, a script or an agent generates a `.docx`, and that is the end of the interaction.

That works for simple cases. But real document workflows rarely work that way.

Documents get reviewed, revised, annotated, shared, corrected, and versioned over time. A document goes through ten iterations, gets comments from multiple reviewers, and accumulates edits across many sessions. Each step matters, and each step creates the risk of losing what was already there.

What was missing was a tool that treats Word documents not as the final output of a one-shot process, but as something that evolves — where humans and agents can keep working on the same document, progressively, with confidence that each step is safe and auditable.

That is what wordflow is.

---

## What this solves in practice

Word automation tends to break in predictable ways:

- The file is open in Word while the script tries to write to it
- The cloud sync provider has not finished syncing the latest version
- The `.docx` looks like a Word file but is actually a legacy container that will not survive an edit
- A failed attempt creates `v009`, `v010`, `v011`, each one wrong
- Formatting survives visually but the underlying structure is broken
- There is no way to tell whether a candidate is actually safe to publish

Instead of working around these problems with more ad hoc scripts, wordflow makes the workflow explicit: normalize the source, validate before editing, apply changes in a reproducible way, check the result, and publish only when everything passes.

That sequence can be driven by a human, by an agent, or by both working together on the same document over multiple sessions.

---

## wordflow vs pandoc

### Purpose and approach

| Dimension | wordflow | pandoc |
|---|---|---|
| **Purpose** | *Safe* automation of existing `.docx` files | Conversion between document *formats* |
| **Approach** | Works on a real `.docx`: inserts, validates, reviews, publishes | Converts from A → B (markdown → docx, docx → pdf, etc.) |
| **Pipeline** | normalize → validate → edit → check-fidelity → publish | input → transform → output |
| **Fidelity** | Verifies the result preserves styles, structure, and comments | Does not verify fidelity — just produces a valid output |
| **Protected documents** | Handles migration of protected sources, sessions, history | Assumes the source is directly convertible |
| **Use cases** | Documents that evolve over multiple sessions with humans and agents | Format migration, one-shot document generation |
| **Document state** | Maintains and updates the same document over time | Discards the source, produces a new file |
| **Validation** | `check-fidelity`, `validate-spec`, `publish` with fail-fast | No post-conversion validation |
| **Collaboration** | Sessions, change history, integrated comments | No concept of session or history |
| **Tool type** | Workflow runner for living documents | Format converter |

### Are they competitors? No — they are complementary.

```
pandoc   → converts any source into a normalized .docx
wordflow → takes that .docx and edits, validates, and publishes it with guarantees
```

### Conceptual difference

**pandoc** answers: *"How do I convert this file from format A to format B?"*

**wordflow** answers: *"How do I safely, auditably, and reversibly modify this existing Word document?"*

---

## Configuring agents to use wordflow

When an agent needs to work with a Word document — to edit it, annotate it, insert a table, add review comments, or publish a new version — the default behavior is often to reach for a custom script, a PowerShell one-liner, or Word COM automation.

This section explains how to configure an agent so that wordflow becomes the first tool it reaches for instead.

The instruction blocks below are interchangeable across **Claude Code** (`CLAUDE.md`), **Qwen Code** (`QWEN.md`), and **GitHub Copilot CLI** (`.github/copilot-instructions.md`). Copy the same content into whichever file your agent reads.

### Priority rule

```
wordflow  →  preferred for all document editing, review, validation, and publishing
Word COM  →  allowed only for migration of protected or non-OOXML sources
ad hoc scripts  →  not acceptable as a replacement for wordflow commands
```

### Base instruction block

Add the following to your agent instruction file (`CLAUDE.md`, `QWEN.md`, or `.github/copilot-instructions.md`):

```markdown
## Word document automation

When working with .docx files, always use wordflow as the primary tool.

Priority:
1. Use wordflow CLI commands for all editing, review, commenting, validation, and publishing.
2. Use Word COM only as a fallback for migration of non-OOXML or protected sources (migrate command).
3. Do not write ad hoc PowerShell or Python scripts to manipulate .docx files directly.

Before editing any document:
- Run normalize to verify the source is a valid OOXML package.
- Run validate-spec to check the spec before applying it.
- Use publish instead of insert for final outputs.

Executable: cargo +stable-x86_64-pc-windows-msvc run --
Spec examples: .\examples\document-update.json, .\examples\review-comments.json
```

### Protected document block

Add this block alongside the base instruction when working with protected or encrypted documents:

```markdown
## Protected documents

When prepare-session reports a Protection line in its output:
1. If the output also shows PasswordRequired  true, ask the user for the
   password before running protect. Keep it in conversation context —
   do not ask again in the same session.
2. Run protect as the last step, after the final publish-next.
3. Report the name and location of the protected file delivered.
4. Do not mention the protect command to the user — just deliver the result.

If prepare-session does not report a Protection line, skip the protect
step entirely without mentioning it.

Password handling:
- Ask for the password only once per session, at the moment it is first needed.
- Never store, log, or repeat the password in any output visible to the user
  beyond what is strictly required to pass it as a --password argument.
- If the user says "same password as before", reuse the one already in context.
```

### System prompt pattern — any AI assistant

For assistants that accept a system prompt or instruction block:

```
When the user asks you to create, edit, update, annotate, or publish a Word document (.docx):

1. Use wordflow as the primary document automation tool.
2. Build a JSON spec that describes the changes (insertions, replacements, deletions, comments, tables, etc.).
3. Run the appropriate wordflow command: normalize → validate-spec → publish.
4. Do not generate file content directly or use Word COM for editing tasks.
5. Only use the migrate command when the source is protected or not a valid OOXML package.
6. If prepare-session reports Protection in its output, run protect as the last step.
   Ask the user for the password only if PasswordRequired  true also appears.

wordflow executable: cargo +stable-x86_64-pc-windows-msvc run --
Spec format reference: see examples/ in the wordflow repository.
```

### Agent skill definition pattern

If you are building an agent skill that wraps wordflow, structure it around these stages:

| Stage | wordflow command | When to use |
|---|---|---|
| Preflight | `normalize`, `validate` | Before any edit — check the source is safe |
| Spec validation | `validate-spec` | After authoring the JSON spec, before applying it |
| Edit or review | `insert`, `add-comment`, `publish` | Apply the change with the right level of finality |
| Publish | `publish`, `publish-next` | When the candidate is validated and ready |
| Re-protect | `protect` | After the final publish, when the source was protected |

A skill should never skip the preflight stage, never write directly to the destination without a validate step, and never skip the re-protect step when the source required it.

### What to tell the agent explicitly

When prompting an agent to work on a document, the most reliable requests include:

- The full path to the input `.docx`
- The desired output path or version intent (`continue on latest` vs `publish new version`)
- The section or anchor where the change should go
- Whether changes should be visible (`highlight`, tracked changes) or clean
- Whether review comments should be added instead of direct edits
- Whether you want the final output protected (the agent detects this automatically, but you can mention it explicitly)

---


## Why this tool exists

This project exists because Word document automation becomes fragile very quickly when the environment includes:

- Cloud sync hydration and file locking behavior
- Word documents already open in the desktop app
- Files with `.docx` extension that are not true OOXML zip packages
- Protected, IRM-wrapped, or legacy-container inputs
- Versioned deliverables where a failed attempt must **not** create `v009`, `v010`, `v011`, and so on
- Formatting that must be preserved, not just text

`wordflow` is the attempt to turn that messy workflow into a deterministic toolchain with explicit validation and fail-fast behavior.

## What problem it solves

Without a dedicated tool, the workflow tends to become:

1. Copy files around manually
2. Try Word COM or ad hoc scripts
3. Discover the file is locked or cloud-only
4. Convert the file in a one-off way
5. Edit the document
6. Find out later that formatting or highlights were lost
7. Retry again with a new numbered version

That approach is slow, hard to reason about, and hard to reuse.

With `wordflow`, the goal is:

1. Stage work in `.wordflow/` adjacent to the document
2. Use a single candidate output
3. Validate before publish
4. Reject bad candidates early
5. Publish only when the result is technically and structurally acceptable

## What you get by having this tool

### Operational benefits

- One repeatable CLI instead of document-specific scripts
- A default temp workspace isolated from cloud sync friction
- Explicit OOXML validation
- Explicit fidelity checks for inherited formatting in untouched paragraphs
- A fail-fast publish workflow that does not churn version numbers
- Reusable commands for editing, migration, inspection, diffing, and validation

### Engineering benefits

- Behavior lives in code, not just in operator memory
- Tests protect the core document-editing behavior
- Workflows can be improved once and reused many times
- The tool can evolve into a stronger orchestrator without changing every downstream process

### Product/process benefits

- Fewer manual confirmations for routine document-handling steps
- Safer handling of documents that may already be open in Word
- Cleaner separation between:
  - Migration
  - Editing
  - Validation
  - Final publish

## How this behaves if we do **not** have the tool

If `wordflow` did not exist, the workflow would rely mostly on:

- Word COM automation
- PowerShell glue code
- Manual temp copies
- Manual validation
- Ad hoc logic per document family

That means:

- More operator decisions per run
- More chances to accidentally touch the final destination too early
- More chances to lose formatting while still producing a technically valid `.docx`
- Less reusable logic
- Less confidence that a failed attempt stopped at the right point

In short:

- **with the tool**: process becomes platform-like
- **without the tool**: process stays artisanal

## Feature support matrix

Use this matrix as the quickest source of truth for what the tool supports today.

| Feature area | Supported | Partial | Notes |
|---|---|---|---|
| Paragraph insertion, replacement, and deletion | [x] | [ ] | `insert-paragraphs`, `replace-text`, `delete-paragraphs` |
| Paragraph styles and visible markup | [x] | [ ] | `normal`, `heading1-3`, `list-bullet`, `list-number`, `quote`, plus highlight/bold/italic/underline on inserted content |
| Simple tables | [x] | [ ] | new table insertion with row/cell text, table style, and table-level highlight |
| Complex table layout | [ ] | [x] | merged cells, exact widths, nested tables, and safe editing of inherited complex layouts are not fully modeled yet |
| Hyperlinks | [x] | [ ] | external hyperlink relationships are created automatically |
| Images | [x] | [ ] | image parts and relationships are created automatically when the asset path exists |
| Section breaks | [x] | [ ] | continuous, next-page, even-page, odd-page |
| Comments and review CRUD | [x] | [ ] | add, list, update, and delete comments with in-document reference cleanup |
| Footnotes and endnotes | [x] | [ ] | inserted through `insert-note-after` |
| Content controls and fields | [x] | [ ] | rich-text content controls and complex field insertion |
| Tracked insertions and deletions | [x] | [ ] | review-friendly change markup |
| Header and footer upsert | [x] | [ ] | header/footer content can be created or replaced by section/reference |
| Core document properties | [x] | [ ] | title, subject, creator, description, keywords, last modified by |
| Anchor search and part targeting | [x] | [ ] | document, header/footer, or explicit `word/*.xml` part targeting |
| Validation and inspection | [x] | [ ] | `validate-spec`, `validate`, `inspect`, `list-parts`, `find-anchors`, `diff-docx`, `check-fidelity` |
| Normalization and staged publish | [x] | [ ] | `normalize`, `publish`, reusable sessions, and `publish-next` |
| Protected-source migration | [ ] | [x] | supported workflow, but non-OOXML conversion still depends on the guarded Word-backed migration path |
| Protection round-trip | [x] | [ ] | `prepare-session` captures protection metadata; `protect` re-applies it after N edits — Windows only |

If a feature area is not listed here, treat it as not yet committed as part of the reusable workflow contract.

## Project layout

```text
wordflow/
|- Cargo.toml
|- Cargo.lock
|- README.md
|- INSTRUCTIONS.md
|- examples/
|  |- document-update.json
|  |- review-comments.json
|  \- structured-update.json
\- src/
   |- main.rs
   |- lib.rs
   \- session.rs
```

## Build and test

**Windows (primary target — required for Word COM commands):**

```powershell
cargo +stable-x86_64-pc-windows-msvc test
```

**Linux / macOS (all non-COM commands):**

```bash
cargo test
```

The `migrate` command and the `--allow-word-com-encrypted-package` flag on `normalize` require Windows with Word installed. All other commands — editing, publishing, validation, comments, inspection, sessions — run on Linux and macOS without Word.

The 59 unit tests in `src/lib.rs` and `src/session.rs` cover only the cross-platform surface and pass on any OS.

## Runtime workspace

Each workflow command that needs staging creates a `.wordflow/` directory adjacent to the input document:

```
C:\Work\
├── contrato-v007.docx              ← --input
├── contrato-v008.docx              ← --output (published)
└── .wordflow/
    ├── wordflow.log                ← always-on log file
    ├── contrato.session.json       ← publish audit trail
    ├── run-<pid>-<stamp>/          ← ephemeral staging dir (cleaned on success)
    └── sessions/
        └── <session-id>/           ← persistent normalized working copy
```

If there is no parent directory (e.g. `--input document.docx` with no path prefix), the `.wordflow/` directory is created in the current working directory.

Override the workspace location for any command with `--temp-root`:

```powershell
wordflow publish `
  --input "C:\Work\contrato-v007.docx" `
  --output "C:\Work\contrato-v008.docx" `
  --spec ".\examples\document-update.json" `
  --temp-root "D:\scratch\wordflow"
```

`.wordflow/` is excluded from version control via `.gitignore`.

## Logging

Every run writes structured logs to two destinations simultaneously:

| Destination | Content | Format |
|---|---|---|
| stderr | real-time progress | `timestamp  LEVEL  message` |
| `.wordflow/wordflow.log` | persistent audit trail | `timestamp  LEVEL  message` |

Example output for a `publish` run:

```
2026-05-21T10:23:44.001Z  INFO wordflow 0.1.0
2026-05-21T10:23:44.003Z  INFO publish workflow started input=contrato-v007.docx output=contrato-v008.docx
2026-05-21T10:23:44.087Z  INFO applying spec operations=3
2026-05-21T10:23:44.412Z  INFO output published output=contrato-v008.docx
2026-05-21T10:23:44.415Z  INFO temp workspace cleaned up
```

On failure:

```
2026-05-21T10:23:44.501Z  WARN publish failed — temp workspace preserved for debugging temp_dir=.wordflow/run-123/
```

The default log level is `INFO`. To enable debug output:

```bash
RUST_LOG=debug wordflow publish ...
```

## Command overview

| Command | Purpose | Typical use |
|---|---|---|
| `insert` | apply a spec directly to an OOXML `.docx` | low-level editing |
| `insert --dry-run` | validate a spec against an OOXML `.docx` without writing output | low-risk spec check |
| `normalize` | check source packaging and publish an OOXML-safe copy when possible | source preflight |
| `prepare-session` | normalize once into a reusable temp session and reuse that working copy | iterative editing |
| `publish` | stage source in temp, create one candidate, validate, then publish | final document generation |
| `publish-next` | publish to the next detected versioned filename by default, or continue on the latest version when requested explicitly | iterative versioning |
| `add-comment` | add a review comment after an anchor with reviewer-friendly defaults | document review |
| `list-comments` | read comment metadata and in-document locations | document review |
| `update-comment` | update the body or metadata of an existing comment | document review |
| `delete-comment` | remove a comment and its in-document references | document review |
| `migrate` | convert a protected/ambiguous source into OOXML through Word | establish a safe OOXML base |
| `protect` | re-apply the original document protection after N editing iterations | close a protected round-trip |
| `validate-spec` | reject invalid highlights, anchors, and missing assets before document edits run | spec authoring |
| `validate` | check package structure and XML parsing | technical validation |
| `check-fidelity` | compare inherited formatting between two OOXML docs | trust decision before publish or migration |
| `find-anchors` | locate paragraphs by text/mode/occurrence | spec authoring and debugging |
| `list-parts` | list package parts and hashes | inspection/debugging |
| `inspect` | high-level package overview | quick health check |
| `diff-docx` | compare changed parts across two docs | impact review |

## Command details

### 1. `insert`

Low-level command that applies a JSON spec to an OOXML `.docx`.

Use this when:

- You already trust the input document
- You want direct edit behavior
- You are not publishing a business-facing final version yet

```powershell
cargo +stable-x86_64-pc-windows-msvc run -- insert `
  --input "C:\path\source.docx" `
  --output "C:\path\updated.docx" `
  --spec ".\examples\document-update.json"
```

### 2. `insert --dry-run`

Builds the candidate in memory and validates it without writing the output file.

Use this when:

- Testing a new spec
- Checking anchors
- Validating an operation set before publish

```powershell
cargo +stable-x86_64-pc-windows-msvc run -- insert `
  --input "C:\path\source.docx" `
  --output "C:\path\ignored.docx" `
  --spec ".\examples\document-update.json" `
  --dry-run
```

### 3. `normalize`

Preferred preflight command before editing sources from mixed origins.

Behavior:

1. Reads the source bytes
2. Detects whether the source is already a real OOXML zip package
3. If already normalized, validates it and publishes an OOXML-safe copy
4. If not normalized, reports the detected packaging and stops unless a reliable normalization backend exists

```powershell
cargo +stable-x86_64-pc-windows-msvc run -- normalize `
  --input "C:\path\source.docx" `
  --output "C:\path\normalized.docx"
```

For `ole-encrypted-package` sources only, you can explicitly allow the guarded COM exception:

```powershell
cargo +stable-x86_64-pc-windows-msvc run -- normalize `
  --input "C:\path\source.docx" `
  --output "C:\path\normalized.docx" `
  --allow-word-com-encrypted-package
```

That exception is intentionally narrow:

- It is only available in `normalize`
- It is only used when the detected format is `ole-encrypted-package`
- It is used only to establish a normalized OOXML base
- After that, the rest of the workflow should continue without COM

### 4. `publish`

This is the preferred command for final document updates.

Behavior:

1. Creates or uses `.wordflow/` adjacent to the input document
2. Copies the source there
3. Checks that the staged source is already normalized for OpenXML editing
4. Validates the staged source
5. Builds a single candidate output
6. Validates the candidate
7. Runs source-fidelity checks for untouched paragraphs
8. Publishes to the destination only if all checks pass
9. Cleans temp on success
10. Preserves temp on failure for debugging

```powershell
cargo +stable-x86_64-pc-windows-msvc run -- publish `
  --input "C:\path\source.docx" `
  --output "C:\path\project-plan-v008.docx" `
  --spec ".\examples\document-update.json"
```

Override the temp root if needed:

```powershell
cargo +stable-x86_64-pc-windows-msvc run -- publish `
  --input "C:\path\source.docx" `
  --output "C:\path\final.docx" `
  --spec ".\examples\document-update.json" `
  --temp-root "C:\Temp\custom-docx-work"
```

### 5. `prepare-session`

Prepare or reuse a normalized working copy in `.wordflow/sessions/<session-id>/` adjacent to the input document.

Use this when:

- You expect multiple sequential edits
- The published destination may be rewrapped by a cloud sync provider or protection
- You want the cache to live in the temp session, not in the document

```powershell
cargo +stable-x86_64-pc-windows-msvc run -- prepare-session `
  --input "C:\path\project-plan-v015.docx" `
  --session "project-plan-session"
```

If the source bytes have not changed, the command reuses the cached normalized working copy instead of normalizing again.

### 6. `publish-next`

Publish to the next detected versioned output, either from a direct input or from a prepared session.

Direct input:

```powershell
cargo +stable-x86_64-pc-windows-msvc run -- publish-next `
  --input "C:\path\project-plan-v015.docx" `
  --spec ".\examples\document-update.json"
```

Prepared session:

```powershell
cargo +stable-x86_64-pc-windows-msvc run -- publish-next `
  --session "project-plan-session" `
  --spec ".\examples\document-update.json"
```

Continue working on the latest published version instead of creating `vN+1`:

```powershell
cargo +stable-x86_64-pc-windows-msvc run -- publish-next `
  --session "project-plan-session" `
  --spec ".\examples\document-update.json" `
  --target latest
```

Behavior:

1. Validates the spec
2. Finds the highest existing matching version in the output directory
3. By default, publishes exactly one next version
4. With `--target latest`, updates the latest published version in place instead of incrementing
5. If a session is used, refreshes the cached normalized working copy to the newly validated content

### 7. `validate-spec`

Validate a JSON spec before any `.docx` edit runs.

```powershell
cargo +stable-x86_64-pc-windows-msvc run -- validate-spec `
  --spec ".\examples\document-update.json"
```

This catches generic authoring problems such as:

- Unsupported highlight values
- Empty anchors
- Invalid occurrences
- Unsupported part targets
- Missing image assets

### 8. `migrate`

Preferred when the source is:

- IRM-wrapped
- Ambiguous
- Not yet a real OOXML zip package
- Better handled through Word for conversion only

Behavior:

1. Stages the source in `.wordflow/` adjacent to the input document
2. Exports source text through Word
3. Converts `source -> RTF -> OOXML`
4. Exports migrated text through Word
5. Exports Word-based paragraph signatures from source and candidate
6. Requires the text exports to match
7. Requires the Word-based paragraph signatures to preserve inherited formatting
8. Validates the migrated OOXML candidate
9. Optionally compares fidelity against a trusted OOXML reference
10. Publishes only on success

```powershell
cargo +stable-x86_64-pc-windows-msvc run -- migrate `
  --input "C:\path\source.docx" `
  --output "C:\path\migrated.docx"
```

With an OOXML reference:

```powershell
cargo +stable-x86_64-pc-windows-msvc run -- migrate `
  --input "C:\path\source.docx" `
  --output "C:\path\migrated.docx" `
  --trusted-ooxml "C:\path\trusted-base.docx"
```

### 9. `protect`

Re-applies the original document protection to the latest published version in a session. Requires Windows with Word installed.

Use this when:

- The source was a password-encrypted or protection-restricted document
- You used `prepare-session` + `publish-next` for N editing iterations
- You need to deliver a protected output as a final step

`prepare-session` automatically captures the protection type, password requirement, and IRM status during the Word COM migration. `protect` reads that metadata and re-applies it.

```powershell
cargo +stable-x86_64-pc-windows-msvc run -- protect `
  --session "contrato-session" `
  --output "C:\path\contrato-v010-protected.docx" `
  --password "secret"
```

Without a password (editing restriction only, no encryption):

```powershell
cargo +stable-x86_64-pc-windows-msvc run -- protect `
  --session "contrato-session" `
  --output "C:\path\contrato-v010-protected.docx"
```

`protect` will error descriptively if:

- The session was created from a plain OOXML source — no protection metadata exists
- The original required a password but `--password` was not provided
- The original document had no protection at all

The command uses the `current_version_path` stored in the session metadata as its input, so it always protects the latest published version without needing `--input`.

### 10. `validate`

Checks whether the package is readable as normalized OOXML and whether XML parts parse correctly.

```powershell
cargo +stable-x86_64-pc-windows-msvc run -- validate --input "C:\path\source.docx"
```

### 10. `check-fidelity`

Compares two OOXML files and detects whether inherited formatting disappeared from paragraphs that should have stayed intact.

Use this to answer:

- Did the candidate preserve existing highlights?
- Did it preserve inherited paragraph styles?
- Should we trust this migrated file as a new technical base?

```powershell
cargo +stable-x86_64-pc-windows-msvc run -- check-fidelity `
  --before "C:\path\trusted-base.docx" `
  --after "C:\path\candidate.docx"
```

With a spec:

```powershell
cargo +stable-x86_64-pc-windows-msvc run -- check-fidelity `
  --before "C:\path\trusted-base.docx" `
  --after "C:\path\candidate.docx" `
  --spec ".\examples\document-update.json"
```

### 11. Comment review commands

Add a comment with GitHub Copilot CLI reviewer defaults:

```powershell
cargo +stable-x86_64-pc-windows-msvc run -- add-comment `
  --input "C:\path\source.docx" `
  --output "C:\path\reviewed.docx" `
  --anchor "Strategic principles" `
  --comment-text "Please clarify the approval workflow."
```

List comments:

```powershell
cargo +stable-x86_64-pc-windows-msvc run -- list-comments --input "C:\path\reviewed.docx"
```

Update a comment:

```powershell
cargo +stable-x86_64-pc-windows-msvc run -- update-comment `
  --input "C:\path\reviewed.docx" `
  --output "C:\path\reviewed-updated.docx" `
  --id 1 `
  --comment-text "Please clarify the public review workflow."
```

Delete a comment:

```powershell
cargo +stable-x86_64-pc-windows-msvc run -- delete-comment `
  --input "C:\path\reviewed-updated.docx" `
  --output "C:\path\reviewed-clean.docx" `
  --id 1
```

### 12. Inspection and debugging commands

```powershell
cargo +stable-x86_64-pc-windows-msvc run -- inspect --input "C:\path\source.docx"
cargo +stable-x86_64-pc-windows-msvc run -- list-parts --input "C:\path\source.docx"
cargo +stable-x86_64-pc-windows-msvc run -- find-anchors --input "C:\path\source.docx" --text "Strategic principles" --mode equals --occurrence 2
cargo +stable-x86_64-pc-windows-msvc run -- diff-docx --before "C:\path\old.docx" --after "C:\path\new.docx"
```

## Recommended usage patterns

### Safe final update for a normal OOXML source

1. `normalize`
2. `validate-spec`
3. `validate`
4. Optional `insert --dry-run`
5. `publish`

### Fast iterative update for a versioned document family

1. `prepare-session`
2. `validate-spec`
3. `publish-next --session ... --target latest` while refining content
4. `publish-next --session ...` when you want to close a new published version
5. Repeat as needed while the session working copy remains the active normalized base

### Safe migration for a protected source

1. `inspect`
2. `normalize` or `migrate`, depending on the available backend
3. Optional `check-fidelity`
4. `publish`

### Protected document round-trip with N editing iterations

Use this when the source is password-encrypted or protection-restricted and the final delivery must also be protected.

1. `prepare-session` — normalizes the source via Word COM (slow, once), captures protection metadata
2. `validate-spec` — verify the spec before any edits
3. `publish-next --session ... --target latest` — iterate freely on content (fast, pure Rust)
4. `publish-next --session ...` — close a new numbered version when ready
5. `protect --session ... --output ... --password "..."` — re-apply original protection (slow, once)

The agent asks the user for the password only once, at step 5. Steps 3–4 are as fast as working on a plain OOXML document.

### Investigate why a candidate is bad

1. `inspect`
2. `validate`
3. `check-fidelity`
4. `diff-docx`
5. `list-parts`

## Spec model

The editing commands consume a JSON spec with `operations`.

Common operation types:

- `insert-paragraphs`
- `replace-text`
- `delete-paragraphs`
- `insert-table-after`
- `insert-hyperlink-after`
- `insert-image-after`
- `insert-section-break-after`
- `insert-comment-after`
- `insert-note-after`
- `insert-content-control-after`
- `insert-field-after`
- `track-insert-paragraphs`
- `track-delete-paragraphs`
- `upsert-header-footer`
- `set-core-property`

Anchors can be strings or structured objects:

```json
{
  "text": "Strategic principles",
  "mode": "equals",
  "occurrence": 2
}
```

Supported anchor modes:

- `contains`
- `equals`
- `starts-with`
- `ends-with`

Part targets can be:

- `document`
- `header:1`
- `footer:1`
- Explicit `word/*.xml` paths

## Example spec

See:

- `examples\document-update.json`
- `examples\review-comments.json`
- `examples\structured-update.json`

These samples are intentionally practical:

| Sample | Best for | What it covers |
|---|---|---|
| `examples\document-update.json` | end-to-end content update | headings, bullets, table insertion, comments, notes, content controls, fields, tracked changes, hyperlink, section break, core properties, header update |
| `examples\review-comments.json` | review pass | comment insertion, tracked insertions/deletions, and review notes |
| `examples\structured-update.json` | document structure edits | headings, numbered lists, quote blocks, table insertion, footer upsert, field insertion, content controls, section break, and metadata |

Start from the closest sample, then change anchors, paths, and wording instead of authoring a spec from scratch every time.

## Highlighting behavior

Inserted text uses green highlight by default unless another highlight value is specified explicitly in the spec.

Supported highlight values are:

- `black`
- `blue`
- `cyan`
- `darkBlue`
- `darkCyan`
- `darkGray`
- `darkGreen`
- `darkMagenta`
- `darkRed`
- `darkYellow`
- `green`
- `lightGray`
- `magenta`
- `none`
- `red`
- `white`
- `yellow`

## Practical rules

- Prefer `publish` over `insert` for final outputs
- Prefer `migrate` over ad hoc Word automation for protected inputs
- Do not touch the final destination until validation passes
- Do not rely on automatic retries to solve structural failures
- Treat temp artifacts preserved after failure as debugging evidence, not clutter
- Treat `migrate` as the guarded path that checks both text and Word-observed formatting before accepting a converted base
- Treat `publish-next` as **new-version by default** and use `--target latest` only when you explicitly want to continue working on the latest published document without incrementing the version

## Using with AI agents

This tool works best when the agent is asked to do two things clearly:

1. Understand the document goal
2. Apply the right workflow mode (`summary`, `edit`, `continue working`, or `publish`)

Good requests usually specify:

- The input `.docx`
- The desired output or destination
- Whether to continue on the latest version or publish a new one
- Any required highlight color
- The target section, anchor, or outcome

### Prompt patterns

#### 1. Executive summary

```text
Read C:\path\stakeholder-feedback.docx and create a professional Word summary with key findings, executive summary, and action items at C:\path\feedback-summary.docx.
```

#### 2. Summary with backlog actions

```text
Read C:\path\feedback.docx and create a professional summary document with action items, owners, and the backlog items that should be created to address the pending work.
```

#### 3. Add content after a known section

```text
Update C:\path\report-v015.docx by adding a new section after the target heading. Use magenta highlight for the new text and publish a new version.
```

#### 4. Continue refining the latest published version

```text
Keep working on the latest published version of C:\path\report-v015.docx. Do not create a new version yet. Add the requested edits and continue on the latest version.
```

That maps to the `publish-next --target latest` behavior.

#### 5. Close a new published version

```text
Use the latest version of C:\path\report-v015.docx as the base, apply the requested edits, and publish the next numbered version.
```

That maps to the default `publish-next` behavior.

#### 6. Ask for explicit validation

```text
Before publishing, validate the spec and make sure the document remains a valid OOXML .docx.
```

#### 7. Ask for visible change marking

```text
Add the new content in yellow highlight so the changes are easy to review.
```

Only use supported highlight values:

- `green`
- `yellow`
- `magenta`
- `cyan`
- And the other values listed in the highlighting section above

#### 8. Ask the agent to leave review comments

```text
Read C:\path\report.docx, add review comments for the unclear sections, and save the reviewed copy as C:\path\report-reviewed.docx.
```

#### 9. Ask the agent for a simple status table

```text
Update C:\path\report.docx by adding a status table after the target anchor with columns Workstream, Owner, Due date, and Status. Keep it as a standard Word table and save the result as C:\path\report-with-status-table.docx.
```

That maps well to the current simple-table support.

### Practical guidance

- Ask for **continue working on the latest version** when you are still iterating
- Ask for **publish the next numbered version** when you want to close a revision
- Ask for **review comments** when you want feedback without rewriting the document content immediately
- Mention the **tone** you want for summaries (`professional`, `executive`, `concise`, `detailed`)
- Mention whether you want **action items**, **owners**, **PBIs**, or **follow-up questions**
- If the source may be protected or ambiguous, ask the agent to **normalize first**

## Session history

Every successful `publish` and `publish-next` call automatically writes or appends to a `.session.json` file placed inside `.wordflow/`, adjacent to the document.

The session file name is derived from the document base name, without the version suffix:

```
C:\Work\
├── report-v014.docx
├── report-v015.docx
├── report-v016.docx
└── .wordflow\
    └── report.session.json     ← created automatically, shared across all versions
```

The file is not versioned — it is a single persistent record that accumulates entries over the entire life of the document.

### What the session file contains

```json
{
  "document": "report",
  "session_file": "C:\\Work\\report.session.json",
  "created_at": "2026-01-10T14:22:00Z",
  "updated_at": "2026-05-20T09:15:00Z",
  "entries": [
    {
      "timestamp": "2026-01-10T14:22:00Z",
      "command": "publish",
      "input": "C:\\Work\\report-v014.docx",
      "output": "C:\\Work\\report-v015.docx",
      "spec": "add-executive-summary.json",
      "result": "ok"
    },
    {
      "timestamp": "2026-05-20T09:15:00Z",
      "command": "publish-next",
      "output": "C:\\Work\\report-v016.docx",
      "spec": "add-status-table.json",
      "result": "ok",
      "note": "mode=next-version version=v016"
    }
  ]
}
```

### Reading the session — `show-session`

Accepts either the `.docx` path (derives the session file automatically) or the `.session.json` path directly:

```powershell
cargo +stable-x86_64-pc-windows-msvc run -- show-session `
  --input "C:\path\report-v016.docx"
```

```powershell
cargo +stable-x86_64-pc-windows-msvc run -- show-session `
  --input "C:\path\report.session.json"
```

Output is JSON — designed to be consumed directly by an agent to restore context before continuing work on a document.

### Rebuilding history from existing files — `reconstruct-session`

For documents that already have versioned files but no session file, `reconstruct-session` infers history by running `diff-docx` on each consecutive pair:

```powershell
cargo +stable-x86_64-pc-windows-msvc run -- reconstruct-session `
  --folder "C:\path\" `
  --document "report"
```

The result is saved to `report.session.json` in the same folder and printed to stdout.

Reconstructed entries include:
- Input and output file paths
- Timestamp from the file modification time
- Number of changed XML parts
- A note indicating the entry was inferred, not recorded live

What cannot be reconstructed: the spec that was used, the intent behind the change, or whether there were failed attempts before the published version.

If a session file already exists, `reconstruct-session` merges only the entries that are not yet tracked, preserving any entries recorded live.

## Current limitations

This tool is much stronger than an ad hoc script, but two limits still matter:

1. Protected or ambiguous sources may still require the narrowly scoped Word-backed migration path to establish a normalized OOXML base
2. Complex table layouts such as merged cells, exact column widths, nested tables, and heavy editorial formatting are only partially covered today

The long-term goal is to keep Word as a conversion fallback only, and to expand the table model without turning the CLI into a fragile layout editor.
