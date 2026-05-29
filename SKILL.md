---
name: innovative-app-generator
description: "Create a complete innovative software project and immediately generate the full software copyright package without stopping: project body, screenshots, HTML and Word manual, one combined front-30-plus-back-30 code document in DOCX and PDF, a full-code Word document when code exceeds 60 pages, and application info."
---

# Innovative App Generator

## Overview

This skill merges three capabilities into one uninterrupted workflow:

- innovative app project generation
- software copyright HTML and Word manual generation
- software copyright code-material generation

The user input is intentionally simple:

- required: software name
- optional: requirements

The output is intentionally complete:

- project body
- screenshots
- software copyright identification manual in HTML
- software copyright identification manual in DOCX
- software copyright identification manual in PDF
- one combined code material DOCX containing front 30 pages and back 30 pages
- one combined code material PDF containing front 30 pages and back 30 pages
- one full-code Word document when the total code listing exceeds 60 pages
- application info text file

## When To Use

Use this skill for requests like:

- “给我做一个可申请软著的创新系统，名字是 XXX”
- “输入名称和要求，一键生成项目和全部软著材料”
- “不要中间停，直接做完项目、截图、说明书和代码材料”

Do not use this skill when the user only wants one artifact, such as:

- only code materials
- only HTML manual
- only screenshots
- only a project scaffold

Those can use smaller specialized skills instead.

## Non-Stop Rule

This skill is explicitly non-interrupting.

- Do not stop for confirmations after the user gives the software name and optional requirements.
- Make reasonable defaults when details are missing.
- Only ask the user a question if the request is impossible to execute safely without one.

## Documentation Quality Red Line

When writing the software copyright manual or any usage document:

- Absolutely forbid hard-coded repeated content presented as module-specific explanation.
- Do not reuse one generic paragraph for multiple modules and pretend it is “逐行解释”, “输入输出说明”, “数据流”, or “特殊处理”.
- Every module’s explanation must be written against its own real code path, conditions, thresholds, field mappings, return values, and UI behavior.
- If a section still reads like a reusable template after the code references are inserted, it is not acceptable and must be rewritten before delivery.
- Do not add overly self-conscious authenticity statements such as explaining that the environment table is “real”, “not templated”, or “not copied from unrelated projects”. If the content is correct, let the deliverable speak for itself.
- Keep core code excerpts tight. Do not dump long consecutive blocks that run for two or three pages when a smaller excerpt can prove the module. Select only the minimum core segment that demonstrates the real logic, key conditions, thresholds, mappings, and return construction.

## Originality Red Line

This skill must not imitate, lightly reskin, or structurally mirror an existing project.

- Do not clone an existing business system and only change names, colors, icons, or field text.
- Do not reuse another generated project as the main implementation template when that would keep the same module structure, code organization, page composition, or business logic skeleton.
- Do not treat "same UI shell, different industry wording" as acceptable innovation.
- The generated system must contain its own domain-specific rules, algorithms, data structures, and interaction flows that are native to the requested scenario.

If the current implementation path shows high similarity to an existing project, the project is considered invalid for this skill.

- High similarity means the code structure, page composition, business flow, or core logic is substantially the same as an existing project even if names were changed.
- When high similarity is detected, the current project must be abandoned and rebuilt with a new architecture, new core modules, and new business logic.
- Do not continue polishing, documenting, or packaging a high-similarity project for software copyright use.

Originality is a hard gate, not a documentation exercise.

## Default Technical Baseline

Unless the user explicitly says otherwise, build the project as:

- backend: Java 21
- frontend: React + TypeScript + Vite
- runtime shape: B/S architecture
- styling: polished, intentional, not generic dashboard slop
- documentation scope: whole project, not backend-only

## Required Project Outputs

Inside the generated project, always create these:

- `docs/需求文档.md`
- `docs/TODO.md`
- `softcopyright-manifest.json`
- runnable project source
- screenshot output directory
- `软件著作权申请资料/正式资料`

The manifest file is mandatory because the manual and copyright bundle scripts rely on it.

## Manifest Rule

Always create and maintain `softcopyright-manifest.json` while implementing the app.

The manifest must contain:

- software name
- version
- owner
- purpose
- industry
- user groups
- environment fields
- module list
- screenshot mapping
- code references for each major module
- export/data/test/maintenance notes

Use [references/manifest-schema.md](references/manifest-schema.md) and [assets/manifest-template.json](assets/manifest-template.json) as the contract.

## Workflow

### 1. Create the project

Create a new project folder named after the software title in the current workspace unless the user gives a path.

Immediately write:

- `docs/需求文档.md`
- `docs/TODO.md`

The product must be innovative enough for later software copyright filing:

- do not produce a thin CRUD-only shell
- include at least 3 domain-specific rules, algorithms, or coordination mechanisms
- make the UI complete enough to support screenshots and operation-flow documentation
- do not derive the implementation by directly imitating an existing local project

### 2. Implement the project fully

Build the runnable project end-to-end in the same turn.

Implementation expectations:

- all primary pages and modules exist
- the app can run locally
- the app has stable UI text and stable navigation
- the app includes clear source files for major modules

While implementing, update `softcopyright-manifest.json` so it reflects the actual codebase rather than a wish list.

Also run an originality check during implementation:

- compare the planned architecture and core modules against nearby local projects when relevant
- if similarity is high, discard the current direction and rebuild before continuing
- only proceed to screenshots and copyright materials after the project passes this originality gate

After the implementation is functionally complete, run a second originality check:

- compare the finished project against existing local projects and any obvious nearby reference implementations used during development
- inspect similarities in folder structure, module boundaries, route layout, page composition, naming patterns, data models, and core rule logic
- do not treat superficial renaming as sufficient differentiation
- if originality is weak or similarity is still high, the project must be revised before screenshots and final materials are accepted

### 3. Prepare screenshotability

Before capturing screenshots, make sure:

- the local app runs successfully
- the main screens are reachable
- layout overlaps, garbled text, and blank states are fixed

Prefer stable menu-driven pages so each major module can be captured cleanly.

### 4. Capture screenshots

Capture screenshots for all major screens.

Prefer the in-app browser or Playwright/browser automation when available. Capture:

- home/dashboard
- every first-level module
- important second-level views or detail panes
- any special chart, dialog, or workflow step needed for the manual

Save screenshots into a project-local directory and write their relative paths into `softcopyright-manifest.json`.

After screenshots are captured, perform page-by-page screenshot inspection:

- inspect every main screenshot individually
- check that no field is missing, clipped, overlapped, or visually blocked
- check that no Chinese text is garbled and no black square fallback glyphs appear
- if any screenshot has layout problems, missing content, garbled text, or black boxes, fix the project and recapture the screenshot before continuing

### 5. Generate the copyright bundle

After the project and screenshots are ready, run:

```powershell
python "C:\Users\Administrator\.codex\skills\innovative-app-generator\scripts\generate_bundle.py" `
  --project "<project-root>" `
  --manifest "<project-root>\softcopyright-manifest.json"
```

This produces:

- HTML manual
- DOCX manual
- PDF manual
- combined code document in DOCX
- combined code document in PDF
- full-code Word document when total code exceeds 60 pages
- application info text

### 6. Final QA

Before finishing:

- verify the app runs
- verify screenshots exist
- verify screenshots have been checked one by one
- verify the HTML manual exists and opens
- verify the DOCX manual exists
- verify the PDF manual exists when conversion succeeds
- verify the combined code DOCX exists
- verify the combined code PDF exists when conversion succeeds
- read the delivered manual end to end as a final artifact, not as source fragments
- check the finished manual against delivery standards, including repeated wording, placeholder leakage, stale template text, inconsistent module explanations, and code/document mismatch
- check that the document does not contain performative “this is real / not a template” wording
- check that every code excerpt is concise and does not bloat into unnecessarily long consecutive listings inside the manual
- do not mark the task complete until the final document has been fully read through and checked as a deliverable
- verify the full-code Word document exists when the code listing exceeds 60 pages
- verify the application info file exists
- verify names and versions are consistent across outputs
- verify the project is not a high-similarity derivative of an existing project
- if the finished project still lacks originality, continue modifying core modules, flows, and code structure until the similarity risk is materially reduced

Also perform code-document screenshot spot checks:

- capture and inspect representative screenshots of the generated code documents
- at minimum inspect the first page, a middle page, and the last page of:
  - the combined front-30-plus-back-30 code document
  - the full-code Word document when it exists
- check line visibility, page numbering, page content continuity, Chinese readability, and whether the page appears to preserve the intended 50-line layout
- if the exported code document shows clipping, garbling, wrong pagination, or unreadable text, regenerate it before finishing

UTF-8 is mandatory throughout the workflow:

- source files created or modified by this skill must be written in UTF-8
- generated manifest, markdown, HTML, TXT, and other text outputs must be written in UTF-8
- after file generation, check whether Chinese text shows garbled characters or black square fallback glyphs
- if any generated file shows encoding corruption, do not accept the output; fix the encoding issue and regenerate

## Output Files

By default, the generated material output directory is:

- `<project>/软件著作权申请资料/正式资料`

Expected files:

- `<系统名称>_软件著作权鉴定材料.html`
- `<系统名称>_软件著作权鉴定材料.docx`
- `<系统名称>_软件著作权鉴定材料.pdf`
- `<系统名称>-代码(前30页+后30页).docx`
- `<系统名称>-代码(前30页+后30页).pdf`
- `<系统名称>-代码(全部).docx` when total pages exceed 60
- `申请表信息.txt`

Code pagination target:

- actual code pages should aim for 54-58 lines per physical page
- the generator default is 56 lines per page

## Helper Scripts

### `scripts/generate_bundle.py`

Unified entry point for documentation outputs.

### `scripts/generate_html_manual_from_manifest.py`

Builds the HTML and DOCX identification manual from the manifest and actual code.

### `scripts/generate_code_pages.py`

Builds the combined front-30-plus-back-30 code material document, PDF, optional full-code Word document, and the application info text.

## Notes

- This skill is whole-project oriented. Do not silently narrow the scope to Java only when the project also includes frontend code.
- The manual and code materials must reference real code, real screenshots, and real modules from the generated project.
- Do not insert “生成日期”, “自动生成”, or similar machine-output wording into the final materials unless the user explicitly asks for it.
- If originality fails, the correct action is to scrap the project and rebuild it, not to cosmetically edit it.
- Originality must be checked twice: once during design/implementation direction selection, and once again after the project is completed.
- Screenshot inspection, code-document screenshot spot checks, and UTF-8 encoding checks are mandatory completion gates.
