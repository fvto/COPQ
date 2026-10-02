# Master prompt: Extract a standalone tool into a portable MCP handoff

## Mission and deliverable

You are the extraction agent. Inspect a working standalone project and produce one portable handoff. Choose the smallest complete mode:

~~~text
Mode A: [tool_name]_MCP_IMPLEMENTATION_SPEC.md
Mode B: [tool_name]_MCP_HANDOFF.zip
~~~

A second AI will use the handoff to integrate the workflow into the target MCP project. **Assume the second AI cannot access the standalone source repository or the extractor's computer.** Mode A is allowed only when the Markdown alone contains everything needed to reconstruct locally authored logic and static assets. Otherwise use Mode B: exactly one ZIP containing the Markdown specification and the minimal necessary source, assets, and safe test fixtures. The Markdown is the entry point in both modes. It must let the implementer reproduce observed business behavior, provision every dependency, route each entry point, and prove a real successful result without guessing. A claim of perfect portability is never a substitute for evidence: identify blockers when code, assets, credentials, or runtime behavior cannot be verified.

**Deliver exactly one user-facing artifact: one Markdown file or one ZIP.** In Mode B, the ZIP must contain one top-level folder with the Markdown specification at its root. Keep the manifest, integration instructions, source references, and uncertainties in that Markdown. Include only necessary additional files in the ZIP. Do not include secrets, raw confidential records, or production values.

### Lifecycle, Versioning, and Continuous Evolution
Tools in upstream standalone repositories inevitably evolve over time (e.g. bug fixes, algorithm revisions, new sub-agents, or updated templates). To ensure long-term maintainability:
1. **Evolution Ledger Initialization**: Every extraction must establish a persistent tracking ledger: `docs/tools/versions/[tool_name]_EVOLUTION_LOG.md` (or include it within the handoff package). This ledger stores baseline commit hashes, SHA-256 fingerprints, schema contracts, and algorithm notes.
2. **Companion Delta Prompt (`MCP_TOOL_DELTA_EXTRACTOR.md`)**: Whenever the upstream source repository releases updates in the future, AI agents must **NOT** guess or restart from scratch. Instead, invoke the companion master prompt [`docs/tools/extract-prompt/MCP_TOOL_DELTA_EXTRACTOR.md`](MCP_TOOL_DELTA_EXTRACTOR.md). That prompt reads the previous specification and the Evolution Ledger, runs a 3-way AST/code delta audit, classifies changes (breaking vs. non-breaking), and produces a concise delta migration plan to update the MCP codebase seamlessly.

### ⚡ MANDATORY PRE-EXTRACTION DIRECTIVE: Git Repository Hygiene & Staged Baseline
Before extracting specifications or packaging, the AI on the standalone project machine MUST check whether Git is initialized in `[STANDALONE_ROOT]`:
- **Trường hợp 1: Dự án chưa có Git (`.git` chưa tồn tại)**:
  BẮT BUỘC thực hiện `git init`, tạo file `.gitignore` chuẩn (chặn `.venv/`, `__pycache__/`, `~$*`, output logs, secrets, v.v.), sau đó **commit theo từng stage có cấu trúc** theo chuẩn Conventional Commits (`feat(core)`, `feat(templates)`, `feat(agents)`, `docs(integrate)`) và đánh tag `v1.0.0`.
- **Trường hợp 2: Dự án đã có Git nhưng còn file chưa commit (dirty state / untracked)**:
  BẮT BUỘC gom các file thành từng stage hợp lý và commit với message chuẩn xác thay vì commit gộp một lần với message chung chung (như "commit all" hay "update").
- **Mục đích cốt lõi**:
  Việc thiết lập baseline Git chuẩn và đánh tag `v1.0.0` ngay từ lúc trích xuất ban đầu là điều kiện tiên quyết để prompt [`MCP_TOOL_DELTA_EXTRACTOR.md`](MCP_TOOL_DELTA_EXTRACTOR.md) sau này có thể dùng lệnh `git diff v1.0.0..HEAD` và lịch sử commit nhằm bóc tách chính xác 100% các sai khác thuật toán, file mã nguồn và template mới mà không bị thất thoát hay đoán mò.

## Inputs to this prompt

- Standalone project root: [STANDALONE_ROOT]
- Target MCP project root, if accessible: [TARGET_MCP_ROOT_OR_UNAVAILABLE]
- Desired tool/workflow name, if known: [TOOL_NAME_OR_INFER]
- Owner requirements and supported entry points: [OWNER_NOTES_OR_NONE]
- Safe example inputs or artifacts, if available: [SAMPLES_OR_NONE]
- Whether a non-destructive local run is authorized: [YES / NO / UNKNOWN]
- Handoff mode requested by owner, if any: [AUTO / MARKDOWN_ONLY / ZIP_WITH_MARKDOWN]

Resolve these from the user's request and environment before asking for information. Ask only for a missing business-critical value or access that blocks further analysis. Continue read-only investigation while awaiting an answer. Never request passwords, cookies, tokens, or authentication headers.

Mode selection: honor an explicit owner choice. With AUTO, use Mode A only if every necessary locally authored source component and static asset is fully reconstructible from the Markdown; otherwise use Mode B. If Mode B cannot lawfully or safely include a required file, keep the requested artifact shape but mark the handoff BLOCKED and identify the file. Do not silently omit it or claim readiness.

## Authority and evidence rules

1. The user's instructions and applicable repository instructions take precedence.
2. Read source, tests, configuration, startup scripts, deployment files, logs, and representative artifacts. Runtime behavior and current code outrank a stale README; record any conflict.
3. Mark substantive statements as **CONFIRMED_BY_CODE**, **CONFIRMED_BY_CONFIG**, **CONFIRMED_BY_TEST**, **CONFIRMED_BY_RUNTIME**, **CONFIRMED_BY_USER**, **INFERRED**, or **UNKNOWN**. Cite file paths and symbols; include line numbers where practical. Identify the command, fixture, or artifact behind runtime claims.
4. Separate **current standalone behavior** from **recommended MCP adaptation**. Never present an adaptation as if the standalone program already does it.
5. Never invent APIs, portal selectors, report columns, dates, factory mappings, package versions, endpoints, output counts, or success metrics. Write UNKNOWN or IMPLEMENTATION BLOCKER when evidence is missing.
6. Do not infer success from process exit code, a printed success message, a browser click, or the existence of a single output. Require the workflow's actual terminal artifacts and semantic checks.
7. Do not perform a destructive or production run merely to fill the specification. Inspect logs and existing artifacts first. If a safe run is unavailable, state which assertions remain unverified.
8. If the target MCP repository is accessible, inspect its **current** registry, tool schemas, gateways, config loader, path policy, permissions, routes, tests, and deployment method. Map to real APIs only after checking them. If it is unavailable, provide a conceptual mapping and mark target-specific details UNKNOWN. Do not let a design document override live target source.
9. Record the inspected source revision or file hashes, the extraction date, and any local uncommitted changes that affect behavior. A handoff based on one version must not silently claim parity with a later version.

## Extraction procedure

### 0. Establish repository hygiene, Git baseline, and version anchor (Bắt buộc)

Before analyzing functions and call graphs, enforce version control discipline on `[STANDALONE_ROOT]`:

1. **Check Git Status**:
   - Check if `.git` folder exists. If absent, execute `git init`.
2. **Generate Standard `.gitignore`**:
   Ensure a clean `.gitignore` exists at the project root containing:
   ```gitignore
   # Python & Virtual Environments
   .venv/
   venv/
   env/
   __pycache__/
   *.py[cod]
   .pytest_cache/

   # IDEs & System files
   .idea/
   .vscode/
   .DS_Store
   Thumbs.db

   # Temporary Office lock files
   ~$*.xlsx
   ~$*.pptx
   ~$*.docx

   # Output & Execution Artifacts
   *.log
   *.tmp
   output/
   temp/
   tmp/
   .stage-*/
   _build_*/
   *.zip

   # Secrets & Credentials
   *.env
   *.key
   *.pem
   ```
3. **Semantic Granular Staging & Conventional Commits**:
   Do NOT stage all files in one indiscriminate `git add .`. Commit in logical, semantic stages:
   - **Stage 1 (Core Engine & Data Processors)**:
     ```powershell
     git add *.py (core business processing scripts, excluding agents/tests)
     git commit -m "feat(core): implement core algorithms and data transformation logic"
     ```
   - **Stage 2 (Master Templates & Style Palettes)**:
     ```powershell
     git add Database/ Color_template.xlsx (templates, slide designs, color palettes)
     git commit -m "feat(templates): add master templates, slide designs, and color palettes"
     ```
   - **Stage 3 (Automation Launchers & Agent Frameworks)**:
     ```powershell
     git add agents/ *.bat *.sh (multi-agent coordination, runners, launchers)
     git commit -m "feat(agents): add execution orchestration, watcher, and launcher automation"
     ```
   - **Stage 4 (Documentation & Safe Fixtures)**:
     ```powershell
     git add README.md docs/ tests/ fixtures/
     git commit -m "docs(integrate): add documentation, usage guide, and reference test fixtures"
     ```
4. **Anchor Baseline Version Tag**:
   Create an annotated Git tag to mark the official baseline:
   ```powershell
   git tag -a v1.0.0 -m "Release v1.0.0: Initial extracted baseline for MCP Server integration"
   ```
5. **Record Baseline Anchor**:
   Extract and record the Git commit SHA (`git rev-parse HEAD`), current branch, and tag name. These will be referenced in Section 29 and the Evolution Log.

### 1. Build an entry point and call graph

Find every executable entry point: batch/shell scripts, CLI, package main, service endpoint, scheduler, bot command, GUI, library API, and test helper. Trace what each actually calls, in order, including dynamically imported local modules.

For every entry point record:

| Entry point | Trigger | Calls | Inputs supplied | Working directory | Browser/network use | Side effects | Is this the requested workflow? |
| --- | --- | --- | --- | --- | --- | --- | --- |

**Do not equate a launcher with an entire workflow.** A batch file may only invoke a processor and open an output folder, while a separate exporter creates its inputs. Identify acquisition, processing, delivery, and presentation as separate stages and show their dependencies.

Record the exact callable functions containing business rules. Identify CLI/UI/bootstrap steps that can be replaced by MCP orchestration without changing the business result. If the main function catches errors and still exits successfully, identify a lower-level call path that exposes those failures.

### 2. Reconstruct successful and failure paths

Write a numbered, end-to-end trace from prerequisite acquisition through final delivery. For each step include its caller, input, output, state change, failure behavior, and evidence source. Trace both the normal branch and relevant flags or modes; do not describe a path that was never inspected.

Create a **stage matrix**:

| Stage | Prerequisites | Exact function/action | Output or state | Success evidence | Failure/partial behavior | Retry unit |
| --- | --- | --- | --- | --- | --- | --- |

State which stages must finish before the next begins. If reports are exported for multiple sites, factories, stations, users, or periods, document the loop order and completion criteria **per item and for the whole batch**. Explain whether one item failing permits the remaining items to continue.

### 3. Extract the true input and data contracts

For every real input and option record type, required status, default, source, validation, accepted values, and where consumed. Distinguish a user-provided parameter from an internal constant or a UI-only option. Record configuration precedence only when confirmed.

For every file, API payload, table, or message consumed, capture the **machine-verifiable schema**:

- filename pattern and exact period/site identifier rules;
- format, encoding, sheet/table name, header row, required columns, and relevant types;
- optional fields and their explicit fallback behavior;
- date format, timezone, inclusive range, fiscal/ISO period rule, and year rollover;
- grouping keys, joins, deduplication, ordering, filters, and aggregate semantics;
- whether totals repeat across rows and how double counting is avoided;
- validation of actual file content against filename, selected site, and requested period;
- cardinality: required count per site, report type, period, and entire batch;
- behavior for missing, duplicate, zero-row, corrupt, mislabeled, mixed-period, or wrong-site files.

Do not replace a missing column with zero, a current date, or a generic model unless the original business rule explicitly does so. State exact failure behavior when no valid fallback exists. An exporter and a processor must have a verified shared contract: document the export filename/layout and the processor's discovery pattern together, including suffix/prefix assumptions. Provide a small **schema crosswalk** when upstream and downstream names differ.

For confidential data, give column names and rules only. Use synthetic row shapes if an example is necessary.

### 4. Inventory resources and provenance

Inspect imports, manifests, lock files, startup scripts, runtime configuration, browser setup, services, bundled assets, templates, fonts, certificates, database files, native executables, and local helper modules.

For each requirement record:

| Requirement | Category | Required for which stage | Version or format | Provisioning source | Resolution base | Runtime location | Missing behavior | Evidence |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |

Categories include Python package, system binary, static asset, generated asset, external source checkout, environment variable, secret reference, browser session, network service, OS feature, and configuration. Distinguish import name from package name when they differ. Never copy real secret values into the file.

For a template or source checkout that AI #2 will not have, either describe its **complete reconstruction inside the Markdown** (Mode A) or include the exact required file in the ZIP (Mode B). Record its relative path, purpose, size, SHA-256, and load point in the Markdown manifest. A mere path, repository link, manifest entry, sheet list, or package name is not enough. If an indispensable asset cannot be reconstructed or bundled safely, mark implementation BLOCKED. A path that happens to exist on the extractor's computer is **not** a portable deployment plan.

### 5. Map paths, files, and side effects

Search all reads, writes, glob patterns, downloads, archive operations, moves, deletes, overwrites, template updates, temporary directories, subprocess outputs, logs, and browser profile changes. Resolve each path relative to its **actual base**: script directory, current working directory, configuration directory, input directory, or temporary directory.

Use a lifecycle table:

| Path/pattern | Stage | Origin | Read/write/move/delete | Fixed or user-controlled | Collision behavior | Retained on success | Retained on failure | Evidence |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |

Explicitly identify operations that mutate source input, master templates, prior-period outputs, databases, browser state, or external systems. Include actions hidden in helper functions and in a launcher. Note whether the source archives other periods, removes legacy files, updates its own template, opens a folder, or closes a tab.

For each destructive or source-mutating action, recommend one of: preserve as required behavior with an explicit target and guard, run on an isolated copy, or omit as presentation-only behavior. Explain the choice and what evidence proves no needed output is lost. Specify staging path rules, approved storage, cleanup on success/failure/cancellation, atomic publication where appropriate, and symlink/path-boundary checks. Do not recommend copying executable source from an arbitrary user-controlled directory.

### 6. Analyze ingestion only when the workflow needs it

For every raw input identify how it reaches the standalone processor: portal export, API, email, shared drive, database, scheduled drop, user-selected local file, or unknown. **Do not invent a browser companion for a tool whose supported contract is genuinely local-file input.** Conversely, do not call a manual file drop an automated end-to-end workflow if the requested target experience includes acquisition.

If acquisition is required, specify a companion tool or stage only from evidence. Inspect an existing target tool before recommending reuse. If absent, specify its input/output contract and mark unobserved portal details UNKNOWN. Record the dependency from acquisition to processing: exact filenames, folder layout, report type, period, site, and expected file count.

For browser automation, when applicable, document:

- authentication/session ownership and whether a signed-in browser or CDP port is required;
- navigation, report filters, station/site selection, date commit/recheck, export action, and download completion check;
- page/context/tab lifecycle, spawned pages and popups, and which tab may be closed;
- ordering and waits that prevent closing the shared browser before the next site;
- retries per report, cancellation, partial batch outcomes, and preservation of successful files;
- stable file size, nonzero content, ZIP/format checks, and file-to-site/period verification.

A successful batch requires every required item to reach a verified terminal state. A partial batch must remain partial even if some documents exist.

### 7. Capture business logic and output semantics

Document transformations as reproducible rules, not labels such as “create pivot” or “update database.” Include exact source columns, normalizations, comparison keys, formula definitions, ranking/tie rules, rounding, sheet names, chart/slides, cell/range updates, styles that carry meaning, and treatment of empty or unmatched groups. For complex logic, include pseudocode or a decision table with references to source functions.

For each output specify filename, location, format, required sheets/sections, nonempty conditions, links/charts/formulas, relationship to the input period, overwrite/versioning behavior, and intended destination. Distinguish:

- source downloads;
- intermediate or audit files;
- combined/analysis output;
- updated master database;
- presentation;
- local synced-folder copy;
- verified cloud upload or external delivery.

Do not equate a local OneDrive/SharePoint folder copy with a verified remote upload. Do not claim a presentation was generated if the program only copied a template. Do not claim all rows were handled if code truncates or samples output.

### 8. Define terminal success and failure

Identify every swallowed exception, warning-only failure, fallback, silent skip, broad catch, timeout, and success print that can mask an incomplete run. Describe durable state, process exit, logs, and artifacts separately. Define a **success predicate** with checks for every required stage. Typical checks may include expected file count, nonzero stable size, valid workbook/package, required sheets, expected row coverage, formula/chart presence, database update, and an explicit delivery result. Use only checks relevant to this project.

For each failure record detection point, error meaning, user-visible result, remaining artifacts, retry scope, cleanup, and whether rerun is safe. Distinguish FAILED, PARTIAL, CANCELLED, and SUCCESS. If the standalone returns success despite a failed substep, the spec must recommend an MCP adapter that promotes that substep failure to a non-success status. Never fabricate completed stages from stdout or UI cards.

### 9. Analyze state, concurrency, and platform behavior

Record locks, shared fixed paths, template/database mutation, caches, browser sessions, concurrent requests, scheduled overlap, file handles, Windows permission/locking behavior, and restart/config reload needs. Specify whether parallel runs can safely share a browser or output location. Where isolation is recommended, describe what must be copied and how results are published back without corrupting existing outputs.

Record OS-specific requirements and subprocess arguments, working directory, environment, timeout, stdout/stderr handling, shell use, and process cleanup. Separate a Python package dependency from a native executable. Note any startup script behavior required to reproduce the working environment.

### 10. Map to the target MCP project

**When [TARGET_MCP_ROOT_OR_UNAVAILABLE] is accessible**, inspect current source rather than relying on remembered schemas or documentation. Find the existing purpose-built tool, gateway, registry definition, Pydantic/input schema, config model, path policy, permissions, launcher, routes, bot/web handlers, scheduler, and tests. Reuse the established architecture unless a verified gap requires a new tool.

Create an integration matrix:

| Standalone stage/contract | Existing target component or verified gap | Exact adaptation | Config/path rule | Error/status mapping | Test evidence needed |
| --- | --- | --- | --- | --- | --- |

For each proposed tool, record its exact registered name or a clearly marked proposed name; input/output schema; READ, WRITE, EXECUTE, NETWORK, and any delivery permissions actually needed; confirmation/risk gates; path fields; and which gateway owns each operation. Verify that config keys are accepted by the target loader and that all supported callers use the same resolved configuration. Do not assume a newly edited config has reached an already-running service.

Choose and justify the implementation approach:

1. Import/package trusted business functions.
2. Run a pinned trusted source as a subprocess in an isolated staging directory.
3. Port the algorithm with a source-linked parity test.
4. Compose existing target tools for acquisition, processing, and delivery.

Explain why the chosen approach preserves the standalone behavior and how it avoids arbitrary code execution, source-template mutation, hardcoded developer paths, duplicate logic, and silent success. Define the source and asset reconstruction mechanism for a fresh installation **from the chosen handoff artifact alone**. In Mode A, include the exact algorithm and all required reconstruction details in Markdown. In Mode B, point to the exact bundled source/asset files and describe how they are used. A recommendation to “inspect the original project later” is not a complete handoff.

Map **every supported caller** that invokes the workflow, including MCP, web portal, bot, scheduler, CLI, or API where present. Show how each gets the same config and gateway; call out hardcoded relative paths or bypasses. Identify whether a running service must reload after code/config changes. Do not claim live readiness from source files alone.

**When the target is unavailable**, provide only conceptual roles, required capabilities, and validation tasks for AI #2. Do not invent concrete target APIs, tool names, or config keys.

### 11. Protect data and execution boundaries

List risks only when supported by the project: input/output path escape, symlinks, unsafe deletes, arbitrary executable source path, shell injection, zip traversal, secrets in logs, browser session exposure, unauthorized email/upload, confidential data in model-visible results, and insecure temporary storage. Specify which values stay local, which may appear in structured MCP output, and which must be redacted. Avoid embedding real production values or customer/model names in the handoff.

### 12. Prove the handoff works when the two AIs cannot share a project

Assume AI #2 receives **only the chosen handoff artifact** and can inspect the target MCP repository. It cannot open the friend's computer, the original standalone checkout, private paths, prior chat, or undocumented sample files. Perform a transferability audit for every required business function, helper module, template, model, binary, and external service:

| Required component | Why it is needed | How AI #2 obtains or reproduces it | Pinned version/hash or exact rule | Safe test evidence | Transfer status |
| --- | --- | --- | --- | --- | --- |

Use one of these transfer statuses: **IN_MARKDOWN**, **RECONSTRUCTIBLE_FROM_MARKDOWN**, **IN_ZIP**, **RUNTIME_EXTERNAL_PREREQUISITE**, or **BLOCKED**. The first three mean AI #2 can build the component without opening the friend's source tree or following a private link. The fourth is limited to things that cannot be transferred, such as an authorized account, live service, signed-in browser session, or user-supplied runtime data; document setup and validation without exposing credentials. A local path, repository link, or statement that a template exists does **not** satisfy the first three statuses.

For every locally authored module needed to reproduce behavior, choose one of these forms:

1. **Mode A exact text capsule**: include complete relevant source text in Markdown, with language, original relative path, imports, source revision/hash, and placement instructions. Include all local helpers needed by that code. Remove only proven UI/bootstrap material and document each removal.
2. **Mode A behavioral reconstruction capsule**: include exact ordered operations, formulas, constants, branch/error tables, schemas, normalization, output layout, and executable safe fixture/expected assertions. It must be specific enough for independent implementations to agree. A summary or vague pseudocode is insufficient.
3. **Mode B bundled source**: include the exact necessary source files under the ZIP's standalone/ tree, preserving their relative layout. Identify every file, its hash, imported local helpers, runtime entry point, and adapter relevance in the Markdown. Do not bundle only the launcher while omitting business modules.

For every binary template, workbook, presentation, model, or other static asset, either give an exact inline reconstruction recipe in Mode A or include the authorized file under standalone/ in Mode B. A list of filenames or worksheets is insufficient when formulas, styles, charts, links, embedded objects, or macros affect behavior. Do not include secrets, confidential production data, personal browser profiles, or third-party code/assets without permission. If a required asset cannot safely be included in the chosen artifact or rebuilt from it, mark BLOCKED and state the missing component.

Each capsule or bundled-file entry in section 26 must state: component ID, original relative path, purpose, stage/callers, language or format, source revision and SHA-256, exact Markdown content or ZIP path, target placement, dependencies, and one verification command/assertion. Cross-reference IDs from the call graph and resource table. Never silently abbreviate the algorithm or omit a local helper. If Mode A would be too long, select Mode B unless the owner expressly requires Markdown only.

For Mode B, create exactly one archive with this layout, omitting folders that genuinely do not apply:

~~~text
[tool_name]_MCP_HANDOFF.zip
  [tool_name]_MCP_HANDOFF/
    [tool_name]_MCP_IMPLEMENTATION_SPEC.md
    standalone/    # minimal executable source tree, preserving relative paths
    fixtures/      # synthetic or sanitized input and expected results
~~~

The Markdown must contain a **bundle manifest** with every archived path, role, size, SHA-256, original relative path, and whether it is required at runtime or only for verification. Include dependency manifests, config examples, templates, and safe fixtures if the workflow needs them. Exclude virtual environments, caches, generated production outputs, real credentials, tokens, cookies, browser profiles, unrelated repository files, and raw confidential data. Do not include duplicate copies of the same asset without a documented reason. Keep the root folder portable: use relative paths and no references that only work on the friend's machine.

Before delivery, inspect the archive listing and extract the ZIP into a fresh temporary directory. Reject absolute paths, parent traversal, symlinks escaping the root, and case-colliding names. Recompute every manifest hash after extraction. From that extracted copy, execute available safe preflight and synthetic parity checks, then confirm no unlisted local file was needed. Clean only the verified temporary extraction directory. If this rehearsal cannot run, mark the unverified stages explicitly; do not report a successful portable package based solely on ZIP creation.

Ask the friend-side AI to run a safe synthetic or sanitized example through the **real standalone entry point**, where feasible. In the same Markdown file, record the fixture construction steps, invocation, expected output manifest, and meaningful assertions: values/invariants, file count, sheets, row coverage, formulas, charts, status, and failure case. Record what was actually executed and what remains unverified. A source test passing by itself does not prove the target implementation has parity; AI #2 must run the corresponding target test after integration.

Add a **handoff coverage table** listing each stage as REPRODUCIBLE_FROM_ARTIFACT, NEEDS_RUNTIME_AUTH_OR_SERVICE, or BLOCKED. Treat required source/assets not present or reconstructible in the chosen artifact as BLOCKED. Never use a numeric “100% complete” claim based on the existence of a specification or ZIP. Separate three claims: (1) the behavior was extracted, (2) the target can implement it from the artifact, and (3) the integrated tool was verified live. The extractor can establish the first, may establish readiness for the second, and cannot establish the third on the friend's machine.

Before classifying a cross-person handoff as READY_FOR_IMPLEMENTATION, perform a **receiver rehearsal**: stop using the original standalone checkout and use only the Markdown (Mode A) or extracted ZIP folder (Mode B). Verify that AI #2 can answer all of these: Which exact code or algorithm runs? How are required static assets obtained from the artifact? Which versions are required? What is the input/output contract? How are errors detected? How are results compared to a safe reference run? If any essential answer still says “ask the friend to inspect their project,” mark the handoff BLOCKED and list the smallest missing item. A pinned external dependency may be a runtime prerequisite, but it cannot substitute for missing locally authored business logic.

## Required Markdown output structure

Use these top-level sections in the generated single file. Write “Not applicable — evidence: …” for a genuinely irrelevant section. Do not omit a relevant section or fill it with generic text.

~~~markdown
# [Tool Name] — MCP Implementation Specification

## 1. Scope, readiness, and evidence status
## 2. Source and target project identities
## 3. Entry points and stage dependency graph
## 4. Successful runtime flow
## 5. Input contract and data schema crosswalk
## 6. Acquisition and upstream/downstream handoff
## 7. Business rules and calculations
## 8. File, template, database, and temporary lifecycle
## 9. Runtime dependencies and asset provisioning
## 10. Configuration, paths, and deployment portability
## 11. Browser, network, and subprocess behavior
## 12. Output and delivery contract
## 13. Terminal success predicate and partial/failure semantics
## 14. State, concurrency, cleanup, and idempotency
## 15. Security and confidentiality
## 16. Target MCP integration matrix
## 17. Proposed tool input and structured output schemas
## 18. Gateway and caller responsibilities
## 19. Test and verification matrix
## 20. Deployment and runtime preflight
## 21. Preserve, adapt, and omit decisions
## 22. Source-of-truth index
## 23. Assumptions, unknowns, and blockers
## 24. Cross-person transferability and handoff coverage
## 25. Reference fixture and parity assertions
## 26. Source/asset capsules and bundle manifest
## 27. Receiver rehearsal and archive verification
## 28. Definition of done for AI #2
## 29. Baseline version fingerprint and evolution tracking
~~~

### Readiness classification

Use exactly one status near the beginning:

- **READY_FOR_IMPLEMENTATION**: locally authored source behavior and required static assets are present or fully reconstructible **from the chosen artifact alone**; runtime accounts/services and their setup are identified; data contracts and parity tests are complete. This is implementation readiness, not live deployment proof.
- **READY_WITH_NONBLOCKING_UNKNOWNS**: uncertainties do not prevent faithful implementation or a safe test.
- **BLOCKED**: required behavior or a static source/asset is missing from the chosen artifact, cannot be reconstructed from it, or a required external system contract is unknown. A runtime account may be a stated prerequisite, but missing credentials must never be embedded or guessed.

The status describes **implementation readiness**, not proof that the integrated tool already works. List blocking items with ID, evidence checked, why they block, and the smallest action to resolve them. Do not label a source machine's successful run as deployable if its dependencies cannot be reproduced.

### Schema and output requirements

For proposed MCP inputs, mark each field as **CONFIRMED SOURCE INPUT**, **TARGET ADAPTER INPUT**, or **RECOMMENDATION**. Provide type, default, allowed values, business meaning, path/security handling, and whether omission is safe. Do not guess a business-critical period or site.

For structured output, specify what can safely be returned to the caller, including status, period/site coverage, artifact paths, and failure details. The output must distinguish actual exported, processed, combined, synced, and delivered stages. Do not include invented metrics or confidential raw values. Define a stable partial-result shape for batch workflows.

### Test and verification matrix

Include tests at these levels when relevant:

1. **Contract**: parser and schema, config precedence, allowed paths, source/assets present, input period and site validation.
2. **Unit**: core business functions and important branch/edge cases.
3. **Adapter**: stage ordering, injected failures, timeout, cancellation, cleanup, preservation of originals, and idempotent rerun.
4. **Cross-tool**: exporter output name/layout/columns matches processor discovery and database/presentation inputs.
5. **Caller**: MCP registration plus every web, bot, scheduler, or API route uses the same configured gateway and passes period/site/options correctly.
6. **Semantic integration**: representative safe data produces valid final artifacts with expected sheets, row coverage, formulas/charts, and no template-only false success.
7. **Runtime**: after deployment/reload, inspect the live registered tool schema and status; run a controlled end-to-end workflow only with approved inputs.
8. **Negative**: missing/duplicate/wrong-period/wrong-site/corrupt file, one item failure in a batch, swallowed downstream exception, output collision, browser tab closure, and invalid/partial delivery.

A fake gateway confirms routing, not business parity. A ZIP-valid spreadsheet confirms packaging, not correct calculations. Specify at least one meaningful output comparison or invariant for the real workflow. If a live test is impossible, mark it **UNVERIFIED** and explain why.

### 29. Baseline version fingerprint and evolution tracking

Establish an immutable baseline record so that future upstream source modifications can be tracked and updated incrementally:

1. **Baseline Version & Provenance**: Record the exact Git commit SHA, tag, or folder timestamp from which this extraction was performed.
2. **Cryptographic Fingerprint Manifest**: Provide a table listing every extracted source script, asset, and template with its SHA-256 hash and byte size.
3. **Core Contracts Snapshot**: Document the baseline schema parameters, Excel sheetnames, chart series, and deliverables filenames.
4. **Living Evolution Ledger**: Initialize or update `docs/tools/versions/[tool_name]_EVOLUTION_LOG.md` with the baseline entry.
5. **Future Evolution Protocol**: Explicitly reference [`MCP_TOOL_DELTA_EXTRACTOR.md`](MCP_TOOL_DELTA_EXTRACTOR.md) as the required prompt for any subsequent incremental update when the source repository changes.

## Final extraction audit

Before saving the file, verify all of the following:

- Every function called by the requested entry point has been traced far enough to identify its business effects and hidden dependencies.
- Acquisition, processing, database/presentation generation, and delivery are separated; required companion stages are connected.
- Upstream and downstream file counts, names, schemas, site identifiers, and period rules match.
- Absolute paths, relative path bases, template locations, package dependencies, and config reload behavior are explicit.
- Source mutations and cleanup are accounted for, with an isolation or preservation strategy.
- Every required output has a content-level verification rule; no caught error can yield a false SUCCESS.
- All supported callers and browser/session lifecycle rules are represented.
- Current source behavior and MCP adaptations are clearly distinguished.
- Every nontrivial claim has a source reference; unresolved critical facts are blockers.
- The chosen artifact can be used without this prompt, previous chat context, or access to the original standalone checkout. In Mode B, every needed transferable source/asset file is inside the single ZIP. Authorized live services and user-provided runtime credentials remain external prerequisites.
- Every required component has a transfer status, a reproducible acquisition/implementation route, and a safe reference assertion; essential missing source or assets force BLOCKED.
- The final handoff is exactly one artifact: a complete Markdown file, or one ZIP with a top-level folder containing that Markdown and its manifested supporting files.

## Delivery rule

If file creation is available, write and return only the selected artifact: [tool_name]_MCP_IMPLEMENTATION_SPEC.md (Mode A) or [tool_name]_MCP_HANDOFF.zip (Mode B). In Mode B the Markdown lives **inside** the ZIP; do not deliver a second loose copy. Keep the conversational response to a link to the artifact, its mode, and a short readiness note. If archive creation is unavailable but Mode B is required, do not pretend Mode A is complete: report BLOCKED and name the packaging limitation. If file creation is entirely unavailable and Mode A is complete, begin the response with:

~~~text
FILE: [tool_name]_MCP_IMPLEMENTATION_SPEC.md
~~~

Then output the complete Markdown file and nothing else. Do not output implementation code for the target MCP project unless the user expressly asks for implementation in addition to extraction. Do not claim the target integration is complete from an extraction handoff.

## Project to analyze

Analyze [STANDALONE_ROOT] and, when accessible, [TARGET_MCP_ROOT_OR_UNAVAILABLE]. Apply [OWNER_NOTES_OR_NONE] and the selected handoff mode. Produce either the complete single Markdown specification or one ZIP containing that specification and all necessary transferable files for [TOOL_NAME_OR_INFER].


---

## Appendix: Turnkey Git Baseline Initializer Script (`init_git_baseline.py`)

If the standalone project has not yet initialized Git or has uncommitted changes, the Extractor AI or tool author can run this turnkey Python script directly inside `[STANDALONE_ROOT]` to automatically configure `.gitignore`, perform granular semantic staging, commit with Conventional Commits, and create the baseline tag:

```python
"""
Turnkey Git Baseline Initializer for Standalone Quality Tools.
Initializes Git, writes standard .gitignore, commits in 4 structured stages, and tags v1.0.0.
"""

import os
import subprocess
from pathlib import Path

ROOT = Path.cwd()

GITIGNORE_CONTENT = """# Python & Virtual Environments
.venv/
venv/
env/
__pycache__/
*.py[cod]
.pytest_cache/

# IDEs & System files
.idea/
.vscode/
.DS_Store
Thumbs.db

# Temporary Office lock files
~$*.xlsx
~$*.pptx
~$*.docx

# Output & Execution Artifacts
*.log
*.tmp
output/
temp/
tmp/
.stage-*/
_build_*/
*.zip

# Secrets & Credentials
*.env
*.key
*.pem
"""

def run_cmd(cmd: str):
    print(f">> {cmd}")
    res = subprocess.run(cmd, shell=True, capture_output=True, text=True, cwd=ROOT)
    if res.stdout.strip():
        print(res.stdout.strip())
    if res.returncode != 0 and res.stderr.strip():
        print(f"WARN: {res.stderr.strip()}")
    return res

def main():
    print(f"[*] Initializing Git Baseline in: {ROOT.resolve()}")
    
    # 1. Initialize git if missing
    if not (ROOT / ".git").exists():
        run_cmd("git init")
    
    # 2. Write .gitignore if missing or needs update
    gitignore_path = ROOT / ".gitignore"
    if not gitignore_path.exists():
        gitignore_path.write_text(GITIGNORE_CONTENT, encoding="utf-8")
        print("[+] Created standard .gitignore")
    
    # Helper to check if there are staged changes
    def has_staged():
        res = subprocess.run("git diff --cached --name-only", shell=True, capture_output=True, text=True, cwd=ROOT)
        return bool(res.stdout.strip())

    # 3. Stage 1: feat(core)
    core_files = [f.name for f in ROOT.glob("*.py") if not f.name.startswith("test_")]
    if core_files:
        for cf in core_files:
            run_cmd(f"git add {cf}")
        if has_staged():
            run_cmd('git commit -m "feat(core): implement core algorithms and data transformation logic"')

    # 4. Stage 2: feat(templates)
    template_paths = ["Database", "Color_template.xlsx", "templates"]
    for tp in template_paths:
        if (ROOT / tp).exists():
            run_cmd(f"git add {tp}")
    if has_staged():
        run_cmd('git commit -m "feat(templates): add master templates, slide designs, and color palettes"')

    # 5. Stage 3: feat(agents)
    agent_paths = ["agents", "run_*.bat", "Run_*.bat", "*.sh"]
    for ap in agent_paths:
        run_cmd(f"git add {ap}")
    if has_staged():
        run_cmd('git commit -m "feat(agents): add execution orchestration, watcher, and launcher automation"')

    # 6. Stage 4: docs(integrate) & fixtures
    remaining_paths = ["README.md", "docs", "tests", "fixtures", ".gitignore"]
    for rp in remaining_paths:
        if (ROOT / rp).exists():
            run_cmd(f"git add {rp}")
    if has_staged():
        run_cmd('git commit -m "docs(integrate): add documentation, usage guide, and reference test fixtures"')

    # Any remaining uncommitted files
    run_cmd("git add -A")
    if has_staged():
        run_cmd('git commit -m "chore: track initial project setup and supporting files"')

    # 7. Tag v1.0.0 if not already tagged
    res_tag = subprocess.run("git tag -l v1.0.0", shell=True, capture_output=True, text=True, cwd=ROOT)
    if "v1.0.0" not in res_tag.stdout:
        run_cmd('git tag -a v1.0.0 -m "Release v1.0.0: Initial extracted baseline for MCP Server integration"')
        print("[+] Created tag v1.0.0")

    # 8. Show summary
    res_head = subprocess.run("git rev-parse HEAD", shell=True, capture_output=True, text=True, cwd=ROOT)
    commit_sha = res_head.stdout.strip()
    print("\n" + "="*65)
    print("✅ GIT BASELINE ESTABLISHED SUCCESSFULLY!")
    print(f"📌 Baseline Commit SHA: {commit_sha}")
    print(f"🏷️  Baseline Tag: v1.0.0")
    print("👉 Bây giờ dự án đã sẵn sàng cho MCP_SPEC_EXTRACTOR và")
    print("   tương thích 100% với MCP_TOOL_DELTA_EXTRACTOR trong tương lai.")
    print("="*65 + "\n")

if __name__ == "__main__":
    main()
```
