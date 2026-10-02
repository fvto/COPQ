# MASTER PROMPT — EXTRACT A COMPLETED STANDALONE PROJECT INTO AN MCP IMPLEMENTATION SPECIFICATION

## Purpose

You are analyzing an **already working standalone project/tool**.

Your job is **NOT** to implement the MCP integration.

Your job is to inspect the completed standalone project, understand exactly how it works, and produce **one self-contained Markdown specification file** that can be handed to another AI working inside the `automation-tools / pc_tool_agent` repository.

That second AI should be able to read your Markdown file and implement the MCP tool with minimal or ideally zero additional investigation into the standalone project.

Think of yourself as the **reverse-engineering and specification layer** between:

```text
Completed Standalone Project
        ↓
You: analyze and extract behavior
        ↓
Standalone-to-MCP Specification.md
        ↓
AI inside pc_tool_agent repository
        ↓
Gateway + Tool Registration + Glue Code + Tests
```

The Markdown file you produce is therefore a **technical handoff contract**, not a summary.

---

# 1. PRIMARY OBJECTIVE

Inspect the provided standalone project and extract all information required to reproduce its behavior as an MCP tool.

The final Markdown specification must explain:

1. What the tool does.
2. How a successful execution works from beginning to end.
3. What must exist before execution.
4. What inputs are required.
5. What files/resources are read.
6. What files/resources are created.
7. What files/resources are downloaded.
8. What external programs are required.
9. What Python packages are required.
10. What environment variables or secrets are required.
11. What network services/endpoints are used.
12. What temporary files/directories are created.
13. What cleanup occurs.
14. What subprocesses are executed.
15. What assumptions the original program makes.
16. What errors can occur.
17. What side effects occur.
18. What behavior must be preserved when converting it into MCP.
19. What behavior should NOT be copied because it is only CLI/UI/bootstrap code.
20. What security concerns the MCP implementation must account for.
21. How the standalone concepts should map conceptually into the `pc_tool_agent` Gateway / Registry / Engine architecture.

The final document must be detailed enough that another AI does **not need to guess how the standalone project works**.

---

# 2. IMPORTANT SCOPE RULE

You are analyzing the **standalone project**, not the MCP repository.

Do not invent implementation details for `pc_tool_agent`.

You may describe what the future MCP implementation will need conceptually, but do not fabricate project-specific APIs that you cannot see.

For example, it is acceptable to write:

```text
The MCP implementation must validate `input_file` through the project's sandbox/path policy before opening it.
```

It is NOT acceptable to invent:

```python
registry.register_tool_v2(...)
```

unless that API is actually provided to you.

---

# 3. ZERO-HALLUCINATION RULE

Every important statement must be classified mentally as one of:

```text
CONFIRMED_BY_CODE
CONFIRMED_BY_CONFIG
CONFIRMED_BY_DOCUMENTATION
CONFIRMED_BY_USER
INFERRED_FROM_CODE
UNKNOWN
```

Prefer confirmed facts.

If something cannot be determined, write:

```text
UNKNOWN
```

or:

```text
IMPLEMENTATION BLOCKER
```

Do not silently fill missing information with a plausible assumption.

Examples:

If the code references:

```python
os.getenv("MY_API_KEY")
```

you may state:

```text
Environment variable required: MY_API_KEY
Evidence: source code
```

If the program loads:

```python
Path("templates/report.xlsx")
```

you may state that the file is required.

But if the repository does not reveal where `report.xlsx` originally comes from, write:

```text
Source/provisioning method of templates/report.xlsx: UNKNOWN
```

Do not invent a download URL.

---

# 4. ANALYSIS STRATEGY

Before writing the final Markdown file, inspect the standalone project systematically.

Do not analyze only the main function.

Trace the complete execution path.

At minimum inspect:

- entry points;
- CLI commands;
- Python modules;
- imported local modules;
- configuration files;
- constants;
- environment access;
- filesystem access;
- resource folders;
- templates;
- network calls;
- subprocess calls;
- third-party SDK usage;
- output generation;
- error handling;
- cleanup logic;
- tests if present;
- README/documentation if present;
- Dockerfile / docker-compose files if present;
- requirements files;
- `pyproject.toml`;
- lock files;
- shell/batch/PowerShell scripts;
- startup scripts;
- example input/output files when available.

Follow function calls across files until you understand the real business workflow.

---

# 5. FIND THE TRUE ENTRY POINTS

Identify every supported way the standalone project can start.

Examples:

```text
python main.py
python -m package
CLI command
FastAPI endpoint
Flask endpoint
desktop GUI
scheduled job
shell script
batch file
PowerShell script
library function
```

Determine which entry point actually triggers the behavior being converted to MCP.

Separate:

```text
ENTRY / UI / CLI LAYER
```

from:

```text
BUSINESS LOGIC
```

from:

```text
INFRASTRUCTURE / RESOURCE ACCESS
```

The MCP implementation normally needs the business logic and infrastructure behavior, not the original CLI presentation layer.

---

# 6. RECONSTRUCT THE COMPLETE EXECUTION FLOW

Trace one successful run step by step.

The final specification must contain a flow similar to:

```text
1. User provides ...
2. Program validates ...
3. Program creates ...
4. Program checks ...
5. Program downloads ...
6. Program loads ...
7. Program invokes ...
8. Program transforms ...
9. Program writes ...
10. Program cleans up ...
11. Program returns/displays ...
```

Do not omit preparation steps.

Especially inspect whether the program performs hidden preparation such as:

- downloading files before processing;
- copying bundled templates;
- creating a temporary working directory;
- unpacking archives;
- converting formats;
- installing or locating binaries;
- launching a browser;
- authenticating to a service;
- creating a session;
- fetching metadata;
- resolving relative paths;
- generating intermediary files;
- waiting for an external process;
- reading a `.env` file;
- loading a model;
- loading certificates;
- creating output directories.

The future MCP implementation needs to know all of these steps.

---

# 7. PRE-EXECUTION REQUIREMENTS

Create a full inventory of everything that must be available before the tool can succeed.

Classify each item as one of:

```text
USER_INPUT
STATIC_PROJECT_ASSET
GENERATED_ASSET
DOWNLOADED_ASSET
PYTHON_PACKAGE
SYSTEM_BINARY
ENVIRONMENT_VARIABLE
SECRET
NETWORK_SERVICE
LOCAL_SERVICE
DATABASE
OS_FEATURE
DIRECTORY
CONFIGURATION
OTHER
```

For every requirement determine:

- name;
- type;
- why it is required;
- where it comes from;
- expected location;
- whether it is mandatory;
- when it is checked;
- what happens if it is missing.

---

# 8. FILE AND DIRECTORY ANALYSIS

Search for all filesystem operations.

Examples include:

```python
open(...)
Path(...)
Path.read_text(...)
Path.read_bytes(...)
Path.write_text(...)
Path.write_bytes(...)
shutil.*
os.path.*
os.makedirs(...)
glob.*
tempfile.*
zipfile.*
tarfile.*
pandas.read_*
DataFrame.to_*
openpyxl.*
fitz.*
PIL.Image.open(...)
```

Build a file lifecycle inventory.

For each file or directory identify:

```text
Name / pattern
Purpose
Input / output / temporary / static asset
Created by whom
Read by whom
Written by whom
Deleted or retained
Expected location
Whether user-controlled
Whether path is relative
Whether path is absolute
Whether overwrite can occur
```

---

# 9. RELATIVE PATH AND WORKING DIRECTORY ANALYSIS

Pay special attention to code such as:

```python
Path.cwd()
"./data"
"../templates"
Path("output")
open("config.json")
```

Determine whether the tool assumes a particular current working directory.

Document that assumption explicitly.

For every relative resource, determine whether it is resolved relative to:

```text
current working directory
script directory
package directory
input file directory
configured directory
temporary directory
UNKNOWN
```

This matters because an MCP server may run from a different working directory.

---

# 10. DOWNLOAD AND RESOURCE-PREPARATION ANALYSIS

This section is mandatory if the standalone project downloads or prepares files before the main operation.

For every downloaded resource determine:

```text
Resource name
Purpose
Trigger
Source URL / endpoint
HTTP method
Authentication
Request parameters
Destination
Temporary or persistent
Cache behavior
Expected file type
Validation after download
Retry behavior
Timeout behavior
Partial-download cleanup
Failure behavior
```

Also determine whether the downloaded file is subsequently:

```text
opened directly
converted
renamed
unzipped
parsed
passed to another program
deleted
cached
```

If the source of a required resource cannot be discovered, explicitly state:

```text
RESOURCE SOURCE UNKNOWN
```

Never invent a download location.

---

# 11. PYTHON DEPENDENCY ANALYSIS

Inspect:

```text
requirements.txt
requirements-dev.txt
pyproject.toml
setup.py
setup.cfg
Pipfile
poetry.lock
uv.lock
requirements/*.txt
```

Also inspect actual imports.

Classify imports into:

### Python standard library

Example:

```text
pathlib
json
subprocess
tempfile
```

### Third-party Python packages

Example:

```text
pandas
openpyxl
requests
PyMuPDF
```

### Internal/local modules

Example:

```text
project.utils
project.converter
```

### Optional dependencies

Only classify something as optional if the code actually treats it as optional.

For each third-party package explain why it is used.

Do not guess package versions if no version constraint is present.

---

# 12. SYSTEM / NATIVE DEPENDENCIES

Search for use of external programs such as:

```text
ffmpeg
ffprobe
LibreOffice
Chrome
Chromium
Tesseract
Ghostscript
ImageMagick
Java
Node.js
PowerShell
cmd.exe
bash
7zip
pandoc
Docker
```

Inspect:

```python
subprocess.run(...)
subprocess.Popen(...)
os.system(...)
shutil.which(...)
```

For every external program determine:

```text
Executable name
Purpose
How it is discovered
Expected PATH behavior
Hardcoded location if any
Arguments passed
Input files
Output files
Exit-code handling
Platform dependency
```

Do not confuse a Python package with a system binary.

---

# 13. SUBPROCESS ANALYSIS

Document every subprocess command.

Do not only copy the command string.

Explain its semantics.

For example:

```text
Command:
ffmpeg -i input.mp4 ...

Purpose:
Convert the source video before the Python parser processes it.

Inputs:
input.mp4

Outputs:
temp/audio.wav

Failure detection:
non-zero process exit code

Cleanup:
temp/audio.wav is deleted after processing
```

Identify whether commands use:

```text
shell=True
shell=False
user-controlled arguments
hardcoded executable paths
working directory overrides
environment overrides
```

Flag potential command-injection concerns.

---

# 14. NETWORK ANALYSIS

Search for:

```text
requests
httpx
urllib
aiohttp
socket
websocket
FTP
SFTP
cloud SDKs
OpenAI SDK
Google APIs
AWS SDK
Azure SDK
database clients
browser automation
```

For each network interaction determine:

```text
Protocol
Domain / host
Endpoint
HTTP method
Purpose
Authentication
Headers
Parameters
Body
Upload/download
Timeout
Retry
Expected response
Failure behavior
```

Do not expose secret values in the final specification.

Only document secret variable names and their purpose.

---

# 15. ENVIRONMENT VARIABLES AND SECRETS

Search for:

```python
os.getenv(...)
os.environ[...]
dotenv
load_dotenv(...)
settings classes
config loaders
```

For each variable provide:

```text
Variable name
Required / optional
Purpose
Default behavior
Where it is consumed
Whether it contains sensitive information
Failure if absent
```

Never include actual secret values.

---

# 16. CONFIGURATION ANALYSIS

Inspect configuration sources such as:

```text
.env
.env.example
yaml
yml
json
toml
ini
XML
Python constants
CLI options
environment variables
database configuration
```

Determine precedence where possible.

Example:

```text
CLI option > environment variable > config file > hardcoded default
```

Only state precedence if confirmed.

---

# 17. INPUT CONTRACT EXTRACTION

Determine the real business inputs.

Ignore inputs that exist only for the original UI unless they influence behavior.

For every input include:

| Field | Type | Required | Default | Constraints | Meaning | Example |
|---|---|---|---|---|---|---|

Identify whether each input is:

```text
plain value
file path
directory path
URL
enum
boolean
list
object
credential reference
```

Document implicit constraints found in code.

Examples:

```text
file extension must be .xlsx
value must be positive
sheet must exist
URL must use https
input directory must contain *.csv
```

Do not create constraints merely because they seem reasonable.

---

# 18. OUTPUT CONTRACT EXTRACTION

Determine what a successful execution actually produces.

Separate:

```text
RETURN VALUE
GENERATED FILES
MODIFIED FILES
STDOUT / CONSOLE OUTPUT
LOGS
SIDE EFFECTS
```

If the standalone script only prints a result, describe what structured data the MCP implementation should preserve conceptually.

Example:

```text
Standalone behavior:
prints "Processed 124 rows"

Recommended MCP semantic output:
processed_rows: 124
```

Clearly mark recommendations as recommendations, not original behavior.

---

# 19. FILE OUTPUT SEMANTICS

For every generated output file determine:

```text
filename rule
extension
destination
overwrite behavior
collision behavior
directory creation
content format
encoding
whether it is final or temporary
whether caller needs the path
```

If filenames contain timestamps, random IDs, source names, or counters, document the exact rule.

---

# 20. TEMPORARY FILE AND CLEANUP ANALYSIS

Identify temporary resources.

Examples:

```text
TemporaryDirectory
NamedTemporaryFile
.tmp files
download cache
converted files
browser profile
extracted archive
staging directory
```

For each one determine:

```text
when created
where created
who uses it
normal cleanup
cleanup on error
whether stale files can remain
```

The future MCP implementation must know whether cleanup is part of correct behavior.

---

# 21. STATE AND CONCURRENCY ANALYSIS

Inspect whether the standalone tool uses:

```text
global mutable variables
singleton objects
shared cache
fixed filenames
shared output folders
database state
browser sessions
lock files
process-global configuration
```

Determine whether two simultaneous executions could conflict.

Document possible collisions.

Do not assume MCP requests are serialized.

---

# 22. ERROR AND FAILURE-MODE ANALYSIS

Trace both explicitly handled errors and likely errors from dependencies.

Identify confirmed failure conditions such as:

```text
input file missing
invalid extension
corrupt workbook
required worksheet missing
API authentication failure
HTTP timeout
binary missing
subprocess non-zero exit
permission denied
output file already exists
invalid JSON
missing environment variable
unsupported platform
```

For every meaningful failure provide:

```text
Failure condition
Where detected
Original exception / behavior
User-visible message if present
Partial outputs left behind
Cleanup behavior
Suggested stable MCP error meaning
```

Do not fabricate exact MCP `AgentError` codes unless the target project provides an error-code convention.

You may suggest semantic names, clearly marked as recommendations.

---

# 23. SECURITY ANALYSIS

Analyze only risks relevant to the project.

Consider:

```text
path traversal
absolute paths
symlink escape
user-controlled output names
file overwrite
unsafe deletion
archive traversal / zip slip
command injection
shell=True
SSRF
credential leakage
unsafe temporary files
world-readable temporary files
untrusted document parsing
macro-enabled documents
arbitrary code execution
unsafe deserialization
network downloads
certificate verification
race conditions
```

For each applicable risk explain:

```text
Current behavior
Why it matters in MCP
Required mitigation or consideration
```

Do not redesign the entire project unless necessary.

---

# 24. PLATFORM ANALYSIS

Determine supported or assumed operating systems.

Look for:

```text
Windows paths
drive letters
cmd.exe
PowerShell
registry access
COM automation
Linux paths
bash
/usr/bin
macOS commands
osascript
platform.system()
```

Report:

```text
Confirmed platform support
Platform-specific functionality
Unknown platform behavior
```

---

# 25. DATA-FORMAT ANALYSIS

Document every important format handled by the tool.

Examples:

```text
.xlsx
.xlsm
.csv
.json
.yaml
.pdf
.docx
.png
.mp4
.zip
XML
HTML
```

If format-specific behavior matters, document it.

Example:

```text
The tool preserves workbook formulas but does not evaluate them.
```

Only state this if confirmed by the code/library behavior being relied upon.

---

# 26. BUSINESS LOGIC EXTRACTION

Separate business logic from infrastructure.

Describe:

```text
BUSINESS RULES
```

independently from:

```text
FILE HANDLING
NETWORK HANDLING
PROCESS EXECUTION
CLI HANDLING
```

Example:

```text
Business rule:
Rows whose "Status" column equals "Invalid" are excluded.

Infrastructure:
Workbook is read with openpyxl.
```

This separation helps the MCP implementation preserve behavior while adapting infrastructure safely.

---

# 27. IDENTIFY CODE THAT SHOULD NOT BE PORTED DIRECTLY

Explicitly identify code that exists only because the project is standalone.

Examples:

```text
argparse setup
Tkinter UI
desktop notification
interactive input()
print formatting
CLI progress bars
terminal color handling
__main__ bootstrap
manual pause
open output folder in Explorer
```

Mark each as:

```text
DO NOT PORT DIRECTLY TO MCP
```

and explain what semantic behavior, if any, should replace it.

---

# 28. MCP CONCEPTUAL MAPPING

Without inventing project APIs, map the standalone behavior into these conceptual layers:

```text
MCP Tool Input Schema
        ↓
Engine validation / sandbox
        ↓
Gateway
        ↓
Standalone business operation
        ↓
Structured MCP result
```

Create a table:

| Standalone Concept | MCP Concept | Notes |
|---|---|---|

Examples:

```text
CLI argument        -> Tool input field
input()             -> Tool input field
print() result      -> Structured return data
sys.exit()          -> Controlled tool error
local file path     -> Sandbox-validated path
HTTP request        -> Network-capable gateway behavior
subprocess          -> Execute-capable gateway behavior
generated file      -> Sandbox-safe output artifact
```

Only include mappings relevant to this project.

---

# 29. PERMISSION REQUIREMENT ANALYSIS

Based strictly on actual behavior, report whether the future MCP tool requires:

```text
READ
WRITE
EXECUTE
NETWORK
```

A tool may require more than one semantic capability.

Do NOT force them into a single project permission if you do not know how the MCP repository resolves multi-permission tools.

Instead write, for example:

```text
Semantic capabilities required:
- READ
- WRITE
- NETWORK

Target-project permission mapping:
Must be resolved by the pc_tool_agent implementation because its permission model may allow only one primary permission.
```

---

# 30. RISK CHARACTERISTICS

Describe concrete risk characteristics rather than inventing a project-specific risk level.

Example:

```text
Risk characteristics:
- Reads user-selected workbook.
- Creates a new output workbook.
- Does not overwrite the source.
- Makes no network calls.
- Executes no subprocess.
```

The MCP-side AI can then map this to `RiskLevel` using the repository's own conventions.

---

# 31. FAKE / TEST DOUBLE REQUIREMENTS

Describe what a future fake gateway must simulate.

Examples:

```text
Successful processing
Missing input file
Invalid file format
API failure
subprocess failure
download failure
```

Identify which external dependencies should NOT be required in unit tests.

For example:

```text
Unit tests should not require:
- live network
- ffmpeg installation
- real API credentials
- production Excel files
```

---

# 32. TEST SCENARIOS TO PRESERVE

Extract useful test scenarios from:

```text
existing tests
sample data
README examples
CLI examples
edge-case code branches
```

Create a table:

| Scenario | Input | Expected behavior |
|---|---|---|

At minimum cover:

```text
normal success
required-input validation
missing dependency/resource
filesystem failure if relevant
network failure if relevant
subprocess failure if relevant
edge cases visible in code
```

Do not invent expected behavior that cannot be determined.

---

# 33. ORIGINAL PROJECT FILE MAP

Provide a concise project tree containing only files relevant to the tool.

Example:

```text
project/
├── main.py                  # CLI entry point
├── converter.py             # core conversion logic
├── downloader.py            # downloads required resource
├── config.py                # environment/config loading
├── templates/
│   └── report.xlsx          # required static template
└── requirements.txt
```

For each relevant file, explain its role.

---

# 34. SOURCE-OF-TRUTH INDEX

The final specification must include a source index so the MCP-side AI can understand where conclusions came from.

Example:

| Behavior / Requirement | Source |
|---|---|
| Input must be `.xlsx` | `src/parser.py: validate_input()` |
| Requires `ffmpeg` | `src/audio.py: convert_audio()` |
| Reads `API_TOKEN` | `src/config.py` |
| Output naming rule | `src/export.py: make_output_path()` |

Use real file names and symbols from the inspected project.

Line numbers are helpful if available but not mandatory.

---

# 35. UNCERTAINTY REGISTER

Create a dedicated section containing every unresolved issue.

Each item must use:

```text
ID:
Category:
Question:
Why it matters:
Evidence checked:
Recommended action:
```

Example:

```text
ID: U-01
Category: Runtime asset
Question: Where is `templates/master.xlsx` provisioned from?
Why it matters: The tool cannot execute without it.
Evidence checked: repository search, README, setup scripts.
Recommended action: Ask project owner or inspect deployment environment.
```

Never hide uncertainty inside prose.

---

# 36. IMPLEMENTATION BLOCKERS

Distinguish between:

```text
UNKNOWN BUT NON-BLOCKING
```

and:

```text
IMPLEMENTATION BLOCKER
```

An implementation blocker is information without which the MCP-side AI cannot reliably reproduce the standalone behavior.

Examples:

```text
Unknown API endpoint required at runtime.
Unknown required template source.
Unknown password for encrypted resource.
Critical helper source missing.
Behavior depends on an unavailable proprietary executable.
```

---

# 37. FINAL DELIVERABLE

Produce exactly **one Markdown document**.

Recommended filename:

```text
<tool_name>_MCP_IMPLEMENTATION_SPEC.md
```

The document must be self-contained.

Do not require the MCP-side AI to read this prompt.

Do not output meta-commentary such as:

```text
"I followed your prompt."
```

The Markdown file itself is the deliverable.

---

# 38. REQUIRED FINAL MARKDOWN STRUCTURE

Use the following exact top-level structure unless a section is truly irrelevant.

```markdown
# <Tool Name> — MCP Implementation Specification

## 1. Executive Summary

## 2. Integration Readiness
### Status
### Blocking Issues

## 3. Standalone Project Overview
### Purpose
### Relevant Project Structure
### Entry Points

## 4. Complete Runtime Flow
### Successful Execution Flow
### Pre-Execution Preparation
### Main Processing
### Post-Processing and Cleanup

## 5. Input Contract

## 6. Output Contract

## 7. File and Directory Lifecycle
### Input Files
### Static Resources
### Downloaded Resources
### Temporary Resources
### Generated Outputs
### Cleanup Rules

## 8. Runtime Prerequisites
### Python Version
### Python Packages
### System / Native Dependencies
### Environment Variables
### Secrets
### Configuration Files
### External Services
### Network Requirements
### OS / Platform Requirements

## 9. Download and Resource Preparation Flow

## 10. Business Logic
### Core Rules
### Important Transformations
### Validation Rules

## 11. External Process / Subprocess Behavior

## 12. Network Behavior

## 13. State, Cache, and Concurrency

## 14. Error and Failure Semantics

## 15. Security Considerations

## 16. Standalone-Only Behavior That Should Not Be Ported

## 17. Standalone-to-MCP Conceptual Mapping

## 18. Required MCP Capabilities
### READ
### WRITE
### EXECUTE
### NETWORK

## 19. MCP Gateway Responsibilities

## 20. Suggested MCP Input Schema

## 21. Suggested MCP Structured Output

## 22. Fake Gateway / Test Double Requirements

## 23. Required Test Scenarios

## 24. Dependency Manifest
### Standard Library
### Third-Party Python Packages
### System Dependencies
### Runtime Assets

## 25. Deployment / Preflight Checklist

## 26. Source-of-Truth Index

## 27. Assumptions and Inferences

## 28. Uncertainty Register

## 29. Implementation Blockers

## 30. Definition of Done for the MCP-Side Implementer
```

---

# 39. EXECUTIVE SUMMARY REQUIREMENTS

The executive summary must answer, in a few paragraphs:

```text
What does this tool do?
What are its primary inputs?
What does it produce?
What external dependencies does it require?
Does it read/write files?
Does it use network?
Does it execute external processes?
What is the most important implementation concern for MCP?
```

---

# 40. INTEGRATION READINESS STATUS

Use exactly one of:

```text
READY_FOR_MCP_IMPLEMENTATION
```

```text
READY_WITH_NON_BLOCKING_UNKNOWNS
```

```text
BLOCKED
```

Definitions:

### READY_FOR_MCP_IMPLEMENTATION

All behavior and required runtime dependencies are sufficiently understood.

### READY_WITH_NON_BLOCKING_UNKNOWNS

Some information is uncertain, but it does not prevent faithful implementation.

### BLOCKED

Critical information required to reproduce runtime behavior is missing.

Do not use `READY_FOR_MCP_IMPLEMENTATION` simply because the standalone project currently runs on one machine.

---

# 41. SUGGESTED MCP INPUT SCHEMA

Convert confirmed standalone inputs into a conceptual MCP schema.

Example:

```yaml
input_file:
  type: string
  required: true
  semantic_type: filesystem_path
  description: Path to the source Excel workbook.

threshold:
  type: number
  required: false
  default: 0.05
  description: ...
```

Do not invent fields merely to improve the API.

Clearly distinguish:

```text
CONFIRMED INPUT
```

from:

```text
RECOMMENDED MCP ADAPTATION
```

---

# 42. SUGGESTED MCP STRUCTURED OUTPUT

Convert observable successful behavior into a structured conceptual output.

Example:

```yaml
summary:
  type: string

processed_rows:
  type: integer

output_file:
  type: string
  semantic_type: filesystem_path

warnings:
  type: array[string]
```

The schema must preserve meaningful results from the standalone tool.

Do not fabricate metrics the standalone tool never calculates.

---

# 43. MCP GATEWAY RESPONSIBILITIES

Describe specifically what the future Gateway should own.

Possible examples:

```text
Validate runtime dependency availability.
Resolve sandbox-approved paths.
Load static template.
Download prerequisite resource.
Execute business transformation.
Invoke external binary.
Translate expected errors.
Clean temporary artifacts.
Return structured result.
```

Only include responsibilities required by this tool.

Also state what should remain outside the Gateway when appropriate.

---

# 44. PREFLIGHT CHECKLIST

Include a practical checklist that the MCP-side AI or developer can follow before running the integrated tool.

Example:

```markdown
- [ ] Python package `openpyxl` is installed.
- [ ] `ffmpeg` is available on PATH.
- [ ] `API_TOKEN` is configured.
- [ ] Required template `templates/report.xlsx` is packaged.
- [ ] Input and output paths are inside the MCP allowed directories.
- [ ] Network access to `api.example.com` is available.
```

Every checklist item must have evidence from the project.

---

# 45. BEHAVIOR-PRESERVATION RULE

The MCP implementation should preserve business behavior unless there is a clear security or architecture reason to adapt it.

Your specification must explicitly call out:

```text
MUST PRESERVE
```

for behavior that is semantically important.

Examples:

```text
MUST PRESERVE: output filename format.
MUST PRESERVE: Excel worksheet selection rule.
MUST PRESERVE: filtering algorithm.
MUST PRESERVE: default threshold of 0.05.
```

And:

```text
MAY ADAPT FOR MCP
```

for presentation/bootstrap behavior.

Examples:

```text
MAY ADAPT FOR MCP: CLI progress output.
MAY ADAPT FOR MCP: argparse.
MAY ADAPT FOR MCP: opening the result directory in Windows Explorer.
```

---

# 46. NO SILENT REFACTORING

Do not rewrite or "improve" algorithms in the specification.

If you notice questionable behavior, document:

```text
CURRENT BEHAVIOR
```

and optionally:

```text
IMPLEMENTATION NOTE
```

but do not redefine expected behavior.

The standalone project is the behavioral source of truth.

---

# 47. DO NOT ASSUME SUCCESS FROM README ALONE

Documentation may be outdated.

When README and source code conflict:

```text
source code behavior
```

should normally be treated as stronger evidence for runtime behavior, while the discrepancy must be documented.

Do not silently choose one.

---

# 48. USE TESTS AS BEHAVIORAL EVIDENCE

If tests exist, inspect them carefully.

Tests can reveal:

```text
edge cases
expected exceptions
input constraints
output formats
naming rules
mocked dependencies
hidden assumptions
```

Include important confirmed test behavior in the specification.

---

# 49. INSPECT DEPLOYMENT FILES

If present, inspect:

```text
Dockerfile
docker-compose.yml
GitHub Actions
startup scripts
systemd units
batch files
PowerShell scripts
Makefile
Taskfile
CI configuration
```

These frequently reveal dependencies that the Python code alone does not show.

Examples:

```text
apt install libreoffice
apk add ffmpeg
playwright install chromium
COPY templates/
ENV ...
```

Treat such evidence as runtime/deployment evidence.

---

# 50. FINAL QUALITY BAR

Before producing the Markdown file, verify that another AI working only with:

1. your generated specification, and
2. the `pc_tool_agent` repository

could answer all of these questions:

```text
What exactly am I implementing?
What parameters should the MCP tool accept?
What should it return?
Which standalone functions contain the real business logic?
What files must exist?
What files are generated?
What must be downloaded first?
Where do downloaded files go?
What packages must be installed?
What system binaries must exist?
What environment variables are required?
Which network endpoints are contacted?
Does the tool need READ, WRITE, EXECUTE, or NETWORK capability?
What paths require sandbox protection?
What subprocesses run?
What cleanup is required?
What failures should become controlled MCP errors?
What behavior must remain identical?
What CLI/UI code should not be copied?
What should a fake gateway simulate?
What test scenarios are required?
What facts are still unknown?
Are any unknowns blockers?
```

If any answer is missing but can still be discovered from the standalone repository, continue analyzing before producing the final document.

---

# 51. FINAL RESPONSE RULE

Your final response must contain the completed Markdown specification only.

Do not include:

```text
analysis notes
chain of thought
progress commentary
generic advice
unrelated refactoring suggestions
```

The specification should be written in clear technical English and optimized for an AI software engineer that will implement the MCP integration.

---

# PROJECT TO ANALYZE

Analyze the standalone project currently available to you.

Desired MCP tool name, if already known:

```text
[OPTIONAL: ENTER TOOL NAME]
```

Additional owner-provided behavior or constraints:

```text
[OPTIONAL: ENTER NOTES]
```

Now inspect the entire relevant standalone project and produce:

```text
<tool_name>_MCP_IMPLEMENTATION_SPEC.md
```

following all requirements above.
