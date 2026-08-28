# FORGE Format Profiles

## Core Invariant

```
ContentIntent != SemanticRole != FormatProfile != Template
```

- **ContentIntent** describes what the content is (e.g. `activity_plan`). It
  never decides the final format.
- **SemanticRole** describes what a block does (title, heading_1, body, caption…).
- **FormatProfile** describes how each role should be formatted.
- **Template** is only a DOCX skeleton for generation. A Reference DOCX is not
  a Template unless explicitly registered as one.

## Profile Sources

| Source | Meaning | Created by |
|---|---|---|
| Preset | Built-in profile | shipped in `profiles/registry.py` |
| Reference | Learned from a sample DOCX | `analyze_reference_format` → `create_format_profile(mode="reference")` |
| Custom | Explicit overrides on a base | `create_format_profile(mode="custom")` |
| Guided | Step-by-step key-difference session | `create_format_profile(mode="guided")` + `update_format_profile` |

## Inheritance and Deep Merge

Profiles support `inherits`. `resolve_profile` performs field-level deep merge
from the base profile; children only override what they explicitly provide.
Cycle detection is built in.

Example:

```json
{
  "schema_version": 1,
  "profile_id": "my_notice",
  "name": "我的通知格式",
  "description": "基于标准公文的定制通知",
  "source": "custom",
  "inherits": "official_standard",
  "rules": {
    "title": {"font": "黑体", "size_pt": 22},
    "body": {"font": "仿宋", "size_pt": 16},
    "page_number": {"alignment": "center"}
  }
}
```

## FORGE_HOME

User assets live outside the repository:

```
FORGE_HOME/
  profiles/      # persisted user profiles (JSON, schema_version 1)
  templates/     # registered user templates + manifests
  drafts/        # guided session drafts
  jobs/          # assembly checkpoint/resume workspaces
```

- `FORGE_HOME` env var overrides the default `~/.forge-docx-mcp/`.
- Saves are atomic (`tmp` + `os.replace`).
- Built-in profiles cannot be overwritten (`PROFILE_ID_CONFLICT`) or deleted.
- A user profile inherited by another profile cannot be deleted
  (`PROFILE_IN_USE`).
- Corrupted user JSON files are reported as load warnings and skipped.

## Reference Learning

`analyze_reference_format`:

1. Runs the Faithful Inspector + Semantic Annotator.
2. Collects effective formatting evidence per semantic role
   (docDefaults → style basedOn chain → paragraph pPr/rPr → run rPr).
3. Uses per-property consensus: a dominant value is adopted only when
   `dominance_ratio >= 0.70` (configured in `open_format/config.py`).
4. Missing roles are inherited from the base profile and marked
   `source="inherited"`, `evidence_count=0`.
5. Conflicting properties are reported as `needs_review` with candidates; the
   draft inherits the base for those fields.

Persisted reference profiles contain **no document text** — only formatting
statistics, sample counts, and the source file hash.

## Dynamic Template Registry

- User templates are stored under `FORGE_HOME/templates/`.
- `register_document_template` validates the package, rejects macros and unsafe
  external relationships, and for `kind=docxtpl` requires valid
  `{{ placeholders }}`.
- `generate_by_type` resolves user aliases/ids before falling back to built-in
  legacy types; built-in generation behavior is unchanged.

## Coverage Validation

A profile is considered complete when all renderable slots
(`title`, `subtitle`, `organization`, `author`, `heading_1..3`, `body`,
`caption`, `signature`, `date`, `table`, `image`) have non-empty rules after
resolution.
