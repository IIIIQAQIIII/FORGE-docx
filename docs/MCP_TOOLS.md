# FORGE MCP Tools

Version: v1.2.0. Legacy tools are marked `[legacy]` and are kept unchanged.

All tools return JSON strings. Error paths generally include an `errors` list
and, where possible, a stable error code.

## Generation

### list_document_types
- Purpose: list built-in document types, sets, templates and the format guide.
- Parameters: none.
- Result: `document_types`, `document_sets`, `templates_on_disk`, `format_guide`, `semester_rule`.
- Example: `list_document_types()`

### get_template_schema
- Purpose: schema + example for one template file.
- Parameters: `template_name`.
- Result: `template`, `document_types`, `fields`, `example`, `notes`.
- Errors: template not found / path escapes allowed directory.

### generate_docx
- Purpose: fill a built-in DOCX template with JSON content.
- Parameters: `template_name`, `content`, `output_name`, `force_multipage_notice`.
- Result: `Created: <path>`.
- Errors: invalid output suffix, template not found, one-page notice constraint.

### generate_by_type
- Purpose: friendly-type generation; also resolves user-registered dynamic templates.
- Parameters: `document_type`, `content`, `output_name`, `force_multipage_notice`.
- Example: `generate_by_type("传统公文", {...}, "out.docx")`.

### generate_document_set
- Purpose: generate every document in a named set.
- Parameters: `set_name`, `content`, `output_prefix`, `force_multipage_notice`.

## Inspection & Editing

### inspect_document
- Purpose: paginated structural inspection for safe editing.
- Parameters: `source_path`, `offset=0`, `limit=100`, `query=None`, `roles=None`, `outline_only=False`, `max_text_chars=4000`.
- Result: `source_file_sha256`, `total_blocks`, `offset`, `limit`, `next_offset`, `blocks[]`.
- Each block: `block_id`, `block_type`, `semantic_role`, `role_confidence`, `text_preview`, `source_locator`, `metadata`.
- Example: `inspect_document(source_path="doc.docx", query="活动主题")`.

### edit_document
- Purpose: explicit source-preserving edits.
- Parameters: `source_path`, `expected_source_sha256`, `edits[]`, `output_path=None`, `dry_run=False`.
- Ops: `replace_text`, `insert_paragraph_before`, `insert_paragraph_after`, `append_paragraph`, `delete_paragraph`.
- Result: `status`, `source`, `output`, `source_sha256`, `operations`, `change_summary`, `preservation`, `warnings`, `errors`.
- Common errors: `SOURCE_CHANGED`, `EDIT_TARGET_NOT_FOUND`, `EDIT_OCCURRENCE_MISMATCH`, `EDIT_COMPLEX_BOUNDARY_UNSUPPORTED`, `EDIT_OPERATION_CONFLICT`, `EDIT_CONTRACT_VIOLATION`, `EDIT_UNSAFE_DELETE_TARGET`.

## Reformat & Assemble

### reformat_document
- Purpose: FORGE v1.2 source-preserving reformat.
- Parameters: `source_path`, `output_path=None`, `explicit_profile_id`, `explicit_format_hint`, `reference_profile_id`, `saved_profile_id`, `allow_default=False`.
- Result: `status`, `source`, `output`, `classification`, `resolution`, `format`, `operations`, `content_preservation`, `warnings`, `errors`.
- Common errors: `PROFILE_NOT_FOUND`, `CONTENT_PRESERVATION_FAILED`, `SOURCE_CHANGED`; ambiguous content → `status="needs_guidance"`.

### reformat_docx `[legacy]`
- Purpose: legacy template-based reformat. Behavior unchanged.

### assemble_documents
- Purpose: batch normalize + package-aware assemble.
- Parameters: `source_paths`, `output_path`, `explicit_profile_id`, `explicit_format_hint`, `reference_profile_id`, `saved_profile_id`, `assembly_profile_id`, `output_mode="assembled"`, `order_mode="input"`, `allow_default=False`, `checkpoint=False`, `job_id=None`, `resume=False`.
- Result: `status`, `total`, `processed`, `failed`, `target_profile_id`, `assembly_profile`, `output`, `normalized_outputs`, `items[]`, `content_preservation`, `assembly_payload_sha256`, `warnings`, `errors`.
- Common errors: `ASSEMBLY_INCOMPLETE`, `JOB_CONFIGURATION_CHANGED`, `CHECKPOINT_SOURCE_CHANGED`, `JOB_NOT_FOUND`.

## Validation & Intelligence

### validate_docx
- Purpose: content/layout/leftover-placeholder checks.
- Parameters: `docx_path`.

### fix_docx
- Purpose: safe light repairs (NBSP / trailing spaces).

### resolve_document_format
- Purpose: classifier + resolver on a content description (advisory only).
- Parameters: `description`, `explicit_profile_id`, `explicit_format_hint`, `reference_profile_id`, `saved_profile_id`, `allow_default`, `default_profile_id`.

### recommend_document_type
- Purpose: deterministic keyword-based document-type recommendation.

### get_semester_info
- Purpose: academic-year/semester workday date rules.

## Profiles & Templates

### list_format_profiles
- Purpose: list builtin + persisted user profiles with origins.
- Result: `format_profiles`, `resolved_profiles`, `origins`.

### get_format_profile
- Parameters: `profile_id`.
- Result: `origin`, `raw`, `resolved`.

### analyze_reference_format
- Parameters: `reference_path`, `base_profile_id`, `profile_id`, `name`.
- Result: `status` (`ok` | `needs_review`), `draft`, `conflicts[]`.

### create_format_profile
- Parameters: `mode` (`custom` | `reference` | `guided`), `profile_id`, `name`, `description`, `base_profile_id`, `overrides`, `intent`.
- Guided mode returns `session_id`, `draft_profile`, `questions[]` (≤5).

### update_format_profile
- Parameters: `session_id` (guided) or `profile_id` (user profile), `answers[]`, `overrides`, `name`, `description`.

### delete_format_profile
- Parameters: `profile_id`.
- Errors: `BUILTIN_PROFILE_PROTECTED`, `PROFILE_IN_USE`.

### register_document_template
- Parameters: `template_path`, `template_id`, `name`, `kind`, `profile_id`, `supported_intents`, `aliases`.
- Errors: `TEMPLATE_ID_CONFLICT`, `NO_TEMPLATE_PLACEHOLDERS`, `MACRO_REJECTED`, `MALFORMED_DOCX`, `UNSAFE_EXTERNAL_RELATIONSHIPS`, `ALIAS_CONFLICT`, `PROFILE_NOT_FOUND`.

### list_document_templates
- Purpose: list builtin + user templates with origin/kind/profile/aliases.
