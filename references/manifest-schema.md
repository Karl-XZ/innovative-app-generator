# Manifest Schema

The file `softcopyright-manifest.json` is the contract between project generation and documentation generation.

## Required Top-Level Fields

- `system_name`
- `version`
- `owner`
- `complete_date`
- `purpose`
- `industry`
- `user_groups`
- `environments`
- `technical_features`
- `modules`

## `environments`

Expected keys:

- `dev_hardware`
- `run_hardware`
- `dev_os`
- `dev_tools`
- `run_platform`
- `support_env`
- `languages`

## `modules`

Each module should include:

- `id`
- `title`
- `menu`
- `route`
- `summary`
- `algorithm`
- `input_output`
- `data_flow`
- `special_handling`
- `steps`
- `screenshot`
- `code_refs`

## `screenshot`

Expected keys:

- `file`
- `title`

Use project-relative paths.

## `code_refs`

Each code reference should include:

- `file`
- `title`
- `function_name`
- `start_line`
- `end_line`
- `explanation`
- `input`
- `output`

Optional:

- `variables`
- `line_explanations`

## `variables`

Each variable item:

- `name`
- `meaning`
- `range`

## `line_explanations`

Each line explanation item:

- `lines`
- `explanation`

## Optional Top-Level Fields

- `main_functions`
- `exports`
- `faq`
- `tests`
- `maintenance`
- `screenshots`

If missing, the generator uses conservative defaults.
