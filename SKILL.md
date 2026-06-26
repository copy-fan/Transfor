---
name: feishu-base-from-excel
description: Convert Excel workbooks into Feishu/Lark Base multi-dimensional tables using lark-cli. Use when the user asks to turn an .xlsx/.xls spreadsheet into a Feishu Base, import workbook sheets into Base tables, infer fields from Excel columns, create useful views/dashboards for marketing or operations analysis, or repeat the "Excel file to Feishu multidimensional table" workflow.
---

# Feishu Base From Excel

## Overview

Use this skill to convert a local Excel workbook into a Feishu Base. It supports both raw single-sheet Excel files and workbooks with one main data sheet plus optional helper sheets such as dashboards, dictionaries, raw data, or view suggestions. The user should not need to manually rebuild a normal Excel file into a multi-sheet template before conversion.

The reusable script is `scripts/feishu_base_from_excel.py`. Prefer the script for the conversion so field inference, batching, and verification are consistent.

## Prerequisites

- `lark-cli` must be installed and authenticated with Base scopes.
- Use the bundled Python from `load_workspace_dependencies` when available, because it includes spreadsheet libraries.
- If the request references a file path, verify the file exists before starting.

Useful auth checks:

```bash
lark-cli auth status
lark-cli base +table-list --base-token <token>
```

## Workflow

1. Run a safe preview:

```bash
python3 scripts/feishu_base_from_excel.py --excel /path/to/file.xlsx --prepare-only
```

2. Review the preview:

- Primary sheet chosen for the main Base table.
- Header row detected for the primary sheet. The script scans the first rows so title notes above the table do not become fields.
- Field type inference.
- Planned auxiliary tables.
- Planned marketing/operations views and dashboards.

3. Run the import:

```bash
python3 scripts/feishu_base_from_excel.py --excel /path/to/file.xlsx
```

Optional flags:

- `--base-name "Name"`: override the Base title.
- `--primary-sheet "Sheet"`: force the main data sheet.
- `--header-row N`: force the primary sheet's header row if auto-detection picked the wrong row.
- `--work-dir /path/to/dir`: save generated payloads and logs somewhere specific.
- `--no-views`: skip auto-created views.
- `--no-dashboards`: skip auto-created dashboards.
- `--lark-cli /path/to/lark-cli`: use a specific CLI binary.

4. Verify:

- Confirm the script reports `ok: true` and a Feishu Base URL.
- Read back table list and sample records if the user needs extra assurance.
- Report the Base link and summarize created tables/views/dashboards.

## Conventions

- Treat the primary sheet as the main operational table. Prefer sheet names containing `飞书导入`, `导入主表`, `主表`, `明细`, `数据`, or `raw/main` when no sheet is specified.
- Auto-detect the header row for every imported sheet. This allows files with a title, date range, or notes above the real table.
- Import other non-empty sheets as auxiliary tables.
- Preserve IDs as text.
- Use number fields for numeric metrics, currency style for amount/cost/GMV/revenue fields, percentage style for rate fields, datetime for date/time columns, single-select for low-cardinality categories, and multi-select for tag-like columns.
- Write records in batches of at most 200 rows.
- Use text fields for owner/person columns unless real Feishu user IDs are available.
- For marketing/material analysis workbooks, create action views when matching columns exist:
  - high potential: `ROI >= 2` and spend/cost `>= 50`
  - low efficiency warning: spend/cost `>= 100` and `ROI < 1`
  - tag completion queue: script style/hook/tag fields are empty or marked `待补充`
- Create dashboards only after tables and fields are verified. Use Base dashboard blocks with `table_name` and field names, not IDs.

## If Something Fails

- If Base creation succeeds but later writes fail, do not recreate blindly. Reuse the reported `base_token` and inspect existing tables.
- If field creation fails due to enum/style validation, fix generated field JSON and retry only the failed creation step when possible.
- If the detected header row is wrong, rerun preview/import with `--header-row N` for the primary sheet instead of asking the user to manually remake the Excel file.
- If `@file` payload paths are rejected by `lark-cli`, run commands from the generated work directory and use relative `@filename` paths.
- If search/list permissions are missing, continue with known base tokens and direct Base APIs; do not request extra scopes unless the user asks for search.
