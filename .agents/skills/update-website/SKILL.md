---
name: update-website
description: |
  Regenerate this Jekyll website's content from raw_data/about.xlsx.
  Use when the user says "update the website", "update from about.xlsx",
  "I edited the xlsx", "regenerate site data", or any request to sync the
  site content (tabs: 工作经历/学术兼职/教育经历/荣誉奖项/期刊文章/会议文章/专利)
  with the spreadsheet. Triggers on: xlsx updated, content refresh before deploy.
---

# update-website

This site is data-driven: `raw_data/about.xlsx` is the single source of truth.
The generator converts sheets 2..last into `_data/about.json`; Liquid templates
(`_layouts/section.html`) render all tabs from it. Never hand-edit
`_data/about.json` — always regenerate.

## Workflow

1. Run the generator from the repo root:
   ```bash
   uv run --with openpyxl .agents/skills/update-website/scripts/build_data.py
   ```
2. Check the printed counts against expectations (工作经历/兼职/教育/荣誉/期刊/会议/美国专利/中国专利).
3. Optionally build locally to verify: `bundle exec jekyll build` (or `jekyll build`).
4. Tell the user to review `git diff _data/about.json`, then commit & push to deploy via GitHub Pages.
   Do NOT commit/push unless the user explicitly asks.

## Data contract (consumed by _layouts/section.html)

Keys in `_data/about.json`: `experience`, `appointments`, `education`, `honors`,
`journals`, `conferences`, `thesis`, `patents_intl`, `patents_cn`.
Field names per section are defined in `scripts/build_data.py` — keep the script
and the layout in sync when changing either.

## xlsx conventions relied upon

- Row 1 = merged sheet title, row 2 = header, data starts at row 3.
- `—` cells mean empty; `至今` means "present".
- In 期刊文章/会议文章 the last two columns (英文引用 / 中文说明) are swapped in
  data rows vs. the header; the script detects them by CJK content, not position.
- 会议文章 may contain a PhD-thesis row with 编号 `T`.
- 美国专利 英文名称 may embed inventors; the script extracts the quoted title.
