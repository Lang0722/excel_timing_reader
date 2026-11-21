# excel-timing-chart-reader

An agent skill for reading **timing charts drawn in Excel**: charts built from Excel's basic shapes,
lines, arrows, cell borders and coloured cells. It finds the chart among a workbook's other tabs, tables
and drawings, turns it into signals, states over time and transitions, links the arrows to the
transitions they connect, infers likely links where arrows are missing, and lists everything the drawing
leaves unclear.

It is domain-neutral: it reads what is drawn, and the meaning comes from the chart's own labels, notes
and tables.

## What's in here

```
excel-timing-chart-reader/      the skill
  SKILL.md                      instructions the agent follows
  scripts/xlsx_timing_reader.py the extractor (Python 3.8+, standard library; matplotlib optional)
  examples/timing_styles.xlsx   sample charts in several drawing styles
tests/smoke_test.py             runs the extractor on the examples and checks the results
dist/excel-timing-chart-reader.zip   the skill folder, zipped for upload
```

## Install

- **Claude apps:** upload `dist/excel-timing-chart-reader.zip` as a custom skill in Claude's settings.
- **Claude Code:** copy the `excel-timing-chart-reader` folder into `~/.claude/skills/`
  (or `.claude/skills/` inside a project).

## Use the script directly

```
python excel-timing-chart-reader/scripts/xlsx_timing_reader.py inventory BOOK.xlsx
python excel-timing-chart-reader/scripts/xlsx_timing_reader.py scene BOOK.xlsx --out out
python excel-timing-chart-reader/scripts/xlsx_timing_reader.py timing BOOK.xlsx --sheet "Tab" --region R1 --unit ns --out out
```

`scene` writes an overview picture of each tab with every region boxed; `timing` writes a readable
report (`.md`), all extracted data (`.json`) and a re-drawn chart (`.png`).

## Test

```
python tests/smoke_test.py
```
