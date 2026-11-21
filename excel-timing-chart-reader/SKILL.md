---
name: excel-timing-chart-reader
description: Read timing charts drawn in Excel in any style (shapes, line waveforms, cell borders, coloured cells, bus hexagons, arrows), separate them from other content, recover missing arrows and labels, and explain what the chart describes.
---

# Reading timing charts drawn in Excel

Use this whenever an Excel workbook (.xlsx / .xlsm) contains a **timing chart**: signals or activities drawn
against time. Examples are digital timing diagrams, handshakes, bus cycles, power-up or reset sequences,
and step-by-step operation timelines. People draw these in very different styles, often on tabs that also
hold tables, notes, legends and other drawings, and the drawings are often **incomplete**: arrows
missing between edges, arrows that miss the edge they point at, no time scale, unnamed signals.

The goal is to **find out what the chart is really describing**: which signals exist, what each one does
over time, which transition causes which, the delays and widths involved, and what the drawing leaves
out. Then report it accurately, keeping what is drawn separate from what is inferred.

## The method: every style becomes the same model

Whatever the drawing style, the bundled script converts it into one model:
**lanes (signals) → states over time → transitions (events) → relations between transitions.**

| what is drawn in Excel | what it becomes |
|---|---|
| rectangle / bar in a lane | an active interval ("on"); the text inside is its name or value |
| bars touching end to end, hexagons, parallelograms | bus values; each boundary is a value change |
| two parallel lines (top + bottom rail) with X crossings | a bus; each crossing is a value change |
| horizontal line at the top / bottom / middle of a lane | high / low / mid level (Z or idle) |
| vertical or slanted line between levels | a rise or fall; a slant's width is the transition time |
| cell borders (top, bottom, left, right, diagonal) | treated exactly like drawn lines |
| a run of coloured cells in a row | an active interval |
| hatched fill | unknown / don't-care |
| arrow starting or ending at a transition | cause → effect between the two transitions, with its delay |
| double-headed arrow | a measured parameter, named by its label (tSU, tPD …) |
| long vertical line across lanes | a time marker (T0, T1 …) |
| text on or next to an arrow | that arrow's label |
| other text boxes and callouts | notes / requirements |
| legend swatches | what each colour means |

Lanes are found as horizontal bands of drawn material, then named from label cells or text boxes on
the left. Time comes from a tick row of numbers, text with units ("10ns"), dates or categories (C0,
Mon, Q1). The script tries whether each tick marks the start, centre or end of its cell, and keeps
whichever puts the most drawn edges on tick values.

## When the chart is incomplete

The script is built for imperfect drawings. **Always keep drawn and inferred information apart when
you report.**

- **Missing arrows.** Every transition without a drawn arrow gets likely causes, each with a reason and
  confidence:
  - **high**: same lanes, edges and delay as a drawn arrow
  - **medium**:
    - the same kind of hand-off as a drawn arrow (e.g. end → start with delay 0)
    - a step of a request/acknowledge handshake
    - follows the active edge of a clock
    - happens at the same instant
  - **low**: the nearest earlier transition on another lane, when nothing better exists

  An ending pulse with no other cause is explained as the end of its own duration. Inferred links are
  suggestions to weigh, not facts. Use the signal names and the user's domain to decide which make sense.
- **Arrows that miss.** Ends are snapped to the nearest transition in that lane (within `--snap-time`),
  and every snap is reported. Ends pointing into the middle of a bar or at nothing are listed.
- **No time axis.** Times become column numbers. If there is a clock, every transition also gets its
  clock-cycle number. Look for a scale anywhere (a header cell, a note, another tab) or ask the user.
  Then re-run with `--time-row N` or `--scale 10@C`.
- **No signal names.** Lanes are named `lane1@rows …`. Look for names in shape text, notes or other tabs,
  or ask.
- **Gaps** (stretches of a lane with nothing drawn) are reported as unknown state, never assumed.
- **Transitions with no cause and no effect** are listed. They are external inputs, durations elapsing,
  or a missing part of the drawing.

## Setup

1. The script is bundled with this skill at `scripts/xlsx_timing_reader.py`. Below, `READER` means that
   path (for example `python "<skill folder>/scripts/xlsx_timing_reader.py"`). Do not edit it.
2. Python 3.8+ standard library only. `matplotlib` is optional and draws the PNGs. Without it, add `--no-plot`.
3. Work on a copy of the user's workbook and write outputs to your working directory. The script only reads
   the workbook.
4. `.xls` / `.xlsb`: convert first, with `soffice --headless --convert-to xlsx FILE`.
5. `examples/timing_styles.xlsx` holds sample charts in several styles (bars with few arrows, line waveforms,
   cell-border waveforms, coloured cells, unlabelled lanes). Use it to see what correct output looks like.

## Workflow

### 1. Inventory every tab
```
python READER inventory BOOK.xlsx
```
This lists every tab, including hidden ones, with its cells, coloured cells, bordered cells, merged
ranges and drawing objects. A tab with only bordered or coloured cells can still hold a timing chart.

### 2. Find the timing charts and set everything else aside
```
python READER scene BOOK.xlsx --out out
```
Each tab's drawing is split into **regions**. Each region gets scored hints with evidence: `timing chart`,
`legend`, `note`, `sequence diagram`, `other diagram (flowchart / block diagram)`, `native Excel chart`,
`picture` or `table formatting`. Lanes are grouped by nearness, so a chart stays in one region even
with no arrows joining its lanes.

**Open every `overview_<tab>.png`**. Each shows all objects at their cell positions, the cell text in
grey, and a dashed box per region. If LibreOffice is available, also render the real look
(`soffice --headless --convert-to pdf BOOK.xlsx`, then `pdftoppm -png -r 80`).

Fixing the split:
- Two charts merged → smaller `--gap-rows` / `--gap-cols`, or select by cell range with `--box`.
- One chart split → larger gaps, or combine regions: `--region R1,R2`.
- Table styling taken for bars → `--no-cells`. Table grid lines taken for waveforms → `--no-borders`.

### 3. Read the cells around the chart and on other tabs
```
python READER cells BOOK.xlsx --sheet "Tab" --range A1:Z40 [--fills] [--borders]
```
Look for:
- the title and the time unit (often in the tick-row header)
- signal names and their descriptions
- spec tables (parameter, min, typ, max)
- notes and requirements ("ACK within 10 ns")
- revision info

The `note` regions from step 2 belong here too.

### 4. Extract the chart
```
python READER timing BOOK.xlsx --sheet "Tab" --region R1 --unit ns --out out
```
It writes `timing_*.md` (read this), `timing_*.json` (all numbers) and `timing_*.png` (the chart
re-drawn from the extracted data: black = drawn arrows, orange dashed = inferred links). The md contains:
- **time axis** used, and which tick alignment was chosen and why
- **lanes** top to bottom, with the source of each name, the drawing style (bars / waveform lines / bus /
  bus with two rails) and every state with its start and end
- **transitions in time order**, numbered, each with what causes it (`ARROW from …`, `inferred: …`,
  `end of its own pulse`, `(clock)` or `no cause found`) and what it drives
- **arrows drawn**: the transitions they connect and the delay
- **measured intervals**: double-headed arrows and their values
- **marker lines**: the lanes they cross and the transitions at the same instant
- **clocks**: period, high time, duty cycle
- **setup / hold** of every lane relative to each clock's active edge
- **request/acknowledge patterns**
- **free text and callouts** on the chart
- **missing or unclear**: no axis, unnamed lanes, gaps, snapped or loose arrows, backwards arrows,
  overlaps, uneven clocks, transitions no arrow explains

Options for awkward styles:
- `--time-row N`: the row that holds the ticks.
- `--scale 10@C`: no ticks, 10 units per column with 0 at column C.
- `--no-ticks`: ignore tick rows and use column numbers.
- `--tick-align auto|start|center|end`: where in its cell each tick label points. Override when you
  know better.
- `--label-col A`: the column with signal names. Use it when lanes touch each other: lanes holding two
  labels are split by them.
- `--clock-edge rise|fall`: the active clock edge.
- `--tol`: time tolerance. Default 3% of a tick step.
- `--snap-time`: how far an arrow end may be from a transition and still attach. Default 0.35 of a tick step.
- `--window`: how far back to look for a cause. Default 1.5 tick steps.
- `--snap`: pixels between a line end and a shape that still count as touching.

### 5. Verify before trusting
- Compare `timing_*.png` with the overview (and the real render). Lanes, levels, bars, edges and arrows
  must match. If a lane was merged, split or misread, fix it with the options above and re-run.
- Check that ticks fall where the drawing intends. Edges should land on round values. If most come out
  at .5, change the alignment.
- Every drawn object should be accounted for. Read the "Missing or unclear" list in full.
- All numbers come from the JSON. Never estimate values from pictures when the data exists.

### 6. Interpret: what is the chart saying?
Work through it like an engineer reading a datasheet diagram:
1. **Purpose.** From the title, the signal names and the overall shape, say what operation is shown.
   Examples: a bus write, a req/ack handshake, a clocked data transfer, a power-up sequence, a reset
   release, a schedule of steps.
2. **Story.** Walk the numbered transitions in order and say what happens and why, using the drawn
   arrows first, then the inferred links that make sense for these signals.
3. **Quantities.** Give delays (arrow deltas), pulse widths and value windows (state durations),
   clock period and duty, setup/hold, measured intervals and marker times. State the unit.
4. **Rules.** Compare these against spec tables and notes in the cells (min/max, "within", "before",
   "after"). Give both numbers for every violation and every tight margin.
5. **Relationships.** Note clock-to-data alignment, handshakes, signals that change together, ordering
   constraints such as "B must not start before A ends", and any overlaps.
6. **Gaps.** Say what the drawing does not show: missing arrows, and which inferred links you rely on and
   how confident you are. Also cover undefined states, missing names or units, and arrows that don't
   reach an edge. Ask the user about anything that changes the conclusion.

Common patterns to recognise:
- **Clock + data.** Data changes away from the active edge. Setup = change before the edge,
  hold = stable after it.
- **valid/ready and req/ack.** A four-phase handshake: req↑ → ack↑ → req↓ → ack↓.
- **Chip-select / enable windows.** Other activity must sit inside the window.
- **Address/data phases** of bus cycles.
- **Power and reset sequencing.** The order of rails or steps and the minimum delays between them.
- **Bus values.** Value windows, and when each value is valid.

### 7. Report
1. Start with what the chart shows: its purpose and the main sequence, in a few sentences.
2. Give a table of transitions in time order: time, signal, change, and what causes it. Mark each
   cause as drawn or inferred.
3. List the key timing parameters, with units, against the spec where one exists.
4. Give findings: violations, tight margins, contradictions.
5. Say what is missing or ambiguous in the drawing, and what you assumed (time unit, tick alignment,
   lane names, inferred links).
6. Say what else is on the tab or workbook that you set aside (legend used, notes, other diagrams).
Offer the JSON or the re-drawn PNG if the user wants to reuse them.

## Other cases
- **Sequence diagrams** (time runs downward, lifelines with message arrows) are flagged by `scene`.
  Read them from `scene_<tab>.json`: participants are the header shapes, and messages are the
  horizontal arrows in top-to-bottom order, with nearby text as their labels.
- **Native Excel charts** (e.g. a Gantt built as a stacked bar): `scene` prints every series with
  values read from the cells. If the first series is invisible, start = first series and length =
  second series.
- **Pictures** (screenshots of timing diagrams): `scene` saves the image. Read it visually and say the
  numbers are approximate.
- **Waveforms typed as characters** in cells (e.g. `‾‾\__/‾‾`): read them with `cells`. One character
  is one time step.
- **A style the script does not handle**: work from `scene_<tab>.json`, which lists every primitive
  with position, style and arrow ends. Map it onto the same model of lanes, states, transitions and
  relations, with a short script if needed. If the user has house conventions (for example "red =
  violation", "dashed = optional"), ask, then apply them through the style groups in the scene JSON.

## Limits
- Rotated shapes are read by their unrotated frame. Freeform shapes are read by their bounding box.
- Two lanes drawn on adjacent rows share the same border line, so the bottom of one is the top of the
  other. Use `--label-col` so labels split them, and check the re-drawn chart.
- Conditional-formatting colours are not stored as cell fills, so check the render for them.
- Inferred links come from geometry and timing alone. The signal meanings decide whether they are real.
