# Fab Move-In Simulator: Handover

Status as of 4 September 2026, script version 1.13.1.

## What it is

`fab_movein.py` turns a fab layout (DXF) and a tool move-in schedule
(xlsx/csv) into one self-contained HTML file that animates tools appearing on
the floor plan day by day. It is a single Python file with two dependencies
(`ezdxf`, `openpyxl`). Everything runs on the local machine: the HTML opens by
double-clicking and works offline. There is no server, no upload, no account.

## Where everything is

| Item | Location |
|---|---|
| Source | GitHub repository `BenjaminDoering1/Fab-move-in-simulator-` (private), branch `main` |
| Current version | 1.13.1 (4 Sep 2026), on branch `claude/move-in-conflicts-storage-v9dvth` until merged into `main`. |
| The tool | `fab_movein.py`, the only file a user needs |
| Developer notes | `CLAUDE.md` (architecture, performance invariants) |
| Regression tests | `tests/run_tests.py` |
| Work folder on the laptop | `fab_movein.py` next to a `data\` folder holding `floor1.dxf`, `floor3.dxf`, `move_in_plan.xlsx` |

## Where the real-data work stands

Confirmed on the real drawings:

- **floor1**: the HTML shows all scheduled tools in place (type mode, since 1.2).
- **floor3**: 286 schedule rows matched. The first HTML (45 MB) loaded very
  slowly, played back below 1 fps and zoomed slowly. Versions 1.4 to 1.8
  targeted exactly this.
- The floor3 DXF is about 1.1 GB. `--layers` traced most of the output to one
  tool layer (about 48 MB) in which every tool block is a complete equipment
  drawing. Since 1.8 such blocks are reduced to their footprint outline by
  default.

Not yet confirmed:

- How floor3 behaves with version 1.8 or later on the laptop (load time,
  playback, zoom). This is the next thing to check.
- **DT_545R** is visible in the CAD drawing and present in the schedule but
  not displayed. The likely cause is that the label is stored differently
  inside its block (split into two texts such as `DT_545` + `R`, or with
  different punctuation). `--trace DT_545R` prints the diagnosis, including
  similar labels found in the drawing.

## Set up a laptop

1. Install Python 3.9 or newer from python.org and tick "Add to PATH".
2. In a command prompt: `pip install ezdxf openpyxl`
3. Sign in to GitHub, open `fab_movein.py` on branch `main`, and use
   "Download raw file". Put it in the work folder. Repeat whenever a new
   version is announced; `python fab_movein.py --version` prints the version.
4. Put the drawings and the schedule in `data\`.

Drawings must be DXF. From AutoCAD: SAVEAS, file type "AutoCAD 2013 DXF".
DWG cannot be read.

## Everyday commands

Run from the work folder.

```
python fab_movein.py --dxf data\floor1.dxf --schedule data\move_in_plan.xlsx --type-col Block_ID -o floor1_simulation.html
python fab_movein.py --dxf data\floor3.dxf --schedule data\move_in_plan.xlsx --type-col Block_ID -o floor3_simulation.html
```

While the schedule covers only the north half and the tools in question sit
on the layer `0_ENG 3K tool` (the same labels also exist in the south half):

```
python fab_movein.py --dxf data\floor3.dxf --schedule data\move_in_plan.xlsx --type-col Block_ID --tool-layer "0_ENG 3K tool" --prefer north -o floor3_simulation.html
```

The ENG line (upper fab, floor 3) has two areas with the move-in passageway
between them. Tools north of the passageway go in north to south, tools
south of it south to north. Both boundaries are read from the drawing (see
"Drawing anchors" below):

```
python fab_movein.py --dxf data\floor3.dxf --schedule data\move_in_plan.xlsx --type-col Block_ID --tool-layer "0_ENG 3K tool" --prefer "north-of=text:XPH-D@BAY-ID" --aisle hatch:ANSI31@00_Dummy -o floor3_simulation.html
```

## Drawing anchors (floor 3) -- READ THIS WHEN THE DXF CHANGES

The placement order depends on two objects in `floor3.dxf`. They are
deliberate, simple workarounds, not features of the DXF format, and the
script reads them fresh on every run rather than storing coordinates:

| What it marks | Where it is in the drawing | Option | If the DXF changes |
|---|---|---|---|
| The move-in passageway between the north and the south area of the ENG line | The pathway object on layer `00_Dummy`: the hatch with pattern `ANSI31` (the layer holds other objects too, so the whole layer must not be used). Its extent gives a horizontal centreline | `--aisle hatch:ANSI31@00_Dummy` | Point at the new pathway: another pattern or layer (`hatch:PATTERN@LAYER`, `--inspect` lists the hatch patterns per layer), a text on it (`text:LABEL@LAYER`), a layer holding only the pathway (`layer:NAME`), or the typed Y (`--aisle 18500`) |
| The split between the upper fab (ENG line) and the lower fab | The Y of the text `XPH-D` on layer `BAY-ID` | `--prefer "north-of=text:XPH-D@BAY-ID"` | Name another text (`north-of=text:XPH-E@BAY-ID`), a layer (`north-of=layer:NAME`) or a Y (`north-of=17800`) |

How to see that they still hold:

- The run prints each anchor it resolved, with the object it found and the
  line it derived, for example
  `Anchor : --aisle hatch:ANSI31@00_Dummy = hatch 'ANSI31' on layer '00_Dummy': 1 hatch(es), extent X ... -> horizontal line at Y = 18,412`.
  If the extent or the Y looks wrong (a second hatch with the same pattern,
  the text moved), fix the option. More than one matching hatch is warned
  about with each piece's extent; a whole layer with several objects is
  warned about too.
- A missing layer or text stops the run and lists similar layer names or the
  texts on that layer. `--inspect` lists all layers, `--inspect-full` all
  texts.
- The viewer draws both lines dashed in orange with their option as label
  (View menu, "Placement reference lines"), so a wrong line is visible on
  the floor plan.
- `--trace DT_112` marks every copy `in`/`OUT` against the anchors and prints
  the order actually used.

Then double-click the HTML file. Do not open it through a preview server
(VS Code Live Server, Jupyter, `python -m http.server`): very large files abort
there with "WinError 10053".

The run prints the columns it picked, the matching option, `Matched N of M
scheduled tools`, warnings for unmatched labels, the file size and, for large
files, the heaviest layers with the knobs that shrink them.

## How the matching works on these drawings

This is the knowledge that is easiest to lose.

- A tool in these drawings is a block reference. Copy-pasted blocks get
  anonymous definition names such as `A$C39A4F2E1`, which is what EATTEDIT
  shows. The readable label (`DT_112`) is TEXT drawn inside the block
  definition.
- Text inside a block definition is stored once in the DXF and displayed at
  every insert of that block. That is why AutoCAD FIND reports one hit for
  `DT_112` although many tools carry it.
- The labels are tool **types**, not per-tool IDs: many tools share `DT_112`.
  The schedule has one row per physical tool, and its `Block_ID` column holds
  the type. Unique location IDs do not exist yet.
- The script therefore runs in **type mode** (`--type-col Block_ID`): each
  schedule row takes one still-free placement of its type, in date order.
  Placements are consumed in a fixed positional order, so runs are repeatable.
  Rows get numbered IDs (`DT_112`, `DT_112 (2)`, ...) and the tooltip marks
  these slots as provisional. If the schedule has more rows of a type than the
  drawing has placements, the shortfall is reported per type and the surplus
  rows are listed as unmatched.
- **Which copy of a repeated label a row takes** (1.10): by default the
  copies are consumed west to east, then south to north, so a schedule that
  covers the north half of the fab landed on southern copies of the same
  label. `--tool-layer NAME` (repeatable, `*` wildcards) takes copies on
  that layer first; `--prefer north|south|east|west` takes that side first;
  `--prefer X0,Y0,X1,Y1` takes copies inside the rectangle (drawing units)
  first. They combine: layer first, then direction. A row that still lands
  elsewhere (no free copy left where the options point) is named in a
  warning. `--trace LABEL` prints the order used and marks every copy
  `in`/`OUT` with its layer; `--inspect` prints where the inserts are.
- **Move-in passageway** (1.12): `--aisle` takes the copies of a label
  farthest from the aisle first, so the area north of it fills north to
  south and the area south of it south to north, as tools are pushed in
  from the aisle. Order of precedence: `--tool-layer`, then `--prefer` (a
  rectangle or `north-of=`/`south-of=` half = inside first; `north`/`south`
  combined with a horizontal aisle = that side of the aisle first), then
  distance from the aisle, then west to east. Without `--aisle`, `--prefer
  north` alone sorts all copies north to south, which is wrong for the
  southern area.
- **Drawing anchors** (1.13): `--aisle` and `--prefer north-of=` (also
  `south-of=`, `east-of=`, `west-of=`) accept `hatch:PATTERN@LAYER` (a
  hatched pathway), `layer:NAME` (everything on that layer),
  `text:LABEL@LAYER` (a text's position) or plain coordinates
  (`Y`, `x=X`, `X0,Y0,X1,Y1`). The run prints every resolved anchor, the
  viewer draws them, `--trace` measures against them.
- The matching option `blocktext` is chosen automatically; `--match blocktext`
  forces it. `--loose` ignores punctuation and spacing (`DT-112` matches
  `DT_112`); `--inspect` lists the near misses `--loose` would fix.
- When per-tool location IDs are added to the drawings later (for example an
  attribute on each tool block), drop `--type-col` and let the script match
  IDs directly (`--match attrib:TAG` if needed). Provisional slots then
  become exact locations.

## Move-in conflicts (1.11)

Every run checks the placed tools for potential conflicts and prints them;
the viewer marks them (dashed warning outline on the tools, a diamond marker
that appears on the day the conflict becomes real, a red pill in the header,
a Conflicts tab, lines in the tooltip). `--conflicts-csv FILE` exports the
list, `--no-conflicts` switches the check off.

| Kind | Meaning | Knob |
|---|---|---|
| overlap (high) | Two scheduled tools share floor space (their outlines overlap by more than 2 % of the smaller one), or two schedule IDs point at the same block | none: this is a planning error or a shortfall of placements |
| access (medium) | A tool arrives after the tools on all four of its sides are in place, with no gap wider than 60 % of its smaller dimension to bring it in | walls are not known, so a tool between a wall and earlier tools is not caught |
| capacity (medium) | More move-ins on one day than `--max-per-day N` allows | `--max-per-day` (off by default) |
| storage (medium) | A delivered tool found no free storage slot for its whole wait | add zones, capacity or `--storage-cell` |
| crowding (low) | A cluster of tools closer than `--conflict-gap` (default one typical tool size, edge to edge) moving in within `--conflict-window` days of each other (default 3) | `--conflict-gap 0` disables, `--conflict-window 0` = same day only |

Conflicts are found on the final positions, so in type mode they depend on
which copy of a label a row took (`--tool-layer` / `--prefer`).

## Storage (laydown) space (1.11)

Tools that are delivered before their move-in day wait somewhere. Give the
zones and the script assigns each waiting tool a slot:

```
python fab_movein.py --dxf data\floor3.dxf --schedule data\move_in_plan.xlsx --type-col Block_ID ^
    --storage-layer "LAYDOWN" --storage "Dock B=41200,18300,41900,18800:10" --storage-csv storage_plan.csv
```

- **Zones**: `--storage NAME=X0,Y0,X1,Y1[:CAP]` (drawing units, AutoCAD `ID`
  shows coordinates; `:CAP` = maximum tools at once) and/or `--storage-layer
  NAME` (closed polylines or blocks on that layer, named by a text inside
  them). `--inspect` lists layers whose names look like storage areas.
- **Arrival**: a delivery / arrival / ETA column is auto-detected
  (`--arrival-col` to choose). Without one, every tool arrives
  `--storage-lead` days before its move-in (default 7 when zones are given).
- **Slots**: each zone is cut into a grid of slots sized by the tools
  (`--storage-cell W,H` overrides). A tool takes one slot (or several for a
  large footprint) from arrival until move-in; slots are reused afterwards.
  Tools are placed in arrival order, longest wait first, into the nearest
  zone with room for the whole wait. A `--storage-col` column (e.g. a
  "Storage" column naming the zone) is honoured when that zone has room.
- **Output**: the run prints slots, tools and peak occupancy per zone; the
  viewer draws the zones, shows each waiting tool in its slot from arrival
  until it moves to its final position, counts "waiting in storage" in the
  stat tile, and lists occupancy in the Storage tab (hover a zone for today's
  occupants). `--storage-csv FILE` writes tool, dates, zone and slot.
- Tools that find no slot are reported as `storage` conflicts.

## Filters and colours (1.11)

Besides the type, group and area columns, the viewer picks up model, vendor
/ supplier and module columns automatically (any other with `--facet-col
NAME`). Each becomes a dropdown next to the search box; the dropdowns
combine, hide the tools that do not match, and the counts, the chart and
the lists then describe the selection ("22 of 286 tools match"). The search
box matches IDs, names and all these values. "Color by" offers every field;
the start colour is the Module column if there is one (`--color-by NAME`
chooses another).

## Diagnostics

- `--inspect` (or `--inspect-full` for untruncated lists): what the DXF
  contains (block attributes, block names, text labels, text inside block
  definitions with insert counts), what the schedule contains, what would
  match under each option, and near misses.
- `--layers`: every background layer ranked by its share of the output
  (geometry, placements, symbol bytes, texts), block placements per layer as
  matched tools / free slots / other blocks, names that look like services or
  annotation flagged, and a paste-ready `--exclude-layer` list. Writes no
  HTML. Excluding a layer never removes matched tools, only that layer's
  backdrop.
- `--trace LABEL`: for one label, the schedule rows carrying it, every place
  in the drawing that holds it (attribute, block name, block text, text), the
  placements it got or why not, free placements, and similar labels
  (substring or punctuation-only differences) for split or misspelled labels.

## Performance playbook

Work down the list and stop when the file is comfortable.

1. **Default run.** Since 1.8, `--block-detail auto` keeps only a block's
   outline and major lines, rarely placed blocks (bound xrefs) are merged into
   their layer, and duplicate block definitions share one symbol.
2. **`--layers`, then `--exclude-layer`** for services, annotation and
   anything the story does not need. Wildcards work and the option repeats:
   `--exclude-layer "*PIPE*" --exclude-layer "*HVAC*"`.
3. **`--max-mb 40`** as a hard cap: drops the heaviest background layers until
   the file fits. Tools are never dropped.
4. **Finer knobs:** `--simplify` (coarser curves), `--max-texts` (fewer
   background texts), `--block-detail outline` or `box`.
5. **In the viewer:** View menu, Fine drawing detail "never", Fast rendering,
   untick Drawing text.

What the viewer does on its own: the floor plan is rasterised once in its own
compositor layer, playback only touches tools whose state changes, the plan is
culled in tiles with two levels of detail, and zoom scales the bitmap first
and redraws sharp after 160 ms. Measured on synthetic heavy data under
software rendering: playback 3.6 to 59 fps, zoom 8.7 to 62 fps, and a tool
layer built from full equipment drawings shrank 33 times with block detail
`auto`.

## Using the viewer

- Space plays and pauses, the arrow keys step one day, the slider scrubs, the
  speed selector sets days per second.
- Right panel: This day / Up next / All tools.
- Search box: type a label; Enter zooms to it. Dashed ghost outlines show
  future tools that match.
- Filter dropdowns next to the search box (one per schedule field); they
  combine, and "clear" resets them.
- View menu: pending footprints (dashed), drawing text on or off, tool ID
  label mode, fine drawing detail (auto / always / never), fast rendering,
  conflict markers, storage zones, and each layer with its size in KB.
- Conflicts tab (or the red pill in the header): click a row to jump to the
  day and zoom to the tools. Storage tab: occupancy per zone today, tools
  arriving next, tools without a slot.
- Zoom in, zoom out, fit; PNG saves the current view; the half-circle button
  toggles light and dark.
- Hover the title for a load-time breakdown; hover a tool for its tooltip.

## Known issues and answers

| Symptom | Cause | Answer |
|---|---|---|
| "WinError 10053, connection aborted" while viewing | A local preview server choking on a huge file. The script opens no sockets. | Double-click the HTML; keep the output small. |
| DWG cannot be read | ezdxf reads DXF only | SAVEAS AutoCAD 2013 DXF |
| Tools drawn as plain boxes, warning mentions PROXYGRAPHICS | AutoCAD MEP proxy objects carry no geometry in the DXF | The boxes sit at the true insertion point. For real outlines set PROXYGRAPHICS to 1 in AutoCAD and export again. |
| Only a handful of rows match | Type column taken for an ID column, or punctuation differences | `--inspect`, then `--type-col`, `--loose` or `--match` as the report suggests |
| A label exists in CAD but never appears | Label split into several texts or spelled differently inside the block | `--trace LABEL`, then fix the text in the block definition |
| A label is visible in the north half but the animation shows the copy in the south half | Type mode consumed the copies west to east / south to north | `--tool-layer "0_ENG 3K tool"` and/or `--prefer north`; `--trace LABEL` shows every copy and which one was taken |
| Largest layer tiny but file huge | Symbol bytes were not attributed to layers before 1.7 | Fixed in 1.7 |
| `floor1.dxf not found` | Files live in `data\` | `--dxf data\floor1.dxf` |

## Option reference

| Option | Purpose |
|---|---|
| `--dxf`, `--schedule`, `--sheet` | Inputs; the biggest worksheet is used by default |
| `--id-col`, `--date-col`, `--name-col`, `--group-col`, `--area-col` | Schedule columns; all auto-detected when omitted |
| `--type-col NAME` | Type mode: rows take free placements of their type in date order |
| `--facet-col NAME`, `--color-by NAME` | Extra filter / colour columns in the viewer; the start colour field |
| `--conflict-window DAYS`, `--conflict-gap UNITS`, `--max-per-day N` | Conflict thresholds: crowding window (3) and distance (one tool size); daily move-in cap (off) |
| `--no-conflicts`, `--conflicts-csv FILE` | Skip conflict detection; export the conflict list |
| `--storage NAME=X0,Y0,X1,Y1[:CAP]`, `--storage-layer NAME` | Storage zones by rectangle or from a drawing layer |
| `--arrival-col NAME`, `--storage-lead DAYS` | Delivery date column (auto-detected); or a fixed lead time before move-in (7) |
| `--storage-cell W,H`, `--storage-col NAME`, `--storage-csv FILE` | Slot size; a column naming the wanted zone; export the storage plan |
| `--date-format auto/dmy/mdy/ymd` | How ambiguous numeric dates are read |
| `--match auto/attrib:TAG/blockname/blocktext/text` | How tools are found in the DXF |
| `--loose` | Ignore punctuation and spacing when matching |
| `--tool-layer NAME` | Type mode: copies of a label on this layer are taken first (repeatable, wildcards) |
| `--prefer WHERE` | Type mode: `north`/`south`/`east`/`west` side first, `X0,Y0,X1,Y1` rectangle first, or `north-of=REF` (also south/east/west) with REF a `layer:`, `text:` or coordinate |
| `--aisle WHERE` | Type mode: the move-in passageway (`hatch:PATTERN@LAYER`, `layer:NAME`, `text:LABEL@LAYER`, `Y`, `x=X` or `X0,Y0,X1,Y1`); copies farthest from it are taken first, each side filling towards the aisle |
| `-o`, `--title`, `--start-date`, `--end-date`, `--open` | Output file, page title, timeline clipping, open in browser |
| `--exclude-layer`, `--only-layer` | Drop or keep drawing layers; wildcards, repeatable |
| `--block-detail auto/full/outline/box` | How much of each block's inner geometry to keep (default auto) |
| `--max-mb MB` | Hard cap on output size; drops heaviest background layers, never tools |
| `--simplify UNITS`, `--max-texts N`, `--precision N` | Curve simplification, text cap, coordinate decimals (all auto by default) |
| `--inspect`, `--inspect-full` | Report what is in the files and what would match |
| `--layers` | Layer weight table, placements per layer, suggested excludes |
| `--trace LABEL` | Explain one label end to end (repeatable) |
| `--version` | Print the script version |

## Version history

| Version | Date | Change |
|---|---|---|
| 1.0.0 | 25 Aug 2026 | First version |
| 1.1.0 | 1 Sep 2026 | Rewrite for large real DXFs: correct tool placement, one shared symbol per block definition, robust tolerances, proxy blocks at their insertion point, playback that touches only changed tools |
| 1.2.0 | 1 Sep 2026 | Type mode (`--type-col`), matching on text inside block definitions, richer `--inspect` |
| 1.3.0 | 1 Sep 2026 | `--loose`, near-miss report, `--inspect-full`, numbered duplicate labels |
| 1.4.0, 1.4.1 | 2 Sep 2026 | `--max-mb` size cap, progressive background rendering, load-time readout |
| 1.5.0 | 2 Sep 2026 | Sub-1 fps playback fixed: background cached in its own compositor layer |
| 1.6.0, 1.6.1 | 2 Sep 2026 | Fast zoom (bitmap-scaled background, spatial tiles, level of detail), `--trace`, `--layers`, per-layer sizes in the View menu |
| 1.7.0 | 2 Sep 2026 | Rarely placed blocks inlined, duplicate symbols deduped, honest per-layer sizes |
| 1.8.0, 1.8.1 | 2 Sep 2026 | `--block-detail` footprint reduction; `--layers` shows placements per layer |
| 1.9.0 | 2 Sep 2026 | Consolidation (split functions, memoised indexes, Windows console guard) and the in-repo regression suite |
| 1.10.0 | 3 Sep 2026 | `--tool-layer` and `--prefer`: choose which copy of a repeated label a row takes (layer, side of the fab, rectangle); `--trace` shows the order and layer, `--inspect` shows the insert extents |
| 1.13.1 | 4 Sep 2026 | `hatch:PATTERN@LAYER` anchor for the ANSI31 pathway on `00_Dummy` (the layer holds other objects); `--inspect` lists hatch patterns |
| 1.13.0 | 4 Sep 2026 | Drawing anchors: `--aisle layer:00_Dummy`, `--prefer north-of=text:XPH-D@BAY-ID`; anchors printed per run, drawn in the viewer, missing ones stop the run |
| 1.12.0 | 4 Sep 2026 | `--aisle`: copies of a label fill from the far end of each area towards the move-in passageway (north area north to south, south area south to north) |
| 1.11.0 | 4 Sep 2026 | Move-in conflicts detected and marked (overlap, boxed in, crowding, daily capacity, no storage); storage zones and slot assignment for delivered tools; filters by model / vendor / module / any column, colour by module |

## For whoever changes the code

- The tool stays one file; users copy only `fab_movein.py`. Tests live in
  `tests/`.
- `CLAUDE.md` documents the architecture and six performance invariants that
  must not regress: never explode geometry per insert, tolerances from a
  robust drawing size, proxy blocks at their insertion point, auto-fit ignores
  far-away junk, playback proportional to changed tools with the background in
  its own SVG, and simplification plus caps.
- Before pushing: `python -m pyflakes fab_movein.py` and
  `python tests/run_tests.py` (`--no-browser` without Playwright,
  `--only=small,heavy` to pick scenarios). Add a scenario whenever a
  real-data problem is fixed.
- On new real data reach for `--inspect`, then `--trace`, then `--layers`.

## Open items

1. Run floor3 with 1.9.0 and note file size, load time (hover the title),
   playback and zoom. If still heavy, apply the playbook and record which
   layers were excluded.
2. Resolve DT_545R with `--trace DT_545R`; if the label is split, fix it in
   the block definition.
3. Agree the layer exclusion list for floor3: services, annotation, and the
   tool-category layers (`0_40K tool`, `0_ENG 3K tool`, ...) whose free slots
   and other blocks are only backdrop.
4. When unique per-tool location IDs are assigned in the drawings, switch from
   type mode to ID mode. Until then, run with `--tool-layer` / `--prefer` /
   `--aisle` for the part of the fab the schedule covers (see Everyday
   commands). The passageway and the upper/lower split come from the
   drawing anchors above; check the `Anchor :` lines of every run after a
   new DXF export.
5. Optional: run `python tests\run_tests.py --no-browser` after each script
   download (about 20 seconds).
6. Storage zones for floor1 / floor3: draw the laydown areas as closed
   polylines on one layer (with a text label inside each) and run with
   `--storage-layer NAME`, or pass `--storage NAME=X0,Y0,X1,Y1` from AutoCAD
   `ID` coordinates. Add a delivery-date column to `move_in_plan.xlsx` so the
   waits are real rather than the 7-day default.
7. Review the first conflict report on floor3: `overlap` rows are placement
   or schedule errors, `crowding` needs the gap / window tuned to the fab's
   rigging practice (`--conflict-gap`, `--conflict-window`).
