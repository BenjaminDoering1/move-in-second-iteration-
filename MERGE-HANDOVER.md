# Merge handover: branch `claude/move-in-conflicts-storage-v9dvth`

Written 4 September 2026 for the session that merges this branch with
work done elsewhere. Read this first, then `HANDOVER.md` (user-facing) and
`CLAUDE.md` (architecture). The tool is still one file, `fab_movein.py`,
plus `tests/run_tests.py`.

## Where the branch is

| | |
|---|---|
| Repositories | `BenjaminDoering1/Fab-move-in-simulator-` (real history, branched from `main` at 1.10.0 = commit b28d4aa) and `BenjaminDoering1/move-in-second-iteration-` (mirror: same files, its own commits, no shared history with the first repo) |
| Branch | `claude/move-in-conflicts-storage-v9dvth` in both |
| Version | `__version__ = "1.13.1"` (main is 1.10.0) |
| Commits on top of main | 2b2c131 v1.11.0 conflicts + storage + filters · acd8af3 v1.12.0 `--aisle` · 57ec0f9 v1.13.0 drawing anchors · 68337f4 v1.13.1 hatch anchors · 704e2c0 floor 1 anchor in HANDOVER |
| Diff vs main | fab_movein.py +1446 / −73 lines, tests +319, CLAUDE.md +60, HANDOVER.md +165 |
| Tests | `python tests/run_tests.py` → 136 checks pass (browser checks included); `--no-browser` for the fast subset |

Merge the first repository's branch. The mirror can simply be overwritten
with the merged files afterwards (that is how it has been kept in step).

## What the branch adds, in one paragraph each

1. **Move-in conflicts** (on by default). `find_conflicts` checks the placed
   tools for `overlap` (convex outlines share > 2 % of the smaller, or two IDs
   on one block), `access` (a tool whose four sides each have an earlier
   tool within 0.6 × its smaller dimension), `crowding` (union-find clusters
   closer than `--conflict-gap`, default one typical tool size, moving in
   within `--conflict-window` days, default 3), `capacity` (`--max-per-day`)
   and `storage` (no slot). Printed after the run, `--conflicts-csv`,
   `--no-conflicts`. Viewer: dashed outline on involved tools, diamond marker
   per conflict appearing on its day, header pill, Conflicts tab, tooltip
   lines.
2. **Storage slots.** Zones from `--storage NAME=X0,Y0,X1,Y1[:CAP]` or
   `--storage-layer NAME` (closed polylines / blocks, named by a text
   inside). A tool whose arrival (`--arrival-col`, auto-detected, else
   `--storage-lead` days, default 7) precedes its move-in gets a slot in the
   nearest zone with room for the whole wait (`assign_storage`: zone cut
   into slots, `--storage-cell`, `--storage-col` wish, `:CAP` concurrency).
   `--storage-csv`. Viewer: zones, waiting tools drawn in their slot until
   move-in, Storage tab, zone tooltips, "waiting in storage" count.
3. **Filters and colours.** Model / vendor / module columns (and any
   `--facet-col`) become dropdowns that combine (AND), hide non-matching
   tools and recount stats, chart, legend and lists. Colour by any field,
   default the Module column (`--color-by`).
4. **`--aisle`.** Type mode consumes the copies of a label farthest from the
   move-in passageway first, so the area north of it fills north to south
   and the area south of it south to north. Precedence: `--tool-layer`, then
   `--prefer`, then distance from the aisle, then west to east.
5. **Drawing anchors.** `--aisle` and `--prefer north-of=|south-of=|east-of=|west-of=`
   accept `hatch:PATTERN@LAYER`, `layer:NAME`, `text:LABEL@LAYER` or typed
   coordinates. Resolved on every run and printed (`Anchor :` lines), drawn
   in the viewer as labelled dashed lines, missing ones stop the run with
   candidates. Project-specific values (floor 3: `hatch:ANSI31@00_Dummy`
   and `text:XPH-D@BAY-ID`; floor 1: `text:XIM-C`) live only in
   HANDOVER.md, "Drawing anchors".

## Where the changes sit in `fab_movein.py` (merge hot spots)

The file is one module; these are the regions touched, top to bottom.
Anything not listed is unchanged from 1.10.0.

| Region | Change | Conflict risk with other work |
|---|---|---|
| Module docstring | two paragraphs on conflicts and storage | low |
| `__version__` | 1.10.0 → 1.13.1 | **certain**: pick the higher number, bump again after merging |
| `_VIEWER_TEMPLATE` CSS | new rules appended after `body.hide-ids`, `g.tool.dim`, `input.txt`; `#filters`, `#refG`, `#zoneG`, `#storeG`, `#cfG`, `.cf-row`, `.t-warn`, `.t-store`, `body.wide` | medium if the other work restyles the same selectors |
| `_VIEWER_TEMPLATE` HTML | `#cfPill` before `#warnPill`; `#filters` after `#search`; View menu: `#lblCf/#ckCf`, `#lblZones/#ckZones`, `#lblRefs/#ckRefs` after `#ckFast`; `#statExtra` after `#statDelta`; tabs `#tabCf`, `#tabSt` after "All tools" | medium |
| Viewer JS, top | `dayIdx`, `CF`, `ST`, `STORED`, `CFCUM`, `INSTORE`, `EV` events + `indexEvents`, `MED`; `newIds/NEWP/TOTAL` now rebuilt by `recount()` over `passes(t)`; `TOTAL` is `let`, `TOTAL_ALL` added; `FIELDS = D.fields` | **high**: the old `var TOTAL` / NEWP block was replaced |
| `makeMapSkeleton` | groups `refG`, `zoneG`, `ghostG`, `storeG`, `toolsG`, `cfG` in that order; delegated events for store/cf/zone | medium |
| `addTool` | adds `cf sev-*` classes | low |
| new `buildExtras()` | ref lines, zones, stored tools, markers, `indexEvents()`; called in the init chain before `applyColors()` | low |
| `applyRange` | walks `EV` after the tool loop | medium |
| `setDay` | `#statExtra` text | low |
| `zoomToTool` / new `zoomToBox`, `flashTool`, `focusConflict` | | low |
| `buildColorControls` | options from `FIELDS`, default `D.colorBy`; new `buildFilters`, `applyFilter` | **high** if the other work touched colour-by (`D.fields` is now a list, not `{g, a}`) |
| `renderLegend`, `isSearchMatch`, `makeRow`, `renderList` | filter-aware; new `cf` and `st` tabs; new `makeConflictRow`, `makeStoredRow`; `selectTab()` replaces the inline tab handler | medium |
| `showToolTip` + new `showConflictTip`, `showZoneTip` | | low |
| `buildMenu` | pill, toggles, `body.wide` | low |
| init chain | `buildExtras()` then `applyColors()`; `buildFilters()` after `buildColorControls()` | low |
| `Drawing.__init__/_scan` | `self.hatches` recorded for HATCH entities | low |
| after `hit_centre` | `parse_prefer(spec, dwg)` (new `north-of=` forms, `ref`, `miss`), new `resolve_ref`, `_extent_ref`, `ref_line`, `HALF_PLANES`, `parse_aisle(spec, dwg)`, `placement_policy(args, dwg)` returning `anchors` | **high** if the other work changed `--prefer` / `placement_policy` |
| `match_rows` | `placement_policy(args, dwg)` call | low |
| `build_tools` | tool records get private `_hull`, `_d`, `_ins`, `_arrive`, `_store` (stripped in `build_payload`) and `f0..fn` facets | medium |
| new section before `build_payload` | conflicts (`hull_svg` … `write_conflicts_csv`), storage (`STORAGE_LAYER_HINT` … `write_storage_csv`), `field_list` | low (new code) |
| `day_axis` | starts at the first arrival (`ad`) | low |
| `build_payload` | storage → conflicts → `cf` indexes → strip `_` keys; `ref_lines`; new payload keys | **high**: the payload dict literal changed |
| patterns | `NAME_PATTERNS` lost `model` (now a facet); new `FACET_PATTERNS`, `ARRIVAL_PATTERNS`, `STORE_PATTERNS` | low |
| `build_parser` | new groups "move-in conflicts", "storage (laydown) space"; `--facet-col`, `--color-by` in the columns group; `--aisle` after `--prefer`; `--prefer` help rewritten | medium |
| `resolve_columns` | arrival, storage and facet columns; returns `arrive`, `store`, `facets` | medium |
| `build_rows` | `arrive`, `store`, `facets`, `facet_labels`, `arrive_label` per row; `cols["lead"]` | medium |
| `main` | storage zone parsing before `resolve_columns`; lead-time handling; `args.anchors`; `Anchor :` printing; conflict report and CSV writers after the match summary | **high**: several insertions in the middle of `main` |
| `do_inspect` | storage-layer hint after the layer list; hatch pattern list before texts; `--aisle` hint under the insert extents | low |

## Payload contract (what the viewer expects)

New or changed keys in the JSON handed to `_VIEWER_TEMPLATE`:

- `fields`: ordered list `[{k, l}]` (was `{g, a}`); keys `ty`, `g`, `a`, `f0`, `f1`, …
- `colorBy`: key of the start colour field ("" for single colour)
- `conflicts`: `[{k, s, ids, day, x, y, d}]` (`x`/`y` SVG coordinates or null); tools carry `cf: [indexes]`
- `storage`: `null` or `{zones: [{n, x, y, w, h, cols, rows, cap, src}], cw, ch}`; stored tools carry `ad` (arrival ISO), `sz`, `sl`, `sx`, `sy`
- `refLines`: `[{n, seg: [x0, y0, x1, y1]}]` in SVG coordinates
- `arrivalLabel`: header of the arrival column
- per tool: `f0…` facet values

SVG coordinates are drawing X and negated drawing Y throughout, as before.

## Tests to keep

`tests/run_tests.py` gained the `conflicts` scenario (`make_conflicts` /
`scenario_conflicts`: every conflict kind, zones from a layer and from the
option, CSVs, viewer events) and the `aisle` scenario (`make_aisle` /
`scenario_aisle`: two areas either side of a hatched pathway, the bay-ID
text with a stray duplicate, a lower-fab copy, anchors, error cases).
`make_small` gained Model / Vendor / Module columns and `scenario_small`
checks the filter dropdowns, recounting and clearing in the browser. If the
other work also edited `make_small`, keep both sets of columns; the checks
assume Vendor = `("ACME","Globex","Initech")[i % 3]` and Module =
`f"Module {1 + (i // 8) % 3}"` (22 Globex tools, filters that actually
narrow when combined).

## Things that are easy to get wrong when merging

- `fields` changed shape. Any code reading `D.fields.g` must move to the
  list form or the colour selector breaks silently.
- `placement_policy`, `parse_prefer` and `parse_aisle` now take `dwg`.
  Callers without it die on drawing references.
- Tool records carry private `_`-prefixed keys until `build_payload` strips
  them; anything that serialises tools earlier will leak them.
- `TOTAL` is recomputed by `recount()`; do not reintroduce a `const TOTAL`.
- The `EV` event list must be indexed (`indexEvents()`) after all events are
  pushed and before the first `applyRange`; `buildExtras()` does this.
- `NAME_PATTERNS` no longer matches "model": a schedule whose only
  descriptive column is "Model" shows it as a facet and the list rows fall
  back to `t.nm || t.f0 || t.a`.
- Version: bump to a number above both branches after merging and add a
  row to the HANDOVER version table.

## Open items this branch leaves

1. Floor 1 has no passageway rule yet (`--aisle` for floor 1 is unset; see
   the anchors table in HANDOVER.md).
2. The layer of the `XIM-C` text on floor 1 is not recorded; the first run
   prints it in the `Anchor :` line.
3. Real-data run of floor 3 with `--aisle hatch:ANSI31@00_Dummy --prefer
   "north-of=text:XPH-D@BAY-ID"` has not been done in this session; the
   printed anchors and `--trace` are the check.
4. Storage zones for floor 1 / floor 3 are not defined; the schedule has no
   delivery-date column yet (7-day default lead applies once zones exist).
5. Conflict thresholds (`--conflict-gap`, `--conflict-window`) are defaults
   chosen without real data; expect to tune after the first floor 3 report.
6. A coordinate readout in the viewer (drawing X/Y under the cursor) was
   offered and not built; it would help with `--storage` rectangles.
