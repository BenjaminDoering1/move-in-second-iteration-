# Fab Move-In Simulator

Single-file Python tool (`fab_movein.py`, no other source files) that reads a
fab layout **DXF** plus a tool move-in **schedule** (xlsx/csv) and writes one
self-contained **HTML** file animating tools appearing on the floor plan day
by day. Everything runs locally on the user's machine; the HTML works offline
by double-clicking it — there is no server component at all.

```
pip install ezdxf openpyxl
python fab_movein.py --dxf layout.dxf --schedule movein.xlsx -o simulation.html
python fab_movein.py --dxf layout.dxf --schedule movein.xlsx --inspect   # see what's matchable
```

## Architecture

* `_VIEWER_TEMPLATE` — the entire HTML/CSS/JS viewer as a raw string; the JSON
  payload is substituted for `/*__DATA__*/` in `render_html()`.
* `Drawing` — DXF loading. Block definitions are exploded **once per
  definition** (local coordinates, `blockdef()`); each INSERT stores only a
  2D affine matrix (`insert_matrix()`, row-vector Matrix44 convention:
  a=m00 b=m01 c=m10 d=m11 e=m30 f=m31). SVG use-transform is `flip∘M∘flip` =
  `matrix(a, -b, -c, d, e, -f)` because SVG y is negated DXF y.
* `Schedule` — xlsx/csv reading, header detection, column auto-resolution by
  header patterns then by content (IDs that actually exist in the drawing,
  values that actually parse as dates).
* `choose_strategy` picks the matching option; `build_payload` (v1.9) is a
  ~40-line orchestrator over focused steps: `match_rows` (ID or type mode,
  `--trace`, per-layer placement counts) → `build_tools` (tool records +
  deduped footprint symbols) → `collect_background` (inlined blocks, shared
  symbols, texts per layer) → `cap_texts` → `compute_bounds` →
  `tile_layers` (tiles, fine/coarse, symbol-byte shares) → `day_axis` →
  `shrink_to_budget`. `main` is likewise split: `build_parser` (argument
  groups), `resolve_columns`, `build_rows`, `suggest_type_col`.
* `Drawing` indexes (`index_by_attrib/blockname/text/blocktext`) are memoized
  per run and per `--loose` mode; callers must treat them as read-only.
* **Blocktext matching** (`index_by_blocktext`, v1.2) — text drawn INSIDE a
  block definition is stored once in the DXF but labels every insert of that
  block (why AutoCAD FIND / text search sees it once). The `blocktext`
  strategy maps each such string to all inserts of blocks containing it,
  pooling copy-paste duplicate definitions (`A$C…` anonymous names) that
  share a label. `--inspect` lists these strings with insert counts.
* **Loose matching** (`--loose`, v1.3) — `mkey()` switches all matching keys
  from `norm()` to `compact()` (letters+digits only), so `DT-112` matches
  `DT_112`. `--inspect` prints a near-miss report (schedule values that would
  match under `--loose`) and `--inspect-full` lists everything untruncated.
  In type mode, `--type-col` may equal the ID column (schedules with no
  unique per-tool IDs): repeated labels are numbered `DT_112 (2)`, `(3)`…;
  in ID mode duplicate-ID rows are counted and warned about.
* **Type mode** (`--type-col`, v1.2) — for drawings that label tool TYPES
  (many tools share one label like `DT_112`) instead of unique per-tool IDs:
  each schedule row consumes one still-free placement of its type, in date
  order (deterministic: placements sorted by position). Per-type shortfalls
  are warned; payload gets `byType` + per-tool `ty`; tooltip marks slots
  provisional. Auto-detection guards against mistaking a type column for the
  ID column (heavy value repetition → keep the header ID column and suggest
  `--type-col`); in type mode the ID column need not appear in the drawing.
* **Placement policy** (`--tool-layer`, `--prefer`, v1.10) — in type mode a
  label placed several times hands its copies to rows in a fixed order
  (`hit_centre`: west→east, then south→north). Real-data failure: the
  schedule covered the north half of the fab, the same labels also existed
  in the south, and rows landed on the southern copies. `placement_policy`
  builds the sort key: copies on `--tool-layer NAME` (repeatable, wildcards
  via `layer_wanted`) first, then `--prefer north|south|east|west` (that
  side first) or `--prefer X0,Y0,X1,Y1` (inside the rectangle first), then
  the default order. Rows that still land elsewhere (no free copy left where
  the options point) are warned about by ID and reason. `--trace` prints the
  order and marks each placement `in`/`OUT` with its layer; `--inspect`
  prints where the block inserts are (`insert_extents`). Both options warn
  and do nothing in ID mode, where every copy carrying an ID is drawn anyway.
  **`--aisle`** (v1.12, `parse_aisle`): the move-in passageway as `Y`,
  `x=X` or a line `X0,Y0,X1,Y1`; the sort key becomes (layer, prefer group,
  -distance from the aisle, x, y) so each side fills from its far end
  towards the aisle (north area north→south, south area south→north). A
  directional `--prefer` combined with a matching aisle orientation means
  "that side of the aisle first"; a rectangle still means "inside first".
  **Drawing anchors** (v1.13, `resolve_ref`): `--aisle` and `--prefer
  north-of=|south-of=|east-of=|west-of=` accept `hatch:PATTERN[@LAYER]`
  (extent of the matching HATCH entities, recorded in `Drawing.hatches`
  during `_scan`; several → warn and union), `layer:NAME` (extent of the
  layer's geometry + inserts, warned when it holds more than one entity →
  horizontal centreline if wider than tall, else vertical), `text:LABEL[@LAYER]` (largest matching top-level text,
  then blocktext inserts; several → warn) or coordinates. Missing → `die`
  with candidates. `placement_policy(args, dwg)` returns `anchors`; `main`
  prints them (`Anchor :`), `build_payload` emits `refLines` (SVG segments
  across the bounds) drawn by the viewer's `#refG`. Project-specific
  workarounds (floor3: ANSI31 hatch on `00_Dummy`, bay text `XPH-D` on `BAY-ID`)
  live in HANDOVER.md "Drawing anchors", not in code.
* **Size budget** (`--max-mb`, v1.4) — `shrink_to_budget` drops the
  heaviest background layers (never tools; `strong` walls last, layers under
  10 KB skipped) until the payload fits, then prunes/remaps unreferenced
  `bsyms`. The large-output warning ranks layers by `_layer_weight`
  (geometry + placements + texts). The viewer builds the background in
  time-budgeted rAF tasks (`backgroundTasks`: ~1 MB path pieces split at 'M'
  boundaries, `<use>` batches of 2500), so heavy pages paint progressively
  instead of freezing.
* **Tiles + level of detail + bitmap zoom** (v1.6) — each layer is stored as
  `tiles: [[tileIdx, bbox, coarsePath, finePath, coarseUses, fineUses], …]`
  (8×8 grid over the fitted bounds by piece centre; pieces larger than a
  cell or outside go to global tile -1; tile bbox = union of members so
  culling is exact). "Fine" = pieces smaller than `fineUnit` (drawing size
  / 500). Viewer: `updateBgZoom` hides off-screen tiles (35% margin), toggles
  `body.hide-fine` (auto: shown when `ppu*fineUnit >= 2.5`), sets background
  `stroke-width = 1/ppu` on `MAP.bgRoot` (no vector-effect on background).
  Pan/zoom: `setVB` updates the foreground viewBox live but only applies a
  CSS `translate/scale` to `#bgSvg` computed from `MAP.rendered`; `commitBg`
  (160 ms after the last change, on resize, before PNG export) sets the real
  viewBox and re-culls. View menu: fine detail auto/always/never, fast
  rendering (`shape-rendering:optimizeSpeed`).
* **Rarely-placed blocks are inlined** (v1.7) — a block inserted <= 2 times
  (bound xrefs holding a whole floor, one-off details) is transformed to WCS
  and merged into its layer's geometry instead of becoming a shared symbol:
  it then gets tiles, detail culling, dedupe and honest accounting. Shared
  symbols (`bsyms`) are deduped by geometry signature (copy-paste `A$C…`
  twins), and each layer carries `sw` = its share of referenced symbol bytes
  so `_layer_weight`, `--layers` and `--max-mb` see the true cost. Before
  this, a 45 MB file could report "largest layer 279 KB" because the bulk
  sat in one xref block's symbol.
* **Block detail policy** (`--block-detail`, v1.8) — `Drawing.block_lite()`
  applies a per-block-definition detail policy to background placements,
  inlined blocks and tool footprint symbols alike: `auto` (default) keeps
  pieces >= 12% of the block's size, then falls back to the convex-hull
  footprint (`convex_hull`, monotone chain) when nothing/too little is kept,
  coverage < 60% of the block, or > 1500 points remain; `outline` = hull;
  `box` = bbox; `full` = everything. Blocks >= 15% of the drawing size
  (bound xrefs, bays) are always full. Matching still uses the full-geometry
  bbox. Real-data motivation: a tool layer of 48 MB where every tool block
  was a complete equipment drawing.
* **`--layers`** (v1.6.1) — `report_layers` lists every background layer by
  output weight (share + cumulative), flags names matching `_IRRELEVANT`
  (services/annotation) and prints a paste-ready `--exclude-layer` list, then
  exits without writing HTML. The View menu shows each layer's KB. Since
  v1.8.1 it also prints block placements per layer from
  `payload["layerTools"]` = {layer: [matched tools, free slots (label in the
  schedule but no row consumed it), other blocks]}, so tool-category layers
  (`0_40K tool`, …) can be judged: excluding a layer never removes matched
  tools, only its free slots and other blocks from the backdrop.
* **`--trace LABEL`** (v1.6) — `trace_label` prints schedule rows carrying
  the label, every drawing source holding it (attribute / block name / block
  text / loose text) with placements, the strategy used, assignments from
  `assign_log`, free placements, and *similar* labels (substring or
  punctuation-only differences) for split/misspelled labels.
* **Move-in conflicts** (v1.11) — `find_conflicts` works on the tool records
  in SVG coordinates (`_hull` convex outline, `_d` date, `_ins` placement
  identity; private `_` keys are stripped before output). Kinds: `overlap`
  (convex outlines share > 2% of the smaller area via Sutherland-Hodgman
  clipping, or two IDs on one block), `crowding` (union-find clusters of
  tools closer than `--conflict-gap`, default one typical tool size, moving
  in within `--conflict-window` days, default 3), `access` (a tool whose four
  sides each have an earlier tool within 0.6 × its smaller dimension),
  `capacity` (`--max-per-day`), `storage` (no slot, from `assign_storage`).
  Neighbour queries use `ToolGrid`. Payload: `conflicts` [{k, s, ids, day,
  x, y, d}], per-tool `cf` indexes; `--conflicts-csv` exports; `--no-conflicts`
  skips. Viewer: dashed warning outline on involved tools, a diamond marker
  per conflict that appears on the conflict's day, header pill, Conflicts
  tab, tooltip lines.
* **Storage assignment** (v1.11) — zones come from `--storage NAME=X0,Y0,X1,Y1[:CAP]`
  or `--storage-layer NAME` (`zones_from_layers`: closed polylines / block
  inserts on the layer, named by a text inside them). A tool needs storage
  when its arrival date (`--arrival-col`, auto-detected; else `--storage-lead`
  days before move-in, default 7) precedes its move-in. `assign_storage`
  cuts each zone into slots (`--storage-cell`, default 90th percentile
  footprint × 1.15), places tools in arrival order (longest stay first) into
  the nearest zone with a free slot for the whole interval (a `--storage-col`
  wish first); `:CAP` limits concurrent tools. Payload: `storage.zones`, per
  tool `ad` (arrival), `sz`/`sl`/`sx`/`sy` (zone, slot, SVG position);
  `--storage-csv` exports. Day axis starts at the first arrival.
* **Filters and colour facets** (v1.11) — besides group / area / type, the
  schedule's model, vendor and module columns (`FACET_PATTERNS`) and any
  `--facet-col` become per-tool `f0`, `f1`, … with `payload["fields"]` =
  ordered [{k, l}] and `colorBy` (`--color-by`, default a Module column,
  else group, else area). Viewer: one dropdown per field (≤ 400 values),
  AND-combined; filtered-out tools get class `filt` (hidden) and `recount()`
  rebuilds the prefix sums, chart, legend and lists for the selection.
* Viewer JS — precomputes per-day prefix sums so a playback tick only touches
  tools whose state changed; storage arrivals / departures and conflict
  markers are day EVENTS (`EV`, sorted, prefix-indexed) walked by the same
  `applyRange`; tools are built in rAF chunks; background gets
  `pointer-events:none`; tool tooltips are event-delegated; text labels are
  bucketed by size and culled by zoom.

## Performance invariants (don't regress these)

Real fab DXFs are huge (100k+ entities, thousands of repeated blocks, proxy
objects from AutoCAD MEP). The v1.1 rewrite fixed a real-data test that was
"extremely slow" and showed tools without locations:

1. **Never explode geometry per insert for output.** Background inserts are
   `<use>` references to one shared symbol per block definition (44× smaller
   output on realistic data: 27.7 MB → 0.6 MB).
2. **Tolerances come from a robust drawing size** (`_robust_size()`: 2%–98%
   spread of entity anchor points). Using raw `bbox.extents` breaks badly:
   one stray entity a mile off the plan coarsens flatten/simplify tolerances
   until all real geometry collapses.
3. **Tools with no drawable geometry (proxy blocks) are placed at their block
   insertion point** (`ins["pt"]`), never at (0,0), and the CLI warns with
   the PROXYGRAPHICS re-export hint.
4. **Auto-fit bounds ignore far-away junk**: only background pieces
   intersecting the inflated tool bbox contribute to the initial view.
5. **Playback is O(changed)**: `applyRange` walks the day-sorted prefix
   between two day indexes; counts come from `NEWP` prefix sums.
   **And it never repaints the floor plan**: the background lives in its own
   `#bgSvg` with `will-change:transform` (own compositor layer, rasterized
   once), tools/halos/ghosts in `#mapSvg` on top; `setVB` syncs both
   viewBoxes; PNG export composes the two. Measured 3.6 -> 59 fps during
   playback on a heavy drawing under software rendering — a single shared
   svg re-rasterizes everything on every halo pulse and class toggle.
6. Polylines are RDP-simplified (`simplify_poly`, tol ≈ size/6000,
   `--simplify` overrides, `0` disables); background texts are capped
   (`--max-texts`, largest kept); coordinate decimals auto-scale
   (`digits_for`).

## Testing

```
pip install ezdxf openpyxl playwright     # playwright optional (browser checks)
python tests/run_tests.py                 # or --no-browser
```

`tests/run_tests.py` generates synthetic fixtures into a temp folder and runs
the real CLI: attributed tool blocks (rotated / scaled / mirrored) checked
against ezdxf-exploded ground-truth bboxes, proxy blocks without geometry,
a stray entity far away, labels inside anonymous twin blocks, a split label
(`--trace`), loose matching, type mode, a label placed in both halves of the
fab (`--tool-layer` / `--prefer`), two areas either side of a move-in
passageway (`--aisle`), `--max-mb` with symbol remap, tiles /
fine-detail / bitmap zoom, PNG export, playback cost, an inlined whole-floor
block, footprint reduction, every conflict kind plus storage slots
(`conflicts` scenario), and the facet filters. Run it before every push; extend it with a
new scenario whenever a real-data problem is fixed. The remote env has
Chromium at `/opt/pw-browsers/chromium` (`FABSIM_CHROMIUM` overrides).
The tool itself stays a single file — users copy only `fab_movein.py`.

## Known user-environment issues

* **WinError 10053** while viewing: not from this script (it opens no
  sockets) — it's a local preview server (VS Code Live Server, Jupyter,
  `python -m http.server`) aborting on a huge file. Answer: double-click the
  HTML directly; keep output small.
* DWG input is unsupported (ezdxf reads DXF only) — SAVEAS AutoCAD 2013 DXF.
