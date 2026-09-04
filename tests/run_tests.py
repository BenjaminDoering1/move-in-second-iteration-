#!/usr/bin/env python3
"""
Regression suite for fab_movein.py.

    pip install ezdxf openpyxl            # required
    pip install playwright                # optional: browser checks (needs Chromium)
    python tests/run_tests.py             # everything
    python tests/run_tests.py --no-browser
    python tests/run_tests.py --only=small,heavy     # scenarios: small blocktext prefer aisle heavy xref conflicts

Synthetic DXF/XLSX fixtures are generated into a temporary folder; every
scenario runs the real command line and asserts on its output, and -- when
Playwright is importable -- on the generated HTML inside Chromium.
Set FABSIM_CHROMIUM=/path/to/chrome to use a specific browser binary.
"""
import datetime as dt
import json
import os
import random
import re
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPT = os.path.join(os.path.dirname(HERE), "fab_movein.py")
NO_BROWSER = "--no-browser" in sys.argv
ONLY = {n for a in sys.argv[1:] if a.startswith("--only=") for n in a[7:].split(",")}
FAILS, PASSES = [], 0


def check(name, cond, detail=""):
    global PASSES
    print(("  ok    " if cond else "  FAIL  ") + name + (f"   {detail}" if detail and not cond else ""))
    if cond:
        PASSES += 1
    else:
        FAILS.append(name)


def run(*extra):
    """Run fab_movein.py; returns (exit code, combined output)."""
    p = subprocess.run([sys.executable, SCRIPT, *extra], capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    return p.returncode, p.stdout + p.stderr


def xlsx(path, header, rows):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(header)
    for r in rows:
        ws.append(list(r))
    wb.save(path)


def payload(html_path):
    """The JSON payload embedded in a generated HTML file."""
    s = open(html_path, encoding="utf-8").read()
    i = s.index('<script id="simData" type="application/json">') + len('<script id="simData" type="application/json">')
    return json.loads(s[i:s.index("</script>", i)])


def _layer_geo(L):
    return sum(len(T[2]) + len(T[3]) for T in L["tiles"])


def tool_geometry_bytes(P):
    """Bytes spent on tool footprints: shared symbols + the TOOLS layer geometry."""
    n = sum(len(d) for d in P["syms"])
    for L in P["layers"]:
        if L["n"] == "TOOLS":
            n += sum(len(T[2]) + len(T[3]) for T in L["tiles"])
    return n


# ----------------------------------------------------------------- fixtures
def make_small(d):
    """Attributed tool blocks (rotated / scaled / mirrored), a proxy-like
    block without geometry, repeated background blocks, dense polylines,
    a stray entity far away. Ground-truth bboxes come from ezdxf itself."""
    import ezdxf
    from ezdxf import bbox as ezbbox
    random.seed(42)
    doc = ezdxf.new("R2013")
    msp = doc.modelspace()
    for n in ("WALLS", "GRID", "PIPING", "LABELS", "TOOLS", "JUNK"):
        doc.layers.add(n)
    FW, FH = 4000.0, 2400.0
    msp.add_lwpolyline([(0, 0), (FW, 0), (FW, FH), (0, FH)], close=True, dxfattribs={"layer": "WALLS"})
    for i in range(0, int(FW) + 1, 200):
        msp.add_line((i, 0), (i, FH), dxfattribs={"layer": "GRID"})
    for _ in range(150):
        x0, y0 = random.uniform(0, FW), random.uniform(0, FH)
        x1, y1 = x0 + random.uniform(-400, 400), y0 + random.uniform(-400, 400)
        msp.add_lwpolyline([(x0 + (x1 - x0) * t / 40.0, y0 + (y1 - y0) * t / 40.0) for t in range(41)],
                           dxfattribs={"layer": "PIPING"})
    col = doc.blocks.new("COLUMN")
    col.add_lwpolyline([(-2, -2), (2, -2), (2, 2), (-2, 2)], close=True)
    col.add_line((-2, -2), (2, 2))
    for _ in range(400):
        msp.add_blockref("COLUMN", (random.uniform(0, FW), random.uniform(0, FH)), dxfattribs={"layer": "GRID"})
    ta = doc.blocks.new("TOOL_A")
    ta.add_lwpolyline([(0, 0), (30, 0), (30, 18), (0, 18)], close=True)
    ta.add_circle((15, 9), 6)
    ta.add_attdef("TOOL_ID", (15, 9), dxfattribs={"height": 3})
    tb = doc.blocks.new("TOOL_B")
    tb.add_lwpolyline([(0, 0), (22, 0), (22, 22), (11, 30), (0, 22)], close=True)
    tb.add_attdef("TOOL_ID", (11, 11), dxfattribs={"height": 3})
    doc.blocks.new("PROXY_TOOL").add_attdef("TOOL_ID", (0, 0), dxfattribs={"height": 3})

    start = dt.date(2026, 3, 2)
    rows, expect, proxy = [], {}, {}
    ii = 0

    def add_tool(i, name, x, y, rot=0.0, sx=1.0, sy=1.0):
        tid = f"T-{i:04d}"
        ins = msp.add_blockref(name, (x, y), dxfattribs={"layer": "TOOLS", "rotation": rot,
                                                        "xscale": sx, "yscale": sy})
        ins.add_auto_attribs({"TOOL_ID": tid})
        rows.append((tid, f"Tool {name} #{i}", start + dt.timedelta(days=(i * 7919) % 45),
                     f"P{1 + i % 4}", f"Bay {1 + i % 6}", f"{name}-{1 + i % 5}",
                     ("ACME", "Globex", "Initech")[i % 3], f"Module {1 + (i // 8) % 3}"))
        return tid, ins

    for _ in range(60):
        ii += 1
        name = "TOOL_A" if ii % 2 else "TOOL_B"
        sx = random.choice([1, 1, 2, -1])
        tid, ins = add_tool(ii, name, random.uniform(60, FW - 60), random.uniform(60, FH - 60),
                            random.choice([0, 90, 180, 37.5]), sx, abs(sx))
        if ii in (5, 17, 42):
            ext = ezbbox.extents(ins.virtual_entities(), fast=False)
            expect[tid] = [ext.extmin.x, ext.extmin.y, ext.extmax.x, ext.extmax.y]
    for k in range(4):
        ii += 1
        x, y = 500 + k * 300, 2000
        tid, _ = add_tool(ii, "PROXY_TOOL", x, y)
        proxy[tid] = [x, y]
    for b in range(6):
        msp.add_text(f"BAY {b + 1}", dxfattribs={"layer": "LABELS", "height": 60}
                     ).set_placement((FW / 6 * b + 150, FH - 150))
    msp.add_line((5_000_000, 5_000_000), (5_000_100, 5_000_100), dxfattribs={"layer": "JUNK"})
    doc.saveas(os.path.join(d, "small.dxf"))
    rows.append(("T-9999", "Ghost", start, "P4", "Bay 9", "X", "ACME", "Module 1"))
    xlsx(os.path.join(d, "small.xlsx"), ["Tool ID", "Tool Name", "Move In Date", "Phase", "Bay",
                                         "Model", "Vendor", "Module"], rows)
    return {"tools": expect, "proxy": proxy, "n": len(rows)}


def make_blocktext(d):
    """Labels drawn as TEXT inside anonymous blocks (copy-paste twins), plus
    a label split into two texts ('DT_545' + 'R')."""
    import ezdxf
    random.seed(11)
    doc = ezdxf.new("R2013")
    msp = doc.modelspace()
    for n in ("WALLS", "TOOLS"):
        doc.layers.add(n)
    msp.add_lwpolyline([(0, 0), (3000, 0), (3000, 1800), (0, 1800)], close=True, dxfattribs={"layer": "WALLS"})

    def mk(name, w, h, labels):
        b = doc.blocks.new(name)
        b.add_lwpolyline([(0, 0), (w, 0), (w, h), (0, h)], close=True)
        for i, lab in enumerate(labels):
            b.add_text(lab, dxfattribs={"height": 4}).set_placement((w / 2 + 6 * i, h / 2))
    mk("A$C39A4F2E1", 30, 18, ["DT_112"])
    mk("A$C39A4F999", 30, 18, ["DT_112"])
    mk("A$C7B22D901", 24, 24, ["PMP_7"])
    mk("A$C0000SPLT", 24, 24, ["DT_545", "R"])
    for blk, n in (("A$C39A4F2E1", 6), ("A$C39A4F999", 4), ("A$C7B22D901", 5), ("A$C0000SPLT", 3)):
        for _ in range(n):
            msp.add_blockref(blk, (random.uniform(100, 2900), random.uniform(100, 1700)), dxfattribs={"layer": "TOOLS"})
    msp.add_text("DT_112", dxfattribs={"height": 8}).set_placement((100, 1750))
    doc.saveas(os.path.join(d, "bt.dxf"))
    start = dt.date(2026, 5, 4)
    xlsx(os.path.join(d, "bt.xlsx"), ["Block_ID", "Move In Date"],
         [(lab, start + dt.timedelta(days=i)) for i, lab in
          enumerate(["DT_112"] * 10 + ["PMP_7"] * 5 + ["DT_545R"] * 2)])
    xlsx(os.path.join(d, "bt_loose.xlsx"), ["Block_ID", "Move In Date"],
         [(lab, start + dt.timedelta(days=i)) for i, lab in
          enumerate(["DT-112"] * 4 + ["PMP 7"] * 3)])


def make_prefer(d):
    """One label placed in both halves of the fab (north / south), the
    schedule covering only the north half -- the real-data case where a row
    was drawn on the wrong copy of its label."""
    import ezdxf
    doc = ezdxf.new("R2013")
    msp = doc.modelspace()
    for n in ("WALLS", "TOOLS", "0_ENG 3K tool"):
        doc.layers.add(n)
    msp.add_lwpolyline([(0, 0), (3000, 0), (3000, 1800), (0, 1800)], close=True, dxfattribs={"layer": "WALLS"})
    msp.add_line((0, 900), (3000, 900), dxfattribs={"layer": "WALLS"})

    def mk(name, label):
        b = doc.blocks.new(name)
        b.add_lwpolyline([(0, 0), (30, 0), (30, 18), (0, 18)], close=True)
        b.add_text(label, dxfattribs={"height": 4}).set_placement((15, 9))
    mk("A$C11111111", "DT_112")
    mk("A$C22222222", "PMP_7")
    # DT_112: one copy north (y=1500, on the project's tool layer), one south
    # (y=300, further WEST, so the old west-to-east order picked it). PMP_7
    # exists only in the south.
    msp.add_blockref("A$C11111111", (1200, 1500), dxfattribs={"layer": "0_ENG 3K tool"})
    msp.add_blockref("A$C11111111", (400, 300), dxfattribs={"layer": "TOOLS"})
    msp.add_blockref("A$C22222222", (2000, 200), dxfattribs={"layer": "TOOLS"})
    doc.saveas(os.path.join(d, "pf.dxf"))
    xlsx(os.path.join(d, "pf.xlsx"), ["Block_ID", "Move In Date"],
         [("DT_112", dt.date(2026, 6, 1)), ("PMP_7", dt.date(2026, 6, 2))])


def make_aisle(d):
    """The ENG line: two areas above and below a move-in passageway at
    Y = 1000. Six copies of one label, three per side, plus a copy of a
    second label in each area; the schedule has one row per day."""
    import ezdxf
    doc = ezdxf.new("R2013")
    msp = doc.modelspace()
    for n in ("WALLS", "TOOLS"):
        doc.layers.add(n)
    for n in ("00_Dummy", "BAY-ID"):
        doc.layers.add(n)
    msp.add_lwpolyline([(0, 0), (3000, 0), (3000, 2000), (0, 2000)], close=True, dxfattribs={"layer": "WALLS"})
    msp.add_line((0, 1000), (3000, 1000), dxfattribs={"layer": "WALLS"})
    # the pathway (940..1060) is an ANSI31 hatch on layer 00_Dummy, which also
    # holds other objects (a line far north and a solid hatch), as on the real floor3
    h = msp.add_hatch(dxfattribs={"layer": "00_Dummy"})
    h.set_pattern_fill("ANSI31", scale=20)
    h.paths.add_polyline_path([(100, 940), (2900, 940), (2900, 1060), (100, 1060)], is_closed=True)
    msp.add_line((100, 1900), (2900, 1900), dxfattribs={"layer": "00_Dummy"})
    sol = msp.add_hatch(dxfattribs={"layer": "00_Dummy"})
    sol.paths.add_polyline_path([(2700, 1700), (2800, 1700), (2800, 1800), (2700, 1800)], is_closed=True)
    # the bay ID text that marks the split between the upper and the lower fab
    msp.add_text("XPH-D", dxfattribs={"height": 20, "layer": "BAY-ID"}).set_placement((1500, 250))
    msp.add_text("XPH-D", dxfattribs={"height": 5, "layer": "NOTES"}).set_placement((2800, 1900))  # a smaller stray copy
    b = doc.blocks.new("A$CAISLE001")
    b.add_lwpolyline([(0, 0), (30, 0), (30, 18), (0, 18)], close=True)
    b.add_text("DT_112", dxfattribs={"height": 4}).set_placement((15, 9))
    for y in (1200, 1400, 1600, 800, 600, 400):      # north area rows, then south area rows
        msp.add_blockref("A$CAISLE001", (1000, y), dxfattribs={"layer": "TOOLS"})
    msp.add_blockref("A$CAISLE001", (1000, 100), dxfattribs={"layer": "TOOLS"})   # lower fab copy
    doc.saveas(os.path.join(d, "ai.dxf"))
    start = dt.date(2026, 8, 3)
    xlsx(os.path.join(d, "ai.xlsx"), ["Block_ID", "Move In Date"],
         [("DT_112", start + dt.timedelta(days=i)) for i in range(6)])


def make_heavy(d):
    """Two heavy polyline layers, a shared fixture block placed 3000 times,
    another placed 200 times, 300 tools -- for --max-mb, tiles and zoom."""
    import ezdxf
    random.seed(5)
    doc = ezdxf.new("R2013")
    msp = doc.modelspace()
    for n in ("WALLS", "PIPING", "HVAC", "FIXA", "FIXB", "TOOLS"):
        doc.layers.add(n)
    FW, FH = 4000.0, 2400.0
    msp.add_lwpolyline([(0, 0), (FW, 0), (FW, FH), (0, FH)], close=True, dxfattribs={"layer": "WALLS"})
    for layer, n in (("PIPING", 3000), ("HVAC", 2500)):
        for _ in range(n):
            x0, y0 = random.uniform(0, FW), random.uniform(0, FH)
            msp.add_lwpolyline([(x0 + random.uniform(-150, 150), y0 + random.uniform(-150, 150)) for _ in range(24)],
                               dxfattribs={"layer": layer})
    fa = doc.blocks.new("FIXA"); fa.add_circle((0, 0), 3)
    fa.add_lwpolyline([(-3, -3), (3, -3), (3, 3), (-3, 3)], close=True)
    fb = doc.blocks.new("FIXB"); fb.add_circle((0, 0), 2); fb.add_line((-2, 0), (2, 0))
    for _ in range(3000):
        msp.add_blockref("FIXA", (random.uniform(0, FW), random.uniform(0, FH)), dxfattribs={"layer": "FIXA"})
    for _ in range(200):
        msp.add_blockref("FIXB", (random.uniform(0, FW), random.uniform(0, FH)), dxfattribs={"layer": "FIXB"})
    ta = doc.blocks.new("TOOL_A")
    ta.add_lwpolyline([(0, 0), (30, 0), (30, 18), (0, 18)], close=True)
    ta.add_attdef("TOOL_ID", (15, 9), dxfattribs={"height": 3})
    rows = []
    start = dt.date(2026, 6, 1)
    for i in range(300):
        ins = msp.add_blockref("TOOL_A", (random.uniform(60, FW - 60), random.uniform(60, FH - 60)),
                               dxfattribs={"layer": "TOOLS"})
        ins.add_auto_attribs({"TOOL_ID": f"T-{i:04d}"})
        rows.append((f"T-{i:04d}", start + dt.timedelta(days=i % 40)))
    doc.saveas(os.path.join(d, "heavy.dxf"))
    xlsx(os.path.join(d, "heavy.xlsx"), ["Tool ID", "Move In Date"], rows)


def make_xref_detail(d):
    """A whole-floor block placed once (must be inlined and stay full detail),
    identical twin blocks under different names (must share one symbol),
    and detailed tool blocks that should reduce to footprints."""
    import ezdxf
    random.seed(13)
    doc = ezdxf.new("R2013")
    msp = doc.modelspace()
    for n in ("WALLS", "XREF-ARCH", "FIXT", "TOOLS"):
        doc.layers.add(n)
    FW, FH = 4000.0, 2400.0
    msp.add_lwpolyline([(0, 0), (FW, 0), (FW, FH), (0, FH)], close=True, dxfattribs={"layer": "WALLS"})
    arch = doc.blocks.new("FLOOR_ARCH")
    for _ in range(4000):
        x0, y0 = random.uniform(0, FW), random.uniform(0, FH)
        arch.add_lwpolyline([(x0 + random.uniform(-60, 60), y0 + random.uniform(-60, 60)) for _ in range(6)])
    msp.add_blockref("FLOOR_ARCH", (0, 0), dxfattribs={"layer": "XREF-ARCH"})
    for i in range(10):
        tw = doc.blocks.new(f"A$C00TWIN{i}")
        tw.add_circle((0, 0), 4); tw.add_line((-4, 0), (4, 0))
        for _ in range(3):
            msp.add_blockref(f"A$C00TWIN{i}", (random.uniform(0, FW), random.uniform(0, FH)), dxfattribs={"layer": "FIXT"})
    rows = []
    start = dt.date(2026, 7, 6)
    for i in range(60):
        b = doc.blocks.new(f"A$CTOOL{i:03d}")
        b.add_lwpolyline([(0, 0), (30, 0), (30, 18), (0, 18)], close=True)
        for _ in range(120):                     # internal detail above the simplify tolerance
            x, y = random.uniform(1, 27), random.uniform(1, 15)
            b.add_line((x, y), (x + random.uniform(1, 2.5), y + random.uniform(1, 2.5)))
        b.add_text("DT_%03d" % (i // 3), dxfattribs={"height": 3}).set_placement((15, 9))
        msp.add_blockref(f"A$CTOOL{i:03d}", (random.uniform(60, FW - 60), random.uniform(60, FH - 60)),
                         dxfattribs={"layer": "TOOLS"})
        if i % 3 == 0:
            rows.append(("DT_%03d" % (i // 3), start + dt.timedelta(days=i % 20)))
    doc.saveas(os.path.join(d, "xd.dxf"))
    xlsx(os.path.join(d, "xd.xlsx"), ["Block_ID", "Move In Date"], rows)


def make_conflicts(d):
    """Move-in conflicts of every kind and storage zones: two tools drawn on
    the same floor space, four neighbours moving in within days of each
    other, a tool boxed in by four earlier tools, six move-ins on one day,
    tools delivered weeks before their move-in (some with no room left), a
    laydown zone drawn on its own layer (closed polyline + label) and one
    given on the command line."""
    import ezdxf
    doc = ezdxf.new("R2013")
    msp = doc.modelspace()
    for n in ("WALLS", "TOOLS", "LAYDOWN"):
        doc.layers.add(n)
    msp.add_lwpolyline([(0, 0), (3000, 0), (3000, 1800), (0, 1800)], close=True, dxfattribs={"layer": "WALLS"})
    blk = doc.blocks.new("TOOL")
    blk.add_lwpolyline([(0, 0), (30, 0), (30, 18), (0, 18)], close=True)
    blk.add_attdef("TOOL_ID", (15, 9), dxfattribs={"height": 3})
    msp.add_lwpolyline([(2600, 1500), (2840, 1500), (2840, 1620), (2600, 1620)], close=True,
                       dxfattribs={"layer": "LAYDOWN"})
    msp.add_text("LAYDOWN NORTH", dxfattribs={"height": 8, "layer": "LAYDOWN"}).set_placement((2610, 1560))
    msp.add_lwpolyline([(2600, 1400), (2840, 1400), (2840, 1450)], dxfattribs={"layer": "LAYDOWN"})  # open: no zone

    start = dt.date(2026, 9, 7)
    rows = []

    def tool(tid, x, y, day, arrive=None, rot=0):
        ins = msp.add_blockref("TOOL", (x, y), dxfattribs={"layer": "TOOLS", "rotation": rot})
        ins.add_auto_attribs({"TOOL_ID": tid})
        rows.append((tid, start + dt.timedelta(days=day),
                     (start + dt.timedelta(days=arrive)) if arrive is not None else None, "Bay 1"))

    tool("OV-A", 200, 200, 2, arrive=0)                      # overlap
    tool("OV-B", 215, 205, 5, arrive=1)
    for i in range(4):                                        # crowding: 4-unit gaps, days 10..12
        tool(f"CR-{i + 1}", 600 + i * 34, 200, 10 + (i % 3))
    tool("BX-N", 1000, 520, 20); tool("BX-S", 1000, 480, 21)  # boxed in: 2-unit gaps all round
    tool("BX-E", 1050, 500, 22, rot=90); tool("BX-W", 998, 500, 23, rot=90)
    tool("BX-C", 1000, 500, 30)
    for i in range(6):                                        # capacity: 6 on one day, far apart
        tool(f"CAP-{i + 1}", 200 + i * 400, 1200, 40)
    for i in range(8):                                        # storage: delivered day 50, move in 60..67
        tool(f"ST-{i + 1}", 1500 + (i % 4) * 120, 800 + (i // 4) * 200, 60 + i, arrive=50)
    tool("LATE", 2400, 300, 45, arrive=48)                    # arrives after move-in: no storage
    doc.saveas(os.path.join(d, "cf.dxf"))
    xlsx(os.path.join(d, "cf.xlsx"), ["Tool ID", "Move In Date", "Delivery Date", "Bay"], rows)


# ----------------------------------------------------------------- browser
def browser():
    if NO_BROWSER:
        return None
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("  (playwright not installed -- browser checks skipped)")
        return None
    return sync_playwright


def open_page(pw, path):
    exe = os.environ.get("FABSIM_CHROMIUM") or ("/opt/pw-browsers/chromium"
                                                if os.path.exists("/opt/pw-browsers/chromium") else None)
    b = pw.chromium.launch(executable_path=exe) if exe else pw.chromium.launch()
    pg = b.new_page(viewport={"width": 1400, "height": 900})
    errs = []
    pg.on("pageerror", lambda e: errs.append(str(e)))
    pg.goto("file://" + os.path.abspath(path))
    pg.wait_for_function("!document.getElementById('mapTitle').textContent.includes('…')", timeout=180000)
    return b, pg, errs


def scrub(pg, where):
    pg.evaluate(f"""(() => {{ const sc = document.getElementById('scrub');
        sc.value = {where}; sc.dispatchEvent(new Event('input')); }})()""")
    pg.wait_for_timeout(250)


# ----------------------------------------------------------------- scenarios
def scenario_small(d, expect):
    print("\n[small] ID matching, placement, proxy fallback, viewer basics")
    out = os.path.join(d, "small.html")
    code, log = run("--dxf", f"{d}/small.dxf", "--schedule", f"{d}/small.xlsx", "-o", out)
    check("run succeeds", code == 0, log[-800:])
    check("64 of 65 matched", "Matched 64 of 65" in log, log)
    check("4 proxy tools placed as boxes", "4 matched tool(s) had no drawable geometry" in log)
    check("stray junk left out of the fit", "far-away drawing pieces left out" in log)
    check("matched on the attribute", 'block attribute "TOOL_ID"' in log)
    check("model / vendor / module picked up as filters", "Filters     : Model, Vendor, Module" in log, log[-900:])
    P = payload(out)
    check("viewer fields: group, area and the three facets, coloured by Module",
          [f["l"] for f in P["fields"]] == ["Phase", "Bay", "Model", "Vendor", "Module"]
          and P["colorBy"] == [f["k"] for f in P["fields"] if f["l"] == "Module"][0], str(P["fields"]) + P["colorBy"])
    code, log = run("--dxf", f"{d}/small.dxf", "--schedule", f"{d}/small.xlsx", "--color-by", "Vendor",
                    "--facet-col", "Tool Name", "-o", os.path.join(d, "small2.html"))
    P2 = payload(os.path.join(d, "small2.html"))
    check("--color-by / --facet-col", P2["colorBy"] == [f["k"] for f in P2["fields"] if f["l"] == "Vendor"][0]
          and any(f["l"] == "Tool Name" for f in P2["fields"]), str(P2["fields"]) + P2["colorBy"])
    code, log = run("--dxf", f"{d}/small.dxf", "--schedule", f"{d}/small.xlsx", "--color-by", "Nope", "-o", out)
    check("--color-by unknown column rejected", code != 0 and "not one of the viewer fields" in log)
    code, log = run("--dxf", f"{d}/small.dxf", "--schedule", f"{d}/small.xlsx", "--inspect")
    check("--inspect runs", code == 0 and "attrib:TOOL_ID" in log)
    code, log = run("--dxf", f"{d}/small.dxf", "--schedule", f"{d}/small.xlsx", "--layers")
    check("--layers reports placements", "scheduled tools" in log and re.search(r"TOOLS\s+64", log))
    code, log = run("--dxf", f"{d}/small.dxf", "--schedule", f"{d}/small.xlsx", "--trace", "T-0005", "-o", out)
    check("--trace explains an assigned tool", "assigned : T-0005" in log)

    pw = browser()
    if not pw:
        return
    with pw() as p:
        b, pg, errs = open_page(p, out)
        for tid, (x0, y0, x1, y1) in expect["tools"].items():
            got = pg.evaluate(f"(() => {{ const t = window.DATA.tools.find(t => t.id === '{tid}'); "
                              f"return t ? [t.x, t.y, t.w, t.h] : null; }})()")
            ok = got and abs(got[0] - x0) < 1.5 and abs(got[1] + y1) < 1.5 and abs(got[2] - (x1 - x0)) < 3
            check(f"{tid} placed where ezdxf says", bool(ok), f"{got} vs {x0:.1f},{-y1:.1f}")
        for tid, (px, py) in expect["proxy"].items():
            got = pg.evaluate(f"(() => {{ const t = window.DATA.tools.find(t => t.id === '{tid}'); "
                              f"return t ? [t.x + t.w/2, -(t.y + t.h/2)] : null; }})()")
            check(f"{tid} proxy at its insertion point", bool(got) and abs(got[0] - px) < 2 and abs(got[1] - py) < 2)
        vb = pg.evaluate("document.getElementById('mapSvg').getAttribute('viewBox')")
        check("view excludes the stray entity", float(vb.split()[2]) < 100000, vb)
        scrub(pg, "sc.max")
        n_in, n_tools = pg.evaluate("[document.querySelectorAll('g.tool.in').length, window.DATA.tools.length]")
        check("all tools visible at the last day", n_in == n_tools, f"{n_in}/{n_tools}")
        scrub(pg, 0)
        check("no tools on day 0", pg.evaluate("document.querySelectorAll('g.tool.in').length") == 0)
        pg.fill("#search", "T-0042"); pg.press("#search", "Enter"); pg.wait_for_timeout(400)
        vb2 = pg.evaluate("document.getElementById('mapSvg').getAttribute('viewBox')")
        check("search + Enter zooms in", float(vb2.split()[2]) < float(vb.split()[2]) / 3)
        pg.wait_for_timeout(300)
        synced = pg.evaluate("document.getElementById('bgSvg').getAttribute('viewBox') === "
                             "document.getElementById('mapSvg').getAttribute('viewBox')")
        check("background view committed after zoom", synced)
        tick = pg.evaluate("(() => { const t0 = performance.now(); for (let i = 1; i <= 40; i++) setDay(i); "
                           "return (performance.now() - t0) / 40; })()")
        check("day stepping is cheap", tick < 25, f"{tick:.1f} ms")
        pg.fill("#search", "")
        nsel = pg.evaluate("document.querySelectorAll('#filters select').length")
        check("one filter dropdown per field", nsel == 5, str(nsel))
        check("colour legend shows the modules", pg.evaluate("[...document.querySelectorAll('#legend .lg-chip')].map(c => c.textContent)").__len__() == 3)
        scrub(pg, "sc.max")
        pg.select_option("#filters select:nth-child(4)", "Globex"); pg.wait_for_timeout(300)   # Vendor
        f1 = pg.evaluate("""(() => ({ total: document.getElementById('statTotal').textContent,
            n: document.getElementById('statN').textContent, filt: document.querySelectorAll('g.tool.filt').length,
            title: document.getElementById('mapTitle').textContent, on: document.querySelectorAll('#filters select.on').length,
            rows: document.querySelectorAll('#toolList .tl-row').length }))()""")
        check("vendor filter hides the other tools and recounts", f1["total"] == "22" and f1["n"] == "22" and f1["filt"] == 42
              and f1["title"].startswith("22 of 64 tools") and f1["on"] == 1, str(f1))
        pg.select_option("#filters select:nth-child(5)", "Module 2"); pg.wait_for_timeout(300)  # Module (AND)
        f2 = pg.evaluate("[document.getElementById('statTotal').textContent, document.querySelectorAll('g.tool.filt').length]")
        check("filters combine", 0 < int(f2[0]) < 22 and f2[1] == 64 - int(f2[0]), str(f2))
        pg.click("#filters .clearf"); pg.wait_for_timeout(300)
        f3 = pg.evaluate("[document.getElementById('statTotal').textContent, document.querySelectorAll('g.tool.filt').length, document.querySelectorAll('#filters select.on').length]")
        check("clear restores everything", f3[0] == "64" and f3[1] == 0 and f3[2] == 0, str(f3))
        pg.fill("#search", "globex"); pg.wait_for_timeout(200)
        check("search matches facet values", pg.evaluate("window.DATA.tools.filter(isSearchMatch).length") == 22)
        pg.fill("#search", "")
        with pg.expect_download(timeout=30000) as dl:
            pg.click("#pngBtn")
        check("PNG export works", os.path.getsize(dl.value.path()) > 10000)
        check("no JS errors", not errs, "; ".join(errs[:2]))
        b.close()


def scenario_blocktext(d):
    print("\n[blocktext] labels inside blocks, type mode, loose matching, split labels")
    out = os.path.join(d, "bt.html")
    code, log = run("--dxf", f"{d}/bt.dxf", "--schedule", f"{d}/bt.xlsx", "--type-col", "Block_ID",
                    "--trace", "DT_545R", "-o", out)
    check("blocktext strategy chosen", "text inside the tool block" in log)
    check("15 of 17 assigned (split label unmatched)", "Matched 15 of 17" in log, log[-600:])
    check("split label diagnosed", "similar labels that DO exist" in log and "'DT_545'" in log)
    ids = [t["id"] for t in payload(out)["tools"]]
    check("repeated labels numbered", "DT_112" in ids and "DT_112 (2)" in ids and "DT_112 (10)" in ids,
          ", ".join(sorted(ids)[:12]))
    code, log = run("--dxf", f"{d}/bt.dxf", "--schedule", f"{d}/bt_loose.xlsx", "--type-col", "Block_ID", "--inspect")
    check("near-miss report", "Near misses" in log and "DT-112" in log)
    code, log = run("--dxf", f"{d}/bt.dxf", "--schedule", f"{d}/bt_loose.xlsx", "--type-col", "Block_ID",
                    "--loose", "-o", os.path.join(d, "bt_loose.html"))
    check("--loose matches punctuation drift", "Matched 7 of 7" in log, log[-400:])
    pw = browser()
    if not pw:
        return
    with pw() as p:
        b, pg, errs = open_page(p, out)
        n, distinct, ty = pg.evaluate("(() => { const ts = window.DATA.tools; return [ts.length, "
                                      "new Set(ts.map(t => t.x + ',' + t.y)).size, ts.filter(t => t.ty).length]; })()")
        check("each row on its own placement", n == 15 and distinct == 15 and ty == 15, f"{n}/{distinct}/{ty}")
        pg.fill("#search", "DT_112"); pg.wait_for_timeout(200)
        check("search matches by type", pg.evaluate("window.DATA.tools.filter(isSearchMatch).length") == 10)
        check("no JS errors", not errs, "; ".join(errs[:2]))
        b.close()


def scenario_prefer(d):
    print("\n[prefer] a label placed in both halves: --prefer picks the copy the schedule means")
    out = os.path.join(d, "pf.html")
    dwg_y = lambda P, tid: -[t for t in P["tools"] if t["id"] == tid][0]["y"]   # SVG y -> drawing y

    code, log = run("--dxf", f"{d}/pf.dxf", "--schedule", f"{d}/pf.xlsx", "--type-col", "Block_ID",
                    "--trace", "DT_112", "-o", out)
    check("default order documented in --trace", "order    : west to east" in log, log[-500:])
    check("default takes the western (south) copy", dwg_y(payload(out), "DT_112") < 400)

    code, log = run("--dxf", f"{d}/pf.dxf", "--schedule", f"{d}/pf.xlsx", "--type-col", "Block_ID",
                    "--prefer", "north", "--trace", "DT_112", "-o", out)
    check("--prefer north reported", "Placements  : north first" in log and "order    : north first" in log, log[-700:])
    check("--prefer north takes the north copy", dwg_y(payload(out), "DT_112") > 1400, log[-400:])
    check("both rows matched", "Matched 2 of 2" in log)

    code, log = run("--dxf", f"{d}/pf.dxf", "--schedule", f"{d}/pf.xlsx", "--type-col", "Block_ID",
                    "--prefer", "south", "-o", out)
    check("--prefer south takes the south copy", dwg_y(payload(out), "DT_112") < 400)

    code, log = run("--dxf", f"{d}/pf.dxf", "--schedule", f"{d}/pf.xlsx", "--type-col", "Block_ID",
                    "--prefer", "0,900,3000,1800", "--trace", "PMP_7", "-o", out)
    P = payload(out)
    check("rectangle: DT_112 inside first", dwg_y(P, "DT_112") > 1400)
    check("rectangle: PMP_7 falls back outside with a warning",
          dwg_y(P, "PMP_7") < 300 and "outside the --prefer rectangle" in log and "PMP_7" in log, log[-600:])
    check("trace marks in/out", " OUT" in log and "order    : inside" in log)

    code, log = run("--dxf", f"{d}/pf.dxf", "--schedule", f"{d}/pf.xlsx", "--type-col", "Block_ID",
                    "--tool-layer", "0_ENG 3K tool", "--trace", "DT_112", "-o", out)
    P = payload(out)
    check("--tool-layer takes the copy on that layer", dwg_y(P, "DT_112") > 1400, log[-500:])
    check("--tool-layer reported", "Placements  : on layer '0_ENG 3K tool' first" in log
          and "on '0_ENG 3K tool' in" in log and "on 'TOOLS' OUT" in log, log[-800:])
    check("--tool-layer: PMP_7 falls back with a warning",
          dwg_y(P, "PMP_7") < 300 and "not on --tool-layer" in log and "PMP_7" in log, log[-600:])

    code, log = run("--dxf", f"{d}/pf.dxf", "--schedule", f"{d}/pf.xlsx", "--type-col", "Block_ID",
                    "--tool-layer", "0_ENG*", "--prefer", "south", "-o", out)
    check("--tool-layer wildcard wins over --prefer", dwg_y(payload(out), "DT_112") > 1400
          and "on layer '0_ENG*' first, then south first" in log, log[-500:])

    code, log = run("--dxf", f"{d}/pf.dxf", "--schedule", f"{d}/pf.xlsx", "--type-col", "Block_ID",
                    "--tool-layer", "NOPE", "-o", out)
    check("unknown --tool-layer warned", "matches no block insert" in log, log[-500:])

    code, log = run("--dxf", f"{d}/pf.dxf", "--schedule", f"{d}/pf.xlsx", "--type-col", "Block_ID",
                    "--prefer", "somewhere", "-o", out)
    check("bad --prefer rejected", code != 0 and "not understood" in log)

    code, log = run("--dxf", f"{d}/pf.dxf", "--schedule", f"{d}/pf.xlsx", "--prefer", "north",
                    "--match", "blocktext", "-o", out)
    check("--prefer without --type-col warns, both copies drawn",
          "only matter with --type-col" in log and len(payload(out)["tools"]) == 3, log[-500:])

    code, log = run("--dxf", f"{d}/pf.dxf", "--schedule", f"{d}/pf.xlsx", "--inspect")
    check("--inspect shows where inserts are", "most inserts lie in X" in log and "--prefer" in log)


def scenario_aisle(d):
    print("\n[aisle] two areas either side of a passageway fill from their far ends towards it")
    out = os.path.join(d, "ai.html")
    base = ("--dxf", f"{d}/ai.dxf", "--schedule", f"{d}/ai.xlsx", "--type-col", "Block_ID")

    def order(P):
        """Drawing Y of each copy in move-in order (SVG y -> drawing y, centre)."""
        return [round(-(t["y"] + t["h"] / 2)) for t in sorted(P["tools"], key=lambda t: t["day"])]

    code, log = run(*base, "--trace", "DT_112", "-o", out)
    ys = order(payload(out))
    check("default order ignores the passageway", ys == [109, 409, 609, 809, 1209, 1409], str(ys))

    code, log = run(*base, "--aisle", "1000", "--trace", "DT_112", "-o", out)
    ys = order(payload(out))
    north = [y for y in ys if y > 1000]; south = [y for y in ys if y < 1000]
    check("--aisle: north area fills north to south", north == sorted(north, reverse=True), str(ys))
    check("--aisle: south area fills south to north", south == sorted(south), str(ys))
    check("--aisle: farthest copies first overall", ys[0] == 109 and 1609 in ys[:3], str(ys))
    check("--aisle reported", "Placements  : farthest from the passageway at Y = 1,000 first" in log
          and "order    : farthest from" in log, log[-900:])

    code, log = run(*base, "--aisle", "y=1000", "--prefer", "north", "-o", out)
    ys = order(payload(out))
    check("--prefer north + --aisle: north side first, north to south, then south to north",
          ys == [1609, 1409, 1209, 109, 409, 609], str(ys))
    check("combined order reported", "the north side of the passageway at Y = 1,000 first, then farthest" in log, log[-700:])

    code, log = run(*base, "--aisle", "0,1000,3000,1000", "--prefer", "0,1000,3000,2000", "-o", out)
    ys = order(payload(out))
    check("--prefer rectangle + --aisle line: inside first, far end first",
          ys == [1609, 1409, 1209, 109, 409, 609], str(ys))

    code, log = run(*base, "--aisle", "x=1000", "-o", out)
    ys = order(payload(out))
    check("vertical aisle: all copies at X=1015 tie, default order breaks the tie",
          ys == [109, 409, 609, 809, 1209, 1409], str(ys))

    # --- drawing anchors: the pathway object on its layer and the bay-ID text
    code, log = run(*base, "--aisle", "hatch:ANSI31@00_Dummy", "--prefer", "north-of=text:XPH-D@BAY-ID",
                    "--trace", "DT_112", "-o", out)
    P = payload(out)
    ys = order(P)
    check("anchors: upper fab first (north-of the bay text), each area from its far end",
          ys == [1609, 409, 1409, 609, 1209, 809] or ys == [409, 1609, 609, 1409, 809, 1209], str(ys))
    check("anchors printed in the run header",
          "Anchor      : --aisle hatch:ANSI31@00_Dummy  =  hatch 'ANSI31' on layer '00_Dummy': 1 hatch(es), extent X 100..2,900 Y 940..1,060 -> horizontal line at Y = 1,000" in log
          and "Anchor      : --prefer north-of=text:XPH-D@BAY-ID  =  text 'XPH-D' on layer 'BAY-ID' at (1,500, 250)" in log
          and "re-point these options" in log, log[-1500:])
    check("order line names the anchors", "north of text 'XPH-D' on layer 'BAY-ID' first, then farthest from the passageway at Y = 1,000 (hatch 'ANSI31' on layer '00_Dummy')" in log, log[-1200:])
    check("lower-fab copy marked OUT in the trace", "(1,015.0, 109.0) on 'TOOLS' OUT" in log, log[-1500:])
    check("reference lines in the payload for the viewer", len(P["refLines"]) == 2
          and all(r["seg"][1] == r["seg"][3] for r in P["refLines"])
          and any(r["seg"][1] == -1000 for r in P["refLines"]) and any(r["seg"][1] == -250 for r in P["refLines"]), str(P["refLines"]))

    code, log = run(*base, "--aisle", "layer:00_Dummy", "-o", out)
    check("whole-layer anchor warns about the other objects and uses the full extent",
          "holds 3 entities; using the extent of ALL of them" in log and "Y 940..1,900 -> horizontal line at Y = 1,420" in log, log[-900:])
    code, log = run(*base, "--aisle", "hatch:ANSI31@NOPE", "-o", out)
    check("missing hatch stops the run with the patterns found", code != 0 and "no hatch with pattern 'ANSI31' on layer 'NOPE'" in log
          and "That pattern exists on: ANSI31 on '00_Dummy' x1" in log, log[-600:])
    code, log = run(*base, "--inspect")
    check("--inspect lists hatch patterns per layer", "Hatch patterns" in log and "ANSI31" in log and "SOLID" in log
          and "--aisle hatch:PATTERN@LAYER" in log, log[-1500:])
    code, log = run(*base, "--prefer", "north-of=text:XPH-D", "-o", out)
    check("ambiguous text warned, largest used", "appears 2 times" in log and "using the largest at (1,500, 250)" in log, log[-800:])
    code, log = run(*base, "--aisle", "layer:NOPE", "-o", out)
    check("missing anchor layer stops the run", code != 0 and "no layer 'NOPE'" in log and "--inspect lists all layers" in log, log[-500:])
    code, log = run(*base, "--prefer", "north-of=text:XPH-Z@BAY-ID", "-o", out)
    check("missing anchor text stops the run with candidates", code != 0 and "no text 'XPH-Z' on layer 'BAY-ID'" in log
          and "Texts on that layer: XPH-D" in log, log[-600:])
    code, log = run(*base, "--prefer", "east-of=text:XPH-D@BAY-ID", "-o", out)
    check("orientation mismatch rejected", code != 0 and "needs a vertical line" in log, log[-500:])

    code, log = run(*base, "--aisle", "north", "-o", out)
    check("bad --aisle rejected", code != 0 and "not understood" in log)
    code, log = run("--dxf", f"{d}/ai.dxf", "--schedule", f"{d}/ai.xlsx", "--aisle", "1000", "--match", "blocktext", "-o", out)
    check("--aisle without --type-col warns", "only matter with --type-col" in log)


def scenario_heavy(d):
    print("\n[heavy] --max-mb budget, symbol remap, tiles, bitmap zoom, playback")
    out = os.path.join(d, "heavy.html")
    code, log = run("--dxf", f"{d}/heavy.dxf", "--schedule", f"{d}/heavy.xlsx", "--simplify", "0", "-o", out)
    check("heavy run succeeds", code == 0 and "Matched 300 of 300" in log, log[-400:])
    tiny = os.path.join(d, "tiny.html")
    code, log = run("--dxf", f"{d}/heavy.dxf", "--schedule", f"{d}/heavy.xlsx", "--simplify", "0",
                    "--max-mb", "0.15", "-o", tiny)
    check("budget drops heaviest layers first", "dropped layer PIPING" in log and "dropped layer HVAC" in log)
    check("fixture layer dropped, walls kept", "dropped layer FIXA" in log and "dropped layer WALLS" not in log)
    pw = browser()
    if not pw:
        return
    with pw() as p:
        for f, want_uses in ((out, 3200), (tiny, 200)):
            b, pg, errs = open_page(p, f)
            info = pg.evaluate("""(() => ({
                uses: document.querySelectorAll('use[href^="#b"]').length,
                bad: [...document.querySelectorAll('use[href^="#b"]')].filter(u => +u.getAttribute('href').slice(2) >= window.DATA.bsyms.length).length,
                tiles: MAP.tiles.length, fineHidden: document.body.classList.contains('hide-fine') }))()""")
            check(f"{os.path.basename(f)}: placements + valid symbol refs", info["uses"] == want_uses and info["bad"] == 0, str(info))
            check(f"{os.path.basename(f)}: fine detail hidden at fit view", info["fineHidden"])
            pg.evaluate("(() => { const v = MAP.vb; zoomAt(0.08, v.x + v.w * 0.3, v.y + v.h * 0.3); })()")
            pg.wait_for_timeout(30)
            mid = pg.evaluate("document.getElementById('bgSvg').style.transform !== ''")
            pg.wait_for_timeout(400)
            after = pg.evaluate("""(() => ({ t: document.getElementById('bgSvg').style.transform,
                hidden: MAP.tiles.filter(t => t.el.style.display === 'none').length,
                fine: !document.body.classList.contains('hide-fine') }))()""")
            check(f"{os.path.basename(f)}: bitmap zoom then sharp commit with tile culling",
                  mid and after["t"] == "" and after["hidden"] > 0 and after["fine"], str(after))
            scrub(pg, "sc.max")
            check(f"{os.path.basename(f)}: all 300 tools in", pg.evaluate("document.querySelectorAll('g.tool.in').length") == 300)
            check(f"{os.path.basename(f)}: no JS errors", not errs, "; ".join(errs[:2]))
            b.close()


def scenario_xref_detail(d):
    print("\n[xref/detail] inlined whole-floor block, twin symbols, footprint reduction")
    code, log = run("--dxf", f"{d}/xd.dxf", "--schedule", f"{d}/xd.xlsx", "--type-col", "Block_ID", "--layers")
    check("whole-floor block counted on its layer", bool(re.search(r"XREF-ARCH\s+[\d,]+\s+\d+%", log)) and
          "inlined into their layers" in log, log[-900:])
    m = re.search(r"reduced to their footprint", log)
    check("detailed tool blocks reduced", bool(m))
    check("twin blocks share one symbol", "1 shared symbol" in log, log[-600:])
    full = os.path.join(d, "xd_full.html"); auto = os.path.join(d, "xd_auto.html")
    run("--dxf", f"{d}/xd.dxf", "--schedule", f"{d}/xd.xlsx", "--type-col", "Block_ID", "--block-detail", "full", "-o", full)
    run("--dxf", f"{d}/xd.dxf", "--schedule", f"{d}/xd.xlsx", "--type-col", "Block_ID", "-o", auto)
    Pf, Pa = payload(full), payload(auto)
    gf, ga = tool_geometry_bytes(Pf), tool_geometry_bytes(Pa)
    check("auto detail shrinks tool footprints", ga < gf * 0.25, f"full {gf // 1024} KB vs auto {ga // 1024} KB")
    xf = [L for L in Pf["layers"] if L["n"] == "XREF-ARCH"]
    xa = [L for L in Pa["layers"] if L["n"] == "XREF-ARCH"]
    check("whole-floor block keeps full detail", xf and xa and _layer_geo(xf[0]) == _layer_geo(xa[0]))


def scenario_conflicts(d):
    print("\n[conflicts] overlap / crowding / boxed-in / capacity detection, storage slots")
    out = os.path.join(d, "cf.html")
    cf_csv, st_csv = os.path.join(d, "cf.csv"), os.path.join(d, "st.csv")
    base = ("--dxf", f"{d}/cf.dxf", "--schedule", f"{d}/cf.xlsx")
    code, log = run(*base, "--storage-layer", "LAYDOWN", "--storage", "Laydown South=100,50,400,150:2",
                    "--max-per-day", "5", "--conflicts-csv", cf_csv, "--storage-csv", st_csv, "-o", out)
    check("run succeeds", code == 0 and "Matched 26 of 26" in log, log[-800:])
    check("arrival column auto-detected", "Arrival     : Delivery Date" in log)
    check("zone from the drawing layer, named by its text", "LAYDOWN NORTH" in log and "(layer 'LAYDOWN')" in log)
    check("conflicts summarised", "5 potential move-in conflict(s)" in log
          and all(k in log for k in ("overlap", "crowding", "access", "capacity")), log[-1500:])
    P = payload(out)
    C = {c["k"]: c for c in P["conflicts"]}
    check("payload holds one conflict of each kind", set(C) == {"overlap", "crowding", "access", "capacity"}
          and len(P["conflicts"]) == 5, str([c["k"] for c in P["conflicts"]]))
    check("overlap: the two tools on one footprint, dated by the later one",
          C.get("overlap", {}).get("ids") == ["OV-A", "OV-B"] and C["overlap"]["day"] == "2026-09-12"
          and C["overlap"]["s"] == "high", str(C.get("overlap")))
    check("boxed in: the late tool and its four earlier neighbours",
          C.get("access", {}).get("ids", [""])[0] == "BX-C" and len(C["access"]["ids"]) == 5, str(C.get("access")))
    crowd = [c for c in P["conflicts"] if c["k"] == "crowding"]
    check("crowding: the four neighbours form one cluster",
          any(set(c["ids"]) == {"CR-1", "CR-2", "CR-3", "CR-4"} for c in crowd), str(crowd))
    check("capacity: the six same-day move-ins", len(C.get("capacity", {}).get("ids", [])) == 6
          and C["capacity"]["x"] is None, str(C.get("capacity")))
    check("tools carry their conflict indexes", all("cf" in t for t in P["tools"] if t["id"] in ("OV-A", "BX-C", "CAP-3"))
          and not any("cf" in t for t in P["tools"] if t["id"] in ("LATE", "ST-1")))
    Z = P["storage"]["zones"]
    check("two zones (the open polyline is not one)", len(Z) == 2 and {z["n"] for z in Z} == {"Laydown South", "LAYDOWN NORTH"}, str(Z))
    south = [z for z in Z if z["n"] == "Laydown South"][0]
    check("explicit capacity honoured", south["cap"] == 2 and south["cols"] * south["rows"] > 2, str(south))
    T = {t["id"]: t for t in P["tools"]}
    check("delivered tools get a zone and a slot", all(T[f"ST-{i}"].get("sz") == Z.index([z for z in Z if z["n"] == "LAYDOWN NORTH"][0])
          and T[f"ST-{i}"]["ad"] == "2026-10-27" for i in range(1, 9)), str({k: (v.get("sz"), v.get("sl")) for k, v in T.items() if k.startswith("ST")}))
    check("slots inside the zone and distinct", len({(T[f"ST-{i}"]["sx"], T[f"ST-{i}"]["sy"]) for i in range(1, 9)}) == 8
          and all(-1620 <= T[f"ST-{i}"]["sy"] and T[f"ST-{i}"]["sy"] + T[f"ST-{i}"]["h"] <= -1500 + 1e-6
                  and 2600 <= T[f"ST-{i}"]["sx"] and T[f"ST-{i}"]["sx"] + T[f"ST-{i}"]["w"] <= 2840 + 1e-6 for i in range(1, 9)))
    check("late delivery needs no storage", "ad" not in T["LATE"] and "sz" not in T["LATE"])
    check("timeline starts at the first arrival", P["days"][0] == "2026-09-06", P["days"][0])
    check("storage summary printed", "10 tool(s) need storage, 10 placed" in log and "peak 8 on 2026-10-27" in log, log[-1200:])
    check("CSV exports", os.path.exists(cf_csv) and len(open(cf_csv, encoding="utf-8").read().splitlines()) == 6
          and os.path.exists(st_csv) and len(open(st_csv, encoding="utf-8").read().splitlines()) == 11)

    code, log = run(*base, "--storage", "Laydown South=100,50,400,150:2", "-o", out)
    check("full storage reported as conflicts", "6 WITHOUT a slot" in log
          and sum(1 for c in payload(out)["conflicts"] if c["k"] == "storage") == 6, log[-1200:])
    code, log = run(*base, "--no-conflicts", "--conflict-gap", "0", "-o", out)
    check("--no-conflicts", code == 0 and not payload(out)["conflicts"] and "storage" not in log.lower())
    code, log = run(*base, "--conflict-gap", "0", "--conflict-window", "0", "-o", out)
    check("--conflict-gap 0 disables crowding", not any(c["k"] == "crowding" for c in payload(out)["conflicts"]))
    code, log = run(*base, "--storage", "100,50,400,150", "--storage-lead", "3", "-o", out)
    P = payload(out)
    check("unnamed zone + lead time", P["storage"]["zones"][0]["n"] == "Storage 1"
          and "Arrival     : rows without a readable arrival date arrive 3 day(s) early" in log, log[-800:])
    code, log = run(*base, "--storage-lead", "5", "-o", out)
    check("storage options without zones warn", "need storage zones" in log)
    code, log = run(*base, "--storage", "Bad=1,2,3", "-o", out)
    check("bad --storage rejected", code != 0 and "not understood" in log)
    code, log = run(*base, "--storage-layer", "NOPE", "-o", out)
    check("unknown --storage-layer warned", "matches no layer" in log)
    code, log = run(*base, "--inspect")
    check("--inspect hints at storage layers", "look like storage" in log and "'LAYDOWN'" in log)

    run(*base, "--storage-layer", "LAYDOWN", "--storage", "Laydown South=100,50,400,150:2", "--max-per-day", "5", "-o", out)
    pw = browser()
    if not pw:
        return
    with pw() as p:
        b, pg, errs = open_page(p, out)
        info = pg.evaluate("""(() => ({ cf: CF.length, markers: document.querySelectorAll('#cfG g.cfm').length,
            stored: STORED.length, zones: document.querySelectorAll('#zoneG rect.zone').length,
            cfTools: document.querySelectorAll('g.tool.cf').length, pill: document.getElementById('cfPill').textContent,
            tabs: [...document.querySelectorAll('.tabs button')].filter(b => b.style.display !== 'none').length }))()""")
        check("viewer: markers, zones, outlined tools, pill, tabs", info["cf"] == 5 and info["markers"] == 4
              and info["stored"] == 10 and info["zones"] == 2 and info["cfTools"] == 17
              and info["pill"].startswith("⚠ 5 conflicts") and info["tabs"] == 5, str(info))
        state = lambda: pg.evaluate("""(() => ({ st: document.querySelectorAll('#storeG g.stored.in').length,
            mk: document.querySelectorAll('#cfG g.cfm.in').length, extra: document.getElementById('statExtra').textContent }))()""")
        scrub(pg, 0)
        s0 = state()
        check("day 0: nothing in storage, no markers", s0["st"] == 0 and s0["mk"] == 0, str(s0))
        scrub(pg, "DI.get('2026-10-29')")
        s1 = state()
        check("while the deliveries wait: 8 crates in storage, all markers up",
              s1["st"] == 8 and s1["mk"] == 4 and s1["extra"].startswith("8 waiting in storage"), str(s1))
        scrub(pg, "sc.max")
        s2 = state()
        check("last day: storage empty again", s2["st"] == 0 and s2["mk"] == 4, str(s2))
        scrub(pg, "DI.get('2026-10-29')")
        check("scrubbing back re-fills storage", state()["st"] == 8)
        pg.click("#cfPill"); pg.wait_for_timeout(100)
        check("pill opens the Conflicts tab", pg.evaluate("S.tab") == "cf"
              and pg.evaluate("document.querySelectorAll('#toolList .cf-row').length") == 5)
        vb0 = pg.evaluate("document.getElementById('mapSvg').getAttribute('viewBox')")
        pg.click("#toolList .cf-row"); pg.wait_for_timeout(300)
        vb1 = pg.evaluate("document.getElementById('mapSvg').getAttribute('viewBox')")
        check("clicking a conflict zooms to it", float(vb1.split()[2]) < float(vb0.split()[2]) / 3)
        pg.click("#tabSt"); pg.wait_for_timeout(100)
        heads = pg.evaluate("[...document.querySelectorAll('#toolList .tl-day')].map(e => e.textContent)")
        check("Storage tab lists zones with occupancy", any(h.startswith("LAYDOWN NORTH · 8 / 30") for h in heads), str(heads))
        tick = pg.evaluate("(() => { const t0 = performance.now(); for (let i = 1; i <= 60; i++) setDay(i); "
                           "return (performance.now() - t0) / 60; })()")
        check("day stepping still cheap with events", tick < 25, f"{tick:.1f} ms")
        check("no JS errors", not errs, "; ".join(errs[:2]))
        b.close()


def main():
    if not os.path.exists(SCRIPT):
        sys.exit(f"cannot find {SCRIPT}")
    with tempfile.TemporaryDirectory(prefix="fabsim-tests-") as d:
        print(f"fixtures in {d}")
        want = lambda n: not ONLY or n in ONLY
        if want("small"):
            scenario_small(d, make_small(d))
        if want("blocktext"):
            make_blocktext(d)
            scenario_blocktext(d)
        if want("prefer"):
            make_prefer(d)
            scenario_prefer(d)
        if want("aisle"):
            make_aisle(d)
            scenario_aisle(d)
        if want("heavy"):
            make_heavy(d)
            scenario_heavy(d)
        if want("xref"):
            make_xref_detail(d)
            scenario_xref_detail(d)
        if want("conflicts"):
            make_conflicts(d)
            scenario_conflicts(d)
    print(f"\n{PASSES} passed, {len(FAILS)} failed" + (": " + ", ".join(FAILS) if FAILS else ""))
    sys.exit(1 if FAILS else 0)


if __name__ == "__main__":
    main()
