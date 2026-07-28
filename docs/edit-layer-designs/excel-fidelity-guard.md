
# Excel Fidelity Guard + Gap Strategy — Implementation Design

> **ERRATUM (2026-07-28) — finding #2 below is WRONG and was reverted.**
> `keep_vba=True` on a plain `.xlsx` does **not** have "zero side effects". openpyxl sets
> `wb.vba_archive` unconditionally under that flag, and `vba_archive` being truthy alone
> (a) flips `Workbook.mime_type` to the macro-workbook content type
> (`workbook/workbook.py:360-370`) and (b) appends a `vbaProject` relationship pointing at
> `vbaProject.bin` (`workbook/_writer.py:165-168`) that `_merge_vba` then has no bytes to
> satisfy. The result is an `.xlsx` with a macro content type and a dangling relationship,
> which real Excel refuses to open (`-50 参数错误`). The probes behind finding #2 only checked
> that no *new parts* appeared; they never checked `[Content_Types].xml` or the rels.
> `session.py::_load()` now gates `keep_vba` on `"xl/vbaProject.bin" in <source zip>`, and
> `save_to()` asserts extension/content-type consistency before handing a file over.
> See [bench/BUGS.md OA-1](../../../bench/BUGS.md) and
> `tests/unit/test_bugs_oa_fixes.py`. Everything else in this document stands — including
> `rich_text=True` (finding #1), which was independently exonerated.

All claims below are either (1) inspection of the installed `openpyxl==3.1.5` source at
`/opt/miniconda3/lib/python3.13/site-packages/openpyxl`, or (2) reproduced empirically with
throwaway probes in `/private/tmp/.../scratchpad/excel_probe/` (kept on disk, not committed).
Where something could not be fully verified (pivot cache-record fidelity), that is stated
explicitly rather than assumed.

## 0. Empirical findings that drive the design

**Mechanism of loss.** `ExcelWriter.write_data()` (`writer/excel.py`) does not copy the source
zip forward — it builds an entirely fresh archive from the `Workbook` object graph
(`_write_worksheets`, `_write_charts`, `_write_images`, tables, pivots, external links, styles,
`workbook.xml`). The *only* passthrough is `_merge_vba()`, which — only when
`workbook.vba_archive` is populated (`keep_vba=True` at load) — copies raw bytes for parts
matching a narrow regex: `xl/vba`, `xl/drawings/*vmlDrawing\d.vml`, `xl/ctrlProps`, `xl/activeX`,
`customUI`, `xl/media/*.emf`. Everything else present in the original zip but not represented in
openpyxl's object model is **silently discarded**, whether or not it was correctly OPC-wired
(verified: `xl/threadedComments/*`, `xl/persons/*`, `xl/slicers/*`, `customXml/*` all vanish on a
single load→save cycle with no error — `probe3_zip_surgery_fixture.py`).

**openpyxl's own feature surface round-trips cleanly.** Built one workbook using every modeled
feature (Table, legacy comment, bar chart, image, data validation, conditional formatting,
defined name, freeze panes, merged cells, hyperlink) and round-tripped it: zero parts lost, one
cosmetic size change (whitespace in `drawing1.xml`) (`probe2_openpyxl_own_features.py`).

**This means the currently-shipped heuristic is a false-positive generator.** `session.py`'s
`_detect_lossy_parts` (commit `97292ba`) warns on *any* `xl/charts/`/`xl/media/` presence by
counting zip-name substrings. It also fires for charts/images this very tool created and saved
in a *previous* session (proven safe by the round-trip above) — confirmed by the repo's own test
`test_open_xlsx_with_charts_warns`, which builds its "at risk" fixture with
`ExcelAdapter.create_chart` (openpyxl-authored, therefore actually safe) and asserts a warning.
The mechanism needs to move from "does a risky-looking name appear" to "did a named part actually
disappear on a real round trip" — that is exactly what the part-inventory guard below does, and
it eliminates this false positive as a side effect.

**Catastrophic all-or-nothing failure mode.** `reader/drawings.py::find_images()`:
```python
try:
    drawing = SpreadsheetDrawing.from_tree(tree)
except TypeError:
    warn("DrawingML support is incomplete and limited to charts and images only. "
         "Shapes and drawings will be lost.")
    return [], []
```
If a drawing XML part fails strict deserialization (schema variant, e.g. SmartArt/OLE/ink content
mixed with charts on the same drawing), **both charts and images anchored on that drawing** are
dropped together, silently, via a Python `warnings.warn()` that nothing in this codebase captures
today. Separately — and unconditionally, no TypeError needed — the *writer* rebuilds
`ws._drawing` from only `ws._charts` + `ws._images` (`writer/excel.py:194-196`), so any
freestanding shape/textbox/group/ink/OLE content that the reader *did* parse successfully is
still thrown away, because `find_images()` never retains shape anchors on the object model in the
first place (`probe4_shape_kills_whole_drawing.py` reproduced the parse succeeding; the shape
still isn't retained anywhere Python-visible).

**Pivot tables are genuinely partial, not "safe" or "lost."** Unlike slicers/threaded
comments/sparklines/customXml (zero support — confirmed no reader/writer code path at all),
`openpyxl.pivot.{cache,table,record}` + `workbook/_writer.py::write_pivots` DO read and re-emit
pivot table *definitions* and cache *references*. Cache-record fidelity/refresh behavior was not
verified (building a schema-valid pivot fixture by hand is high-effort and low-confidence without
a real Excel-authored source) — treat as **at risk, unresolved**, not as either safe or doomed;
resolve with a real fixture (test plan §4).

**Two free, currently-missing fidelity wins, found while probing:**
1. `rich_text` defaults to `False` in `load_workbook`. A cell with intra-cell mixed-run
   formatting (`CellRichText`/`TextBlock`, e.g. one bold word in an otherwise plain cell) is
   *silently flattened to a plain string* on load, and the formatting is gone for good the moment
   that happens — not a display quirk, a real, irreversible loss, and it happens even though the
   sheet XML part itself is still present (so the part-inventory guard alone cannot catch it).
   `rich_text=True` preserves it losslessly through a full load→save→reload cycle, with zero
   observed side effects on ordinary plain-string/number cells (`probe6_rich_text.py`).
2. **[REVERTED — see the erratum at the top of this file; this finding is false.]**
   `keep_vba` is currently gated on the `.xlsm` extension in `session.py::_load()`. But
   `keep_vba=True` on a *plain* `.xlsx` with no macros has **zero side effects** (no spurious VBA
   parts appear in the output) while still rescuing `ctrlProps`/`activeX`/`vmlDrawing`/`customUI`
   parts if the file happens to have form controls (checkboxes, buttons) without any macro
   (`probe5_xlsm_keepvba.py`, `probe7` follow-up). Cost: `wb.vba_archive` holds a full extra copy
   of the source zip in memory for the life of the session — negligible at this tool's scale.

**AppleScript native-automation exploration (for fixture-building, §4d).** Two independent
attempts at `make new chart object … with properties {chart type: bar clustered, …}` and a
save-as-xlsx variant both hit AppleScript syntax errors (-2741) with no working syntax found in
the time available — `sdef` (the dictionary dumper) requires full Xcode, not just Command Line
Tools, which is what's installed here, so I could not browse the exact enumerator names. This
independently reproduces the flavor of the user's own `-1728` chart-creation failure: Excel's
AppleScript dictionary enum names are non-obvious and this is a recurring speed bump, not a one-off.
Simple value writes (`set value of range "A1" to "Cat"`) worked without complaint in the same
script, consistent with the user's own experience that cell writes are easy. Recommendation:
don't invest further in AppleScript object *creation* syntax right now (see §2 Phase 3); use
zip-surgery for fidelity-guard test fixtures instead (§4), which is fully validated below.

---

## (a) Part-inventory guard — design

### Where it lives
`src/office_agent/core/session.py`, generalizing/replacing `EditSession._detect_lossy_parts`
(lines 73–100 today).

### Data model
```python
import io, zipfile
from dataclasses import dataclass

@dataclass(frozen=True)
class FidelityLossEntry:
    path: str      # zip member name, e.g. "xl/threadedComments/threadedComment1.xml"
    size: int      # bytes, from the ORIGINAL file
    category: str  # human label, e.g. "threaded comment"

class FidelityLossError(PermissionError):
    """Saving would silently drop OOXML parts openpyxl's writer cannot preserve.
    Subclasses PermissionError so it flows through ToolRegistry.dispatch's existing
    caught-exception tuple with ZERO changes to registry.py — same precedent as the
    overwrite guard in save_to()."""

_FIDELITY_CATEGORIES = [  # longest/most-specific prefix first
    ("xl/threadedComments/", "threaded comment"),
    ("xl/persons/",          "comment-author registry"),
    ("xl/slicerCaches/",     "slicer cache"),
    ("xl/slicers/",          "slicer"),
    ("xl/pivotCache/",       "pivot cache"),
    ("xl/pivotTables/",      "pivot table"),
    ("xl/richData/",         "rich data type (stocks/geography)"),
    ("xl/metadata.xml",      "rich-value/dynamic-array metadata"),
    ("customXml/",           "custom XML part"),
    ("xl/ctrlProps/",        "form control"),
    ("xl/activeX/",          "ActiveX object"),
    ("customUI/",            "ribbon customization"),
    ("xl/media/",            "embedded image/media"),
    ("xl/charts/",           "chart"),
    ("xl/drawings/",         "drawing (shape/textbox/chart anchor)"),
]

def _categorize(name: str) -> str:
    for prefix, label in _FIDELITY_CATEGORIES:
        if name.startswith(prefix):
            return label
    return "unrecognized part"

def _zip_inventory(source) -> dict:
    """source: Path or bytes. Returns {member_name: size}. Cheap — reads only
    the central directory, no decompression."""
    fh = io.BytesIO(source) if isinstance(source, (bytes, bytearray)) else source
    with zipfile.ZipFile(fh) as zf:
        return {i.filename: i.file_size for i in zf.infolist()}
```

### Wiring into `EditSession`
```python
class EditSession:
    def __init__(self, file_path: str):
        ...
        shutil.copy2(original, self.working_path)
        self._baseline_inventory = (
            _zip_inventory(self.working_path) if self.doc_type == "excel" else {}
        )
        self.doc = self._load()
        self.fidelity_report = self._compute_fidelity_report()   # replaces preservation_warnings
        ...

    def _load(self):
        if self.doc_type == "word":
            return WordDocument(str(self.working_path))
        return load_workbook(
            str(self.working_path),
            keep_vba=True,        # was: gated on .xlsm — now always on, see §0 finding #2
            rich_text=True,       # NEW — see §0 finding #1
        )

    def _compute_fidelity_report(self) -> list[FidelityLossEntry]:
        """Computed ONCE at open. Loss is fixed at load time — openpyxl's object
        model either captured a part or it never did; further edits operate
        within that already-loaded model and cannot un-lose anything. One
        trial serialize here is the same cost class as the baseline undo
        snapshot SnapshotManager already takes right after open."""
        if self.doc_type != "excel":
            return []
        buf = io.BytesIO()
        self.doc.save(buf)
        current = _zip_inventory(buf.getvalue())
        return [
            FidelityLossEntry(name, size, _categorize(name))
            for name, size in self._baseline_inventory.items()
            if name not in current
        ]

    def fidelity_warnings(self) -> list[str]:
        """Human-readable strings for open_document's `warnings` field."""
        if not self.fidelity_report:
            return []
        by_cat: dict[str, list[FidelityLossEntry]] = {}
        for e in self.fidelity_report:
            by_cat.setdefault(e.category, []).append(e)
        parts = ", ".join(f"{len(v)} {k}(s)" for k, v in by_cat.items())
        return [
            f"⚠️ This workbook contains part(s) the current editing engine CANNOT "
            f"preserve — {parts} will be LOST if you save: "
            f"{', '.join(e.path for e in self.fidelity_report)}. Inform the user "
            "before saving. For simple cell/formula/style edits elsewhere in this "
            "sheet, prefer the excel_native_* tools to avoid losing these; "
            "otherwise save_document will require accept_fidelity_loss=True after "
            "explicit user consent."
        ]

    def save_to(self, dest=None, overwrite=False, accept_fidelity_loss=False) -> Path:
        target = Path(dest).expanduser().resolve() if dest else self.original_path
        # ... existing overwrite checks unchanged ...
        self.flush()  # existing call — writes doc to working_path
        if self.doc_type == "excel" and not accept_fidelity_loss:
            current = _zip_inventory(self.working_path)
            lost = [
                FidelityLossEntry(name, size, _categorize(name))
                for name, size in self._baseline_inventory.items()
                if name not in current
            ]
            if lost:
                raise FidelityLossError(_format_fidelity_error(lost, self.original_path))
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(self.working_path, target)
        self.written_paths.add(str(target))
        return target
```

`_format_fidelity_error` produces a message like:
```
Saving would silently DROP 4 part(s) that this editing engine (openpyxl) cannot
preserve because it fully re-serializes the workbook from its own object model:
  - threaded comment: xl/threadedComments/threadedComment1.xml (372 bytes)
  - comment-author registry: xl/persons/person.xml (253 bytes)
  - slicer: xl/slicers/slicer1.xml (206 bytes)
  - custom XML part: customXml/item1.xml (91 bytes)
The original file at /Users/.../report.xlsx is untouched.
Options: (1) show the user exactly what would be lost and, only after explicit
confirmation, retry save_document with accept_fidelity_loss=True; (2) for simple
cell/formula/style edits, use the excel_native_* tools (native Excel automation)
instead — they preserve everything; (3) save to a new path and continue editing
the original directly in Excel for these specific features.
```

### Why diff by *name presence*, not size
Even a fully lossless round trip changes byte *sizes* of nearly every part (whitespace,
attribute-order, re-encoded shared strings) — proven in `probe2` (0 parts lost, but
`drawing1.xml`'s size changed). A size-based trigger would false-positive on **every single
save**, training users/agents to ignore it. The guard blocks only on a part *name disappearing
entirely*. Size-changed-but-present parts are still recorded to `session.audit()` as a soft,
non-blocking signal (paper trail only) — this is a deliberate, empirically-justified choice, not
an oversight.

### Scope boundary
Word is untouched — python-docx edits the lxml tree **in place** on the original OPC package (per
the given verified facts), so this whole class of "wholesale reserialization eats parts it
doesn't understand" risk does not apply there. `_compute_fidelity_report`/`save_to`'s check are
both no-ops (`doc_type != "excel"` → `[]`/skip).

### Interaction with the existing render-verify loop (why this guard is *necessary*, not redundant)
`Renderer.render()` calls `self.session.flush()` before every AppleScript export
(`render/renderer.py:31`) — meaning every rendered screenshot, including the very first one taken
right after `open_document`, already reflects the post-load, already-lossy re-serialized working
copy. The verifier's diff-based approach (per the given facts: "independent stateless multimodal
verifier judges *changed* pages") is blind to loss that happened at **open** time, before any
edit-diff baseline existed — there is no "pristine original" render to diff against. The part-
inventory guard is the only mechanism that can catch this, because it compares against the
*original file on disk*, not against a previous render.

---

## (b) Alternatives when the guard trips

### Option 1 — Raw-XML editing of the unpacked package (unpack → patch `sheetN.xml`/`styles.xml` → repack)
- **What it buys**: byte-identical preservation of every untouched part, since nothing gets
  reserialized wholesale — this is the architecturally "right" model (matches Anthropic's own
  docx/xlsx skills: unpack → edit XML → repack + validate) and the PPTArena/PPTPilot hybrid
  pattern the user cited.
- **What it costs**: for the "simple" cases the user names (values, styles) it is genuinely
  tractable but non-trivial:
  - **Values**: writing a *new* string requires either appending to `sharedStrings.xml` (and
    correctly maintaining `count`/`uniqueCount`, optionally deduping against existing entries) or
    switching that one cell to `t="inlineStr"` (simpler, always safe, mild non-canonical bloat —
    Excel reads it fine).
  - **Styles**: a cell's `s` attribute indexes into `cellXfs` in `xl/styles.xml`; adding a new
    combination of font/fill/border/numFmt means either finding a matching existing `cellXfs`
    entry or appending one and cross-referencing `fonts`/`fills`/`borders`/`numFmts` — this is
    reimplementing a meaningful slice of what openpyxl already does internally, just by hand.
  - **Charts/pivots/slicers/comments**: raw-XML editing does **not** rescue any of these — they'd
    still need their own from-scratch XML generation (drawing anchors, rels, content-type
    overrides), which is most of the actual engineering effort and risk in this option, for
    exactly the features that matter most.
- **Risk**: a hand-rolled OOXML writer is a brand-new, large surface for silent corruption bugs —
  this cuts against the project's core thesis (trust the file structure; verify with real
  rendering) by introducing a second, parallel, harder-to-verify write path.

### Option 2 — Route to native Microsoft Excel automation (AppleScript)
- **What it buys**: this is the **most faithful possible answer**, because Excel is authoring its
  own file — there is no reimplementation risk at all for the parts that survive, since Excel
  wrote them in the first place and Excel's own save routine already understands its own
  pivot/slicer/threaded-comment format perfectly.
- **What it costs**:
  - Reuses already-validated zero-dialog infrastructure (`render/applescript.py`): sandbox
    container paths, `display alerts none`/`false`, `_AUTOMATION_LOCK` serialization, dual
    AppleScript+subprocess timeouts, the `RenderError`/`RenderTimeout`/`AutomationDenied`/
    `DocumentCorrupted` classification taxonomy. Extending it with write functions is
    architecturally "more of the same," not a new paradigm.
  - Confirmed low-friction for exactly the ops named as "easy": `set value of range "B2" of
    worksheet "Sheet1" of workbook 1 to ...` worked without issue in probing.
  - Confirmed high-friction for anything beyond simple values (chart/pivot/slicer *creation*) —
    two independent probe attempts hit AppleScript syntax errors; treat as **not yet solved**,
    see Phase 3.
  - **State-sync hazard** (the main real risk): AppleScript acts on a file via the real app, not
    on `session.doc` in memory. Every native write must be followed by
    `session.reload_from_bytes(result_bytes)` so the live openpyxl object never diverges from
    disk — otherwise the next python-lib edit's `flush()`/snapshot would silently clobber the
    native change. This must be enforced inside the tool wrapper, not left to agent discipline.
  - **Latency**: subprocess + UI automation (~1–3s per call after warm-up) vs. microseconds for
    an in-process openpyxl mutation — acceptable for an occasional escape hatch, not for bulk
    edits.
  - **Isolation**: write via a dedicated scratch copy inside the sandbox
    (`session.session_dir / "native_scratch.xlsx"`), not `session.working_path` directly, so
    Excel's own file lock never collides with the session's assumptions about that path.

### Option 3 — Refuse, with save-as-copy guidance
- Zero engineering cost, zero risk, zero new capability. This is the *default fallback* for
  anything neither Option 1 nor Option 2 covers yet (chart/pivot/slicer creation).

### Recommendation
**Default posture: block-with-guidance (Option 3) for everything, upgraded to native-Excel-
automation (Option 2) for a narrow, explicit set of "simple write" operations** (cell values,
formulas, basic font/fill/number-format styling) **when the guard trips because of unrelated
pre-existing at-risk content elsewhere in the workbook**. Raw-XML editing (Option 1) is
**explicitly not recommended for the initial build** — it is the highest-effort option, it only
ever covers the "simple" half of the problem (never rescues charts/pivots/slicers, which is most
of what's actually at risk), and it adds a second write path whose correctness this project has
no existing way to verify as rigorously as it verifies openpyxl's. Revisit Option 1 only if
real usage shows Option 2's latency/state-sync overhead is a genuine blocker *and* there is
demonstrated demand — not preemptively.

No silent auto-fallback: `save_document` never *itself* switches engines. It blocks and names the
lost parts; the agent/user decides whether to accept the loss (`accept_fidelity_loss=True`,
mirroring the existing `overwrite=True` consent pattern) or redo the specific edits via the
native tool and retry. Predictability over cleverness, consistent with the project's existing
explicit-gate philosophy.

---

## (c) `keep_vba` / `keep_links` / rich-read options

- **`keep_vba`**: change from `.xlsm`-gated to **always `True`** for Excel loads. Verified zero
  side effects on plain `.xlsx` (no spurious VBA parts appear in output); still correctly
  preserves `xl/vbaProject.bin` + associated `ctrlProps`/`activeX`/`vmlDrawing`/`customUI`/EMF
  media whenever they exist, macro or not. Minor cost: `wb.vba_archive` holds a full copy of the
  source zip bytes in RAM for the session's lifetime — negligible at this tool's scale.
- **`keep_links`**: leave at its openpyxl default (`True`) but pass it **explicitly** in
  `_load()` rather than relying on the implicit default, purely for auditability/documentation —
  external-workbook-reference formulas are one of the better-supported features (both
  `reader/workbook.py` and `writer/excel.py::_write_external_links` model them), so no behavior
  change is needed, just make the choice visible in the code.
- **`rich_text`**: change from default `False` to **`True`**. This is the one real, silent,
  currently-shipping fidelity bug found during probing (§0 finding #1) — turn it on
  unconditionally; verified no downside for ordinary plain cells.
- **Coupled fix required**: `ExcelAdapter.read_cells` (`adapters/excel_adapter.py:18-38`) must
  stringify `CellRichText` values before they reach the tool's returned dict — `cell.value if not
  isinstance(cell.value, CellRichText) else str(cell.value)`. Verified `str(CellRichText(...))`
  produces the clean concatenated text (not a verbose repr). Without this, a rich-text cell's
  value flows into `agent/loop.py`'s `json.dumps(result, default=str)` and gets a *usable* but
  undocumented fallback; making the ExcelAdapter contract explicit is cleaner than depending on a
  downstream serialization detail three layers away.

---

## (d) Test plan, including fixture-building strategy

### Fixture-building strategy (answering the AppleScript-chart-creation question directly)
AppleScript native chart/pivot/slicer *creation* is not a reliable path to build test fixtures
right now (§0) — `sdef` needs full Xcode (unavailable here; Command Line Tools only), and two
independent syntax attempts failed with `-2741`/mirrored the user's own `-1728`. **Recommended
approach, validated and used throughout this design's probing:**
1. **Zip surgery on an openpyxl base** for anything that doesn't need semantically-correct
   internals to prove the guard's detection logic — threaded comments, slicers, customXml,
   sparkline-style extLst content. The guard only needs to know a named part *disappeared*; it
   doesn't validate the part's own internal correctness, so a structurally-plausible-but-not-
   Excel-perfect injected part is sufficient and fully reproducible in pytest (no binary fixtures
   to check in). This is exactly what `probe3_zip_surgery_fixture.py` does — turn it into a
   reusable pytest helper, e.g. `tests/unit/_xlsx_fixtures.py::inject_foreign_part(path, kind)`.
2. **One hand-authored fixture, checked in once**, for pivot tables specifically — this is the
   one feature where "does it actually work when Excel refreshes it" cannot be faked
   structurally with confidence. Recommend: create a small workbook with one pivot table in real
   Excel, save as `tests/fixtures/excel_fidelity/pivot_table_manual.xlsx`, document its
   provenance in a comment at the top of the test that uses it. This resolves the "unverified"
   status flagged in §0 without further blind AppleScript spelunking.

### Unit tests (`tests/unit/test_excel_fidelity_guard.py`, new file)
1. `_zip_inventory` on Path vs bytes — trivial parity check.
2. **No false positive**: build the full openpyxl-native feature surface (table, comment, chart,
   image, data validation, conditional formatting, defined name, freeze panes, merged cells,
   hyperlink) → `open_document` → `session.fidelity_report == []` → `save_document` succeeds with
   `accept_fidelity_loss` omitted (default `False`).
3. **True positive, correct naming**: inject threaded-comment/persons/slicer/customXml parts via
   zip surgery → `open_document`'s `warnings` names all four with correct categories →
   `save_document` (no `accept_fidelity_loss`) raises, envelope shows
   `{"success": False, "error_type": "FidelityLossError", "error": "...xl/threadedComments/..."}`
   → retry with `accept_fidelity_loss=True` succeeds, and the saved file's inventory confirms
   those parts really are gone (the tool must not claim a loss it doesn't actually incur, nor
   silently keep something it said it would drop).
4. **`.xlsm` + fake `vbaProject.bin`**: guard does not flag `xl/vbaProject.bin` as lost; a second
   case with `keep_vba` forced off proves the guard *would* have caught it (regression pin for
   finding #2).
5. **Rich text regression**: mixed-run cell → open → make an unrelated edit elsewhere → save →
   reopen → formatting intact; `excel_read_cells` returns a plain JSON-safe string, not a
   `CellRichText` object, for that cell.
6. **Size-change is not a false trigger**: assert a lossless round trip with a chart (size of
   `drawing1.xml` changes) does *not* appear in `fidelity_report` and does not block save.

### Integration
7. Shape-in-drawing fixture (`probe4`-style, generalized): confirms the guard's *name-presence*
   diff catches the outcome regardless of whether the specific failure mode is "shape silently
   unretained" or "TypeError kills the whole drawing" — the design deliberately doesn't need to
   distinguish these at the guard level, only detect the resulting part loss.
8. Pivot-table manual fixture (see above): open → record whatever `fidelity_report` says today
   (may turn out empty if openpyxl's partial support is better than expected, or non-empty) →
   **this is a "go find out" test, not a "prove a known result" test** — whichever way it comes
   out, update the friendly-category confidence level in code comments and README accordingly.

### Performance
9. Benchmark the open-time trial-serialize on a large-ish sheet (e.g. 20k rows × 20 cols) —
   Excel-scale latency hasn't been validated the way Word's 60-page case was (per commit
   history). The trial serialize is one *additional* instance of a cost the system already pays
   on every mutating call (via `SnapshotManager`), so it should be proportionally small, but this
   needs a real number before shipping, not an assumption.

### Native-automation escape hatch (Phase 2, `live`-gated like existing `OFFICE_AGENT_LIVE_TEST=1`
tests)
10. Inject an at-risk part into a workbook → `excel_native_write_cells` updates a value via real
    Excel → reopen and confirm (a) the value changed, (b) the at-risk part is *still present*
    (Excel, not openpyxl, performed the save), (c) a subsequent `excel_write_cells` (python-lib)
    call sees the native-written value too (proves the `reload_from_bytes` resync actually
    happened and the two engines never diverged).

### Required update to existing tests
`tests/unit/test_review_fixes.py::test_open_xlsx_with_charts_warns` currently asserts a warning
fires for an *openpyxl-authored* chart — which round-trips safely (§0) — so it encodes the old,
inaccurate heuristic. **This must be rewritten**, not left passing by accident: either (a) change
its fixture to use a zip-surgery-injected genuinely-at-risk part instead of
`ExcelAdapter.create_chart`, keeping the same test name/intent, or (b) split it into
"safe-chart-produces-no-warning" (new, using the current fixture) + "genuinely-foreign-part-
produces-warning" (new, using zip surgery). This is called out explicitly as an intentional
correction of a shipped false positive, not a regression.

---

## Phased rollout

- **Phase 0 (do first, ~1 hr, no API surface change)**: flip `rich_text=True` and un-gate
  `keep_vba=True`; add the `CellRichText`→`str` normalization in `ExcelAdapter.read_cells`.
  Strictly additive, zero risk, ships independently of everything else.
- **Phase 1 (core ask of this task, ~1–2 days)**: replace `_detect_lossy_parts` with the
  part-inventory guard (`fidelity_report` at open, blocking check in `save_to`); add
  `accept_fidelity_loss` to `SaveDocumentInput`/`save_document`; add a cheap read-only
  `excel_fidelity_report(doc_id)` tool exposing `session.fidelity_report` on demand (agents
  shouldn't have to attempt-and-fail a save just to find out); rewrite the affected existing
  test; add the new fixture-building tests. This is the highest-leverage change: it converts an
  already-shipping *silent* loss into a *loud, informed-consent* one, with a more accurate
  detector than what's live today.
- **Phase 2 (~2–3 days incl. live testing)**: `render/applescript.py` gains
  `write_cells_via_excel(...)`; new `tools/excel_native_tools.py` with
  `excel_native_write_cells` (values/formulas) and a basic-style variant; wraps the
  scratch-copy + `reload_from_bytes` resync protocol; the `FidelityLossError` message points the
  agent at these tools by name.
- **Phase 3 (not scheduled — explicitly not recommended right now)**: raw-XML sheet/style
  editing subsystem, and/or AppleScript-driven chart/pivot/slicer *creation*. Both require a
  dedicated spike first — for AppleScript, ideally with full Xcode's `sdef` or interactive Script
  Editor dictionary browsing (unavailable in this environment) to nail exact enumerator names
  before committing engineering time; for raw-XML, only revisit if Phase 2's latency/state-sync
  overhead proves to be a real, measured blocker with demonstrated user demand.

---

## Concrete file-level change list

- `src/office_agent/core/session.py`: add `FidelityLossEntry`, `FidelityLossError`,
  `_zip_inventory`, `_categorize`, `_format_fidelity_error`; replace `_detect_lossy_parts` with
  `_compute_fidelity_report`/`fidelity_warnings`; edit `_load()` (`keep_vba=True` unconditional,
  `rich_text=True`, explicit `keep_links=True`); edit `save_to()` signature + blocking check.
- `src/office_agent/tools/session_tools.py`: `SaveDocumentInput.accept_fidelity_loss: bool =
  False`; `save_document` passes it through; `open_document` uses `session.fidelity_warnings()`
  instead of `session.preservation_warnings`; new `ExcelFidelityReportInput`/
  `excel_fidelity_report` tool.
- `src/office_agent/adapters/excel_adapter.py`: `read_cells` stringifies `CellRichText`.
- `src/office_agent/render/applescript.py`: new `write_cells_via_excel(...)`, reusing `_q`,
  `_AUTOMATION_LOCK`, `_classify`.
- `src/office_agent/tools/excel_native_tools.py` (new file, Phase 2): `excel_native_write_cells`
  and friends, with the scratch-copy + `reload_from_bytes` resync protocol.
- `tests/unit/test_excel_fidelity_guard.py` (new file): tests 1–9 above.
- `tests/unit/_xlsx_fixtures.py` (new file): `inject_foreign_part(path, kind)` zip-surgery
  helper, generalized from `probe3_zip_surgery_fixture.py`.
- `tests/fixtures/excel_fidelity/pivot_table_manual.xlsx` (new, hand-authored once).
- `tests/unit/test_review_fixes.py`: rewrite `test_open_xlsx_with_charts_warns` per §4.
- `README.md`: update the "已知限制" section — replace the current blanket "打开本来就带图表/图片
  的 .xlsx 再保存会丢失" line with the more accurate (and less alarming) statement that
  openpyxl-authored charts/images are safe, only genuinely-foreign parts (pivot/slicer/threaded-
  comments/etc.) are at risk, and the guard now blocks + names them explicitly rather than just
  warning.


## Key decisions
- Replace the shipped chart/media-substring heuristic (_detect_lossy_parts, commit 97292ba) with an actual load->trial-serialize->diff computed once at open. This fixes a proven false positive (it currently warns on openpyxl's own charts, which round-trip safely) and generalizes detection to every at-risk OOXML part, not just charts/media.
- Block (not just warn) at save_document by default when the part-inventory guard detects any lost part; add a new accept_fidelity_loss flag on SaveDocumentInput, parallel to the existing overwrite flag, framed as requiring explicit user consent. Implemented as FidelityLossError(PermissionError) so it needs zero changes to registry.py's exception handling.
- Diff by part-name presence only, never by size. A lossless round trip still changes byte sizes of nearly every part (whitespace/reordering), so a size-based trigger would false-positive on every save; only a part disappearing entirely blocks the save.
- Default fallback when the guard trips is native Microsoft Excel automation (AppleScript) for simple write ops (cell values/formulas/basic styles) -- not raw-XML editing. Raw-XML editing is explicitly deferred: it's the highest-effort option and only ever covers the 'simple' half of the risk (never rescues charts/pivots/slicers).
- Always pass keep_vba=True (not just for .xlsm) and add rich_text=True to load_workbook -- two verified-zero-downside fidelity wins found during probing: keep_vba=True rescues form-control parts (ctrlProps/activeX/vmlDrawing) on plain .xlsx files with no macros; rich_text=True stops silently flattening intra-cell mixed-run formatting, a real bug in the current code.
- Guard computation piggybacks on serialization the system already performs (baseline undo-snapshot at open, mandatory flush() at save) so it introduces no new expensive operation class, only additional cheap zip-inventory scans.
- No silent auto-fallback: save_document never itself switches editing engines. It blocks and names exactly what would be lost; the agent/user decides whether to accept the loss or redo the edits via the native-automation tool.
- The existing test test_open_xlsx_with_charts_warns encodes the old, inaccurate heuristic and must be rewritten as part of this change -- called out explicitly as an intentional correction, not a regression.

## Risks
- Full-workbook trial re-serialize at open time adds latency; Excel-scale performance (unlike Word's validated 60-page case) has not been benchmarked. Needs a benchmark on a large sheet (e.g. 20k rows) before shipping, not an assumption that it's fine.
- The native-Excel-automation escape hatch creates a second live copy of the document (the real Excel app's in-memory/on-disk state) that must be explicitly resynced into session.doc via reload_from_bytes after every native write. A missed resync would let a subsequent python-lib edit silently clobber the native change -- this must be enforced by the tool wrapper, not left to agent discipline.
- AppleScript dictionary syntax for anything beyond simple value/formula writes (chart/pivot/slicer creation) is fragile and non-obvious: two independent probe attempts in this session hit syntax errors (-2741), mirroring the user's own -1728 experience. Do not assume this escape hatch generalizes to complex object creation without a dedicated dictionary-discovery spike (ideally with full Xcode's sdef or interactive Script Editor browsing, unavailable in this environment).
- Pivot table fidelity is only partially characterized: openpyxl's reader/writer do have real code paths for pivot definitions and cache references, but cache-record/refresh fidelity was not empirically verified (a safe, schema-valid fixture could not be fabricated with confidence in the time available). The design flags pivot tables as 'at risk, unresolved' rather than claiming certainty either way, and requires a real manually-authored fixture test to close this out.
- openpyxl's reader silently swallows a whole class of failures via warnings.warn() (e.g., the find_images() TypeError path that drops both charts and images together) rather than raising or logging by default. The part-inventory diff still catches the resulting data loss, but without capturing these warnings during load, the error message can name WHAT was lost without being able to say WHY.
- Blocking saves by default is a behavior change from today's warn-only posture for Excel. Existing workflows that relied on 'warn but proceed' will now need an explicit accept_fidelity_loss=True step -- a deliberate tightening for safety, but a user-facing behavior change worth calling out plainly rather than shipping quietly.
- Extending render/applescript.py's zero-dialog automation pattern to WRITE operations (not just PDF export) is new-code surface reusing validated infrastructure, but it is still new: sheet-name/range-not-found error classification, and the scratch-copy isolation protocol, need their own test coverage before being trusted the way the existing render path is.

## Estimated effort
Phase 0 (rich_text/keep_vba flags + CellRichText read-path fix): ~1 hour, no API surface change. Phase 1 (the core part-inventory guard: generalized fidelity_report, blocking save_to, accept_fidelity_loss flag, excel_fidelity_report tool, rewritten existing test, new zip-surgery test suite): ~1-2 days. Phase 2 (native-Excel-automation escape hatch: write_cells_via_excel, excel_native_tools.py, state-resync protocol, live-gated tests): ~2-3 days including live testing against real Excel. Phase 3 (raw-XML sheet/style editing, and/or AppleScript-driven chart/pivot/slicer creation): explicitly not scheduled -- each would need its own multi-day-to-multi-week spike and is not recommended until Phase 2 proves insufficient with demonstrated demand.
