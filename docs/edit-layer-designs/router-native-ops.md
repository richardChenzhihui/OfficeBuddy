
# Design Task 3 — Skill Router + Native-Automation Modality

Grounded in the actual code at `/Users/chenzhihui/Desktop/Interest/Project/office_agent`:
`core/session.py` (EditSession/SessionManager), `core/snapshot_manager.py`,
`tools/registry.py` (ToolContext/ToolRegistry/dispatch), `tools/word_tools.py`,
`tools/excel_tools.py`, `render/applescript.py`, `agent/loop.py`, `agent/budget.py`,
`agent/prompts.py`. All AppleScript syntax claims below were verified LIVE this
session with `osascript` dry runs against throwaway copies inside the Word/Excel
sandbox containers (`~/Library/Containers/com.microsoft.{Word,Excel}/Data/tmp/`),
never against user files, always closed with `saving no`, never `activate`d, and
all probe files were removed afterward. Where a guess turned out wrong, that is
reported honestly below — it's directly relevant to the design.

---

## 0. The three modalities (recap + one assumption)

1. **Structured tools** — `word_*` / `excel_*` in `tools/word_tools.py` /
   `tools/excel_tools.py` → `adapters/` → python-docx / openpyxl. Mutates the
   session's live in-memory object directly. Already exists, 121 tests green.
2. **`xml_patch`** — raw OOXML fragment grafted onto `doc.element` (Word) or a
   targeted part (Excel), for content structured tools can't express (tracked-
   change authoring, footnotes, custom field codes). **Not (re)designed here** —
   this is a companion deliverable. I assume only a minimal shape so the routing
   text below is coherent: `word_xml_patch(doc_id, target: selector, xml: str,
   mode: insert_before|insert_after|append_child|replace)`, validated by parsing
   the fragment before grafting and rejecting disallowed namespaces/size. If its
   real design differs, only the routing *examples* need updating, not the
   mechanism recommended in §1.
3. **Native automation (AppleScript ops)** — this task's subject: §2 below.

---

## 1. Routing mechanism (part a)

### 1.1 Three candidate mechanisms, compared

| | Prompt-guidance-only | Capability-error redirects | Hard harness-side router |
|---|---|---|---|
| **What it is** | System prompt + tool descriptions tell the model which modality for which ask; model picks freely | Tool/adapter exceptions, when a structured tool hits a real gap, name the exact right modality/tool to call next | A pre-model classification step (rules or an LLM call) restricts/dispatches before the main model chooses |
| **Cost per turn** | ~free (already-loaded prompt tokens) | Free until a gap is hit; then 1 extra round-trip | An extra LLM call (2x latency/cost) or a maintained rule table |
| **Fits single-model MiniMax-M3?** | Yes — this is exactly how the codebase already steers the model everywhere else (`SYSTEM_EXECUTOR`'s "Efficiency rules", `_word_session`'s "use excel_* tools instead", `RenderError._classify`'s remediation text) | Yes — same style, same file (`registry.py` dispatch already turns every exception into an actionable `{success:false,error}` envelope that drives `agent/budget.py`'s escalation ladder) | Conflicts with two things: (1) the user's own global rule "semantic judgment must come from an LLM, never from code" — a *rule-based* router classifying "does this need native automation" is exactly the kind of meaning-judgment that rule bans; a *second-LLM-call* router avoids that but doubles cost/latency for every single turn just to pick a tool category the model could pick itself from tool descriptions; (2) the harness's existing philosophy is "trust the model to choose tools, supervise the *outcome*" (plan → execute → render → verify), not "gate the *input* choice" |
| **Failure mode if wrong** | Model tries the wrong modality once, gets a real tool error, self-corrects (this is the *existing* recovery path, already relied on) | None new for structured→native (the "gap" a redirect fires from often doesn't exist — see below) | A rule table that's stale or a second model call that mis-classifies produces silent misrouting the main model never sees, worse than a wrong-tool error the model *can* read and recover from |
| **Maintenance** | One prompt section | A growing set of precise exception messages at known gap points, same style as existing code | A second classification surface to keep in sync with tool descriptions — double the maintenance |

**A specific limitation of capability-error redirects for the native-automation
modality**: for structured-tools↔xml_patch the redirect fires naturally (a
structured tool call fails, the error names the right modality). But most
native ops have **no structured-tool call to fail in the first place** —
"accept all tracked changes" has no `word_*` call that could raise
"unsupported, try native automation," because python-docx has no revision-
acceptance concept to attempt at all. So native-op discovery cannot rely on a
failure signal; it must be **prompt-guidance-first** and self-describing tool
names/descriptions (`accept_all_revisions`, `recalculate_workbook`) do the
rest.

### 1.2 Recommendation

- **Primary**: prompt-guidance-only for the categorical decision ("which
  modality class handles this kind of ask"). Near-free, and this is the
  mechanism the codebase already trusts MiniMax-M3 with everywhere.
- **Secondary (enforcement)**: capability-error redirects at every *known*
  structured-tool gap boundary, in the exact style already used
  (`_word_session`, `_resolve`, `RenderError._classify`) — precise, actionable,
  names the next tool to call. This is where real engineering effort goes:
  a growing table of "if the model tries X the wrong way, tell it exactly Y,"
  added incrementally as gaps are discovered (see Phase 2 below).
- **Explicitly NOT recommended**: a hard harness-side semantic router. It
  violates the user's own "no code-level semantic judgment" rule if
  rule-based, and doubles cost/latency if LLM-based, for a decision the model
  already has enough signal to make from tool descriptions + the prompt below.
  The one *mechanical* (non-semantic) filtering that IS worth doing cheaply:
  `_word_session`/`_excel_session`-style runtime type guards (already exist)
  extended to native tools — reject at dispatch time, don't try to hide the
  tool from the schema list. Hiding tools per-session would require the
  registry to know which doc_id the *next* call targets before it's made,
  which it can't; the existing runtime guard is simpler and already proven.

### 1.3 Actual routing text (add to `agent/prompts.py`, `SYSTEM_EXECUTOR`)

Insert as a new section after "# Selectors" and before "# Efficiency rules"
(same terse, hyphenated register as the rest of the prompt):

```
# Modality routing — three ways to edit, pick the cheapest that works
1. word_*/excel_* structured tools are the DEFAULT for everything they cover: \
text, formatting, tables, rows/cols, charts, cell values/formulas, borders/\
shading. Always try these first.
2. word_xml_patch/excel_xml_patch (raw OOXML fragment) is ONLY for content \
structured tools cannot express at all (tracked-change authoring, footnotes, \
custom field codes). Reach for it when a structured-tool error explicitly \
tells you to, using the fragment shape the error describes — never \
speculatively, and never as a first attempt.
3. Native ops (accept_all_revisions, reject_all_revisions, update_all_fields, \
recalculate_workbook) are ONLY for whole-document, application-level \
operations with no per-element target and no representation in python-docx/ \
openpyxl's object model: resolving ALL tracked changes at once, recomputing \
ALL field/TOC display text, recalculating ALL formulas. Recognize these by \
the ask itself ("accept the changes", "update the table of contents", \
"recalculate the sheet") — they reopen the real app and are slower and less \
granular than editing the in-memory object, so never use one for a targeted \
content edit a structured tool already handles, and never call one "just to \
be safe."
Only the named native ops exist — never invent an AppleScript command or ask \
for one to be run; if no existing tool fits, say so and ask the user.
```

The last line is load-bearing for §3 (no free-form scripting) — it forecloses
the model ever *asking* for a capability that doesn't exist as a named tool.

---

## 2. Native-automation modality (part b)

### 2.1 Curated op library — what's worth having, and why

Enumerated with the reasoning test: *does this need Word/Excel's own engine
because the operation is whole-document/application-level, has no
python-docx/openpyxl representation, and re-implementing it would mean
reinventing part of Word/Excel's own document-resolution logic?*

| Op | App | Verified live? | Argument |
|---|---|---|---|
| `accept_all_revisions` | Word | **Yes** | python-docx can *author* a tracked insert via raw `OxmlElement` (proven to survive round-trip per the empirical findings), but it has no concept of *resolving* revision state — accepting requires coherently merging every `w:ins`/`w:del`/format-change marker across runs, paragraphs, and tables, which is exactly what Word's own revision engine already does correctly. Reimplementing it in lxml is high-risk, low-value. |
| `reject_all_revisions` | Word | **Yes** | Same argument, opposite direction. |
| `update_all_fields` | Word | **Yes** (generic `fields`; TOC-specific `update` **not** independently verified — see §2.4) | Fields inserted via xml_patch/structured tools (PAGE, TOC, REF) have no cached display text — a human opening the saved file without pressing F9 sees blank/stale text. Only Word's own layout engine knows the real page count / heading text to bake in. We already pay for Word automation (render pipeline) — this reuses it to also fix the *saved artifact*, not just the verification screenshot. |
| `recalculate_workbook(mode)` | Excel | **Yes** (`calculate` / `calculate full rebuild`) | openpyxl never *computes* formula values — it only stores whatever was cached at last real-Excel save. After `excel_write_cells`/`excel_edit_formula` touch cells other formulas depend on, every dependent's cached value is stale until something recalculates. Only real Excel can do this. |
| `refresh_pivot_tables` / `refresh_all_data_connections` | Excel | **No — unverified** | High value in principle (pivot caches are exactly the "openpyxl can't represent this" empirical gap already flagged for chart/image round-trip), but I only confirmed the `pivot tables of sheet` collection *exists* (count=0 on a pivot-less test file); I did not have a workbook with a real pivot table to verify the `refresh` command against. **Do not ship without a dedicated spike.** See §2.5 (Tier 2) for why this also needs an architecture change, not just a syntax fix. |
| `protect_sheet` / `unprotect_sheet` | Excel | **Attempted, both failed** | Tried `protect sheet 1 of theBook` (syntax error) and `protect sheet 1 of theBook contents true` (syntax error). The correct incantation is unresolved — needs a real dictionary lookup (this machine's `xcode-select` points at Command Line Tools only, so `sdef` is unavailable; use Script Editor's Window → Open Dictionary, or Microsoft's VBA object-model docs, since Office's AppleScript nouns/verbs largely mirror VBA parameter names). **Not shipped.** Also note: protection state changes are consent-sensitive (defeats a protection someone set for a reason) — even once verified, gate behind explicit user confirmation in the tool description, not silent `mutates=True` execution. |

**Explicitly evaluated and rejected as native ops:**

- **`add_page_numbers`** (the "?" in the task prompt, resolved): this is
  **not** a native-automation op. Inserting a `PAGE` field is pure OOXML
  content insertion (`w:fldSimple`/`w:instrText` sequence) — exactly the
  pattern already empirically proven to survive save/reopen via `OxmlElement`.
  It belongs to the `xml_patch` modality (or even a small structured-tool
  helper), not to whole-document application automation. No live Word engine
  capability is needed to *insert* the field — only `update_all_fields`
  (already in the library) is needed afterward to bake in its *displayed*
  text.
- **Selective/indexed revision accept-reject** (accept only the 3rd change):
  no verified targeting mechanism exists in Word's AppleScript bridge for
  "the Nth revision matching a description," and building one would need a
  read-only `list_revisions` companion tool plus an index-stable accept/reject
  template — real engineering cost for a rare ask. Deferred, not phase-1.

### 2.2 Fixed AppleScript templates (`render/native_ops.py`, new file, same
invariants as `render/applescript.py`)

All four verified live this session. Reused invariants from
`render/applescript.py`: sandbox-container paths only, `display alerts`
off, no `activate`, `_q()` escaping, dual subprocess+AppleScript timeout,
and — critically — **route through the existing `run_applescript()` +
`_AUTOMATION_LOCK`**, not a separate subprocess call, so a native op can never
race a render export (both use `active document`/`active workbook`).

```python
"""render/native_ops.py — curated, hand-verified AppleScript templates for
whole-document ops with no python-docx/openpyxl representation. Every op was
verified with a live osascript dry run against a throwaway sandbox-container
copy before being added here — Office's AppleScript dictionaries are
idiosyncratic and DO fail in surprising, silent ways (see the design doc:
`repeat with d in documents` + `count` is not understood by Word 2024; a
direct-by-path document reference for an already-open doc is not either).
Never add an op without repeating that verification."""
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Optional

from .applescript import _q, run_applescript, RenderError


@dataclass(frozen=True)
class NativeOpSpec:
    name: str
    app: str          # "Microsoft Word" | "Microsoft Excel"
    tier: int         # 1 = safe to reload into the live EditSession object
                       # 2 = finishing-only, never reloaded (see design §2.5)
    build_script: Callable[..., str]
    default_timeout: float = 60.0


def _word_accept_all(path: Path) -> str:
    return f'''
with timeout of 55 seconds
    tell application "Microsoft Word"
        set display alerts to alerts none
        open (POSIX file "{_q(path)}")
        set theDoc to active document
        accept all revisions in theDoc
        save theDoc
        close theDoc saving no
    end tell
end timeout
'''


def _word_reject_all(path: Path) -> str:
    return f'''
with timeout of 55 seconds
    tell application "Microsoft Word"
        set display alerts to alerts none
        open (POSIX file "{_q(path)}")
        set theDoc to active document
        reject all revisions in theDoc
        save theDoc
        close theDoc saving no
    end tell
end timeout
'''


def _word_update_fields(path: Path) -> str:
    return f'''
with timeout of 55 seconds
    tell application "Microsoft Word"
        set display alerts to alerts none
        open (POSIX file "{_q(path)}")
        set theDoc to active document
        repeat with f in fields of theDoc
            try
                update f
            end try
        end repeat
        save theDoc
        close theDoc saving no
    end tell
end timeout
'''
# NOTE: TOC-specific `update` on the `table of contents` collection was NOT
# independently verified (no TOC-bearing test doc was available this
# session) -- see design §2.4 open item before relying on this for
# TOC-heavy documents.


def _excel_recalculate(path: Path, mode: str = "normal") -> str:
    verb = "calculate full rebuild" if mode == "full_rebuild" else "calculate"
    return f'''
with timeout of 55 seconds
    tell application "Microsoft Excel"
        set display alerts to false
        open "{_q(path)}"
        set theBook to active workbook
        {verb}
        save theBook
        close theBook saving no
    end tell
end timeout
'''


OPS = {
    "accept_all_revisions": NativeOpSpec("accept_all_revisions", "Microsoft Word", 1, _word_accept_all),
    "reject_all_revisions": NativeOpSpec("reject_all_revisions", "Microsoft Word", 1, _word_reject_all),
    "update_all_fields":    NativeOpSpec("update_all_fields",    "Microsoft Word", 1, _word_update_fields),
    "recalculate_workbook": NativeOpSpec("recalculate_workbook", "Microsoft Excel", 1, _excel_recalculate),
}


def run_native_op(session, op_name: str, timeout: Optional[float] = None, **params) -> None:
    """The reload+snapshot-safe sequence every Tier-1 op must follow. See
    design §2.3 for why each step exists and what happens if it's skipped."""
    spec = OPS[op_name]
    session.flush()                                   # 1
    pre_bytes = session.working_path.read_bytes()      # 2
    script = spec.build_script(session.working_path, **params)
    try:
        run_applescript(script, timeout or spec.default_timeout, spec.app)
    except RenderError:
        session.working_path.write_bytes(pre_bytes)    # 3 (rollback)
        raise
    session.reload_from_disk()                          # 4
```

`tools/native_tools.py` (new file, same shape as `word_tools.py`):

```python
"""Native-automation tools: fixed, curated AppleScript-backed ops for
whole-document operations with no python-docx/openpyxl representation. Every
input model here is typed parameters only, substituted into a KNOWN-SAFE
script skeleton in render/native_ops.py — there is deliberately no field
anywhere in this file that accepts script text. See design doc §3 for why."""
from typing import Literal

from pydantic import BaseModel, Field

from ..render.native_ops import run_native_op
from .registry import REGISTRY, ToolContext


def _word_session(ctx: ToolContext, doc_id: str):
    session = ctx.sessions.get(doc_id)
    if session.doc_type != "word":
        raise ValueError(f"doc_id '{doc_id}' is an Excel document; this native op is Word-only.")
    return session


def _excel_session(ctx: ToolContext, doc_id: str):
    session = ctx.sessions.get(doc_id)
    if session.doc_type != "excel":
        raise ValueError(f"doc_id '{doc_id}' is a Word document; this native op is Excel-only.")
    return session


class DocIdInput(BaseModel):
    doc_id: str = Field(..., description="Document id")


@REGISTRY.register(
    "accept_all_revisions",
    "Accept ALL tracked changes via the real Word engine (python-docx cannot "
    "resolve revision state). Whole-document — not for one specific change.",
    DocIdInput, mutates=True,
)
def accept_all_revisions(ctx: ToolContext, p: DocIdInput) -> dict:
    session = _word_session(ctx, p.doc_id)
    run_native_op(session, "accept_all_revisions")
    return {"applied": "accept_all_revisions"}


@REGISTRY.register(
    "reject_all_revisions",
    "Reject ALL tracked changes via the real Word engine, restoring "
    "pre-change content. Whole-document.",
    DocIdInput, mutates=True,
)
def reject_all_revisions(ctx: ToolContext, p: DocIdInput) -> dict:
    session = _word_session(ctx, p.doc_id)
    run_native_op(session, "reject_all_revisions")
    return {"applied": "reject_all_revisions"}


@REGISTRY.register(
    "update_all_fields",
    "Recompute and bake in the displayed text of every field (PAGE, "
    "NUMPAGES, TOC, REF, ...) via the real Word engine. Run as a finishing "
    "step after edits that could shift pagination or TOC-referenced headings.",
    DocIdInput, mutates=True,
)
def update_all_fields(ctx: ToolContext, p: DocIdInput) -> dict:
    session = _word_session(ctx, p.doc_id)
    run_native_op(session, "update_all_fields")
    return {"applied": "update_all_fields"}


class RecalculateWorkbookInput(BaseModel):
    doc_id: str = Field(..., description="Document id")
    mode: Literal["normal", "full_rebuild"] = Field(
        "normal",
        description=(
            "normal: recompute dependents of changed cells (fast). "
            "full_rebuild: force-recompute every formula (slow; only if "
            "'normal' left stale-looking values)."
        ),
    )


@REGISTRY.register(
    "recalculate_workbook",
    "Recalculate all formulas via the real Excel engine and bake results "
    "into the saved file (openpyxl never computes formula values itself). "
    "Run after writing/editing formulas whose dependents need correct "
    "cached values.",
    RecalculateWorkbookInput, mutates=True,
)
def recalculate_workbook(ctx: ToolContext, p: RecalculateWorkbookInput) -> dict:
    session = _excel_session(ctx, p.doc_id)
    run_native_op(session, "recalculate_workbook", mode=p.mode)
    result = {"applied": "recalculate_workbook", "mode": p.mode}
    if session.preservation_warnings:
        result["warnings"] = session.preservation_warnings
    return result
```

Wire-up: add `native_tools` to the side-effect import list in
`tools/__init__.py` (currently `from . import excel_tools, session_tools,
word_tools`).

### 2.3 How native ops fit `EditSession` — the reload+snapshot sequence

**The core tension**: `EditSession` holds exactly ONE live in-memory object
(`self.doc`) backed by `self.working_path` on disk (`core/session.py:34-121`).
Every existing mutating tool mutates `self.doc` directly; the file on disk is
only a projection, refreshed by `flush()`. Native ops invert this: Word/Excel
opens `working_path`, mutates it via an application-level command
python-docx/openpyxl can't express, and **saves it back to the same path** —
so `self.doc` becomes stale the instant the AppleScript call returns.

Proof this in-place save works (verified live, `probe3.docx`/`probe3.xlsx`):
`open (POSIX file ...)` → `accept all revisions in theDoc` → **`save theDoc`**
(no format argument — saves in the file's current native format) → `close
theDoc saving no` (already saved; "no" here just means don't re-prompt) →
reopening the result with `python-docx.Document(...)` succeeds. Same pattern
confirmed for Excel with `save theBook` + `openpyxl.load_workbook(...)`.

The sequence in `run_native_op` (§2.2) is:

1. **`session.flush()` before automation.** `self.doc` may hold pending
   mutations from earlier calls in the same step that were never written to
   `working_path` (`session.dirty`). If we skipped this, the native op would
   operate on stale content, and the subsequent reload would **silently
   discard** those in-memory-only edits. This is the single most important
   ordering rule in the whole design — get it wrong and a native op quietly
   reverts the model's own prior work in the same turn.
2. **Snapshot the pre-state bytes** (`pre_bytes = working_path.read_bytes()`)
   for rollback, not for undo (undo already comes for free — see step 4).
3. **Run the AppleScript**, through the existing `run_applescript()` +
   `_AUTOMATION_LOCK` (never a bespoke subprocess call — that would reopen
   the exact render/native-op race the lock exists to prevent).
4. **On success: `session.reload_from_disk()`** — a small new method on
   `EditSession` (additive, `core/session.py`):
   ```python
   def reload_from_disk(self) -> None:
       """Reload the in-memory object from the CURRENT working_path bytes,
       WITHOUT rewriting working_path — for native ops where a real Office
       app already wrote the mutated file directly. Unlike
       reload_from_bytes(), this never touches the on-disk file."""
       self.doc = self._load()
       self.dirty = False
   ```
   This is deliberately distinct from the existing `reload_from_bytes()`
   (used by `SnapshotManager.restore`), which *writes* bytes to
   `working_path` first — here the bytes are already correct on disk (Word
   wrote them), writing them again would be redundant and racy.
5. **On failure: restore `pre_bytes` to `working_path` before re-raising.**
   If the AppleScript errored before Word ever saved, this is a harmless
   no-op (bytes are already identical). If it errored *after* a save but
   before/during close (e.g. the "close" step itself times out), this
   prevents a half-mutated file from silently becoming the session's working
   copy — see the failure-mode discussion below for what it does *not* fix.
6. **No new snapshot code needed.** `registry.dispatch()` already snapshots
   generically after every successful `mutates=True` call, via
   `session.to_bytes()` → `self.doc.save(buf)` — i.e. it re-serializes
   *whatever `self.doc` currently is*. Because step 4 already reloaded
   `self.doc` to reflect the native mutation, this generic mechanism
   snapshots the *post-native-op* state for free, and `undo`/
   `list_snapshots`/`restore_snapshot` work unmodified.
   - **For Word**, this re-serialization is safe: the empirical finding is
     that python-docx's load→save round-trip preserves all ZIP parts.
   - **For Excel**, this re-serialization carries the SAME pre-existing risk
     already surfaced by `_detect_lossy_parts`/`preservation_warnings` in
     `core/session.py` (openpyxl fully re-serializes and can silently drop
     charts/images/pivots outside its object model) — this design does not
     make that risk worse (any subsequent `flush()`/`save_document` already
     carries it), but `recalculate_workbook`'s tool result explicitly echoes
     `session.preservation_warnings` so the caveat surfaces every time
     instead of staying silent.

### 2.4 Failure modes (all empirically probed, not theoretical)

- **AppleScript errors before Word/Excel saves** → `pre_bytes` restore is a
  no-op, clean re-raise as `RenderError`/`RenderTimeout`/`AutomationDenied`
  (already-classified by `_classify()` in `render/applescript.py` — reused
  as-is, no new error taxonomy needed).
- **AppleScript errors after save but the "close" step doesn't complete**
  (e.g. subprocess timeout kills `osascript` mid-script): `working_path` is
  restored from `pre_bytes`, BUT **the real Word/Excel app may still have the
  document open in a window**, holding it in a state our restore didn't
  affect. I tried to build an automatic recovery for this and it **failed
  live, twice**:
  - `repeat with d in documents ... if (POSIX path of (full name of d)) is
    targetPath then close d ... ` → silently did **nothing** (the `try`
    swallowed a real error: `"every document" doesn't understand the "count"
    message` — Word 2024's AppleScript bridge does not support iterating/
    counting the `documents` collection the way this pattern assumes).
  - `close document (POSIX file "...") saving no` → `"The object you are
    trying to access does not exist"` — this reference form only works for
    *opening*, not for referencing an already-open document by path.
  - The one thing that DID work: `close every document saving no` — but this
    is a blunt, workbook/app-wide hammer that would also discard any
    *unrelated* document the user happens to have open in the same app,
    which is unacceptable (this system's own safety posture never discards
    user data without confirmation).
  - **Conclusion, stated plainly in the design**: there is no verified,
    safe, automatic recovery for a stuck native-op window in this Office
    build. Phase 1 does **not** attempt one. Instead: treat a native-op
    timeout/failure as a **hard stop** surfaced to the model/user with a
    specific message ("a previous automation step may have left
    Microsoft Word/Excel with a document open in an unexpected state;
    close it manually if you see a window, or run `office-agent doctor`"),
    consistent with the existing philosophy of classified, actionable errors
    rather than blind auto-retry.
- **Reload produces a `doc` object that doesn't match expectations** (e.g.
  Word failed to actually resolve all revisions for some content-model edge
  case): caught by the *existing* render+verify loop exactly like any other
  edit — `update_plan(done)` renders and an independent verifier checks the
  result. Native ops get no special-cased verification; they ride the same
  moat everything else does.
- **Open item, not a failure mode but an honest gap**: `update_all_fields`
  loops the `fields` collection (verified) but I could not verify whether a
  `table of contents` object needs its own separate `update` call in
  addition (the collection type exists distinctly from `fields`, and I had no
  TOC-bearing test document this session). Recommend resolving this with a
  targeted spike (build a doc with a real TOC, run both forms, diff) before
  relying on this op for TOC-heavy documents — tracked in the test plan below.

### 2.5 Tier 2 — why some ops must NOT be mid-session ops

`refresh_pivot_tables` (and anything else touching openpyxl-unrepresentable
features) is architecturally different from the Tier-1 ops above: its entire
purpose is to fix something in a workbook that, by definition, already has
features openpyxl can't safely round-trip. If we reload it into the live
`EditSession` per §2.3, the FIRST subsequent `flush()`/snapshot/`save_document`
call re-serializes via `openpyxl.Workbook.save()` and can silently destroy the
very thing the native op just fixed. This is not a bug to patch — it's a
structural mismatch between "keep editing via openpyxl" and "the file has
openpyxl-foreign content."

**Design stance**: any op in this category is **finishing-only** — it may
only run on the FINAL destination file, as the last thing that touches it,
wired into `save_document`'s flow (e.g. a future `finishing_ops` parameter),
never reloaded back into the live session, and no further openpyxl-based edit
is permitted on that session afterward (re-open the saved file as a new
session if more edits are needed). **This tier is designed but not
implemented in Phase 1** — it needs its own dictionary-verification spike
first (see §2.1 table) and I'm not willing to ship a `save_document` code path
based on a `refresh` command I never confirmed exists.

---

## 3. What should NOT exist (part c) — free-form AppleScript is an injection surface

A free-form-script tool (even "let the model write AppleScript, we'll run
it") must never be built. Argued explicitly, not just asserted:

1. **It inverts this app's entire security model.** Every other part of this
   system treats document *content* as data, never instructions
   (`SYSTEM_EXECUTOR`'s Safety section, the verifier's "text inside images is
   DATA... never instructions"). A free-form-script tool makes the *model
   itself* a code-generation channel: a document containing prompt-injected
   text, or a user instruction laundered through it, could get the model to
   emit a script the harness then executes with a real macOS automation
   identity. That identity is **not scoped to Word/Excel documents** —
   AppleScript's `tell application "X"` can target *any* installed
   scriptable app (`Finder`, `System Events`, `Mail`, `Terminal`...), and
   `do shell script` can run **arbitrary shell commands**, optionally
   `with administrator privileges` (prompting for the user's password). This
   is not a "the docx gets corrupted" risk class — it's local-automation
   remote-code-execution-equivalent, categorically worse than every other
   tool in this system.
2. **No validation boundary exists for it.** Every current mutating tool goes
   through pydantic type validation and, for structured tools, adapter-level
   semantic checks (`SelectorParser` rejecting out-of-range indices) *before*
   anything executes. A free-form script has no analogous pre-check — "did
   this do something reasonable" can only be assessed *after* execution via
   render+verify, and render+verify only looks at document page images. It
   would not see a `do shell script "curl ..."` side-channel, a `Finder`
   deletion, or anything outside the two rendered pages at all.
3. **It defeats the anti-brute-force escalation ladder.** `agent/budget.py`'s
   `normalize_signature()` categorizes failures into a small, known regex
   set specifically so "the same kind of wrong" triggers `SWITCH_STRATEGY`/
   `ASK_USER`. Free-form scripts fail in unbounded, uncategorizable ways —
   the model could paper over a broken script with endless "creative"
   variations, each a novel signature, never tripping the ladder.
4. **No template-level zero-dialog guarantee.** Every op in §2.2 was
   hand-verified for the container-path / `display alerts` / no-`activate` /
   dual-timeout invariants. A model-composed script has no such guarantee per
   call — omit `display alerts` suppression, forget `close ... saving no`,
   reference a path outside the sandbox container (reintroducing exactly the
   macOS file-access dialog this whole architecture exists to avoid), or
   `activate` (stealing focus) — any one of these regresses the "zero
   dialogs, zero focus-steal" guarantee the README calls out as a core
   feature.
5. **A "restricted DSL" middle ground doesn't actually help.** Composing
   scripts from an allow-listed set of verbs/nouns still lets adversarial
   content steer *which* allow-listed sequence executes, and building an
   interpreter/sandboxer for a constrained AppleScript subset is a bespoke,
   under-tested security boundary — strictly worse than enumerating the
   handful of actually-needed whole-document ops as fixed templates with zero
   free script-text parameters (every parameter in §2.2 is a typed value —
   an enum, a bool, a path already owned by the session — substituted via
   the existing `_q()` escaping helper into a KNOWN-SAFE skeleton).
6. **Empirical evidence from this very session reinforces the point.**
   Writing *safe, well-intentioned* AppleScript for internal recovery
   purposes — with full context and no adversarial pressure — still produced
   two silently-wrong scripts in a row (§2.4). If a careful author with full
   context gets Word's AppleScript dictionary wrong twice in a row, an LLM
   improvising a script per-call, under no obligation to verify it, cannot be
   trusted to do better — curation-then-verification isn't just a security
   stance here, it's a correctness necessity given how idiosyncratic Office's
   AppleScript surface actually is.

---

## 4. Test plan

Fixtures reused from the existing suite: `word_doc_path`/`excel_doc_path`
(generate a minimal `.docx`/`.xlsx` into `tmp_path`, see `tests/conftest.py`;
used by e.g. `tests/tools/test_direct_apply_no_double_exec.py`), `ctx`
(`ToolContext()` + `close_all()` teardown).

**Unit — `tests/unit/test_session_snapshot.py` (extend)**
- `reload_from_disk()` round-trip: write different bytes directly to
  `session.working_path`, call `reload_from_disk()`, assert `session.doc`
  reflects the new content and `dirty is False`; then snapshot and `undo(1)`,
  assert it restores the pre-write state — mirrors the existing
  `test_snapshot_undo_roundtrip` but for the disk-write-then-reload path.

**Unit — `tests/unit/test_native_ops.py` (new, mocked AppleScript)**
- Monkeypatch `render.native_ops.run_applescript` to a fake that (a) on
  success, writes deterministic bytes to the path it's given and returns; (b)
  on failure, raises `RenderTimeout`. Assert:
  - call order is `flush → read pre_bytes → run_applescript → reload_from_disk`
    (success path) — a spy/call-order assertion.
  - on failure, `working_path` bytes equal `pre_bytes` (rollback), and the
    exception propagates un-swallowed.
  - a native op called on a session with unflushed in-memory edits does NOT
    lose them (write via `session.doc`, don't flush, call `run_native_op`,
    assert the fake automation saw the flushed content).

**Integration — `tests/tools/test_native_tools_registry.py` (new)**
- `REGISTRY.dispatch(ctx, "accept_all_revisions", {...})` through the real
  registry with `native_ops.run_native_op` monkeypatched to a no-op stub;
  assert `result["success"]`, `result["snapshot_id"]` present,
  `session.dirty` correctly reset, and an `audit.jsonl` entry was written —
  confirms native ops don't bypass any of the registry's generic machinery.
- `recalculate_workbook` on a workbook with `preservation_warnings` set:
  assert the tool result echoes the warnings.
- **Guardrail regression test**: assert no tool in `REGISTRY.tools` accepts a
  free-form script/command string — e.g. assert none of the pydantic input
  models has a field literally named `script`/`applescript`/`command`, and
  assert `"run_applescript"` (or any equivalent) is never a registered tool
  name. This exists specifically to catch a future accidental regression
  toward §3's rejected design.

**Live battery additions (extends the existing 13-scenario battery,
`tests/smoke/` / manual, since these require the real apps):**
1. Word doc with actual tracked changes (author via a real Word track-changes
   session, or a raw `w:ins`/`w:del` `OxmlElement` per the empirical
   finding) → `accept_all_revisions` → render+verify the content is baked in
   and confirm via `python-docx`/raw XML inspection that `w:ins`/`w:del` are
   gone.
2. Same setup → `reject_all_revisions` → verify reverted content.
3. Word doc with an existing TOC whose heading text is then changed via a
   structured tool → `update_all_fields` → verify via render that the TOC
   entry now matches the new heading (this is also the live test that
   resolves the §2.4 open item about `table of contents`-specific `update`).
4. Excel: write a formula via `excel_edit_formula` that depends on a cell
   changed by another structured-tool call in the same step →
   `recalculate_workbook(mode="normal")` → `excel_read_cells` (data_only) and
   assert the value matches an independently computed expected result.
5. **Manual, once**: deliberately induce a native-op timeout (e.g. a
   near-zero timeout against a real doc) and confirm by hand (a)
   `working_path` is restored, (b) whether a subsequent normal `render()`
   call still succeeds despite a possibly-stuck Word window — document
   whatever is actually observed rather than assuming the §2.4 hard-fail
   stance is sufficient; if it isn't, that's a finding for a follow-up, not
   something to paper over.

**Process gate (not code): dictionary-verification-before-ship.** Before any
new native op is added to `OPS`, require a written-down `osascript` dry-run
transcript (exact command + exit code/output) against a throwaway
sandbox-container copy, in the PR description — the exact discipline used to
produce this design. This is how the `protect_sheet` failure and the two
failed recovery-script attempts were caught here instead of being shipped as
if confirmed.

---

## 5. Phased rollout

- **Phase 0** (near-zero risk, fully additive): land `reload_from_disk()` on
  `EditSession`. No behavior change for any existing path; unit-testable in
  isolation.
- **Phase 1**: land `render/native_ops.py` + `tools/native_tools.py` with
  **exactly the four osascript-verified Tier-1 ops**
  (`accept_all_revisions`, `reject_all_revisions`, `update_all_fields`,
  `recalculate_workbook`), wire into `tools/__init__.py`, add the §1.3
  routing text to `SYSTEM_EXECUTOR`, land the full unit+integration test
  suite from §4. Independent of the `xml_patch` modality landing — no
  dependency either way.
- **Phase 2**: capability-error-redirect audit pass over
  `word_tools.py`/`excel_tools.py`/adapters — add precise "structured tool
  can't do X, call Y" messages at every *known* gap, matching existing
  exception style. Incremental, low-risk, additive-only.
- **Phase 3**: run the live-battery scenarios (§4) against the real apps;
  specifically probe modality confusion (ask MiniMax-M3 to "accept all the
  changes" and confirm it reaches for `accept_all_revisions` rather than
  looping `word_edit_text` calls) since this is the one part of the design
  that is *not* mechanically verifiable and rests on the model reading tool
  descriptions correctly.
- **Phase 4 (optional, demand-gated)**: resolve the Tier-2 candidates
  (`refresh_pivot_tables`, `protect_sheet`) properly — a real dictionary
  lookup (Script Editor GUI, since `sdef` needs full Xcode which isn't
  installed here) followed by the same live-verification discipline, then
  design the `save_document`-integrated finishing-pass architecture from
  §2.5. Do not build this speculatively; only if a concrete workload needs
  it, since it requires an architectural addition (a save-time-only
  execution path with no reload back into the live session) that doesn't
  exist yet.
- **Never phased in**: selective/indexed revision accept-reject (§2.1), any
  free-form-script tool or restricted-DSL middle ground (§3),
  `add_page_numbers` as a native op (resolved to belong to `xml_patch`/
  structured tools instead, §2.1).

---

## Honest summary of what's solid vs. what's a hypothesis

**Solid (verified live this session, safe to build on):** `accept all
revisions in theDoc`, `reject all revisions in theDoc`, generic `fields`
iteration + `update`, in-place `save theDoc`/`save theBook` (native format,
no dialog), `calculate`/`calculate full rebuild`, `pivot tables of sheet`
collection exists syntactically, `close every document saving no` works.

**Hypothesis / explicitly flagged as unresolved:** `table of contents`-object-
specific `update` (only the generic `fields` path was exercised), any
`refresh` command on a pivot table (existence of the collection confirmed,
verb not), any `protect`/`unprotect` sheet syntax (two attempts both failed
with syntax errors), automatic recovery from a stuck native-op window (two
attempts both failed — one silently).


## Key decisions
- Routing: prompt-guidance-only is the PRIMARY mechanism (categorical modality choice, near-free, matches how the codebase already steers MiniMax-M3 everywhere else); capability-error redirects (existing exception style in _word_session/_resolve/RenderError._classify) are the SECONDARY enforcement layer for structured-tool gaps; a hard harness-side semantic router is explicitly rejected — a rule-based one violates the user's own 'no code-level semantic judgment' rule, an LLM-based one doubles cost/latency per turn for a single-model setup.
- Native automation is a closed, curated op library of fixed AppleScript templates with ONLY typed parameters (enum/bool/path) substituted into hand-verified script skeletons — never free-form script generation, argued explicitly as a cross-app RCE-equivalent injection surface (do shell script, tell application "Finder"/"System Events"), not just a docx-corruption risk.
- Two-tier native-op classification: Tier 1 (accept_all_revisions, reject_all_revisions, update_all_fields, recalculate_workbook) is safe to reload back into the live EditSession object because python-docx/openpyxl round-trip it losslessly; Tier 2 (pivot/data-connection refresh, sheet protection) touches openpyxl-unrepresentable features and must become a save-time-only finishing pass with no reload into the live session — designed in this doc but explicitly NOT implemented in Phase 1 pending further dictionary verification.
- Every native op follows session.flush() -> read pre_bytes -> run AppleScript -> (rollback pre_bytes on failure) -> session.reload_from_disk() (new additive EditSession method) -> success. No new snapshot code is needed: the registry's existing generic post-call snapshot (session.to_bytes()) already snapshots correctly once self.doc has been reloaded.
- add_page_numbers is explicitly NOT a native-automation op (resolves the '?' in the task) -- it's pure OOXML field insertion, belongs to xml_patch/structured tools; only update_all_fields (baking in the field's displayed text) needs the live app.
- Ship only the 4 ops verified live this session in Phase 1; explicitly exclude protect_sheet/unprotect_sheet (two live syntax attempts both failed) and refresh_pivot_tables (collection existence confirmed, refresh verb unverified) until a dedicated dictionary-verification spike is done.

## Risks
- MiniMax-M3's reliability at choosing correctly among three overlapping modalities (structured/xml_patch/native) at scale is unverified beyond today's reasoning -- only the existing 2-way (word_*/excel_* + harness tools) split has empirical evidence (121 tests, 13-scenario battery); Phase 3's live battery must specifically probe modality confusion (e.g. does 'accept all changes' correctly reach for accept_all_revisions instead of looping word_edit_text calls).
- No verified, safe automatic recovery exists for a native-op script that fails mid-flight after Word/Excel has already saved but before it closed cleanly -- two live recovery attempts this session both failed (a documents-collection iterate-and-close loop silently no-op'd due to an unsupported 'count' message; a direct by-path document reference errored 'object does not exist'). The only working bulk-close idiom ('close every document saving no') is unsafe to run automatically since it would discard any unrelated document the user has open in the same app. Phase 1 ships a hard-fail-and-surface-to-user stance instead of auto-recovery, which may prove too blunt in practice.
- Two Tier-2 candidates are unresolved: protect_sheet/unprotect_sheet syntax failed twice live this session (root cause not identified), and refresh_pivot_tables' refresh verb was never exercised against a real pivot table. Neither should be shipped from this design alone -- both need a dedicated dictionary-verification spike (Script Editor's Open Dictionary, since sdef requires full Xcode which this machine's Command Line Tools install does not provide).
- Office's AppleScript dictionaries are demonstrably quirky and version-specific (confirmed live: bare 'documents' collection iteration + 'count' unsupported, direct-path document reference unsupported) -- a future silent Microsoft AutoUpdate could silently break a verified template with no compile-time signal; recommend a doctor-time canary that dry-runs each native op against a throwaway probe file so a breakage is caught at `office-agent doctor` time, not mid-task.
- recalculate_workbook inherits the PRE-EXISTING openpyxl re-serialization risk for workbooks containing charts/images/pivots outside openpyxl's object model (already flagged via session.preservation_warnings for any editing tool, not unique to this design) -- this design surfaces the caveat in the tool result but does not solve the underlying re-serialization risk.
- The update_all_fields op's coverage of Word's distinct 'table of contents' object collection (separate from the generic 'fields' collection) was not independently verified this session (no TOC-bearing test document was available) -- ship it, but treat TOC-specific correctness as unconfirmed until the Phase 3 live-battery scenario 3 (heading-change -> update_all_fields -> TOC text check) actually runs.

## Estimated effort
Phase 0 (reload_from_disk addition): under an hour, additive and low-risk. Phase 1 (4 verified native ops + render/native_ops.py + tools/native_tools.py + routing prompt text + full unit/integration test suite from §4): roughly 2-3 focused engineering days, building on the live osascript verification already done in this session. Phase 2 (capability-error-redirect audit across word_tools.py/excel_tools.py/adapters): 1-2 days, incremental. Phase 3 (live battery: 4 scripted scenarios + 1 manual timeout-recovery probe against real Word/Excel): about 1 day. Phase 4 (Tier-2 dictionary spike for refresh_pivot_tables/protect_sheet + save_document-integrated finishing-pass architecture): 1-2 days, optional and demand-gated -- do not schedule unless a concrete workload needs it.
