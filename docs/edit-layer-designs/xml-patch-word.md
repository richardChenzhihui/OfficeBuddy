
# Word raw-XML modality: `xml_query` / `xml_patch` on the live python-docx tree

All claims below marked **[verified]** were checked by running real code against this repo's exact
dependency versions (python-docx 1.2.0, openpyxl 3.1.5) and, for rendering claims, through the
project's actual AppleScript export path against a real Microsoft Word install on this machine —
not asserted from memory. Probe scripts lived in `/private/tmp/.../scratchpad`, wrote temporary
files only under that scratchpad and under `~/Library/Containers/com.microsoft.Word/Data/tmp/`
(cleaned up afterward), and never touched anything under the project directory.

## 0. Where this plugs into the existing architecture

No new pipeline. Two new tools operate on the *same* `session.doc` (a live `docx.document.Document`
wrapping a live `lxml` tree over an in-memory `docx.package.Package`) that every other Word tool
already mutates. Concretely:

- `src/office_agent/adapters/word_xml_adapter.py` — new, stateless `WordXmlAdapter` (mirrors the
  existing `WordAdapter` pattern in `adapters/word_adapter.py`): fragment parsing, xpath query,
  the 7 patch ops, part resolution, the footnotes/endnotes part bootstrapper, the validation gate.
- `src/office_agent/tools/word_xml_tools.py` — new, two `REGISTRY.register(...)`-decorated tools:
  `xml_query` (mutates=False) and `xml_patch` (mutates=True). Registered in
  `src/office_agent/tools/__init__.py` alongside the existing `word_tools` import.
- `src/office_agent/core/session.py` — small addition: a per-session handle registry
  (`EditSession._xml_handles`), cleared on `reload_from_bytes` (i.e. on undo/restore_snapshot).
- `src/office_agent/tools/word_tools.py` — one-line augmentation to `word_insert_element`'s
  existing rejection message (capability redirect).
- `src/office_agent/agent/prompts.py` — a few sentences added to `SYSTEM_EXECUTOR`'s existing
  "Capability notes" bullet (routing guidance), and optionally one sentence to `SYSTEM_VERIFIER`.
- `src/office_agent/agent/loop.py` — one conditional line in `_verify_doc` (inject an `extra_note`
  when `xml_patch` was used this step) and one conditional line in `_apply_escalation` (mention the
  word_* fallback when the repeatedly-failing tool is `xml_patch`).

Nothing in `SnapshotManager`, `Renderer`, `verifier.py`, or the AppleScript export layer needs to
change — see §3 and §4 for why, with empirical proof.

## 1. Tool schemas

### 1.1 `xml_query` (read-only, `mutates=False`)

```python
# src/office_agent/tools/word_xml_tools.py
from pydantic import BaseModel, Field
from typing import Optional

class XmlQueryInput(BaseModel):
    doc_id: str = Field(..., description="Document id")
    xpath: str = Field(
        ...,
        description=(
            "XPath over the selected OOXML part. Known prefixes: w (wordprocessingml, "
            "the one you'll use 99% of the time), r (relationships), wp/a/pic (drawings/"
            "images), m (math), w14 (Word 2010 extensions). Examples: \"//w:p[.//w:t"
            "[contains(.,'Executive Summary')]]\" (paragraph containing text), \"//w:tbl\" "
            "(all tables), \"//w:ins\" (all tracked insertions), \"//w:fldSimple\" | "
            "\"//w:fldChar[@w:fldCharType='begin']\" (fields)."
        ),
    )
    part: str = Field(
        "body",
        description=(
            "Which OOXML part to query: 'body' (default, main document.xml), "
            "'header:N' / 'footer:N' (section N's header/footer), 'footnotes', "
            "'endnotes', 'comments', 'styles', 'numbering', 'settings'. "
            "'footnotes'/'endnotes' only exist after xml_patch op='ensure_standard_part'."
        ),
    )
    limit: int = Field(20, description="Max matches to return (each match is also size-capped).")


@REGISTRY.register(
    "xml_query",
    "LOW-LEVEL ESCAPE HATCH. Read raw OOXML nodes by XPath and get back stable handles for "
    "xml_patch. Use ONLY for what word_* tools cannot express: footnotes, endnotes, fields "
    "(TOC, PAGEREF, ...), tracked-change inspection/accept/reject, or other raw-XML constructs. "
    "For ordinary text/formatting/table edits use word_edit_text / word_edit_style / "
    "word_insert_element / word_find_replace instead — they are cheaper and already validated. "
    "Handles are valid until the next undo/restore_snapshot on this doc_id; call xml_query again "
    "after any undo.",
    XmlQueryInput,
)
def xml_query(ctx: ToolContext, p: XmlQueryInput) -> dict: ...
```

Output shape:
```json
{
  "count": 3,
  "returned": 3,
  "matches": [
    {"handle": "h4", "tag": "w:p", "path_hint": "w:body/w:p[12]",
     "text_preview": "Executive Summary", "xml": "<w:p>...</w:p>"}
  ]
}
```
`xml` is `lxml.etree.tostring(el, pretty_print=True)` truncated to 4000 chars with a
`"...[truncated, 8213 chars total]"` suffix when long; `path_hint` is a best-effort ancestor/
sibling-index breadcrumb (not a real XPath — just orientation, computed via `getparent()` +
`.index()` walks, so it can't be fed back into another query).

### 1.2 `xml_patch` (mutating, `mutates=True`)

```python
class XmlPatchOp(str, Enum):
    INSERT_BEFORE = "insert_before"
    INSERT_AFTER = "insert_after"
    APPEND = "append"                    # append fragment as target's last child
    REPLACE = "replace"                  # swap target node for fragment
    REMOVE = "remove"                    # delete target node + subtree
    SET_ATTR = "set_attr"                # add/change/remove attributes on target
    UNWRAP = "unwrap"                    # replace target with ITS OWN children, in place
    ENSURE_STANDARD_PART = "ensure_standard_part"  # idempotent part bootstrap

class XmlPatchInput(BaseModel):
    doc_id: str = Field(..., description="Document id")
    op: XmlPatchOp = Field(..., description="insert_before/after/append/replace need "
        "fragment_xml; remove/set_attr/unwrap need only handle; set_attr also needs attrs; "
        "ensure_standard_part needs standard_part instead of handle.")
    handle: Optional[str] = Field(None, description="Handle from a prior xml_query call.")
    fragment_xml: Optional[str] = Field(
        None,
        description=(
            "ONE raw OOXML element, e.g. '<w:r><w:t>hi</w:t></w:r>'. Namespace prefixes "
            "(w:, r:, ...) are auto-declared if you omit xmlns — you may write bare "
            "'<w:r>...' without xmlns=\"...\" boilerplate. Must contain exactly one root "
            "element; call xml_patch once per sibling for multiple insertions."
        ),
    )
    attrs: Optional[Dict[str, Optional[str]]] = Field(
        None,
        description="For set_attr: {'w:id': '5'} sets/changes; {'w:id': None} removes the "
        "attribute. Keys are prefixed like 'w:id' or 'r:id'.",
    )
    standard_part: Optional[str] = Field(
        None, description="For ensure_standard_part: 'footnotes' or 'endnotes'."
    )


@REGISTRY.register(
    "xml_patch",
    "LOW-LEVEL ESCAPE HATCH — mutates raw OOXML. Applies immediately to the working copy; "
    "snapshotted and undoable exactly like word_* tools, and triggers the same real-Word "
    "render + independent visual verification when the step is marked done. Validated before "
    "commit (malformed fragments and gross package corruption are rejected with the document "
    "left UNCHANGED) but this does NOT guarantee OOXML-schema correctness — only Word's own "
    "rendering (checked by the verifier) is ground truth for that. Use ONLY for footnotes, "
    "endnotes, fields (TOC etc.), tracked-change accept/reject, or constructs word_* tools "
    "reject — prefer word_edit_text/word_edit_style/word_insert_element for anything they can "
    "express.",
    XmlPatchInput,
    mutates=True,
)
def xml_patch(ctx: ToolContext, p: XmlPatchInput) -> dict: ...
```

Per-op required-field validation happens in the handler body (same style as
`word_edit_text`'s `TextOperation` branch in `tools/word_tools.py:133-135` — pydantic models stay
permissive/optional, the handler enforces the actual combination and raises `ValueError` with a
specific message naming the missing field).

## 2. Implementation core (`adapters/word_xml_adapter.py`)

### 2.1 Namespace handling **[verified]**

`docx.oxml.ns.nsmap` already contains the full prefix table used everywhere in this design
(`a, c, cp, dc, dcmitype, dcterms, dgm, m, pic, r, sl, w, w14, wp, xml, xsi` — 16 entries, printed
directly from the installed library). Two concrete facts settle the "how do namespaces work"
question:

1. `doc.element.xpath(xpath_str)` (and any node from python-docx's OWN registered class hierarchy —
   e.g. a header/footer/styles/numbering/settings part root) already has the full `nsmap` baked in
   by `BaseOxmlElement.xpath()`'s override (`docx/oxml/xmlchemy.py`) — **no `namespaces=` kwarg
   needed** when the root object is one python-docx itself created.
2. Elements **we** build via `docx.oxml.parse_xml(...)` or `OxmlElement(...)` (a raw `w:ins`
   fragment, a hand-built `w:footnotes` part root) come back as plain `lxml.etree._Element` —
   `BaseOxmlElement`'s override does NOT apply to them, and calling `.xpath()` on them WITHOUT an
   explicit `namespaces=nsmap` kwarg raises `lxml.etree.XPathEvalError: Undefined namespace
   prefix`. **Verified**: `ins_matches[0].xpath(".//w:t/text()")` failed until `namespaces=nsmap`
   was passed explicitly.

   Design consequence: `WordXmlAdapter`'s internal query helper NEVER relies on the implicit
   override — it always calls `root.xpath(xpath_str, namespaces=WORDML_NSMAP)` explicitly, so
   behavior is uniform regardless of whether the root happens to be a registered python-docx class
   or a bare part we bootstrapped ourselves (footnotes/endnotes roots are the latter).

Fragment input ergonomics — the model should be able to write `<w:r><w:t>hi</w:t></w:r>` without
remembering `xmlns:w="http://schemas..."` boilerplate every time:

```python
from docx.oxml import parse_xml
from docx.oxml.ns import nsmap as WORDML_NSMAP
from lxml import etree

class XmlOpError(ValueError):
    """A patch/query op could not be applied. Subclasses ValueError so it's caught by
    registry.dispatch's existing except clause and gets its own error_type tag."""

def parse_fragment(xml_str: str):
    xml_str = (xml_str or "").strip()
    if not xml_str:
        raise XmlOpError("fragment_xml must be one non-empty OOXML element, e.g. '<w:r>...</w:r>'.")
    try:
        return parse_xml(xml_str)               # works if the caller included xmlns decls
    except etree.XMLSyntaxError:
        pass
    decls = " ".join(f'xmlns:{pfx}="{uri}"' for pfx, uri in WORDML_NSMAP.items())
    try:
        wrapper = parse_xml(f"<_oa_wrap {decls}>{xml_str}</_oa_wrap>")
    except etree.XMLSyntaxError as exc:
        raise XmlOpError(f"fragment_xml is not well-formed XML: {exc}") from exc
    children = list(wrapper)
    if len(children) != 1:
        raise XmlOpError(
            f"fragment_xml must contain exactly ONE root element; got {len(children)}. "
            "Call xml_patch once per sibling for multiple insertions."
        )
    return children[0]
```

Attribute keys for `set_attr` go through `docx.oxml.ns.qn()` **[verified]** — `qn('w:id')` →
Clark-notation `{http://schemas.../wordprocessingml/2006/main}id`; `qn('bogus:id')` raises
`KeyError('bogus')`, which the adapter re-raises as `XmlOpError("Unknown namespace prefix "
"'bogus' in attribute key 'bogus:id' — known prefixes: a, c, cp, ..., w, w14, wp, xml, xsi.")`.

### 2.2 Part resolution

```python
def resolve_part(session, part: str):
    doc = session.doc
    if part == "body":
        return doc.element
    if part.startswith("header:") or part.startswith("footer:"):
        kind, _, idx = part.partition(":")
        i = int(idx)
        if not 0 <= i < len(doc.sections):
            raise XmlOpError(f"Section index {i} out of range: {len(doc.sections)} section(s).")
        target = doc.sections[i].header if kind == "header" else doc.sections[i].footer
        return target.part.element
    if part == "styles":
        return doc.part._styles_part.element
    if part == "numbering":
        return doc.part.numbering_part.element
    if part == "settings":
        return doc.part._settings_part.element
    if part == "comments":
        return doc.part._comments_part.element
    if part in ("footnotes", "endnotes"):
        from docx.opc.constants import RELATIONSHIP_TYPE as RT
        rt = RT.FOOTNOTES if part == "footnotes" else RT.ENDNOTES
        try:
            return doc.part.part_related_by(rt).element
        except KeyError:
            raise XmlOpError(
                f"Part '{part}' does not exist in this document yet. Call xml_patch "
                f"op='ensure_standard_part' standard_part='{part}' first."
            )
    raise XmlOpError(f"Unknown part '{part}': expected body, header:N, footer:N, footnotes, "
                      "endnotes, comments, styles, numbering, or settings.")
```
Every branch here is **[verified]** against a real `docx.Document()` — `.element` exists uniformly
on every part type (`CT_HdrFtr`, `CT_Numbering`, `CT_Styles`, `CT_Settings` all confirmed present).

### 2.3 Stable node handles — design and lifetime

Handles are per-`EditSession`, not global: `EditSession` (in `core/session.py`) gets

```python
self._xml_handles: Dict[str, Any] = {}
self._xml_handle_seq = 0

def register_xml_handle(self, el) -> str:
    self._xml_handle_seq += 1
    handle = f"h{self._xml_handle_seq}"
    self._xml_handles[handle] = el
    return handle

def resolve_xml_handle(self, handle: str):
    el = self._xml_handles.get(handle)
    if el is None:
        raise XmlOpError(
            f"Unknown or stale handle '{handle}'. Handles are invalidated whenever the "
            "document is reloaded (undo, restore_snapshot). Call xml_query again."
        )
    if el.getroottree() is not self.doc.element.getroottree():
        # Node was detached by an earlier patch this turn (e.g. removed, or unwrapped away).
        raise XmlOpError(
            f"Handle '{handle}' refers to a node no longer attached to the document "
            "(an earlier xml_patch call this turn may have removed or replaced it). "
            "Call xml_query again to get a fresh handle."
        )
    return el
```

and `reload_from_bytes` (already the sole place `self.doc` is reassigned — used by
`SnapshotManager.restore`, i.e. `undo`/`restore_snapshot`) gets one added line:
`self._xml_handles.clear()`.

Why a plain dict of direct element references is safe and sufficient **[verified by construction]**:
- `EditSession.doc` is set once in `__init__` and only ever *reassigned* (not mutated in place) by
  `reload_from_bytes`. Every other operation — `flush()` (calls `self.doc.save(...)`), `to_bytes()`
  (`doc.save(BytesIO())`), and every `word_*`/`xml_patch` mutation — touches the SAME live object,
  so a handle issued at step N is still valid at step N+50 as long as no undo/restore happened in
  between and the node itself wasn't later detached by another patch. Both invalidation cases
  (reload, detachment) are checked explicitly above.
- No cross-document collision is possible: the registry lives on the `EditSession` instance, so
  `doc_id` A's `"h1"` and `doc_id` B's `"h1"` are different dicts.
- Renderer.render() also calls `session.flush()` internally (`render/renderer.py:31`) — same
  non-reassigning flush, so handles survive render/verify cycles too, meaning an agent can legally
  reuse a handle from an earlier `xml_query` to make a follow-up fix to the same node without
  re-querying, as long as nothing rewound the document since.

### 2.4 The seven ops

```python
class WordXmlAdapter:
    @staticmethod
    def query(session, xpath: str, part: str, limit: int) -> dict:
        root = resolve_part(session, part)
        try:
            all_matches = root.xpath(xpath, namespaces=WORDML_NSMAP)
        except etree.XPathEvalError as exc:
            raise XmlOpError(f"Invalid XPath '{xpath}': {exc}") from exc
        shown = all_matches[:max(1, min(limit, 100))]
        matches = []
        for el in shown:
            if not isinstance(el, etree._Element):
                continue  # xpath can return attribute/text nodes too; skip for handle purposes
            xml = etree.tostring(el, pretty_print=True).decode()
            truncated = len(xml) > 4000
            handle = session.register_xml_handle(el)
            matches.append({
                "handle": handle,
                "tag": etree.QName(el).localname and f"{_prefix_for(el.tag)}:{etree.QName(el).localname}",
                "path_hint": _path_hint(el),
                "text_preview": "".join(el.itertext())[:120],
                "xml": xml[:4000] + (f"...[truncated, {len(xml)} chars total]" if truncated else ""),
            })
        return {"count": len(all_matches), "returned": len(matches), "matches": matches}

    @staticmethod
    def apply(session, op: XmlPatchOp, handle, fragment_xml, attrs, standard_part) -> dict:
        if op == XmlPatchOp.ENSURE_STANDARD_PART:
            return _ensure_standard_part(session.doc, standard_part)

        if handle is None:
            raise XmlOpError(f"op='{op.value}' requires 'handle' from a prior xml_query call.")
        target = session.resolve_xml_handle(handle)

        if op in (XmlPatchOp.INSERT_BEFORE, XmlPatchOp.INSERT_AFTER, XmlPatchOp.APPEND, XmlPatchOp.REPLACE):
            if not fragment_xml:
                raise XmlOpError(f"op='{op.value}' requires 'fragment_xml'.")
            new_el = parse_fragment(fragment_xml)
            if op == XmlPatchOp.INSERT_BEFORE:
                target.addprevious(new_el)
            elif op == XmlPatchOp.INSERT_AFTER:
                target.addnext(new_el)
            elif op == XmlPatchOp.APPEND:
                target.append(new_el)
            else:  # REPLACE
                parent = target.getparent()
                if parent is None:
                    raise XmlOpError("Cannot replace a part's root element.")
                parent.replace(target, new_el)
            return {"applied": op.value, "new_handle": session.register_xml_handle(new_el)}

        if op == XmlPatchOp.REMOVE:
            parent = target.getparent()
            if parent is None:
                raise XmlOpError("Cannot remove a part's root element.")
            parent.remove(target)
            return {"applied": "remove"}

        if op == XmlPatchOp.SET_ATTR:
            if not attrs:
                raise XmlOpError("op='set_attr' requires 'attrs'.")
            for key, value in attrs.items():
                qkey = qn(key) if ":" in key else key
                if value is None:
                    target.attrib.pop(qkey, None)
                else:
                    target.set(qkey, value)
            return {"applied": "set_attr", "attrs": attrs}

        if op == XmlPatchOp.UNWRAP:
            parent = target.getparent()
            if parent is None:
                raise XmlOpError("Cannot unwrap a part's root element.")
            # Spec-correctness special case: restoring a rejected deletion's content must
            # convert w:delText/w:delInstrText back to w:t/w:instrText — a bare w:delText
            # outside a w:del wrapper is not valid content. This is the ONE hardcoded rename;
            # xml_patch does not expose a general-purpose tag-rename primitive.
            if target.tag == qn("w:del"):
                for el in target.iter(qn("w:delText")):
                    el.tag = qn("w:t")
                for el in target.iter(qn("w:delInstrText")):
                    el.tag = qn("w:instrText")
            idx = list(parent).index(target)
            for i, child in enumerate(list(target)):  # lxml re-parents on insert automatically
                parent.insert(idx + i, child)
            parent.remove(target)
            return {"applied": "unwrap"}
```

### 2.5 `ensure_standard_part` — footnotes/endnotes bootstrap **[verified end-to-end]**

python-docx has zero footnote/endnote support and a fresh `docx.Document()` has no
`word/footnotes.xml` part at all (confirmed: `doc.part.package.iter_parts()` lists 13 parts, none
of them footnotes). But `doc.part.comments` **[verified]** shows python-docx's own internal
pattern for lazily creating a brand-new part + relationship + content-type registration
(`docx/parts/document.py:129-140`, the `_comments_part` property) — `XxxPart.default(package)` +
`self.relate_to(part, RT.XXX)`. There's no `FootnotesPart` class, but the underlying opc-python
primitives it's built on (`docx.opc.part.XmlPart`, `docx.opc.packuri.PackURI`,
`docx.opc.constants.CONTENT_TYPE.WML_FOOTNOTES`, `docx.opc.constants.RELATIONSHIP_TYPE.FOOTNOTES`
— **both constants already exist** in the installed library even though the high-level API doesn't
use them) let us replicate the exact same pattern by hand:

```python
def _ensure_standard_part(doc, kind: str):
    from docx.opc.packuri import PackURI
    from docx.opc.part import XmlPart
    from docx.opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT

    if kind not in ("footnotes", "endnotes"):
        raise XmlOpError(f"standard_part must be 'footnotes' or 'endnotes', got {kind!r}.")
    rt, ct, partname, tag, sep_ids = {
        "footnotes": (RT.FOOTNOTES, CT.WML_FOOTNOTES, "/word/footnotes.xml", "footnotes", ...),
        "endnotes":  (RT.ENDNOTES,  CT.WML_ENDNOTES,  "/word/endnotes.xml",  "endnotes",  ...),
    }[kind]
    try:
        doc.part.part_related_by(rt)
        return {"applied": "ensure_standard_part", "part": kind, "created": False}  # idempotent
    except KeyError:
        pass
    boilerplate = _separator_boilerplate_xml(tag)   # w:{footnote,endnote} type=separator id=-1
                                                     # + type=continuationSeparator id=0
    element = parse_xml(boilerplate)
    part = XmlPart(PackURI(partname), ct, element, doc.part.package)
    doc.part.relate_to(part, rt)
    return {"applied": "ensure_standard_part", "part": kind, "created": True}
```

**Empirically verified** in full: built this exact structure (part + relationship + content-type),
appended a `<w:footnote w:id="1">` with real footnote text into the new part's root, added a
`<w:footnoteReference w:id="1"/>` run (with `FootnoteReference` character style) to a body
paragraph, called `doc.save(BytesIO())`, and confirmed via `zipfile`: `word/footnotes.xml` present,
`word/_rels/document.xml.rels` contains the new relationship, `[Content_Types].xml` registers the
content type, and re-opening the bytes with `docx.Document(...)` finds the new part on
`iter_parts()` with the correct partname/content-type. This is a real, working recipe, not a guess.

Scope note: this is intentionally a **narrow whitelist** (`footnotes` | `endnotes` only), not a
general "create any OPC part" tool — see §9.

### 2.6 The validation gate

Three tiers, explicitly answering the task's question ("lxml schema? python-docx re-parse?
re-open probe? render-as-validity-check we already have?"):

1. **Fragment well-formedness** (`parse_fragment`, above) — cheap, happens before any tree
   mutation. Catches unbalanced tags, unknown-but-undeclared namespace prefixes, multi-root
   fragments. **No mutation is attempted if this fails.**
2. **Save/reopen probe** — wraps the whole `xml_patch` handler:
   ```python
   def xml_patch(ctx: ToolContext, p: XmlPatchInput) -> dict:
       session = _word_session(ctx, p.doc_id)
       pre_bytes = session.to_bytes()          # cheap in-memory full-package serialize
       try:
           result = WordXmlAdapter.apply(session, p.op, p.handle, p.fragment_xml, p.attrs, p.standard_part)
           session.flush()                     # forces doc.save() to disk — catches serializer failures
           probe = WordDocument(str(session.working_path))  # independent re-open, NOT session.doc
           _ = probe.paragraphs, probe.tables, probe.sections  # touch the common structural paths
       except Exception as exc:
           session.reload_from_bytes(pre_bytes)   # full rollback: doc, working copy, handle table
           raise XmlOpError(
               f"xml_patch rejected and rolled back ({type(exc).__name__}: {exc}). The document "
               "is unchanged. This usually means a required child/attribute is missing, or the "
               "fragment is nested somewhere invalid."
           ) from exc
       return result
   ```
   This reuses `EditSession.to_bytes()` / `flush()` / `reload_from_bytes()` verbatim — the exact
   primitives `SnapshotManager` already trusts — rather than inventing a parallel serialization
   path. Rollback restores the pre-patch bytes AND clears `_xml_handles` (since
   `reload_from_bytes` clears them), which is correct: any handle from before a rolled-back patch
   attempt that referenced now-reverted content is still valid content-wise, but forcing a re-query
   is the safer, simpler invariant (one rule — "reload always clears" — rather than two).
3. **Real Word render + independent verifier** — this is the EXISTING loop, unmodified (§4). It is
   the actual ground truth for "does this look right", run asynchronously at step-done time, not
   synchronously inside the tool call (would cost an AppleScript round-trip, ~0.6–3s, on every
   single patch — too slow for a tight query/patch loop).

**What tier 2 does NOT catch — verified empirically, not assumed:**
- `doc.save(BytesIO())` performs **zero XSD/schema validation**. Built a deliberately-illegal tree
  (a `<w:tbl>` nested directly inside a `<w:r>`, which is schema-illegal) — `doc.save()` succeeded
  silently, and `docx.Document(reopened_bytes)` opened it fine and read the surrounding paragraph
  text correctly. lxml only checks well-formedness; python-docx's OPC layer only checks that the
  ZIP/parts/relationships/content-types graph is self-consistent, not that any individual part's
  XML conforms to the OOXML schema.
- Went one step further and **rendered that same illegal file through the real AppleScript
  pipeline** (`export_docx_to_pdf`): it exported successfully, with zero dialogs, and Word simply
  ignored the illegal empty table — the surrounding text rendered fine, no repair dialog, no crash.
  This is good news for the zero-dialog invariant (malformed content from xml_patch doesn't hang or
  prompt) but it means **tier 2 + a "successful" render do not jointly guarantee schema-clean
  output** — Word can silently swallow structurally-invalid content that produces no visible
  symptom. The render+verify loop is a correctness backstop for *visible* regressions, not a
  substitute for good-faith schema-correct fragment authoring. State this honestly rather than
  oversell it (see §9, "what NOT to build" — this is exactly why a full schema validator is not
  worth building: even if we built one, Word's own leniency means schema-clean isn't suffient and
  schema-dirty isn't always harmful, so the real signal remains the render).

## 3. Snapshot/undo coverage — confirmed, no changes needed

`SnapshotManager.snapshot()` calls `session.to_bytes()`, which calls `self.doc.save(buf)` — this
serializes the **entire OPC package** (every part reachable from the main document part via
relationships), not just `document.xml`. **Verified concretely**: after using
`ensure_standard_part` to add a brand-new `word/footnotes.xml` part + relationship + content-type
entry to a document, `doc.save(BytesIO())` produced bytes that, when re-opened, still had the new
part with correct partname and content-type. So:
- A snapshot taken after an `xml_patch` call captures the new part along with everything else —
  **undo already works correctly for xml_patch, including part-creation ops, with zero changes to
  `SnapshotManager`.**
- `restore()`/`undo()` call `session.reload_from_bytes(data)`, which reassigns `self.doc` — this is
  precisely the boundary at which `_xml_handles` must be cleared (§2.3), and is the only such
  boundary in the whole codebase.

## 4. Visual verifier coverage — confirmed, one recommended addition

`xml_patch` is registered `mutates=True` and is **not** added to `agent/loop.py`'s
`NON_VISUAL_TOOLS = {"excel_write_cells", "excel_edit_formula"}` (that set is Excel-only data
edits). Following the dispatch path in `loop.py:_handle_tool_inner` → `visual_dirty` gets the
doc_id marked exactly like any `word_*` mutation, and `_on_update_plan`'s `status="done"` path
renders via the real Word app and calls `verify_edit` with an independent multimodal reviewer,
unchanged. **No code changes needed for basic coverage.**

One **recommended, non-required** addition, motivated by an empirical rendering-behavior finding
that the verifier isn't currently told about (§6.3): in `loop.py:_verify_doc`, when
`"xml_patch" in tools_used`, append a short `extra_note` to the existing `extra_note` parameter
already threaded through `verify_edit()`:
```python
if "xml_patch" in tools_used:
    extra_note += (
        "\nNOTE: this step used raw-XML editing. Word's review markup can render tracked "
        "insertions as inline colored/underlined text, and tracked DELETIONS as a margin "
        "callout box connected by a line (not inline strikethrough) — this is normal Word "
        "rendering, not a layout defect. Footnote markers appear as superscript numbers in "
        "the body with note text at the page bottom; a freshly-inserted TOC field may show "
        "placeholder text until the user updates it in Word (F9) — that is expected, not a bug."
    )
```
This is a one-line, deterministic, tool-name-keyed addition (not a semantic classifier) — it fires
whenever `xml_patch` was used this step, regardless of what it did.

## 5. Capability-error redirects

Design principle: **redirects are static text, never intent-guessing code.** (This matches an
explicit standing rule: semantic judgment — "what did the model actually want" — must never be
inferred by string/keyword matching in code; if it needs judgment, it needs an LLM, and here the
LLM is the *agent itself*, which should be told the rule up front rather than have code try to
guess its intent after a failed call.)

Two channels:

1. **Proactive (primary)** — `SYSTEM_EXECUTOR`'s existing "Capability notes" bullet in
   `agent/prompts.py` (the same section that already documents borders/bg_color/Table Grid
   defaults) gets one more line:
   > Word has no native support for footnotes, endnotes, fields (TOC, PAGEREF, etc.), or
   > authoring/accepting/rejecting tracked changes via word_* tools. Use xml_query (to inspect raw
   > OOXML) and xml_patch (to insert/modify/remove it) for these — each tool's description
   > includes a worked recipe. Prefer word_* tools for everything else.
   This is read on every turn before any tool call, so the model self-selects correctly rather
   than discovering the gap by trial and error.
2. **Reactive backstop** — `word_insert_element`'s existing generic rejection
   (`tools/word_tools.py:377-380`, currently `f"Unsupported element type '{element_type}': "
   "expected table, paragraph, or page_break."`) gets a fixed, always-appended suffix:
   > For footnotes, endnotes, TOC/other fields, comments, bookmarks, or tracked changes, use
   > xml_query/xml_patch instead (see their tool descriptions for recipes).
   This appends unconditionally for ANY unrecognized `element_type` — it's not classifying what
   the caller meant, just consistently pointing at the escape hatch.

Related, smaller finding worth flagging (not required for this task, low-risk cleanup): today,
`StyleParams` (schemas/operations.py) has no `model_config`, so pydantic v2's default `extra="ignore"`
means a model inventing an unmodeled field (e.g. `{"strikethrough": true}`) is **silently dropped**,
not rejected — a genuine silent-capability-gap risk distinct from the xml_patch redirects above.
Recommend `model_config = ConfigDict(extra="forbid")` on `StyleParams` as a small follow-up so an
unknown style key becomes a loud, actionable pydantic `ValidationError` ("extra fields not
permitted") instead of a silent no-op. Flagging, not doing — out of this task's direct scope.

## 6. Three worked examples (each grounded in a real probe run against real Word on this machine)

### 6.1 Add a footnote

```
xml_query(doc_id, xpath="//w:p[.//w:t[contains(.,'needs a citation')]]", part="body")
  → {"matches": [{"handle": "h3", "tag": "w:p", ...}]}

xml_patch(doc_id, op="ensure_standard_part", standard_part="footnotes")
  → {"applied": "ensure_standard_part", "part": "footnotes", "created": true}
     # or created:false if the doc already had footnotes — idempotent either way

xml_query(doc_id, xpath="//w:r[last()]", part="body")   # last run of that paragraph, to append after
  → {"matches": [{"handle": "h5", ...}]}

xml_patch(doc_id, op="insert_after", handle="h5", fragment_xml=
  '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>'
  '<w:footnoteReference w:id="2"/></w:r>')

xml_query(doc_id, xpath="/*", part="footnotes")          # get the footnotes root as a handle
  → {"matches": [{"handle": "h7", "tag": "w:footnotes", ...}]}

xml_patch(doc_id, op="append", handle="h7", fragment_xml=
  '<w:footnote w:id="2"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>'
  '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>'
  '<w:r><w:t xml:space="preserve"> Smith, J. (2024). Example Citation.</w:t></w:r>'
  '</w:p></w:footnote>')
```
Footnote id must be picked by the caller/adapter (not auto-incremented by python-docx, which has
no concept of footnotes) — a small helper `_next_footnote_id(footnotes_part)` scanning existing
`w:footnote/@w:id` and returning `max+1` (starting above 0, since -1/0 are reserved for the
separator/continuationSeparator) belongs in the adapter, not the agent's responsibility.
**Verified**: this exact structure (part + relationship + content-type + body reference +
footnote-part content) round-trips correctly through `doc.save()`/re-open.
**Known gap**: content added this way is invisible to `word_read_content`/`get_structure`/
`text_match` selectors — python-docx's `Paragraph.text` walks only `w:r` children directly under
`w:p` via its own fixed accessor and doesn't know about footnotes at all (nor would it look inside
`word/footnotes.xml`). Read footnote content back via `xml_query(part="footnotes")`, not the word_*
read tools.

### 6.2 Insert a TOC field

```
xml_query(doc_id, xpath="//w:p[1]", part="body")   # anchor: first paragraph, insert before it
  → {"matches": [{"handle": "h1", ...}]}

xml_patch(doc_id, op="insert_before", handle="h1", fragment_xml="<w:p/>")
  → {"applied": "insert_before", "new_handle": "h9"}   # empty paragraph to hold the field

xml_patch(doc_id, op="append", handle="h9", fragment_xml=
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>')
xml_patch(doc_id, op="append", handle="h9", fragment_xml=
  '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>')
xml_patch(doc_id, op="append", handle="h9", fragment_xml=
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>')
xml_patch(doc_id, op="append", handle="h9", fragment_xml=
  '<w:r><w:t xml:space="preserve">Heading One .......... 1&#10;Heading Two .......... 2</w:t></w:r>')
xml_patch(doc_id, op="append", handle="h9", fragment_xml=
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>')
```

**Important, empirically-confirmed limitation, stated honestly**: rendered this exact
`begin/instrText/separate/<cached text>/end` field structure through the real AppleScript
`export_docx_to_pdf` path and extracted the resulting PDF's text — **the cached placeholder text
came through completely unchanged**; Word's "Save As PDF" via AppleScript does **not** recalculate
field results. This means:
- A hand-authored TOC field will show *exactly whatever cached text we put in the `separate…end`
  run* in the rendered PDF the verifier sees — never real, current page numbers, because computing
  those requires Word's layout engine, which only runs on an actual field-update (F9 / right-click
  "Update Field" / "Update fields before printing" — none of which the current zero-dialog export
  path triggers, and forcing it via UI automation — bringing Word to the foreground, sending
  keystrokes — would violate the documented "never activate — no focus stealing" invariant in
  `render/applescript.py` and reintroduce exactly the fragility that invariant exists to avoid).
- Best-effort mitigation, and what's actually worth building: synthesize the cached text from real
  data we DO have — read heading-styled paragraphs via the existing `get_structure`/`doc.paragraphs`
  (already available, no new capability needed) and populate the cached run with a plausible
  heading list — **without page numbers** (we have no pagination info; don't fabricate numbers).
  State clearly in the tool's returned result and in the agent's final summary to the user that
  opening the document in Word and updating the field (F9, or right-click → Update Field) will
  populate real page numbers — this is a real, load-bearing limitation to disclose, not paper over.
- This is a case for `word_insert_element`'s capability-note text (§5) to be explicit about too:
  "TOC fields render with placeholder/no-page-number text until the user updates fields in Word."

### 6.3 Accept / reject a pre-existing tracked change

```
xml_query(doc_id, xpath="//w:ins | //w:del", part="body")
  → {"matches": [
       {"handle": "h2", "tag": "w:ins", "text_preview": "EXTREMELY", ...},
       {"handle": "h4", "tag": "w:del", "text_preview": "old wording", ...}
     ]}

# Accept the insertion (keep the new text, drop the tracked-change wrapper):
xml_patch(doc_id, op="unwrap", handle="h2")

# Reject the same insertion instead (undo it — remove the inserted content entirely):
xml_patch(doc_id, op="remove", handle="h2")

# Accept the deletion (the struck content really goes away):
xml_patch(doc_id, op="remove", handle="h4")

# Reject the deletion instead (restore the content — unwrap converts w:delText -> w:t
# automatically, since raw w:delText outside a w:del wrapper isn't valid content):
xml_patch(doc_id, op="unwrap", handle="h4")
```

**Empirically confirmed rendering behavior** (rendered a real doc with an inline `w:ins` and a
separate paragraph's `w:del` through the actual AppleScript pipeline and inspected both the
extracted PDF text and the rendered page image):
- The tracked **insertion** rendered inline, in red, underlined — clean and directly visible to a
  multimodal verifier.
- The tracked **deletion** did **not** render as inline strikethrough. Word's default "Simple
  Markup" review view rendered it as a **margin callout box** ("删除了: TRACKED_DELETE_MARKER…" —
  localized to the Word UI's language, Chinese on this machine — "Deleted: ...") connected to the
  deletion point by a dotted leader line, with a small red change-bar marker in the left margin.
  This is genuinely useful information for both the worked example (accept/reject logic above is
  correct regardless of how it renders) and for the verifier (§4's recommended `extra_note`
  addition exists specifically because of this finding — without it, a correctly-applied deletion
  could read to the verifier as "nothing changed in the body text", a false blocking failure).

## 7. Routing guidance — keeping the model from overusing XML mode

1. **Tool description framing** (§1): both tools open with "LOW-LEVEL ESCAPE HATCH" and name the
   cheaper alternatives by name — this text is present in every turn's tool list, not a one-time
   instruction that can be forgotten mid-session.
2. **`SYSTEM_EXECUTOR` ordering rule** (§5.1): explicit "prefer word_* for X; reach for xml_query/
   xml_patch only for Y" bullet, with Y spelled out as a closed, bounded list (footnotes, endnotes,
   fields, tracked-change accept/reject) rather than "anything creative you can think of."
3. **Escalation-ladder hint** — `loop.py:_apply_escalation`'s `SWITCH_STRATEGY` branch already
   fires after two same-signature failures on any tool; add one conditional clause: if the failing
   tool name is `xml_patch`, append "...or, if this is something word_edit_text/word_edit_style/
   word_insert_element can express, use that tool instead — it's simpler and already validated."
   This is keyed on **tool name**, a deterministic fact already available in scope, not on
   inferring what the model meant.
4. **Deliberately NOT building** (see §9): a separate opt-in gate (e.g. `enable_xml_mode`) that
   hides these tools until explicitly unlocked. There's no evidence yet that description+prompt
   nudging is insufficient; that's a Phase-2 escalation to reach for only if the live battery
   (§8) shows the model reaching for xml_patch on things word_* tools already cover.

## 8. Test plan

Unit (new): `tests/unit/test_word_xml_adapter.py`
- `parse_fragment`: with/without xmlns decls, multi-root rejection, malformed-XML rejection.
- `resolve_part`: body/header:N/footer:N/styles/numbering/settings resolve correctly; footnotes/
  endnotes raise the "call ensure_standard_part first" error before it's been created, and resolve
  correctly after.
- Each of the 7 ops against a small fixture `Document()` (mirroring the existing
  `tests/unit/test_word_adapter.py` fixture style): insert_before/after/append/replace/remove
  produce the expected tree shape (assert via `etree.tostring` diffs or child counts); `set_attr`
  add/change/remove; `unwrap` — including the `w:del`-specific `delText→t` rename, verified by
  asserting the unwrapped paragraph's python-docx `.text` now includes the restored content (proof
  it's valid `w:t`, not dangling `w:delText`).
- `ensure_standard_part("footnotes")` twice → second call returns `created: False`, no duplicate
  part/relationship (assert `len(doc.part.package.iter_parts())` unchanged on the second call).
- Full byte-round-trip assertions for all three worked examples (§6): build, save to `BytesIO()`,
  reopen, assert the expected parts/relationships/content-types/body content are present —
  essentially codifying the probes already run manually in this design session.
- Rollback correctness: feed a fragment that parses fine standalone but fails at the reopen-probe
  stage (simplest reliable trigger: monkeypatch the probe's `WordDocument(...)` call to raise) and
  assert (a) `XmlOpError` raised with the underlying exception named, (b) `session.to_bytes()`
  before and after are byte-identical, (c) `_xml_handles` was cleared.

Extend existing: `tests/unit/test_session_snapshot.py`
- A handle obtained via `register_xml_handle` before `undo()`/`restore_snapshot()` raises the
  "stale handle" `XmlOpError` when resolved afterward.

Extend existing: `tests/tools/test_direct_apply_no_double_exec.py`
- `xml_patch` produces exactly one snapshot per successful call (reusing the existing
  double-exec-guard test pattern already validated for other mutating tools).

Extend existing: `tests/loop/test_loop_offline.py` (uses `tests/fake_llm.py`)
- Scenario: `propose_plan` → `xml_query` → `xml_patch` → `update_plan(done)` against the offline
  fake-verifier harness; assert `visual_dirty`/`pending_ops` tracked `xml_patch` correctly and that
  the `extra_note` passed to the (faked) `verify_edit` call contains the tracked-change/footnote/
  TOC rendering-behavior hint from §4 when `xml_patch` was the tool used.

Live battery (not unit tests — matches the project's existing "13-scenario live battery" practice,
real Word + real independent verifier): add 3 new scenarios, one per worked example in §6, each
run end-to-end through the actual agent loop and judged by the real multimodal verifier. This is
the actual acceptance gate for Phase 1 — the unit tests above prove the plumbing is correct; only
a live run proves Word actually renders the result acceptably.

## 9. Phased rollout

**Phase 1 (this design → implementation):** everything in §1–§7 above. Ships `xml_query`/
`xml_patch`, the capability-note additions, and the two small `loop.py` hints. Gate: all new/
extended unit tests green, plus the 3 new live-battery scenarios passing.

**Phase 2 (only if warranted by Phase-1 live-battery evidence):**
- If the verifier still misjudges tracked-change/TOC/footnote renders despite the `extra_note`
  hint, iterate on `SYSTEM_VERIFIER` wording specifically (small, targeted prompt change).
- If real usage shows the model reaching for `xml_patch` on things word_* tools already cover
  despite the routing guidance in §7, consider the opt-in gate mentioned there — not before.
- If `StyleParams`'s silent-extra-field gap (§5) turns out to matter in practice, land the
  `extra="forbid"` hardening as its own small change.
- If accept/reject-all-tracked-changes turns out to be a common real request and one-call-per-node
  is a genuine pain point (not hypothetical), consider a convenience wrapper — but only after
  seeing that pain, not speculatively.

## What NOT to build, and why (all reasoned from the probes above, not guessed)

1. **A full ECMA-376/ISO 29500 XSD schema validator.** Empirically confirmed lxml and python-docx's
   OPC layer do zero schema checking, and empirically confirmed Word itself silently tolerates
   at least some schema-illegal structures without complaint — so a schema-clean fragment isn't
   sufficient for correctness and a schema-dirty one isn't always harmful. The real, already-built
   ground truth is the render+verify loop; a schema validator would be substantial engineering
   effort solving a question we don't actually need answered.
2. **A separate unpack-zip→edit-files→repack-zip pipeline** (the Anthropic-skill pattern). Would be
   a second, divergent mutation path that has to stay synchronized with the live in-memory
   `EditSession.doc` — pure complexity for no benefit, since `doc.part.package` already gives full
   in-memory access to every part.
3. **A generic "create any OPC part with any name/content-type" tool.** Wrong content-type or a
   malformed relationship graph is a real corruption class; `ensure_standard_part` is deliberately
   a narrow, hand-verified whitelist (footnotes, endnotes) rather than an open-ended primitive.
4. **A generic tag-rename / cross-part subtree-move primitive.** The one case we need
   (`w:delText`→`w:t` on rejected-deletion restore) is a narrow, named special case inside `unwrap`,
   not a general capability — keeps the validation surface small and each behavior individually
   justifiable.
5. **Forcing real TOC field recalculation via UI automation** (bringing Word to the foreground,
   sending F9/keystrokes via System Events). Empirically confirmed the current export path does
   not recalculate fields; forcing it would violate the documented "never activate — no focus
   stealing" zero-dialog invariant and reintroduce exactly the fragility that invariant exists to
   prevent. Ship best-effort static cached content and disclose the limitation instead.
6. **Hiding xml_query/xml_patch behind an opt-in gate by default.** Speculative complexity with no
   current evidence it's needed — see Phase 2.
7. **Any code that infers "what capability the model actually wanted" from matching keywords in
   its failed call.** Redirects are static, deterministic text (tool descriptions, fixed message
   suffixes) — never a heuristic classifier guessing intent from strings.


## Key decisions
- Two new tools only (xml_query read-only, xml_patch mutating) operating directly on the existing EditSession.doc.element/package — no separate unpack/repack pipeline; new files adapters/word_xml_adapter.py and tools/word_xml_tools.py mirror the existing word_adapter.py/word_tools.py split.
- xml_patch supports 7 ops: insert_before/insert_after/append/replace/remove/set_attr (as requested) plus unwrap (replace node with its own children — needed for tracked-change accept) and ensure_standard_part (idempotent, whitelisted footnotes/endnotes part bootstrap) — chosen because these two are the minimum needed to make the footnote and tracked-change worked examples actually work, without opening a general-purpose tag-rename or arbitrary-part-creation surface.
- Stable node handles are per-EditSession dict entries (handle string -> live lxml element reference), stored on EditSession itself and cleared exactly when EditSession.reload_from_bytes runs (i.e. on undo/restore_snapshot) — verified this is the only place doc identity changes, so it's the only invalidation boundary needed; handle resolution also checks getroottree() attachment to catch same-turn detachment (e.g. a handle into a subtree an earlier patch this turn already removed).
- Validation gate is three tiers: (1) fragment well-formedness via parse_fragment before any mutation, (2) a save-to-disk + independent re-open probe reusing EditSession.to_bytes/flush/reload_from_bytes verbatim (full rollback on failure), (3) the EXISTING real-Word-render + independent visual verifier, unmodified. Deliberately no XSD schema validation layer — empirically confirmed neither lxml/python-docx nor Word itself schema-check, so building one would solve a question that doesn't gate real correctness.
- Confirmed (not assumed) that SnapshotManager/undo already fully covers xml_patch, including brand-new OPC parts, because to_bytes()/reload_from_bytes() serialize the whole package via doc.save(), verified with an actual footnotes-part round-trip test.
- Confirmed xml_patch gets visual verification for free via the existing mutates=True + NON_VISUAL_TOOLS exclusion logic in agent/loop.py; recommended (not required) one extra_note addition in _verify_doc so the verifier isn't blindsided by empirically-observed Word rendering quirks (tracked deletions as margin callouts, not inline strikethrough).
- Capability-error redirects are static text only (tool descriptions + a fixed suffix on word_insert_element's existing generic-rejection message + a SYSTEM_EXECUTOR prompt bullet) — never code that pattern-matches the model's intent from a failed call, consistent with a firm 'semantic judgment must come from an LLM, never code heuristics' rule.
- TOC worked example ships with an explicit, disclosed limitation rather than a hacky workaround: empirically confirmed Word's AppleScript PDF export does NOT recalculate field results, so a freshly-authored TOC field will show whatever placeholder text we cache, not real page numbers, until the user manually updates fields in Word (F9) — forcing recalculation via UI automation was rejected as it would violate the documented no-focus-stealing zero-dialog invariant.

## Risks
- Word is lenient beyond both lxml's and python-docx's checks: a deliberately schema-illegal nesting (a <w:tbl> inside a <w:r>) passed every in-process validation gate AND rendered through the real AppleScript export with zero errors/dialogs, with Word simply silently ignoring the illegal content. This means a 'successful' xml_patch call plus a passing render/verify cycle does NOT guarantee the resulting document is OOXML-schema-clean — only that nothing visibly broke this time; latent structural garbage could surface later under a different Word version, LibreOffice, Google Docs, or a stricter downstream tool.
- Content inserted via xml_patch that isn't wrapped in a plain <w:r>/<w:t> that python-docx recognizes (e.g. text inside a raw w:ins wrapper) is invisible to word_read_content/get_structure/text_match selectors -- empirically confirmed python-docx's Paragraph.text returned '' for a paragraph whose only content was a hand-built w:ins run. An agent that forgets this and tries to verify/find its own xml_patch output via ordinary word_* read tools will wrongly conclude the edit didn't happen.
- TOC fields will render with stale/placeholder text (no real page numbers) in the exported PDF the verifier sees, since Word's PDF-export path does not recalculate fields -- this is a genuine, unavoidable-with-current-infrastructure limitation that must be disclosed to the user, not silently shipped as if it were a real computed TOC.
- Tracked deletions render as a margin callout box connected by a leader line (localized to the Word UI's language), not inline strikethrough, in the default 'Simple Markup' view -- if the independent multimodal verifier isn't told about this convention, it could misjudge a correctly-applied deletion as 'no visible change happened', causing false-blocking verification failures and unnecessary repair loops.
- Handle staleness bugs are a new failure class this feature introduces: a multi-call turn that removes/unwraps a node and then tries to reuse a handle into that now-detached subtree must be caught (designed for via a getroottree() attachment check), but this is new logic that hasn't been exercised under the project's full escalation-ladder/circuit-breaker machinery yet -- worth extra scrutiny in the live battery.
- Footnote-id and similar counter allocation (next available w:id) has to be computed by scanning existing part content since python-docx offers no next_id() helper for parts it doesn't model; an off-by-one or collision with reserved ids (-1/0 for separators) would silently corrupt the footnotes part in a way the in-process gates (per the schema-leniency finding above) likely would not catch.
- Performance: the save+reopen validation probe adds a full package serialize/deserialize per xml_patch call; fine for occasional footnote/TOC/tracked-change work, but a workflow that calls xml_patch many times in a tight loop (e.g. bulk-accepting dozens of tracked changes one node at a time) will accumulate that fixed per-call overhead -- not measured here, flagged as a Phase-2 concern only if it proves to matter in practice.
- Routing guidance (tool descriptions + prompt bullet + escalation hint) is a soft nudge, not an enforced constraint -- there is currently no hard mechanism stopping a model from reaching for xml_patch on something a word_* tool already handles; if live-battery usage shows this happening, the Phase-2 opt-in-gate idea would need to actually be built, which is a real possibility, not a certainty this design resolves.

## Estimated effort

