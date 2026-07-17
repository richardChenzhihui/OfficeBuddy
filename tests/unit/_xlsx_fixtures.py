"""Zip-surgery helpers for Excel fidelity-guard tests.

The guard detects a part NAME disappearing on a round trip; it does not
validate the part's internal correctness. So structurally-plausible injected
parts are sufficient and fully reproducible in pytest — no binary fixtures.
"""
import shutil
import zipfile
from pathlib import Path

_FOREIGN_PARTS = {
    "threaded_comment": (
        "xl/threadedComments/threadedComment1.xml",
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">'
        '<threadedComment ref="A1" id="{1}" personId="{2}"><text>probe</text></threadedComment>'
        "</ThreadedComments>",
    ),
    "person": (
        "xl/persons/person.xml",
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"/>',
    ),
    "slicer": (
        "xl/slicers/slicer1.xml",
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<slicers xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"/>',
    ),
    "custom_xml": (
        "customXml/item1.xml",
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><probe>data</probe>',
    ),
}


def inject_foreign_parts(xlsx_path: Path, kinds: list) -> list:
    """Append foreign parts to an existing xlsx. Returns the injected names."""
    tmp = xlsx_path.with_suffix(".tmp.xlsx")
    shutil.copy2(xlsx_path, tmp)
    injected = []
    with zipfile.ZipFile(xlsx_path) as src, zipfile.ZipFile(
        tmp, "w", zipfile.ZIP_DEFLATED
    ) as dst:
        for item in src.infolist():
            dst.writestr(item, src.read(item.filename))
        for kind in kinds:
            name, xml = _FOREIGN_PARTS[kind]
            dst.writestr(name, xml)
            injected.append(name)
    tmp.replace(xlsx_path)
    return injected
