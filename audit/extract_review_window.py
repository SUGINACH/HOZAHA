#!/usr/bin/env python3
"""Extract a narrow, numbered manuscript window for the next editorial review."""
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET

DOCX = Path("כתב_יד_עם_הערות_עריכה_E0001-E1002_עם_השלמת_פער_E0022-E0105_סשן_20261106_P02520-P02554.docx")
OUT = Path("חלון-כתב-יד-להמשך-ביקורת-P02555.md")
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
def q(name): return f"{{{W}}}{name}"
def text(el):
    return "".join((x.text or "") for x in el.iter() if x.tag in (q("t"), q("delText")))

if not DOCX.exists():
    raise SystemExit(f"Missing cumulative review DOCX: {DOCX}")
with ZipFile(DOCX) as z:
    bad = z.testzip()
    if bad: raise SystemExit(f"Corrupt DOCX member: {bad}")
    doc = ET.fromstring(z.read("word/document.xml"))
    body = doc.find(".//" + q("body"))
    paras = body.findall(q("p")) if body is not None else []
    if len(paras) != 4603: raise SystemExit(f"Unexpected paragraph count: {len(paras)}")
    comments = {}
    if "word/comments.xml" in z.namelist():
        root = ET.fromstring(z.read("word/comments.xml"))
        for c in root.findall(".//" + q("comment")):
            comments[c.get(q("id"), "?")] = text(c)
    start, end = 2548, 2595
    if end > len(paras): end = len(paras)
    rows = [
        "# קטעי כתב־יד להמשך ביקורת — סביב P02555",
        "",
        f"- קובץ מקור: {DOCX.name}",
        f"- מספר פסקאות במסמך: {len(paras)}",
        "- המספור P מתייחס למספר פסקת כתב־היד; מספר פסקה בקובץ Word הוא P+1.",
        "- הטקסט להלן חולץ לקריאה בלבד; לא נעשה שינוי בקובץ Word.",
        "",
    ]
    for n in range(start, end + 1):
        p = paras[n - 1]
        label = f"P{n-1:05d}"
        body_text = text(p)
        rows += [f"## {label} (פסקה {n})", "", body_text or "_[פסקה ריקה]_", ""]
        active = {}
        snippets = {}
        for child in list(p):
            if child.tag == q("commentRangeStart"):
                active[child.get(q("id"), "?")] = []
            elif child.tag == q("commentRangeEnd"):
                cid = child.get(q("id"), "?")
                snippets[cid] = "".join(active.pop(cid, []))
            else:
                part = text(child)
                if part:
                    for pieces in active.values(): pieces.append(part)
        for cid, snippet in snippets.items():
            comment = comments.get(cid, "[טקסט הערה חסר]")
            rows += [f"> הערה קיימת (Word ID {cid}; עוגן: {snippet}): {comment}", ""]
OUT.write_text("\n".join(rows), encoding="utf-8")
print(f"Wrote {OUT} with paragraphs P{start-1:05d}–P{end-1:05d} ({end-start+1} paragraphs).")
