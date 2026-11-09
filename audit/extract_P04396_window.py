#!/usr/bin/env python3
"""Read-only extraction of P04396–P04445 from the latest cumulative reviewed DOCX."""
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET
import hashlib, json

DOCX = Path("כתב_יד_S027_P04346-P04395.docx")
OUT = Path("review-windows/חלון-כתב-יד-להמשך-ביקורת-P04396.md")
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def q(name):
    return f"{{{W}}}{name}"

def text(el):
    return "".join((x.text or "") for x in el.iter() if x.tag in (q("t"), q("delText")))

if not DOCX.is_file():
    raise SystemExit(f"Missing latest cumulative review DOCX: {DOCX}")

with ZipFile(DOCX) as z:
    bad = z.testzip()
    if bad:
        raise SystemExit(f"Corrupt DOCX member: {bad}")
    names = set(z.namelist())
    if not {"word/document.xml", "word/comments.xml"}.issubset(names):
        raise SystemExit("Missing document.xml or comments.xml")
    doc = ET.fromstring(z.read("word/document.xml"))
    body = doc.find(".//" + q("body"))
    paras = body.findall(q("p")) if body is not None else []
    if len(paras) != 4603:
        raise SystemExit(f"Unexpected paragraph count: {len(paras)}")
    ptexts = [text(p) for p in paras]
    digest = hashlib.sha256(json.dumps(ptexts, ensure_ascii=False).encode()).hexdigest()
    expected = "b945de862e3f90a06a6129aba60f8af6fa10e0bdab68356de9b00c53cc07b97d"
    if digest != expected:
        raise SystemExit(f"Unexpected manuscript-body hash: {digest}")
    root = ET.fromstring(z.read("word/comments.xml"))
    comment_map = {c.get(q("id"), "?"): text(c) for c in root.findall(".//" + q("comment"))}
    if len(comment_map) != 734:
        raise SystemExit(f"Unexpected Word comment count: {len(comment_map)}")
    start, end = 4391, 4446  # P04390–P04445; Word paragraph index is P+1.
    rows = [
        "# חלון כתב־יד להמשך ביקורת — P04396 ואילך", "",
        f"- קובץ מקור: {DOCX.name}",
        f"- מספר פסקאות במסמך: {len(paras)}",
        f"- מספר הערות Word: {len(comment_map)}",
        f"- SHA-256 של גוף כתב־היד: {digest}",
        "- המספור P הוא מספר פסקת כתב־היד; מספר הפסקה ב־Word הוא P+1.",
        "- חולץ לקריאה בלבד מן המסמך המצטבר לאחר הוספת ההערות S-20261109-136 עד S-20261109-142; קובץ המקור לא שונה.",
        "- החלון כולל חפיפה P04390–P04395 ואת P04396–P04445 להמשך היחידה.", ""
    ]
    for n in range(start, end + 1):
        p = paras[n - 1]
        rows += [f"## P{n-1:05d} (פסקה {n})", "", text(p) or "_[פסקה ריקה]_", ""]
        active, snippets = {}, {}
        for child in list(p):
            if child.tag == q("commentRangeStart"):
                active[child.get(q("id"), "?")] = []
            elif child.tag == q("commentRangeEnd"):
                cid = child.get(q("id"), "?")
                snippets[cid] = "".join(active.pop(cid, []))
            else:
                part = text(child)
                if part:
                    for parts in active.values():
                        parts.append(part)
        for cid, snippet in snippets.items():
            rows += [f"> הערה קיימת (Word ID {cid}; עוגן: {snippet}): {comment_map.get(cid, '[טקסט הערה חסר]')}", ""]
OUT.parent.mkdir(parents=True, exist_ok=True)
OUT.write_text("\n".join(rows) + "\n", encoding="utf-8")
print(f"Wrote {OUT}: P{start-1:05d}–P{end-1:05d}, {end-start+1} paragraphs; 734 comments; body hash verified.")
