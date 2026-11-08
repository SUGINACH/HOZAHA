#!/usr/bin/env python3
"""Extract the next manuscript review window from the newest cumulative DOCX."""
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET
import hashlib, json

DOCX=Path("כתב_יד_עם_הערות_עריכה_E0001-E1002_עם_השלמת_פער_E0022-E0105_סשן_20261108_P02646-P02695.docx")
OUT=Path("חלון-כתב-יד-להמשך-ביקורת-P02696.md")
W="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
def q(name): return f"{{{W}}}{name}"
def text(el):
    return "".join((x.text or "") for x in el.iter() if x.tag in (q("t"),q("delText")))

if not DOCX.is_file():
    raise SystemExit(f"Missing cumulative review DOCX: {DOCX}")
with ZipFile(DOCX) as z:
    bad=z.testzip()
    if bad: raise SystemExit(f"Corrupt DOCX member: {bad}")
    names=set(z.namelist())
    if not {"word/document.xml","word/comments.xml"}.issubset(names):
        raise SystemExit("Missing document.xml or comments.xml")
    doc=ET.fromstring(z.read("word/document.xml"))
    body=doc.find(".//"+q("body"))
    paras=body.findall(q("p")) if body is not None else []
    if len(paras)!=4603:
        raise SystemExit(f"Unexpected paragraph count: {len(paras)}")
    ptexts=[text(p) for p in paras]
    digest=hashlib.sha256(json.dumps(ptexts,ensure_ascii=False).encode()).hexdigest()
    expected="b945de862e3f90a06a6129aba60f8af6fa10e0bdab68356de9b00c53cc07b97d"
    if digest!=expected: raise SystemExit(f"Unexpected manuscript-body hash: {digest}")
    cr=ET.fromstring(z.read("word/comments.xml"))
    comment_map={}
    for c in cr.findall(".//"+q("comment")):
        comment_map[c.get(q("id"),"?")]=text(c)
    if len(comment_map)!=549:
        raise SystemExit(f"Unexpected Word comment count: {len(comment_map)}")
    start,end=2691,2746  # P02690–P02745; Word paragraph index is P+1.
    rows=[
        "# חלון כתב־יד להמשך ביקורת — P02696 ואילך","",
        f"- קובץ מקור: {DOCX.name}",
        f"- מספר פסקאות במסמך: {len(paras)}",
        f"- מספר הערות Word: {len(comment_map)}",
        f"- SHA-256 של גוף כתב־היד: {digest}",
        "- המספור P הוא מספר פסקת כתב־היד; מספר הפסקה ב־Word הוא P+1.",
        "- חולץ לקריאה בלבד מתוך גרסת DOCX העדכנית; קובץ המקור לא שונה.",
        "- החלון כולל חפיפה P02690–P02695 לצורך שמירת ההקשר, ואת P02696–P02745 להמשך.",
        "",
    ]
    for n in range(start,end+1):
        p=paras[n-1]
        label=f"P{n-1:05d}"
        rows += [f"## {label} (פסקה {n})","",text(p) or "_[פסקה ריקה]_",""]
        active={}
        snippets={}
        for child in list(p):
            if child.tag==q("commentRangeStart"):
                active[child.get(q("id"),"?")]=[]
            elif child.tag==q("commentRangeEnd"):
                cid=child.get(q("id"),"?")
                snippets[cid]="".join(active.pop(cid,[]))
            else:
                part=text(child)
                if part:
                    for parts in active.values(): parts.append(part)
        for cid,snippet in snippets.items():
            rows += [f"> הערה קיימת (Word ID {cid}; עוגן: {snippet}): {comment_map.get(cid,'[טקסט הערה חסר]')}",""]
OUT.write_text("\n".join(rows)+"\n",encoding="utf-8")
print(f"Wrote {OUT}: P{start-1:05d}–P{end-1:05d}, {end-start+1} paragraphs; 549 comments; body hash verified.")
