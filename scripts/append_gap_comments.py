#!/usr/bin/env python3
"""Append a reviewed gap batch as real Word comments without changing manuscript text."""
from __future__ import annotations
import json
import re
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from lxml import etree

ROOT = Path(".")
BASE = ROOT / "כתב_יד_עם_הערות_עריכה_E0001-E1002_מתוקן_גופן.docx"
BATCH = ROOT / "scripts" / "gap_batch_E0022_E0054.json"
OUT = ROOT / "כתב_יד_עם_הערות_עריכה_E0001-E1002_עם_השלמת_פער_ראשונה_E0022-E0054.docx"
REPORT = ROOT / "דוח_השלמת_פער_הערות_E0022-E0054.md"

WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XMLNS = "http://www.w3.org/XML/1998/namespace"
NS = {"w": WNS}
def q(local: str) -> str:
    return f"{{{WNS}}}{local}"

if not BASE.exists():
    raise SystemExit(f"Missing expected base DOCX: {BASE}")
if not BATCH.exists():
    raise SystemExit(f"Missing review batch: {BATCH}")

batch = json.loads(BATCH.read_text(encoding="utf-8"))
if not batch:
    raise SystemExit("Review batch is empty")
seqs = [int(x["seq"]) for x in batch]
if len(seqs) != len(set(seqs)):
    raise SystemExit("Duplicate E-sequence IDs in batch")
if seqs != list(range(min(seqs), max(seqs)+1)):
    raise SystemExit("Batch E-sequence range is not continuous")

with zipfile.ZipFile(BASE, "r") as zin:
    names = set(zin.namelist())
    if "word/document.xml" not in names or "word/comments.xml" not in names:
        raise SystemExit("Base DOCX has no document/comments part")
    document = etree.fromstring(zin.read("word/document.xml"))
    comments = etree.fromstring(zin.read("word/comments.xml"))
    body = document.find(".//" + q("body"))
    if body is None:
        raise SystemExit("Missing document body")
    paragraphs = body.findall(q("p"))
    if not paragraphs:
        raise SystemExit("No top-level manuscript paragraphs found")

    existing_comments = comments.findall(q("comment"))
    numeric_ids = []
    labels = []
    for c in existing_comments:
        cid = c.get(q("id"))
        if cid is not None and cid.isdigit():
            numeric_ids.append(int(cid))
        ct = "".join(c.xpath(".//w:t/text()", namespaces=NS))
        labels += re.findall(r"E-(\d{4})", ct)
    next_word_id = max(numeric_ids, default=-1) + 1
    existing_seqs = {int(s) for s in labels}
    if existing_seqs.intersection(seqs):
        raise SystemExit("One or more E-sequence labels already exist in base document")

    stamp = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    added_rows = []
    for item in batch:
        pnum = int(item["p"])
        if pnum < 1 or pnum > len(paragraphs):
            raise SystemExit(f"Paragraph {pnum} outside document (1-{len(paragraphs)})")
        para = paragraphs[pnum-1]
        if para.xpath(".//w:commentRangeStart | .//w:commentRangeEnd | .//w:commentReference", namespaces=NS):
            raise SystemExit(f"Paragraph {pnum} already has comment markup; refusing overlap")

        wid = str(next_word_id)
        next_word_id += 1
        ppr = para.find(q("pPr"))
        start_at = 1 if ppr is not None else 0
        para.insert(start_at, etree.Element(q("commentRangeStart"), attrib={q("id"): wid}))
        para.append(etree.Element(q("commentRangeEnd"), attrib={q("id"): wid}))

        ref_run = etree.Element(q("r"))
        ref_rpr = etree.SubElement(ref_run, q("rPr"))
        etree.SubElement(ref_rpr, q("rStyle")).set(q("val"), "CommentReference")
        etree.SubElement(ref_rpr, q("rtl")).set(q("val"), "1")
        etree.SubElement(ref_rpr, q("vertAlign")).set(q("val"), "superscript")
        etree.SubElement(ref_run, q("commentReference")).set(q("id"), wid)
        para.append(ref_run)

        comment = etree.SubElement(comments, q("comment"))
        comment.set(q("id"), wid)
        comment.set(q("author"), "OpenAI – ביקורת עריכה")
        comment.set(q("initials"), "AI")
        comment.set(q("date"), stamp)
        cp = etree.SubElement(comment, q("p"))
        cppr = etree.SubElement(cp, q("pPr"))
        etree.SubElement(cppr, q("pStyle")).set(q("val"), "CommentText")
        etree.SubElement(cppr, q("bidi")).set(q("val"), "1")
        etree.SubElement(cppr, q("jc")).set(q("val"), "right")
        cr = etree.SubElement(cp, q("r"))
        crpr = etree.SubElement(cr, q("rPr"))
        fonts = etree.SubElement(crpr, q("rFonts"))
        for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
            fonts.set(q(attr), "Arial")
        etree.SubElement(crpr, q("rtl")).set(q("val"), "1")
        lang = etree.SubElement(crpr, q("lang"))
        lang.set(q("val"), "he-IL")
        lang.set(q("bidi"), "he-IL")
        etree.SubElement(crpr, q("sz")).set(q("val"), "20")
        t = etree.SubElement(cr, q("t"))
        t.set(f"{{{XMLNS}}}space", "preserve")
        t.text = (f"E-{int(item['seq']):04d} | חומרה: {item['severity']} | "
                  f"קטגוריה: {item['category']}. {item['body']}")
        added_rows.append((int(item["seq"]), pnum, wid, item["severity"], item["category"]))

    rels = etree.fromstring(zin.read("word/_rels/document.xml.rels"))
    rel_ok = any(rel.get("Type", "").endswith("/comments") and rel.get("Target") == "comments.xml" for rel in rels)
    if not rel_ok:
        raise SystemExit("Comments relationship missing in base document")
    content_types = etree.fromstring(zin.read("[Content_Types].xml"))
    ct_ok = any(el.get("PartName") == "/word/comments.xml" for el in content_types)
    if not ct_ok:
        raise SystemExit("Comments content-type override missing in base document")

    for seq, pnum, wid, sev, cat in added_rows:
        starts = document.xpath(f'.//w:commentRangeStart[@w:id="{wid}"]', namespaces=NS)
        ends = document.xpath(f'.//w:commentRangeEnd[@w:id="{wid}"]', namespaces=NS)
        refs = document.xpath(f'.//w:commentReference[@w:id="{wid}"]', namespaces=NS)
        defs = comments.xpath(f'./w:comment[@w:id="{wid}"]', namespaces=NS)
        if not (len(starts) == len(ends) == len(refs) == len(defs) == 1):
            raise SystemExit(f"Comment linkage integrity failed for E-{seq:04d}")

    doc_xml = etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone=True)
    comments_xml = etree.tostring(comments, xml_declaration=True, encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(OUT, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name in zin.namelist():
            if name == "word/document.xml":
                zout.writestr(name, doc_xml)
            elif name == "word/comments.xml":
                zout.writestr(name, comments_xml)
            else:
                zout.writestr(name, zin.read(name))

with zipfile.ZipFile(OUT, "r") as check:
    bad = check.testzip()
    if bad:
        raise SystemExit(f"Corrupt DOCX ZIP member: {bad}")
    d2 = etree.fromstring(check.read("word/document.xml"))
    c2 = etree.fromstring(check.read("word/comments.xml"))
    for seq, pnum, wid, sev, cat in added_rows:
        if len(d2.xpath(f'.//w:commentRangeStart[@w:id="{wid}"]', namespaces=NS)) != 1:
            raise SystemExit(f"Final verification failed at E-{seq:04d}")
    all_labels = []
    for c in c2.findall(q("comment")):
        all_labels.extend(re.findall(r"E-(\d{4})", "".join(c.xpath(".//w:t/text()", namespaces=NS))))
    if len(all_labels) != len(set(all_labels)):
        raise SystemExit("Duplicate E-labels in final document")

report = [
    "# דוח השלמת פער הערות – אצווה ראשונה",
    "",
    f"- קובץ בסיס: {BASE.name}",
    f"- קובץ פלט: {OUT.name}",
    f"- אצווה שנוספה: E-{min(seqs):04d} עד E-{max(seqs):04d} ({len(added_rows)} הערות).",
    f"- מספר הפסקאות במסמך: {len(paragraphs)}.",
    "- כל הערה נוספה כהערת Word אמיתית, עם עוגן בפסקה המתאימה.",
    "- גוף הטקסט לא נערך; נוספו סימוני עוגן והערות בלבד.",
    "- לכל הערה חדשה הוגדרו Arial, כיווניות RTL ושפת he-IL.",
    "- קובץ זה אינו מציג את הפער כולו כסגור: עדיין חסרות ההערות E-0055 עד E-0767, והן דורשות ביקורת ושחזור תוכני.",
    "",
    "## מיקומי ההערות באצווה",
    "",
    "| מזהה | פסקה (מספור המסמך) | חומרה | קטגוריה |",
    "|---|---:|---|---|",
]
for seq, pnum, wid, sev, cat in added_rows:
    report.append(f"| E-{seq:04d} | {pnum} | {sev} | {cat} |")
REPORT.write_text("\n".join(report) + "\n", encoding="utf-8")
print(f"Created {OUT} with {len(added_rows)} comments, E-{min(seqs):04d}–E-{max(seqs):04d}.")
print(f"Paragraph count: {len(paragraphs)}. Comment count after insertion: {len(comments.findall(q('comment')))}.")
print(f"Report: {REPORT}")
