#!/usr/bin/env python3
"""Append all reconstructed gap-comment batches to the fixed Hebrew DOCX base."""
from __future__ import annotations
import json
import re
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from lxml import etree

ROOT = Path(".")
BASE = ROOT / "כתב_יד_עם_הערות_עריכה_E0001-E1002_מתוקן_גופן.docx"
SCRIPTS = ROOT / "scripts"
OUT = ROOT / "כתב_יד_עם_הערות_עריכה_E0001-E1002_עם_השלמת_פער_E0022-E0105.docx"
REPORT = ROOT / "דוח_השלמת_פער_הערות_E0022-E0105.md"

WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XMLNS = "http://www.w3.org/XML/1998/namespace"
NS = {"w": WNS}

def q(local: str) -> str:
    return f"{{{WNS}}}{local}"

if not BASE.exists():
    raise SystemExit(f"Missing expected base DOCX: {BASE}")

batch_paths = sorted(
    SCRIPTS.glob("gap_batch_E*.json"),
    key=lambda p: int(re.search(r"_E(\d+)_E(\d+)\.json$", p.name).group(1))
        if re.search(r"_E(\d+)_E(\d+)\.json$", p.name) else 10**9
)
if not batch_paths:
    raise SystemExit("No gap batches found")

all_items = []
batch_ranges = []
for path in batch_paths:
    match = re.search(r"_E(\d+)_E(\d+)\.json$", path.name)
    if not match:
        raise SystemExit(f"Unexpected batch filename: {path.name}")
    batch = json.loads(path.read_text(encoding="utf-8"))
    if not batch:
        raise SystemExit(f"Empty batch: {path.name}")
    seqs = [int(x["seq"]) for x in batch]
    if len(seqs) != len(set(seqs)):
        raise SystemExit(f"Duplicate E-sequence IDs in {path.name}")
    if seqs != list(range(min(seqs), max(seqs) + 1)):
        raise SystemExit(f"Non-contiguous sequence range in {path.name}")
    if min(seqs) != int(match.group(1)) or max(seqs) != int(match.group(2)):
        raise SystemExit(f"Filename/range mismatch in {path.name}")
    batch_ranges.append((min(seqs), max(seqs), path.name, len(batch)))
    all_items.extend(batch)

all_seqs = [int(x["seq"]) for x in all_items]
if len(all_seqs) != len(set(all_seqs)):
    raise SystemExit("Duplicate E-sequence IDs across batches")
all_seqs_sorted = sorted(all_seqs)
if all_seqs_sorted != list(range(min(all_seqs_sorted), max(all_seqs_sorted) + 1)):
    raise SystemExit("Combined E-sequence range has gaps")
if all_seqs_sorted[0] != 22:
    raise SystemExit("Gap batches must begin at E-0022")
all_items.sort(key=lambda x: int(x["seq"]))

with zipfile.ZipFile(BASE, "r") as zin:
    if "word/document.xml" not in zin.namelist() or "word/comments.xml" not in zin.namelist():
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
        comment_text = "".join(c.xpath(".//w:t/text()", namespaces=NS))
        labels.extend(re.findall(r"E-(\d{4})", comment_text))
    next_word_id = max(numeric_ids, default=-1) + 1
    existing_seqs = {int(s) for s in labels}
    if existing_seqs.intersection(all_seqs):
        raise SystemExit("One or more E-sequence labels already exist in base document")

    stamp = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    added_rows = []
    for item in all_items:
        pnum = int(item["p"])
        if pnum < 1 or pnum > len(paragraphs):
            raise SystemExit(f"Paragraph {pnum} outside document (1-{len(paragraphs)})")
        para = paragraphs[pnum - 1]
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
        t.text = (
            f"E-{int(item['seq']):04d} | חומרה: {item['severity']} | "
            f"קטגוריה: {item['category']}. {item['body']}"
        )
        added_rows.append((int(item["seq"]), pnum, wid, item["severity"], item["category"]))

    rels = etree.fromstring(zin.read("word/_rels/document.xml.rels"))
    if not any(
        rel.get("Type", "").endswith("/comments") and rel.get("Target") == "comments.xml"
        for rel in rels
    ):
        raise SystemExit("Comments relationship missing in base document")
    content_types = etree.fromstring(zin.read("[Content_Types].xml"))
    if not any(el.get("PartName") == "/word/comments.xml" for el in content_types):
        raise SystemExit("Comments content-type override missing in base document")

    for seq, pnum, wid, severity, category in added_rows:
        checks = (
            document.xpath(f'.//w:commentRangeStart[@w:id="{wid}"]', namespaces=NS),
            document.xpath(f'.//w:commentRangeEnd[@w:id="{wid}"]', namespaces=NS),
            document.xpath(f'.//w:commentReference[@w:id="{wid}"]', namespaces=NS),
            comments.xpath(f'./w:comment[@w:id="{wid}"]', namespaces=NS),
        )
        if any(len(part) != 1 for part in checks):
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
    for seq, pnum, wid, severity, category in added_rows:
        if len(d2.xpath(f'.//w:commentRangeStart[@w:id="{wid}"]', namespaces=NS)) != 1:
            raise SystemExit(f"Final verification failed at E-{seq:04d}")
    final_labels = []
    for c in c2.findall(q("comment")):
        final_labels.extend(re.findall(r"E-(\d{4})", "".join(c.xpath(".//w:t/text()", namespaces=NS))))
    final_ints = [int(s) for s in final_labels]
    if len(final_ints) != len(set(final_ints)):
        raise SystemExit("Duplicate E-labels in final document")
    total_comments = len(c2.findall(q("comment")))

report_lines = [
    "# דוח השלמת פער הערות – גרסה מצטברת",
    "",
    f"- קובץ בסיס: {BASE.name}",
    f"- קובץ פלט: {OUT.name}",
    f"- מספור ההערות החדשות שנוסף: E-{min(all_seqs_sorted):04d} עד E-{max(all_seqs_sorted):04d} ({len(added_rows)} הערות).",
    f"- מספר ההערות הכולל בקובץ: {total_comments}.",
    f"- מספר הפסקאות במסמך: {len(paragraphs)}.",
    "- כל הערה חדשה נוספה כהערת Word אמיתית, עם טווח עיגון בפסקה המתאימה.",
    "- גוף כתב־היד לא נערך; נוספו סימוני עוגן והערות בלבד.",
    "- לכל ההערות החדשות הוגדרו Arial, כיווניות RTL ושפת he-IL.",
    "- הקובץ עדיין אינו סוגר את הפער ההיסטורי כולו: ההערות E-0106 עד E-0767 טרם נכתבו ונבדקו.",
    "",
    "## אצוות שנכללו",
    "",
    "| קובץ אצווה | טווח | מספר הערות |",
    "|---|---:|---:|",
]
for start, end, name, count in batch_ranges:
    report_lines.append(f"| {name} | E-{start:04d}–E-{end:04d} | {count} |")
report_lines += [
    "",
    "## מיקומי ההערות",
    "",
    "| מזהה | פסקה (מספור המסמך) | חומרה | קטגוריה |",
    "|---|---:|---|---|",
]
for seq, pnum, wid, severity, category in added_rows:
    report_lines.append(f"| E-{seq:04d} | {pnum} | {severity} | {category} |")
REPORT.write_text("\n".join(report_lines) + "\n", encoding="utf-8")
print(f"Created {OUT} with {len(added_rows)} new comments, E-{min(all_seqs_sorted):04d}–E-{max(all_seqs_sorted):04d}.")
print(f"Total comments: {total_comments}. Paragraph count: {len(paragraphs)}.")
print(f"Report: {REPORT}")
