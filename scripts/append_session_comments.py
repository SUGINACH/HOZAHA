#!/usr/bin/env python3
"""Append session Word comments; never alter manuscript paragraph text."""
from __future__ import annotations
import hashlib, json, re, zipfile
from datetime import datetime, timezone
from pathlib import Path
from lxml import etree
ROOT=Path("."); CFG=ROOT/"scripts/session_comments.json"
W="http://schemas.openxmlformats.org/wordprocessingml/2006/main"; X="http://www.w3.org/XML/1998/namespace"; NS={"w":W}
def q(x): return f"{{{W}}}{x}"
def txt(p): return "".join((n.text or "") for n in p.iter() if n.tag in (q("t"),q("delText")))
def sha(ps): return hashlib.sha256(json.dumps([txt(p) for p in ps],ensure_ascii=False).encode()).hexdigest()
cfg=json.loads(CFG.read_text(encoding="utf-8")); sid=str(cfg["session_id"])
if not re.fullmatch(r"[A-Za-z0-9-]+",sid): raise SystemExit("Unsafe session id")
base=ROOT/cfg["base_docx"]; out=ROOT/cfg["output_docx"]; report=ROOT/cfg["report_path"]; items=cfg["comments"]
next_section=str(cfg.get("next_section","לא הוגדרה נקודת המשך בתצורת הסשן"))
if not base.exists() or not items: raise SystemExit("Missing base DOCX or session comments")
labels=[x["label"] for x in items]; nums=[int(x["p"]) for x in items]
if len(set(labels))!=len(labels) or len(set(nums))!=len(nums): raise SystemExit("Duplicate comment labels or paragraph targets")
if any(not re.fullmatch(r"S-\d{8}-\d{2,}", label) for label in labels): raise SystemExit("Invalid session comment label format")
with zipfile.ZipFile(base) as z:
    required={"word/document.xml","word/comments.xml","word/_rels/document.xml.rels","[Content_Types].xml"}
    if not required.issubset(z.namelist()): raise SystemExit("Base DOCX lacks Word comments package parts")
    doc=etree.fromstring(z.read("word/document.xml")); comments=etree.fromstring(z.read("word/comments.xml"))
    body=doc.find(".//"+q("body")); ps=body.findall(q("p")) if body is not None else []
    expected=int(cfg.get("expected_paragraph_count",len(ps)))
    if len(ps)!=expected: raise SystemExit(f"Paragraph count mismatch: {len(ps)} != {expected}")
    before=sha(ps); old=comments.findall(q("comment"))
    oldids=[int(c.get(q("id"))) for c in old if (c.get(q("id")) or "").isdigit()]
    oldtexts=["".join(c.xpath(".//w:t/text()",namespaces=NS)) for c in old]
    for label in labels:
        if any(label in t for t in oldtexts): raise SystemExit("Session label already exists: "+label)
    wid=max(oldids,default=-1)+1; stamp=datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00","Z")
    added=[]
    for item in items:
        n=int(item["p"])
        if not 1<=n<=len(ps): raise SystemExit("Paragraph target outside DOCX: "+str(n))
        p=ps[n-1]
        if not txt(p).startswith(item["expected_prefix"]): raise SystemExit("Paragraph text mismatch at "+str(n))
        if p.xpath(".//w:commentRangeStart | .//w:commentRangeEnd | .//w:commentReference",namespaces=NS):
            raise SystemExit("Paragraph already has comment anchors: "+str(n))
        runs=[r for r in p.findall(q("r")) if any(a.tag in (q("t"),q("delText")) and (a.text or "") for a in r.iter())]
        if not runs: raise SystemExit("No text runs in paragraph "+str(n))
        cid=str(wid); wid+=1
        start=etree.Element(q("commentRangeStart"),attrib={q("id"):cid}); end=etree.Element(q("commentRangeEnd"),attrib={q("id"):cid})
        p.insert(list(p).index(runs[0]),start); p.insert(list(p).index(runs[-1])+1,end)
        ref=etree.Element(q("r")); rp=etree.SubElement(ref,q("rPr"))
        etree.SubElement(rp,q("rStyle")).set(q("val"),"CommentReference")
        etree.SubElement(rp,q("rtl")).set(q("val"),"1")
        etree.SubElement(rp,q("vertAlign")).set(q("val"),"superscript")
        etree.SubElement(ref,q("commentReference")).set(q("id"),cid); p.insert(list(p).index(end)+1,ref)
        c=etree.SubElement(comments,q("comment"))
        c.set(q("id"),cid); c.set(q("author"),"OpenAI – ביקורת עריכה"); c.set(q("initials"),"AI"); c.set(q("date"),stamp)
        cp=etree.SubElement(c,q("p")); pp=etree.SubElement(cp,q("pPr"))
        etree.SubElement(pp,q("pStyle")).set(q("val"),"CommentText"); etree.SubElement(pp,q("bidi")).set(q("val"),"1"); etree.SubElement(pp,q("jc")).set(q("val"),"right")
        r=etree.SubElement(cp,q("r")); rp=etree.SubElement(r,q("rPr")); rf=etree.SubElement(rp,q("rFonts"))
        for k in ("ascii","hAnsi","eastAsia","cs"): rf.set(q(k),"Arial")
        etree.SubElement(rp,q("rtl")).set(q("val"),"1")
        lang=etree.SubElement(rp,q("lang")); lang.set(q("val"),"he-IL"); lang.set(q("bidi"),"he-IL")
        etree.SubElement(rp,q("sz")).set(q("val"),"20")
        t=etree.SubElement(r,q("t")); t.set(f"{{{X}}}space","preserve")
        t.text=f"{item['label']} | חומרה: {item['severity']} | קטגוריה: {item['category']}. {item['body']}"
        added.append({"label":item["label"],"p":n,"id":cid,"severity":item["severity"],"category":item["category"]})
    rels=etree.fromstring(z.read("word/_rels/document.xml.rels"))
    if not any(a.get("Type","").endswith("/comments") and a.get("Target")=="comments.xml" for a in rels): raise SystemExit("Missing comments relationship")
    types=etree.fromstring(z.read("[Content_Types].xml"))
    if not any(a.get("PartName")=="/word/comments.xml" for a in types): raise SystemExit("Missing comments content-type override")
    dx=etree.tostring(doc,xml_declaration=True,encoding="UTF-8",standalone=True)
    cx=etree.tostring(comments,xml_declaration=True,encoding="UTF-8",standalone=True)
    with zipfile.ZipFile(out,"w",compression=zipfile.ZIP_DEFLATED) as dst:
        for name in z.namelist(): dst.writestr(name,dx if name=="word/document.xml" else cx if name=="word/comments.xml" else z.read(name))
with zipfile.ZipFile(out) as z:
    bad=z.testzip()
    if bad: raise SystemExit("DOCX ZIP error: "+bad)
    d=etree.fromstring(z.read("word/document.xml")); croot=etree.fromstring(z.read("word/comments.xml"))
    ps2=d.find(".//"+q("body")).findall(q("p")); cs=croot.findall(q("comment")); ctexts=["".join(c.xpath(".//w:t/text()",namespaces=NS)) for c in cs]
    for a in added:
        p=ps2[a["p"]-1]
        for tag in ("commentRangeStart","commentRangeEnd","commentReference"):
            if len([x for x in p.iter() if x.tag==q(tag) and x.get(q("id"))==a["id"]])!=1: raise SystemExit("Anchor integrity failure: "+a["label"])
        if sum(a["label"] in s for s in ctexts)!=1: raise SystemExit("Comment text integrity failure: "+a["label"])
        c=next(x for x in cs if x.get(q("id"))==a["id"]); rs=c.xpath(".//w:r",namespaces=NS)
        if len(rs)!=1: raise SystemExit("Unexpected comment run count: "+a["label"])
        rp=rs[0].find(q("rPr")); rf=rp.find(q("rFonts")) if rp is not None else None
        rtl=rp.find(q("rtl")) if rp is not None else None; lang=rp.find(q("lang")) if rp is not None else None
        if rf is None or any(rf.get(q(k))!="Arial" for k in ("ascii","hAnsi","eastAsia","cs")): raise SystemExit("Font check failed: "+a["label"])
        if rtl is None or rtl.get(q("val"))!="1": raise SystemExit("RTL check failed: "+a["label"])
        if lang is None or lang.get(q("val"))!="he-IL" or lang.get(q("bidi"))!="he-IL": raise SystemExit("Language check failed: "+a["label"])
    after=sha(ps2)
    if after!=before or len(ps2)!=expected: raise SystemExit("Body text or paragraph count changed")
    all_labels=[]
    label_pattern=re.compile(r"^\s*((?:E-\d{4}|S-\d{8}-\d{2,}))\s*\|")
    for comment_text in ctexts:
        match=label_pattern.match(comment_text)
        if match: all_labels.append(match.group(1))
    if len(all_labels)!=len(set(all_labels)): raise SystemExit("Duplicate comment labels in output")
report.parent.mkdir(parents=True,exist_ok=True)
lines=[f"# הערות Word — סשן {sid}","",f"- קובץ בסיס: {base.name}",f"- קובץ פלט חדש: {out.name}",f"- הערות לפני התוספת: {len(old)}",f"- הערות שנוספו: {len(added)}",f"- הערות לאחר התוספת: {len(cs)}",f"- מספר פסקאות: {len(ps2)}",f"- גיבוב טקסט גוף כתב־היד לפני ואחרי: {after} (זהה)","- גוף כתב־היד לא נערך; נוספו עוגני הערות Word בלבד.","- אומתו שלושת סוגי העוגנים, גופן Arial, RTL ושפת he-IL.","- מספור ההערות ההיסטורי E-0106–E-0767 לא נוצל.","","| מזהה | פסקה במסמך | חומרה | קטגוריה |","|---|---:|---|---|"]
for a in added: lines.append(f"| {a['label']} | {a['p']} | {a['severity']} | {a['category']} |")
report.write_text("\n".join(lines)+"\n",encoding="utf-8")
log=ROOT/"project-management/יומן-דחיפות-והמשכיות.md"
if log.exists():
    s=log.read_text(encoding="utf-8"); marker=f"## הערות Word — סשן {sid}"
    if marker not in s:
        s+="\n\n"+marker+"\n"+f"- קובץ בסיס: {base.name}\n- קובץ Word חדש: {out.name}\n"
        s+=f"- נוספו {len(added)} הערות Word: "+", ".join(f"{a['label']} (פסקה {a['p']})" for a in added)+".\n"
        s+=f"- מספר הערות כולל: {len(cs)}; מספר פסקאות: {len(ps2)}; גיבוב טקסט גוף ללא שינוי: {after}\n"
        s+="- לא נערך גוף כתב־היד ולא שונו הערות קודמות; נוספו עוגני הערות בלבד.\n"
        s+=f"- דוח הסשן: {report.name}\n- נקודת המשך: {next_section}; יש לבסס את הסשן הבא על {out.name}.\n"
        log.write_text(s,encoding="utf-8")
print(f"Created {out}: added {len(added)} session comments; total comments={len(cs)}; paragraphs={len(ps2)}; body SHA={after}")
