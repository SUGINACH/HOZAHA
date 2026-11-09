#!/usr/bin/env python3
"""Append narrowly anchored editorial comments to an existing cumulative DOCX."""
from __future__ import annotations
import hashlib, json, re, zipfile, os
from datetime import datetime, timezone
from pathlib import Path
from lxml import etree

ROOT=Path(".")
CONFIG=Path(os.environ.get("REVIEW_BATCH_CONFIG", "review_batches/P02555-P02571.json"))
W="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
X="http://www.w3.org/XML/1998/namespace"
NS={"w":W}
def q(s): return f"{{{W}}}{s}"
def ptxt(p):
    return "".join(n.text or "" for n in p.iter() if n.tag in (q("t"),q("delText")))
def digest(ps):
    return hashlib.sha256(json.dumps([ptxt(p) for p in ps],ensure_ascii=False).encode("utf-8")).hexdigest()
def ctxt(c): return "".join(c.xpath(".//w:t/text()",namespaces=NS))
def labels(cs):
    pat=re.compile(r"^\s*((?:E-\d{4}|S-\d{8}-\d{2,}))\s*\|")
    out=[]
    for c in cs:
        m=pat.match(ctxt(c))
        if m: out.append(m.group(1))
    return out

def main():
    cfg=json.loads(CONFIG.read_text(encoding="utf-8"))
    base=ROOT/cfg["base_docx"]; out=ROOT/cfg["output_docx"]; report=ROOT/cfg["report_path"]
    expected_n=int(cfg["expected_paragraph_count"])
    expected_c=int(cfg["expected_existing_comment_count"])
    items=cfg["comments"]
    if not base.is_file(): raise SystemExit(f"Missing base DOCX: {base}")
    if out.exists(): raise SystemExit(f"Refusing to overwrite output DOCX: {out}")
    if not items: raise SystemExit("Empty review batch")
    ids=[x["label"] for x in items]
    if len(ids)!=len(set(ids)) or any(not re.fullmatch(r"S-\d{8}-\d{3,}",x) for x in ids):
        raise SystemExit("Invalid or duplicate review labels")

    with zipfile.ZipFile(base,"r") as zin:
        bad=zin.testzip()
        if bad: raise SystemExit(f"Corrupt base DOCX member: {bad}")
        names=set(zin.namelist())
        required={"word/document.xml","word/comments.xml","word/_rels/document.xml.rels","[Content_Types].xml"}
        if not required.issubset(names): raise SystemExit("Base DOCX lacks required Word comments parts")
        doc=etree.fromstring(zin.read("word/document.xml"))
        cr=etree.fromstring(zin.read("word/comments.xml"))
        body=doc.find(".//"+q("body"))
        ps=body.findall(q("p")) if body is not None else []
        old=cr.findall(q("comment"))
        if len(ps)!=expected_n: raise SystemExit(f"Paragraph count mismatch: {len(ps)} != {expected_n}")
        if len(old)!=expected_c: raise SystemExit(f"Existing comment count mismatch: {len(old)} != {expected_c}")
        before=digest(ps)
        if before!=cfg["expected_body_sha256"]: raise SystemExit(f"Unexpected manuscript hash: {before}")
        oldlabels=labels(old)
        if len(oldlabels)!=len(set(oldlabels)): raise SystemExit("Duplicate labels already present in base")
        if set(ids).intersection(oldlabels): raise SystemExit("New label already exists in base")
        oldtexts=[ctxt(c) for c in old]
        word_ids=[int(c.get(q("id"))) for c in old if (c.get(q("id")) or "").isdigit()]
        next_id=max(word_ids,default=-1)+1

        planned=[]; used=set()
        for item in items:
            pnum=int(item["p"])
            if not 1<=pnum<=len(ps): raise SystemExit(f"Target paragraph out of range: {pnum}")
            if pnum in used: raise SystemExit(f"Duplicate paragraph target: {pnum}")
            used.add(pnum); p=ps[pnum-1]
            if not ptxt(p).startswith(item["expected_prefix"]):
                raise SystemExit(f"Paragraph prefix mismatch at Word paragraph {pnum}")
            if p.xpath(".//w:commentRangeStart | .//w:commentRangeEnd | .//w:commentReference",namespaces=NS):
                raise SystemExit(f"Refusing to overlap an existing comment in paragraph {pnum}")
            runs=[r for r in p.findall(q("r")) if any(n.tag in (q("t"),q("delText")) and (n.text or "") for n in r.iter())]
            if not runs: raise SystemExit(f"No text runs at paragraph {pnum}")
            planned.append((item,pnum,p,runs[0],runs[-1],str(next_id)))
            next_id+=1

        stamp=datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00","Z")
        added=[]
        for item,pnum,p,first,last,wid in planned:
            start=etree.Element(q("commentRangeStart"),attrib={q("id"):wid})
            end=etree.Element(q("commentRangeEnd"),attrib={q("id"):wid})
            p.insert(list(p).index(first),start)
            p.insert(list(p).index(last)+1,end)
            ref=etree.Element(q("r")); rp=etree.SubElement(ref,q("rPr"))
            etree.SubElement(rp,q("rStyle")).set(q("val"),"CommentReference")
            etree.SubElement(rp,q("rtl")).set(q("val"),"1")
            etree.SubElement(rp,q("vertAlign")).set(q("val"),"superscript")
            etree.SubElement(ref,q("commentReference")).set(q("id"),wid)
            p.insert(list(p).index(end)+1,ref)

            c=etree.SubElement(cr,q("comment"))
            c.set(q("id"),wid); c.set(q("author"),"OpenAI – ביקורת עריכה")
            c.set(q("initials"),"AI"); c.set(q("date"),stamp)
            cp=etree.SubElement(c,q("p")); pp=etree.SubElement(cp,q("pPr"))
            etree.SubElement(pp,q("pStyle")).set(q("val"),"CommentText")
            etree.SubElement(pp,q("bidi")).set(q("val"),"1")
            etree.SubElement(pp,q("jc")).set(q("val"),"right")
            r=etree.SubElement(cp,q("r")); rpr=etree.SubElement(r,q("rPr"))
            fonts=etree.SubElement(rpr,q("rFonts"))
            for k in ("ascii","hAnsi","eastAsia","cs"): fonts.set(q(k),"Arial")
            etree.SubElement(rpr,q("rtl")).set(q("val"),"1")
            lang=etree.SubElement(rpr,q("lang")); lang.set(q("val"),"he-IL"); lang.set(q("bidi"),"he-IL")
            etree.SubElement(rpr,q("sz")).set(q("val"),"20")
            t=etree.SubElement(r,q("t")); t.set(f"{{{X}}}space","preserve")
            t.text=f"{item['label']} | חומרה: {item['severity']} | קטגוריה: {item['category']}. {item['body']}"
            added.append({"label":item["label"],"word_id":wid,"pnum":pnum,
                          "manuscript_p":item["manuscript_p"],"severity":item["severity"],
                          "category":item["category"],"body":item["body"]})

        rels=etree.fromstring(zin.read("word/_rels/document.xml.rels"))
        if not any(x.get("Type","").endswith("/comments") and x.get("Target")=="comments.xml" for x in rels):
            raise SystemExit("Comments relationship missing")
        types=etree.fromstring(zin.read("[Content_Types].xml"))
        if not any(x.get("PartName")=="/word/comments.xml" for x in types):
            raise SystemExit("Comments content-type override missing")
        dx=etree.tostring(doc,xml_declaration=True,encoding="UTF-8",standalone=True)
        cx=etree.tostring(cr,xml_declaration=True,encoding="UTF-8",standalone=True)
        with zipfile.ZipFile(out,"w",compression=zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                if info.filename=="word/document.xml": zout.writestr(info,dx)
                elif info.filename=="word/comments.xml": zout.writestr(info,cx)
                else: zout.writestr(info,zin.read(info.filename))

    with zipfile.ZipFile(out,"r") as z:
        bad=z.testzip()
        if bad: raise SystemExit(f"Corrupt output DOCX member: {bad}")
        d2=etree.fromstring(z.read("word/document.xml")); c2=etree.fromstring(z.read("word/comments.xml"))
        b2=d2.find(".//"+q("body")); ps2=b2.findall(q("p")) if b2 is not None else []
        cs2=c2.findall(q("comment"))
        if len(ps2)!=expected_n or digest(ps2)!=before: raise SystemExit("Manuscript text or paragraph count changed")
        if len(cs2)!=expected_c+len(added): raise SystemExit("Incorrect final Word comment count")
        labs=labels(cs2)
        if len(labs)!=len(set(labs)) or not set(oldlabels).issubset(labs) or not set(ids).issubset(labs):
            raise SystemExit("Final editorial label validation failed")
        texts2=[ctxt(c) for c in cs2]
        if any(texts2.count(t)!=1 for t in oldtexts): raise SystemExit("Existing comment text changed or duplicated")
        for a in added:
            p=ps2[a["pnum"]-1]; wid=a["word_id"]
            for tag in ("commentRangeStart","commentRangeEnd","commentReference"):
                matches=[x for x in p.iter() if x.tag==q(tag) and x.get(q("id"))==wid]
                if len(matches)!=1: raise SystemExit(f"Missing or duplicate Word anchor: {a['label']}")
            found=[c for c in cs2 if c.get(q("id"))==wid]
            if len(found)!=1: raise SystemExit(f"Missing/duplicate comment body: {a['label']}")
            runs=found[0].xpath(".//w:r",namespaces=NS)
            if len(runs)!=1: raise SystemExit(f"Unexpected comment run count: {a['label']}")
            rpr=runs[0].find(q("rPr"))
            f=rpr.find(q("rFonts")) if rpr is not None else None
            rtl=rpr.find(q("rtl")) if rpr is not None else None
            lang=rpr.find(q("lang")) if rpr is not None else None
            if f is None or any(f.get(q(k))!="Arial" for k in ("ascii","hAnsi","eastAsia","cs")):
                raise SystemExit(f"Arial check failed: {a['label']}")
            if rtl is None or rtl.get(q("val"))!="1": raise SystemExit(f"RTL check failed: {a['label']}")
            if lang is None or lang.get(q("val"))!="he-IL" or lang.get(q("bidi"))!="he-IL":
                raise SystemExit(f"Language check failed: {a['label']}")

    lines=[f"# הערות Word — סשן {cfg['session_id']}","",
           f"- קובץ בסיס: {base.name}",f"- קובץ פלט חדש: {out.name}",
           f"- הערות לפני התוספת: {expected_c}",f"- הערות שנוספו: {len(added)}",
           f"- הערות לאחר התוספת: {expected_c+len(added)}",
           f"- מספר פסקאות: {expected_n}",f"- גיבוב גוף כתב־היד: {before} (זהה לפני ואחרי)",
           "- גוף כתב־היד לא נערך; נוספו עוגני הערות Word בלבד.",
           "- כל הערה חדשה עוגנה לפסקה שלא הכילה עוגן קודם; לא נוצרה חפיפה.",
           "- אומתו start/end/reference, Arial, RTL ושפת he-IL; תוכן כל ההערות הקודמות נשמר.",
           "","| מזהה | פסקת כתב־יד | פסקה במסמך Word | חומרה | קטגוריה |","|---|---:|---:|---|---|"]
    for a in added:
        lines.append(f"| {a['label']} | {a['manuscript_p']} | {a['pnum']} | {a['severity']} | {a['category']} |")
    lines += ["","## תוכן ההערות",""]
    for a in added: lines += [f"### {a['label']} — {a['manuscript_p']}","",a["body"],""]
    lines += ["## אימות סופי","",f"- SHA-256 של גוף כתב־היד: {before}.",
              f"- מספר הערות Word: {expected_c+len(added)}.",
              f"- מספר פסקאות: {expected_n}.",
              "- לא שונו תוויות, טקסטים או עוגנים של הערות קודמות.",
              "- קובץ המקור לא שונה; קובץ הפלט נכתב בשם חדש."]
    report.parent.mkdir(parents=True,exist_ok=True)
    report.write_text("\n".join(lines)+"\n",encoding="utf-8")

    log=ROOT/cfg.get("worklog_path","project-management/יומן-דחיפות-והמשכיות.md")
    marker=f"## הערות Word — סשן {cfg['session_id']}"
    if log.exists():
        s=log.read_text(encoding="utf-8")
        if marker not in s:
            s+="\n\n"+marker+"\n"
            s+=f"- קובץ בסיס: {base.name}\n- קובץ Word חדש: {out.name}\n"
            s+=f"- נוספו {len(added)} הערות: "+", ".join(f"{a['label']} ({a['manuscript_p']})" for a in added)+".\n"
            s+=f"- מספר הערות כולל: {expected_c+len(added)}; מספר פסקאות: {expected_n}; גיבוב גוף ללא שינוי: {before}.\n"
            s+="- גוף כתב־היד והערות קודמות לא שונו.\n"
            s+=f"- דוח הסשן: {report.name}\n- נקודת המשך: {cfg['next_section']}; יש לבסס את הסשן הבא על {out.name}.\n"
            log.write_text(s,encoding="utf-8")
    print(f"Created {out}; added {len(added)} comments; total={expected_c+len(added)}; paragraphs={expected_n}; body SHA={before}")

if __name__=="__main__":
    main()
