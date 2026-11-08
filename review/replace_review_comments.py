#!/usr/bin/env python3
"""Replace only explicitly identified, erroneous editorial comment text in a new DOCX copy."""
from __future__ import annotations
import hashlib, json, os, zipfile
from datetime import datetime, timezone
from pathlib import Path
from lxml import etree

W="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
X="http://www.w3.org/XML/1998/namespace"
NS={"w":W}
def q(name): return f"{{{W}}}{name}"
def ptext(p):
    return "".join(n.text or "" for n in p.iter() if n.tag in (q("t"),q("delText")))
def body_hash(paragraphs):
    return hashlib.sha256(json.dumps([ptext(p) for p in paragraphs],ensure_ascii=False).encode("utf-8")).hexdigest()
def comment_text(c):
    return "".join(c.xpath(".//w:t/text()",namespaces=NS))

def main():
    cfg_path=Path(os.environ.get("REVIEW_CORRECTION_CONFIG","review_batches/correct-source-annotations.json"))
    cfg=json.loads(cfg_path.read_text(encoding="utf-8"))
    base=Path(cfg["base_docx"]); out=Path(cfg["output_docx"]); report=Path(cfg["report_path"])
    if not base.is_file(): raise SystemExit(f"Missing base DOCX: {base}")
    if out.exists(): raise SystemExit(f"Refusing to overwrite output DOCX: {out}")
    expected_n=int(cfg["expected_paragraph_count"])
    expected_c=int(cfg["expected_comment_count"])
    expected_body=cfg["expected_body_sha256"]
    corrections=cfg["corrections"]
    if len({x["comment_id"] for x in corrections})!=len(corrections):
        raise SystemExit("Duplicate comment IDs in correction batch")

    with zipfile.ZipFile(base,"r") as zin:
        bad=zin.testzip()
        if bad: raise SystemExit(f"Corrupt source DOCX member: {bad}")
        names=set(zin.namelist())
        if not {"word/document.xml","word/comments.xml"}.issubset(names):
            raise SystemExit("Missing document.xml or comments.xml")
        document_bytes=zin.read("word/document.xml")
        doc=etree.fromstring(document_bytes)
        body=doc.find(".//"+q("body"))
        paragraphs=body.findall(q("p")) if body is not None else []
        if len(paragraphs)!=expected_n:
            raise SystemExit(f"Paragraph count mismatch: {len(paragraphs)} != {expected_n}")
        source_body_hash=body_hash(paragraphs)
        if source_body_hash!=expected_body:
            raise SystemExit(f"Manuscript body SHA mismatch: {source_body_hash}")
        comments_bytes=zin.read("word/comments.xml")
        comments_root=etree.fromstring(comments_bytes)
        all_comments=comments_root.findall(".//"+q("comment"))
        if len(all_comments)!=expected_c:
            raise SystemExit(f"Word comment count mismatch: {len(all_comments)} != {expected_c}")

        before={c.get(q("id")):comment_text(c) for c in all_comments}
        metadata_before={
            c.get(q("id")):(c.get(q("author")),c.get(q("initials")),c.get(q("date")))
            for c in all_comments
        }
        changed=[]
        for item in corrections:
            cid=str(item["comment_id"])
            matches=[c for c in all_comments if c.get(q("id"))==cid]
            if len(matches)!=1: raise SystemExit(f"Comment ID {cid} missing or duplicated")
            c=matches[0]
            old=comment_text(c)
            if not old.startswith(item["label"]+" |"):
                raise SystemExit(f"Comment ID {cid} label mismatch")
            old_sha=hashlib.sha256(old.encode("utf-8")).hexdigest()
            if old_sha!=item["expected_old_text_sha256"]:
                raise SystemExit(f"Comment ID {cid} old text changed since review")
            nodes=c.xpath(".//w:t",namespaces=NS)
            if len(nodes)!=1:
                raise SystemExit(f"Comment {cid} has {len(nodes)} text nodes; refusing nontrivial reconstruction")
            nodes[0].text=item["new_text"]
            nodes[0].set(f"{{{X}}}space","preserve")
            changed.append({
                "comment_id":cid,"label":item["label"],"old_sha256":old_sha,
                "old_text":old,"new_text":item["new_text"],
                "author":c.get(q("author")),"date":c.get(q("date"))
            })

        out_comments=etree.tostring(comments_root,xml_declaration=True,encoding="UTF-8",standalone=True)
        with zipfile.ZipFile(out,"w",compression=zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                if info.filename=="word/comments.xml": zout.writestr(info,out_comments)
                else: zout.writestr(info,zin.read(info.filename))

    with zipfile.ZipFile(out,"r") as zout:
        bad=zout.testzip()
        if bad: raise SystemExit(f"Corrupt output DOCX member: {bad}")
        if zout.read("word/document.xml")!=document_bytes:
            raise SystemExit("document.xml changed; comment-only correction violated")
        doc2=etree.fromstring(zout.read("word/document.xml"))
        body2=doc2.find(".//"+q("body"))
        paras2=body2.findall(q("p")) if body2 is not None else []
        if len(paras2)!=expected_n or body_hash(paras2)!=source_body_hash:
            raise SystemExit("Manuscript body or paragraph count changed")
        croot2=etree.fromstring(zout.read("word/comments.xml"))
        comments2=croot2.findall(".//"+q("comment"))
        if len(comments2)!=expected_c: raise SystemExit("Comment count changed")
        after={c.get(q("id")):comment_text(c) for c in comments2}
        if set(after)!=set(before): raise SystemExit("Comment ID set changed")
        for cid,old in before.items():
            correction=next((x for x in corrections if str(x["comment_id"])==cid),None)
            expected=correction["new_text"] if correction else old
            if after[cid]!=expected: raise SystemExit(f"Unexpected comment text change at ID {cid}")
        for c in comments2:
            cid=c.get(q("id"))
            if metadata_before[cid]!=(c.get(q("author")),c.get(q("initials")),c.get(q("date"))):
                raise SystemExit(f"Comment metadata changed at ID {cid}")
        for item in corrections:
            cid=str(item["comment_id"])
            c=next(c for c in comments2 if c.get(q("id"))==cid)
            runs=c.xpath(".//w:r",namespaces=NS)
            if len(runs)!=1: raise SystemExit(f"Unexpected run structure in comment {cid}")
            rpr=runs[0].find(q("rPr"))
            fonts=rpr.find(q("rFonts")) if rpr is not None else None
            rtl=rpr.find(q("rtl")) if rpr is not None else None
            lang=rpr.find(q("lang")) if rpr is not None else None
            if fonts is None or any(fonts.get(q(k))!="Arial" for k in ("ascii","hAnsi","eastAsia","cs")):
                raise SystemExit(f"Arial check failed for comment {cid}")
            if rtl is None or rtl.get(q("val")) not in (None,"1"): raise SystemExit(f"RTL check failed for comment {cid}")
            if lang is None or lang.get(q("val"))!="he-IL" or lang.get(q("bidi"))!="he-IL":
                raise SystemExit(f"Hebrew language check failed for comment {cid}")

    lines=[
        "# תיקון נקודתי של הערות מקור ב־Word","",
        f"- מסמך בסיס: {base.name}",f"- מסמך חדש: {out.name}",
        f"- מספר פסקאות: {expected_n}",f"- מספר הערות Word: {expected_c}",
        f"- SHA-256 של גוף כתב־היד: {source_body_hash}",
        "- לא נגעו ב־word/document.xml; טקסט כתב־היד וכל העוגנים נשארו ללא שינוי.",
        "- נשמרו מזהי ההערות, המחברים, התאריכים והעוגנים; שונו רק הטקסטים של ההערות המפורטות להלן.",
        "- כל שאר הערות Word נשמרו בדיוק כפי שהיו.",
        "",
        "| מזהה | Word ID | סוג התיקון |",
        "|---|---:|---|"
    ]
    for x in changed: lines.append(f"| {x['label']} | {x['comment_id']} | תיקון מראה מקום/קריאת מקור |")
    for x in changed:
        lines += ["",f"## {x['label']} (Word ID {x['comment_id']})","",
                  "### נוסח קודם", "", x["old_text"], "",
                  "### נוסח מתוקן", "", x["new_text"], ""]
    lines += [
        "## אימות","",
        f"- מספר פסקאות: {expected_n}; מספר הערות: {expected_c}.",
        f"- SHA-256 גוף כתב־היד לפני ואחרי: {source_body_hash}.",
        "- document.xml זהה byte-for-byte לקובץ הבסיס.",
        "- כל מזהי ההערות והמטא־נתונים נשמרו.",
        "- כל ההערות שלא צוינו ברשימת התיקונים זהות בדיוק.",
        "- אומתו Arial, RTL ושפת he-IL לכל הערה שתוקנה.",
        "- קובץ הבסיס לא שונה; נוצר קובץ DOCX חדש."
    ]
    report.parent.mkdir(parents=True,exist_ok=True)
    report.write_text("\n".join(lines)+"\n",encoding="utf-8")

    log_path=Path(cfg.get("worklog_path","יומן-דחיפות-והמשכיות.md"))
    marker=f"## תיקון הערות מקור — {cfg['session_id']}"
    if log_path.exists():
        log=log_path.read_text(encoding="utf-8")
        if marker not in log:
            log+="\n\n"+marker+"\n"
            log+=f"- מסמך בסיס: {base.name}\n- מסמך Word חדש: {out.name}\n"
            log+="- תוקנו טקסטי ההערות E-0937 ו־E-0945 בלבד; מזהי Word והעוגנים נשמרו.\n"
            log+=f"- מספר הערות נשאר {expected_c}; מספר פסקאות {expected_n}; גוף כתב־היד ללא שינוי ({source_body_hash}).\n"
            log+=f"- דוח: {report.as_posix()}.\n"
            log+=f"- נקודת המשך: {cfg['next_section']}; בסיס ההמשך הוא {out.name}.\n"
            log_path.write_text(log,encoding="utf-8")
    print(f"Created {out}; corrected {len(changed)} comment texts; comments={expected_c}; paragraphs={expected_n}; body SHA unchanged.")
if __name__=="__main__": main()
