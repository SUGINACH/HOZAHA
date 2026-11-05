#!/usr/bin/env python3
"""Read-only structural audit of manuscript/review DOCX files in HOZAHA."""
from __future__ import annotations
import hashlib, json, os, re, zipfile
from collections import Counter, defaultdict
from pathlib import Path
from xml.etree import ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
def q(local): return f"{{{W}}}{local}"
FILES = [
"כתב_יד_עם_הערות_עריכה.docx",
"כתב_יד_עם_הערות_עריכה_מתוקן_גופן.docx",
"כתב_יד_עם_הערות_עריכה_E0001-E1002_מתוקן_גופן.docx",
"כתב_יד_עם_הערות_עריכה_E0001-E1002_עם_השלמת_פער_ראשונה_E0022-E0054_מתוקן_גופן.docx",
"כתב_יד_עם_הערות_עריכה_E0001-E1002_עם_השלמת_פער_E0022-E0080.docx",
"המשך_ביקורת_הערות_E0768-E1002_מתוקן_גופן.docx",
"המשך_ביקורת_הערות_E0768-E0882_מתוקן_גופן.docx",
"המשך_ביקורת_הערות_E0768-E0841_מתוקן_גופן.docx",
]
def txt(el):
    return "" if el is None else "".join((x.text or "") for x in el.iter() if x.tag in (q("t"), q("delText")))
def tag_counts(root,names):
    return {n:len(root.findall(".//"+q(n))) for n in names}

def inspect(path):
    d={"name":path.name,"exists":path.exists()}
    if not path.exists(): return d
    d["bytes"]=path.stat().st_size
    try:
        with zipfile.ZipFile(path) as z:
            bad=z.testzip()
            names=set(z.namelist())
            d["zip_test"]="OK" if bad is None else "BAD: "+str(bad)
            d["comment_related_parts"]=sorted(n for n in names if "comment" in n.lower() or "people" in n.lower())
            if "word/document.xml" not in names:
                d["error"]="word/document.xml missing"; return d
            doc=ET.fromstring(z.read("word/document.xml"))
            body=doc.find(".//"+q("body"))
            paras=body.findall(q("p")) if body is not None else []
            ptexts=[txt(p) for p in paras]
            d["paragraph_texts"]=ptexts
            d["paragraph_count"]=len(paras)
            d["body_text_sha256"]=hashlib.sha256(json.dumps(ptexts,ensure_ascii=False).encode()).hexdigest()
            rev=tag_counts(doc,["ins","del","moveFrom","moveTo","moveFromRangeStart","moveFromRangeEnd","moveToRangeStart","moveToRangeEnd","pPrChange","rPrChange","tblPrChange","trPrChange","tcPrChange"])
            d["tracked_changes_document_xml"]={k:v for k,v in rev.items() if v}
            anchors=defaultdict(lambda:{"start":0,"end":0,"ref":0,"paragraphs":[]})
            for i,p in enumerate(paras,1):
                for e in p.iter():
                    if e.tag in (q("commentRangeStart"),q("commentRangeEnd"),q("commentReference")):
                        cid=e.get(q("id"),"?")
                        kind={q("commentRangeStart"):"start",q("commentRangeEnd"):"end",q("commentReference"):"ref"}[e.tag]
                        anchors[cid][kind]+=1
                        if i not in anchors[cid]["paragraphs"]: anchors[cid]["paragraphs"].append(i)
            comments=[]
            comment_xml_elements=[]
            if "word/comments.xml" in names:
                root=ET.fromstring(z.read("word/comments.xml"))
                for c in root.findall(".//"+q("comment")):
                    cid=c.get(q("id"),"?"); p=c.find(".//"+q("p"))
                    comments.append({"id":cid,"author":c.get(q("author"),""),"initials":c.get(q("initials"),""),"date":c.get(q("date"),""),"text":txt(c),"paraId":p.get(f"{{{W14}}}paraId") if p is not None else None,"anchor":anchors.get(cid,{"start":0,"end":0,"ref":0,"paragraphs":[]})})
                    comment_xml_elements.append(c)
            d["comment_count"]=len(comments)
            d["comment_authors"]=dict(Counter(c["author"] or "(author missing)" for c in comments))
            nums=[int(c["id"]) for c in comments if c["id"].isdigit()]
            d["comment_id_minmax"]=[min(nums),max(nums)] if nums else []
            labels=[int(x) for c in comments for x in re.findall(r"E-(\d{4})",c["text"])]
            d["E_label_count"]=len(labels)
            d["E_label_minmax"]=[min(labels),max(labels)] if labels else []
            d["E_label_duplicates"]=sorted(n for n,v in Counter(labels).items() if v>1)
            d["comment_anchor_counts"]=tag_counts(doc,["commentRangeStart","commentRangeEnd","commentReference"])
            d["commented_paragraphs"]=sorted(i for i,p in enumerate(paras,1) if any(e.tag in (q("commentRangeStart"),q("commentRangeEnd"),q("commentReference")) for e in p.iter()))
            d["comments_with_bad_anchors"]=[{"id":c["id"],"author":c["author"],"labels":re.findall(r"E-\d{4}",c["text"]),**c["anchor"]} for c in comments if any(c["anchor"][k]!=1 for k in ("start","end","ref"))]
            d["all_comments"]=comments
            if path.name=="כתב_יד_עם_הערות_עריכה.docx":
                d["legacy_comment_records"]=comments
            threaded=[]
            if "word/commentsExtended.xml" in names:
                er=ET.fromstring(z.read("word/commentsExtended.xml"))
                for e in er.findall(".//"+f"{{{W15}}}commentEx"):
                    threaded.append({"paraId":e.get(f"{{{W15}}}paraId"),"parentParaId":e.get(f"{{{W15}}}paraIdParent"),"done":e.get(f"{{{W15}}}done")})
            pmap={c["paraId"]:c for c in comments if c["paraId"]}
            d["thread_metadata_records"]=len(threaded)
            d["thread_replies"]=[{"reply_id":pmap.get(e["paraId"],{}).get("id"),"reply_author":pmap.get(e["paraId"],{}).get("author"),"parent_id":pmap.get(e["parentParaId"],{}).get("id"),"parent_author":pmap.get(e["parentParaId"],{}).get("author"),"done":e["done"]} for e in threaded if e["parentParaId"]]
            fonts=Counter(); rtls=Counter(); langs=Counter()
            for c in comment_xml_elements:
                for r in c.iter(q("r")):
                    rp=r.find(q("rPr"))
                    if rp is None: fonts["(no rPr)"]+=1; continue
                    rf=rp.find(q("rFonts"))
                    fonts["/".join(rf.get(q(k),"-") for k in ("ascii","hAnsi","eastAsia","cs")) if rf is not None else "(no rFonts)"]+=1
                    rtl=rp.find(q("rtl")); rtls[rtl.get(q("val"),"1") if rtl is not None else "(missing)"]+=1
                    lang=rp.find(q("lang")); langs[(lang.get(q("val"),"-")+"/"+lang.get(q("bidi"),"-")) if lang is not None else "(missing)"]+=1
            d["comment_run_fonts"]=dict(fonts); d["comment_run_rtl"]=dict(rtls); d["comment_run_languages"]=dict(langs)
            if "word/styles.xml" in names:
                styles=ET.fromstring(z.read("word/styles.xml"))
                st=next((s for s in styles.findall(".//"+q("style")) if s.get(q("styleId"))=="CommentText"),None)
                if st is not None:
                    pp=st.find(q("pPr")); rp=st.find(q("rPr")); rf=rp.find(q("rFonts")) if rp is not None else None
                    d["CommentText_style"]={"bidi":pp.find(q("bidi")).get(q("val"),"1") if pp is not None and pp.find(q("bidi")) is not None else None,"jc":pp.find(q("jc")).get(q("val")) if pp is not None and pp.find(q("jc")) is not None else None,"fonts":{k:rf.get(q(k)) for k in ("ascii","hAnsi","eastAsia","cs")} if rf is not None else None}
            other={}
            for n in names:
                if n.endswith(".xml") and n!="word/document.xml":
                    try:
                        x=ET.fromstring(z.read(n)); c=tag_counts(x,["ins","del","moveFrom","moveTo","pPrChange","rPrChange"])
                        if any(c.values()): other[n]={k:v for k,v in c.items() if v}
                    except Exception: pass
            d["tracked_changes_other_parts"]=other
    except Exception as e: d["error"]=f"{type(e).__name__}: {e}"
    return d

def main():
    ds=[inspect(Path(x)) for x in FILES]; by={d["name"]:d for d in ds}
    lines=["# דוח ביקורת מבנית של מסמכי DOCX — HOZAHA","",
    f"- ריצה: {os.environ.get('GITHUB_RUN_ID','local')}",
    f"- commit שנבדק: {os.environ.get('GITHUB_SHA','unknown')}",
    "- הבדיקה קריאתית בלבד: לא נערך אף קובץ DOCX.",
    "- נבדקו חלקי OOXML: הערות, שרשורים, עוגנים, שינויים מסומנים, סגנונות וגופנים.","",
    "## תמונת מצב לפי קובץ","","| קובץ | גודל | פסקאות | הערות Word | מחברים | תוויות E | ZIP |","|---|---:|---:|---:|---|---:|---|"]
    for d in ds:
        if not d["exists"]: lines.append(f"| {d['name']} | — | — | — | — | — | חסר |"); continue
        authors="; ".join(f"{a}: {n}" for a,n in d.get("comment_authors",{}).items()) or "—"
        er=d.get("E_label_minmax",[])
        labels=f"{er[0]}–{er[1]} ({d.get('E_label_count',0)})" if er else "—"
        lines.append(f"| {d['name']} | {d.get('bytes',0):,} | {d.get('paragraph_count','?')} | {d.get('comment_count','?')} | {authors} | {labels} | {d.get('zip_test','?')} |")
    lines += ["","## שגיאות סריקה",""]
    errors=[f"{d['name']}: {d['error']}" for d in ds if d.get("error")]
    lines.extend(["- "+e for e in errors] if errors else ["לא נרשמו שגיאות סריקה."])
    base=by.get("כתב_יד_עם_הערות_עריכה.docx",{})
    lines += ["","## השוואת גוף כתב־היד",""]
    for n in FILES[1:5]:
        d=by.get(n,{})
        if base.get("body_text_sha256") and d.get("body_text_sha256"):
            same=base["body_text_sha256"]==d["body_text_sha256"] and base.get("paragraph_count")==d.get("paragraph_count")
            lines.append(f"- מול {n}: **{'זהה' if same else 'שונה'}**; פסקאות {base.get('paragraph_count')} / {d.get('paragraph_count')}.")
        elif not d.get("exists"): lines.append(f"- {n}: קובץ חסר ב־HEAD.")
        else: lines.append(f"- {n}: השוואה לא זמינה ({d.get('error','נתונים חסרים')}).")
    lines += ["","## הערות בקובץ כתב־היד הוותיק","","| מזהה Word | מחבר | ראשי תיבות | פסקה | תוכן |","|---:|---|---|---:|---|"]
    legacy=base.get("legacy_comment_records",[])
    if not legacy: lines.append("לא נמצאו הערות או שלא ניתן לקרוא אותן.")
    for c in legacy:
        content=c["text"].replace("|","\\|").replace("\n"," ")
        p=", ".join(map(str,c["anchor"]["paragraphs"])) or "עוגן חסר"
        lines.append(f"| {c['id']} | {c['author']} | {c['initials']} | {p} | {content} |")
    lines += ["","## בדיקת התנגשויות באצוות E-0022–E-0105",""]
    batches=[]
    all_targets=defaultdict(list)
    for bp in sorted(Path("scripts").glob("gap_batch_E*.json")):
        try:
            items=json.loads(bp.read_text(encoding="utf-8"))
            batches.append((bp.name,items))
            for item in items:
                all_targets[int(item["p"])].append(int(item["seq"]))
        except Exception as e:
            lines.append(f"- שגיאה בקריאת {bp.name}: {type(e).__name__}: {e}")
    base_doc=by.get("כתב_יד_עם_הערות_עריכה_E0001-E1002_מתוקן_גופן.docx",{})
    base_ps=set(base_doc.get("commented_paragraphs",[]))
    for batch_name,items in batches:
        local=defaultdict(list)
        for item in items:
            local[int(item["p"])].append(int(item["seq"]))
        repeated={p:seqs for p,seqs in local.items() if len(seqs)>1}
        preexisting={p:seqs for p,seqs in local.items() if p in base_ps}
        lines.append(f"### {batch_name}")
        lines.append(f"- מספר הערות: {len(items)}; פסקאות יעד ייחודיות: {len(local)}.")
        lines.append(f"- פסקאות יעד שכבר מכילות עוגן בקובץ הבסיס: {json.dumps(preexisting,ensure_ascii=False)}.")
        lines.append(f"- פסקאות יעד החוזרות בתוך האצווה: {json.dumps(repeated,ensure_ascii=False)}.")
    cross={p:seqs for p,seqs in all_targets.items() if len(seqs)>1}
    lines.append(f"- חזרות בין כל האצוות: {json.dumps(cross,ensure_ascii=False)}.")
    for pnum,seqs in cross.items():
        text_p=(base_doc.get("paragraph_texts") or [])
        para_text=text_p[pnum-1] if 0 < pnum <= len(text_p) else "(הפסקה לא נמצאה)"
        lines.append(f"- נוסח פסקה {pnum} (יעד E-{seqs[0]:04d}, E-{seqs[1]:04d}):")
        lines.append("  " + para_text.replace("\\n"," "))
    lines.append("- מגבלת הבונה: כל הערה חדשה מעוגנת לכל הפסקה, והסקריפט מסרב להוסיף הערה אם הפסקה כבר מכילה סימון הערה כלשהו. לכן חזרה על אותה פסקה או עוגן קודם בפסקת היעד הם חסם בנייה, גם אם תוכן שתי ההערות נוגע למשפטים שונים.")

    lines += ["","## שרשורי תגובות",""]
    replies=False
    for d in ds:
        if d.get("thread_replies"):
            replies=True; lines.append(f"### {d['name']}")
            for r in d["thread_replies"]: lines.append(f"- תגובה {r['reply_id']} ({r['reply_author']}) להערה {r['parent_id']} ({r['parent_author']}); done={r['done']}.")
    if not replies: lines.append("לא זוהו תגובות משורשרות באמצעות commentsExtended.xml. אין בכך לשלול הערות מחבר שהוזנו כהערות Word רגילות.")
    lines += ["","## עוגנים, שינויים מסומנים וגופנים",""]
    for d in ds:
        if not d["exists"]: continue
        a=d.get("comment_anchor_counts",{})
        lines += [f"### {d['name']}",
        f"- עוגנים: התחלה {a.get('commentRangeStart',0)}, סוף {a.get('commentRangeEnd',0)}, הפניה {a.get('commentReference',0)}.",
        f"- הערות עם עוגנים חסרים/כפולים: {len(d.get('comments_with_bad_anchors',[]))}.",
        f"- פירוט חריגות: {json.dumps(d.get('comments_with_bad_anchors',[]),ensure_ascii=False)}",
        f"- תוויות E כפולות: {len(d.get('E_label_duplicates',[]))}.",
        f"- שינויים מסומנים במסמך הראשי: {json.dumps(d.get('tracked_changes_document_xml',{}),ensure_ascii=False)}",
        f"- מטא־נתוני שרשורים: {d.get('thread_metadata_records',0)}; תגובות מזוהות: {len(d.get('thread_replies',[]))}.",
        f"- גופני ריצות הערות: {json.dumps(d.get('comment_run_fonts',{}),ensure_ascii=False)}.",
        f"- RTL: {json.dumps(d.get('comment_run_rtl',{}),ensure_ascii=False)}; שפות: {json.dumps(d.get('comment_run_languages',{}),ensure_ascii=False)}."]
        if d.get("CommentText_style") is not None: lines.append(f"- CommentText style: {json.dumps(d['CommentText_style'],ensure_ascii=False)}.")
        if d.get("tracked_changes_other_parts"): lines.append(f"- שינויים מסומנים בחלקים אחרים: {json.dumps(d['tracked_changes_other_parts'],ensure_ascii=False)}.")
    lines += ["","## מגבלות","",
    "- הערת Word רגילה אינה בהכרח תגובה של המחבר; יש לפרש מחבר, תאריך ותוכן לפי ההקשר.",
    "- תגובות משורשרות זוהו לפי paraIdParent בחלקי Word המתאימים.",
    "- hash גוף הטקסט משווה טקסט בפסקאות הראשיות; אינו תחליף לבדיקת עימוד חזותית או לבדיקה הלכתית.",
    "- הדוח בודק שלמות מבנית, לא את נכונות התוכן של ההערות.",""]
    Path("דוח_ביקורת_מבנית_DOCX.md").write_text("\n".join(lines),encoding="utf-8")
    log=Path("יומן-דחיפות-והמשכיות.md")
    if log.exists():
        old=log.read_text(encoding="utf-8")
        entry=["","## ביקורת מבנית אוטומטית של קובצי DOCX","",
        f"- ריצה: {os.environ.get('GITHUB_RUN_ID','local')}; commit בסיס: {os.environ.get('GITHUB_SHA','unknown')}.",
        "- נוספו audit/inspect_docx.py ו־workflow ייעודי. שניהם קוראים את קובצי ה־DOCX ואינם משנים אותם.",
        "- הופק דוח_ביקורת_מבנית_DOCX.md: ספירת הערות, מחברים, שרשורים, עוגנים, שינויים מסומנים, גופנים והשוואת גוף כתב־היד.",
        "- נקודת המשך: לבחון את הערות המחבר והתגובות לפי הדוח; רק לאחר מכן להחליט איזו גרסה תהיה בסיס העבודה. אין לדרוס DOCX בשלב זה.",""]
        log.write_text(old.rstrip()+"\n"+"\n".join(entry),encoding="utf-8")
    print(f"Audited {sum(1 for d in ds if d.get('exists'))}/{len(FILES)} DOCX files.")

if __name__=="__main__": main()
