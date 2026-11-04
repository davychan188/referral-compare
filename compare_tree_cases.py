#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Referral compare. Supports weekly KC/WTT sheets and Summary table format."""
import sys, re, ctypes
from pathlib import Path
from datetime import datetime, date, timedelta

try:
    from openpyxl import load_workbook, Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo
except ImportError:
    print("ERROR: openpyxl missing"); input("Press Enter..."); sys.exit(1)

def _win_dialog(title, save=False, default=""):
    if sys.platform != "win32":
        return input(title + " (full path): ").strip().strip('"')
    class OPENFILENAME(ctypes.Structure):
        _fields_ = [
            ("lStructSize", ctypes.c_uint), ("hwndOwner", ctypes.c_void_p),
            ("hInstance", ctypes.c_void_p), ("lpstrFilter", ctypes.c_wchar_p),
            ("lpstrCustomFilter", ctypes.c_wchar_p), ("nMaxCustFilter", ctypes.c_uint),
            ("nFilterIndex", ctypes.c_uint), ("lpstrFile", ctypes.c_wchar_p),
            ("nMaxFile", ctypes.c_uint), ("lpstrFileTitle", ctypes.c_wchar_p),
            ("nMaxFileTitle", ctypes.c_uint), ("lpstrInitialDir", ctypes.c_wchar_p),
            ("lpstrTitle", ctypes.c_wchar_p), ("Flags", ctypes.c_uint),
            ("nFileOffset", ctypes.c_ushort), ("nFileExtension", ctypes.c_ushort),
            ("lpstrDefExt", ctypes.c_wchar_p), ("lCustData", ctypes.c_void_p),
            ("lpfnHook", ctypes.c_void_p), ("lpTemplateName", ctypes.c_wchar_p),
            ("pvReserved", ctypes.c_void_p), ("dwReserved", ctypes.c_uint), ("FlagsEx", ctypes.c_uint),
        ]
    buf = ctypes.create_unicode_buffer(4096)
    if default: buf.value = default
    ofn = OPENFILENAME()
    ofn.lStructSize = ctypes.sizeof(ofn)
    ofn.lpstrFilter = "Excel\0*.xlsx;*.xls;*.xlsm\0All\0*.*\0"
    ofn.lpstrFile = ctypes.cast(buf, ctypes.c_wchar_p)
    ofn.nMaxFile = 4096
    ofn.lpstrTitle = title
    ofn.lpstrDefExt = "xlsx"
    ofn.Flags = 0x2 if save else 0x1000
    fn = ctypes.windll.comdlg32.GetSaveFileNameW if save else ctypes.windll.comdlg32.GetOpenFileNameW
    return buf.value if fn(ctypes.byref(ofn)) else ""

def win_open_file(t="Select Excel"): return _win_dialog(t)
def win_save_file(t="Save report", default="Diff.xlsx"): return _win_dialog(t, True, default)

HEADER_MAP = {
    "source of referral": "Source of Referral", "estate": "Estate", "estate (eng)": "Estate",
    "case no.": "Case No.", "case no": "Case No.", "case number": "Case No.",
    "date received": "Date Received", "date of referral request": "Date Received",
    "tree id": "Tree ID", "tree no.": "Tree ID", "tree no": "Tree ID",
    "observatrion": "Observation", "observation": "Observation", "case description": "Observation",
    "mitigation": "Mitigation", "recommendation": "Mitigation",
    "case deadline*\n(normal case)": "Case deadline", "case deadline* (normal case)": "Case deadline",
    "case deadline": "Case deadline", "anticipated completion date by ts": "Case deadline",
    "works completion date": "Works Completion Date", "actual completion date": "Works Completion Date",
    "remarks": "Remarks", "region": "Region",
}
WORD_NORMALIZE = [(r"\bbranches\b","branch"),(r"\btwigs\b","twig"),(r"\bleaves\b","leaf"),(r"\bfronds\b","frond"),
                  (r"\bwounds\b","wound"),(r"\bcavities\b","cavity"),(r"\broots\b","root"),(r"\btrees\b","tree"),
                  (r"\bstubs\b","stub"),(r"\bhangers\b","hanger")]
TEXT_COLS = {"Observation","Mitigation","Remarks","Source of Referral","Estate"}
COMPARE_COLS = ["Observation", "Mitigation", "Works Completion Date", "Date Received", "Case deadline"]
RISK_WORDS = ("dead tree", "wilting", "wilted", "severe sparse foliage", "root plate movement", "fungal fruiting")

def norm_header(h):
    if h is None: return None
    s = str(h).strip().lower().replace("\r\n","\n").replace("\r","\n")
    s2 = re.sub(r"\s+"," ", s.replace("\n"," ")).strip()
    if s2 in HEADER_MAP: return HEADER_MAP[s2]
    if "case no" in s2: return "Case No."
    if s2.startswith("tree id") or s2.startswith("tree no"): return "Tree ID"
    if "observ" in s2 or "description" in s2: return "Observation"
    if "mitigation" in s2 or s2=="recommendation": return "Mitigation"
    if "completion date" in s2 and "actual" in s2: return "Works Completion Date"
    if "anticipated completion" in s2 or "deadline" in s2: return "Case deadline"
    if "referral request" in s2 or "date received" in s2: return "Date Received"
    if s2=="estate": return "Estate"
    if "remarks" in s2: return "Remarks"
    if s2=="region": return "Region"
    return str(h).strip()

def norm_val(v):
    if v is None or v=="" or str(v).strip().upper() in ("N/A","NA","NONE","-"): return None
    if isinstance(v, datetime): return v.date()
    if isinstance(v, date): return v
    s=str(v).strip()
    for fmt in ("%Y-%m-%d","%d/%m/%Y","%Y/%m/%d","%d-%m-%Y"):
        try: return datetime.strptime(s, fmt).date()
        except ValueError: pass
    return s

def norm_text(v):
    if v is None: return None
    s=str(v).strip().lower().replace("&"," and ")
    s=re.sub(r"[^\w\s]"," ", s)
    s=re.sub(r"\s+"," ", s).strip()
    for pat,rep in WORD_NORMALIZE: s=re.sub(pat, rep, s)
    return s

def values_equal(col, va, vb):
    if va is None and vb is None: return True
    if va is None or vb is None: return False
    if isinstance(va, date) and isinstance(vb, date): return abs((va-vb).days)<=1
    if isinstance(va, date) or isinstance(vb, date): return va==vb
    if col in TEXT_COLS or isinstance(va,str) or isinstance(vb,str):
        return norm_text(va)==norm_text(vb)
    return va==vb

def deadline_is_one_day_from_received(rec):
    a,b=rec.get("Date Received"), rec.get("Case deadline")
    return isinstance(a,date) and isinstance(b,date) and abs((b-a).days)<=1

def is_incomplete(rec):
    return rec.get("Works Completion Date") is None

def is_overdue_incomplete(rec, as_of=None):
    if not is_incomplete(rec) or deadline_is_one_day_from_received(rec): return False
    d=rec.get("Case deadline")
    if not isinstance(d, date): return False
    return d <= (as_of or date.today())

def risk_hits(rec):
    obs = norm_text(rec.get("Observation")) or ""
    return [w for w in RISK_WORDS if w in obs]

def follow_up_status(rec):
    done = rec.get("Works Completion Date")
    mit = norm_text(rec.get("Mitigation")) or ""
    acted = any(w in mit for w in ("removal","remove","fell","felling","tpc","pruning","prune","done"))
    return "已跟進" if done or acted else "未跟進"

def completion_status(rec):
    return "已完成" if rec.get("Works Completion Date") is not None else "未完成"

def obs_n(rec):
    return norm_text(rec.get("Observation")) or ""

def mit_n(rec):
    return norm_text(rec.get("Mitigation")) or ""

def is_dead_branch_obs(rec):
    """dead branch / broken branch / hanger 等（norm 後 branches→branch）。"""
    obs = obs_n(rec)
    if "dead branch" in obs or "broken branch" in obs:
        return True
    if "dead and broken" in obs and "branch" in obs:
        return True
    if "hanger" in obs and ("dead" in obs or "broken" in obs or "branch" in obs):
        return True
    return False

def is_pruning_mit(rec):
    mit = mit_n(rec)
    return "pruning" in mit or "prune" in mit

def is_dead_tree_like_obs(rec):
    """dead tree / wilted / wilting / dieback / die back / abnormal defoliation。"""
    obs = obs_n(rec)
    keys = (
        "dead tree", "wilted", "wilting", "dieback", "die back",
        "abnormal defoliation", "wilted crown", "wilting crown",
    )
    return any(k in obs for k in keys)

def is_non_tpc_mit(rec):
    mit = mit_n(rec)
    # norm_text 會把 non-tpc → non tpc；亦容許 nontpc
    return "non tpc" in mit or "nontpc" in mit

def is_tpc_mit(rec):
    """純 TPC（唔計 Non-TPC）。"""
    mit = mit_n(rec)
    if not mit or is_non_tpc_mit(rec):
        return False
    return bool(re.search(r"\btpc\b", mit))

def is_tree_removal_mit(rec):
    mit = mit_n(rec)
    return any(w in mit for w in ("tree removal", "removal", "remove", "fell", "felling"))

def is_dead_tree_anomaly(rec):
    """Dead-tree-like 但 Mitigation 空白，又或者唔係 Non-TPC／TPC／tree removal。"""
    if not is_dead_tree_like_obs(rec):
        return False
    if is_non_tpc_mit(rec) or is_tpc_mit(rec) or is_tree_removal_mit(rec):
        return False
    return True

def is_delayed_remark(rec):
    raw=str(rec.get("Remarks") or "").lower()
    return any(k in raw for k in ("liaised with emo","traffic obstruction","attempt to carry out work","延後"))

def display_row_extra(rec, key=None, extra=None):
    row = display_row(rec, key)
    if extra:
        row = list(row) + list(extra)
    return row

def make_key(case_no, tree_id):
    c=str(case_no).strip() if case_no is not None else ""
    t=str(tree_id).strip() if tree_id is not None else ""
    return (c,t)

def display_row(rec, key=None):
    return [rec.get("Source of Referral"), rec.get("Estate"), key[0] if key else rec.get("Case No."),
            rec.get("Date Received"), key[1] if key else rec.get("Tree ID"),
            rec.get("Observation"), rec.get("Mitigation"), rec.get("Case deadline")]

def load_sheet(ws, date_from=None, date_to=None):
    rows=ws.iter_rows(values_only=True)
    headers=None
    for _ in range(8):
        try: row=next(rows)
        except StopIteration: return {}
        cand=[norm_header(h) for h in row]
        if "Case No." in cand and "Tree ID" in cand:
            headers=cand; break
    if not headers:
        print(f"  [WARN] no Case No./Tree ID in {ws.title}")
        return {}
    ci, ti = headers.index("Case No."), headers.index("Tree ID")
    records={}
    for row in rows:
        if not row or all(c is None or str(c).strip()=="" for c in row[:8]): continue
        key=make_key(row[ci] if ci<len(row) else None, row[ti] if ti<len(row) else None)
        if key==("",""): continue
        rec={}
        for i,h in enumerate(headers):
            if h: rec[h]=norm_val(row[i] if i<len(row) else None)
        d=rec.get("Date Received")
        if isinstance(d, date):
            if date_from and d < date_from: continue
            if date_to and d > date_to: continue
        elif date_from or date_to:
            continue
        records[key]=rec
    return records

def compare_rec(a,b):
    diffs=[]
    for col in COMPARE_COLS:
        va,vb=a.get(col), b.get(col)
        if vb is None: continue
        if not values_equal(col, va, vb): diffs.append((col, va, vb))
    return diffs

def style_header(ws, row=1):
    fill=PatternFill("solid", fgColor="1F4E79"); font=Font(bold=True, color="FFFFFF", size=11)
    thin=Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    for cell in ws[row]:
        cell.fill, cell.font = fill, font
        cell.alignment=Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border=thin

def auto_width(ws):
    for i in range(1, min(ws.max_column+1, 12)):
        letter=get_column_letter(i)
        m=max((min(len(str(c.value)),55) for c in ws[letter] if c.value), default=10)
        ws.column_dimensions[letter].width=min(max(m+2,10),45)

def sort_rows(rows, by="received", reverse=False):
    idx=7 if by=="deadline" else 3
    def k(r):
        v=r[idx] if len(r)>idx else None
        if isinstance(v, datetime): v=v.date()
        if isinstance(v, date): return (0, v.toordinal(), str(r[2] or ""), str(r[4] or ""))
        return (1,0,str(r[2] or ""), str(r[4] or ""))
    return sorted(rows, key=k, reverse=reverse)

def write_block(ws, title, rows, headers, fill, table_name, sort_by="received", newest_first=True):
    ws.append([title])
    ws.cell(ws.max_row,1).font=Font(bold=True, size=12, color="1F4E79")
    header_row=ws.max_row+1
    ws.append(headers)
    style_header(ws, header_row)
    rows=sort_rows(rows, sort_by, newest_first)
    for row in rows:
        ws.append(row)
        if fill:
            for cell in ws[ws.max_row]: cell.fill=fill
    end=ws.max_row
    if end==header_row:
        ws.append([""]*len(headers)); end=ws.max_row
    ref=f"A{header_row}:{get_column_letter(len(headers))}{end}"
    try:
        tab=Table(displayName=table_name, ref=ref)
        tab.tableStyleInfo=TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(tab)
    except Exception:
        pass
    ws.append([])

def safe_name(s):
    s=re.sub(r"[^A-Za-z0-9_]", "", s)[:12] or "X"
    if s[0].isdigit(): s="N"+s
    return s

def should_skip_sheet(name):
    """Skip M KC / M WTT / Crystal Used maintenance sheets."""
    s = str(name or "").strip()
    if re.match(r"^M\s", s, re.I):
        return True
    if re.search(r"crystal", s, re.I):
        return True
    return False

def load_all(wb, date_from, date_to):
    """Load every sheet that has Case No. + Tree ID. Return {region: records}."""
    out={}
    for name in wb.sheetnames:
        if should_skip_sheet(name):
            continue
        recs=load_sheet(wb[name], date_from, date_to)
        if not recs: continue
        # Summary sheet: split by Region column if present
        if name=="Summary":
            for k,r in recs.items():
                reg=(r.get("Region") or "Summary")
                if should_skip_sheet(reg):
                    continue
                out.setdefault(str(reg), {})[k]=r
        else:
            out[name]=recs
    return out

def run_compare(a_path, b_path, output_path, name_a="A", name_b="B", date_from=None, date_to=None, as_of=None):
    print(f"Compare: {name_a} vs {name_b}")
    print(f"Scope: {date_from or '不限'} 至 {date_to or '不限'}")
    print(f"Overdue as of: {as_of or date.today()}")
    wb_a=load_workbook(a_path, data_only=True, read_only=True)
    wb_b=load_workbook(b_path, data_only=True, read_only=True)
    out=Workbook(); out.remove(out.active)
    a_map=load_all(wb_a, date_from, date_to)
    b_map=load_all(wb_b, date_from, date_to)
    if not a_map and not b_map:
        raise ValueError(f"兩邊都讀唔到 Case No./Tree ID。A={wb_a.sheetnames} B={wb_b.sheetnames}")
    regions=list(dict.fromkeys(list(a_map)+list(b_map)))
    print("A sheets/regions:", list(a_map))
    print("B sheets/regions:", list(b_map))
    loaded={r: dict(a_recs=a_map.get(r, {}), b_recs=b_map.get(r, {})) for r in regions}
    all_a=set(); all_b=set()
    for region in regions:
        a_recs, b_recs = loaded[region]["a_recs"], loaded[region]["b_recs"]
        a_set, b_set = set(a_recs), set(b_recs)
        only_a, only_b = sorted(a_set-b_set), sorted(b_set-a_set)
        common = sorted(a_set & b_set)
        content_diff=[x for x in ((k, compare_rec(a_recs[k], b_recs[k])) for k in common) if x[1]]
        print(f"  [{region}] A {len(a_recs)} B {len(b_recs)} onlyA {len(only_a)} onlyB {len(only_b)} diff {len(content_diff)}")
        all_a |= {k[0] for k in a_set if k[0]}
        all_b |= {k[0] for k in b_set if k[0]}
        loaded[region].update(only_a=only_a, only_b=only_b, content_diff=content_diff)
    na, nb = safe_name(name_a), safe_name(name_b)
    per={}
    for region in regions:
        d=loaded[region]; a_recs,b_recs=d["a_recs"], d["b_recs"]
        new, seen=[], set()
        for key in d["only_b"]:
            if key[0] and key[0] not in all_a and key[0] not in seen:
                seen.add(key[0]); new.append(display_row(b_recs[key], key))
        overdue, pruning_branch, dieback_nontpc, removal_need_tpc, removal_tpc, nontpc_tpc, dead_anom, delayed, risk = (
            [], [], [], [], [], [], [], [], []
        )
        combined=list(b_recs.items())+[(k, a_recs[k]) for k in d["only_a"]]
        for key, rec in combined:
            hits = risk_hits(rec)
            if hits:
                risk.append(display_row(rec, key)+[", ".join(hits), follow_up_status(rec), rec.get("Works Completion Date")])
            if deadline_is_one_day_from_received(rec): continue
            row=display_row(rec, key)
            if is_overdue_incomplete(rec, as_of):
                overdue.append(row)
            if is_dead_branch_obs(rec) and is_pruning_mit(rec):
                pruning_branch.append(row + [completion_status(rec), rec.get("Works Completion Date")])
            if is_dead_tree_like_obs(rec) and is_non_tpc_mit(rec):
                dieback_nontpc.append(row + ["Non-TPC", completion_status(rec), rec.get("Works Completion Date")])
            if is_tree_removal_mit(rec) and not is_tpc_mit(rec) and not is_non_tpc_mit(rec):
                removal_need_tpc.append(row + [completion_status(rec), rec.get("Works Completion Date")])
            elif is_tree_removal_mit(rec) and is_tpc_mit(rec):
                removal_tpc.append(row + ["TPC", completion_status(rec), rec.get("Works Completion Date")])
            if is_non_tpc_mit(rec):
                nontpc_tpc.append(row + ["Non-TPC", completion_status(rec), rec.get("Works Completion Date")])
            elif is_tpc_mit(rec):
                nontpc_tpc.append(row + ["TPC", completion_status(rec), rec.get("Works Completion Date")])
            if is_dead_tree_anomaly(rec):
                dead_anom.append(row + [completion_status(rec), rec.get("Mitigation")])
            if is_delayed_remark(rec):
                delayed.append(row)
        per[region]=dict(
            new=new, overdue=overdue, pruning_branch=pruning_branch,
            dieback_nontpc=dieback_nontpc, removal_need_tpc=removal_need_tpc,
            removal_tpc=removal_tpc, nontpc_tpc=nontpc_tpc, dead_anom=dead_anom,
            delayed=delayed, risk=risk,
        )
        for label, keys, recs in ((f"{region}_Only_{na}", d["only_a"], a_recs),(f"{region}_Only_{nb}", d["only_b"], b_recs)):
            if not keys: continue
            ws=out.create_sheet(label[:31])
            ws.append(["Source of Referral","Estate","Case No.","Date Received","Tree ID","Observation","Mitigation","Case deadline"])
            style_header(ws)
            for key in keys: ws.append(display_row(recs[key], key))
            auto_width(ws)
        if d["content_diff"]:
            ws=out.create_sheet(f"{region}_Diff")
            ws.append(["Case No.","Tree ID","Estate","Date Received","Field", f"{name_a} Value", f"{name_b} Value"])
            style_header(ws)
            yellow=PatternFill("solid", fgColor="FFF2CC")
            for key, diffs in d["content_diff"]:
                estate=a_recs[key].get("Estate") or b_recs[key].get("Estate")
                dr=a_recs[key].get("Date Received") or b_recs[key].get("Date Received")
                for col,va,vb in diffs:
                    ws.append([key[0], key[1], estate, dr, col, "" if va is None else str(va), "" if vb is None else str(vb)])
                    for cell in ws[ws.max_row]: cell.fill=yellow
            auto_width(ws)
    cols=["Source of Referral","Estate","Case No.","Date Received","Tree ID","Observation","Mitigation","Case deadline"]
    cols_done=cols+["有冇完成","Works Completion Date"]
    cols_class=cols+["類別","有冇完成","Works Completion Date"]
    cols_anom=cols+["有冇完成","Mitigation 原文"]
    fills=[PatternFill("solid", fgColor=c) for c in (
        "E2EFDA","FCE4D6","C6EFCE","DDEBF7","F8CBAD","FFF2CC","E2F0D9","FCE4D6","F4B183"
    )]
    first=True
    scope=f"{date_from or '不限'} 至 {date_to or '不限'}" if (date_from or date_to) else "全部日期"
    for region in regions:
        d=per[region]
        ws=out.create_sheet(f"Summary_{region}", 0 if first else None)
        first=False
        ws.append([f"{region}  {name_a} vs {name_b}"])
        ws["A1"].font=Font(bold=True, size=14, color="1F4E79")
        ws.append([f"範圍: {scope}   逾期以 {as_of or date.today()} 計"])
        ws.append([])
        rn=safe_name(region)
        done_n=sum(1 for r in d["pruning_branch"] if len(r)>8 and r[8]=="已完成")
        write_block(ws, f"新 Case（{name_b} 有、{name_a} 沒有）— {len(d['new'])}", d["new"], cols, fills[0], f"{rn}_{nb}New", newest_first=True)
        write_block(ws, f"0. Case deadline 已過且未完成 — {len(d['overdue'])}", d["overdue"], cols, fills[1], f"{rn}Overdue", sort_by="deadline", newest_first=False)
        write_block(ws, f"1. Dead branch／Broken branch + pruning（已完成 {done_n}／共 {len(d['pruning_branch'])}）— {len(d['pruning_branch'])}", d["pruning_branch"], cols_done, fills[2], f"{rn}Prune")
        write_block(ws, f"2. Dead tree／wilted／wilting／dieback + Non-TPC — {len(d['dieback_nontpc'])}", d["dieback_nontpc"], cols_class, fills[3], f"{rn}DieNontpc")
        write_block(ws, f"3. Tree removal 應為 TPC（異常／待標 TPC）— {len(d['removal_need_tpc'])}", d["removal_need_tpc"], cols_done, fills[4], f"{rn}RemNeed")
        write_block(ws, f"3b. Tree removal／TPC — {len(d['removal_tpc'])}", d["removal_tpc"], cols_class, fills[5], f"{rn}RemTpc")
        write_block(ws, f"4. Non-TPC／TPC — {len(d['nontpc_tpc'])}", d["nontpc_tpc"], cols_class, fills[6], f"{rn}TpcClass")
        write_block(ws, f"異常：Dead tree-like 但 Mitigation 未標 Non-TPC／TPC／removal — {len(d['dead_anom'])}", d["dead_anom"], cols_anom, fills[7], f"{rn}DeadAnom")
        write_block(ws, f"延後 — {len(d['delayed'])}", d["delayed"], cols, fills[3], f"{rn}Delayed")
        open_n=sum(1 for r in d["risk"] if r[-2]=="未跟進")
        write_block(ws, f"高風險未跟進 {open_n} / 共 {len(d['risk'])}", [r[:8] for r in d["risk"] if r[-2]=="未跟進"], cols, PatternFill("solid", fgColor="F4B183"), f"{rn}RiskOpen")
        auto_width(ws)
        if d["risk"]:
            ws2=out.create_sheet(f"{region}_Risk")
            headers=cols+["命中字眼","有冇跟進","Works Completion Date"]
            ws2.append(headers); style_header(ws2)
            red=PatternFill("solid", fgColor="F4B183"); green=PatternFill("solid", fgColor="C6EFCE")
            for row in sort_rows(d["risk"], "received", True):
                ws2.append(row)
                fill=red if row[-2]=="未跟進" else green
                for cell in ws2[ws2.max_row]: cell.fill=fill
            auto_width(ws2)
    out.save(output_path)
    print("Done", output_path)

def parse_month(s, end=False):
    s=(s or "").strip().replace("-","").replace("/","")
    if not s: return None
    if len(s)==6 and s.isdigit():
        y,m=int(s[:4]), int(s[4:6])
        if not 1<=m<=12: raise ValueError(f"月份要 01-12：{s}")
        if not end: return date(y,m,1)
        return date(y,12,31) if m==12 else date(y,m+1,1)-timedelta(days=1)
    raise ValueError(f"月份格式要 202607：{s}")

def pick_files():
    print("\n=== 對比設定 ===")
    na=input("A 名稱（預設 A公司）: ").strip() or "A公司"
    nb=input("B 名稱（預設 B公司）: ").strip() or "B公司"
    print("對比時間：1=全部  2=最近3個月  3=自訂月份（202607）")
    choice=input("選擇 [2]: ").strip() or "2"
    d1=d2=None
    if choice=="2":
        d2=date.today(); d1=d2-timedelta(days=90)
    elif choice=="3":
        d1=parse_month(input("開始月份（例如 202607）: ").strip(), False)
        d2=parse_month(input("結束月份（例如 202609）: ").strip(), True)
    asof_s=input("Audit 日期（例如 20261028，留空用今日）: ").strip()
    as_of=None
    if len(asof_s)==8 and asof_s.isdigit():
        as_of=date(int(asof_s[:4]), int(asof_s[4:6]), int(asof_s[6:8]))
    print("\n請選擇 A 公司 Excel...")
    a=win_open_file("選擇 "+na+" Excel")
    if not a: return {}
    print("請選擇 B 公司 Excel...")
    b=win_open_file("選擇 "+nb+" Excel")
    if not b: return {}
    default=f"Diff_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    print("請選擇報告儲存位置...")
    out=win_save_file("儲存報告", default) or str(Path(a).parent/default)
    return dict(a=a,b=b,out=out,na=na,nb=nb,d1=d1,d2=d2,as_of=as_of)

def main():
    print("Housing Referral Compare")
    try:
        if len(sys.argv)>=3:
            a,b=sys.argv[1],sys.argv[2]
            out=sys.argv[3] if len(sys.argv)>=4 else str(Path(a).parent/f"Diff_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")
            na=sys.argv[4] if len(sys.argv)>=5 else "A"
            nb=sys.argv[5] if len(sys.argv)>=6 else "B"
            run_compare(a,b,out,na,nb)
        else:
            r=pick_files()
            if not r: print("Cancelled"); input("Press Enter..."); return
            print(f"A: {r['a']}\nB: {r['b']}\nOut: {r['out']}")
            run_compare(r["a"],r["b"],r["out"],r["na"],r["nb"],r["d1"],r["d2"],r.get("as_of"))
            print(f"\n報告已儲存:\n{r['out']}")
    except Exception as e:
        print("\n=== ERROR ===")
        import traceback; traceback.print_exc()
        print(f"\n錯誤: {e}")
    input("\nPress Enter to exit...")

if __name__=="__main__":
    main()
