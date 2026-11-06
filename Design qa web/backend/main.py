"""Cyncly Design & QA Intelligence Portal - API. Run: uvicorn main:app --reload --port 8000
Every number is computed from the uploaded workbook; source data is never modified (raw copy kept)."""
import io, re, json, uuid, datetime as dt
from pathlib import Path
import numpy as np, pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter as LET
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, PlainTextResponse
from fastapi.staticfiles import StaticFiles

app = FastAPI(title="Cyncly Design & QA Intelligence API")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])
SESS, HIST = {}, Path(__file__).parent / "history.json"
FRONT = Path(__file__).parent.parent / "frontend"

cl = lambda h: re.sub(r"\s+", " ", str(h).replace("\xa0", " ")).strip()
num = lambda s: pd.to_numeric(s, errors="coerce")
dte = lambda s: pd.to_datetime(s, errors="coerce")
ratio = lambda a, b: None if not b else round(float(a) / float(b), 4)

# short key -> exact (cleaned) header in the workbook
C = dict(name="Project Name", mfr="Manufacturer Name", type="Type of Catalog", code="Project Codes", pos="Region & POS",
 contract="Contract Type", work="Work Type", am="Account Manager", owner="Project Owner", phase="Phase", est="Estimated Hours",
 cons="Consumption hours", start="Project Start Date", prop="Proposed Delivery Date to Customer", track="Project On Track",
 why="Projects Not on Track - Comment", act="Actual Delivery Date to Customer", iter="Total Number of iterations till acceptance",
 pdreq="Total Number of Post Delivery Request Received", corr="Total Number of Correction", chg="Total Number of Change Request",
 add="Total Number of New addition", nai="Total Number of Not an issue", sw="Software limitation",
 acc="Acceptance received (YES/NO)", accd="Acceptance received Date", qareq="QA Required at Cyncly", qah="QA hours",
 qd1="Proposed Delivery Date to QA - 1st Draft", qs1="QA Status - 1st Draft", qs2="QA Status - 2nd Draft",
 r1="Total No. of Product/Update Requested by Client - 1st Draft", t1="Total No. of Product/Update Testing Done By QA - 1st Draft",
 ok1="Number of OK Product/Update - 1st Draft", no1="Number of Not OK Product/Update - 1st Draft",
 r2="Total No. of Product/Update Requested - 2nd Draft", t2="Total No. of Product/Update Testing Done By QA - 2nd Draft",
 ok2="Number of OK Product/Update - 2nd Draft", no2="Number of Not OK Product/Update - 2nd Draft",
 qry1="Number of 1st Draft QA Queries", bu="Quality Check After Production - 1st Draft", rep="QA Report Location",
 person="Responsible QA/Self QA Person", q1t="Number Of Product/Feedback Tested in Q1", q1ok='Number Of Product/Feedback "OK" in Q1',
 q1pct="Catalog Quality in Q1 %", pdt="Total No. of Post Delivery Updates Testing Done By QA", pds="Post Delivery QA Status")
GATES = ["Product Commercial Data and Graphical Data Issue", "Add ons Issue", "Global Features and Option Issue", "Catalog Pricing Issue", "POS Specific Requirement Issue"]
G1 = [f"1st Draft Gate {i+1} - {g}" for i, g in enumerate(GATES)] + ["1st Draft - Combination Gate Issue"]
GP = [f"Post Delivery Gate {i+1} - {g}" for i, g in enumerate(GATES)] + ["Post Delivery Combination Gate Issue"]
GL = [f"Gate {i+1}: {g.replace(' Issue','')}" for i, g in enumerate(GATES)] + ["Combination"]
LISTMAP = {"Contract Type": "Contract Type", "Work Type": "Work Type", "Region & POS": "Region & POS", "Phase": "Phase", "Account Manager": "Account Manager",
 "Backup Uploaded On VSTS": "Backup Uploaded On VSTS", "Type of Catalog": "Catalog Type", "QA Required at Cyncly": "QA Required at 2020 (Yes/No)",
 "QA Status - 1st Draft": "QA Status - 1st Draft", "QA Status - 2nd Draft": "QA Status - 2nd Draft", "Project On Track": "Project On Track", "Project Owner": "Project Owner"}
STATUS = lambda p: {"06-Live": "Completed", "08-On-Hold": "On Hold", "10-Retired": "Retired"}.get(p, "In Progress")

# ---------------------------------------------------------------- load
def load(raw: bytes, filename: str):
    wbv = load_workbook(io.BytesIO(raw), data_only=True); wbf = load_workbook(io.BytesIO(raw), data_only=False)
    if "Design_Dashboard" not in wbv.sheetnames: raise HTTPException(422, "Sheet 'Design_Dashboard' not found - not a valid weekly workbook.")
    ws, wf = wbv["Design_Dashboard"], wbf["Design_Dashboard"]
    rows = list(ws.iter_rows(values_only=True)); frows = list(wf.iter_rows(values_only=True))
    ncol = max(i for i, h in enumerate(rows[0]) if h is not None or any(r[i] is not None for r in rows[2:] if i < len(r))) + 1
    heads = [cl(h) if h is not None else f"(unnamed {LET(i+1)})" for i, h in enumerate(rows[0][:ncol])]
    body = [(i + 1, list(r[:ncol])) for i, r in enumerate(rows) if i >= 1]
    tmpl = [n for n, r in body if str(r[0] or "").startswith("COPY & PASTE")]
    body = [(n, r) for n, r in body if n not in tmpl and r[0] not in (None, "")]
    raw_df = pd.DataFrame([r for _, r in body], columns=heads); raw_df.insert(0, "_xl_row", [n for n, _ in body])
    fcount = {h: sum(1 for n, _ in body if isinstance(frows[n-1][i] if i < len(frows[n-1]) else None, str) and str(frows[n-1][i]).startswith("="))
              for i, h in enumerate(heads)}
    notes = {}
    if "Notes" in wbv.sheetnames:
        for r in wbv["Notes"].iter_rows(min_row=2, values_only=True):
            if r[1]: notes[cl(r[1])] = dict(letter=r[0], owner=r[2], intent=r[3])
    lists = {}
    if "List" in wbv.sheetnames:
        L = list(wbv["List"].iter_rows(values_only=True)); hd = [cl(x) if x else None for x in L[0]]
        for j, h in enumerate(hd):
            if h: lists[h] = {cl(r[j]) for r in L[1:] if j < len(r) and r[j] not in (None, "")}
    extra = dict(est_rows=(wbv["Estimation Matrix"].max_row - 1) if "Estimation Matrix" in wbv.sheetnames else None,
                 pivot_heads=[cl(c.value) for c in wbv["Pivot Table"][4] if c.value] if "Pivot Table" in wbv.sheetnames else [],
                 sheets=wbv.sheetnames, template_rows=len(tmpl), filename=filename)
    return raw_df, fcount, notes, lists, extra

def prep(raw_df):
    d = raw_df.copy()
    for k in ["name", "mfr", "type", "code", "pos", "contract", "work", "am", "owner", "phase", "track", "acc", "qareq", "qs1", "qs2", "person"]:
        d[C[k]] = d[C[k]].map(lambda v: cl(v) if isinstance(v, str) else v)
    for k in ["est", "cons", "qah", "r1", "t1", "ok1", "no1", "r2", "t2", "ok2", "no2", "q1t", "q1ok", "iter", "pdreq", "corr", "chg", "add", "nai", "sw", "qry1", "pdt"] + []:
        d[C[k]] = num(d[C[k]])
    for g in G1 + GP: d[g] = num(d[g])
    for k in ["start", "prop", "act", "accd", "qd1"]: d[C[k]] = dte(d[C[k]])
    d["_status"] = d[C["phase"]].map(STATUS)
    d["_late"] = (d[C["act"]] - d[C["prop"]]).dt.days
    # Quality is ALWAYS recomputed from counts. Column BU holds a literal '95%' placeholder on 268 rows and is ignored.
    d["_q1pct"] = np.where(d[C["q1t"]] > 0, d[C["q1ok"]] / d[C["q1t"]], np.nan)
    d["_fdq"] = np.where(d[C["t1"]] > 0, d[C["ok1"]] / d[C["t1"]], np.nan)
    known = {cl(v).title() for v in d[C["person"]].dropna() if not re.search(r"[/\\,&]", str(v))}
    def split(v):
        if pd.isna(v): return []
        ps = [cl(p).title() for p in re.split(r"[/\\,&]", str(v)) if cl(p)]
        return [p if p in known else "Unresolved (abbreviated names)" for p in ps]
    d["_people"] = d[C["person"]].map(split)
    return d

# ---------------------------------------------------------------- helpers
def J(o):
    if isinstance(o, dict): return {str(k): J(v) for k, v in o.items()}
    if isinstance(o, (list, tuple)): return [J(v) for v in o]
    if isinstance(o, (bool, np.bool_)): return bool(o)
    if isinstance(o, (np.integer,)): return int(o)
    if isinstance(o, (np.floating, float)): return None if pd.isna(o) else round(float(o), 4)
    if isinstance(o, (pd.Timestamp, dt.datetime, dt.date)): return None if pd.isna(o) else str(o)[:10]
    return None if o is pd.NaT else o

def vc(s, top=None):
    s = s.dropna().map(lambda v: cl(v) if isinstance(v, str) else v); s = s[s != ""]
    r = s.value_counts(); r = r.head(top) if top else r
    return [dict(name=str(k), value=int(v)) for k, v in r.items()]

def month(s): return s.dt.to_period("M").astype(str)

def qa_totals(d, t, ok):
    m = d[d[C[t]].notna() & d[C[ok]].notna()]
    T, O = m[C[t]].sum(), m[C[ok]].sum()
    return dict(entries=len(m), tested=T, ok=O, not_ok=T - O, quality=ratio(O, T))

def filtered(d, owner, pos, f, t):
    if owner: d = d[d[C["owner"]].isin(owner.split("|"))]
    if pos: d = d[d[C["pos"]].isin(pos.split("|"))]
    if f: d = d[d[C["start"]] >= pd.Timestamp(f + "-01")]
    if t: d = d[d[C["start"]] < (pd.Timestamp(t + "-01") + pd.offsets.MonthBegin(1))]
    return d

# ---------------------------------------------------------------- analytics
def compute(d):
    A = {}
    n = len(d); ph = d["_status"].value_counts()
    fd, q1 = qa_totals(d, "t1", "ok1"), qa_totals(d, "q1t", "q1ok")
    trk = d[C["track"]].value_counts(); acc = d[C["acc"]].value_counts()
    A["kpi"] = dict(entries=n, projects=d[C["code"]].nunique(), manufacturers=d[C["mfr"]].nunique(),
        completed=int(ph.get("Completed", 0)), in_progress=int(ph.get("In Progress", 0)), on_hold=int(ph.get("On Hold", 0)),
        delivered=int(d[C["act"]].notna().sum()), est_hours=d[C["est"]].sum(), cons_hours=d[C["cons"]].sum(),
        hour_util=ratio(d[C["cons"]].sum(), d[C["est"]].sum()), on_track_pct=ratio(trk.get("Yes", 0), trk.get("Yes", 0) + trk.get("No", 0)),
        late_pct=ratio((d["_late"] > 0).sum(), d["_late"].notna().sum()), accept_pct=ratio(acc.get("Yes", 0), acc.get("Yes", 0) + acc.get("No", 0)),
        qa_hours=d[C["qah"]].sum(), qa_hours_entries=int(d[C["qah"]].notna().sum()),
        internal_qa_entries=int((d[C["qareq"]] == "QA Required").sum()), self_qa_entries=int((d[C["qareq"]] == "Delivered with Self QA Only").sum()),
        q1_entries=int((d[C["qareq"]] == "Delivered with Q1 QA Only").sum()),
        fd_tested=fd["tested"], fd_ok=fd["ok"], fd_not_ok=fd["not_ok"], fd_quality=fd["quality"], fd_entries=fd["entries"],
        q1_tested=q1["tested"], q1_ok=q1["ok"], q1_quality=q1["quality"], q1_rows=q1["entries"],
        pd_requests=d[C["pdreq"]].sum(), corrections=d[C["corr"]].sum())
    # trends (clip to 2025-01+ to keep legacy start dates from stretching axes; excluded count reported)
    def mtr(col, name):
        s = d[d[C[col]] >= "2025-01-01"][C[col]]; return month(s).value_counts().sort_index()
    ms, md, ma = mtr("start", "s"), mtr("act", "d"), mtr("accd", "a")
    months = sorted(set(ms.index) | set(md.index) | set(ma.index))
    A["monthly"] = dict(months=months, started=[int(ms.get(m, 0)) for m in months], delivered=[int(md.get(m, 0)) for m in months], accepted=[int(ma.get(m, 0)) for m in months])
    q = d[d[C["t1"]].notna() & d[C["ok1"]].notna() & d[C["qd1"]].notna()].copy(); q["m"] = month(q[C["qd1"]]); q = q[q["m"] >= "2025-01"]
    g = q.groupby("m").agg(t=(C["t1"], "sum"), o=(C["ok1"], "sum"), n=(C["t1"], "size")).reset_index()
    A["qa_monthly"] = dict(months=list(g["m"]), tested=list(g["t"]), ok=list(g["o"]), n=list(g["n"]), quality=[ratio(o, t) for o, t in zip(g["o"], g["t"])])
    wk = d[d[C["start"]] >= d[C["start"]].max() - pd.Timedelta(weeks=15)][C["start"]].dt.to_period("W-SUN").dt.start_time.dt.strftime("%d %b").value_counts()
    wk_order = d[d[C["start"]] >= d[C["start"]].max() - pd.Timedelta(weeks=15)][C["start"]].dt.to_period("W-SUN").dt.start_time.sort_values().dt.strftime("%d %b").unique()
    A["weekly"] = dict(weeks=list(wk_order), entries=[int(wk.get(w, 0)) for w in wk_order])
    # design & production
    A["design"] = dict(status=vc(d["_status"]), phase=vc(d[C["phase"]]), type=vc(d[C["type"]]), pos=vc(d[C["pos"]]), contract=vc(d[C["contract"]]), work=vc(d[C["work"]]),
        delay_buckets=[dict(name=k, value=int(v)) for k, v in pd.cut(d["_late"].dropna(), [-9999, -1, 0, 3, 7, 14, 9999], labels=["Early", "On time", "1-3 days late", "4-7 days late", "8-14 days late", ">14 days late"]).value_counts().sort_index().items()],
        avg_days_late=d.loc[d["_late"] > 0, "_late"].mean(), not_on_track_reasons=vc(d[C["why"]].map(lambda v: cl(v).lower().capitalize() if isinstance(v, str) else v), 10),
        feedback=[dict(name=nm, value=float(d[C[k]].sum())) for nm, k in [("Corrections", "corr"), ("Change requests", "chg"), ("New additions", "add"), ("Not an issue", "nai"), ("Software limitation", "sw")]],
        customer_gates=[dict(name=GL[i], value=float(d[GP[i]].sum())) for i in range(6)])
    ov = d[(d[C["est"]] >= 2) & (d[C["cons"]] / d[C["est"]] > 1.2)].assign(over=lambda x: x[C["cons"]] / x[C["est"]]).sort_values("over", ascending=False).head(10)
    A["design"]["overruns"] = [dict(name=r[C["name"]], owner=r[C["owner"]], est=r[C["est"]], cons=r[C["cons"]], pct=r["over"]) for _, r in ov.iterrows()]
    # internal QA
    sg = lambda cols, lab: [dict(name=lab[i], value=float(d[c].sum()), n=int(d[c].notna().sum())) for i, c in enumerate(cols)]
    fd2 = qa_totals(d, "t2", "ok2"); both = d[d[C["t1"]].notna() & d[C["t2"]].notna()]
    rw = d[(d["_fdq"] < 0.95)].sort_values("_fdq").head(15)
    A["iqa"] = dict(first=fd, second=fd2, second_requested=d[C["r2"]].sum(), first_requested=d[C["r1"]].sum(), post_delivery_populated=int(d[C["pdt"]].notna().sum()),
        gates_first=sg(G1, GL), gates_customer=sg(GP, GL), status1=vc(d[C["qs1"]], 8), status2=vc(d[C["qs2"]]), routing=vc(d[C["qareq"]]),
        queries=d[C["qry1"]].sum(), report_available=int(d[C["rep"]].notna().sum()), tested_rows=fd["entries"], qa_hours=d[C["qah"]].sum(),
        rework_n=int((d["_fdq"] < 0.95).sum()), rework=[dict(name=r[C["name"]], owner=r[C["owner"]], qa=r[C["person"]], tested=r[C["t1"]], ok=r[C["ok1"]], q=r["_fdq"]) for _, r in rw.iterrows()],
        both_drafts=len(both), draft_q_threshold=0.95)
    # "External QA": the workbook has no external-QA field group. Report what exists, labelled honestly.
    cm = d[d["_q1pct"].notna() & d["_fdq"].notna()]
    A["ext"] = dict(available=False, q1=q1, routing=vc(d[C["qareq"]]), q1_by_routing=[dict(name=k, tested=v[C["q1t"]].sum(), ok=v[C["q1ok"]].sum(), rows=int(v[C["q1t"]].notna().sum()), quality=ratio(v[C["q1ok"]].sum(), v[C["q1t"]].sum()))
        for k, v in d.groupby(C["qareq"]) if v[C["q1t"]].notna().any()],
        overlap=dict(rows=len(cm), q1=ratio(cm[C["q1ok"]].sum(), cm[C["q1t"]].sum()), fd=ratio(cm[C["ok1"]].sum(), cm[C["t1"]].sum()), tested_q1=cm[C["q1t"]].sum(), tested_fd=cm[C["t1"]].sum()),
        cic=int((d[C["qareq"]] == "CiC QA required").sum()))
    # resources
    def own(o, v):
        f1, fq = qa_totals(v, "t1", "ok1"), qa_totals(v, "q1t", "q1ok"); t = v["_late"].notna().sum(); tr = v[C["track"]].value_counts()
        return dict(name=o, entries=len(v), est=v[C["est"]].sum(), cons=v[C["cons"]].sum(), util=ratio(v[C["cons"]].sum(), v[C["est"]].sum()), delivered=int(v[C["act"]].notna().sum()),
            completed=int((v["_status"] == "Completed").sum()), on_hold=int((v["_status"] == "On Hold").sum()), in_progress=int((v["_status"] == "In Progress").sum()),
            on_track=ratio(tr.get("Yes", 0), tr.get("Yes", 0) + tr.get("No", 0)), late_pct=ratio((v["_late"] > 0).sum(), t), pd_req=v[C["pdreq"]].sum(), corr=v[C["corr"]].sum(),
            avg_iter=v[C["iter"]].mean(), fd_tested=f1["tested"], fd_ok=f1["ok"], fd_quality=f1["quality"], q1_tested=fq["tested"], q1_quality=fq["quality"], low_volume=len(v) < 10)
    A["owners"] = sorted([own(o, v) for o, v in d.groupby(C["owner"])], key=lambda x: -x["entries"])
    hm = d[d[C["start"]] >= "2026-01-01"].assign(m=lambda x: month(x[C["start"]])); mm = sorted(hm["m"].unique())
    A["owner_heat"] = dict(months=mm, owners=[o["name"] for o in A["owners"]], data=[[int(((hm[C["owner"]] == o["name"]) & (hm["m"] == m)).sum()) for m in mm] for o in A["owners"]])
    ex = d.explode("_people"); ex = ex[ex["_people"].notna()]; ex["w"] = 1 / ex["_people"].map(lambda x: 1) if False else 1
    ex["w"] = 1 / d.loc[ex.index, "_people"].map(len)
    qrows = []
    for p, v in ex.groupby("_people"):
        w = v["w"]; tt = (v[C["t1"]] * w); oo = (v[C["ok1"]] * w); mk = v[C["t1"]].notna() & v[C["ok1"]].notna()
        qrows.append(dict(name=p, entries=w.sum(), shared=int((w < 1).sum()), qa_entries=int((v[C["qareq"]] == "QA Required").sum()), self_entries=int((v[C["qareq"]] == "Delivered with Self QA Only").sum()),
            tested=tt[mk].sum(), ok=oo[mk].sum(), quality=ratio(oo[mk].sum(), tt[mk].sum()), qa_hours=(v[C["qah"]] * w).sum(), reports=int(v[C["rep"]].notna().sum()), low_volume=w.sum() < 10))
    A["qa_people"] = sorted(qrows, key=lambda x: -x["entries"])
    A["am"] = vc(d[C["am"]], 12)
    # quality & catalog
    def by(col):
        out = []
        for k, v in d.groupby(C[col]):
            a, b = qa_totals(v, "t1", "ok1"), qa_totals(v, "q1t", "q1ok")
            if a["entries"] or b["entries"]: out.append(dict(name=k, fd_tested=a["tested"], fd_quality=a["quality"], q1_tested=b["tested"], q1_quality=b["quality"]))
        return out
    A["quality"] = dict(by_type=by("type"), by_pos=by("pos"), by_contract=by("contract"), fd=fd, q1=q1, customer=dict(
        delivered=int(d[C["act"]].notna().sum()), with_feedback=int(d[C["pdreq"]].notna().sum()), corrections=d[C["corr"]].sum(), iter_gt0=int((d[C["iter"]] > 0).sum()), iter_known=int(d[C["iter"]].notna().sum())))
    # explorer records
    cols = dict(name="name", mfr="mfr", type="type", pos="pos", contract="contract", work="work", owner="owner", am="am", phase="phase", est="est", cons="cons", track="track", acc="acc", qareq="qareq", qs1="qs1", person="person", r1="r1", t1="t1", ok1="ok1", no1="no1", q1t="q1t", q1ok="q1ok", qah="qah", pdreq="pdreq", corr="corr", iter="iter", why="why", rep="rep")
    rec = pd.DataFrame({k: d[C[v]] for k, v in cols.items()})
    for k in ["start", "prop", "act"]: rec[k] = d[C[k]]
    rec["status"], rec["late"], rec["fdq"], rec["q1q"], rec["row"] = d["_status"], d["_late"], d["_fdq"], d["_q1pct"], d["_xl_row"]
    A["records"] = rec.to_dict("records")
    return A

def profile(raw_df, fcount, notes):
    out = []
    for i, h in enumerate([c for c in raw_df.columns if c != "_xl_row"]):
        s = raw_df[h].replace("", np.nan); nn = s.dropna(); n = len(nn)
        kind = "empty" if n == 0 else "date" if all(isinstance(v, (dt.datetime, pd.Timestamp)) for v in nn) else "number" if all(isinstance(v, (int, float, np.number)) and not isinstance(v, bool) for v in nn) else "mixed" if any(isinstance(v, (int, float)) for v in nn) and any(isinstance(v, str) for v in nn) else "text"
        if re.search(r"^(Responsible QA|Account Manager|Project Owner)", h): cat = "Resource"
        elif "Q1" in h: cat = "Q1 QA"
        elif re.search(r"Post Delivery Feedback|^Total Number of (Post|Correction|Change|New|Not an)|^Software|^Post Delivery (Gate|Combination)", h): cat = "Post-delivery feedback"
        elif re.search(r"QA|Draft|Quality Check|^CG$", h): cat = "Internal QA"
        elif re.search(r"Hours|hours|Consumption|Completion", h): cat = "Hours & progress"
        elif re.search(r"Delivery|Warranty|Acceptance|Publication|closed|Backup|Fixes", h): cat = "Delivery & acceptance"
        elif "Date" in h or "iteration" in h: cat = "Schedule"
        elif re.search(r"Comment", h): cat = "Comments"
        else: cat = "Project"
        dist = s.dropna().astype(str).nunique()
        role = "Not usable (empty)" if kind == "empty" else "Date - trend axis" if kind == "date" else "KPI / measure" if kind == "number" else "Dimension / filter" if dist <= 25 else "Descriptor / drilldown"
        out.append(dict(letter=LET(i + 1), header=h, category=cat, dtype=kind, populated=n, pct=round(n / len(raw_df), 3), formulas=fcount.get(h, 0), distinct=dist, role=role,
                        meaning=(notes.get(h) or {}).get("intent") or "Not documented in Notes sheet", notes_letter=(notes.get(h) or {}).get("letter")))
    return out

def quality_report(raw_df, d, lists, notes, extra, prof, fcount):
    W, I = [], []
    def add(sev, area, msg, n=None): W.append(dict(severity=sev, area=area, message=msg, count=n))
    empty = [p for p in prof if p["populated"] == 0]
    if empty: add("info", "Missing fields", "Columns with headers but zero data: " + ", ".join(f"{p['letter']} {p['header']}" for p in empty), len(empty))
    bu = raw_df[C["bu"]].map(lambda v: isinstance(v, str) and v.strip() == "95%").sum()
    if bu: add("warn", "Placeholder values", f"'{C['bu']}' contains a literal '95%' text on {bu} rows (no QA counts behind it). Ignored - quality is recomputed as OK / tested.", int(bu))
    err = int(sum(raw_df[c].map(lambda v: isinstance(v, str) and v.startswith("#")).sum() for c in raw_df.columns))
    if err: add("warn", "Formula errors", f"{err} cells hold Excel errors (e.g. #DIV/0!). Treated as missing, never as 0.", err)
    t, o, nk = d[C["t1"]], d[C["ok1"]], d[C["no1"]]; m = t.notna()
    r = int((m & (o + nk.fillna(0) != t)).sum()); add("warn", "QA reconciliation", f"{r} first-draft rows where OK + Not OK != Tested ({int((m & nk.isna()).sum())} lack a Not OK value). Dashboard derives Not OK = Tested - OK.", r) if r else None
    x = int((t > d[C["r1"]]).sum()); add("warn", "QA reconciliation", "Tested exceeds requested on first-draft rows.", x) if x else None
    nd = int(((d[C["acc"]] == "No") & d[C["accd"]].notna()).sum()); add("warn", "Inconsistent values", "Acceptance flag = 'No' but an acceptance date exists (flag is formula-driven; check its logic).", nd) if nd else None
    nd = int((d[C["start"]] < "2024-01-01").sum()); add("warn", "Outlier dates", f"Project Start Date before 2024 (earliest {d[C['start']].min().date()}). Excluded from monthly trends only.", nd) if nd else None
    dup = int(raw_df.duplicated([C["name"], C["code"], C["start"]]).sum()); add("warn", "Duplicate records", "Rows identical on Project Name + Code + Start Date.", dup) if dup else None
    rp = d[C["name"]].duplicated().sum(); I.append(f"{rp} rows repeat a project name with different dates/hours - treated as separate update entries, not duplicates.")
    inv = []
    for col, lh in LISTMAP.items():
        if col in raw_df and lh in lists:
            bad = raw_df[col].dropna().map(cl).loc[lambda s: ~s.isin(lists[lh]) & (s != "")]
            if len(bad): inv.append(dict(field=col, count=int(len(bad)), values=bad.value_counts().head(6).to_dict()))
    if sum(x["count"] for x in inv): add("warn", "Invalid values", "Values not found in the List sheet (see table).", sum(x["count"] for x in inv))
    hs = [p["header"] for p in prof]; nh = set(notes); miss = sorted(nh - set(hs)); new = sorted(set(hs) - nh - {p["header"] for p in empty if p["header"].startswith("(unnamed")})
    drift = sum(1 for p in prof if p["notes_letter"] and p["notes_letter"] != p["letter"])
    if drift: add("warn", "Schema drift", f"Notes sheet column letters no longer match the Design_Dashboard sheet for {drift} columns. The portal maps by header name, not letter.", drift)
    if new: add("info", "Schema changes", f"{len(new)} columns are in the sheet but not documented in Notes.", len(new))
    q2 = [h for h in extra["pivot_heads"] if "Q2" in h]
    if q2: add("warn", "Schema drift", f"Pivot Table references '{q2[0]}' but no Q2 column exists in Design_Dashboard.", 1)
    nz = int(((d[C["est"]] <= 0) | d[C["est"]].isna()).sum()); add("info", "Hours", "Entries without positive estimated hours (excluded from % utilisation per row).", nz) if nz else None
    mis = int(((d[C["cons"]] / d[C["est"]]).round(4) != raw_df.loc[d.index, "% of Hour Consumption"].pipe(num).round(4)).sum()); I.append(f"Independent check of '% of Hour Consumption' vs Consumption/Estimated: {mis} rows differ.")
    return dict(warnings=W, notes=I, invalid=inv, missing_in_sheet=miss, undocumented=new, empty=[f"{p['letter']} {p['header']}" for p in empty], rows=len(raw_df), sheets=extra["sheets"], estimation_rows=extra["est_rows"], template_rows=extra["template_rows"])

# ---------------------------------------------------------------- endpoints
def hist_load(): return json.loads(HIST.read_text()) if HIST.exists() else []

@app.post("/api/upload")
async def upload(file: UploadFile = File(...)):
    if not file.filename.lower().endswith((".xlsx", ".xlsm")): raise HTTPException(400, "Please upload an .xlsx workbook.")
    raw = await file.read()
    try: raw_df, fcount, notes, lists, extra = load(raw, file.filename)
    except HTTPException: raise
    except Exception as e: raise HTTPException(422, f"Could not read workbook: {e}")
    d = prep(raw_df); prof = profile(raw_df, fcount, notes); dq = quality_report(raw_df, d, lists, notes, extra, prof, fcount)
    tok = uuid.uuid4().hex[:12]; SESS[tok] = dict(d=d, raw=raw_df, prof=prof)
    meta = dict(filename=file.filename, uploaded=dt.datetime.now().strftime("%Y-%m-%d %H:%M"), rows=len(raw_df), warnings=len(dq["warnings"]), size_kb=round(len(raw) / 1024))
    h = hist_load() + [meta]; HIST.write_text(json.dumps(h[-25:]))
    opts = dict(owners=sorted(d[C["owner"]].dropna().unique()), pos=sorted(d[C["pos"]].dropna().unique()), min=str(d[C["start"]][d[C["start"]] >= "2025-01-01"].min())[:7], max=str(d[C["start"]].max())[:7])
    return J(dict(token=tok, meta=meta, options=opts, schema=prof, dq=dq, history=h[::-1], dash=compute(d)))

@app.get("/api/dashboard")
def dashboard(token: str, owner: str = "", pos: str = "", frm: str = "", to: str = ""):
    if token not in SESS: raise HTTPException(404, "Session expired - re-upload the workbook.")
    d = filtered(SESS[token]["d"], owner, pos, frm, to)
    if d.empty: raise HTTPException(404, "No records match the filters.")
    return J(compute(d))

@app.get("/api/schema.csv")
def schema_csv(token: str):
    if token not in SESS: raise HTTPException(404, "Session expired")
    return PlainTextResponse(pd.DataFrame(SESS[token]["prof"]).to_csv(index=False), media_type="text/csv", headers={"Content-Disposition": "attachment; filename=column_mapping.csv"})

@app.get("/api/history")
def history(): return hist_load()[::-1]

app.mount("/assets", StaticFiles(directory=FRONT), name="assets")
@app.get("/")
def index(): return FileResponse(FRONT / "index.html")
