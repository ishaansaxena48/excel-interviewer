# app.py
"""
Excel Mock Interviewer — Rule-based with improved UX & reporting
Features added:
- Intro card (on Start)
- Confidence flags (High / Medium / Low)
- Human-readable feedback templates
- Review queue for low-confidence answers
- Enhanced final report (Strengths, Areas to Improve, Next Steps)
- Hands-on CSV/XLSX pivot validation for the hands-on question
- 10 questions, Submit/Skip only
Run:
pip install streamlit pandas openpyxl
streamlit run app.py
"""

import re
import json
from datetime import datetime
from dateutil.relativedelta import relativedelta
import streamlit as st
import pandas as pd

# -------------------------
# Basic page styling
# -------------------------
st.set_page_config(page_title="Excel Mock Interviewer", layout="wide")
st.markdown(
    """
    <style>
      .stApp { background: linear-gradient(180deg,#fbfdff,#ffffff); }
      .topbar { display:flex; align-items:center; gap:12px; margin-bottom:12px; }
      .brand { width:52px; height:52px; border-radius:10px; background:linear-gradient(135deg,#2563eb,#06b6d4); color:white; display:flex; align-items:center; justify-content:center; font-weight:700; font-size:18px; }
      .title { font-size:20px; font-weight:700; margin:0; }
      .subtitle { color:#475569; font-size:13px; margin-top:3px; }
      .card { background: white; border-radius:12px; padding:16px; box-shadow: 0 8px 24px rgba(2,6,23,0.06); border: 1px solid rgba(15,23,42,0.04); margin-bottom:12px; }
      .muted { color:#6b7280; font-size:13px; }
      .badge { background:#eef2ff; color:#1e3a8a; padding:6px 10px; border-radius:999px; font-weight:600; }
      .small { font-size:13px; color:#6b7280; }
      .conf-high { color: #065f46; font-weight:700; }
      .conf-med { color: #b45309; font-weight:700; }
      .conf-low { color: #b91c1c; font-weight:700; }
      pre.json { background:#f1f5f9; padding:10px; border-radius:8px; overflow:auto }
    </style>
    """,
    unsafe_allow_html=True,
)

# -------------------------
# Question bank (10)
# -------------------------
QUESTIONS = [
    {"id":"q1","type":"objective","prompt":"Write an Excel formula that returns the SUM of column B only for rows where column A equals 'Sales'.","example_answer":"=SUMIFS(B:B, A:A, \"Sales\")","weight":2},
    {"id":"q2","type":"objective","prompt":"Give a formula to lookup value in C2 from a table where key is in column A, robust to column re-ordering.","example_answer":"XLOOKUP(C2, A:A, B:B)","weight":2},
    {"id":"q3","type":"objective","prompt":"Write an Excel formula to count how many cells in range D2:D100 contain text.","example_answer":"=COUNTIF(D2:D100,\"*\")","weight":2},
    {"id":"q4","type":"objective","prompt":"Write a formula to calculate the average of numbers in column E, ignoring blanks.","example_answer":"=AVERAGE(E:E)","weight":2},
    {"id":"q5","type":"objective","prompt":"Write a formula that extracts the year from a date in cell F2.","example_answer":"=YEAR(F2)","weight":1},
    {"id":"q6","type":"debug","prompt":"You get #DIV/0! error in Excel. What are two common causes?","example_answer":"Dividing by zero; blank denominator.","weight":2},
    {"id":"q7","type":"debug","prompt":"Your SUM formula =SUM(A1:A10) is returning 0 even though numbers are visible. List likely causes.","example_answer":"Cells formatted as text; hidden characters/spaces; not actually numeric values.","weight":3},
    {"id":"q8","type":"concept","prompt":"Explain difference between Absolute ($A$1) and Relative (A1) references in formulas.","example_answer":"Absolute stays fixed when copied; Relative shifts with cell position.","weight":2},
    {"id":"q9","type":"concept","prompt":"Explain difference between VLOOKUP and XLOOKUP in Excel.","example_answer":"XLOOKUP is more flexible, can search both directions, no need for column index, handles errors.","weight":3},
    {"id":"q10","type":"hands_on","prompt":"Upload a CSV/XLSX with Date and Sales columns. How would you create a chart to show monthly sales trend?","example_answer":"Insert → Line chart, set X=Month(Date), Y=Sum(Sales).","weight":4}
]

# -------------------------
# Helpers & rule-based checks
# -------------------------
def normalize(s): 
    return re.sub(r"\s+", "", (s or "")).upper()

def check_sumifs(ans):
    a = normalize(ans)
    if "SUMIFS" in a: return 1.0, ["Uses SUMIFS"]
    if "SUMIF" in a: return 0.9, ["Uses SUMIF (acceptable)"]
    return 0.0, ["SUMIF(S) not detected"]

def check_xlookup(ans):
    a = normalize(ans)
    if "XLOOKUP" in a or ("INDEX" in a and "MATCH" in a): return 1.0, ["Uses XLOOKUP / INDEX+MATCH"]
    return 0.0, ["XLOOKUP / INDEX+MATCH not detected"]

def check_countif(ans):
    a = normalize(ans)
    if "COUNTIF" in a or "COUNTA" in a: return 1.0, ["Uses COUNTIF/COUNTA"]
    return 0.0, ["COUNTIF/COUNTA not detected"]

def check_average(ans):
    a = normalize(ans)
    if "AVERAGE" in a or "AVERAGEIF" in a: return 1.0, ["Uses AVERAGE/AverageIf"]
    return 0.0, ["AVERAGE not detected"]

def check_year(ans):
    a = normalize(ans)
    if "YEAR(" in a: return 1.0, ["Uses YEAR()"]
    return 0.0, ["YEAR() not detected"]

def check_div0(ans):
    t = (ans or "").lower()
    hits = 0
    if "zero" in t or "divide" in t: hits += 1
    if "blank" in t or "empty" in t: hits += 1
    return min(1.0, hits/2.0), [f"matched {hits} causes"]

def check_sum_zero(ans):
    t = (ans or "").lower()
    hits = 0
    if "text" in t: hits += 1
    if "format" in t or "formatted" in t: hits += 1
    if "hidden" in t or "space" in t: hits += 1
    return min(1.0, hits/3.0), [f"matched {hits} reasons"]

def check_pivot_steps(ans):
    t = (ans or "").lower()
    hits = sum(1 for w in ("pivot","chart","rows","columns","values","month","sum","line") if w in t)
    return min(1.0, hits/4.0), [f"matched {hits} chart/pivot keywords"]

# grade mapping
def grade(q, answer):
    if q["id"]=="q1": return check_sumifs(answer)
    if q["id"]=="q2": return check_xlookup(answer)
    if q["id"]=="q3": return check_countif(answer)
    if q["id"]=="q4": return check_average(answer)
    if q["id"]=="q5": return check_year(answer)
    if q["id"]=="q6": return check_div0(answer)
    if q["id"]=="q7": return check_sum_zero(answer)
    if q["id"]=="q8":
        t = (answer or "").lower()
        ok = any(w in t for w in ("absolute","relative","$a$1","a1"))
        return (1.0, ["Explains absolute vs relative"]) if ok else (0.0, ["Does not explain absolute vs relative"])
    if q["id"]=="q9":
        t = (answer or "").lower()
        ok = any(w in t for w in ("xlookup","vlookup","index","match","both directions"))
        return (1.0, ["Mentions key differences"]) if ok else (0.0, ["Does not mention differences"])
    if q["id"]=="q10": return check_pivot_steps(answer)
    return (0.0, ["no rule"])

# -------------------------
# Human readable feedback templates
# -------------------------
def feedback_from_notes(qid, notes):
    # basic mapping from notes to friendly tips
    tips = []
    joined = " ".join(notes).lower()
    if "sumifs" in joined:
        tips.append("Good: you used SUMIFS. Tip: consider using Table references for robustness.")
    if "sumif" in joined and "sumifs" not in joined:
        tips.append("Partial: SUMIF works for single conditions; SUMIFS handles multiple conditions.")
    if "xlookup" in joined or "index" in joined:
        tips.append("Good: using XLOOKUP/INDEX+MATCH is robust to column re-ordering.")
    if "countif" in joined or "counta" in joined:
        tips.append("COUNTIF / COUNTA are good for counting text/non-empty cells.")
    if "average" in joined:
        tips.append("AVERAGE ignores blanks; AVERAGEIF can help with conditions.")
    if "year" in joined:
        tips.append("YEAR() extracts year from dates — useful for grouping by year.")
    if "div0" in joined or "divide" in joined:
        tips.append("Check denominators and handle zero or blank cases (IFERROR or conditional checks).")
    if "pivot" in joined or "chart" in joined:
        tips.append("Good: use PivotTables or Group by Month to summarize time-series data.")
    if not tips:
        # generic suggestions based on q type
        if qid.startswith("q1") or qid.startswith("q2") or qid.startswith("q3") or qid.startswith("q4") or qid.startswith("q5"):
            tips.append("If unsure, include the exact formula syntax (e.g., =SUMIFS(...)).")
        else:
            tips.append("Try to mention concrete functions or steps (keywords help the grader).")
    return " ".join(tips[:2])

# -------------------------
# Confidence calculation
# -------------------------
def confidence_label(score):
    if score >= 0.8:
        return "High"
    if score >= 0.4:
        return "Medium"
    return "Low"

# -------------------------
# Hands-on pivot validation
# -------------------------
def pivot_validate_dataframe(df):
    note = ""
    ok = False
    try:
        # normalize columns
        cols = [c.strip().lower() for c in df.columns]
        # attempt to find date and sales columns
        date_col = None
        sales_col = None
        for c in df.columns:
            lc = c.strip().lower()
            if "date" in lc:
                date_col = c
            if "sales" in lc or "amount" in lc or "revenue" in lc:
                sales_col = c
        if date_col is None or sales_col is None:
            note = f"Missing columns: date_col={date_col}, sales_col={sales_col}"
            return False, note, None
        # coerce date
        df2 = df.copy()
        df2[date_col] = pd.to_datetime(df2[date_col], errors="coerce")
        if df2[date_col].isna().all():
            note = "All date values could not be parsed."
            return False, note, None
        # drop rows without sales numeric
        df2[sales_col] = pd.to_numeric(df2[sales_col], errors="coerce")
        # group by month
        df2["month"] = df2[date_col].dt.to_period("M").dt.to_timestamp()
        monthly = df2.groupby("month")[sales_col].sum().reset_index().sort_values("month")
        note = f"Pivot OK — found {len(monthly)} months; sample total for first month: {monthly.iloc[0][sales_col] if len(monthly)>0 else 'N/A'}"
        ok = True
        return ok, note, monthly
    except Exception as e:
        return False, f"Pivot validation error: {e}", None

# -------------------------
# Session state init
# -------------------------
if "started" not in st.session_state:
    st.session_state.started = False
    st.session_state.current_q = 0
    st.session_state.responses = {}
    st.session_state.scores = {}
    st.session_state.notes = {}
    st.session_state.started_at = None
    st.session_state.upload_preview = None
    st.session_state.pivot_monthly = None
    st.session_state.show_intro = False

# -------------------------
# Header / top bar
# -------------------------
col1, col2 = st.columns([4,1])
with col1:
    st.markdown(
        f"""
        <div class="topbar">
          <div class="brand">XL</div>
          <div>
            <div class="title">Excel Mock Interviewer</div>
            <div class="subtitle">Rule-based grader · 10 questions · deterministic & auditable</div>
          </div>
        </div>
        """, unsafe_allow_html=True
    )
with col2:
    if st.session_state.responses:
        total_w = sum(q["weight"] for q in QUESTIONS)
        score_sum = sum((st.session_state.scores.get(q["id"], 0.0) * q["weight"]) for q in QUESTIONS)
        overall = round(100 * (score_sum / total_w), 1) if total_w else 0.0
    else:
        overall = 0.0
    st.markdown(f"<div style='text-align:right'><span class='badge'>Overall: {overall} / 100</span></div>", unsafe_allow_html=True)

st.write("")  # spacer

# -------------------------
# Sidebar: controls & progress
# -------------------------
with st.sidebar:
    st.header("Controls")
    candidate = st.text_input("Candidate name", value=st.session_state.get("candidate_name", "Candidate"))
    if st.button("Start / Reset"):
        st.session_state.started = True
        st.session_state.started_at = datetime.utcnow().isoformat()
        st.session_state.current_q = 0
        st.session_state.responses = {}
        st.session_state.scores = {}
        st.session_state.notes = {}
        st.session_state.upload_preview = None
        st.session_state.pivot_monthly = None
        st.session_state.show_intro = True
        st.session_state.candidate_name = candidate
    st.write("---")
    st.write(f"Question {min(st.session_state.current_q+1, len(QUESTIONS))} / {len(QUESTIONS)}")
    st.progress(st.session_state.current_q / max(1, len(QUESTIONS)))
    st.write("---")
    st.markdown("Tip: paste formula answers like `=SUMIFS(B:B, A:A, \"Sales\")`")
    st.write("")
    if st.session_state.responses:
        if st.button("Download last transcript"):
            payload = {
                "candidate": st.session_state.get("candidate_name"),
                "started_at": st.session_state.get("started_at"),
                "finished_at": datetime.utcnow().isoformat(),
                "responses": st.session_state.responses,
                "scores": st.session_state.scores,
                "notes": st.session_state.notes,
                "overall": overall,
                "upload_preview": st.session_state.get("upload_preview")
            }
            st.download_button("Download JSON", json.dumps(payload, indent=2), file_name="transcript.json")

# -------------------------
# Intro card (display when started and show_intro True)
# -------------------------
if st.session_state.started and st.session_state.show_intro:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("## Welcome — Interview overview")
    st.markdown(
        f"- Candidate: **{st.session_state.get('candidate_name','-')}**  \n"
        f"- This interview has **{len(QUESTIONS)}** questions and is rule-based.  \n"
        "- For hands-on questions you may optionally upload a CSV/XLSX file.  \n"
        "- Use **Submit** to save an answer or **Skip** to move on.  \n"
        "- At the end you'll get a constructive summary and a downloadable transcript."
    )
    st.markdown("**Estimated time:** ~10–20 minutes (depends on hands-on tasks).")
    st.markdown("**What happens next:** The interviewer will ask questions one-by-one and score them deterministically. Low-confidence answers will be flagged for human review.")
    st.markdown("</div>", unsafe_allow_html=True)
    # hide after showing once
    st.session_state.show_intro = False

# -------------------------
# Main interview flow
# -------------------------
if not st.session_state.started:
    st.markdown('<div class="card"><strong>Welcome</strong><div class="small">Click Start / Reset in the left to begin the interview.</div></div>', unsafe_allow_html=True)
else:
    idx = st.session_state.current_q
    if idx < len(QUESTIONS):
        q = QUESTIONS[idx]
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown(f"**Question {idx+1} / {len(QUESTIONS)}**")
        st.markdown(f"### {q['prompt']}")
        st.markdown(f"<div class='muted'>Example: <code>{q['example_answer']}</code></div>", unsafe_allow_html=True)

        # hands-on upload support for q10
        pivot_note = None
        if q["type"] == "hands_on":
            uploaded = st.file_uploader("Upload CSV/XLSX (optional)", type=["csv", "xlsx"])
            if uploaded:
                try:
                    if uploaded.name.lower().endswith(".csv"):
                        df = pd.read_csv(uploaded)
                    else:
                        df = pd.read_excel(uploaded)
                    st.write("Preview (first 8 rows):")
                    st.dataframe(df.head(8))
                    st.session_state.upload_preview = df.head(50).to_dict(orient="records")
                    ok, pivot_note, monthly = pivot_validate_dataframe(df)
                    st.session_state.pivot_monthly = monthly
                    # include pivot validation note in UI immediately
                    if ok:
                        st.success("Hands-on validation: OK — monthly grouping detected.")
                        st.write("Monthly sums (first 8 rows):")
                        st.dataframe(monthly.head(8).assign(month=lambda d: d['month'].dt.strftime('%Y-%m')))
                    else:
                        st.warning(f"Hands-on validation: {pivot_note}")
                except Exception as e:
                    st.error("Could not read file: " + str(e))

        answer = st.text_area("Your answer", key=q["id"], height=160)

        # Actions: Submit / Skip only (follow-up removed)
        c1, c2 = st.columns([1,1])
        with c1:
            if st.button("Submit", key=f"submit_{q['id']}"):
                st.session_state.responses[q['id']] = answer
                score, notes = grade(q, answer)
                # incorporate pivot note into notes if present and on hands_on
                if q["type"]=="hands_on" and pivot_note:
                    notes = notes + [pivot_note]
                st.session_state.scores[q['id']] = score
                st.session_state.notes[q['id']] = notes
                st.session_state.current_q += 1
        with c2:
            if st.button("Skip", key=f"skip_{q['id']}"):
                st.session_state.responses[q['id']] = ""
                st.session_state.scores[q['id']] = 0.0
                st.session_state.notes[q['id']] = ["skipped"]
                st.session_state.current_q += 1

        st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.success("Interview complete — see the constructive summary below.")
        st.markdown('</div>', unsafe_allow_html=True)

# -------------------------
# Review panel & enhanced final report
# -------------------------
if st.session_state.responses:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("Answers & notes (detail)")
    review_queue = []
    strengths = []
    weaknesses = []

    for q in QUESTIONS:
        qid = q["id"]
        if qid in st.session_state.responses:
            ans = st.session_state.responses.get(qid, "")
            score = st.session_state.scores.get(qid, 0.0)
            notes = st.session_state.notes.get(qid, [])
            conf = confidence_label(score)
            # display question block
            st.markdown(f"**Q:** {q['prompt']}")
            st.write("Answer:", ans or "(skipped)")
            # confidence badge
            if conf == "High":
                st.markdown(f"<div class='small'>Confidence: <span class='conf-high'>{conf}</span></div>", unsafe_allow_html=True)
            elif conf == "Medium":
                st.markdown(f"<div class='small'>Confidence: <span class='conf-med'>{conf}</span></div>", unsafe_allow_html=True)
            else:
                st.markdown(f"<div class='small'>Confidence: <span class='conf-low'>{conf}</span></div>", unsafe_allow_html=True)
            st.write("Score (0–1):", round(score, 3))
            st.write("Notes:", notes)
            # human-readable feedback
            friendly = feedback_from_notes(qid, notes)
            st.write("Feedback:", friendly)
            st.write("---")

            # collect strengths / weaknesses / review queue
            if score >= 0.8:
                strengths.append(q["prompt"])
            if score <= 0.4:
                weaknesses.append((q["prompt"], friendly))
            if score < 0.4:
                review_queue.append({"question": q["prompt"], "answer": ans, "notes": notes})

    # enhanced summary: strengths, weaknesses, next steps
    st.subheader("Constructive Summary")
    st.markdown("**Strengths**")
    if strengths:
        for s in strengths:
            st.write("- " + s)
    else:
        st.write("- (none detected)")

    st.markdown("**Areas to improve**")
    if weaknesses:
        for qtxt, tip in weaknesses:
            st.write("- ", qtxt)
            st.write("  - Suggested next step:", tip)
    else:
        st.write("- (none detected)")

    st.markdown("**Flagged for human review**")
    if review_queue:
        st.write(f"{len(review_queue)} answers flagged as Low confidence:")
        for r in review_queue:
            st.write("- Question:", r["question"])
            st.write("  - Answer:", r["answer"] or "(skipped)")
            st.write("  - Notes:", r["notes"])
    else:
        st.write("- None")

    st.markdown("**Next steps / Resources**")
    st.write("- Review Excel tables & structured references (official docs / Microsoft Learn).")
    st.write("- Practice PivotTables and grouping by month (try sample datasets).")
    st.write("- For formulas, practice exact syntax and edge cases (COUNTIF/AVERAGEIF/SUMIFS).")

    # overall numeric score
    total_w = sum(q["weight"] for q in QUESTIONS)
    total = sum((st.session_state.scores.get(q["id"], 0.0) * q["weight"]) for q in QUESTIONS)
    overall = round(100 * (total / total_w), 1) if total_w else 0.0
    st.markdown(f"### Overall score: **{overall} / 100**")

    # include pivot monthly preview if available
    if st.session_state.pivot_monthly is not None:
        st.markdown("**Hands-on monthly totals (preview)**")
        try:
            preview_df = st.session_state.pivot_monthly.copy()
            preview_df = preview_df.assign(month=preview_df['month'].dt.strftime('%Y-%m'))
            st.dataframe(preview_df.head(12))
        except Exception:
            pass

    # transcript payload (concise and detailed both)
    transcript = {
        "candidate": st.session_state.get("candidate_name"),
        "started_at": st.session_state.get("started_at"),
        "finished_at": datetime.utcnow().isoformat(),
        "responses": st.session_state.responses,
        "scores": st.session_state.scores,
        "notes": st.session_state.notes,
        "overall": overall,
        "strengths": strengths,
        "weaknesses": weaknesses,
        "review_queue": review_queue,
        "upload_preview": st.session_state.get("upload_preview")
    }

    st.download_button("Download transcript (JSON)", json.dumps(transcript, indent=2), file_name="transcript.json")
    st.markdown('</div>', unsafe_allow_html=True)

# Footer
st.write("")
st.caption("Deterministic rule-based PoC. You can refine rules, add more validations, or re-enable LLM scoring later.")
