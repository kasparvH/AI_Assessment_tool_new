import io
import json
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
from dotenv import load_dotenv
load_dotenv(Path(__file__).parent / ".env")

from excel_engine import (
    EXCEL_PATH, LOGO_PATH, OUTPUT_DIR, SESSIONS_DIR,
    calculate_scores, export_filled_excel, get_compliance_summary,
    get_dimension_summary, get_overall_summary, load_questions,
)
from session_manager import (
    generate_session_code, load_session, mark_completed,
    save_session, session_exists,
)
from document_analyzer import analyze_document_against_answers, get_file_type
from config_manager import is_test_mode, get_test_questions_per_dim

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="AI Readiness Assessment — Straightable",
    page_icon="🤖",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Brand CSS ─────────────────────────────────────────────────────────────────
CUSTOM_CSS = """
<style>
    /* Sidebar background */
    [data-testid="stSidebar"] { background-color: #1B3B6F !important; }

    /* Sidebar text — only target text nodes, not form inputs */
    [data-testid="stSidebar"] p,
    [data-testid="stSidebar"] span,
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] div.stMarkdown,
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3,
    [data-testid="stSidebar"] small { color: white !important; }

    /* Main panel inputs — explicitly enabled and readable */
    .stTextInput input,
    .stTextInput textarea {
        color: #111111 !important;
        background-color: #ffffff !important;
        opacity: 1 !important;
        -webkit-text-fill-color: #111111 !important;
    }
    .stSelectbox div[data-baseweb="select"] {
        background-color: #ffffff !important;
        opacity: 1 !important;
    }
    .stSelectbox span {
        color: #111111 !important;
        -webkit-text-fill-color: #111111 !important;
    }
    .stTextArea textarea {
        color: #111111 !important;
        background-color: #ffffff !important;
        -webkit-text-fill-color: #111111 !important;
    }

    /* Buttons */
    .stButton > button {
        background-color: #F7941D !important;
        color: white !important;
        border: none !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
    }
    .stButton > button:hover { background-color: #d97c0f !important; }

    /* Progress bar */
    .stProgress > div > div { background-color: #F7941D !important; }

    /* Radio buttons */
    .stRadio > label { font-size: 0.95rem !important; }

    /* Metric cards */
    [data-testid="metric-container"] {
        background-color: white;
        border: 1px solid #e0e0e0;
        border-radius: 8px;
        padding: 12px;
    }

    /* Inconsistency callout */
    .inconsistency-box {
        background-color: #FFF3CD;
        border-left: 4px solid #F7941D;
        border-radius: 4px;
        padding: 12px 16px;
        margin: 8px 0;
    }

    /* Page header banner */
    .main-header {
        background-color: #1B3B6F;
        padding: 12px 24px;
        border-radius: 8px;
        margin-bottom: 24px;
    }

    /* Hide duplicate radio label */
    div[data-testid="stRadio"] > label { display: none; }

    /* Upload box */
    [data-testid="stFileUploader"] {
        border: 2px dashed #F7941D !important;
        border-radius: 8px !important;
        padding: 8px !important;
    }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

DIMENSIONS_ORDER = [
    "Data Readiness",
    "Operations, Culture & Delivery",
    "Risk, Compliance & Responsible AI",
    "Skills & Organizational Capability",
    "Strategic & Leadership Readiness",
    "Technology Infrastructure & Tooling",
    "Use Case Maturity & Adoption",
]

SECTORS = [
    "Zakelijke dienstverlening",
    "Gezondheidszorg",
    "Detailhandel",
    "Industrie (Manufacturing)",
    "Onderwijs",
    "Bouwnijverheid",
    "Horeca",
    "Logistiek (Vervoer & Opslag)",
    "ICT & Informatievoorziening",
    "Openbaar Bestuur",
    "Financiële dienstverlening",
    "Landbouw, Bosbouw & Visserij",
    "Cultuur, Sport & Recreatie",
    "Onroerend goed",
    "Energievoorziening",
    "E-commerce",
    "Specialistische zakelijke diensten",
    "Water & Afvalbeheer",
    "Persoonlijke dienstverlening",
    "Delfstoffenwinning",
]


# ── Helpers ───────────────────────────────────────────────────────────────────

@st.cache_data
def get_questions():
    return load_questions()


def get_active_questions():
    """Returns the question DataFrame, sliced to N per dimension when test mode is on."""
    df = get_questions()
    if is_test_mode():
        n = get_test_questions_per_dim()
        df = df.groupby("Dimension", group_keys=False).head(n)
    return df


def total_questions():
    return len(get_active_questions())


def build_session_state() -> dict:
    ss = st.session_state
    return {
        "session_code": ss.get("session_code", ""),
        "org_name": ss.get("org_name", ""),
        "respondent_name": ss.get("respondent_name", ""),
        "role": ss.get("role", ""),
        "sector": ss.get("sector", ""),
        "started_at": ss.get("started_at", datetime.now().isoformat()),
        "current_dimension": ss.get("current_dimension", DIMENSIONS_ORDER[0]),
        "answers": {str(k): v for k, v in ss.get("answers", {}).items()},
        "uploaded_documents": ss.get("uploaded_documents", []),
        "inconsistencies_flagged": ss.get("inconsistencies_flagged", []),
        "phase": ss.get("phase", "assessment"),
    }


def on_answer_change(question_id: int, value: int):
    st.session_state.answers[question_id] = value
    save_session(st.session_state.session_code, build_session_state())


def render_inconsistency_alert(inconsistency: dict):
    severity_color = {"high": "#dc3545", "medium": "#F7941D", "low": "#ffc107"}
    color = severity_color.get(inconsistency.get("severity", "medium"), "#F7941D")
    st.markdown(f"""
    <div style="border-left: 4px solid {color}; background: #FFF8F0;
                padding: 12px 16px; border-radius: 4px; margin: 8px 0;">
        <strong>⚠️ Document Inconsistency Detected</strong><br>
        <em>From: {inconsistency.get('source_file', 'uploaded document')}</em><br><br>
        {inconsistency['explanation']}<br><br>
        <strong>Document says:</strong> <em>"{inconsistency['document_evidence']}"</em><br><br>
        <small>Please review your answer and correct it if needed.</small>
    </div>
    """, unsafe_allow_html=True)


def dim_progress(df, answers, dim):
    q_ids = df[df["Dimension"] == dim].index.tolist()
    answered = sum(1 for qid in q_ids if qid in answers)
    return answered, len(q_ids)


def sidebar_logo():
    if LOGO_PATH.exists():
        st.sidebar.image(str(LOGO_PATH), use_container_width=True)
    st.sidebar.markdown("---")


def init_session():
    for key, default in [
        ("phase", "welcome"),
        ("answers", {}),
        ("session_code", ""),
        ("org_name", ""),
        ("respondent_name", ""),
        ("role", ""),
        ("sector", ""),
        ("started_at", datetime.now().isoformat()),
        ("current_dimension", DIMENSIONS_ORDER[0]),
        ("uploaded_documents", []),
        ("inconsistencies_flagged", []),
        ("report_text", ""),
        ("report_generated", False),
        ("paused", False),
    ]:
        if key not in st.session_state:
            st.session_state[key] = default


# ── Screen: Welcome ───────────────────────────────────────────────────────────

def screen_welcome():
    sidebar_logo()
    st.sidebar.markdown("### AI Readiness Assessment")
    st.sidebar.markdown("*by Straightable Innovation & Strategy*")

    st.markdown("""
    <div class="main-header">
        <h1 style="color:white;margin:0;">AI Readiness Assessment</h1>
        <p style="color:#F7941D;margin:0;">Powered by Straightable Innovation & Strategy</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    Discover the AI maturity level of your organization across **7 key dimensions**
    in approximately **45–60 minutes**.
    """)

    if is_test_mode():
        n = get_test_questions_per_dim()
        st.warning(
            f"🧪 **Test mode actief** — deze sessie bevat {n} vragen per dimensie "
            f"({7 * n} vragen totaal in plaats van 265)."
        )

    col1, col2 = st.columns([1, 1], gap="large")

    with col1:
        st.markdown("### ▶ Start New Assessment")
        org = st.text_input("Organization name", key="new_org")
        resp = st.text_input("Respondent name", key="new_resp")
        role = st.text_input("Role / function", key="new_role")
        sector = st.selectbox("Sector", SECTORS, key="new_sector")

        if st.button("Start Assessment →", use_container_width=True):
            if not org or not resp:
                st.warning("Please fill in organization name and respondent name.")
            else:
                code = generate_session_code()
                st.session_state.session_code = code
                st.session_state.org_name = org
                st.session_state.respondent_name = resp
                st.session_state.role = role
                st.session_state.sector = sector
                st.session_state.started_at = datetime.now().isoformat()
                st.session_state.answers = {}
                st.session_state.current_dimension = DIMENSIONS_ORDER[0]
                st.session_state.inconsistencies_flagged = []
                st.session_state.uploaded_documents = []
                st.session_state.phase = "assessment"
                save_session(code, build_session_state())
                st.info(f"✅ **Your session code: {code}** — Save this to resume later!")
                st.rerun()

    with col2:
        st.markdown("### ↩ Resume Assessment")
        st.markdown("Enter your session code to continue where you left off.")
        resume_code = st.text_input("Session code (e.g. STR-4K9)", key="resume_code")

        if st.button("Resume →", use_container_width=True):
            code_clean = resume_code.strip().upper()
            if session_exists(code_clean):
                data = load_session(code_clean)
                st.session_state.session_code = code_clean
                st.session_state.org_name = data.get("org_name", "")
                st.session_state.respondent_name = data.get("respondent_name", "")
                st.session_state.role = data.get("role", "")
                st.session_state.sector = data.get("sector", "")
                st.session_state.started_at = data.get("started_at", "")
                st.session_state.current_dimension = data.get("current_dimension", DIMENSIONS_ORDER[0])
                st.session_state.answers = {int(k): v for k, v in data.get("answers", {}).items()}
                st.session_state.uploaded_documents = data.get("uploaded_documents", [])
                st.session_state.inconsistencies_flagged = data.get("inconsistencies_flagged", [])
                st.session_state.phase = "assessment"
                n = len(st.session_state.answers)
                st.success(f"✅ **Session restored** — You have answered {n} of {total_questions()} questions. Continue from: **{st.session_state.current_dimension}**")
                st.rerun()
            else:
                st.error("⚠️ Session code not found. Please check and try again.")


# ── Screen: Assessment ────────────────────────────────────────────────────────

def screen_assessment():
    df = get_active_questions()
    answers = st.session_state.answers
    inconsistencies = st.session_state.inconsistencies_flagged
    n_total = total_questions()
    test_mode = is_test_mode()

    # ── Sidebar ──
    sidebar_logo()

    if test_mode:
        st.sidebar.markdown(
            "<div style='background:#F7941D;color:white;padding:5px 8px;"
            "border-radius:5px;font-size:0.8rem;font-weight:600;text-align:center;'>"
            "🧪 TEST MODE — 4 vragen/dim</div>",
            unsafe_allow_html=True,
        )

    st.sidebar.markdown(f"**Session:** 📋 `{st.session_state.session_code}`")
    if st.sidebar.button("📋 Copy code"):
        st.sidebar.code(st.session_state.session_code)

    st.sidebar.markdown("---")
    n_answered = len([k for k in answers if k in df.index])
    st.sidebar.markdown(f"**Progress: {n_answered} / {n_total} answered**")
    st.sidebar.progress(n_answered / n_total if n_total > 0 else 0)

    st.sidebar.markdown("---")
    st.sidebar.markdown("**Dimensions**")
    for dim in DIMENSIONS_ORDER:
        answered, total = dim_progress(df, answers, dim)
        if answered == total:
            icon = "✅"
        elif answered > 0:
            icon = "🔄"
        else:
            icon = "⬜"
        label = f"{icon} {dim} ({answered}/{total})"
        if st.sidebar.button(label, key=f"nav_{dim}", use_container_width=True):
            st.session_state.current_dimension = dim
            st.rerun()

    st.sidebar.markdown("---")
    if inconsistencies:
        st.sidebar.warning(f"⚠️ {len(inconsistencies)} inconsistencies found")

    st.sidebar.markdown("---")
    if st.sidebar.button("💾 Save & Pause", use_container_width=True):
        save_session(st.session_state.session_code, build_session_state())
        st.session_state.paused = True
        st.rerun()

    remaining = n_total - n_answered
    if remaining == 0:
        if st.sidebar.button("💾 Assessment opslaan →", use_container_width=True):
            st.session_state.phase = "results"
            st.rerun()
    else:
        st.sidebar.markdown(f"*Complete all questions first. **{remaining}** remaining.*")

    # ── Pause confirmation ──
    if st.session_state.get("paused"):
        st.success("✅ **Assessment Paused** — Your progress has been saved.")
        st.info(f"Your session code is: **{st.session_state.session_code}**\n\nWrite this down or save it — you'll need it to resume.")
        if st.button("Continue answering"):
            st.session_state.paused = False
            st.rerun()
        return

    # ── Main panel ──
    dim = st.session_state.current_dimension
    dim_df = df[df["Dimension"] == dim]
    answered_in_dim, total_in_dim = dim_progress(df, answers, dim)

    st.markdown(f"## {dim}")
    st.markdown(f"*{answered_in_dim} of {total_in_dim} questions answered*")

    # ── Document upload — prominent panel ──
    with st.expander("📎 Upload evidence documents (optional — validates your answers automatically)", expanded=False):
        st.markdown(
            "Upload company documents such as AI policies, strategy docs, or data governance frameworks. "
            "They will be compared against your answers to flag any inconsistencies."
        )
        uploaded_files = st.file_uploader(
            "Accepted formats: PDF, Word (.docx), images (PNG/JPG), plain text (.txt)",
            accept_multiple_files=True,
            type=["pdf", "docx", "doc", "png", "jpg", "jpeg", "txt"],
            key="doc_uploader",
        )

        already_uploaded = st.session_state.get("uploaded_documents", [])
        if already_uploaded:
            st.markdown(f"**Previously analyzed:** {', '.join(d['filename'] for d in already_uploaded)}")

        if st.button("🔍 Analyze Documents", use_container_width=True) and uploaded_files:
            answered_ids = list(answers.keys())
            new_inconsistencies = []
            analysis_errors = []

            with st.spinner("Analyzing documents..."):
                for uf in uploaded_files:
                    file_bytes = uf.read()
                    ftype = get_file_type(uf.name)
                    try:
                        found = analyze_document_against_answers(
                            uf.name, file_bytes, ftype, answers, df, answered_ids
                        )
                        for item in found:
                            item["source_file"] = uf.name
                        new_inconsistencies.extend(found)
                        st.session_state.uploaded_documents.append({
                            "filename": uf.name,
                            "uploaded_at": datetime.now().isoformat(),
                            "analyzed": True,
                        })
                    except Exception as e:
                        analysis_errors.append(f"**{uf.name}:** {str(e)}")

            st.session_state.inconsistencies_flagged = (
                st.session_state.inconsistencies_flagged + new_inconsistencies
            )
            # Store errors persistently so they don't vanish on rerun
            if analysis_errors:
                st.session_state["_doc_errors"] = analysis_errors
            else:
                st.session_state.pop("_doc_errors", None)

            save_session(st.session_state.session_code, build_session_state())
            st.rerun()

        # Show persistent errors from previous run
        if st.session_state.get("_doc_errors"):
            st.error("⚠️ The following errors occurred during document analysis:")
            for err in st.session_state["_doc_errors"]:
                st.markdown(f"- {err}")
            st.markdown("**Possible causes:** unsupported file encoding, empty document, or API connection issue. Try again or use a different file.")
            if st.button("Dismiss errors", key="dismiss_doc_errors"):
                st.session_state.pop("_doc_errors", None)
                st.rerun()

        # Show result of last analysis
        if not st.session_state.get("_doc_errors"):
            flagged = st.session_state.inconsistencies_flagged
            if flagged:
                st.warning(f"⚠️ {len(flagged)} inconsistencies found — see alerts inline below.")
            elif already_uploaded:
                st.success("✅ No inconsistencies found in the uploaded documents.")

    st.markdown("---")

    # Build inconsistency lookup by question_id
    incon_by_qid = {}
    for inc in inconsistencies:
        qid = inc.get("question_id")
        if qid:
            incon_by_qid.setdefault(qid, []).append(inc)

    for qid, row in dim_df.iterrows():
        badges = []
        for fw, label in [("EU_AI_ACT", "🇪🇺 EU AI Act"), ("NIST_AI_RMF", "📋 NIST"),
                           ("ISO_42001", "🏅 ISO 42001"), ("AI_TRISM", "🛡️ AI TRISM")]:
            if str(row.get(fw, "")).upper() == "YES":
                badges.append(label)
        badge_str = "  &nbsp;".join(badges) if badges else ""

        st.markdown(f"""
        <div style="margin-bottom:4px;">
            <span style="color:#1B3B6F;font-weight:600;">Q{qid}</span>
            {"&nbsp;&nbsp;" + badge_str if badge_str else ""}
        </div>
        """, unsafe_allow_html=True)

        st.markdown(f"**{row['Question_text']}**")

        current_answer = answers.get(qid)
        options = [1, 2, 3]
        idx = (current_answer - 1) if current_answer in options else None

        answer = st.radio(
            label=f"q_{qid}",
            options=options,
            format_func=lambda x, r=row: r.get(f"Answer_options_{x}", f"Option {x}"),
            index=idx,
            key=f"q_{qid}",
            label_visibility="collapsed",
        )
        if answer and answer != current_answer:
            on_answer_change(qid, answer)
            st.rerun()

        # Show inconsistencies inline
        if qid in incon_by_qid:
            for inc in incon_by_qid[qid]:
                render_inconsistency_alert(inc)

        st.markdown("---")

    # ── Next dimension button ──
    current_idx = DIMENSIONS_ORDER.index(dim) if dim in DIMENSIONS_ORDER else -1
    if current_idx < len(DIMENSIONS_ORDER) - 1:
        next_dim = DIMENSIONS_ORDER[current_idx + 1]
        st.markdown(" ")
        col_l, col_r = st.columns([3, 1])
        with col_r:
            if st.button(f"Volgende: {next_dim} →", use_container_width=True):
                st.session_state.current_dimension = next_dim
                save_session(st.session_state.session_code, build_session_state())
                st.rerun()
    else:
        # Last dimension — show go to results if all answered
        remaining = n_total - len([k for k in answers if k in df.index])
        st.markdown(" ")
        if remaining == 0:
            col_l, col_r = st.columns([3, 1])
            with col_r:
                if st.button("💾 Sla assessment op →", use_container_width=True):
                    st.session_state.phase = "results"
                    st.rerun()
        else:
            st.info(f"Nog **{remaining}** vragen te beantwoorden voor je de resultaten kunt bekijken.")

def screen_results():
    df = get_active_questions()
    answers = st.session_state.answers
    scored = calculate_scores(df, answers)
    dim_summary = get_dimension_summary(scored)
    overall = get_overall_summary(dim_summary)
    compliance = get_compliance_summary(scored)
    inconsistencies = st.session_state.inconsistencies_flagged

    sidebar_logo()
    st.sidebar.markdown(f"**Session:** `{st.session_state.session_code}`")
    st.sidebar.markdown("---")
    if st.sidebar.button("← Back to Assessment"):
        st.session_state.phase = "assessment"
        st.rerun()

    st.markdown("# Your AI Readiness Results")
    st.markdown(f"**Organization:** {st.session_state.org_name} &nbsp;&nbsp; **Date:** {datetime.now().strftime('%d %B %Y')}")

    # Overall score box
    tier = overall["maturity_tier"]
    score = overall["overall_score_0_5"]
    pct = overall["overall_pct"]
    st.markdown(f"""
    <div style="background:#1B3B6F;padding:20px;border-radius:8px;margin-bottom:20px;text-align:center;">
        <h2 style="color:white;margin:0;">Overall Maturity Score</h2>
        <h1 style="color:#F7941D;margin:4px 0;">{score:.1f} / 5.0</h1>
        <h3 style="color:white;margin:0;">Tier: {tier} ({pct:.0f}%)</h3>
    </div>
    """, unsafe_allow_html=True)

    # Dimension scores
    st.markdown("### Dimension Scores")
    for _, row in dim_summary.sort_values("score_pct", ascending=False).iterrows():
        col1, col2, col3 = st.columns([4, 1, 1])
        with col1:
            st.markdown(f"**{row['dimension']}**")
            st.progress(row["score_pct"] / 100)
        with col2:
            st.markdown(f"**{row['score_0_5']:.1f}/5**")
        with col3:
            st.markdown(f"**{row['score_pct']:.0f}%**")

    if inconsistencies:
        st.warning(f"⚠️ **{len(inconsistencies)} document inconsistencies** were flagged during this assessment.")
        with st.expander("Review inconsistencies & pas antwoorden aan", expanded=True):
            st.markdown(
                "Het geüploade document wijkt op onderstaande punten af van jouw antwoorden. "
                "Je kunt per punt jouw antwoord corrigeren naar de werkelijke situatie."
            )
            st.markdown("---")

            # Load full question dataframe (all 265) for answer option lookup
            df_full = get_questions()

            changed_any = False
            for idx, inc in enumerate(inconsistencies):
                qid = inc.get("question_id")
                sev = inc.get("severity", "medium")
                sev_color = {"high": "#dc3545", "medium": "#F7941D", "low": "#ffc107"}.get(sev, "#F7941D")
                sev_nl = {"high": "Hoog", "medium": "Middel", "low": "Laag"}.get(sev, sev)

                st.markdown(f"""
                <div style="border-left:4px solid {sev_color};padding:4px 12px;margin-bottom:4px;">
                    <strong>Q{qid} — Ernst: {sev_nl}</strong><br>
                    {inc.get('question_text', '')}
                </div>
                """, unsafe_allow_html=True)

                col_a, col_b = st.columns(2)
                with col_a:
                    st.markdown(f"**Jouw antwoord:**")
                    st.markdown(f"*{inc.get('selected_answer', '')}*")
                with col_b:
                    st.markdown(f"**Document stelt:**")
                    st.markdown(f"*{inc.get('document_evidence', '')}*")

                st.markdown(f"**Toelichting:** {inc.get('explanation', '')}")

                # Show answer options for this question
                row_match = df_full[df_full.index == qid] if qid in df_full.index else df_full[df_full['Question_id'] == qid] if 'Question_id' in df_full.columns else None

                current_answer = answers.get(qid)
                already_adjusted = st.session_state.get(f"_adj_confirmed_{qid}", False)

                if already_adjusted:
                    st.success(f"✅ Antwoord voor Q{qid} is aangepast naar optie {answers.get(qid)}.")
                else:
                    if row_match is not None and not row_match.empty:
                        row = row_match.iloc[0]
                        opt1 = row.get('Answer_options_1', 'Optie 1')
                        opt2 = row.get('Answer_options_2', 'Optie 2')
                        opt3 = row.get('Answer_options_3', 'Optie 3')

                        st.markdown("**Selecteer het juiste antwoord op basis van het document:**")
                        new_answer = st.radio(
                            label=f"Aanpassing Q{qid}",
                            options=[1, 2, 3],
                            format_func=lambda x, o1=opt1, o2=opt2, o3=opt3: {1: o1, 2: o2, 3: o3}[x],
                            index=(current_answer - 1) if current_answer in [1, 2, 3] else 0,
                            key=f"adj_radio_{qid}_{idx}",
                            label_visibility="collapsed",
                        )

                        if st.button(f"✅ Bevestig aanpassing Q{qid}", key=f"adj_btn_{qid}_{idx}", use_container_width=False):
                            on_answer_change(qid, new_answer)
                            st.session_state[f"_adj_confirmed_{qid}"] = True
                            changed_any = True
                            st.rerun()

                st.markdown("---")

            if changed_any:
                st.success("Antwoorden bijgewerkt. De scores hieronder zijn herberekend.")
                st.rerun()

    st.info("✅ Je assessment is volledig ingevuld en opgeslagen. De Straightable consultant genereert het rapport voor je en stuurt dit toe.")


# ── Screen: Report (admin only) ───────────────────────────────────────────────

def screen_report():
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import numpy as np

    from report_generator import build_report_input, generate_report_safe
    from pdf_export import create_radar_chart, export_pdf

    df = get_active_questions()
    answers = st.session_state.answers
    scored = calculate_scores(df, answers)

    # Fix: use explicit None check — DataFrame cannot be used as bool in 'or'
    _ds = st.session_state.get("_dim_summary")
    dim_summary = _ds if _ds is not None else get_dimension_summary(scored)
    _ov = st.session_state.get("_overall")
    overall = _ov if _ov is not None else get_overall_summary(dim_summary)
    _co = st.session_state.get("_compliance")
    compliance = _co if _co is not None else get_compliance_summary(scored)
    inconsistencies = st.session_state.inconsistencies_flagged

    sidebar_logo()
    st.sidebar.markdown(f"**Session:** `{st.session_state.session_code}`")
    st.sidebar.markdown(f"*Session {st.session_state.session_code} completed.*")

    st.markdown(f"# AI Readiness Report — {st.session_state.org_name}")
    st.markdown(f"*Generated: {datetime.now().strftime('%d %B %Y')}*")

    # Generate report if not already done
    if not st.session_state.report_generated:
        report_input = build_report_input(
            st.session_state.org_name,
            st.session_state.respondent_name,
            st.session_state.sector,
            overall, dim_summary, compliance, inconsistencies,
        )
        with st.spinner("Claude is analyzing your results..."):
            report_text, success = generate_report_safe(report_input)
        st.session_state.report_text = report_text
        st.session_state.report_generated = True
        mark_completed(st.session_state.session_code)

    report_text = st.session_state.report_text

    # Radar chart
    radar_buf = io.BytesIO()
    dimensions = dim_summary["dimension"].tolist()
    scores = dim_summary["score_0_5"].tolist()
    N = len(dimensions)
    angles = np.linspace(0, 2 * np.pi, N, endpoint=False).tolist()
    scores_plot = scores + [scores[0]]
    angles_plot = angles + angles[:1]
    fig, ax = plt.subplots(figsize=(7, 7), subplot_kw=dict(polar=True))
    ax.plot(angles_plot, scores_plot, "o-", linewidth=2.5, color="#F7941D")
    ax.fill(angles_plot, scores_plot, alpha=0.2, color="#1B3B6F")
    ax.set_ylim(0, 5)
    ax.set_yticks([1, 2, 3, 4, 5])
    ax.set_xticks(angles)
    ax.set_xticklabels(dimensions, fontsize=9, color="#1B3B6F", fontweight="bold")
    ax.set_facecolor("#F8F9FA")
    ax.grid(color="#cccccc", linestyle="--", linewidth=0.7)
    ax.set_title("AI Maturity by Dimension", fontsize=13, color="#1B3B6F", fontweight="bold", pad=20)
    plt.tight_layout()
    plt.savefig(radar_buf, format="png", dpi=120, bbox_inches="tight", facecolor="white")
    plt.close()
    radar_buf.seek(0)
    st.image(radar_buf, use_container_width=False, width=550)

    # Report text
    st.markdown(report_text)

    # Compliance table
    st.markdown("### Framework Coverage")
    comp_data = []
    for fw, data in compliance.items():
        comp_data.append({
            "Framework": fw.replace("_", " "),
            "Questions covered": f"{data['covered']} / {data['total']}",
            "Avg. Score": f"{data['avg_score']:.1f} / 5.0",
        })
    st.table(pd.DataFrame(comp_data))

    # Downloads
    col1, col2 = st.columns(2)
    with col1:
        if st.button("📄 Download PDF Report", use_container_width=True):
            pdf_path = str(OUTPUT_DIR / f"report_{st.session_state.session_code}.pdf")
            with st.spinner("Generating PDF..."):
                export_pdf(
                    report_text=report_text,
                    dim_summary=dim_summary,
                    overall=overall,
                    compliance=compliance,
                    org_name=st.session_state.org_name,
                    respondent_name=st.session_state.respondent_name,
                    logo_path=str(LOGO_PATH),
                    output_path=pdf_path,
                )
            with open(pdf_path, "rb") as f:
                st.download_button(
                    "⬇️ Download PDF",
                    f,
                    file_name=f"AI_Readiness_Report_{st.session_state.org_name}.pdf",
                    mime="application/pdf",
                )
    with col2:
        if st.button("📥 Download Filled Excel", use_container_width=True):
            out_path = str(OUTPUT_DIR / f"assessment_{st.session_state.session_code}.xlsx")
            export_filled_excel(df, answers, out_path)
            with open(out_path, "rb") as f:
                st.download_button(
                    "⬇️ Download Excel",
                    f,
                    file_name=f"AI_Assessment_{st.session_state.org_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

    st.caption(f"Session {st.session_state.session_code} completed. Session data will be retained for 30 days.")


# ── Main router ───────────────────────────────────────────────────────────────

def main():
    init_session()
    phase = st.session_state.phase

    if phase == "welcome":
        screen_welcome()
    elif phase == "assessment":
        screen_assessment()
    elif phase == "results":
        screen_results()
    elif phase == "report":
        screen_report()
    else:
        screen_welcome()


if __name__ == "__main__":
    main()
