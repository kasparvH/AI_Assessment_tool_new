import io
import json
import os
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
    [data-testid="stSidebar"] { background-color: #1B3B6F !important; }
    [data-testid="stSidebar"] p,
    [data-testid="stSidebar"] span,
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] div.stMarkdown,
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3,
    [data-testid="stSidebar"] small { color: white !important; }
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
    .stButton > button {
        background-color: #F7941D !important;
        color: white !important;
        border: none !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
    }
    .stButton > button:hover { background-color: #d97c0f !important; }
    .stProgress > div > div { background-color: #F7941D !important; }
    .stRadio > label { font-size: 0.95rem !important; }
    [data-testid="metric-container"] {
        background-color: white;
        border: 1px solid #e0e0e0;
        border-radius: 8px;
        padding: 12px;
    }
    .inconsistency-box {
        background-color: #FFF3CD;
        border-left: 4px solid #F7941D;
        border-radius: 4px;
        padding: 12px 16px;
        margin: 8px 0;
    }
    .main-header {
        background-color: #1B3B6F;
        padding: 12px 24px;
        border-radius: 8px;
        margin-bottom: 24px;
    }
    div[data-testid="stRadio"] > label { display: none; }
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
        ("admin_authenticated", False),
    ]:
        if key not in st.session_state:
            st.session_state[key] = default


# ── Admin: password check ─────────────────────────────────────────────────────

def check_admin_password() -> bool:
    """
    Shows a password screen. Returns True if already authenticated.
    Uses ADMIN_PASSWORD from .env or Streamlit Secrets.
    """
    if st.session_state.get("admin_authenticated"):
        return True

    admin_pw = os.getenv("ADMIN_PASSWORD", "")
    if not admin_pw:
        st.error("⚠️ ADMIN_PASSWORD is not set. Add it to your .env file or Streamlit Secrets.")
        st.stop()

    sidebar_logo()
    st.sidebar.markdown("### 🔒 Admin Portal")

    st.markdown("""
    <div class="main-header">
        <h2 style="color:white;margin:0;">🔒 Admin Portal — Straightable</h2>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("Enter the admin password to access the portal.")

    with st.form("admin_login_form"):
        entered = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Login →", use_container_width=True)
        if submitted:
            if entered == admin_pw:
                st.session_state.admin_authenticated = True
                st.rerun()
            else:
                st.error("❌ Incorrect password. Please try again.")

    return False


# ── Admin: portal ─────────────────────────────────────────────────────────────

def render_admin_portal():
    """Full admin portal rendered inside app.py when ?admin=true."""
    sidebar_logo()
    st.sidebar.markdown("### 🔒 Admin Portal")
    st.sidebar.markdown("---")

    if st.sidebar.button("🚪 Logout", use_container_width=True):
        st.session_state.admin_authenticated = False
        st.rerun()

    st.markdown("""
    <div class="main-header">
        <h2 style="color:white;margin:0;">Admin Portal — Straightable</h2>
        <p style="color:#F7941D;margin:0;">AI Readiness Assessment</p>
    </div>
    """, unsafe_allow_html=True)

    tab_sessions, tab_reports, tab_new = st.tabs([
        "📋 All Sessions",
        "📄 Reports & Downloads",
        "➕ Create Session for Client",
    ])

    # ── Tab 1: All sessions ──
    with tab_sessions:
        st.markdown("### All Active Sessions")

        session_files = list(SESSIONS_DIR.glob("*.json")) if SESSIONS_DIR.exists() else []

        if not session_files:
            st.info("No sessions found yet.")
        else:
            rows = []
            for sf in sorted(session_files, key=lambda x: x.stat().st_mtime, reverse=True):
                try:
                    with open(sf, "r", encoding="utf-8") as f:
                        data = json.load(f)
                    n_answered = len(data.get("answers", {}))
                    rows.append({
                        "Code": data.get("session_code", sf.stem),
                        "Organization": data.get("org_name", "—"),
                        "Respondent": data.get("respondent_name", "—"),
                        "Sector": data.get("sector", "—"),
                        "Answered": f"{n_answered} / 265",
                        "Phase": data.get("phase", "—"),
                        "Started": data.get("started_at", "")[:10],
                        "Last saved": datetime.fromtimestamp(sf.stat().st_mtime).strftime("%d-%m-%Y %H:%M"),
                    })
                except Exception:
                    continue

            st.dataframe(pd.DataFrame(rows), use_container_width=True)

            st.markdown("---")
            st.markdown("### Delete a session")
            del_code = st.text_input("Session code to delete", key="admin_del_code")
            if st.button("🗑️ Delete session", key="admin_del_btn"):
                target = SESSIONS_DIR / f"{del_code.strip().upper()}.json"
                if target.exists():
                    target.unlink()
                    st.success(f"Session {del_code.upper()} deleted.")
                    st.rerun()
                else:
                    st.error("Session code not found.")

    # ── Tab 2: Reports & Downloads ──
    with tab_reports:
        st.markdown("### Generate or Download Reports")

        session_files = list(SESSIONS_DIR.glob("*.json")) if SESSIONS_DIR.exists() else []

        if not session_files:
            st.info("No sessions found yet.")
        else:
            session_options = {}
            for sf in sorted(session_files, key=lambda x: x.stat().st_mtime, reverse=True):
                try:
                    with open(sf, "r", encoding="utf-8") as f:
                        data = json.load(f)
                    code = data.get("session_code", sf.stem)
                    org = data.get("org_name", "Unknown")
                    n = len(data.get("answers", {}))
                    session_options[f"{code} — {org} ({n}/265 answered)"] = sf
                except Exception:
                    continue

            selected_label = st.selectbox("Select a session", list(session_options.keys()))
            selected_sf = session_options[selected_label]

            with open(selected_sf, "r", encoding="utf-8") as f:
                session_data = json.load(f)

            code = session_data.get("session_code", "")
            org = session_data.get("org_name", "")
            respondent = session_data.get("respondent_name", "")
            sector = session_data.get("sector", "")
            answers_raw = session_data.get("answers", {})
            answers = {int(k): v for k, v in answers_raw.items()}
            inconsistencies = session_data.get("inconsistencies_flagged", [])

            n_answered = len(answers)
            st.markdown(f"**Organization:** {org} &nbsp;|&nbsp; **Respondent:** {respondent} &nbsp;|&nbsp; **Answered:** {n_answered}/265")

            if n_answered < 265:
                st.warning(f"⚠️ This session is not fully completed ({n_answered}/265 questions answered). You can still generate a partial report.")

            col1, col2 = st.columns(2)

            with col1:
                if st.button("📝 Generate PDF Report", use_container_width=True, key="admin_gen_pdf"):
                    import matplotlib
                    matplotlib.use("Agg")
                    from report_generator import build_report_input, generate_report_safe
                    from pdf_export import export_pdf

                    df = get_questions()
                    scored = calculate_scores(df, answers)
                    dim_summary = get_dimension_summary(scored)
                    overall = get_overall_summary(dim_summary)
                    compliance = get_compliance_summary(scored)

                    report_input = build_report_input(
                        org, respondent, sector,
                        overall, dim_summary, compliance, inconsistencies,
                    )

                    with st.spinner("Claude is generating the report..."):
                        report_text, success = generate_report_safe(report_input)

                    if not success:
                        st.warning("⚠️ Report generated using fallback (Claude API unavailable).")

                    pdf_path = str(OUTPUT_DIR / f"report_{code}.pdf")
                    with st.spinner("Exporting PDF..."):
                        export_pdf(
                            report_text=report_text,
                            dim_summary=dim_summary,
                            overall=overall,
                            compliance=compliance,
                            org_name=org,
                            respondent_name=respondent,
                            logo_path=str(LOGO_PATH),
                            output_path=pdf_path,
                        )

                    with open(pdf_path, "rb") as f:
                        st.download_button(
                            "⬇️ Download PDF Report",
                            f,
                            file_name=f"AI_Readiness_Report_{org}.pdf",
                            mime="application/pdf",
                            key="admin_dl_pdf",
                        )
                    st.success("✅ Report generated successfully.")

            with col2:
                if st.button("📊 Generate Excel Export", use_container_width=True, key="admin_gen_excel"):
                    df = get_questions()
                    out_path = str(OUTPUT_DIR / f"assessment_{code}.xlsx")
                    export_filled_excel(df, answers, out_path)
                    with open(out_path, "rb") as f:
                        st.download_button(
                            "⬇️ Download Excel",
                            f,
                            file_name=f"AI_Assessment_{org}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="admin_dl_excel",
                        )
                    st.success("✅ Excel export ready.")

            # Show existing PDF if already generated
            existing_pdf = OUTPUT_DIR / f"report_{code}.pdf"
            if existing_pdf.exists():
                st.markdown("---")
                st.markdown(f"*Previously generated report found for session {code}:*")
                with open(existing_pdf, "rb") as f:
                    st.download_button(
                        "⬇️ Re-download existing PDF",
                        f,
                        file_name=f"AI_Readiness_Report_{org}.pdf",
                        mime="application/pdf",
                        key="admin_redl_pdf",
                    )

    # ── Tab 3: Create session for client ──
    with tab_new:
        st.markdown("### Pre-create a session for a client")
        st.markdown(
            "Create a session code that you can send to a client so they can start "
            "the assessment directly without having to fill in a new code themselves."
        )

        with st.form("admin_create_session"):
            new_org = st.text_input("Organization name")
            new_resp = st.text_input("Client contact name")
            new_role = st.text_input("Role / function")
            new_sector = st.selectbox("Sector", SECTORS)
            create_btn = st.form_submit_button("➕ Create session →", use_container_width=True)

        if create_btn:
            if not new_org or not new_resp:
                st.warning("Please fill in organization name and contact name.")
            else:
                new_code = generate_session_code()
                session_payload = {
                    "session_code": new_code,
                    "org_name": new_org,
                    "respondent_name": new_resp,
                    "role": new_role,
                    "sector": new_sector,
                    "started_at": datetime.now().isoformat(),
                    "current_dimension": DIMENSIONS_ORDER[0],
                    "answers": {},
                    "uploaded_documents": [],
                    "inconsistencies_flagged": [],
                    "phase": "assessment",
                }
                save_session(new_code, session_payload)

                st.success(f"✅ Session created successfully!")
                st.markdown(f"""
                <div style="background:#1B3B6F;padding:20px;border-radius:8px;text-align:center;">
                    <p style="color:white;margin:0;">Session code for {new_org}:</p>
                    <h1 style="color:#F7941D;margin:8px 0;letter-spacing:4px;">{new_code}</h1>
                    <p style="color:white;margin:0;font-size:0.9rem;">
                        Send this code to the client. They enter it on the welcome screen to start.
                    </p>
                </div>
                """, unsafe_allow_html=True)


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

    if st.session_state.get("paused"):
        st.success("✅ **Assessment Paused** — Your progress has been saved.")
        st.info(f"Your session code is: **{st.session_state.session_code}**\n\nWrite this down or save it — you'll need it to resume.")
        if st.button("Continue answering"):
            st.session_state.paused = False
            st.rerun()
        return

    dim = st.session_state.current_dimension
    dim_df = df[df["Dimension"] == dim]
    answered_in_dim, total_in_dim = dim_progress(df, answers, dim)

    st.markdown(f"## {dim}")
    st.markdown(f"*{answered_in_dim} of {total_in_dim} questions answered*")

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
            if analysis_errors:
                st.session_state["_doc_errors"] = analysis_errors
            else:
                st.session_state.pop("_doc_errors", None)

            save_session(st.session_state.session_code, build_session_state())
            st.rerun()

        if st.session_state.get("_doc_errors"):
            st.error("⚠️ The following errors occurred during document analysis:")
            for err in st.session_state["_doc_errors"]:
                st.markdown(f"- {err}")
            st.markdown("**Possible causes:** unsupported file encoding, empty document, or API connection issue.")
            if st.button("Dismiss errors", key="dismiss_doc_errors"):
                st.session_state.pop("_doc_errors", None)
                st.rerun()

        if not st.session_state.get("_doc_errors"):
            flagged = st.session_state.inconsistencies_flagged
            if flagged:
                st.warning(f"⚠️ {len(flagged)} inconsistencies found — see alerts inline below.")
            elif already_uploaded:
                st.success("✅ No inconsistencies found in the uploaded documents.")

    st.markdown("---")

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

        if qid in incon_by_qid:
            for inc in incon_by_qid[qid]:
                render_inconsistency_alert(inc)

        st.markdown("---")

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


# ── Screen: Results ───────────────────────────────────────────────────────────

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
                    st.markdown("**Jouw antwoord:**")
                    st.markdown(f"*{inc.get('selected_answer', '')}*")
                with col_b:
                    st.markdown("**Document stelt:**")
                    st.markdown(f"*{inc.get('document_evidence', '')}*")

                st.markdown(f"**Toelichting:** {inc.get('explanation', '')}")

                row_match = df_full[df_full.index == qid] if qid in df_full.index else df_full[df_full['Question_id'] == qid] if 'Question_id' in df_full.columns else None
                current_answer = answers.get(qid)
                already_adjusted = st.session_state.get(f"_adj_confirmed_{qid}", False)

                if already_adjusted:
                    st.success(f"✅ Antwoord voor Q{qid} is aangepast naar optie {answers.get(qid)}.")
                else:
                    if row_match is not None and not row_match.empty:
                        row_q = row_match.iloc[0]
                        opt1 = row_q.get('Answer_options_1', 'Optie 1')
                        opt2 = row_q.get('Answer_options_2', 'Optie 2')
                        opt3 = row_q.get('Answer_options_3', 'Optie 3')

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


# ── Screen: Report ────────────────────────────────────────────────────────────

def screen_report():
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import numpy as np

    from report_generator import build_report_input, generate_report_safe
    from pdf_export import export_pdf

    df = get_active_questions()
    answers = st.session_state.answers
    scored = calculate_scores(df, answers)

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

    st.markdown(report_text)

    # Compliance table — corrected headers and explanatory note
    st.markdown("### Compliance Framework Indicators")
    st.info(
        "ℹ️ The table below shows how many of the 265 assessment questions are relevant "
        "to each international framework, and how your organization scored on those specific "
        "questions. The question count reflects the design of this assessment — it is not "
        "a measure of how much of the framework your organization covers or complies with."
    )

    def score_color_emoji(score: float) -> str:
        if score < 2.0: return "🔴"
        if score < 3.0: return "🟠"
        if score < 4.0: return "🟡"
        return "🟢"

    comp_rows = []
    fw_labels = {
        "EU_AI_ACT": "EU AI Act",
        "NIST_AI_RMF": "NIST AI RMF",
        "ISO_42001": "ISO 42001",
        "AI_TRISM": "AI TRiSM",
    }
    for fw_key, fw_label in fw_labels.items():
        data = compliance.get(fw_key, {})
        covered = data.get("covered", 0)
        total_q = data.get("total", 265)
        avg = data.get("avg_score", 0.0)
        pct = round(covered / total_q * 100) if total_q > 0 else 0
        comp_rows.append({
            "Framework": fw_label,
            "Questions in assessment with relevance": f"{covered} / {total_q} ({pct}%)",
            "Your avg. score on these questions": f"{score_color_emoji(avg)}  {avg:.1f} / 5.0",
        })
    st.table(pd.DataFrame(comp_rows))

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

    # ── Admin portal routing ──────────────────────────────────────────────────
    # Access via: https://yourapp.streamlit.app/?admin=true
    # Locally:    http://localhost:8501/?admin=true
    if st.query_params.get("admin", "") == "true":
        if check_admin_password():
            render_admin_portal()
        return  # never show the client app when ?admin=true

    # ── Normal client app routing ─────────────────────────────────────────────
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
