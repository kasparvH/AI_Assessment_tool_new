import json
import os
import shutil
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
from dotenv import load_dotenv

load_dotenv(Path(__file__).parent / ".env")

from excel_engine import (
    LOGO_PATH, OUTPUT_DIR, SESSIONS_DIR,
    load_questions, calculate_scores,
    get_dimension_summary, get_overall_summary,
)
from session_manager import generate_session_code, save_session, delete_session
from config_manager import load_config, save_config, set_test_mode, is_test_mode

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Admin — AI Readiness Assessment",
    page_icon="🔒",
    layout="wide",
    initial_sidebar_state="expanded",
)

ADMIN_PASSWORD = os.getenv("ADMIN_PASSWORD", "straightable2025")

CUSTOM_CSS = """
<style>
    [data-testid="stSidebar"] { background-color: #1B3B6F !important; }
    [data-testid="stSidebar"] * { color: white !important; }
    .stButton > button {
        background-color: #F7941D !important;
        color: white !important;
        border: none !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
    }
    .stButton > button:hover { background-color: #d97c0f !important; }
    .danger-btn > button {
        background-color: #dc3545 !important;
    }
    [data-testid="metric-container"] {
        background-color: white;
        border: 1px solid #e0e0e0;
        border-radius: 8px;
        padding: 12px;
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


# ── Auth ──────────────────────────────────────────────────────────────────────

def check_password() -> bool:
    if st.session_state.get("admin_authenticated"):
        return True
    st.markdown("## 🔒 Admin Portal — Straightable")
    st.markdown("Enter the admin password to continue.")
    pwd = st.text_input("Password", type="password", key="admin_pwd_input")
    if st.button("Login"):
        if pwd == ADMIN_PASSWORD:
            st.session_state.admin_authenticated = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    return False


# ── Data helpers ──────────────────────────────────────────────────────────────

def load_all_sessions() -> list[dict]:
    sessions = []
    for f in sorted(SESSIONS_DIR.glob("*.json"), key=lambda x: x.stat().st_mtime, reverse=True):
        try:
            data = json.loads(f.read_text(encoding="utf-8"))
            data["_file"] = f.name
            sessions.append(data)
        except Exception:
            pass
    return sessions


def session_to_row(s: dict) -> dict:
    answers = s.get("answers", {})
    n_answered = len(answers)
    pct = round(n_answered / 265 * 100, 1)
    phase = s.get("phase", "assessment")
    last_saved = s.get("last_saved_at", "")
    if last_saved:
        try:
            last_saved = datetime.fromisoformat(last_saved).strftime("%d-%m-%Y %H:%M")
        except Exception:
            pass
    return {
        "Code": s.get("session_code", ""),
        "Organization": s.get("org_name", ""),
        "Respondent": s.get("respondent_name", ""),
        "Sector": s.get("sector", ""),
        "Progress": f"{n_answered}/265 ({pct}%)",
        "Phase": phase.capitalize(),
        "Last saved": last_saved,
        "Inconsistencies": len(s.get("inconsistencies_flagged", [])),
    }


def get_report_files() -> list[Path]:
    return sorted(OUTPUT_DIR.glob("rapport_*.docx"), key=lambda x: x.stat().st_mtime, reverse=True)


def get_excel_files() -> list[Path]:
    return sorted(OUTPUT_DIR.glob("assessment_*.xlsx"), key=lambda x: x.stat().st_mtime, reverse=True)


# ── Sidebar ───────────────────────────────────────────────────────────────────

def render_sidebar():
    if LOGO_PATH.exists():
        st.sidebar.image(str(LOGO_PATH), use_container_width=True)
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🔐 Admin Portal")
    st.sidebar.markdown("*Straightable Innovation & Strategy*")
    st.sidebar.markdown("---")

    # ── Test mode toggle ──────────────────────────────────────────────────────
    cfg = load_config()
    test_mode = cfg.get("test_mode", False)
    n_per_dim = cfg.get("test_questions_per_dimension", 4)

    st.sidebar.markdown("### 🧪 Test mode")
    new_test_mode = st.sidebar.toggle(
        "Actief (4 vragen per dimensie)",
        value=test_mode,
        key="test_mode_toggle",
        help="Limiteert het klantportal tot 4 vragen per dimensie zodat je de volledige flow snel kunt testen.",
    )
    if new_test_mode != test_mode:
        cfg["test_mode"] = new_test_mode
        save_config(cfg)
        st.rerun()

    if new_test_mode:
        st.sidebar.markdown(
            f"<div style='background:#F7941D;color:white;padding:6px 10px;"
            f"border-radius:6px;font-size:0.82rem;font-weight:600;'>"
            f"⚠️ TEST MODE AAN — klantportal toont {n_per_dim} vragen/dim "
            f"({7 * n_per_dim} totaal)</div>",
            unsafe_allow_html=True,
        )
    else:
        st.sidebar.markdown(
            "<div style='background:#28a745;color:white;padding:6px 10px;"
            "border-radius:6px;font-size:0.82rem;font-weight:600;'>"
            "✅ PRODUCTIE — klantportal toont alle 265 vragen</div>",
            unsafe_allow_html=True,
        )
    st.sidebar.markdown("---")

    tabs = {
        "📋 Session Overview": "overview",
        "📄 Reports & Downloads": "reports",
        "➕ Create Session": "create",
        "🗑️ Manage Sessions": "manage",
    }
    for label, key in tabs.items():
        if st.sidebar.button(label, use_container_width=True, key=f"nav_{key}"):
            st.session_state.admin_tab = key

    st.sidebar.markdown("---")
    if st.sidebar.button("🚪 Logout", use_container_width=True):
        st.session_state.admin_authenticated = False
        st.rerun()

    return st.session_state.get("admin_tab", "overview")


# ── Tab: Overview ─────────────────────────────────────────────────────────────

def tab_overview():
    st.markdown("## 📋 Session Overview")

    sessions = load_all_sessions()

    if not sessions:
        st.info("No sessions found yet.")
        return

    # Summary metrics
    total = len(sessions)
    completed = sum(1 for s in sessions if s.get("phase") == "completed")
    in_progress = sum(1 for s in sessions if s.get("phase") == "assessment")
    total_inconsistencies = sum(len(s.get("inconsistencies_flagged", [])) for s in sessions)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total sessions", total)
    c2.metric("Completed", completed)
    c3.metric("In progress", in_progress)
    c4.metric("Total inconsistencies flagged", total_inconsistencies)

    st.markdown("---")

    rows = [session_to_row(s) for s in sessions]
    df = pd.DataFrame(rows)

    # Filter
    search = st.text_input("🔍 Search by organization or code", "")
    if search:
        mask = (
            df["Organization"].str.contains(search, case=False, na=False) |
            df["Code"].str.contains(search, case=False, na=False)
        )
        df = df[mask]

    st.dataframe(df, use_container_width=True, hide_index=True)

    # Detail view
    st.markdown("---")
    st.markdown("### Session detail")
    codes = [s.get("session_code", "") for s in sessions]
    selected_code = st.selectbox("Select a session to inspect", ["— select —"] + codes)

    if selected_code and selected_code != "— select —":
        s = next((x for x in sessions if x.get("session_code") == selected_code), None)
        if s:
            st.markdown(f"**Organization:** {s.get('org_name')}  |  **Respondent:** {s.get('respondent_name')}  |  **Sector:** {s.get('sector')}")
            st.markdown(f"**Started:** {s.get('started_at', '')}  |  **Last saved:** {s.get('last_saved_at', '')}")
            st.markdown(f"**Phase:** {s.get('phase', '').capitalize()}  |  **Current dimension:** {s.get('current_dimension', '')}")

            answers = {int(k): v for k, v in s.get("answers", {}).items()}
            n = len(answers)
            st.markdown(f"**Answers given:** {n} / 265")

            if n > 0:
                try:
                    df_q = load_questions()
                    scored = calculate_scores(df_q, answers)
                    dim_sum = get_dimension_summary(scored)
                    overall = get_overall_summary(dim_sum)
                    st.markdown(f"**Current score:** {overall['overall_score_0_5']:.1f}/5.0 — {overall['overall_pct']:.0f}% — *{overall['maturity_tier']}*")

                    st.markdown("**Scores per dimension:**")
                    for _, row in dim_sum.sort_values("score_pct", ascending=False).iterrows():
                        col1, col2 = st.columns([5, 1])
                        col1.progress(row["score_pct"] / 100, text=f"{row['dimension']}")
                        col2.markdown(f"**{row['score_0_5']:.1f}/5**")
                except Exception as e:
                    st.warning(f"Could not compute scores: {e}")

            inconsistencies = s.get("inconsistencies_flagged", [])
            if inconsistencies:
                with st.expander(f"⚠️ {len(inconsistencies)} inconsistencies"):
                    for inc in inconsistencies:
                        st.markdown(f"**Q{inc.get('question_id')}** — {inc.get('question_text', '')}")
                        st.markdown(f"- Severity: `{inc.get('severity', '?')}`")
                        st.markdown(f"- {inc.get('explanation', '')}")
                        st.markdown("---")

            # ── Report & export actions ──
            if n > 0:
                st.markdown("---")
                st.markdown("### 📄 Rapport & export")
                col_r1, col_r2 = st.columns(2)

                with col_r1:
                    if st.button("📝 Genereer Word-rapport", key=f"gen_report_{selected_code}", use_container_width=True):
                        try:
                            from report_generator import build_report_input, generate_report_safe
                            from word_export import generate_word_report
                            from excel_engine import get_compliance_summary

                            df_q = load_questions()
                            scored = calculate_scores(df_q, answers)
                            dim_sum = get_dimension_summary(scored)
                            overall_r = get_overall_summary(dim_sum)
                            compliance = get_compliance_summary(scored)

                            report_input = build_report_input(
                                s.get("org_name", ""),
                                s.get("respondent_name", ""),
                                s.get("sector", ""),
                                overall_r, dim_sum, compliance,
                                s.get("inconsistencies_flagged", []),
                            )
                            with st.spinner("Rapport wordt geschreven..."):
                                report_text, success = generate_report_safe(report_input)

                            docx_path = str(OUTPUT_DIR / f"rapport_{selected_code}.docx")
                            generate_word_report(
                                report_text=report_text,
                                dim_summary=dim_sum,
                                overall=overall_r,
                                compliance=compliance,
                                org_name=s.get("org_name", ""),
                                respondent_name=s.get("respondent_name", ""),
                                sector=s.get("sector", ""),
                                logo_path=str(LOGO_PATH),
                                inconsistencies=s.get("inconsistencies_flagged", []),
                                output_path=docx_path,
                            )
                            st.session_state[f"_docx_path_{selected_code}"] = docx_path
                            st.success("✅ Word-rapport gegenereerd!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Fout bij genereren rapport: {e}")

                with col_r2:
                    if st.button("📊 Genereer Excel-export", key=f"gen_excel_{selected_code}", use_container_width=True):
                        try:
                            from excel_engine import export_filled_excel
                            df_q = load_questions()
                            xlsx_path = str(OUTPUT_DIR / f"assessment_{selected_code}.xlsx")
                            export_filled_excel(df_q, answers, xlsx_path)
                            st.session_state[f"_xlsx_path_{selected_code}"] = xlsx_path
                            st.success("✅ Excel-export klaar!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Fout bij Excel-export: {e}")

                # Download buttons if files exist
                docx_path = st.session_state.get(f"_docx_path_{selected_code}") or str(OUTPUT_DIR / f"rapport_{selected_code}.docx")
                if Path(docx_path).exists():
                    with open(docx_path, "rb") as fp:
                        st.download_button(
                            "⬇️ Download Word-rapport",
                            fp,
                            file_name=f"AI_Rapport_{s.get('org_name', selected_code)}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"dl_docx_{selected_code}",
                        )

                xlsx_path = st.session_state.get(f"_xlsx_path_{selected_code}") or str(OUTPUT_DIR / f"assessment_{selected_code}.xlsx")
                if Path(xlsx_path).exists():
                    with open(xlsx_path, "rb") as fx:
                        st.download_button(
                            "⬇️ Download Excel",
                            fx,
                            file_name=f"AI_Assessment_{s.get('org_name', selected_code)}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"dl_xlsx_{selected_code}",
                        )


# ── Tab: Reports & Downloads ──────────────────────────────────────────────────

def tab_reports():
    st.markdown("## 📄 Reports & Downloads")

    word_files = get_report_files()
    xlsx_files = get_excel_files()

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("### 📝 Word-rapporten")
        if not word_files:
            st.info("Nog geen rapporten gegenereerd. Ga naar 'Session Overview', selecteer een sessie en klik op 'Genereer Word-rapport'.")
        for f in word_files:
            code = f.stem.replace("rapport_", "")
            sessions = load_all_sessions()
            s = next((x for x in sessions if x.get("session_code") == code), {})
            org = s.get("org_name", code)
            date_str = datetime.fromtimestamp(f.stat().st_mtime).strftime("%d-%m-%Y %H:%M")
            with st.expander(f"📝 {org} — {date_str}"):
                st.markdown(f"**Sessie:** `{code}`  |  **Bestand:** `{f.name}`")
                with open(f, "rb") as fp:
                    st.download_button(
                        "⬇️ Download Word-rapport",
                        fp,
                        file_name=f"AI_Rapport_{org}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"dl_docx_{code}",
                    )

    with col2:
        st.markdown("### 📊 Excel-exports")
        if not xlsx_files:
            st.info("Nog geen Excel-exports beschikbaar.")
        for f in xlsx_files:
            code = f.stem.replace("assessment_", "")
            sessions = load_all_sessions()
            s = next((x for x in sessions if x.get("session_code") == code), {})
            org = s.get("org_name", code)
            date_str = datetime.fromtimestamp(f.stat().st_mtime).strftime("%d-%m-%Y %H:%M")
            with st.expander(f"📊 {org} — {date_str}"):
                st.markdown(f"**Sessie:** `{code}`  |  **Bestand:** `{f.name}`")
                with open(f, "rb") as fp:
                    st.download_button(
                        "⬇️ Download Excel",
                        fp,
                        file_name=f"AI_Assessment_{org}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_xlsx_{code}",
                    )


# ── Tab: Create Session ───────────────────────────────────────────────────────

def tab_create():
    st.markdown("## ➕ Create Session for Customer")
    st.markdown("Pre-fill customer details and generate a session code to share with them.")

    with st.form("create_session_form"):
        org = st.text_input("Organization name *")
        resp = st.text_input("Respondent name *")
        role = st.text_input("Role / function")
        sector = st.selectbox("Sector", SECTORS)
        note = st.text_area("Internal note (not shown to customer)", placeholder="e.g. Follow-up call booked for 15 July")
        submitted = st.form_submit_button("Generate Session Code →")

    if submitted:
        if not org or not resp:
            st.warning("Organization name and respondent name are required.")
        else:
            code = generate_session_code()
            state = {
                "session_code": code,
                "org_name": org,
                "respondent_name": resp,
                "role": role,
                "sector": sector,
                "started_at": datetime.now().isoformat(),
                "last_saved_at": datetime.now().isoformat(),
                "current_dimension": DIMENSIONS_ORDER[0],
                "answers": {},
                "uploaded_documents": [],
                "inconsistencies_flagged": [],
                "phase": "assessment",
                "_admin_note": note,
            }
            save_session(code, state)

            st.success(f"✅ Session created for **{org}**")
            st.markdown("---")

            col1, col2 = st.columns(2)
            with col1:
                st.markdown("### 📋 Session Code")
                st.markdown(f"## `{code}`")
                st.markdown("Share this code with the customer so they can start and resume their assessment.")

            with col2:
                st.markdown("### 📧 Email template")
                email_body = f"""Dear {resp},

We have prepared an AI Readiness Assessment for {org}.

To start your assessment, go to:
[Your assessment URL, e.g. http://localhost:8501]

Your personal session code is: {code}

Save this code — you can use it to pause and resume the assessment at any time.

The assessment covers 7 key dimensions and takes approximately 45–60 minutes.

Kind regards,
Straightable Innovation & Strategy
"""
                st.code(email_body, language=None)
                st.download_button(
                    "⬇️ Download as .txt",
                    email_body,
                    file_name=f"invite_{code}.txt",
                    mime="text/plain",
                    key=f"dl_invite_{code}",
                )


# ── Tab: Manage Sessions ──────────────────────────────────────────────────────

def tab_manage():
    st.markdown("## 🗑️ Manage Sessions")
    st.warning("⚠️ Deleting a session is permanent and cannot be undone.")

    sessions = load_all_sessions()

    if not sessions:
        st.info("No sessions found.")
        return

    codes = [f"{s.get('session_code')} — {s.get('org_name', '?')} ({s.get('phase', '?')})"
             for s in sessions]
    selected = st.selectbox("Select session", ["— select —"] + codes)

    if selected and selected != "— select —":
        code = selected.split(" — ")[0]
        s = next((x for x in sessions if x.get("session_code") == code), None)

        if s:
            st.markdown(f"**Organization:** {s.get('org_name')}  |  **Respondent:** {s.get('respondent_name')}")
            st.markdown(f"**Phase:** {s.get('phase', '').capitalize()}  |  **Answers:** {len(s.get('answers', {}))}/265")
            st.markdown(f"**Last saved:** {s.get('last_saved_at', '')}")

            st.markdown("---")
            col1, col2 = st.columns(2)

            with col1:
                st.markdown("**Mark as completed**")
                st.markdown("Closes the session without deleting it. The customer can no longer answer questions.")
                if st.button("✅ Mark completed", key=f"complete_{code}"):
                    s["phase"] = "completed"
                    save_session(code, s)
                    st.success(f"Session {code} marked as completed.")
                    st.rerun()

            with col2:
                st.markdown("**Delete session**")
                st.markdown("Permanently removes the session file. Reports in `/output` are NOT deleted.")
                confirm = st.checkbox(f"I confirm I want to delete session `{code}`", key=f"confirm_{code}")
                if st.button("🗑️ Delete session", key=f"delete_{code}"):
                    if confirm:
                        delete_session(code)
                        st.success(f"Session {code} deleted.")
                        st.rerun()
                    else:
                        st.error("Please check the confirmation box first.")

    st.markdown("---")
    st.markdown("### Bulk delete completed sessions")
    completed_sessions = [s for s in sessions if s.get("phase") == "completed"]
    if not completed_sessions:
        st.info("No completed sessions to bulk delete.")
    else:
        st.markdown(f"Found **{len(completed_sessions)}** completed sessions.")
        confirm_bulk = st.checkbox("I confirm I want to delete all completed sessions")
        if st.button("🗑️ Delete all completed sessions"):
            if confirm_bulk:
                for s in completed_sessions:
                    delete_session(s.get("session_code", ""))
                st.success(f"Deleted {len(completed_sessions)} completed sessions.")
                st.rerun()
            else:
                st.error("Please check the confirmation box first.")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    if not check_password():
        return

    tab = render_sidebar()

    if tab == "overview":
        tab_overview()
    elif tab == "reports":
        tab_reports()
    elif tab == "create":
        tab_create()
    elif tab == "manage":
        tab_manage()
    else:
        tab_overview()


if __name__ == "__main__":
    main()
