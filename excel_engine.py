from pathlib import Path
import pandas as pd

ROOT = Path(__file__).parent
ASSETS_DIR = ROOT / "assets"
DATA_DIR = ROOT / "data"
SESSIONS_DIR = ROOT / "sessions"
OUTPUT_DIR = ROOT / "output"
LOGO_PATH = ASSETS_DIR / "Logo_oranje.png"
EXCEL_PATH = DATA_DIR / "AI_Assessment_Framework_V6.xlsx"

SESSIONS_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

SCORE_MAP = {1: 1, 2: 3, 3: 5}

MATURITY_TIERS = [
    (0, 34.9, "Emerging"),
    (35, 59.9, "Developing"),
    (60, 79.9, "Accelerating"),
    (80, 100, "Leading"),
]


def load_questions(filepath: str = None) -> pd.DataFrame:
    path = filepath or str(EXCEL_PATH)
    df = pd.read_excel(path, sheet_name="Sheet")
    df.columns = [c.strip() for c in df.columns]
    df["Scoring_rule"] = "1→1, 2→3, 3→5"
    df["Weight"] = pd.to_numeric(df["Weight"], errors="coerce").fillna(1)
    df = df.set_index("Question_id")
    return df


def calculate_scores(df: pd.DataFrame, answers: dict) -> pd.DataFrame:
    df = df.copy()
    df["Selected_option"] = None
    df["Points"] = None
    df["Weighted_points"] = None
    df["Max_points"] = None
    df["Normalized_0_5"] = None

    for qid, opt in answers.items():
        if qid in df.index and opt in (1, 2, 3):
            pts = SCORE_MAP[opt]
            w = df.at[qid, "Weight"]
            max_pts = 5 * w
            df.at[qid, "Selected_option"] = opt
            df.at[qid, "Points"] = pts
            df.at[qid, "Weighted_points"] = pts * w
            df.at[qid, "Max_points"] = max_pts
            df.at[qid, "Normalized_0_5"] = (pts * w / max_pts) * 5
    return df


def get_dimension_summary(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for dim, grp in df.groupby("Dimension"):
        answered = grp.dropna(subset=["Weighted_points"])
        n_answered = len(answered)
        sum_wp = answered["Weighted_points"].sum()
        sum_mp = answered["Max_points"].sum()
        score_0_5 = (sum_wp / sum_mp * 5) if sum_mp > 0 else 0.0
        score_pct = (sum_wp / sum_mp * 100) if sum_mp > 0 else 0.0

        def count_yes(col):
            if col not in grp.columns:
                return 0
            subset = answered if n_answered > 0 else grp
            return int((subset[col].str.upper() == "YES").sum()) if n_answered > 0 else 0

        rows.append({
            "dimension": dim,
            "n_questions": len(grp),
            "n_answered": n_answered,
            "sum_weighted_points": sum_wp,
            "sum_max_points": sum_mp,
            "score_0_5": round(score_0_5, 2),
            "score_pct": round(score_pct, 1),
            "eu_ai_act_count": count_yes("EU_AI_ACT"),
            "nist_count": count_yes("NIST_AI_RMF"),
            "iso_count": count_yes("ISO_42001"),
            "ai_trism_count": count_yes("AI_TRISM"),
        })
    return pd.DataFrame(rows)


def get_overall_summary(dim_summary: pd.DataFrame) -> dict:
    total_wp = dim_summary["sum_weighted_points"].sum()
    total_mp = dim_summary["sum_max_points"].sum()
    score_0_5 = round((total_wp / total_mp * 5) if total_mp > 0 else 0.0, 2)
    score_pct = round((total_wp / total_mp * 100) if total_mp > 0 else 0.0, 1)

    tier = "Emerging"
    for lo, hi, name in MATURITY_TIERS:
        if lo <= score_pct <= hi:
            tier = name
            break

    return {
        "overall_score_0_5": score_0_5,
        "overall_pct": score_pct,
        "maturity_tier": tier,
    }


def get_compliance_summary(df: pd.DataFrame) -> dict:
    answered = df.dropna(subset=["Weighted_points"])
    result = {}
    for fw, col in [("EU_AI_ACT", "EU_AI_ACT"), ("NIST_AI_RMF", "NIST_AI_RMF"),
                    ("ISO_42001", "ISO_42001"), ("AI_TRISM", "AI_TRISM")]:
        if col not in df.columns:
            result[fw] = {"covered": 0, "total": len(df), "avg_score": 0}
            continue
        covered = df[df[col].str.upper() == "YES"] if df[col].notna().any() else df.iloc[0:0]
        covered_answered = answered[answered[col].str.upper() == "YES"] if answered[col].notna().any() else answered.iloc[0:0]
        avg = 0.0
        if len(covered_answered) > 0:
            mp = covered_answered["Max_points"].sum()
            wp = covered_answered["Weighted_points"].sum()
            avg = round((wp / mp * 5) if mp > 0 else 0.0, 2)
        result[fw] = {
            "covered": len(covered),
            "total": len(df),
            "avg_score": avg,
        }
    return result


def export_filled_excel(df: pd.DataFrame, answers: dict, output_path: str):
    scored = calculate_scores(df, answers)
    scored = scored.reset_index()
    scored.to_excel(output_path, index=False, sheet_name="Sheet")
