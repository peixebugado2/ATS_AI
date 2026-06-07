import os
import re
import json
import time
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Tuple

import streamlit as st
from dotenv import load_dotenv
from google import genai
from docx import Document
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_LEFT

try:
    from pypdf import PdfReader
except Exception:
    PdfReader = None

load_dotenv()

APP_TITLE = "CareerOps Studio"
APP_SUBTITLE = "Executive Resume Optimization Platform"
DEFAULT_MODEL = "gemini-2.5-flash-lite"
ROOT = Path(__file__).parent
TEMPLATE_DIR = ROOT / "templates"

HISTORY_FILE = ROOT / "careerops_history.json"
MAX_HISTORY_ITEMS = 30

MODEL_OPTIONS = [
    "gemini-2.5-flash-lite",
    "gemini-2.5-flash",
    "gemini-2.0-flash-lite",
    "gemini-2.0-flash",
    "gemini-1.5-flash",
    "gemini-1.5-flash-8b",
]

CONTEXT_PROFILES = {
    "Fast": {"instructions": 8000, "master": 26000, "manager": 8000, "director": 8000},
    "Balanced": {"instructions": 14000, "master": 48000, "manager": 12000, "director": 12000},
    "Deep": {"instructions": 24000, "master": 80000, "manager": 18000, "director": 18000},
}

MANDATORY_BULLET_COUNTS = """
MANDATORY BULLET COUNTS

Director CV:
- Sr. EMEA Transport Manager: exactly 7 achievement lines
- Director Logistics: exactly 8 achievement lines
- Trade Compliance Manager EMEA: do not include this role
- Head of Logistics EMEA: exactly 7 achievement lines
- Sr. EMEA Logistics & Spare Parts Inventory Manager: exactly 5 achievement lines

Manager CV:
- Sr. EMEA Transport Manager: exactly 7 achievement lines
- Director Logistics: do not include this role
- Trade Compliance Manager EMEA: exactly 8 achievement lines
- Head of Logistics EMEA: exactly 7 achievement lines
- Sr. EMEA Logistics & Spare Parts Inventory Manager: exactly 5 achievement lines

These counts are mandatory and must be followed exactly.
For Manager CV, Trade Compliance Manager EMEA must have exactly 8 achievement lines, not 7.
Do not add extra roles. Do not add fewer or more achievement lines.
"""

FINAL_OUTPUT_POLICY = """
FINAL OUTPUT POLICY

Process internally and do not show reasoning, scoring steps, long analysis, internal ranking, or achievement selection process.

The final answer must contain only:
1. Selected CV Type: Manager CV or Director CV
2. ATS Validation Status table with key scores/statistics only
3. Final tailored CV
4. Cover Letter only when requested in the output package
5. Final ATS Report with realistic simulated scores only

Do not include:
- Detailed reasoning
- Full keyword extraction
- Full achievement scoring
- Long explanations
- Internal validation steps
- Draft alternatives
"""

COVER_LETTER_EXTRACTION_RULES = """
COVER LETTER EXTRACTION RULES

The cover letter must be generated from the Job Description and the selected verified achievements.

Internal extraction requirements:
- Extract the company name from the Job Description whenever available.
- Extract the exact job title from the Job Description whenever available.
- Extract the hiring context, location, seniority, department and business area whenever available.
- Extract the most important responsibilities, hard skills, soft skills, systems, certifications and ATS keywords.
- Use only achievements selected for the tailored CV.
- Never use generic placeholders if the information exists in the Job Description.
- Never write [Company Name], [Hiring Manager Name], [Job Title], or similar placeholders in the final output.
- If the company name is available, address the letter to: Dear Hiring Team at [Company Name],
- If the company name is not available, address the letter to: Dear Hiring Manager,
- If the job title is available, mention the exact job title naturally in the opening paragraph.
- If the job title is not available, infer the closest professional title from the Job Description.
- The letter must be personalized to the role requirements, company context, ATS keywords and selected achievements.
- The letter must not invent achievements, metrics, certifications, systems, names or company information.
- The final cover letter must be 1,790 to 1,810 characters including spaces.
- Target exactly 1,800 characters including spaces.
- If the cover letter is above 1,810 characters, shorten it before final output.
- If the cover letter is below 1,790 characters, expand it before final output.
- Do not show character counts in the final output.
"""

ATS_VALIDATION_RULES = """
ATS VALIDATION RULES

The ATS validation is a simulation based on the Job Description, the generated CV, and the operating instructions.
It is not a live connection to LinkedIn, Jobscan, SkillSyncer, Resume Worded, Rezi or any external website.

The system must internally simulate these ATS-style checks:
- LinkedIn ATS style keyword coverage
- Jobscan style hard skill and job title match
- SkillSyncer style skill extraction and skill matching
- Resume Worded style impact, clarity and leadership match
- Rezi style ATS formatting and keyword alignment

Skill extraction requirements:
- Extract hard skills directly from the Job Description.
- Extract soft skills directly from the Job Description.
- Extract tools, systems, certifications, industries and operational keywords directly from the Job Description.
- Match extracted skills against the tailored CV.
- ATS skills shown in the CV must be directly aligned with the extracted Job Description skills.
- Additional ATS Skills must contain exactly 10 skills.
- Do not invent skills that are not supported by the Job Description or the candidate experience.

Final ATS report requirements:
- Keep it concise.
- Show realistic scores only.
- Never claim the scores came from live external sites.
- Never claim 100% unless every critical keyword is naturally covered.
"""

STRICT_OUTPUT_FORMAT_RULES = """
STRICT OUTPUT FORMAT RULES

The final response must follow this exact structure:

SELECTED CV TYPE
Manager CV
or
Director CV

ATS VALIDATION STATUS
A concise table only.

FINAL TAILORED CV
Use the exact role headings below when included:
Sr. EMEA Transport Manager
Director Logistics
Trade Compliance Manager EMEA
Head of Logistics EMEA
Sr. EMEA Logistics & Spare Parts Inventory Manager

Each achievement line must start with "ACHIEVEMENT: ".
Each achievement line must be 170 to 190 characters including spaces AFTER the "ACHIEVEMENT: " prefix is removed.
Do not use bullet symbols.

COVER LETTER
Only include this section when requested.
The cover letter body must be 1,790 to 1,810 characters including spaces, target exactly 1,800.
Do not include placeholder text.

FINAL ATS REPORT
A concise report only.

Do not show internal reasoning, keyword dumps, scoring logs, chain-of-thought, draft alternatives or validation steps.
"""


def load_history() -> List[Dict]:
    if not HISTORY_FILE.exists():
        return []

    try:
        with open(HISTORY_FILE, "r", encoding="utf-8") as file:
            data = json.load(file)

        if isinstance(data, list):
            return data

        return []

    except Exception:
        return []


def save_history(history: List[Dict]) -> None:
    try:
        history = history[:MAX_HISTORY_ITEMS]

        with open(HISTORY_FILE, "w", encoding="utf-8") as file:
            json.dump(history, file, ensure_ascii=False, indent=2)

    except Exception:
        pass


def add_history_item(kind: str, title: str, job_description: str, result: str, model: str, context_profile: str) -> None:
    history = load_history()

    item = {
        "id": datetime.now().strftime("%Y%m%d_%H%M%S"),
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "kind": kind,
        "title": title[:120] if title else kind,
        "job_description": job_description or "",
        "result": result or "",
        "model": model or "",
        "context_profile": context_profile or "",
    }

    history.insert(0, item)
    save_history(history)


def get_history_title(job_description: str, fallback: str = "CareerOps Output") -> str:
    if not job_description:
        return fallback

    clean_lines = [line.strip() for line in job_description.splitlines() if line.strip()]

    for line in clean_lines[:8]:
        if len(line) >= 8:
            return line[:100]

    return fallback


def render_history_panel() -> None:
    st.markdown("### History")

    history = load_history()

    if not history:
        st.caption("No saved outputs yet.")
        return

    labels = [
        f"{item.get('created_at', '')} · {item.get('kind', 'Output')} · {item.get('title', 'Untitled')}"
        for item in history
    ]

    selected_index = st.selectbox(
        "Saved outputs",
        range(len(labels)),
        format_func=lambda index: labels[index],
        key="history_selected_index",
    )

    selected_item = history[selected_index]

    c1, c2 = st.columns(2)

    with c1:
        if st.button("Load", use_container_width=True):
            st.session_state["last_result"] = selected_item.get("result", "")
            st.session_state["last_job"] = selected_item.get("job_description", "")
            st.session_state["loaded_history_item"] = selected_item
            st.success("History item loaded.")

    with c2:
        if st.button("Clear", use_container_width=True):
            save_history([])
            st.session_state.pop("loaded_history_item", None)
            st.success("History cleared.")
            st.rerun()

    with st.expander("Preview selected"):
        st.caption(f"Model: {selected_item.get('model', '')}")
        st.caption(f"Context: {selected_item.get('context_profile', '')}")
        st.text((selected_item.get("result", "") or "")[:2500])

def inject_css():
    st.markdown(
        """
        <style>
        :root {
            --bg: #0a0f1a;
            --surface: #111827;
            --surface-2: #1a2332;
            --text: #f3f4f6;
            --muted: #9ca3af;
            --border: #2a3444;
            --border-soft: #1f2937;
            --primary: #d4af37;
            --primary-2: #b89020;
            --success: #15803d;
            --danger: #b42318;
            --shadow: 0 10px 30px rgba(16, 24, 40, .07);
        }
        .stApp {
            color: var(--text);
            background:
        radial-gradient(circle at top left, rgba(212,175,55,.12), transparent 30%),
        radial-gradient(circle at top right, rgba(80,120,255,.08), transparent 35%),
        linear-gradient(180deg, #05070c 0%, #0b1120 100%);
        }
        [data-testid="stHeader"] {
            background: rgba(248, 250, 252, .92);
            backdrop-filter: blur(10px);
            border-bottom: 1px solid var(--border-soft);
        }
        .block-container {
            max-width: 1180px;
            padding-top: 1.4rem;
            padding-bottom: 2rem;
        }
        [data-testid="stSidebar"] {
            background:
            linear-gradient(180deg,#0a0f1a 0%,#111827 100%);
            border-right: 1px solid rgba(255,255,255,.08);
        }
        [data-testid="stSidebar"] * { color: #f8fafc; }
        [data-testid="stSidebar"] .stCaption,
        [data-testid="stSidebar"] small,
        [data-testid="stSidebar"] p {
            color: rgba(248,250,252,.68) !important;
        }
        .app-header {
            background: var(--surface);
            border: 1px solid var(--border-soft);
            border-radius: 18px;
            padding: 24px 28px;
            box-shadow: var(--shadow);
            margin-bottom: 18px;
        }
        .app-eyebrow {
            color: var(--primary);
            font-size: .76rem;
            font-weight: 800;
            letter-spacing: .16em;
            text-transform: uppercase;
            margin-bottom: 8px;
        }
        .app-title {
            font-size: 2.05rem;
            line-height: 1.1;
            font-weight: 820;
            letter-spacing: -.045em;
            margin: 0;
            color: #101828;
        }
        .app-subtitle {
            margin-top: 10px;
            color: var(--muted);
            max-width: 760px;
            font-size: .98rem;
            line-height: 1.55;
        }
        .status-row {
            display: grid;
            grid-template-columns: repeat(4, minmax(0, 1fr));
            gap: 12px;
            margin: 0 0 18px 0;
        }
        .status-card {
            background: var(--surface);
            border: 1px solid var(--border-soft);
            border-radius: 14px;
            padding: 14px 16px;
            box-shadow: 0 4px 16px rgba(16, 24, 40, .045);
        }
        .status-label {
            color: var(--muted);
            font-size: .72rem;
            letter-spacing: .12em;
            text-transform: uppercase;
            font-weight: 750;
        }
        .status-value {
            color: #101828;
            font-size: .98rem;
            font-weight: 760;
            margin-top: 6px;
            overflow-wrap: anywhere;
        }
        .card {
            background: var(--surface);
            border: 1px solid var(--border-soft);
            border-radius: 18px;
            padding: 24px 26px;
            box-shadow: var(--shadow);
            margin-bottom: 16px;
        }
        .section-title {
            font-size: 1.28rem;
            font-weight: 820;
            letter-spacing: -.03em;
            color: #101828;
            margin: 0 0 6px 0;
        }
        .section-note {
            color: var(--muted);
            margin: 0 0 18px 0;
            line-height: 1.5;
        }
        .result-shell {
            background: var(--surface);
            border: 1px solid var(--border-soft);
            border-radius: 18px;
            padding: 26px;
            margin-top: 18px;
            box-shadow: var(--shadow);
            color: var(--text);
        }
        .result-shell h1,
        .result-shell h2,
        .result-shell h3 { color: #101828; }
        .audit-box {
            border: 1px solid #bbf7d0;
            background: #f0fdf4;
            color: #14532d;
            border-radius: 14px;
            padding: 12px 14px;
            margin: 14px 0;
            font-size: .92rem;
        }
        .footer-note {
            text-align:center;
            color: var(--muted);
            font-size:.80rem;
            padding:24px 0 10px;
        }
        .stButton > button,
        .stDownloadButton > button {
            border-radius: 12px !important;
            border: 1px solid var(--border) !important;
            font-weight: 750 !important;
            padding: .72rem 1rem !important;
            background: var(--surface) !important;
            color: var(--text) !important;
        }
        .stButton > button:hover,
        .stDownloadButton > button:hover {
            border-color: var(--primary) !important;
            color: var(--primary) !important;
        }
        .stButton > button[kind="primary"] {
            color: #ffffff !important;
            border: 0 !important;
            background: linear-gradient(90deg, var(--primary), var(--primary-2)) !important;
            box-shadow: 0 8px 20px rgba(31, 78, 121, .22);
        }
        textarea,
        input,
        .stTextArea textarea,
        .stTextInput input {
            border-radius: 12px !important;
            border-color: var(--border) !important;
            background: #ffffff !important;
            color: var(--text) !important;
        }
        [data-baseweb="select"] > div {
            border-radius: 12px !important;
            border-color: var(--border) !important;
            background: #ffffff !important;
        }
        [data-testid="stFileUploader"] {
            border: 1px dashed rgba(255,255,255,.28);
            background: rgba(255,255,255,.05);
            border-radius: 14px;
            padding: 10px;
        }
        div[data-testid="stTabs"] button {
            border-radius: 999px !important;
            padding: 10px 18px !important;
            color: var(--text) !important;
        }
        div[data-testid="stTabs"] button[aria-selected="true"] {
            background: #e8eef6 !important;
            color: var(--primary) !important;
            font-weight: 800 !important;
        }
        hr { border-color: var(--border-soft); }
        .stMetric {
            background: var(--surface-2);
            border: 1px solid var(--border-soft);
            border-radius: 12px;
            padding: 12px;
        }
        @media (max-width: 900px) {
            .status-row { grid-template-columns: repeat(2, minmax(0, 1fr)); }
            .app-title { font-size: 1.62rem; }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

def load_template_text(filename: str, fallback: str = "") -> str:
    path = TEMPLATE_DIR / filename
    if not path.exists():
        return fallback
    try:
        if path.suffix.lower() == ".docx":
            doc = Document(path)
            parts = []
            for p in doc.paragraphs:
                if p.text.strip():
                    parts.append(p.text)
            for table in doc.tables:
                for row in table.rows:
                    cells = [cell.text.strip() for cell in row.cells]
                    if any(cells):
                        parts.append(" | ".join(cells))
            return "\n".join(parts)
        return path.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        return fallback


def load_first_available(filenames: List[str]) -> str:
    for filename in filenames:
        text = load_template_text(filename)
        if text.strip():
            return text
    return ""


def get_clients():
    keys = [
        os.getenv("GEMINI_API_KEY_1"),
        os.getenv("GEMINI_API_KEY_2"),
        os.getenv("GEMINI_API_KEY_3"),
        os.getenv("GEMINI_API_KEY_4"),
        os.getenv("GEMINI_API_KEY_5"),
        os.getenv("GEMINI_API_KEY"),
    ]

    clients = []

    for index, key in enumerate(keys, start=1):
        if key and key.strip():
            clients.append({
                "label": f"API Key {index}",
                "client": genai.Client(api_key=key.strip()),
            })

    if not clients:
        st.error(
            "No Gemini API key found. Add GEMINI_API_KEY_1=your_key_here to the .env file. "
            "You can also add GEMINI_API_KEY_2, GEMINI_API_KEY_3, GEMINI_API_KEY_4 and GEMINI_API_KEY_5."
        )
        st.stop()

    return clients


def read_uploaded_file(uploaded_file) -> str:
    if uploaded_file is None:
        return ""
    uploaded_file.seek(0)
    name = uploaded_file.name.lower()
    data = uploaded_file.read()
    if name.endswith((".txt", ".md", ".csv")):
        return data.decode("utf-8", errors="ignore")
    if name.endswith(".docx"):
        doc = Document(BytesIO(data))
        parts = []
        for p in doc.paragraphs:
            if p.text.strip():
                parts.append(p.text)
        for table in doc.tables:
            for row in table.rows:
                cells = [cell.text.strip() for cell in row.cells]
                if any(cells):
                    parts.append(" | ".join(cells))
        return "\n".join(parts)
    if name.endswith(".pdf"):
        if PdfReader is None:
            return "[PDF uploaded, but pypdf is not installed.]"
        reader = PdfReader(BytesIO(data))
        return "\n".join([(page.extract_text() or "") for page in reader.pages])
    return data.decode("utf-8", errors="ignore")


def truncate_text(text: str, max_chars: int) -> str:
    text = text or ""
    if len(text) <= max_chars:
        return text
    return text[:max_chars] + "\n\n[Content truncated by application limit]"


def make_docx(content: str, title: str = "CareerOps Studio Output") -> bytes:
    doc = Document()
    styles = doc.styles
    styles["Normal"].font.name = "Calibri"
    doc.add_heading(title, level=1)
    for line in content.splitlines():
        stripped = line.strip()
        if not stripped:
            doc.add_paragraph("")
        elif stripped.startswith("# "):
            doc.add_heading(stripped[2:], level=1)
        elif stripped.startswith("## "):
            doc.add_heading(stripped[3:], level=2)
        elif stripped.startswith("### "):
            doc.add_heading(stripped[4:], level=3)
        else:
            doc.add_paragraph(stripped)
    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()


def make_pdf(content: str, title: str = "CareerOps Studio Output") -> bytes:
    bio = BytesIO()
    pdf = SimpleDocTemplate(
        bio,
        pagesize=A4,
        rightMargin=1.65 * cm,
        leftMargin=1.65 * cm,
        topMargin=1.55 * cm,
        bottomMargin=1.55 * cm,
    )
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="ExecutiveTitle", parent=styles["Title"], alignment=TA_LEFT, fontSize=17, leading=21, spaceAfter=12))
    styles.add(ParagraphStyle(name="ExecutiveNormal", parent=styles["Normal"], fontSize=9.3, leading=12.2, spaceAfter=4.8))
    story = [Paragraph(title, styles["ExecutiveTitle"])]
    for line in content.splitlines():
        stripped = line.strip()
        if not stripped:
            story.append(Spacer(1, 0.18 * cm))
        else:
            safe = stripped.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            story.append(Paragraph(safe, styles["ExecutiveNormal"]))
    pdf.build(story)
    return bio.getvalue()


def clean_model_text(text: str) -> str:
    if not text:
        return ""
    markers = ["done thinking.", "</think>"]
    lower = text.lower()
    for marker in markers:
        idx = lower.rfind(marker)
        if idx != -1:
            text = text[idx + len(marker):]
            break
    return text.strip()


def call_model(prompt: str, primary_model: str, failover_models: List[str]) -> Tuple[str, str, List[str]]:
    clients = get_clients()

    models_to_try = []
    for item in [primary_model] + failover_models:
        item = (item or "").strip()
        if item and item not in models_to_try:
            models_to_try.append(item)

    if len(prompt) > 180000:
        prompt = prompt[:180000] + "\n\n[Prompt truncated by application limit]"

    attempts = []
    last_error = None

    for client_info in clients:
        client_label = client_info["label"]
        client = client_info["client"]

        for current_model in models_to_try:
            for attempt in range(2):
                try:
                    attempts.append(f"{client_label} / {current_model}: attempt {attempt + 1}")

                    response = client.models.generate_content(
                        model=current_model,
                        contents=prompt,
                    )

                    text = clean_model_text(response.text or "")

                    if not text:
                        raise RuntimeError("Empty response returned by provider.")

                    return text, f"{current_model} ({client_label})", attempts

                except Exception as exc:
                    last_error = exc
                    error_text = str(exc).lower()
                    attempts.append(f"FAILED: {client_label} / {current_model}")

                    if any(code in error_text for code in ["503", "unavailable", "high demand", "overloaded", "500"]):
                        time.sleep(2 + attempt * 3)
                        continue

                    if any(code in error_text for code in ["429", "quota", "rate", "resource_exhausted"]):
                        break

                    if any(code in error_text for code in ["404", "not found", "invalid", "model"]):
                        break

                    break

    raise RuntimeError(
        "All configured API keys and models failed. Last error: "
        + str(last_error)
        + "\n\nLast attempts:\n"
        + "\n".join(attempts[-20:])
    )

def detect_cv_type(result: str) -> str:
    lower = (result or "").lower()
    if "selected cv type" in lower:
        selected_block = lower.split("selected cv type", 1)[1][:200]
        if "director cv" in selected_block:
            return "Director CV"
        if "manager cv" in selected_block:
            return "Manager CV"
    if "director cv" in lower and "manager cv" not in lower:
        return "Director CV"
    if "manager cv" in lower and "director cv" not in lower:
        return "Manager CV"
    return "Unknown"


ROLE_NAMES = [
    "Sr. EMEA Transport Manager",
    "Director Logistics",
    "Trade Compliance Manager EMEA",
    "Head of Logistics EMEA",
    "Sr. EMEA Logistics & Spare Parts Inventory Manager",
]

EXPECTED_COUNTS = {
    "Director CV": {
        "Sr. EMEA Transport Manager": 7,
        "Director Logistics": 8,
        "Trade Compliance Manager EMEA": 0,
        "Head of Logistics EMEA": 7,
        "Sr. EMEA Logistics & Spare Parts Inventory Manager": 5,
    },
    "Manager CV": {
        "Sr. EMEA Transport Manager": 7,
        "Director Logistics": 0,
        "Trade Compliance Manager EMEA": 8,
        "Head of Logistics EMEA": 7,
        "Sr. EMEA Logistics & Spare Parts Inventory Manager": 5,
    },
}


def normalize_role_name(line: str) -> str:
    return re.sub(r"^[#\-\*\s\d\.\)]+", "", line or "").replace(":", "").strip()


def extract_role_achievements(result: str) -> Dict[str, List[str]]:
    role_hits = {role: [] for role in ROLE_NAMES}
    current_role = None

    for raw_line in (result or "").splitlines():
        line = raw_line.strip()
        normalized = normalize_role_name(line)

        for role in ROLE_NAMES:
            if normalized.lower() == role.lower():
                current_role = role
                break

        if not current_role:
            continue

        if line.upper().startswith("ACHIEVEMENT:"):
            achievement = line.split(":", 1)[1].strip()
            role_hits[current_role].append(achievement)

    return role_hits


def extract_cover_letter_text(result: str) -> str:
    text = result or ""
    lower = text.lower()

    if "cover letter" not in lower:
        return ""

    start = lower.find("cover letter")
    after = text[start:]
    lines = after.splitlines()

    if lines:
        lines = lines[1:]

    stop_markers = ["FINAL ATS REPORT", "ATS REPORT", "FINAL REPORT"]
    collected = []

    for line in lines:
        stripped = line.strip()
        if stripped.upper() in stop_markers:
            break
        collected.append(line)

    return "\n".join(collected).strip()


def validate_generated_output(result: str, output_scope: str) -> Tuple[bool, List[str]]:
    issues = []
    cv_type = detect_cv_type(result)

    if cv_type not in EXPECTED_COUNTS:
        issues.append("Selected CV Type is missing or unclear. Must be exactly Manager CV or Director CV.")
        return False, issues

    role_achievements = extract_role_achievements(result)
    expected = EXPECTED_COUNTS[cv_type]

    for role, required_count in expected.items():
        actual_count = len(role_achievements.get(role, []))

        if required_count == 0 and actual_count > 0:
            issues.append(f"{role}: must not be included for {cv_type}, but {actual_count} achievement lines were found.")

        if required_count > 0 and actual_count != required_count:
            issues.append(f"{role}: expected exactly {required_count} achievement lines, found {actual_count}.")

        for index, achievement in enumerate(role_achievements.get(role, []), start=1):
            char_count = len(achievement)
            if char_count < 170 or char_count > 190:
                issues.append(f"{role} achievement {index}: {char_count} characters. Required 170-190 including spaces.")

    if "Cover Letter" in output_scope or "Cover Letter" in (result or ""):
        cover_letter = extract_cover_letter_text(result)

        if cover_letter:
            cover_len = len(cover_letter)
            if cover_len < 1790 or cover_len > 1810:
                issues.append(f"Cover Letter length is {cover_len} characters. Required 1,790-1,810, target 1,800.")
        elif "Cover Letter" in output_scope:
            issues.append("Cover Letter was requested but not found.")

    lower = (result or "").lower()
    if any(site in lower for site in ["linkedin ats", "jobscan", "skillsyncer", "resume worded", "rezi"]):
        if not any(term in lower for term in ["simulation", "simulated", "ats-style", "internal"]):
            issues.append("ATS validation mentions external ATS sites but does not clearly state that scores are simulated internally.")

    return len(issues) == 0, issues


def build_repair_prompt(original_prompt: str, previous_result: str, issues: List[str], output_scope: str) -> str:
    issue_text = "\n".join(f"- {issue}" for issue in issues)

    return f"""
You must repair the previous output. Do not explain the repair.

The previous output failed strict validation.

Validation issues:
{issue_text}

Hard requirements:
- Use the same source facts only.
- Keep the same selected CV type unless the previous selection was clearly wrong.
- Use exact role headings.
- Every achievement line must start with "ACHIEVEMENT: ".
- Count characters after the "ACHIEVEMENT: " prefix only.
- Every achievement line must be 170-190 characters including spaces.
- Manager CV: Trade Compliance Manager EMEA must have exactly 8 achievement lines.
- Director CV: Trade Compliance Manager EMEA must not be included.
- Cover Letter, if included, must be 1,790-1,810 characters including spaces, target 1,800.
- ATS validation must be described as simulated/internal, not live website data.
- Return only the final corrected output.

Original task:
{original_prompt}

Previous output:
{previous_result}

Output package:
{output_scope}
"""

def build_context(instructions, master, manager, director, profile: str):
    limits = CONTEXT_PROFILES.get(profile, CONTEXT_PROFILES["Balanced"])
    return f"""
[OPERATING INSTRUCTIONS]
{truncate_text(instructions, limits['instructions'])}

[EXPERIENCE REPOSITORY - SOURCE OF TRUTH]
{truncate_text(master, limits['master'])}

[MANAGER CV TEMPLATE]
{truncate_text(manager, limits['manager'])}

[DIRECTOR CV TEMPLATE]
{truncate_text(director, limits['director'])}
"""


def render_downloads(result: str, base_name: str):
    st.markdown('<div class="section-title">Export</div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    with c1:
        st.download_button("TXT", result.encode("utf-8"), f"{base_name}.txt", "text/plain", use_container_width=True)
    with c2:
        st.download_button("DOCX", make_docx(result), f"{base_name}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
    with c3:
        st.download_button("PDF", make_pdf(result), f"{base_name}.pdf", "application/pdf", use_container_width=True)


def strict_cv_prompt(context: str, job_description: str, language: str, output_scope: str) -> str:
    return f"""
You are CareerOps Studio, an executive CV optimization engine. Do not describe yourself as AI in the output.

PRIMARY OBJECTIVE
Create a tailored CV from the Job Description using the Experience Repository as the only source of truth. The final CV text must be clean, usable, ATS-aligned and faithful to the candidate's real achievements.

ABSOLUTE RULES
1. Never invent achievements, metrics, companies, roles, dates, systems, certifications, countries, savings, headcount, budget or KPIs.
2. Select the best achievements from the Experience Repository based on the Job Description.
3. Keep all facts and metrics faithful to the Experience Repository.
4. Choose Manager CV or Director CV using the instructions and the Job Description.
5. Follow mandatory role inclusion/exclusion and exact achievement counts.
6. Each achievement line must start with "ACHIEVEMENT: ".
7. Count characters after the "ACHIEVEMENT: " prefix only.
8. Every achievement line must be 170-190 characters including spaces.
9. Manager CV must include exactly 8 Trade Compliance Manager EMEA achievements, not 7.
10. Additional ATS Skills must contain exactly 10 skills extracted from the Job Description and supported by candidate experience.
11. ATS validation is simulated internally; do not claim it comes from live external ATS websites.
12. Cover Letter must be 1,790-1,810 characters including spaces, target exactly 1,800, same language as the job description unless the user selected another language.
13. All analysis, scoring, ranking and validation must be done internally. Do not expose internal reasoning in the final output.

MANDATORY BULLET COUNTS
{MANDATORY_BULLET_COUNTS}

FINAL OUTPUT POLICY
{FINAL_OUTPUT_POLICY}

COVER LETTER RULES
{COVER_LETTER_EXTRACTION_RULES}

ATS VALIDATION RULES
{ATS_VALIDATION_RULES}

STRICT OUTPUT FORMAT
{STRICT_OUTPUT_FORMAT_RULES}

MANDATORY PROCESS - INTERNAL ONLY
A. Analyze the Job Description internally: company name, exact job title, seniority, location, industry, hard skills, soft skills, systems, certifications, leadership requirements and critical ATS keywords.
B. Extract ATS skills from the Job Description and use them to evaluate CV alignment.
C. Select CV Type internally: Manager CV or Director CV.
D. Apply the mandatory bullet count table exactly.
E. Score and select achievements internally by keyword match, responsibility match, industry match, leadership match and systems/tools match.
F. Build the CV using selected template and selected achievements only.
G. Build the Cover Letter only when requested, using extracted company name, exact job title, ATS keywords and selected achievements; keep it 1,790-1,810 characters.
H. Run the optimization loop internally and validate every achievement character count.
I. Return only the concise final output requested by FINAL_OUTPUT_POLICY.

OUTPUT REQUIRED
{output_scope}

OUTPUT LANGUAGE
{language}

CONTEXT FILES
{context}

JOB DESCRIPTION
{job_description}
"""

def ats_validation_prompt(context: str, cv_text: str, job_description: str, language: str) -> str:
    return f"""
You are CareerOps Studio. Validate the CV against the Job Description and the operating instructions.

Check:
- Manager vs Director selection
- Hard skills match
- Soft skills match
- Job title match
- Industry match
- Keyword gaps
- Whether the CV uses only Experience Repository facts
- Whether Professional Experience achievement lines are 170-190 characters
- Whether Additional ATS Skills has exactly 10 skills
- Realistic simulated ATS score across LinkedIn ATS-style, Jobscan-style, SkillSyncer-style, Resume Worded-style and Rezi-style checks

Do not claim 100% unless every critical keyword is covered. Do not claim validation comes from live external websites.
Return practical corrections if needed.
Language: {language}

CONTEXT FILES
{context}

JOB DESCRIPTION
{job_description}

CV TO VALIDATE
{cv_text}
"""


def show_header(provider_status: str, context_profile: str, selected_model: str):
    st.markdown(
        f"""
        <div class="app-header">
            <div class="app-eyebrow">{APP_TITLE}</div>
            <h1 class="app-title">Executive resume optimization workspace</h1>
            <div class="app-subtitle">
                Generate tailored CVs, validate ATS alignment, select the correct Manager or Director route,
                and export client-ready documents from verified source material.
            </div>
        </div>

        <div class="status-row">
            <div class="status-card">
                <div class="status-label">Provider</div>
                <div class="status-value">{provider_status}</div>
            </div>
            <div class="status-card">
                <div class="status-label">Primary model</div>
                <div class="status-value">{selected_model}</div>
            </div>
            <div class="status-card">
                <div class="status-label">Context mode</div>
                <div class="status-value">{context_profile}</div>
            </div>
            <div class="status-card">
                <div class="status-label">Output</div>
                <div class="status-value">PDF · DOCX · TXT</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

st.set_page_config(page_title=APP_TITLE, page_icon="◆", layout="wide")
inject_css()

with st.sidebar:
    st.markdown("### Settings")
    primary_model = st.selectbox("Primary model", MODEL_OPTIONS, index=0)
    failover_models = st.multiselect(
        "Failover models",
        MODEL_OPTIONS,
        default=["gemini-2.5-flash", "gemini-2.0-flash-lite", "gemini-1.5-flash"],
        help="If the primary model fails, CareerOps Studio will automatically try the next selected model.",
    )
    custom_model = st.text_input("Optional custom model", value="")
    if custom_model.strip():
        failover_models = [custom_model.strip()] + failover_models
    context_profile = st.selectbox("Context depth", ["Fast", "Balanced", "Deep"], index=1)
    st.caption("Use Fast when limits are strict. Add multiple keys in .env as GEMINI_API_KEY_1, GEMINI_API_KEY_2, etc.")
    st.divider()
    render_history_panel()
    st.divider()
    st.markdown("### Source files")
    instructions_file = st.file_uploader("Instructions", type=["txt", "md", "docx", "pdf"], key="instructions")
    master_file = st.file_uploader("Experience Repository", type=["txt", "md", "docx", "pdf"], key="master")
    manager_file = st.file_uploader("Manager CV Template", type=["txt", "md", "docx", "pdf"], key="manager")
    director_file = st.file_uploader("Director CV Template", type=["txt", "md", "docx", "pdf"], key="director")

instructions_text = read_uploaded_file(instructions_file) or load_first_available([
    "My_ideas/AI_CV_Instructions_Master_Rev7.docx", "AI_CV_Instructions_Master.docx", "AI_CV_Instructions_Master_Rev7.docx"
])
master_text = read_uploaded_file(master_file) or load_first_available([
    "Master Experience File.docx", "My_ideas/Master_Experience_File_Structured_Model.docx", "Master_Experience_File_Structured_Model.docx"
])
manager_text = read_uploaded_file(manager_file) or load_first_available([
    "cv manager x AI.docx", "My_ideas/Manager_CV_Template.docx", "Manager_CV_Template.docx"
])
director_text = read_uploaded_file(director_file) or load_first_available([
    "cv Director x AI.docx", "My_ideas/Director_CV_Template.docx", "Director_CV_Template.docx"
])
context = build_context(instructions_text, master_text, manager_text, director_text, context_profile)

provider_status = "Failover enabled" if failover_models else "Single model"
show_header(provider_status, context_profile, primary_model)


if st.session_state.get("loaded_history_item"):
    loaded = st.session_state["loaded_history_item"]
    st.info(f"Loaded from history: {loaded.get('created_at', '')} · {loaded.get('kind', '')} · {loaded.get('title', '')}")
    with st.expander("Loaded history result"):
        st.markdown(loaded.get("result", ""))
        render_downloads(
            loaded.get("result", ""),
            f"careerops_history_{loaded.get('id', datetime.now().strftime('%Y%m%d_%H%M'))}",
        )

main_tab, validate_tab, files_tab = st.tabs(["Generate CV", "Validate ATS", "Source Files"])

with main_tab:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Generate tailored CV</div>', unsafe_allow_html=True)
    st.markdown('<p class="section-note">Paste one complete job description. The system selects Manager CV or Director CV, enforces exact role counts, validates every achievement at 170-190 characters including spaces, extracts ATS skills from the job description, and maps only verified achievements.</p>', unsafe_allow_html=True)
    job_description = st.text_area("Job description", height=360, placeholder="Paste the full job description here...")
    c1, c2 = st.columns([1, 1])
    with c1:
        language = st.selectbox("Output language", ["Same as job description", "English", "Português do Brasil"])
    with c2:
        output_scope = st.selectbox(
            "Output package",
            [
                "ATS Validation Status, Tailored CV, Cover Letter (strict 1800 chars), Final ATS Report",
                "Tailored CV only with ATS Validation Status",
                "ATS Analysis and achievement recommendations only",
            ],
        )
    generate = st.button("Generate CV", type="primary", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    if generate:
        if not job_description.strip():
            st.warning("Paste the job description first.")
        elif len(master_text.strip()) < 500:
            st.error("The Experience Repository is missing or too small. Upload the real Master Experience File before generating a CV.")
        else:
            prompt = strict_cv_prompt(context, job_description, language, output_scope)
            with st.spinner("Building the tailored CV and running strict validation..."):
                try:
                    result, used_model, attempts = call_model(prompt, primary_model, failover_models)

                    is_valid, issues = validate_generated_output(result, output_scope)

                    repair_round = 0
                    while not is_valid and repair_round < 2:
                        repair_round += 1
                        repair_prompt = build_repair_prompt(prompt, result, issues, output_scope)
                        repaired_result, repair_model, repair_attempts = call_model(repair_prompt, primary_model, failover_models)
                        attempts.extend(repair_attempts)
                        used_model = repair_model
                        result = repaired_result
                        is_valid, issues = validate_generated_output(result, output_scope)

                    if not is_valid:
                        st.warning("Strict validation still found issues. Review the validation list below.")
                        with st.expander("Strict validation issues"):
                            for issue in issues:
                                st.write("- " + issue)

                except Exception as exc:
                    st.error(str(exc))
                    st.stop()
            st.session_state["last_result"] = result
            st.session_state["last_job"] = job_description
            add_history_item(
                kind="CV",
                title=get_history_title(job_description, "Tailored CV"),
                job_description=job_description,
                result=result,
                model=used_model,
                context_profile=context_profile,
            )
            st.markdown(f'<div class="audit-box">Processed with: <b>{used_model}</b><br>Attempts: {" → ".join(attempts)}</div>', unsafe_allow_html=True)
            st.markdown('<div class="result-shell">', unsafe_allow_html=True)
            st.markdown(result)
            st.markdown('</div>', unsafe_allow_html=True)
            render_downloads(result, f"careerops_cv_{datetime.now().strftime('%Y%m%d_%H%M')}")

with validate_tab:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Validate existing CV against a job</div>', unsafe_allow_html=True)
    st.markdown('<p class="section-note">Use this when the client already has a CV and wants ATS match, keyword gaps and role alignment checked.</p>', unsafe_allow_html=True)
    existing_cv = st.text_area("Existing CV text", height=260, value=st.session_state.get("last_result", ""))
    validation_job = st.text_area("Job description for validation", height=240, value=st.session_state.get("last_job", ""))
    validate = st.button("Validate CV", type="primary", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)
    if validate:
        if not existing_cv.strip() or not validation_job.strip():
            st.warning("Paste both the CV and the job description.")
        else:
            prompt = ats_validation_prompt(context, existing_cv, validation_job, "Same as job description")
            with st.spinner("Validating ATS fit and testing failover if needed..."):
                try:
                    result, used_model, attempts = call_model(prompt, primary_model, failover_models)
                except Exception as exc:
                    st.error(str(exc))
                    st.stop()
            add_history_item(
                kind="Validation",
                title=get_history_title(validation_job, "ATS Validation"),
                job_description=validation_job,
                result=result,
                model=used_model,
                context_profile=context_profile,
            )
            st.markdown(f'<div class="audit-box">Processed with: <b>{used_model}</b><br>Attempts: {" → ".join(attempts)}</div>', unsafe_allow_html=True)
            st.markdown('<div class="result-shell">', unsafe_allow_html=True)
            st.markdown(result)
            st.markdown('</div>', unsafe_allow_html=True)
            render_downloads(result, f"careerops_validation_{datetime.now().strftime('%Y%m%d_%H%M')}")

with files_tab:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Source file status</div>', unsafe_allow_html=True)
    st.markdown('<p class="section-note">Confirm the source material was loaded correctly. The Experience Repository remains the source of truth.</p>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Instructions", f"{len(instructions_text):,}".replace(",", "."))
    c2.metric("Experience Repository", f"{len(master_text):,}".replace(",", "."))
    c3.metric("Manager Template", f"{len(manager_text):,}".replace(",", "."))
    c4.metric("Director Template", f"{len(director_text):,}".replace(",", "."))
    st.markdown('</div>', unsafe_allow_html=True)
    with st.expander("Preview instructions"):
        st.text(truncate_text(instructions_text, 6000))
    with st.expander("Preview Experience Repository"):
        st.text(truncate_text(master_text, 6000))

st.markdown('<div class="footer-note">CareerOps Studio - Private executive resume optimization workspace</div>', unsafe_allow_html=True)
