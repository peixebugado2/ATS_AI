import os
import json
import re
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

MANDATORY_BULLET_COUNTS = """
MANDATORY BULLET COUNTS - NON NEGOTIABLE

Director CV:
- Sr. EMEA Transport Manager: exactly 7 achievement lines
- Director Logistics: exactly 8 achievement lines
- Trade Compliance Manager EMEA: must not appear
- Head of Logistics EMEA: exactly 7 achievement lines
- Sr. EMEA Logistics & Spare Parts Inventory Manager: exactly 5 achievement lines

Manager CV:
- Sr. EMEA Transport Manager: exactly 7 achievement lines
- Director Logistics: must not appear
- Trade Compliance Manager EMEA: exactly 8 achievement lines, not 7
- Head of Logistics EMEA: exactly 7 achievement lines
- Sr. EMEA Logistics & Spare Parts Inventory Manager: exactly 5 achievement lines

Do not add extra roles. Do not omit required roles. Do not create fewer or more achievement lines.
"""

STRICT_FORMAT_RULES = """
STRICT FORMAT RULES - MACHINE VALIDATED

The final output must use these exact headings:
SELECTED CV TYPE
ATS VALIDATION STATUS
FINAL TAILORED CV
COVER LETTER
FINAL ATS REPORT

Professional experience role headings must match exactly:
Sr. EMEA Transport Manager
Director Logistics
Trade Compliance Manager EMEA
Head of Logistics EMEA
Sr. EMEA Logistics & Spare Parts Inventory Manager

Every achievement line must start exactly with:
ACHIEVEMENT: 

The character count rule applies only to the text after ACHIEVEMENT: .
Every achievement text after ACHIEVEMENT: must be 170 to 190 characters including spaces.
Do not use bullet symbols for achievement lines.
"""

COVER_LETTER_RULES = """
COVER LETTER RULES - MACHINE VALIDATED

The cover letter must be generated from the Job Description and the selected verified achievements.
Extract the company name, exact job title, location, seniority, department, responsibilities, systems, hard skills, soft skills and ATS keywords from the Job Description whenever available.
Never use placeholders such as [Company Name], [Hiring Manager Name], [Job Title] or similar text.
If company name is available, address it to: Dear Hiring Team at [Company Name],
If company name is not available, address it to: Dear Hiring Manager,
The cover letter body must be 1,790 to 1,810 characters including spaces. Target exactly 1,800 characters.
Do not show the character count in the final output.
Do not invent facts, achievements, metrics, certifications, systems or company information.
"""

ATS_RULES = """
ATS VALIDATION RULES - NO LIVE SITE CLAIMS

The ATS validation is an internal simulation based only on the Job Description, generated CV and source files.
It is not connected to LinkedIn, Jobscan, SkillSyncer, Resume Worded, Rezi or any external ATS website.
Never claim that the score came directly from live websites.
Use wording like: simulated ATS-style score, internal ATS simulation, LinkedIn-style, Jobscan-style, SkillSyncer-style.

Skill extraction requirements:
- Extract hard skills directly from the Job Description.
- Extract soft skills directly from the Job Description.
- Extract tools, systems, certifications, industry terms, leadership requirements and operational keywords directly from the Job Description.
- Match extracted skills against the tailored CV.
- Additional ATS Skills must contain exactly 10 skills and must be aligned with the Job Description and candidate experience.
- Do not invent skills unsupported by the Job Description or Experience Repository.
"""

FINAL_OUTPUT_POLICY = """
FINAL OUTPUT POLICY

Do all analysis, scoring, ranking and validation internally.
Do not show reasoning, scoring steps, internal ranking, draft alternatives, keyword dumps or validation logs.

Final answer must contain only:
1. SELECTED CV TYPE
2. ATS VALIDATION STATUS - concise table with key simulated statistics
3. FINAL TAILORED CV
4. COVER LETTER only when requested
5. FINAL ATS REPORT - concise realistic simulated scores
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
            --bg: #070910;
            --panel: rgba(16, 19, 27, .86);
            --panel2: rgba(22, 27, 38, .84);
            --panel3: rgba(255,255,255,.045);
            --border: rgba(230, 210, 170, .16);
            --border2: rgba(230, 210, 170, .28);
            --text: #f7f1e8;
            --muted: rgba(247,241,232,.66);
            --gold: #d6b56e;
            --gold2: #f2d99e;
            --blue: #8fb4dc;
            --green: #94d6a8;
            --red: #ff7272;
        }
        .stApp {
            color: var(--text);
            background:
              radial-gradient(circle at top left, rgba(214,181,110,.16), transparent 30%),
              radial-gradient(circle at 90% 12%, rgba(143,180,220,.11), transparent 28%),
              linear-gradient(135deg, #05060a 0%, #0b1019 45%, #080a0f 100%);
        }
        [data-testid="stHeader"] { background: rgba(0,0,0,0); }
        .block-container { max-width: 1240px; padding-top: 1.2rem; }
        [data-testid="stSidebar"] {
            background: linear-gradient(180deg, rgba(9,12,18,.99), rgba(13,17,25,.94));
            border-right: 1px solid var(--border);
        }
        [data-testid="stSidebar"] * { color: var(--text); }
        .studio-shell {
            border: 1px solid var(--border);
            background: linear-gradient(135deg, rgba(255,255,255,.06), rgba(255,255,255,.025));
            border-radius: 28px;
            padding: 22px 26px;
            box-shadow: 0 26px 80px rgba(0,0,0,.34);
            margin-bottom: 18px;
        }
        .studio-top { display:flex; justify-content:space-between; align-items:flex-start; gap:18px; }
        .studio-brand { color:var(--gold); letter-spacing:.24em; text-transform:uppercase; font-size:.75rem; font-weight:700; }
        .studio-title { margin-top:8px; font-size:2.3rem; line-height:1.02; letter-spacing:-.055em; font-weight:850; }
        .studio-subtitle { color:var(--muted); margin-top:10px; max-width:760px; font-size:.98rem; }
        .status-pill { border:1px solid var(--border2); background:rgba(214,181,110,.12); border-radius:999px; padding:8px 12px; color:var(--gold2); font-size:.78rem; white-space:nowrap; }
        .mini-grid { display:grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap:12px; margin: 18px 0 2px; }
        .mini-card { border:1px solid var(--border); background:rgba(255,255,255,.035); border-radius:18px; padding:14px 16px; }
        .mini-label { color:var(--muted); font-size:.72rem; letter-spacing:.13em; text-transform:uppercase; }
        .mini-value { margin-top:6px; font-size:1.03rem; font-weight:780; }
        .card {
            border:1px solid var(--border); background:var(--panel); border-radius:24px;
            padding:24px 26px; box-shadow: 0 20px 60px rgba(0,0,0,.24); margin-bottom:16px;
        }
        .section-title { font-size:1.35rem; font-weight:820; letter-spacing:-.035em; margin:0 0 6px; }
        .section-note { color:var(--muted); margin:0 0 18px; }
        .result-shell { border:1px solid var(--border); background:rgba(255,255,255,.055); border-radius:22px; padding:24px; margin-top:18px; }
        .audit-box { border:1px solid rgba(148,214,168,.25); background:rgba(148,214,168,.08); border-radius:16px; padding:12px 14px; margin:14px 0; color:rgba(247,241,232,.88); }
        .footer-note { text-align:center; color:var(--muted); font-size:.80rem; padding:26px 0 10px; }
        .stButton > button, .stDownloadButton > button {
            border-radius: 14px !important; border:1px solid var(--border) !important; font-weight:750 !important;
            padding:.72rem 1rem !important;
        }
        .stButton > button[kind="primary"] {
            color:#080910 !important; border:0 !important;
            background: linear-gradient(90deg, #f3d995, #caa760) !important;
        }
        textarea, input, .stTextArea textarea, .stTextInput input { border-radius: 16px !important; }
        [data-testid="stFileUploader"] { border:1px dashed rgba(214,181,110,.28); background:rgba(255,255,255,.035); border-radius:18px; padding:12px; }
        div[data-testid="stTabs"] button { border-radius:999px !important; padding:11px 18px !important; }
        hr { border-color: var(--border); }
        @media (max-width: 900px) { .mini-grid { grid-template-columns: repeat(2, minmax(0,1fr)); } .studio-title { font-size:1.8rem; } }
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

def normalize_role_name(line: str) -> str:
    return re.sub(r"^[#\-\*\s\d\.\)]+", "", line or "").replace(":", "").strip()


def detect_cv_type(result: str) -> str:
    text = result or ""
    lower = text.lower()
    m = re.search(r"selected\s+cv\s+type\s*[:\n\-]*\s*(manager cv|director cv)", lower)
    if m:
        return "Manager CV" if "manager" in m.group(1) else "Director CV"
    if "director cv" in lower and "manager cv" not in lower:
        return "Director CV"
    if "manager cv" in lower and "director cv" not in lower:
        return "Manager CV"
    return "Unknown"


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
    match = re.search(r"(?is)\bCOVER LETTER\b\s*\n(.*?)(?:\n\s*FINAL ATS REPORT\b|\Z)", text)
    if not match:
        return ""
    cover = match.group(1).strip()
    return re.sub(r"\n{3,}", "\n\n", cover)


def has_requested_cover_letter(output_scope: str) -> bool:
    return "cover letter" in (output_scope or "").lower()


def validate_generated_output(result: str, output_scope: str) -> Tuple[bool, List[str]]:
    issues = []
    cv_type = detect_cv_type(result)
    if cv_type not in EXPECTED_COUNTS:
        issues.append("SELECTED CV TYPE missing or unclear. Must be exactly Manager CV or Director CV.")
        return False, issues

    role_achievements = extract_role_achievements(result)
    expected = EXPECTED_COUNTS[cv_type]
    for role, required_count in expected.items():
        actual_count = len(role_achievements.get(role, []))
        if required_count == 0 and actual_count > 0:
            issues.append(f"{role}: must not appear in {cv_type}; found {actual_count} achievement lines.")
        elif required_count > 0 and actual_count != required_count:
            issues.append(f"{role}: expected exactly {required_count} achievement lines; found {actual_count}.")
        for idx, achievement in enumerate(role_achievements.get(role, []), start=1):
            char_count = len(achievement)
            if char_count < 170 or char_count > 190:
                issues.append(f"{role} achievement {idx}: {char_count} characters. Required 170-190 including spaces.")

    if has_requested_cover_letter(output_scope):
        cover = extract_cover_letter_text(result)
        if not cover:
            issues.append("Cover Letter requested but COVER LETTER section was not found.")
        else:
            cover_len = len(cover)
            if cover_len < 1790 or cover_len > 1810:
                issues.append(f"Cover Letter length is {cover_len} characters. Required 1,790-1,810; target 1,800.")
            if re.search(r"\[[^\]]+\]", cover):
                issues.append("Cover Letter contains placeholder text in square brackets.")

    lower = (result or "").lower()
    external_terms = ["linkedin ats", "jobscan", "skillsyncer", "resume worded", "rezi"]
    if any(term in lower for term in external_terms):
        if not any(term in lower for term in ["simulated", "simulation", "internal", "ats-style", "style"]):
            issues.append("ATS validation references external tools but does not state that scores are internal simulations.")
    return len(issues) == 0, issues


def build_repair_prompt(original_prompt: str, previous_result: str, issues: List[str], output_scope: str) -> str:
    issue_text = "\n".join(f"- {issue}" for issue in issues)
    return f"""
Repair the previous output. Return only the corrected final output. Do not explain.

STRICT VALIDATION FAILURES:
{issue_text}

NON-NEGOTIABLE REPAIR RULES:
- Keep facts faithful to the Experience Repository.
- Keep the selected CV type unless it is missing or clearly wrong.
- Use exact role headings from STRICT FORMAT RULES.
- Every achievement line must start exactly with ACHIEVEMENT: .
- Count only the text after ACHIEVEMENT: . It must be 170-190 characters including spaces.
- Manager CV must have exactly 8 Trade Compliance Manager EMEA achievements. Not 7.
- Director CV must not include Trade Compliance Manager EMEA.
- Cover Letter, if requested, must be 1,790-1,810 characters including spaces, target 1,800.
- ATS must be described as internal simulated ATS-style validation, not live website validation.
- Additional ATS Skills must be extracted from the Job Description and supported by source facts.

ORIGINAL TASK:
{original_prompt}

PREVIOUS OUTPUT:
{previous_result}

OUTPUT PACKAGE:
{output_scope}
"""


def generate_with_strict_validation(prompt: str, output_scope: str, primary_model: str, failover_models: List[str]) -> Tuple[str, str, List[str], List[str]]:
    result, used_model, attempts = call_model(prompt, primary_model, failover_models)
    is_valid, issues = validate_generated_output(result, output_scope)
    repair_round = 0
    while not is_valid and repair_round < 3:
        repair_round += 1
        repair_prompt = build_repair_prompt(prompt, result, issues, output_scope)
        repaired, repair_model, repair_attempts = call_model(repair_prompt, primary_model, failover_models)
        attempts.extend(repair_attempts)
        result = repaired
        used_model = repair_model
        is_valid, issues = validate_generated_output(result, output_scope)
    return result, used_model, attempts, issues

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
Create a tailored CV from the Job Description using the Experience Repository as the only source of truth. The final output must satisfy strict machine validation.

ABSOLUTE RULES
1. Never invent achievements, metrics, companies, roles, dates, systems, certifications, countries, savings, headcount, budget or KPIs.
2. Select achievements only from the Experience Repository based on the Job Description.
3. Choose Manager CV or Director CV using the Job Description and instructions.
4. Follow exact role inclusion/exclusion and exact achievement counts.
5. Every achievement line must start exactly with ACHIEVEMENT: .
6. Count only the text after ACHIEVEMENT: . It must be 170-190 characters including spaces.
7. Manager CV must include exactly 8 Trade Compliance Manager EMEA achievement lines, not 7.
8. Additional ATS Skills must contain exactly 10 skills extracted from the Job Description and supported by candidate experience.
9. Cover Letter, when requested, must be 1,790-1,810 characters including spaces. Target exactly 1,800.
10. ATS validation is an internal simulation. Never claim it is coming from live external websites.
11. Do all analysis internally. Do not show reasoning, scoring steps, keyword dumps, validation logs, draft alternatives or chain-of-thought.

{MANDATORY_BULLET_COUNTS}

{STRICT_FORMAT_RULES}

{COVER_LETTER_RULES}

{ATS_RULES}

{FINAL_OUTPUT_POLICY}

MANDATORY INTERNAL PROCESS
A. Extract company name, exact job title, seniority, location, industry, hard skills, soft skills, systems, certifications, leadership requirements and critical ATS keywords from the Job Description.
B. Select Manager CV or Director CV.
C. Apply mandatory role counts exactly.
D. Select and rewrite achievements from source facts only.
E. Validate each achievement line internally until it is 170-190 characters including spaces.
F. If Cover Letter is requested, write it from Job Description + selected achievements and keep it 1,790-1,810 characters.
G. Produce internal simulated ATS-style statistics and skill matching.
H. Return only the final output sections.

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
You are CareerOps Studio. Validate the CV against the Job Description and operating instructions.

Validation must be an internal ATS-style simulation only. Do not claim live website access.

Check:
- Manager CV vs Director CV selection
- Exact mandatory role inclusion/exclusion
- Exact number of achievement lines per role
- Trade Compliance Manager EMEA must be exactly 8 achievements for Manager CV
- Every achievement line must be 170-190 characters including spaces after ACHIEVEMENT: prefix
- Cover Letter length when present: 1,790-1,810 characters including spaces
- Hard skills extracted from the Job Description
- Soft skills extracted from the Job Description
- Systems, tools, certifications and industry keywords extracted from the Job Description
- Whether Additional ATS Skills has exactly 10 skills aligned to the Job Description
- Whether the CV uses only Experience Repository facts
- Realistic simulated scores for LinkedIn-style, Jobscan-style, SkillSyncer-style, Resume Worded-style and Rezi-style checks

Never claim 100% unless every critical keyword is naturally covered.
Return concise corrections only.
Language: {language}

{MANDATORY_BULLET_COUNTS}

{ATS_RULES}

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
        <div class="studio-shell">
            <div class="studio-top">
                <div>
                    <div class="studio-brand">{APP_TITLE}</div>
                    <div class="studio-title">Executive CV optimization workspace.</div>
                    <div class="studio-subtitle">Map real achievements to job requirements, route Manager / Director CVs, validate ATS fit, and export client-ready files.</div>
                </div>
                <div class="status-pill">{provider_status}</div>
            </div>
            <div class="mini-grid">
                <div class="mini-card"><div class="mini-label">Primary model</div><div class="mini-value">{selected_model}</div></div>
                <div class="mini-card"><div class="mini-label">Context mode</div><div class="mini-value">{context_profile}</div></div>
                <div class="mini-card"><div class="mini-label">Routing</div><div class="mini-value">Manager / Director</div></div>
                <div class="mini-card"><div class="mini-label">Exports</div><div class="mini-value">PDF / DOCX / TXT</div></div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


st.set_page_config(page_title=APP_TITLE, page_icon="◆", layout="wide")
inject_css()

with st.sidebar:
    st.markdown("### Workspace")
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
    st.markdown("### Source documents")
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

main_tab, validate_tab, files_tab = st.tabs(["Application workspace", "ATS validation", "Source control"])

with main_tab:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Create tailored CV</div>', unsafe_allow_html=True)
    st.markdown('<p class="section-note">Paste one complete job description. The system enforces exact role counts, validates every achievement at 170-190 characters including spaces, keeps Trade Compliance Manager at 8 achievements for Manager CV, and treats ATS as internal simulation.</p>', unsafe_allow_html=True)
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
                "ATS Validation Status and Final ATS Report only",
            ],
        )
    generate = st.button("Generate client-ready CV", type="primary", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    if generate:
        if not job_description.strip():
            st.warning("Paste the job description first.")
        elif len(master_text.strip()) < 500:
            st.error("The Experience Repository is missing or too small. Upload the real Master Experience File before generating a CV.")
        else:
            prompt = strict_cv_prompt(context, job_description, language, output_scope)
            with st.spinner("Building CV and running strict validation..."):
                try:
                    result, used_model, attempts, strict_issues = generate_with_strict_validation(
                        prompt,
                        output_scope,
                        primary_model,
                        failover_models,
                    )
                    if strict_issues:
                        st.warning("Strict validation still found issues after repair attempts. Review below before sending to client.")
                        with st.expander("Strict validation issues"):
                            for issue in strict_issues:
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
    validate = st.button("Run ATS validation", type="primary", use_container_width=True)
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
    st.markdown('<div class="section-title">Loaded source files</div>', unsafe_allow_html=True)
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
