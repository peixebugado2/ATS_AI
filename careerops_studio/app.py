import os
import json
import re
import time
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Tuple, Optional

import streamlit as st
from dotenv import load_dotenv
from google import genai
from docx import Document
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_LEFT
from reportlab.lib import colors

try:
    from pypdf import PdfReader
except Exception:
    PdfReader = None

load_dotenv()

# --- CONFIGURAÇÕES E CONSTANTES ORIGINAIS ---
APP_TITLE = "CareerOps Studio"
APP_SUBTITLE = "Executive Resume Optimization Platform"
DEFAULT_MODEL = "gemini-2.5-flash-lite"
ROOT = Path(__file__).parent
TEMPLATE_DIR = ROOT / "templates"
HISTORY_FILE = ROOT / "careerops_chat_history.json"
MAX_HISTORY_ITEMS = 50

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

REQUIRED_CV_SECTIONS = [
    "Name",
    "Contact Details",
    "Additional ATS Skills",
    "Professional Summary",
    "Professional Experience",
    "Education",
    "Systems / Tools",
    "ATS Analysis",
]

# --- REGRAS DE NEGÓCIO E PROMPTS ORIGINAIS ---
MANDATORY_CV_TYPE_RULES = """
CV TYPE SELECTION RULES - NON NEGOTIABLE

Director CV:
Use when the Job Description mentions Director, Head of, Senior Leadership, Strategic Leadership, Global Leadership, Regional Leadership, executive ownership or director-level accountability.
Mandatory sections: Sr. EMEA Transport Manager; Director Logistics; Head of Logistics EMEA; Sr. EMEA Logistics & Spare Parts Inventory Manager.
Director CV must NOT include Trade Compliance Manager EMEA.

Manager CV:
Use when the Job Description mentions Manager, Operations Manager, Supply Chain Manager, Logistics Manager, Warehouse Manager or operational management.
Mandatory sections: Sr. EMEA Transport Manager; Trade Compliance Manager EMEA; Head of Logistics EMEA; Sr. EMEA Logistics & Spare Parts Inventory Manager.
Manager CV must include Trade Compliance Manager EMEA instead of Director Logistics.
"""

MANDATORY_BULLET_COUNTS = """
MANDATORY ACHIEVEMENT COUNTS - NON NEGOTIABLE

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
- Head of Logistics EMEA: exactly 7 achievement lines, not 6
- Sr. EMEA Logistics & Spare Parts Inventory Manager: exactly 5 achievement lines

Do not add extra roles. Do not omit required roles. Do not create fewer or more achievement lines.
"""

STRICT_FORMAT_RULES = """
STRICT COPY-PASTE FORMAT RULES - MACHINE VALIDATED

The output must be compact, ATS optimized and easy to copy and paste into the existing CV template.

Mandatory section order inside FINAL TAILORED CV:
1. Name
2. Contact Details
3. Additional ATS Skills
4. Professional Summary
5. Professional Experience
6. Education
7. Systems / Tools
8. ATS Analysis
9. Cover Letter only if requested

Additional ATS Skills formatting:
- Exactly 10 skills.
- One skill per line.
- No bullets.
- No numbering.
- No blank lines between skills.
- No comma-separated skills.
- Compact vertical list only.

Professional Experience formatting:
- Each role heading must be on its own line.
- Each achievement must be on its own separate new line.
- DO NOT write "ACHIEVEMENT:".
- DO NOT write "Achievement:".
- DO NOT write "Bullet:".
- DO NOT write "Result:".
- DO NOT write "Success:".
- DO NOT write "KPI:".
- DO NOT use bullet symbols such as -, *, •, or numbered bullets for achievements.
- Do not place multiple achievements on the same paragraph.
- One blank line between roles only.
- No blank lines between achievements inside the same role.
- Each achievement must occupy exactly one line.
- Achievement text must be plain text only, ready to paste into the CV.

Professional experience role headings must match exactly:
Sr. EMEA Transport Manager
Director Logistics
Trade Compliance Manager EMEA
Head of Logistics EMEA
Sr. EMEA Logistics & Spare Parts Inventory Manager

Every achievement line must be 170 to 190 characters including spaces.
"""

COVER_LETTER_RULES = """
COVER LETTER RULES - MACHINE VALIDATED

The cover letter must be generated from the Job Description and the selected verified achievements.
Extract the company name, exact job title, location, seniority, department, responsibilities, systems, hard skills, soft skills and ATS keywords from the Job Description whenever available.
Never use placeholders such as [Company Name], [Hiring Manager Name], [Job Title] or similar text.
If company name is available, address it to: Dear Hiring Team at [Company Name],
If company name is not available, address it to: Dear Hiring Manager,
The cover letter body must be 1,795 to 1,805 characters including spaces. Target exactly 1,800 characters.
The Cover Letter must be written in the same language as the Job Description unless the user explicitly selected another output language.
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

Mandatory ATS gap analysis requirements:
- Always identify Matched Keywords.
- Always identify Partial Match Keywords with 60-80% estimated coverage.
- Always identify Missing Important Keywords.
- Always consider partial matches and missing keywords during CV creation.
- Always list Partial Match Keywords and Missing Important Keywords in the Final ATS Report.
- The objective is zero deviations.
"""

ADDITIONAL_ATS_SKILLS_RULES = """
ADDITIONAL ATS SKILLS RULES - NON NEGOTIABLE

Additional ATS Skills must always be displayed vertically:
- Exactly 10 skills.
- One skill per line.
- No bullets.
- No numbering.
- No commas separating multiple skills.
- No blank lines between skills.
- No extra explanation before or after the skill list.
- Skills must be extracted from the Job Description and supported by candidate experience.
"""

ATS_GAP_ANALYSIS_RULES = """
MANDATORY ATS GAP ANALYSIS - NON NEGOTIABLE

The final output must always include ATS gap analysis at the end, after the Cover Letter when Cover Letter is requested.

The final ATS section must include these Markdown tables:

1. Keyword Coverage Table
| Category | Keyword | Status |
|---|---|---|
| Matched | keyword | Covered |
| Partial Match | keyword | 60-80% |
| Missing Important | keyword | High Priority |

2. Partial Match Keywords Table
| Partial Match Keyword | Estimated Coverage | Action Taken |
|---|---|---|
| keyword | 60-80% | Improved naturally in CV when supported by source facts |

3. Missing Important Keywords Table
| Missing Important Keyword | Priority | Action |
|---|---|---|
| keyword | High | Address if supported by candidate experience |

Rules:
- Always include Partial Match Keywords.
- Always include Missing Important Keywords.
- Partial matches are keywords with estimated 60-80% coverage.
- Missing Important Keywords must be extracted from the Job Description.
- Never invent keywords.
- Every partial match and missing important keyword must be considered during CV creation.
- The objective is zero ATS deviations.
- If a keyword cannot be added because it is not supported by the Experience Repository, keep it listed as Missing Important.
- Use Markdown tables so DOCX and PDF exports render real table cells.
"""

LANGUAGE_ALIGNMENT_RULES = """
LANGUAGE ALIGNMENT RULES

- Detect the language of the Job Description.
- The Cover Letter must always be written in the same language as the Job Description.
- If the Job Description is English, the Cover Letter must be English.
- If the Job Description is Dutch, the Cover Letter must be Dutch.
- If the Job Description is German, the Cover Letter must be German.
- If the Job Description is French, the Cover Letter must be French.
- If the Job Description is Portuguese, the Cover Letter must be Portuguese.
- Never translate the Cover Letter into another language unless the user explicitly selected a different output language.
"""

FINAL_OUTPUT_POLICY = """
FINAL OUTPUT POLICY

Do all analysis, scoring, ranking and validation internally.
Do not show reasoning, scoring steps, internal ranking, draft alternatives, keyword dumps or validation logs.

Final answer must contain only:
1. SELECTED CV TYPE
2. ATS VALIDATION STATUS - concise Markdown table with columns Metric and Score
3. FINAL TAILORED CV
4. FINAL ATS REPORT - concise realistic simulated scores, using Markdown tables for metrics

If Cover Letter is requested, include it inside FINAL TAILORED CV after ATS Analysis.

At the end of FINAL ATS REPORT, always include:
- Keyword Coverage Table
- Partial Match Keywords Table
- Missing Important Keywords Table

All tables must use Markdown table syntax so DOCX and PDF exports render formatted cells.
"""

# --- LÓGICA DE PERSISTÊNCIA E UTILITÁRIOS ---

def load_history() -> List[Dict]:
    if not HISTORY_FILE.exists(): return []
    try:
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            return data if isinstance(data, list) else []
    except: return []

def save_history(history: List[Dict]):
    try:
        with open(HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(history[:MAX_HISTORY_ITEMS], f, ensure_ascii=False, indent=2)
    except: pass

def read_uploaded_file(uploaded_file) -> str:
    if uploaded_file is None: return ""
    uploaded_file.seek(0)
    name = uploaded_file.name.lower()
    data = uploaded_file.read()
    if name.endswith((".txt", ".md", ".csv")): return data.decode("utf-8", errors="ignore")
    if name.endswith(".docx"):
        doc = Document(BytesIO(data))
        parts = []
        for p in doc.paragraphs:
            if p.text.strip(): parts.append(p.text)
        for table in doc.tables:
            for row in table.rows:
                cells = [cell.text.strip() for cell in row.cells]
                if any(cells): parts.append(" | ".join(cells))
        return "\n".join(parts)
    if name.endswith(".pdf") and PdfReader:
        reader = PdfReader(BytesIO(data))
        return "\n".join([(page.extract_text() or "") for page in reader.pages])
    return data.decode("utf-8", errors="ignore")

def load_template_text(filename: str, fallback: str = "") -> str:
    path = TEMPLATE_DIR / filename
    if not path.exists(): return fallback
    try:
        if path.suffix.lower() == ".docx":
            doc = Document(path)
            parts = []
            for p in doc.paragraphs:
                if p.text.strip(): parts.append(p.text)
            for table in doc.tables:
                for row in table.rows:
                    cells = [cell.text.strip() for cell in row.cells]
                    if any(cells): parts.append(" | ".join(cells))
            return "\n".join(parts)
        return path.read_text(encoding="utf-8", errors="ignore")
    except: return fallback

def load_first_available(filenames: List[str]) -> str:
    for f in filenames:
        text = load_template_text(f)
        if text.strip(): return text
    return ""

def truncate_text(text: str, max_chars: int) -> str:
    text = text or ""
    if len(text) <= max_chars: return text
    return text[:max_chars] + "\n\n[Content truncated by application limit]"

# --- EXPORTADORES (PDF/DOCX) ---

def is_markdown_table_line(line: str) -> bool:
    stripped = (line or "").strip()
    return stripped.startswith("|") and stripped.endswith("|") and stripped.count("|") >= 2

def is_markdown_separator_line(line: str) -> bool:
    stripped = (line or "").strip()
    if not is_markdown_table_line(stripped): return False
    cells = [cell.strip() for cell in stripped.strip("|").split("|")]
    return bool(cells) and all(re.fullmatch(r":?-{3,}:?", cell or "") for cell in cells)

def parse_markdown_table(lines: List[str], start_index: int) -> Tuple[List[List[str]], int]:
    table_lines = []
    index = start_index
    while index < len(lines) and is_markdown_table_line(lines[index]):
        table_lines.append(lines[index])
        index += 1
    rows = []
    for raw in table_lines:
        if is_markdown_separator_line(raw): continue
        cells = [cell.strip() for cell in raw.strip().strip("|").split("|")]
        rows.append(cells)
    return rows, index

def detect_pipe_table_block(lines: List[str], start_index: int) -> bool:
    if start_index >= len(lines): return False
    if not is_markdown_table_line(lines[start_index]): return False
    if start_index + 1 < len(lines) and is_markdown_separator_line(lines[start_index + 1]): return True
    return False

def add_docx_markdown_table(doc: Document, rows: List[List[str]]) -> None:
    if not rows: return
    max_cols = max(len(row) for row in rows)
    normalized_rows = [row + [""] * (max_cols - len(row)) for row in rows]
    table = doc.add_table(rows=len(normalized_rows), cols=max_cols)
    table.style = "Table Grid"
    for row_index, row in enumerate(normalized_rows):
        for col_index, cell_text in enumerate(row):
            cell = table.cell(row_index, col_index)
            cell.text = cell_text
    doc.add_paragraph("")

def make_docx(content: str, title: str = "CareerOps Studio Output") -> bytes:
    doc = Document()
    doc.add_heading(title, level=1)
    lines = content.splitlines()
    index = 0
    while index < len(lines):
        line = lines[index].strip()
        if not line:
            doc.add_paragraph("")
            index += 1
            continue
        if detect_pipe_table_block(lines, index):
            rows, next_index = parse_markdown_table(lines, index)
            add_docx_markdown_table(doc, rows)
            index = next_index
            continue
        doc.add_paragraph(line)
        index += 1
    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

def make_pdf(content: str, title: str = "CareerOps Studio Output") -> bytes:
    bio = BytesIO()
    pdf = SimpleDocTemplate(bio, pagesize=A4, rightMargin=1.65*cm, leftMargin=1.65*cm, topMargin=1.55*cm, bottomMargin=1.55*cm)
    styles = getSampleStyleSheet()
    story = [Paragraph(title, styles["Title"])]
    for line in content.splitlines():
        if line.strip(): story.append(Paragraph(line.strip(), styles["Normal"]))
    pdf.build(story)
    return bio.getvalue()

# --- LÓGICA DE IA (GEMINI) ---

def get_clients():
    keys = [os.getenv(f"GEMINI_API_KEY_{i}") for i in range(1, 6)] + [os.getenv("GEMINI_API_KEY")]
    clients = []
    for i, k in enumerate(keys):
        if k: clients.append({"label": f"Key {i+1}", "client": genai.Client(api_key=k.strip())})
    return clients

def call_model(prompt: str, primary_model: str, failover_models: List[str]) -> Tuple[str, str, List[str]]:
    clients = get_clients()
    if not clients: raise RuntimeError("No Gemini API key found.")
    models = [primary_model] + failover_models
    attempts = []
    for client_info in clients:
        for m in models:
            try:
                attempts.append(f"{client_info['label']} / {m}")
                resp = client_info["client"].models.generate_content(model=m, contents=prompt)
                return resp.text, m, attempts
            except Exception as e:
                attempts.append(f"Fail: {str(e)[:50]}")
    raise RuntimeError("All models failed.")

# --- LÓGICA DE VALIDAÇÃO E PROMPTS ORIGINAIS ---

def normalize_role_name(line: str) -> str:
    cleaned = re.sub(r"^[#\-\*\•\s\d\.\)]+", "", line or "").strip()
    return cleaned.replace(":", "").strip()

def normalize_for_match(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", (text or "").lower()).strip()

def detect_cv_type(result: str) -> str:
    lower = (result or "").lower()
    if "director cv" in lower[:500]: return "Director CV"
    if "manager cv" in lower[:500]: return "Manager CV"
    return "Unknown"

def is_role_heading(line: str) -> str:
    normalized = normalize_role_name(line).lower()
    for role in ROLE_NAMES:
        if normalized == role.lower(): return role
    return ""

def is_plain_achievement_line(line: str) -> bool:
    stripped = (line or "").strip()
    if not stripped or len(stripped) < 80: return False
    if stripped.startswith(("-", "*", "•")) or re.match(r"^\d+[\.\)]\s+", stripped): return False
    return True

def extract_role_achievements(result: str) -> Dict[str, List[str]]:
    role_hits = {role: [] for role in ROLE_NAMES}
    current_role = None
    inside_experience = False
    for line in (result or "").splitlines():
        line = line.strip()
        if not line: continue
        norm = normalize_for_match(line)
        if norm == "professional experience":
            inside_experience = True
            current_role = None
            continue
        if inside_experience and norm in ["education", "systems tools", "ats analysis"]: break
        role = is_role_heading(line)
        if role:
            current_role = role
            continue
        if inside_experience and current_role and is_plain_achievement_line(line):
            role_hits[current_role].append(line)
    return role_hits

def validate_generated_output(result: str, output_scope: str) -> Tuple[bool, List[str]]:
    issues = []
    cv_type = detect_cv_type(result)
    if cv_type not in EXPECTED_COUNTS:
        issues.append("SELECTED CV TYPE missing or unclear.")
        return False, issues
    role_achievements = extract_role_achievements(result)
    expected = EXPECTED_COUNTS[cv_type]
    for role, required_count in expected.items():
        actual_count = len(role_achievements.get(role, []))
        if required_count == 0 and actual_count > 0:
            issues.append(f"{role}: must not appear in {cv_type}.")
        elif required_count > 0 and actual_count != required_count:
            issues.append(f"{role}: expected {required_count} achievements; found {actual_count}.")
        for idx, ach in enumerate(role_achievements.get(role, []), start=1):
            if len(ach) < 170 or len(ach) > 190:
                issues.append(f"{role} achievement {idx}: {len(ach)} chars. Required 170-190.")
    return len(issues) == 0, issues

def build_repair_prompt(original_prompt: str, previous_result: str, issues: List[str]) -> str:
    issue_text = "\n".join(f"- {issue}" for issue in issues)
    return f"Repair the previous output. Failures:\n{issue_text}\n\nOriginal Task:\n{original_prompt}\n\nPrevious Output:\n{previous_result}"

def generate_with_strict_validation(prompt: str, output_scope: str, primary_model: str, failover_models: List[str]) -> Tuple[str, str, List[str], List[str]]:
    result, used_model, attempts = call_model(prompt, primary_model, failover_models)
    is_valid, issues = validate_generated_output(result, output_scope)
    round = 0
    while not is_valid and round < 3:
        round += 1
        repair_prompt = build_repair_prompt(prompt, result, issues)
        result, used_model, repair_attempts = call_model(repair_prompt, primary_model, failover_models)
        attempts.extend(repair_attempts)
        is_valid, issues = validate_generated_output(result, output_scope)
    return result, used_model, attempts, issues

def strict_cv_prompt(context: str, job_description: str, language: str, output_scope: str) -> str:
    return f"""
You are CareerOps Studio. Create a tailored CV.
Rules: {MANDATORY_CV_TYPE_RULES}\n{MANDATORY_BULLET_COUNTS}\n{STRICT_FORMAT_RULES}\n{ATS_GAP_ANALYSIS_RULES}\n{FINAL_OUTPUT_POLICY}
Context: {context}
JD: {job_description}
Language: {language}
"""

# --- UI CSS (CHAT MODERN STYLE) ---

def inject_chat_css():
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    
    html, body, [data-testid="stAppViewContainer"] {
        background-color: #0d0d0d !important;
        color: #e3e3e3 !important;
        font-family: 'Inter', sans-serif !important;
    }

    [data-testid="stSidebar"] {
        background-color: #000000 !important;
        border-right: 1px solid #262626 !important;
        width: 300px !important;
    }

    .stChatInput {
        background-color: #0d0d0d !important;
        padding-bottom: 20px;
    }

    .stChatInput > div {
        background-color: #212121 !important;
        border: 1px solid #333 !important;
        border-radius: 12px !important;
    }

    .message-row {
        display: flex;
        padding: 24px 0;
        gap: 20px;
        max-width: 850px;
        margin: 0 auto;
    }

    .avatar {
        width: 32px; height: 32px; border-radius: 4px;
        display: flex; align-items: center; justify-content: center;
        font-weight: bold; flex-shrink: 0;
    }
    .avatar.user { background: #5436da; color: #fff; }
    .avatar.ai { background: #10a37f; color: #fff; }

    .message-body {
        flex-grow: 1;
        font-size: 1rem;
        line-height: 1.6;
    }

    .sidebar-label {
        font-size: 0.7rem; color: #666; font-weight: 700;
        margin-top: 20px; margin-bottom: 8px; text-transform: uppercase;
        letter-spacing: 1px;
    }

    .download-container {
        display: flex; gap: 8px; margin-top: 12px;
    }

    #MainMenu, footer, header { visibility: hidden; }
    .main .block-container { padding-bottom: 150px; }
    
    /* Tabelas no chat */
    .message-body table {
        width: 100%; border-collapse: collapse; margin: 1rem 0;
        background: rgba(255,255,255,0.03); border-radius: 8px;
    }
    .message-body th { background: rgba(16,163,127,0.1); padding: 8px; text-align: left; }
    .message-body td { padding: 8px; border-bottom: 1px solid rgba(255,255,255,0.05); }
    </style>
    """, unsafe_allow_html=True)

# --- APP PRINCIPAL ---

def main():
    st.set_page_config(page_title=APP_TITLE, page_icon="✨", layout="wide")
    inject_chat_css()

    if "chat_id" not in st.session_state:
        st.session_state.chat_id = datetime.now().strftime("%Y%m%d_%H%M%S")
        st.session_state.messages = []
        st.session_state.chat_title = "Novo Chat"

    history = load_history()

    # --- SIDEBAR ---
    with st.sidebar:
        st.markdown(f"### {APP_TITLE}")
        if st.button("➕ Novo Chat", use_container_width=True):
            st.session_state.chat_id = datetime.now().strftime("%Y%m%d_%H%M%S")
            st.session_state.messages = []
            st.session_state.chat_title = "Novo Chat"
            st.rerun()

        st.markdown('<div class="sidebar-label">Histórico</div>', unsafe_allow_html=True)
        for chat in history:
            active = " (Ativo)" if chat["id"] == st.session_state.chat_id else ""
            if st.button(f"💬 {chat['title'][:25]}{active}", key=f"h_{chat['id']}", use_container_width=True):
                st.session_state.chat_id = chat["id"]
                st.session_state.messages = chat["messages"]
                st.session_state.chat_title = chat["title"]
                st.rerun()

        st.markdown('<div class="sidebar-label">Configurações</div>', unsafe_allow_html=True)
        model = st.selectbox("Modelo", MODEL_OPTIONS, index=0)
        depth = st.selectbox("Contexto", ["Fast", "Balanced", "Deep"], index=1)
        
        st.markdown('<div class="sidebar-label">Documentos</div>', unsafe_allow_html=True)
        inst_f = st.file_uploader("Instruções", type=["docx", "pdf", "txt"])
        exp_f = st.file_uploader("Repositório", type=["docx", "pdf", "txt"])
        
        inst_text = read_uploaded_file(inst_f) or load_first_available(["AI_CV_Instructions_Master.docx"])
        exp_text = read_uploaded_file(exp_f) or load_first_available(["Master Experience File.docx"])

    # --- ÁREA DE CHAT ---
    if not st.session_state.messages:
        st.markdown("<div style='text-align: center; margin-top: 15vh;'><h1>CareerOps Studio</h1><p style='color: #8e8ea0;'>Cole a descrição da vaga para começar.</p></div>", unsafe_allow_html=True)
    else:
        for i, msg in enumerate(st.session_state.messages):
            st.markdown(f"""
            <div class="message-row {msg['role']}">
                <div class="avatar {'user' if msg['role'] == 'user' else 'ai'}">{'U' if msg['role'] == 'user' else 'AI'}</div>
                <div class="message-body">{msg['content']}</div>
            </div>
            """, unsafe_allow_html=True)
            
            if msg['role'] == "assistant" and "FINAL TAILORED CV" in msg['content']:
                c1, c2, _ = st.columns([1, 1, 4])
                with c1: st.download_button("📄 PDF", make_pdf(msg['content']), f"cv_{i}.pdf", key=f"pdf_{i}")
                with c2: st.download_button("📋 DOCX", make_docx(msg['content']), f"cv_{i}.docx", key=f"docx_{i}")

    # --- INPUT ---
    user_input = st.chat_input("Cole a descrição da vaga aqui...")

    if user_input:
        st.session_state.messages.append({"role": "user", "content": user_input})
        if st.session_state.chat_title == "Novo Chat":
            st.session_state.chat_title = user_input[:30] + "..."

        with st.spinner("Processando..."):
            try:
                prompt = strict_cv_prompt(f"Instructions: {inst_text[:2000]}\nExp: {exp_text[:5000]}", user_input, "Português", "Full")
                result, used_model, attempts, issues = generate_with_strict_validation(prompt, "Full", model, ["gemini-2.0-flash"])
                
                st.session_state.messages.append({"role": "assistant", "content": result})
                
                new_h = [c for c in history if c["id"] != st.session_state.chat_id]
                new_h.insert(0, {"id": st.session_state.chat_id, "title": st.session_state.chat_title, "messages": st.session_state.messages})
                save_history(new_h)
                st.rerun()
            except Exception as e:
                st.error(f"Erro: {str(e)}")

if __name__ == "__main__":
    main()
