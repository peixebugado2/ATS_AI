import os
import re
from datetime import datetime
from io import BytesIO
from pathlib import Path

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
APP_SUBTITLE = "Executive resume optimization workspace"
DEFAULT_MODEL = "gemini-2.5-flash"
ROOT = Path(__file__).parent
TEMPLATE_DIR = ROOT / "templates"


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
The cover letter body must be exactly 1,800 characters including spaces, or as close as technically possible within 1,790 to 1,810 characters.
Do not include placeholder text.

FINAL ATS REPORT
A concise report only.

Do not show internal reasoning, keyword dumps, scoring logs, chain-of-thought, draft alternatives or validation steps.
"""

STRICT_VALIDATION_POLICY = """
STRICT VALIDATION POLICY

Before final answer, internally verify:
- Correct CV type selected.
- Correct role inclusion/exclusion.
- Exact mandatory achievement line counts.
- Every achievement line between 170 and 190 characters including spaces.
- Manager CV has exactly 8 Trade Compliance Manager EMEA achievement lines, not 7.
- Cover Letter section, when present, is approximately 1,800 characters and must not be too long.
- ATS validation is clearly an internal simulation, not live website data.
- ATS skills are extracted from the Job Description and matched against the CV.
"""

def inject_css():
    st.markdown(
        """
        <style>
        :root {
            --bg: #090b10;
            --panel: rgba(20, 24, 32, .78);
            --panel-strong: rgba(25, 30, 40, .92);
            --border: rgba(218, 197, 167, .18);
            --text: #f4efe8;
            --muted: rgba(244,239,232,.68);
            --gold: #d7b879;
            --gold-soft: rgba(215,184,121,.14);
            --blue: #9bb6d3;
            --green: #8fd0a4;
            --danger: #e88c8c;
        }
        .stApp {
            color: var(--text);
            background:
                radial-gradient(circle at 16% 0%, rgba(215,184,121,.16), transparent 28%),
                radial-gradient(circle at 88% 12%, rgba(155,182,211,.12), transparent 24%),
                linear-gradient(145deg, #06070a 0%, #0d1118 45%, #090b10 100%);
        }
        [data-testid="stHeader"] { background: rgba(0,0,0,0); }
        [data-testid="stSidebar"] {
            background: linear-gradient(180deg, rgba(11,14,20,.98), rgba(13,17,24,.92));
            border-right: 1px solid var(--border);
        }
        [data-testid="stSidebar"] * { color: var(--text); }
        .block-container { padding-top: 1.6rem; max-width: 1220px; }
        .topbar {
            display:flex; align-items:center; justify-content:space-between; gap:16px;
            padding: 10px 2px 22px 2px;
        }
        .brand { font-size: .82rem; letter-spacing:.28em; text-transform:uppercase; color:var(--gold); }
        .status-chip {
            border:1px solid var(--border); background:var(--gold-soft); color:var(--text);
            border-radius:999px; padding:7px 11px; font-size:.78rem;
        }
        .hero {
            border: 1px solid var(--border);
            background:
                linear-gradient(135deg, rgba(255,255,255,.065), rgba(255,255,255,.025)),
                linear-gradient(135deg, rgba(215,184,121,.13), rgba(155,182,211,.055));
            border-radius: 28px;
            padding: 34px 36px;
            box-shadow: 0 30px 90px rgba(0,0,0,.42);
            margin-bottom: 20px;
        }
        .hero h1 { margin:0; font-size:3.1rem; letter-spacing:-.065em; line-height:.95; }
        .hero p { margin:14px 0 0; color:var(--muted); font-size:1.02rem; max-width:760px; }
        .pill-row { display:flex; flex-wrap:wrap; gap:10px; margin-top:22px; }
        .pill { border:1px solid var(--border); background:rgba(255,255,255,.045); border-radius:999px; padding:8px 12px; color:rgba(244,239,232,.84); font-size:.83rem; }
        .card {
            border:1px solid var(--border); background:var(--panel); border-radius:22px;
            padding:22px 24px; box-shadow: 0 18px 55px rgba(0,0,0,.28); margin-bottom:16px;
        }
        .metric-card {
            border:1px solid var(--border); background:rgba(255,255,255,.045); border-radius:20px;
            padding:18px; min-height:122px;
        }
        .metric-kicker { color:var(--muted); font-size:.78rem; text-transform:uppercase; letter-spacing:.12em; }
        .metric-value { margin-top:8px; font-size:1.45rem; font-weight:750; letter-spacing:-.035em; }
        .metric-hint { margin-top:9px; color:var(--muted); font-size:.84rem; }
        .section-title { font-size:1.32rem; font-weight:780; letter-spacing:-.03em; margin:0 0 6px; }
        .section-note { color:var(--muted); margin:0 0 18px; }
        .result-shell { border:1px solid var(--border); background:rgba(255,255,255,.055); border-radius:22px; padding:22px; margin-top:18px; }
        .footer-note { text-align:center; color:var(--muted); font-size:.82rem; padding:30px 0 10px; }
        .stButton > button, .stDownloadButton > button {
            border-radius: 14px !important; border:1px solid var(--border) !important; font-weight:700 !important;
            padding:.68rem 1rem !important;
        }
        .stButton > button[kind="primary"] {
            color:#090b10 !important; border:0 !important;
            background: linear-gradient(90deg, #f0d696, #c7a45f) !important;
        }
        textarea, input, .stTextArea textarea, .stTextInput input {
            border-radius: 15px !important;
        }
        [data-testid="stFileUploader"] {
            border:1px dashed rgba(215,184,121,.28); background:rgba(255,255,255,.035); border-radius:16px; padding:10px;
        }
        div[data-testid="stTabs"] button { border-radius:999px !important; padding:10px 18px !important; }
        hr { border-color: var(--border); }
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


def get_client():
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        st.error("GEMINI_API_KEY não encontrada. Crie um arquivo .env com GEMINI_API_KEY=sua_chave")
        st.stop()
    return genai.Client(api_key=api_key)


def read_uploaded_file(uploaded_file) -> str:
    if uploaded_file is None:
        return ""
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
            return "[PDF enviado, mas pypdf não está instalado.]"
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
    styles["Normal"].font.size = None
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
        elif stripped.startswith("| "):
            doc.add_paragraph(stripped)
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


def call_model(prompt: str, model: str) -> str:
    client = get_client()
    response = client.models.generate_content(model=model, contents=prompt)
    return response.text or ""


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


def normalize_role_name(line: str) -> str:
    cleaned = re.sub(r"^[#\-\*\s\d\.\)]+", "", line or "").strip()
    cleaned = cleaned.replace(":", "").strip()
    return cleaned


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

    # Remove section title line.
    lines = after.splitlines()
    if lines:
        lines = lines[1:]

    stop_markers = [
        "FINAL ATS REPORT",
        "ATS REPORT",
        "FINAL REPORT",
        "SOURCE CONTROL",
    ]

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
                issues.append(f"Cover Letter length is {cover_len} characters. Required approximately 1,800, accepted 1,790-1,810.")
        elif "Cover Letter" in output_scope:
            issues.append("Cover Letter was requested but not found.")

    lower = (result or "").lower()
    if "linkedin ats" in lower or "jobscan" in lower or "skillsyncer" in lower or "resume worded" in lower or "rezi" in lower:
        if "simulation" not in lower and "simulated" not in lower and "ats-style" not in lower:
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

def build_context(instructions, master, manager, director):
    return f"""
[OPERATING INSTRUCTIONS]
{truncate_text(instructions, 35000)}

[EXPERIENCE REPOSITORY - SOURCE OF TRUTH]
{truncate_text(master, 90000)}

[MANAGER CV TEMPLATE]
{truncate_text(manager, 18000)}

[DIRECTOR CV TEMPLATE]
{truncate_text(director, 18000)}
"""


def render_downloads(result: str, base_name: str):
    st.markdown('<div class="section-title">Exportar</div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    with c1:
        st.download_button("TXT", result.encode("utf-8"), f"{base_name}.txt", "text/plain", use_container_width=True)
    with c2:
        st.download_button("DOCX", make_docx(result), f"{base_name}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
    with c3:
        st.download_button("PDF", make_pdf(result), f"{base_name}.pdf", "application/pdf", use_container_width=True)


def file_metric(label: str, text: str, required: bool = False):
    loaded = bool(text.strip())
    status = "Ready" if loaded else ("Required" if required else "Optional")
    value = f"{len(text):,}".replace(",", ".")
    st.markdown(
        f"""
        <div class="metric-card">
          <div class="metric-kicker">{label}</div>
          <div class="metric-value">{status}</div>
          <div class="metric-hint">{value} characters indexed</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def strict_cv_prompt(context: str, job_description: str, language: str, output_scope: str) -> str:
    return f"""
You are CareerOps Studio, an executive CV optimization engine. Do not describe yourself as AI in the output.

PRIMARY OBJECTIVE
Create a tailored CV from the Job Description using the Experience Repository as the only source of truth. The user normally copies the output into an existing CV, so the final CV text must be clean, usable, and faithful to the candidate's real achievements.

ABSOLUTE RULES
1. Never invent achievements, metrics, companies, roles, dates, systems, certifications, countries, savings, headcount, budget or KPIs.
2. Select the best achievements from the Experience Repository based on the Job Description.
3. You may adapt wording, merge compatible evidence, and improve ATS keyword alignment, but facts and metrics must remain faithful.
4. Keep the candidate's roles; do not change role titles unless the template explicitly contains that role title.
5. Choose Manager CV or Director CV using the rules in the instructions.
6. Every achievement line in Professional Experience must be 170-190 characters including spaces.
7. Final CV section must not use bullet symbols. Put each achievement on its own line.
8. No blank lines between achievements inside the same role. Use one blank line between roles.
9. Additional ATS Skills must contain exactly 10 skills in a clean vertical list.
10. ATS score must be realistic. Never claim 100% unless every critical keyword is naturally covered.
11. Cover letter must be approximately 1800 characters including spaces, same language as the job description unless the user selected another language.

MANDATORY PROCESS
A. Analyze the Job Description: title, seniority, industry, hard skills, soft skills, systems, certifications, leadership requirements and critical ATS keywords.
B. Select CV Type: Manager or Director, with a short reason.
C. Score and select achievements by keyword match, responsibility match, industry match, leadership match and systems/tools match.
D. Build the CV using the selected template and the selected achievements only.
E. Run the optimization loop: identify keyword gaps, inject missing keywords naturally, and recalculate ATS.
F. Validate all achievement character counts internally and rewrite until compliant.
G. Return a final ATS report with realistic scores.

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
- Realistic ATS score across LinkedIn ATS, Jobscan, SkillSyncer, Resume Worded and Rezi

Do not claim 100% unless every critical keyword is covered.
Return practical corrections if needed.
Language: {language}

CONTEXT FILES
{context}

JOB DESCRIPTION
{job_description}

CV TO VALIDATE
{cv_text}
"""


st.set_page_config(page_title=APP_TITLE, page_icon="◆", layout="wide")
inject_css()

st.markdown(
    f"""
    <div class="topbar">
      <div class="brand">{APP_TITLE}</div>
      <div class="status-chip">Private workspace</div>
    </div>
    <div class="hero">
      <h1>Resume optimization, built for executive applications.</h1>
      <p>{APP_SUBTITLE}. Select the right CV path, map real achievements to a role, validate ATS fit, and export clean client-ready documents.</p>
      <div class="pill-row">
        <span class="pill">Experience-led</span>
        <span class="pill">ATS validation</span>
        <span class="pill">Manager / Director routing</span>
        <span class="pill">PDF & DOCX export</span>
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.markdown("### Workspace")
    model = st.text_input("Model", value=DEFAULT_MODEL)
    st.caption("Keep provider details internal when sharing with clients.")
    st.divider()
    st.markdown("### Source documents")
    instructions_file = st.file_uploader("Instructions", type=["txt", "md", "docx", "pdf"], key="instructions")
    master_file = st.file_uploader("Experience Repository", type=["txt", "md", "docx", "pdf"], key="master")
    manager_file = st.file_uploader("Manager CV Template", type=["txt", "md", "docx", "pdf"], key="manager")
    director_file = st.file_uploader("Director CV Template", type=["txt", "md", "docx", "pdf"], key="director")
    st.divider()
    st.caption("For production: keep these files preloaded on the server, not uploaded by the client.")

instructions_text = read_uploaded_file(instructions_file) or load_template_text("AI_CV_Instructions_Master_Rev7.docx")
master_text = read_uploaded_file(master_file) or load_template_text("Master_Experience_File_Structured_Model.docx")
manager_text = read_uploaded_file(manager_file) or load_template_text("Manager_CV_Template.docx")
director_text = read_uploaded_file(director_file) or load_template_text("Director_CV_Template.docx")
context = build_context(instructions_text, master_text, manager_text, director_text)

c1, c2, c3, c4 = st.columns(4)
with c1:
    file_metric("Instructions", instructions_text)
with c2:
    file_metric("Experience repository", master_text, required=True)
with c3:
    file_metric("Manager template", manager_text)
with c4:
    file_metric("Director template", director_text)

main_tab, validate_tab, files_tab = st.tabs(["Application workspace", "ATS validation", "Source control"])

with main_tab:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Create tailored CV</div>', unsafe_allow_html=True)
    st.markdown('<p class="section-note">Paste one complete job description. The system will select the CV type and map the most valuable real achievements to the role.</p>', unsafe_allow_html=True)
    job_description = st.text_area("Job description", height=360, placeholder="Paste the full job description here...")
    c1, c2 = st.columns([1, 1])
    with c1:
        language = st.selectbox("Output language", ["Same as job description", "English", "Português do Brasil"])
    with c2:
        output_scope = st.selectbox(
            "Output package",
            [
                "Full package: Job Analysis, ATS Validation Status, Tailored CV, Cover Letter, Final ATS Report",
                "Tailored CV only with ATS Validation Status",
                "ATS Analysis and achievement recommendations only",
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
            with st.spinner("Building the tailored CV..."):
                result = call_model(prompt, model)
            st.session_state["last_result"] = result
            st.session_state["last_job"] = job_description
            st.markdown('<div class="result-shell">', unsafe_allow_html=True)
            st.markdown(result)
            st.markdown('</div>', unsafe_allow_html=True)
            render_downloads(result, f"careerops_cv_{datetime.now().strftime('%Y%m%d_%H%M')}")

with validate_tab:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Validate existing CV against a job</div>', unsafe_allow_html=True)
    st.markdown('<p class="section-note">Use this when the client already has a CV and wants to verify ATS match, keyword gaps and role alignment.</p>', unsafe_allow_html=True)
    existing_cv = st.text_area("Existing CV text", height=260, value=st.session_state.get("last_result", ""))
    validation_job = st.text_area("Job description for validation", height=240, value=st.session_state.get("last_job", ""))
    validate = st.button("Run ATS validation", type="primary", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)
    if validate:
        if not existing_cv.strip() or not validation_job.strip():
            st.warning("Paste both the CV and the job description.")
        else:
            prompt = ats_validation_prompt(context, existing_cv, validation_job, "Same as job description")
            with st.spinner("Validating ATS fit..."):
                result = call_model(prompt, model)
            st.markdown('<div class="result-shell">', unsafe_allow_html=True)
            st.markdown(result)
            st.markdown('</div>', unsafe_allow_html=True)
            render_downloads(result, f"careerops_validation_{datetime.now().strftime('%Y%m%d_%H%M')}")

with files_tab:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Loaded source files</div>', unsafe_allow_html=True)
    st.markdown('<p class="section-note">Confirm whether the source material was loaded correctly. The Experience Repository is the source of truth.</p>', unsafe_allow_html=True)
    st.write("Instructions characters:", len(instructions_text))
    st.write("Experience Repository characters:", len(master_text))
    st.write("Manager Template characters:", len(manager_text))
    st.write("Director Template characters:", len(director_text))
    st.markdown('</div>', unsafe_allow_html=True)
    with st.expander("Preview instructions"):
        st.text(truncate_text(instructions_text, 6000))
    with st.expander("Preview Experience Repository"):
        st.text(truncate_text(master_text, 6000))

st.markdown('<div class="footer-note">CareerOps Studio - Private executive resume optimization workspace</div>', unsafe_allow_html=True)
