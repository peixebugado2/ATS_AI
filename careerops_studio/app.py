import os
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
    st.caption("multiple keys")
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
    st.markdown('<p class="section-note">Paste one complete job description. The system selects the CV type and maps the strongest verified achievements to the target role.</p>', unsafe_allow_html=True)
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
            with st.spinner("Building the tailored CV and testing failover if needed..."):
                try:
                    result, used_model, attempts = call_model(prompt, primary_model, failover_models)
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
