import os
from dotenv import load_dotenv
load_dotenv()
import streamlit as st
import pinecone
from pinecone import Pinecone, ServerlessSpec
from langchain_openai import ChatOpenAI, OpenAIEmbeddings
from langchain_core.messages import HumanMessage, SystemMessage
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
from datetime import datetime
import json
import hashlib
from collections import Counter
import pandas as pd
import httpx
from urllib.parse import quote_plus

# ----- BRANDING -----
COMPANY_NAME = "Quantium Insights LLC"
APP_NAME = "AfriCareer AI"
TAGLINE = "Empowering African Youth Through Intelligent Career Solutions"

# Supported languages — rendered at the TOP of the main page (mobile-first: no sidebar needed)
LANGUAGES = {
    "🇬🇧 English": "English",
    "🇫🇷 Français": "French",
    "🇰🇪 Kiswahili": "Swahili",
    "🇸🇦 العربية": "Arabic",
    "🇳🇬 Hausa": "Hausa",
    "🇳🇬 Pidgin": "Nigerian Pidgin",
    "🇵🇹 Português": "Portuguese",
    "🇪🇸 Español": "Spanish",
    "🇪🇹 አማርኛ": "Amharic",
}
target_lang = "English"  # module-level default; overwritten by the picker in main()

# Admin credentials (override via env/secret in production; defaults kept for local dev)
ADMIN_USERNAME = os.getenv("ADMIN_USERNAME", "admin")
ADMIN_PASSWORD_HASH = hashlib.sha256(
    os.getenv("ADMIN_PASSWORD", "africareer2024").encode()
).hexdigest()

# ----- ACCESS CONTROL & COST GUARDRAILS -----
# Optional invite code for pilot gating. If unset, the app is open (local/dev).
APP_ACCESS_CODE = os.getenv("APP_ACCESS_CODE", "").strip()
# Cap AI calls per browser session to protect the OpenAI budget on public deployments.
MAX_LLM_CALLS_PER_SESSION = int(os.getenv("MAX_LLM_CALLS_PER_SESSION", "30"))
QUOTA_MESSAGE = (
    "⚠️ You've reached this session's usage limit for AI guidance. "
    "This limit keeps the service free and available for everyone. "
    "Please refresh the page later to continue, or contact the AfriCareer AI team for extended access."
)

# ===== SAFETY GUARDRAILS SYSTEM MESSAGE - MAXIMALLY HELPFUL =====
SAFETY_SYSTEM_MESSAGE = """You are AfriCareer AI, a comprehensive career guidance assistant for African youth.

YOUR MISSION: Help with ANY career, job, education, or professional development question. Be as helpful as possible!

YOU SHOULD ANSWER ALL QUESTIONS ABOUT:
✅ Career guidance, planning, and transitions (any industry, any level)
✅ Resume/CV creation, improvement, and tailoring (for any job)
✅ Job searching strategies and platforms
✅ Interview preparation (for ANY specific job or company)
✅ Interview tips, common questions, how to answer behavioral questions
✅ Salary negotiation and job offer evaluation
✅ Educational pathways, courses, certifications, degrees
✅ Professional development and skill building
✅ Entrepreneurship, business planning, startups
✅ Workplace issues, conflicts, communication
✅ Work-life balance and career satisfaction
✅ Networking and professional relationships
✅ Performance reviews and career advancement
✅ Industry-specific guidance (tech, healthcare, finance, agriculture, etc.)
✅ Job market trends and employment statistics
✅ Questions about AfDB, UNICEF, ILO frameworks
✅ TVET, STEM, digital skills programs
✅ Informal sector work and apprenticeships
✅ Freelancing and remote work
✅ Career changes and pivots
✅ Leadership and management skills
✅ Personal branding and online presence

IMPORTANT: If someone asks how to prepare for an interview at a specific company (Google, Microsoft, banks, NGOs, etc.) or for a specific role (software engineer, nurse, teacher, etc.), provide detailed, helpful guidance!

YOU MUST REFUSE ONLY THESE THREE THINGS:
❌ Sexual or explicit content
❌ Violence, harmful behavior, or instructions for illegal activities  
❌ That's it! These are the ONLY restrictions.

WHAT YOU SHOULD NOT REFUSE:
✅ Questions about salary, compensation, benefits (this is career advice!)
✅ Questions about workplace conflicts or difficult bosses (this is professional development!)
✅ Questions about specific companies, industries, or job roles
✅ Questions about career strategies that might be considered "political" (working in government, NGOs, international organizations)
✅ Questions about personal career decisions (this is career counseling!)

If and ONLY if a request involves sexual content, violence, or illegal activity, respond with:
"I'm AfriCareer AI for career guidance. I can't help with that specific topic, but I'm here to help with any career, job, education, or professional development question. What career-related question can I help you with?"

RESPONSE STYLE:
- Be MAXIMALLY HELPFUL - if it's career-related, answer it thoroughly!
- Provide specific, actionable, practical advice
- When relevant frameworks exist in knowledge base (AfDB SEPA, UNICEF, ILO), cite them
- When frameworks don't exist, use general career guidance best practices
- Ground responses in African context when applicable
- Be conversational, supportive, and encouraging
- Don't be overly cautious or restrictive - if someone asks a career question, help them!"""

# ----- PREMIUM THEME & STYLING -----
st.set_page_config(
    page_title=f"{APP_NAME} - Career Guidance for African Youth",
    page_icon="🌍",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ===== LIGHT, MODERN THEME (OpenAI-style; base colors set in .streamlit/config.toml) =====
st.markdown("""
<style>
    :root {
        --ink: #15161B;
        --muted: #5B5F6B;
        --accent: #4C6FFF;
        --accent2: #7A5CFF;
        --line: #E6E8EE;
        --soft: #F4F6FA;
    }

    /* Clean white canvas with a soft cool wash at the top */
    .stApp {
        background:
            radial-gradient(900px 380px at 82% -6%, rgba(122,92,255,0.08), transparent 60%),
            radial-gradient(900px 380px at 12% -6%, rgba(76,111,255,0.08), transparent 60%),
            #FFFFFF;
        overflow-x: hidden;
    }

    .block-container { max-width: 1000px; padding-top: 2.2rem; }

    /* Clean modern sans-serif everywhere */
    html, body, [class*="css"], .stMarkdown, p, li, label, input, textarea, button, select {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif !important;
    }
    h1, h2, h3, h4 { color: var(--ink) !important; font-weight: 700; letter-spacing: -0.01em; }
    p, li, label, .stMarkdown { color: #2B2D36; }
    [data-testid="stCaptionContainer"], .stCaption { color: var(--muted) !important; }
    a, a:visited { color: var(--accent) !important; text-decoration: none; }
    a:hover { text-decoration: underline; }
    hr { border-color: var(--line) !important; }

    /* Primary buttons: friendly blue -> indigo */
    .stButton > button {
        background: linear-gradient(135deg, #4C6FFF 0%, #7A5CFF 100%);
        color: #FFFFFF !important;
        border: none;
        border-radius: 10px;
        padding: 11px 24px;
        font-weight: 600;
        box-shadow: 0 6px 16px rgba(76,111,255,0.22);
        transition: all .2s ease;
    }
    .stButton > button:hover { transform: translateY(-1px); box-shadow: 0 10px 22px rgba(76,111,255,0.30); }

    /* Download buttons: fresh green accent for the "take-away" action */
    .stDownloadButton > button {
        background: linear-gradient(135deg, #17B26A 0%, #0E9E5C 100%) !important;
        color: #FFFFFF !important;
        border: none; border-radius: 10px; padding: 11px 24px; font-weight: 600;
        box-shadow: 0 6px 16px rgba(23,178,106,0.22);
    }

    /* Tabs: clean underline */
    .stTabs [data-baseweb="tab-list"] { gap: 4px; border-bottom: 1px solid var(--line); flex-wrap: wrap; }
    .stTabs [data-baseweb="tab"] { background: transparent; color: var(--muted) !important; border-radius: 8px 8px 0 0; padding: 10px 18px; font-weight: 600; }
    .stTabs [data-baseweb="tab"] p { color: inherit !important; font-weight: 600; }
    .stTabs [aria-selected="true"] { background: rgba(76,111,255,0.08) !important; border-bottom: 2px solid var(--accent); }
    .stTabs [aria-selected="true"] p { color: var(--accent) !important; }

    /* Inputs: white with soft border */
    .stTextInput input,
    .stTextArea textarea,
    .stSelectbox div[data-baseweb="select"] > div {
        background: #FFFFFF !important;
        border: 1px solid var(--line) !important;
        border-radius: 10px !important;
        color: var(--ink) !important;
    }

    /* File uploader dropzone */
    [data-testid="stFileUploaderDropzone"] {
        background: var(--soft) !important;
        border: 1px dashed rgba(76,111,255,0.45) !important;
        border-radius: 12px;
    }

    .stAlert { border-radius: 12px; }

    /* Sidebar: soft light with a divider */
    [data-testid="stSidebar"] { background: var(--soft); border-right: 1px solid var(--line); }

    /* Landing hero: warm+cool blended gradient (drop in a licensed photo later via background-image) */
    .africareer-hero {
        border-radius: 20px;
        padding: 3rem 2rem;
        text-align: center;
        background:
            radial-gradient(600px 300px at 15% 20%, rgba(255,159,67,0.55), transparent 60%),
            radial-gradient(600px 300px at 85% 25%, rgba(122,92,255,0.60), transparent 60%),
            radial-gradient(700px 360px at 50% 115%, rgba(23,178,106,0.55), transparent 60%),
            linear-gradient(135deg, #4C6FFF 0%, #7A5CFF 100%);
        box-shadow: 0 20px 50px rgba(76,111,255,0.25);
    }
    .africareer-hero h1 { color: #FFFFFF !important; font-size: 2.4rem; margin: 0 0 .5rem 0; }
    .africareer-hero p { color: rgba(255,255,255,0.94) !important; font-size: 1.08rem; max-width: 640px; margin: 0 auto; }
    .hero-chips { margin-top: 1.4rem; display: flex; gap: .5rem; flex-wrap: wrap; justify-content: center; }
    .hero-chip { background: rgba(255,255,255,0.18); color: #fff; padding: .4rem .9rem; border-radius: 999px; font-size: .9rem; font-weight: 600; }

    /* Hide Streamlit chrome for a cleaner, app-like surface */
    #MainMenu { visibility: hidden; }
    footer { visibility: hidden; }

    /* ===== MOBILE RESPONSIVENESS (phones, <= 640px) ===== */
    @media (max-width: 640px) {
        /* Stack columns vertically instead of cramming them side-by-side */
        [data-testid="stHorizontalBlock"] { flex-wrap: wrap !important; }
        [data-testid="stHorizontalBlock"] > div,
        [data-testid="column"],
        [data-testid="stColumn"] {
            flex: 1 1 100% !important;
            min-width: 100% !important;
        }

        /* Use the full narrow width; less wasted padding */
        .block-container {
            padding: 1rem 0.8rem 3rem 0.8rem !important;
        }

        /* Full-width, easy-to-tap buttons */
        .stButton > button,
        .stDownloadButton > button {
            width: 100% !important;
            padding: 14px 18px !important;
            font-size: 1rem !important;
        }

        /* Headings scale down so they don't wrap awkwardly */
        h1 { font-size: 1.5rem !important; }
        h2 { font-size: 1.25rem !important; }
        h3 { font-size: 1.1rem !important; }
        .africareer-hero { padding: 2rem 1rem; }
        .africareer-hero h1 { font-size: 1.7rem; }

        /* Tabs: scroll horizontally, keep labels tappable */
        .stTabs [data-baseweb="tab-list"] {
            overflow-x: auto !important;
            flex-wrap: nowrap !important;
        }
        .stTabs [data-baseweb="tab"] {
            padding: 10px 14px !important;
            font-size: 0.9rem !important;
        }

        /* 16px inputs stop iOS Safari from auto-zooming when typing */
        textarea, input, .stTextInput input, .stTextArea textarea {
            font-size: 16px !important;
        }
    }
</style>
""", unsafe_allow_html=True)

# ----- PINECONE & OPENAI SETUP -----
try:
    PINECONE_API_KEY = os.getenv("PINECONE_API_KEY")
    OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
    
    if not PINECONE_API_KEY or not OPENAI_API_KEY:
        st.error("The service is not fully configured yet. Please contact the AfriCareer AI team.")
        st.stop()
    
    pc = Pinecone(api_key=PINECONE_API_KEY)
    
    INDEX_NAME = "africareer-kb"
    if INDEX_NAME not in [idx.name for idx in pc.list_indexes()]:
        pc.create_index(
            name=INDEX_NAME,
            dimension=1536,
            metric="cosine",
            spec=ServerlessSpec(cloud="aws", region="us-east-1")
        )
    
    index = pc.Index(INDEX_NAME)
    embeddings = OpenAIEmbeddings(openai_api_key=OPENAI_API_KEY)
    
    # ===== CRITICAL: LLM WITH SAFETY SYSTEM MESSAGE =====
    llm = ChatOpenAI(
        temperature=0.7,
        model="gpt-4o-mini",
        openai_api_key=OPENAI_API_KEY
    )
    
except Exception as e:
    st.error(f"Setup Error: {str(e)}")
    st.stop()

# ----- ANALYTICS (Supabase when configured; local JSON fallback) -----
# On ephemeral hosts (HF Spaces, Render free) the JSON file resets on restart.
# Set SUPABASE_URL + SUPABASE_KEY (and create an `analytics` table) for durable storage.
ANALYTICS_FILE = "africareer_analytics.json"
SUPABASE_URL = os.getenv("SUPABASE_URL", "").strip()
SUPABASE_KEY = os.getenv("SUPABASE_KEY", "").strip()

@st.cache_resource
def _get_supabase():
    """Return a Supabase client if configured and the package is installed, else None."""
    if not (SUPABASE_URL and SUPABASE_KEY):
        return None
    try:
        from supabase import create_client
        return create_client(SUPABASE_URL, SUPABASE_KEY)
    except Exception:
        return None

def log_analytics(event_type, details=None):
    record = {
        "timestamp": datetime.now().isoformat(),
        "event": event_type,
        "details": details,
    }
    # Prefer durable storage when available
    client = _get_supabase()
    if client is not None:
        try:
            client.table("analytics").insert(record).execute()
            return
        except Exception:
            pass  # fall through to local file
    # Local JSON fallback (ephemeral on cloud hosts)
    try:
        if os.path.exists(ANALYTICS_FILE):
            with open(ANALYTICS_FILE, 'r') as f:
                analytics = json.load(f)
        else:
            analytics = []
        analytics.append(record)
        with open(ANALYTICS_FILE, 'w') as f:
            json.dump(analytics, f)
    except Exception:
        pass

def load_analytics():
    """Load analytics from Supabase if configured, else from local JSON."""
    client = _get_supabase()
    if client is not None:
        try:
            resp = client.table("analytics").select("*").order("timestamp").execute()
            return resp.data or []
        except Exception:
            pass
    try:
        if os.path.exists(ANALYTICS_FILE):
            with open(ANALYTICS_FILE, 'r') as f:
                return json.load(f)
    except Exception:
        pass
    return []

# ===== ENHANCED RAG RETRIEVAL WITH GUARDRAILS =====
def retrieve_career_guidance(query, top_k=5):
    """Retrieve relevant guidance from knowledge base (AfDB, UNICEF, ILO)"""
    try:
        query_vec = embeddings.embed_query(query)
        results = index.query(vector=query_vec, top_k=top_k, include_metadata=True)
        
        context_pieces = []
        sources = []
        
        for match in results["matches"]:
            if match.get("metadata") and match.get("score", 0) > 0.7:
                context_pieces.append(match["metadata"]["text"])
                if "source" in match["metadata"]:
                    sources.append(match["metadata"]["source"])
        
        context = "\n\n".join(context_pieces) if context_pieces else ""
        
        # Add source attribution
        if sources:
            unique_sources = list(set(sources))
            context += f"\n\n[Sources: {', '.join(unique_sources)}]"
        
        return context
    except Exception as e:
        return ""

# ===== SAFE LLM CALL WITH GUARDRAILS =====
def safe_llm_call(user_prompt, rag_context="", language="English"):
    """
    Makes LLM call with safety guardrails and RAG grounding.
    ALL LLM calls in the app should use this function.
    """
    # ---- Cost guardrail: cap AI calls per session to protect the OpenAI budget ----
    _n = st.session_state.get("llm_calls", 0)
    if _n >= MAX_LLM_CALLS_PER_SESSION:
        return QUOTA_MESSAGE
    st.session_state["llm_calls"] = _n + 1

    # Build system message with safety rules
    system_message = SystemMessage(content=SAFETY_SYSTEM_MESSAGE)
    
    # Build user message with RAG context
    if rag_context:
        full_prompt = f"""Language: {language}

CONTEXT FROM AUTHORITATIVE SOURCES (AfDB SEPA, UNICEF Education Strategy, ILO Youth Employment):
{rag_context}

USER REQUEST:
{user_prompt}

Provide your response in {language}, grounded in the context above. Cite specific frameworks when relevant (e.g., "According to AfDB SEPA..." or "UNICEF's Education Strategy emphasizes...")"""
    else:
        full_prompt = f"""Language: {language}

USER REQUEST:
{user_prompt}

Provide your response in {language}."""
    
    user_message = HumanMessage(content=full_prompt)
    
    try:
        response = llm.invoke([system_message, user_message])
        return response.content
    except Exception as e:
        return f"Error: {str(e)}"

# ===== PREMIUM CV DOCUMENT GENERATOR =====
def generate_premium_cv_docx(cv_json_str):
    """Generate a premium 2-page ATS-optimized CV as DOCX from LLM JSON output."""
    from docx.shared import Pt, RGBColor, Inches, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    # Parse LLM JSON
    try:
        cv = json.loads(cv_json_str)
    except json.JSONDecodeError:
        # Attempt to extract JSON from markdown code block
        import re
        match = re.search(r'```(?:json)?\s*([\s\S]*?)```', cv_json_str)
        if match:
            cv = json.loads(match.group(1))
        else:
            cv = json.loads(cv_json_str.strip())

    doc = Document()

    # Page margins
    for section in doc.sections:
        section.top_margin = Cm(1.5)
        section.bottom_margin = Cm(1.2)
        section.left_margin = Cm(2.0)
        section.right_margin = Cm(2.0)

    NAVY = RGBColor(0x0F, 0x2B, 0x4C)
    TEAL = RGBColor(0x14, 0xB8, 0xA6)
    GRAY = RGBColor(0x4A, 0x55, 0x68)
    DARK = RGBColor(0x2D, 0x2D, 0x2D)

    def add_divider(doc_obj):
        p = doc_obj.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single')
        bottom.set(qn('w:sz'), '8')
        bottom.set(qn('w:color'), '14B8A6')
        pBdr.append(bottom)
        pPr.append(pBdr)

    def set_run(run, size=11, color=DARK, bold=False, italic=False, font_name="Georgia"):
        run.font.size = Pt(size)
        run.font.color.rgb = color
        run.bold = bold
        run.italic = italic
        run.font.name = font_name

    # ===== HEADER: NAME =====
    name_para = doc.add_paragraph()
    name_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    name_para.paragraph_format.space_after = Pt(2)
    r = name_para.add_run(cv.get("full_name", "CANDIDATE NAME").upper())
    set_run(r, size=18, color=NAVY, bold=True)
    r.font.character_spacing = Pt(2)

    # Credentials line
    creds = cv.get("credentials", "")
    if creds:
        cred_para = doc.add_paragraph()
        cred_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cred_para.paragraph_format.space_after = Pt(2)
        r = cred_para.add_run(creds)
        set_run(r, size=9.5, color=GRAY, italic=True)

    # Contact line
    contact = cv.get("contact_line", "")
    if contact:
        contact_para = doc.add_paragraph()
        contact_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        contact_para.paragraph_format.space_after = Pt(2)
        r = contact_para.add_run(contact)
        set_run(r, size=9.5, color=GRAY)

    add_divider(doc)

    # ===== PROFESSIONAL SUMMARY =====
    def add_section_heading(title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        r = p.add_run(title.upper())
        set_run(r, size=12, color=NAVY, bold=True)
        # Add thin teal underline
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single')
        bottom.set(qn('w:sz'), '4')
        bottom.set(qn('w:color'), '14B8A6')
        pBdr.append(bottom)
        pPr.append(pBdr)

    summary = cv.get("professional_summary", "")
    if summary:
        add_section_heading("Professional Summary")
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(6)
        r = p.add_run(summary)
        set_run(r, size=10.5, color=DARK)

    # ===== CORE COMPETENCIES =====
    competencies = cv.get("core_competencies", [])
    if competencies:
        add_section_heading("Core Competencies")
        # Display as pipe-separated single block for ATS
        comp_text = "  •  ".join(competencies)
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(6)
        r = p.add_run(comp_text)
        set_run(r, size=10, color=DARK)

    # ===== WORK EXPERIENCE =====
    experience = cv.get("work_experience", [])
    if experience:
        add_section_heading("Professional Experience")
        for job in experience:
            # Job title + Company line
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(1)
            r = p.add_run(job.get("title", ""))
            set_run(r, size=11, color=NAVY, bold=True)
            if job.get("company"):
                r = p.add_run(f" — {job['company']}")
                set_run(r, size=11, color=DARK)

            # Location + Dates
            loc_date = []
            if job.get("location"):
                loc_date.append(job["location"])
            if job.get("dates"):
                loc_date.append(job["dates"])
            if loc_date:
                p = doc.add_paragraph()
                p.paragraph_format.space_after = Pt(3)
                r = p.add_run(" | ".join(loc_date))
                set_run(r, size=9.5, color=GRAY, italic=True)

            # Bullets
            for bullet in job.get("bullets", []):
                p = doc.add_paragraph(style='List Bullet')
                p.paragraph_format.space_after = Pt(1)
                p.paragraph_format.left_indent = Cm(0.8)
                p.clear()
                r = p.add_run(bullet)
                set_run(r, size=10, color=DARK)

    # ===== EDUCATION =====
    education = cv.get("education", [])
    if education:
        add_section_heading("Education")
        for edu in education:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(1)
            r = p.add_run(edu.get("degree", ""))
            set_run(r, size=11, color=NAVY, bold=True)
            if edu.get("institution"):
                p2 = doc.add_paragraph()
                p2.paragraph_format.space_after = Pt(1)
                r = p2.add_run(f"{edu['institution']}")
                set_run(r, size=10, color=DARK)
                if edu.get("dates"):
                    r = p2.add_run(f" — {edu['dates']}")
                    set_run(r, size=10, color=GRAY, italic=True)

    # ===== SELECTED PUBLICATIONS =====
    publications = cv.get("publications", [])
    if publications:
        add_section_heading("Selected Publications")
        for pub in publications:
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(2)
            r = p.add_run(pub)
            set_run(r, size=10, color=DARK)

    # ===== PROJECTS / CERTIFICATIONS =====
    projects = cv.get("projects", [])
    if projects:
        add_section_heading("Selected Projects & Deployments")
        for proj in projects:
            p = doc.add_paragraph(style='List Bullet')
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.left_indent = Cm(0.8)
            p.clear()
            r = p.add_run(proj)
            set_run(r, size=10, color=DARK)

    certifications = cv.get("certifications", [])
    if certifications:
        add_section_heading("Certifications & Training")
        for cert in certifications:
            p = doc.add_paragraph(style='List Bullet')
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.left_indent = Cm(0.8)
            p.clear()
            r = p.add_run(cert)
            set_run(r, size=10, color=DARK)

    # ===== TECHNICAL SKILLS =====
    tech_skills = cv.get("technical_skills", "")
    if tech_skills:
        add_section_heading("Technical Skills")
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(4)
        r = p.add_run(tech_skills)
        set_run(r, size=10, color=DARK)

    # ===== LANGUAGES =====
    languages_list = cv.get("languages", [])
    if languages_list:
        add_section_heading("Languages")
        p = doc.add_paragraph()
        r = p.add_run("  •  ".join(languages_list))
        set_run(r, size=10, color=DARK)

    # Footer
    add_divider(doc)
    footer_p = doc.add_paragraph()
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = footer_p.add_run(f"Generated by {APP_NAME} • {datetime.now().strftime('%B %d, %Y')}")
    set_run(r, size=8, color=GRAY, italic=True)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio


# ===== PREMIUM COVER LETTER DOCUMENT GENERATOR =====
def generate_premium_cover_letter_docx(letter_json_str):
    """Generate a premium cover letter following the Quantium Insights template structure."""
    from docx.shared import Pt, RGBColor, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    try:
        cl = json.loads(letter_json_str)
    except json.JSONDecodeError:
        import re
        match = re.search(r'```(?:json)?\s*([\s\S]*?)```', letter_json_str)
        if match:
            cl = json.loads(match.group(1))
        else:
            cl = json.loads(letter_json_str.strip())

    doc = Document()

    for section in doc.sections:
        section.top_margin = Cm(1.8)
        section.bottom_margin = Cm(1.5)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)

    NAVY = RGBColor(0x0F, 0x2B, 0x4C)
    TEAL_HEX = '14B8A6'
    GRAY = RGBColor(0x4A, 0x55, 0x68)
    DARK = RGBColor(0x2D, 0x2D, 0x2D)

    def set_run(run, size=11, color=DARK, bold=False, italic=False):
        run.font.size = Pt(size)
        run.font.color.rgb = color
        run.bold = bold
        run.italic = italic
        run.font.name = "Georgia"

    def add_teal_divider():
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(16)
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single')
        bottom.set(qn('w:sz'), '8')
        bottom.set(qn('w:color'), TEAL_HEX)
        pBdr.append(bottom)
        pPr.append(pBdr)

    # ===== CENTERED HEADER =====
    name_p = doc.add_paragraph()
    name_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    name_p.paragraph_format.space_after = Pt(2)
    r = name_p.add_run(cl.get("full_name", "CANDIDATE NAME").upper())
    set_run(r, size=16, color=NAVY, bold=True)
    r.font.character_spacing = Pt(2)

    creds = cl.get("credentials", "")
    if creds:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(creds)
        set_run(r, size=9.5, color=GRAY, italic=True)

    contact = cl.get("contact_line", "")
    if contact:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(contact)
        set_run(r, size=9.5, color=GRAY)

    add_teal_divider()

    # ===== DATE =====
    date_str = cl.get("date", datetime.now().strftime("%B %d, %Y"))
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(10)
    r = p.add_run(date_str)
    set_run(r, size=11, color=DARK)

    # ===== ADDRESSEE BLOCK =====
    for line in cl.get("addressee_lines", ["Hiring Committee"]):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(line)
        set_run(r, size=11, color=DARK)

    # Spacing after addressee
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # ===== RE LINE =====
    re_line = cl.get("re_line", "")
    if re_line:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(14)
        r = p.add_run("RE: ")
        set_run(r, size=11, color=NAVY, bold=True)
        r = p.add_run(re_line)
        set_run(r, size=11, color=NAVY, bold=True)

    # ===== SALUTATION =====
    salutation = cl.get("salutation", "Dear Hiring Manager,")
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(10)
    r = p.add_run(salutation)
    set_run(r, size=11, color=DARK)

    # ===== BODY PARAGRAPHS =====
    for para_text in cl.get("body_paragraphs", []):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(10)
        p.paragraph_format.line_spacing = Pt(15)
        r = p.add_run(para_text)
        set_run(r, size=11, color=DARK)

    # ===== CLOSING =====
    closing_text = cl.get("closing_line", "Respectfully submitted,")
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(closing_text)
    set_run(r, size=11, color=DARK)

    doc.add_paragraph()  # Signature space

    sig_name = cl.get("signature_name", cl.get("full_name", ""))
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(sig_name)
    set_run(r, size=11, color=NAVY, bold=True)

    sig_title = cl.get("signature_title", "")
    if sig_title:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(sig_title)
        set_run(r, size=10, color=GRAY)

    sig_contact = cl.get("signature_contact", "")
    if sig_contact:
        p = doc.add_paragraph()
        r = p.add_run(sig_contact)
        set_run(r, size=10, color=GRAY)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio


def initial_populate_rag():
    """Initialize knowledge base with foundational content"""
    foundational_docs = [
        {
            "id": "unicef_youth_employment",
            "text": """UNICEF Youth Employment Guidelines for Africa:
1. Skills Development: Focus on foundational learning and transferable skills
2. Equity and Inclusion: Reach marginalized youth, especially girls
3. Education in Emergencies: Ensure continuity in fragile contexts
4. Cross-sectoral Linkages: Connect education with health and protection
5. System Strengthening: Build resilient education systems
Source: UNICEF Education Strategy 2019-2030""",
            "source": "UNICEF"
        },
        {
            "id": "afdb_sepa_priorities",
            "text": """AfDB Skills for Employability and Productivity in Africa (SEPA) Priorities:
1. STEM Education: Prioritize Science, Technology, Engineering, Mathematics
2. TVET Development: Expand Technical and Vocational Education and Training
3. Digital Skills: Prepare youth for 4th Industrial Revolution (230M digital jobs by 2030)
4. Entrepreneurship: Integrate business skills in education
5. Skills Enhancement Zones: Link training to industrial clusters
Key Statistics: 83% of 18M African youth entering labor market annually remain unemployed or underemployed in informal sector
Source: AfDB SEPA Action Plan 2022-2025""",
            "source": "AfDB"
        },
        {
            "id": "ilo_youth_employment",
            "text": """ILO Global Employment Trends for Youth:
1. Skills Mismatch: 46% of employed African youth report skills mismatch
2. Informal Sector: 95% of young Africans work in informal economy
3. NEET Rates: 20.8% youth Not in Employment, Education or Training (25.9% female)
4. Decent Work: Promote productive employment yielding above poverty-line returns
5. Social Protection: Support workers in informal economy
Source: ILO Global Employment Trends for Youth 2022""",
            "source": "ILO"
        },
        {
            "id": "africa_job_sectors",
            "text": """Priority Employment Sectors for African Youth:
1. Agriculture & Agribusiness: Value chain development, agro-processing
2. Digital Economy: ICT, software development, digital services
3. Renewable Energy: Solar, wind, green technologies
4. Healthcare: Nursing, pharmacy, health technology
5. Manufacturing: Light industry, textiles, processing
6. Creative Industries: Media, arts, entertainment
7. Tourism & Hospitality: Service sector growth
Source: AfDB, ILO, UNICEF frameworks""",
            "source": "Multi-source"
        }
    ]
    
    try:
        vectors_to_upsert = []
        for doc in foundational_docs:
            vec = embeddings.embed_query(doc["text"])
            vectors_to_upsert.append({
                "id": doc["id"],
                "values": vec,
                "metadata": {
                    "text": doc["text"],
                    "source": doc.get("source", "Unknown")
                }
            })
        
        if vectors_to_upsert:
            index.upsert(vectors=vectors_to_upsert)
            return len(vectors_to_upsert)
        return 0
    except Exception as e:
        return 0

# ----- SIDEBAR -----
with st.sidebar:
    st.markdown(f"## {APP_NAME}")
    st.markdown(f"*{TAGLINE}*")
    
    # Admin Access
    st.markdown("---")
    admin_access = st.checkbox("Admin Access")
    
    if admin_access:
        st.markdown("### Admin Login")
        username = st.text_input("Username", key="admin_user")
        password = st.text_input("Password", type="password", key="admin_pass")
        
        if st.button("Login"):
            if username == ADMIN_USERNAME and hashlib.sha256(password.encode()).hexdigest() == ADMIN_PASSWORD_HASH:
                st.session_state['admin_logged_in'] = True
                st.success("Admin access granted.")
                st.rerun()
            else:
                st.error("Invalid credentials.")
    
    st.markdown("---")
    st.caption("Choose your language at the top of the main page.")
    st.markdown("---")
    st.markdown("## Features")
    st.markdown("""
    - Youth-Centric Design
    - ATS-Optimized CV Builder
    - 9 African Languages
    - AI-Powered Guidance
    - Culturally Grounded via RAG
    """)
    
    st.markdown("---")
    st.markdown("## Knowledge Base")

    if st.button("Initialize / Update Knowledge Base"):
        with st.spinner("Loading foundational knowledge..."):
            count = initial_populate_rag()
            if count > 0:
                st.success(f"Loaded {count} foundational documents.")
            else:
                st.info("ℹ️ Knowledge base already initialized")
    
    st.markdown("---")
    st.caption("© Quantium Insights LLC")

# ----- ADMIN DASHBOARD -----
def admin_dashboard():
    st.title("Admin Dashboard")
    
    col1, col2 = st.columns([3, 1])
    with col2:
        if st.button("Logout"):
            st.session_state['admin_logged_in'] = False
            st.rerun()
    
    tabs = st.tabs(["Analytics", "Knowledge Base Management"])
    
    with tabs[0]:
        st.markdown("## Usage Analytics")
        
        analytics = load_analytics()
        if analytics:
            df = pd.DataFrame(analytics)

            st.metric("Total Events", len(analytics))

            event_counts = Counter(df['event'])
            st.markdown("### Event Distribution")
            st.bar_chart(event_counts)

            st.markdown("### Recent Events")
            st.dataframe(df.tail(20))
        else:
            st.info("No analytics data yet")
    
    with tabs[1]:
        st.markdown("## Knowledge Base Management")
        
        # Display current stats
        try:
            stats = index.describe_index_stats()
            total_vectors = stats.get('total_vector_count', 0)
            st.info(f"Current knowledge base: {total_vectors} total chunks from all documents")
            
            # Estimate number of documents (assuming ~100-200 chunks per doc)
            est_docs = max(1, total_vectors // 150)
            st.info(f"Estimated ~{est_docs} documents loaded")
        except:
            st.warning("Unable to fetch knowledge base statistics")
        
        st.markdown("---")
        st.markdown("### Upload New Document")
        st.info("ℹ️ Each upload ADDS to your existing knowledge base (does not replace)")
        
        uploaded_doc = st.file_uploader(
            "Upload Document to Knowledge Base",
            type=['pdf', 'txt', 'docx'],
            help="Limit 25MB per file • TXT, PDF, DOCX"
        )
        
        doc_source = st.text_input(
            "Document Source (e.g., UNICEF, ILO)",
            placeholder="AfDB",
            help="Organization or source of the document"
        )
        
        doc_category = st.text_input(
            "Topic/Category",
            placeholder="Youth Employment Strategy",
            help="Main topic or category"
        )
        
        if st.button("Add to Knowledge Base"):
            if uploaded_doc and doc_source:
                with st.spinner("Processing document..."):
                    try:
                        # Read document based on type
                        if uploaded_doc.name.endswith('.txt'):
                            doc_text = uploaded_doc.read().decode('utf-8')
                        elif uploaded_doc.name.endswith('.pdf'):
                            import PyPDF2
                            pdf_reader = PyPDF2.PdfReader(uploaded_doc)
                            doc_text = ""
                            for page in pdf_reader.pages:
                                doc_text += page.extract_text()
                        elif uploaded_doc.name.endswith('.docx'):
                            from docx import Document
                            doc = Document(uploaded_doc)
                            doc_text = "\n".join([para.text for para in doc.paragraphs])
                        
                        # Split into chunks
                        chunk_size = 500
                        chunks = [doc_text[i:i+chunk_size] for i in range(0, len(doc_text), chunk_size)]
                        
                        # Create unique document ID with timestamp
                        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                        doc_id = f"{doc_source}_{uploaded_doc.name.replace('.pdf', '').replace('.docx', '').replace('.txt', '')}_{timestamp}"
                        
                        # Upsert to Pinecone
                        vectors_to_upsert = []
                        for idx, chunk in enumerate(chunks):
                            if chunk.strip():
                                vec = embeddings.embed_query(chunk)
                                vectors_to_upsert.append({
                                    "id": f"{doc_id}_chunk_{idx}",
                                    "values": vec,
                                    "metadata": {
                                        "text": chunk,
                                        "source": doc_source,
                                        "category": doc_category,
                                        "document_id": doc_id
                                    }
                                })
                        
                        if vectors_to_upsert:
                            index.upsert(vectors=vectors_to_upsert)
                            
                            # Get updated stats
                            new_stats = index.describe_index_stats()
                            new_total = new_stats.get('total_vector_count', 0)
                            
                            st.success(f"Document added — {len(chunks)} chunks indexed.")
                            st.success(f"New total: {new_total} chunks in knowledge base (+{len(chunks)} from this upload).")
                            st.info(f"Document ID: {doc_id}")
                        else:
                            st.error("No valid content to index")
                    
                    except Exception as e:
                        st.error(f"Error: {str(e)}")
            else:
                st.warning("Please upload a document and specify source")

# ----- SECTION 1: RESUME ANALYSIS -----
def resume_analysis_section():
    log_analytics('section_accessed', 'Resume Analysis')
    
    st.markdown("## Professional Resume Analysis")
    st.markdown("Upload your resume to receive expert feedback grounded in African job market context.")

    uploaded_file = st.file_uploader(
        "Upload your resume",
        type=['pdf', 'docx', 'txt'],
        help="Limit 25MB per file • PDF, DOCX, TXT"
    )

    city = st.text_input(
        "Your city (optional)",
        placeholder="e.g., Lagos, Nairobi, Accra",
        help="Helps provide location-specific advice"
    )

    additional_info = st.text_area(
        "Additional information",
        placeholder="Target industry, preferred roles...",
        help="Any additional context that helps us provide better feedback"
    )

    if st.button("Analyze Resume", key="analyze_btn"):
        if uploaded_file:
            log_analytics('resume_upload', f"File: {uploaded_file.name}")

            with st.spinner("Analyzing your resume..."):
                try:
                    # Extract text from resume
                    if uploaded_file.name.endswith('.txt'):
                        resume_text = uploaded_file.read().decode('utf-8')
                    elif uploaded_file.name.endswith('.pdf'):
                        import PyPDF2
                        pdf_reader = PyPDF2.PdfReader(uploaded_file)
                        resume_text = ""
                        for page in pdf_reader.pages:
                            resume_text += page.extract_text()
                    elif uploaded_file.name.endswith('.docx'):
                        doc = Document(uploaded_file)
                        resume_text = "\n".join([para.text for para in doc.paragraphs])
                    
                    # ===== CRITICAL: RETRIEVE RAG CONTEXT =====
                    rag_context = retrieve_career_guidance(f"resume improvement African job market {city if city else 'Africa'}")
                    
                    # ===== USE SAFE LLM CALL WITH GUARDRAILS =====
                    analysis_prompt = f"""Analyze this resume for the African job market.

Resume Content:
{resume_text[:12000]}

Location Context: {city if city else 'General African market'}
Additional Info: {additional_info if additional_info else 'None'}

Provide:
1. ATS Compatibility Score (1-100)
2. Top 3 Strengths
3. Top 5 Areas for Improvement
4. African Market Relevance Assessment
5. 3 Actionable Next Steps"""

                    feedback = safe_llm_call(analysis_prompt, rag_context, target_lang)
                    
                    # Store in session state for downstream generation
                    st.session_state['resume_text'] = resume_text
                    st.session_state['resume_feedback'] = feedback
                    st.session_state['resume_city'] = city
                    st.session_state['resume_additional'] = additional_info
                    st.session_state['analysis_done'] = True
                    
                    st.success("Analysis complete.")
                    st.markdown("### Your Resume Analysis")
                    st.markdown(feedback)
                    
                    log_analytics('analysis_completed')
                
                except Exception as e:
                    st.error(f"Error: {str(e)}")
        else:
            st.warning("Please upload a resume first.")

    # ===== POST-ANALYSIS: PREMIUM CV & COVER LETTER GENERATION =====
    if st.session_state.get('analysis_done', False):
        st.markdown("---")
        st.markdown("### Generate Premium Documents")
        st.markdown("Based on your resume analysis, generate a polished, ATS-optimized CV and/or a professional cover letter.")
        
        # Target job details for cover letter
        col_job1, col_job2 = st.columns(2)
        with col_job1:
            target_position = st.text_input(
                "Target position (for cover letter)",
                placeholder="e.g., Data Analyst, Project Manager, Nurse",
                key="target_position_input"
            )
        with col_job2:
            target_company = st.text_input(
                "Target company / organization",
                placeholder="e.g., WHO, Dangote Group, Safaricom",
                key="target_company_input"
            )
        
        col_cv, col_cl = st.columns(2)
        
        # ===== GENERATE UPDATED CV =====
        with col_cv:
            if st.button("Generate Updated CV", key="gen_cv_btn"):
                resume_text = st.session_state.get('resume_text', '').strip()
                feedback = st.session_state.get('resume_feedback', '').strip()
                
                if resume_text:
                    log_analytics('premium_cv_generation')
                    with st.spinner("Generating your premium ATS-optimized CV..."):
                        try:
                            rag_context_cv = retrieve_career_guidance("professional CV resume best practices African job market ATS optimization")
                            
                            cv_gen_prompt = f"""You are a professional CV writer specializing in ATS-optimized resumes for the African job market.

ORIGINAL RESUME (this is the ground truth — use ONLY what appears here):
{resume_text[:16000]}

ANALYSIS FEEDBACK (apply wording/structure improvements ONLY — do NOT add new facts):
{feedback[:2000]}

Rewrite the resume into an improved, premium, ATS-optimized CV.

RESPOND ONLY WITH VALID JSON (no markdown, no code blocks, no preamble). Use this exact structure:
{{
  "full_name": "full name exactly as in the resume",
  "credentials": "degree abbreviations that appear in the resume (e.g., Ph.D., M.P.H., B.Sc.)",
  "contact_line": "email | phone | city, country | LinkedIn — only the parts present in the resume",
  "professional_summary": "3-4 sentence summary built ONLY from resume facts, tailored for the African market",
  "core_competencies": ["8-12 ATS keyword-rich skills that are supported by the resume"],
  "work_experience": [
    {{
      "title": "Job title from resume",
      "company": "Company/Organization from resume",
      "location": "City, Country if in resume, else omit this key",
      "dates": "dates if in resume, else omit this key",
      "bullets": ["3-5 achievement bullets rewritten from what the resume actually says"]
    }}
  ],
  "education": [
    {{"degree": "Degree, Field from resume", "institution": "Institution from resume", "dates": "Year if in resume, else omit this key"}}
  ],
  "publications": ["only publications that appear in the resume; else empty list"],
  "projects": ["only projects that appear in the resume; else empty list"],
  "certifications": ["only certifications that appear in the resume; else empty list"],
  "technical_skills": "only tools/software mentioned in the resume, comma-separated",
  "languages": ["ONLY languages explicitly stated in the resume; return an empty list [] if none are stated"]
}}

CRITICAL RULES (accuracy matters more than polish):
- Include EVERY employer and role that appears in the resume — do NOT drop any position (all organizations such as HJF, USAID, UNICEF, etc. must appear if they are in the resume).
- Do NOT invent or add anything absent from the resume: no new employers, dates, metrics, degrees, publications, tools, or languages. If the resume does not list a language, the languages array MUST be empty.
- You may rephrase and quantify ONLY using numbers already present in the resume.
- Return ONLY the JSON object, nothing else."""

                            cv_json = safe_llm_call(cv_gen_prompt, rag_context_cv, "English")

                            if cv_json.startswith("⚠️"):
                                st.warning(cv_json)
                            else:
                                cv_docx = generate_premium_cv_docx(cv_json)
                                st.session_state['cv_docx'] = cv_docx
                                st.session_state['cv_generated'] = True
                                st.success("Premium CV generated.")
                        
                        except Exception as e:
                            st.error(f"CV generation error: {str(e)}")
                            st.info("Tip: try again — occasionally a second attempt is needed for complex resumes.")
        
        # ===== GENERATE COVER LETTER =====
        with col_cl:
            if st.button("Generate Cover Letter", key="gen_cl_btn"):
                resume_text = st.session_state.get('resume_text', '')
                # Read directly from session state widget keys to avoid column scoping issues
                cl_position = st.session_state.get('target_position_input', '').strip()
                cl_company = st.session_state.get('target_company_input', '').strip()
                
                if resume_text and cl_position and cl_company:
                    log_analytics('premium_cover_letter_generation', f"{cl_position} at {cl_company}")
                    with st.spinner("Crafting your premium cover letter..."):
                        try:
                            rag_context_cl = retrieve_career_guidance(f"cover letter professional {cl_position} {cl_company} African job market")

                            # Deep research on the target organization (requires TAVILY_API_KEY)
                            org_research = web_research(
                                f"{cl_company} company mission, values, products, and recent priorities relevant to a {cl_position} role"
                            )

                            cl_gen_prompt = f"""You are a professional cover letter writer. Generate a premium, compelling, tailored cover letter.

CANDIDATE'S RESUME (ground truth — use only what appears here):
{resume_text[:16000]}

TARGET POSITION: {cl_position}
TARGET COMPANY: {cl_company}
LOCATION CONTEXT: {st.session_state.get('resume_city', 'Africa')}

ORGANIZATION RESEARCH (verified web results about the company; use these specifics to tailor the letter — do NOT invent facts beyond this):
{org_research if org_research else "(No live research available — rely on widely-known general facts about the company and do not fabricate specifics.)"}

RESPOND ONLY WITH VALID JSON (no markdown, no code blocks, no preamble). Use this exact structure:
{{
  "full_name": "FULL NAME FROM RESUME",
  "credentials": "Degree abbreviations from resume",
  "contact_line": "email | phone | city (from resume)",
  "date": "{datetime.now().strftime('%B %d, %Y')}",
  "addressee_lines": ["Hiring Committee", "{cl_company}"],
  "re_line": "{cl_position} Position",
  "salutation": "Dear Hiring Manager,",
  "body_paragraphs": [
    "Opening paragraph: Express interest, mention the specific role, and give a 1-sentence summary of why you are an ideal fit.",
    "Paragraph 2: Highlight most relevant experience from resume that directly maps to the target role. Use specific achievements and metrics from the resume.",
    "Paragraph 3: Highlight additional relevant skills, technical capabilities, or domain expertise. Reference specific projects, tools, or publications if applicable.",
    "Paragraph 4: Show knowledge of the target company/organization and explain why you are drawn to their mission. Connect your values/goals to theirs.",
    "Closing paragraph: Reaffirm interest, mention availability, and invite further discussion."
  ],
  "closing_line": "Respectfully submitted,",
  "signature_name": "Full Name with credentials",
  "signature_title": "Current Title | Current Organization",
  "signature_contact": "City | email | phone"
}}

RULES:
- Use ONLY factual details from the resume for the candidate's experience — do NOT invent achievements, employers, or metrics.
- Make the company paragraph SPECIFIC using the ORGANIZATION RESEARCH above (reference a real product, mission point, or recent priority). If no research is available, keep company statements general and do not fabricate specifics.
- Each body paragraph should be 3-5 sentences, substantive and specific.
- Use a confident, professional tone; highlight quantified achievements only where the resume provides the numbers.
- Return ONLY the JSON object, nothing else"""

                            cl_json = safe_llm_call(cl_gen_prompt, rag_context_cl, "English")

                            if cl_json.startswith("⚠️"):
                                st.warning(cl_json)
                            else:
                                cl_docx = generate_premium_cover_letter_docx(cl_json)
                                st.session_state['cl_docx'] = cl_docx
                                st.session_state['cl_generated'] = True
                                st.success("Premium cover letter generated.")

                        except Exception as e:
                            st.error(f"Cover letter generation error: {str(e)}")
                            st.info("Tip: try again — occasionally a second attempt is needed.")

                elif not cl_position or not cl_company:
                    st.warning("Please enter the target position and company above to generate a cover letter.")
        
        # ===== DOWNLOAD BUTTONS =====
        st.markdown("---")
        dl_col1, dl_col2 = st.columns(2)
        
        with dl_col1:
            if st.session_state.get('cv_generated', False):
                cv_data = st.session_state['cv_docx']
                cv_data.seek(0)
                st.download_button(
                    label="Download Updated CV (.docx)",
                    data=cv_data,
                    file_name=f"AfriCareer_Premium_CV_{datetime.now().strftime('%Y%m%d')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="dl_cv_btn"
                )
        
        with dl_col2:
            if st.session_state.get('cl_generated', False):
                cl_data = st.session_state['cl_docx']
                cl_data.seek(0)
                st.download_button(
                    label="Download Cover Letter (.docx)",
                    data=cl_data,
                    file_name=f"AfriCareer_CoverLetter_{datetime.now().strftime('%Y%m%d')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="dl_cl_btn"
                )

# ----- SECTION 2: CAREER GUIDANCE & CV BUILDER -----
def career_guidance_section():
    log_analytics('section_accessed', 'Career Guidance')

    st.markdown("## Career Guidance & CV Builder")
    st.markdown("For youth and job seekers: get personalized guidance and build a professional CV from scratch.")

    with st.expander("Contact details (used on your CV)"):
        cc1, cc2 = st.columns(2)
        with cc1:
            cv_name = st.text_input("Full name", key="cg_name", placeholder="e.g., Amina Bello")
            cv_email = st.text_input("Email", key="cg_email", placeholder="e.g., amina@email.com")
            cv_phone = st.text_input("Phone", key="cg_phone", placeholder="e.g., +234 800 000 0000")
        with cc2:
            cv_city = st.text_input("City, Country", key="cg_city", placeholder="e.g., Kano, Nigeria")
            cv_linkedin = st.text_input("LinkedIn / Portfolio (optional)", key="cg_linkedin", placeholder="e.g., linkedin.com/in/aminabello")

    questions_text = """
**Tell us about yourself**

Please answer these 5 simple questions (number your answers 1-5):

1. **What are you interested in?**
   (Example: I like computers, helping people, cooking, fixing things, etc.)

2. **What are you good at? What skills do you have?**
   (Example: I'm good at math, I can speak 3 languages, I know how to use Excel, etc.)

3. **What work or experience do you have?**
   (Include ANY experience: part-time jobs, helping a family business, volunteer work, school projects, etc.)

4. **What is your education?**
   (Example: I finished secondary school in 2020, I'm studying at university, I completed a training course, etc.)

5. **What job do you want? What are your goals?**
   (Example: I want to work in a bank, I want to be a nurse, I want to start my own business, etc.)

Write your answers below. Be honest — there are no wrong answers.
"""
    st.markdown(questions_text)

    user_answers = st.text_area(
        "Your answers (number them 1-5)",
        height=250,
        placeholder="1. I'm interested in...\n2. My strengths...\n3. I have experience...\n4. My education...\n5. My goals...",
        key="answers_section2"
    )

    col1, col2 = st.columns(2)

    with col1:
        if st.button("Get Career Guidance", key="advice_btn"):
            if user_answers.strip():
                log_analytics('query', f"Career guidance: {user_answers[:50]}...")
                with st.spinner("Preparing your guidance..."):
                    rag_context_advice = retrieve_career_guidance(
                        f"career paths employment opportunities skills development Africa {user_answers[:100]}"
                    )
                    advice_prompt = f"""Provide career guidance for African youth based on their profile.

Their Answers: {user_answers}

Provide:
1. Top 3 Career Paths in Africa (with specific reasons why they match)
2. Key Skills to Develop (7-10 skills with brief explanations)
3. Action Plan (5 concrete steps they can take now)

Ground advice in African job market realities and cite relevant frameworks when applicable."""
                    st.session_state['cg_guidance'] = safe_llm_call(advice_prompt, rag_context_advice, target_lang)
            else:
                st.warning("Please answer the questions first.")

    with col2:
        if st.button("Generate Premium CV", key="cv_btn"):
            if user_answers.strip():
                log_analytics('cv_generation', f"User: {user_answers[:50]}...")
                with st.spinner("Building your premium ATS-optimized CV..."):
                    try:
                        rag_context_cv = retrieve_career_guidance(
                            "professional CV resume best practices African job market ATS optimization"
                        )
                        contact_bits = [b for b in [
                            cv_email.strip(), cv_phone.strip(), cv_city.strip(), cv_linkedin.strip()
                        ] if b]
                        contact_line = " | ".join(contact_bits)
                        full_name = cv_name.strip()

                        cv_gen_prompt = f"""You are a professional CV writer creating an ATS-optimized CV for the African job market, based on a young jobseeker's answers to 5 questions.

CANDIDATE ANSWERS:
{user_answers}

CANDIDATE CONTACT (use verbatim; do not invent):
- Full name: {full_name if full_name else "(not provided)"}
- Contact line: {contact_line if contact_line else "(not provided)"}

RESPOND ONLY WITH VALID JSON (no markdown, no code blocks, no preamble). Use this exact structure:
{{
  "full_name": "the candidate's full name, or empty string if not provided",
  "credentials": "degree abbreviations if clearly stated (e.g., B.Sc.), else empty string",
  "contact_line": "the contact line provided, or empty string",
  "professional_summary": "3-4 sentence summary based ONLY on what the candidate said, tailored to the African market",
  "core_competencies": ["6-10 ATS keyword-rich skills drawn from their answers"],
  "work_experience": [
    {{
      "title": "role or what they did",
      "company": "organization ONLY if they named one, else omit this key",
      "location": "ONLY if they gave one, else omit this key",
      "dates": "ONLY if they gave dates, else omit this key",
      "bullets": ["2-4 achievement-oriented bullets based on what they actually described"]
    }}
  ],
  "education": [
    {{
      "degree": "their education level or qualification as they described it",
      "institution": "ONLY if they named one, else omit this key",
      "dates": "ONLY if they gave a year, else omit this key"
    }}
  ],
  "certifications": ["ONLY if they mentioned any; else empty list"],
  "technical_skills": "comma-separated tools/software they mentioned, or empty string",
  "languages": ["languages they mentioned with proficiency, or empty list"]
}}

CRITICAL RULES:
- Use ONLY facts the candidate provided. Do NOT invent employers, job titles, dates, metrics, or degrees.
- NEVER output placeholder text such as "[Your Name]", "[Start Date]", "[City]", or "[Company]". If a detail is unknown, OMIT that key entirely (or use an empty string/list where the schema requires the key).
- For informal experience (family business, volunteering, school projects), represent it honestly and professionally.
- Return ONLY the JSON object, nothing else."""

                        cv_json = safe_llm_call(cv_gen_prompt, rag_context_cv, "English")
                        if cv_json.startswith("⚠️"):
                            st.warning(cv_json)
                        else:
                            st.session_state['cg_cv_docx'] = generate_premium_cv_docx(cv_json)
                            st.session_state['cg_cv_generated'] = True
                            st.success("Premium CV generated.")
                    except Exception as e:
                        st.error(f"CV generation error: {str(e)}")
                        st.info("Tip: add a few more details to your answers and try again.")
            else:
                st.warning("Please answer the questions first.")

    if st.session_state.get('cg_guidance'):
        st.markdown("---")
        st.markdown("### Your Career Roadmap")
        st.markdown(st.session_state['cg_guidance'])

    if st.session_state.get('cg_cv_generated'):
        st.markdown("---")
        cv_data = st.session_state['cg_cv_docx']
        cv_data.seek(0)
        st.download_button(
            label="Download Premium CV (.docx)",
            data=cv_data,
            file_name=f"AfriCareer_Premium_CV_{datetime.now().strftime('%Y%m%d')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="cg_dl_cv_btn"
        )

# ----- LINK VERIFICATION (guarantees no hallucinated / dead course links) -----
TAVILY_API_KEY = os.getenv("TAVILY_API_KEY", "").strip()
_BROWSER_UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
               "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")

# Reputable providers -> search deep-link templates. A search query cannot 404 to a fake course.
_PROVIDER_SEARCH = {
    "coursera": "https://www.coursera.org/search?query={q}",
    "edx": "https://www.edx.org/search?q={q}",
    "udemy": "https://www.udemy.com/courses/search/?q={q}",
    "udacity": "https://www.udacity.com/catalog?searchValue={q}",
    "class central": "https://www.classcentral.com/search?q={q}",
    "classcentral": "https://www.classcentral.com/search?q={q}",
    "freecodecamp": "https://www.freecodecamp.org/news/search/?query={q}",
    "khan academy": "https://www.khanacademy.org/search?page_search_query={q}",
    "khanacademy": "https://www.khanacademy.org/search?page_search_query={q}",
    "linkedin learning": "https://www.linkedin.com/learning/search?keywords={q}",
    "youtube": "https://www.youtube.com/results?search_query={q}",
    "alison": "https://alison.com/courses?query={q}",
    "futurelearn": "https://www.futurelearn.com/search?q={q}",
}

def provider_search_url(provider: str, query: str) -> str:
    """Build a real search deep-link on the given provider (defaults to Class Central)."""
    q = quote_plus((query or "").strip())
    p = (provider or "").lower()
    for key, tmpl in _PROVIDER_SEARCH.items():
        if key in p:
            return tmpl.format(q=q)
    return f"https://www.classcentral.com/search?q={q}"

# Provider-grounded cost classification (don't trust the model's self-reported cost)
_PAID_PROVIDERS = ("udemy", "udacity", "linkedin learning")
_FREE_PROVIDERS = ("freecodecamp", "khan academy", "khanacademy", "youtube",
                   "mit opencourseware", "ocw", "alison")

def classify_cost(provider: str, llm_cost: str) -> str:
    """Return 'Free' or 'Paid', grounded by provider first, then the model's label.
    Free = the content can be accessed at no cost (Coursera/edX audit free; certificate may cost)."""
    p = (provider or "").lower()
    if any(k in p for k in _PAID_PROVIDERS):
        return "Paid"
    if any(k in p for k in _FREE_PROVIDERS):
        return "Free"
    lc = (llm_cost or "").lower()
    if "paid" in lc and "free" not in lc:
        return "Paid"
    return "Free"  # Coursera/edX audit-free; Class Central lists free options

def cost_matches(cost: str, pref: str) -> bool:
    """Enforce the user's cost preference in code (the model is unreliable at this)."""
    if pref.startswith("Free &"):   # both
        return True
    if pref.startswith("Free"):
        return cost == "Free"
    if pref.startswith("Paid"):
        return cost == "Paid"
    return True

@st.cache_data(ttl=3600, show_spinner=False)
def verify_url(url: str, timeout: float = 6.0) -> bool:
    """Return True if the URL resolves for a real user.

    2xx/3xx = OK. Anti-bot blocks (401/403/405/429/999) also count as OK: the page
    exists for real browsers even though it refuses automated requests (Class Central,
    Udemy, LinkedIn do this). Only clear failures — 404/410/5xx, timeouts, or DNS/
    connection errors — return False.
    """
    if not url or not url.startswith(("http://", "https://")):
        return False
    try:
        with httpx.Client(follow_redirects=True, timeout=timeout,
                          headers={"User-Agent": _BROWSER_UA}) as client:
            code = client.get(url).status_code
        return code < 400 or code in (401, 403, 405, 429, 999)
    except Exception:
        return False

@st.cache_data(ttl=1800, show_spinner=False)
def web_search_links(query: str, max_results: int = 4):
    """Real course URLs from Tavily web search if TAVILY_API_KEY is set, else []."""
    if not TAVILY_API_KEY:
        return []
    try:
        with httpx.Client(timeout=12.0) as client:
            resp = client.post(
                "https://api.tavily.com/search",
                json={"api_key": TAVILY_API_KEY, "query": query,
                      "max_results": max_results, "search_depth": "basic"},
            )
            return [{"title": r.get("title", r.get("url", "")), "url": r.get("url", "")}
                    for r in resp.json().get("results", []) if r.get("url")]
    except Exception:
        return []

@st.cache_data(ttl=1800, show_spinner=False)
def web_research(query: str, max_results: int = 5) -> str:
    """Return a short grounded research brief (answer + result snippets) from Tavily, or ''."""
    if not TAVILY_API_KEY:
        return ""
    try:
        with httpx.Client(timeout=15.0) as client:
            resp = client.post(
                "https://api.tavily.com/search",
                json={"api_key": TAVILY_API_KEY, "query": query,
                      "max_results": max_results, "search_depth": "advanced",
                      "include_answer": True},
            )
            data = resp.json()
        parts = []
        if data.get("answer"):
            parts.append("Summary: " + data["answer"])
        for r in data.get("results", []):
            c = (r.get("content") or "").strip()
            if c:
                parts.append(f"- {r.get('title', '')}: {c[:400]}")
        return "\n".join(parts)[:4000]
    except Exception:
        return ""

# ----- SECTION 3: LEARNING RESOURCES -----
def learning_resources_section():
    log_analytics('section_accessed', 'Learning Resources')

    st.markdown("## Learning Resources & Courses")
    st.markdown(
        "Get course recommendations matched to your goals. Every link below is built by the app "
        "and checked live before it is shown — so you never get a broken or made-up link."
    )

    learning_interest = st.text_area(
        "What would you like to learn?",
        placeholder="e.g., data analysis, solar installation, digital marketing, tailoring & small business skills",
        help="The more specific, the better the recommendations."
    )

    col1, col2 = st.columns(2)
    with col1:
        course_type = st.selectbox("Cost preference", ["Free & Paid", "Free only", "Paid only"])
    with col2:
        skill_level = st.selectbox("Your level", ["Beginner", "Intermediate", "Advanced"])

    if st.button("Find Courses", key="course_search_btn"):
        if not learning_interest.strip():
            st.warning("Please describe what you want to learn.")
            return

        log_analytics('course_search', learning_interest)
        with st.spinner("Finding and verifying courses..."):
            rag_context_learning = retrieve_career_guidance(
                f"skills development training courses {learning_interest} African youth"
            )
            # Preference-specific guidance so the model biases the right providers
            if course_type.startswith("Free"):
                cost_guidance = ("Recommend ONLY courses whose content is accessible at NO cost. Strongly prefer "
                                 "freeCodeCamp, Khan Academy, YouTube, Alison, MIT OpenCourseWare, Class Central, "
                                 "and free-to-audit Coursera/edX courses. Do NOT include Udemy, Udacity, or LinkedIn Learning.")
            elif course_type.startswith("Paid"):
                cost_guidance = ("Recommend paid courses and certifications. Prefer Udemy, Udacity, LinkedIn Learning, "
                                 "and paid Coursera/edX certificates or specializations.")
            else:
                cost_guidance = "Include a healthy mix of free and paid options."

            course_prompt = f"""Recommend 8 high-quality, real learning resources for an African youth who wants to learn: {learning_interest}
Level: {skill_level}

COST REQUIREMENT: {cost_guidance}
Definition of "Free" = the learning content can be accessed at no cost (free-to-audit counts as Free even if an optional certificate costs extra).

Return ONLY valid JSON (no markdown, no code fences) - an array of objects with this exact shape:
[
  {{
    "title": "specific, real course or specialization title",
    "provider": "one of: Coursera, edX, Udemy, Udacity, Class Central, freeCodeCamp, Khan Academy, LinkedIn Learning, YouTube, Alison, FutureLearn, MIT OpenCourseWare",
    "cost": "Free | Paid",
    "level": "Beginner | Intermediate | Advanced",
    "duration": "approx. time commitment, e.g. '3 months'",
    "why": "one sentence on why it fits this learner and the African job market"
  }}
]

RULES:
- Do NOT include any URLs or links - the app builds and verifies links itself.
- "cost" must be exactly "Free" or "Paid" per the definition above.
- Prefer well-known, currently-offered courses.
- Return ONLY the JSON array, nothing else."""

            raw = safe_llm_call(course_prompt, rag_context_learning, "English")

            if raw.startswith("⚠️"):
                st.warning(raw)
                return

            recs = []
            try:
                recs = json.loads(raw)
            except json.JSONDecodeError:
                import re
                m = re.search(r'\[[\s\S]*\]', raw)
                if m:
                    try:
                        recs = json.loads(m.group(0))
                    except Exception:
                        recs = []

            # Ground the cost label by provider, then ENFORCE the user's preference in code
            items = []
            if isinstance(recs, list):
                for rec in recs:
                    if not isinstance(rec, dict):
                        continue
                    title = str(rec.get("title", "")).strip()
                    if not title:
                        continue
                    provider = str(rec.get("provider", "")).strip()
                    cost = classify_cost(provider, str(rec.get("cost", "")))
                    if not cost_matches(cost, course_type):
                        continue
                    items.append({
                        "title": title, "provider": provider, "cost": cost,
                        "level": str(rec.get("level", "")).strip(),
                        "duration": str(rec.get("duration", "")).strip(),
                        "why": str(rec.get("why", "")).strip(),
                    })
            items = items[:6]

            if items:
                st.markdown(f"### Recommended Courses ({course_type})")
                st.caption('Cost is checked against the provider; "Free" means the content is accessible at no cost '
                           '(Coursera/edX can be audited free — a certificate may cost extra). Links are verified live.')
                for it in items:
                    url = provider_search_url(it["provider"], it["title"])
                    ok = verify_url(url)
                    label_provider = it["provider"] or "Class Central"
                    if not ok:
                        url = provider_search_url("class central", it["title"])
                        ok = verify_url(url)
                        label_provider = "Class Central"
                    meta = " · ".join([x for x in [it["provider"], it["cost"], it["level"], it["duration"]] if x])
                    st.markdown(f"**{it['title']}**")
                    if meta:
                        st.caption(meta)
                    if it["why"]:
                        st.markdown(it["why"])
                    if ok:
                        st.markdown(f"[Find this course on {label_provider} →]({url})")
                    else:
                        st.caption("Live link check failed — search this title on classcentral.com.")
                    st.markdown("")
            else:
                st.info(f"No strictly {course_type.lower()} matches came back for that topic. "
                        "Try 'Free & Paid', or a broader topic.")

            # Live web search, tuned to the chosen cost (requires TAVILY_API_KEY)
            if TAVILY_API_KEY:
                if course_type.startswith("Free"):
                    q = f"free {learning_interest} online course {skill_level}"
                    heading = "Verified Free Courses (live web search)"
                elif course_type.startswith("Paid"):
                    q = f"{learning_interest} paid certification course {skill_level}"
                    heading = "Verified Paid Courses (live web search)"
                else:
                    q = f"{learning_interest} online course {skill_level}"
                    heading = "Verified Courses (live web search)"
                verified_live = [l for l in web_search_links(q) if verify_url(l["url"])]
                if verified_live:
                    st.markdown(f"### {heading}")
                    for l in verified_live:
                        st.markdown(f"- [{l['title']}]({l['url']})")
            else:
                st.info("Tip: add a TAVILY_API_KEY secret to also pull real, verified course links "
                        "from a live web search tuned to your Free/Paid choice.")

# ----- SECTION 4: AI ASSISTANT -----
def ai_assistant_section():
    log_analytics('section_accessed', 'AI Assistant')
    
    st.markdown("## AI Career Assistant")
    st.markdown("Ask any career-related question and get personalized guidance.")

    st.info("This assistant focuses on careers, education, job searching, and professional development, grounded in AfDB, UNICEF, and ILO frameworks for African youth.")

    question = st.text_area(
        "Ask your question",
        placeholder="What skills should African youth prioritize according to international development frameworks for employability?",
        help="Ask about career paths, skills, education, job search, etc.",
        height=120
    )

    if st.button("Ask AI Assistant", key="ask_ai_btn"):
        if question:
            log_analytics('query', f"AI Assistant: {question}")

            with st.spinner("Thinking..."):
                rag_context = retrieve_career_guidance(question)
                response_content = safe_llm_call(question, rag_context, target_lang)

                st.markdown("### Response")
                st.markdown(response_content)
        else:
            st.warning("Please enter a question above before clicking 'Ask AI Assistant'.")

# ----- LANDING PAGE / LOGIN -----
def render_landing():
    """Show the landing page + sign-in. Returns True once the user has entered."""
    if st.session_state.get("access_granted"):
        return True

    st.markdown(f"""
    <div class="africareer-hero">
        <h1>{APP_NAME}</h1>
        <p>{TAGLINE}. Build an ATS-ready CV, get personalized career guidance, and find
        verified courses — free, multilingual, and made for African youth.</p>
        <div class="hero-chips">
            <span class="hero-chip">ATS-optimized CVs</span>
            <span class="hero-chip">9 African languages</span>
            <span class="hero-chip">Verified learning links</span>
            <span class="hero-chip">24/7 AI guidance</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("")
    _, mid, _ = st.columns([1, 1.5, 1])
    with mid:
        st.markdown("#### Sign in to continue")
        name = st.text_input("Your name", key="login_name", placeholder="e.g., Amina Bello")
        code = ""
        if APP_ACCESS_CODE:
            code = st.text_input("Access code", type="password", key="login_code",
                                 placeholder="Enter your pilot access code")
        if st.button("Enter AfriCareer AI", key="login_btn"):
            if APP_ACCESS_CODE and code.strip() != APP_ACCESS_CODE:
                st.error("Invalid access code. Please contact the AfriCareer AI team.")
            elif not name.strip():
                st.warning("Please enter your name to continue.")
            else:
                st.session_state["access_granted"] = True
                st.session_state["user_name"] = name.strip()
                log_analytics("login", name.strip())
                st.rerun()
        st.caption("Free to use. Your details are only used to personalize your experience.")

    return False

# ----- MAIN APP -----
def main():
    global target_lang

    if st.session_state.get('admin_logged_in', False):
        admin_dashboard()
        return

    if not render_landing():
        return

    if 'user_logged' not in st.session_state:
        log_analytics('user_visit')
        st.session_state['user_logged'] = True
    
    st.title(APP_NAME)
    st.markdown(f"### *{TAGLINE}*")
    st.markdown("**Developed by:** Quantium Insights LLC")

    # Language selector at the TOP — visible immediately on mobile, no sidebar drawer needed
    lang_col, _ = st.columns([1, 2])
    with lang_col:
        selected_display = st.selectbox(
            "🌐 Language / Langue / لغة / Harshe",
            list(LANGUAGES.keys()),
            index=0,
            key="language_selector",
        )
    target_lang = LANGUAGES[selected_display]

    tabs = st.tabs([
        "About",
        "Career Guidance",
        "Learning Resources",
        "AI Assistant",
        "Resume Analysis",
    ])

    with tabs[0]:
        st.markdown("""
        ## About AfriCareer AI

        **AfriCareer AI** is an AI-powered career guidance platform for African youth.

        ### Mission
        Empower urban African youth with professional, accessible career services.

        ### Key Features
        - **9 Languages:** English, French, Swahili, Arabic, Hausa, Pidgin, Portuguese, Spanish, Amharic
        - **ATS-Optimized CV Builder:** Clean, recruiter-ready formatting
        - **Cultural Grounding:** RAG grounded in UNICEF, ILO, and AfDB frameworks
        - **Youth-Centric:** Designed for adolescents and young professionals
        - **Safety Guardrails:** Focused, appropriate career guidance

        ### Knowledge Base
        Our AI is grounded in authoritative frameworks:
        - **AfDB SEPA (2022–2025):** Skills for Employability and Productivity in Africa
        - **UNICEF Education Strategy (2019–2030):** Every Child Learns
        - **ILO Global Employment Trends for Youth (2022):** Investing in Transforming Futures

        ### Developer
        **Dr. Amobi Andrew Onovo**
        - PhD Global Health, MPH, PGDip Data Science
        - Global health & data science specialist, Nigeria

        ### Safety & Ethics
        AfriCareer AI includes guardrails to ensure:
        - Mission-aligned responses (career guidance)
        - No inappropriate or harmful content
        - Culturally appropriate advice for the African context
        - Evidence-based recommendations from trusted sources
        """)

    with tabs[1]:
        career_guidance_section()

    with tabs[2]:
        learning_resources_section()

    with tabs[3]:
        ai_assistant_section()

    with tabs[4]:
        resume_analysis_section()

if __name__ == "__main__":
    main()
