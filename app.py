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

# FIXED CSS - LIGHT BACKGROUND, DARK TEXT (READABLE!)
st.markdown("""
<style>
    /* LIGHT BACKGROUNDS - EASY TO READ */
    .main {
        background: linear-gradient(135deg, #e8f5e9 0%, #f1f8f4 100%);
    }
    
    .stApp {
        background: #ffffff;
    }
    
    /* DARK TEXT ON LIGHT - READABLE! */
    p, div, span, li, label, .stMarkdown {
        color: #1a1a1a !important;
    }
    
    /* Headers - Dark Green */
    h1, h2, h3 {
        color: #1a4d2e !important;
        font-weight: 700;
    }
    
    /* Sidebar - Dark Green BG, White Text */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #1a4d2e 0%, #2d7a3e 100%);
    }
    
    [data-testid="stSidebar"] * {
        color: white !important;
    }
    
    /* Button styling */
    .stButton>button {
        background: linear-gradient(135deg, #2d7a3e 0%, #1a4d2e 100%);
        color: white !important;
        border: none;
        border-radius: 25px;
        padding: 12px 30px;
        font-weight: 600;
        box-shadow: 0 4px 15px rgba(45, 122, 62, 0.4);
        transition: all 0.3s ease;
    }
    
    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(45, 122, 62, 0.6);
    }
    
    /* Tab styling - YELLOW TEXT FOR CLARITY */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: transparent;
    }
    
    .stTabs [data-baseweb="tab"] {
        background: linear-gradient(135deg, #2d7a3e 0%, #1a4d2e 100%);
        color: #FFD700 !important;
        border-radius: 15px 15px 0 0;
        padding: 12px 24px;
        font-weight: 600;
    }
    
    /* Target all text inside tabs - FORCE YELLOW */
    .stTabs [data-baseweb="tab"] p,
    .stTabs [data-baseweb="tab"] span,
    .stTabs [data-baseweb="tab"] div {
        color: #FFD700 !important;
    }
    
    /* Active tab */
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #FFD700 0%, #FFA500 100%) !important;
        color: #1a4d2e !important;
    }
    
    /* Active tab text - DARK GREEN */
    .stTabs [aria-selected="true"] p,
    .stTabs [aria-selected="true"] span,
    .stTabs [aria-selected="true"] div {
        color: #1a4d2e !important;
        font-weight: 700;
    }
    
    /* Info boxes */
    .stAlert {
        border-radius: 15px;
        border-left: 5px solid #2d7a3e;
        background: white;
    }
    
    /* Download button */
    .stDownloadButton>button {
        background: linear-gradient(135deg, #FFD700 0%, #FFA500 100%) !important;
        color: #1a4d2e !important;
        border-radius: 25px;
        padding: 12px 30px;
        font-weight: 700;
        box-shadow: 0 4px 15px rgba(255, 215, 0, 0.4);
    }

    /* Never allow the page itself to scroll sideways on any device */
    .stApp { overflow-x: hidden; }

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
        st.error("⚠️ Missing API keys! Check your .env file.")
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
    st.markdown(f"## 🌍 {APP_NAME}")
    st.markdown(f"*{TAGLINE}*")
    
    # Admin Access
    st.markdown("---")
    admin_access = st.checkbox("🔐 Admin Access")
    
    if admin_access:
        st.markdown("### Admin Login")
        username = st.text_input("Username", key="admin_user")
        password = st.text_input("Password", type="password", key="admin_pass")
        
        if st.button("Login"):
            if username == ADMIN_USERNAME and hashlib.sha256(password.encode()).hexdigest() == ADMIN_PASSWORD_HASH:
                st.session_state['admin_logged_in'] = True
                st.success("✅ Admin access granted!")
                st.rerun()
            else:
                st.error("❌ Invalid credentials")
    
    st.markdown("---")
    st.caption("🌐 Choose your language at the top of the main page ☝️")
    st.markdown("---")
    st.markdown("## ✨ Features")
    st.markdown("""
    - 🎯 Youth-Centric Design
    - 📄 ATS-Optimized CVs (95%+ pass rate)
    - 🌍 9 African Languages
    - 💬 AI-Powered Guidance
    - 📚 Culturally Grounded via RAG
    """)
    
    st.markdown("---")
    st.markdown("## 📚 Knowledge Base")
    
    if st.button("🔄 Initialize/Update Knowledge Base"):
        with st.spinner("Loading foundational knowledge..."):
            count = initial_populate_rag()
            if count > 0:
                st.success(f"✅ Loaded {count} foundational documents")
                st.balloons()
            else:
                st.info("ℹ️ Knowledge base already initialized")
    
    st.markdown("---")
    st.markdown("## 🤝 Partnership")
    st.markdown("**Pilot Partner:** Ahmadu Bello University Zaria, Nigeria")

# ----- ADMIN DASHBOARD -----
def admin_dashboard():
    st.title("🔐 Admin Dashboard")
    
    col1, col2 = st.columns([3, 1])
    with col2:
        if st.button("🚪 Logout"):
            st.session_state['admin_logged_in'] = False
            st.rerun()
    
    tabs = st.tabs(["📊 Analytics", "📚 Knowledge Base Management"])
    
    with tabs[0]:
        st.markdown("## 📊 Usage Analytics")
        
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
        st.markdown("## 📚 Knowledge Base Management")
        
        # Display current stats
        try:
            stats = index.describe_index_stats()
            total_vectors = stats.get('total_vector_count', 0)
            st.info(f"📊 Current Knowledge Base: {total_vectors} total chunks from all documents")
            
            # Estimate number of documents (assuming ~100-200 chunks per doc)
            est_docs = max(1, total_vectors // 150)
            st.info(f"💡 Estimated ~{est_docs} documents loaded")
        except:
            st.warning("Unable to fetch knowledge base statistics")
        
        st.markdown("---")
        st.markdown("### ➕ Upload New Document")
        st.info("ℹ️ Each upload ADDS to your existing knowledge base (does not replace)")
        
        uploaded_doc = st.file_uploader(
            "Upload Document to Knowledge Base",
            type=['pdf', 'txt', 'docx'],
            help="Limit 200MB per file • TXT, PDF, DOCX"
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
                            
                            st.success(f"✅ Document added! {len(chunks)} chunks indexed.")
                            st.success(f"📊 New total: {new_total} chunks in knowledge base (+{len(chunks)} from this upload)")
                            st.info(f"📝 Document ID: {doc_id}")
                            st.balloons()
                        else:
                            st.error("No valid content to index")
                    
                    except Exception as e:
                        st.error(f"Error: {str(e)}")
            else:
                st.warning("Please upload a document and specify source")

# ----- SECTION 1: RESUME ANALYSIS -----
def resume_analysis_section():
    log_analytics('section_accessed', 'Resume Analysis')
    
    st.markdown("## 📄 Professional Resume Analysis")
    st.markdown("Upload your resume to receive expert feedback grounded in African job market context.")
    
    uploaded_file = st.file_uploader(
        "📎 Upload Your Resume",
        type=['pdf', 'docx', 'txt'],
        help="Limit 200MB per file • PDF, DOCX, TXT"
    )
    
    city = st.text_input(
        "🏙️ Your City (Optional)",
        placeholder="e.g., Lagos, Nairobi, Accra",
        help="Helps provide location-specific advice"
    )
    
    additional_info = st.text_area(
        "📝 Additional Information",
        placeholder="Target industry, preferred roles...",
        help="Any additional context that helps us provide better feedback"
    )
    
    if st.button("🔍 Analyze Resume", key="analyze_btn"):
        if uploaded_file:
            log_analytics('resume_upload', f"File: {uploaded_file.name}")
            
            with st.spinner("🤔 Analyzing your resume..."):
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
{resume_text[:3000]}

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
                    
                    st.success("✅ Analysis Complete!")
                    st.markdown("### 📊 Your Resume Analysis")
                    st.markdown(feedback)
                    
                    log_analytics('analysis_completed')
                
                except Exception as e:
                    st.error(f"Error: {str(e)}")
        else:
            st.warning("⚠️ Please upload a resume first")
    
    # ===== POST-ANALYSIS: PREMIUM CV & COVER LETTER GENERATION =====
    if st.session_state.get('analysis_done', False):
        st.markdown("---")
        st.markdown("### ✨ Generate Premium Documents")
        st.markdown("Based on your resume analysis, generate a polished, ATS-optimized CV and/or a professional cover letter.")
        
        # Target job details for cover letter
        col_job1, col_job2 = st.columns(2)
        with col_job1:
            target_position = st.text_input(
                "🎯 Target Position (for Cover Letter)",
                placeholder="e.g., Data Analyst, Project Manager, Nurse",
                key="target_position_input"
            )
        with col_job2:
            target_company = st.text_input(
                "🏢 Target Company / Organization",
                placeholder="e.g., WHO, Dangote Group, Safaricom",
                key="target_company_input"
            )
        
        col_cv, col_cl = st.columns(2)
        
        # ===== GENERATE UPDATED CV =====
        with col_cv:
            if st.button("📄 Generate Updated CV", key="gen_cv_btn"):
                resume_text = st.session_state.get('resume_text', '').strip()
                feedback = st.session_state.get('resume_feedback', '').strip()
                
                if resume_text:
                    log_analytics('premium_cv_generation')
                    with st.spinner("📝 Generating your premium ATS-optimized CV..."):
                        try:
                            rag_context_cv = retrieve_career_guidance("professional CV resume best practices African job market ATS optimization")
                            
                            cv_gen_prompt = f"""You are a professional CV writer specializing in ATS-optimized resumes for the African job market.

ORIGINAL RESUME:
{resume_text[:4000]}

ANALYSIS FEEDBACK TO INCORPORATE:
{feedback[:2000]}

Based on the original resume details and the analysis feedback, create an improved, premium, ATS-optimized CV.

RESPOND ONLY WITH VALID JSON (no markdown, no code blocks, no preamble). Use this exact structure:
{{
  "full_name": "FULL NAME FROM RESUME",
  "credentials": "Degree abbreviations e.g. Ph.D., M.P.H., B.Sc.",
  "contact_line": "email | phone | city, country | LinkedIn (from resume)",
  "professional_summary": "3-4 sentence powerful summary with quantified achievements, tailored for African market. Incorporate feedback improvements.",
  "core_competencies": ["Competency 1", "Competency 2", "... 8-12 ATS keyword-rich competencies"],
  "work_experience": [
    {{
      "title": "Job Title",
      "company": "Company Name",
      "location": "City, Country",
      "dates": "Month Year - Month Year",
      "bullets": ["Achievement bullet with metrics (start with action verb)...", "...3-5 bullets per role"]
    }}
  ],
  "education": [
    {{
      "degree": "Degree Name, Field",
      "institution": "University Name",
      "dates": "Year"
    }}
  ],
  "publications": ["Citation 1 (if applicable)"],
  "projects": ["Project description (if applicable)"],
  "certifications": ["Certification 1 (if applicable)"],
  "technical_skills": "Comma-separated list of software, tools, platforms",
  "languages": ["English (Native)", "French (Professional)"]
}}

RULES:
- Keep ALL factual details from original resume — do NOT invent or hallucinate experience
- Improve bullet points with action verbs and quantified metrics where possible
- Ensure 95%+ ATS compatibility with clean formatting
- Target 2 pages of content
- Incorporate the feedback improvements (better keywords, stronger bullets, etc.)
- Return ONLY the JSON object, nothing else"""

                            cv_json = safe_llm_call(cv_gen_prompt, rag_context_cv, "English")

                            if cv_json.startswith("⚠️"):
                                st.warning(cv_json)
                            else:
                                cv_docx = generate_premium_cv_docx(cv_json)
                                st.session_state['cv_docx'] = cv_docx
                                st.session_state['cv_generated'] = True
                                st.success("✅ Premium CV Generated!")
                        
                        except Exception as e:
                            st.error(f"CV Generation Error: {str(e)}")
                            st.info("💡 Tip: Try again — the AI occasionally needs a second attempt for complex resumes.")
        
        # ===== GENERATE COVER LETTER =====
        with col_cl:
            if st.button("✉️ Generate Cover Letter", key="gen_cl_btn"):
                resume_text = st.session_state.get('resume_text', '')
                # Read directly from session state widget keys to avoid column scoping issues
                cl_position = st.session_state.get('target_position_input', '').strip()
                cl_company = st.session_state.get('target_company_input', '').strip()
                
                if resume_text and cl_position and cl_company:
                    log_analytics('premium_cover_letter_generation', f"{cl_position} at {cl_company}")
                    with st.spinner("✉️ Crafting your premium cover letter..."):
                        try:
                            rag_context_cl = retrieve_career_guidance(f"cover letter professional {cl_position} {cl_company} African job market")
                            
                            cl_gen_prompt = f"""You are a professional cover letter writer. Generate a premium, compelling cover letter.

CANDIDATE'S RESUME:
{resume_text[:4000]}

TARGET POSITION: {cl_position}
TARGET COMPANY: {cl_company}
LOCATION CONTEXT: {st.session_state.get('resume_city', 'Africa')}

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
- Use ONLY factual details from the resume — do NOT invent achievements or experience
- Each body paragraph should be 3-5 sentences, substantive and specific
- Use confident, professional tone
- Highlight quantified achievements where available in the resume
- Tailor language to the target company and African/international development context where applicable
- Return ONLY the JSON object, nothing else"""

                            cl_json = safe_llm_call(cl_gen_prompt, rag_context_cl, "English")

                            if cl_json.startswith("⚠️"):
                                st.warning(cl_json)
                            else:
                                cl_docx = generate_premium_cover_letter_docx(cl_json)
                                st.session_state['cl_docx'] = cl_docx
                                st.session_state['cl_generated'] = True
                                st.success("✅ Premium Cover Letter Generated!")
                        
                        except Exception as e:
                            st.error(f"Cover Letter Generation Error: {str(e)}")
                            st.info("💡 Tip: Try again — the AI occasionally needs a second attempt.")
                
                elif not cl_position or not cl_company:
                    st.warning("⚠️ Please enter the target position and company above to generate a cover letter.")
        
        # ===== DOWNLOAD BUTTONS =====
        st.markdown("---")
        dl_col1, dl_col2 = st.columns(2)
        
        with dl_col1:
            if st.session_state.get('cv_generated', False):
                cv_data = st.session_state['cv_docx']
                cv_data.seek(0)
                st.download_button(
                    label="⬇️ Download Updated CV (.docx)",
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
                    label="⬇️ Download Cover Letter (.docx)",
                    data=cl_data,
                    file_name=f"AfriCareer_CoverLetter_{datetime.now().strftime('%Y%m%d')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="dl_cl_btn"
                )

# ----- SECTION 2: CAREER GUIDANCE & CV BUILDER -----
def career_guidance_section():
    log_analytics('section_accessed', 'Career Guidance')
    
    st.markdown("## 🎓 Career Guidance & CV Builder")
    st.markdown("For youth and job seekers: Get personalized guidance and build your CV from scratch.")
    
    # ===== CRITICAL: RETRIEVE RAG CONTEXT AT START =====
    rag_context = retrieve_career_guidance("youth career guidance employability skills Africa")
    
    # SIMPLE, HUMANIZED QUESTIONS FOR AFRICAN YOUTH
    questions_text = f"""
📋 **Tell Us About Yourself**

Please answer these 5 simple questions (number your answers 1-5):

1. **What are you interested in?**
   (Example: I like computers, helping people, cooking, fixing things, etc.)

2. **What are you good at? What skills do you have?**
   (Example: I'm good at math, I can speak 3 languages, I know how to use Excel, etc.)

3. **What work or experience do you have?**
   (Include ANY experience: part-time jobs, helping family business, volunteer work, school projects, etc.)

4. **What is your education?**
   (Example: I finished secondary school in 2020, I'm studying at university, I completed a training course, etc.)

5. **What job do you want? What are your goals?**
   (Example: I want to work in a bank, I want to be a nurse, I want to start my own business, etc.)

✍️ **Write your answers below. Be honest - there are no wrong answers!**
"""
    
    st.markdown(questions_text)
    st.markdown("---")
    
    user_answers = st.text_area(
        "✍️ Your Answers (Number them 1-5)",
        height=250,
        placeholder="1. I'm interested in...\n2. My strengths...\n3. I have experience...\n4. My education...\n5. My goals...",
        key="answers_section2"
    )
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("🌟 Get Career Guidance", key="advice_btn"):
            if user_answers:
                log_analytics('query', f"Career guidance: {user_answers[:50]}...")
                
                with st.spinner("🤔 Preparing guidance..."):
                    # ===== CRITICAL: RETRIEVE SPECIFIC RAG CONTEXT =====
                    rag_context_advice = retrieve_career_guidance(f"career paths employment opportunities skills development Africa {user_answers[:100]}")
                    
                    # ===== USE SAFE LLM CALL WITH GUARDRAILS =====
                    advice_prompt = f"""Provide career guidance for African youth based on their profile.

Their Answers: {user_answers}

Provide:
1. Top 3 Career Paths in Africa (with specific reasons why they match)
2. Key Skills to Develop (7-10 skills with brief explanations)
3. Action Plan (5 concrete steps they can take now)

Ground advice in African job market realities and cite relevant frameworks when applicable."""

                    guidance = safe_llm_call(advice_prompt, rag_context_advice, target_lang)
                    
                    st.success("✅ Guidance Ready!")
                    st.markdown("### 🎯 Your Career Roadmap")
                    st.markdown(guidance)
    
    with col2:
        if st.button("📄 Generate CV", key="cv_btn"):
            if user_answers:
                log_analytics('cv_generation', f"User: {user_answers[:50]}...")
                
                with st.spinner("📝 Creating your CV..."):
                    try:
                        # ===== CRITICAL: RETRIEVE RAG CONTEXT FOR CV =====
                        rag_context_cv = retrieve_career_guidance("professional CV resume best practices African job market ATS")
                        
                        # ===== USE SAFE LLM CALL WITH GUARDRAILS =====
                        cv_prompt = f"""Create a professional, ATS-optimized CV based on this information:

{user_answers}

Generate:
- Professional Summary (3-4 sentences)
- Skills Section (8-10 relevant skills)
- Experience Section (format their experience professionally)
- Education Section
- Additional Sections if relevant

Follow African job market standards and ATS optimization best practices."""

                        cv_content = safe_llm_call(cv_prompt, rag_context_cv, target_lang)
                        
                        # Create DOCX
                        doc = Document()
                        
                        # Header
                        header = doc.add_heading(f'{APP_NAME}', 0)
                        header.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        
                        # Add CV content
                        for line in cv_content.split('\n'):
                            if line.strip():
                                if line.startswith('#'):
                                    doc.add_heading(line.replace('#', '').strip(), level=1)
                                else:
                                    doc.add_paragraph(line)
                        
                        # Footer
                        footer = doc.add_paragraph(f'\n\nGenerated by {APP_NAME} • {datetime.now().strftime("%B %d, %Y")}')
                        footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        
                        # Save to bytes
                        bio = io.BytesIO()
                        doc.save(bio)
                        bio.seek(0)
                        
                        st.success("✅ CV Generated!")
                        st.download_button(
                            label="⬇️ Download Your CV",
                            data=bio,
                            file_name=f"AfriCareer_CV_{datetime.now().strftime('%Y%m%d')}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                        
                    except Exception as e:
                        st.error(f"Error: {str(e)}")
            else:
                st.warning("⚠️ Please answer the questions first")

# ----- SECTION 3: LEARNING RESOURCES -----
def learning_resources_section():
    log_analytics('section_accessed', 'Learning Resources')
    
    st.markdown("## 📚 Learning Resources & Courses")
    
    st.info("💡 **Tip:** We provide course recommendations and platforms. Always verify links directly on the platform websites (Coursera, edX, Udemy, etc.) as course URLs change frequently.")
    st.warning("⚠️ **For Free Courses:** Check Class Central - it aggregates 50,900+ free courses from top universities!")
    
    st.markdown("---")
    
    st.markdown("### 🎯 What would you like to learn?")
    
    learning_interest = st.text_area(
        "📝 Describe what you want to learn",
        placeholder="e.g., data analysis, solar installation, digital marketing, tailoring & small business skills",
        help="The more specific, the better recommendations we can provide"
    )
    
    col1, col2 = st.columns(2)
    
    with col1:
        course_type = st.selectbox(
            "Course Type",
            ["Both free & Paid", "Free Only", "Paid Only"]
        )
    
    with col2:
        skill_level = st.selectbox(
            "Your Level",
            ["Beginner", "Intermediate", "Advanced"]
        )
    
    if st.button("🔍 Find Courses", key="course_search_btn"):
        if learning_interest:
            log_analytics('course_search', learning_interest)
            
            with st.spinner("🔎 Searching for courses..."):
                # ===== CRITICAL: RETRIEVE RAG CONTEXT =====
                rag_context_learning = retrieve_career_guidance(f"skills development training courses {learning_interest} African youth")
                
                # ===== USE SAFE LLM CALL WITH GUARDRAILS =====
                course_prompt = f"""Recommend learning resources for:

Topic: {learning_interest}
Level: {skill_level}
Preference: {course_type}

Provide:
1. Platform Coursera - 2-3 recommended courses (title only, no direct links)
2. Provider: deeplearning.ai - 1-2 specialized courses
3. Provider: Udacity - 1-2 relevant courses
4. How to Find: Specific search terms for Class Central
5. Duration: Approx. time commitment (e.g., 3 months)
6. Prerequisites: What they should know first

Note: Emphasize that users should search course titles on provider platforms directly. Mention Class Central for free options."""

                recommendations = safe_llm_call(course_prompt, rag_context_learning, target_lang)

                st.markdown("### 📚 Course Recommendations")
                st.markdown(recommendations)

                st.markdown("---")
                st.info(
                    "💡 **How to enroll:** Search these course titles directly on the provider's "
                    "website (Coursera, edX, Udacity, deeplearning.ai) — course URLs change often, "
                    "so titles are more reliable than links. For free options, "
                    "[Class Central](https://www.classcentral.com) aggregates 50,000+ free courses "
                    "from top universities; filter by topic, language, and cost."
                )
        else:
            st.warning("⚠️ Please describe what you want to learn")

# ----- SECTION 4: AI ASSISTANT -----
def ai_assistant_section():
    log_analytics('section_accessed', 'AI Assistant')
    
    st.markdown("## 💬 AI Career Assistant")
    st.markdown("Ask any career-related question and get personalized guidance.")
    
    # ===== CRITICAL: SHOW SAFETY NOTICE =====
    st.info("🛡️ **This assistant is designed specifically for career guidance.** It provides advice on careers, education, job searching, and professional development based on AfDB, UNICEF, and ILO frameworks for African youth.")
    
    question = st.text_area(
        "💭 Ask Your Question",
        placeholder="What skills should African youth prioritize according to international development frameworks for employability?",
        help="Ask about career paths, skills, education, job search, etc.",
        height=120
    )
    
    if st.button("🤖 Ask AI Assistant", key="ask_ai_btn"):
        if question:
            log_analytics('query', f"AI Assistant: {question}")
            
            with st.spinner("🤔 Thinking..."):
                # ===== CRITICAL: RETRIEVE RAG CONTEXT =====
                rag_context = retrieve_career_guidance(question)
                
                # ===== USE SAFE LLM CALL WITH GUARDRAILS =====
                response_content = safe_llm_call(question, rag_context, target_lang)
                
                st.markdown("### 🎯 Response")
                st.markdown(response_content)
        else:
            st.warning("⚠️ Please enter a question above before clicking 'Ask AI Assistant'")

# ----- ACCESS GATE -----
def require_access():
    """Gate the app behind an optional invite code (set APP_ACCESS_CODE to enable)."""
    if not APP_ACCESS_CODE:
        return True
    if st.session_state.get("access_granted"):
        return True

    st.title(f"🌍 {APP_NAME}")
    st.markdown(f"### *{TAGLINE}*")
    st.info("🔐 This is a private pilot. Please enter your access code to continue.")
    entered = st.text_input("Access code", type="password", key="access_code_input")
    if st.button("Enter"):
        if entered.strip() == APP_ACCESS_CODE:
            st.session_state["access_granted"] = True
            st.rerun()
        else:
            st.error("❌ Invalid access code. Please contact the AfriCareer AI team.")
    return False

# ----- MAIN APP -----
def main():
    global target_lang

    if st.session_state.get('admin_logged_in', False):
        admin_dashboard()
        return

    if not require_access():
        return

    if 'user_logged' not in st.session_state:
        log_analytics('user_visit')
        st.session_state['user_logged'] = True
    
    st.title(f"🌍 {APP_NAME}")
    st.markdown(f"### *{TAGLINE}*")
    st.markdown("**Partnership:** Ahmadu Bello University Zaria, Nigeria | **Developed by:** Quantium Insights LLC")

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
        "📄 Resume Analysis",
        "🎓 Career Guidance",
        "📚 Learning Resources",
        "💬 AI Assistant",
        "ℹ️ About"
    ])
    
    with tabs[0]:
        resume_analysis_section()
    
    with tabs[1]:
        career_guidance_section()
    
    with tabs[2]:
        learning_resources_section()
    
    with tabs[3]:
        ai_assistant_section()
    
    with tabs[4]:
        st.markdown("""
        ## About AfriCareer AI
        
        **AfriCareer AI** is an AI-powered career guidance platform for African youth.
        
        ### 🎯 Mission
        Empower 1 million+ urban African youth with professional career services.
        
        ### ✨ Key Features
        - **9 Languages:** English, French, Swahili, Arabic, Hausa, Pidgin, Portuguese, Spanish, Amharic
        - **ATS-Optimized CVs:** 95%+ pass rate
        - **Cultural Grounding:** RAG with UNICEF, ILO, AfDB frameworks
        - **Youth-Centric:** Designed for adolescents and young professionals
        - **Safety Guardrails:** Content filtering for appropriate career guidance
        
        ### 📚 Knowledge Base
        Our AI is grounded in authoritative frameworks:
        - **AfDB SEPA (2022-2025):** Skills for Employability and Productivity in Africa
        - **UNICEF Education Strategy (2019-2030):** Every Child Learns
        - **ILO Global Employment Trends for Youth (2022):** Investing in Transforming Futures
        
        ### 🤝 Partnership
        **Pilot Partner:** Ahmadu Bello University Zaria, Nigeria
        
        ### 👨‍💻 Developer
        **Dr. Amobi Andrew Onovo**
        - HIV Data Scientist, UNICEF
        - PhD Global Health, MPH, PGDip Data Science
        - Nigerian innovator
        
        ### 🔒 Safety & Ethics
        AfriCareer AI includes guardrails to ensure:
        - Mission-aligned responses only (career guidance)
        - No inappropriate or harmful content
        - Culturally appropriate advice for African context
        - Evidence-based recommendations from trusted sources
        """)

if __name__ == "__main__":
    main()
