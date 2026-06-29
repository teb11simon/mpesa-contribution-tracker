import os
import sys
import tempfile
from datetime import datetime
from pathlib import Path
import io
import json
import pandas as pd
import streamlit as st

# Ensure src directory is in path
src_path = Path(__file__).parent / "src"
if str(src_path) not in sys.path:
    sys.path.insert(0, str(src_path))

try:
    from mpesa_parser import MpesaParser, Transaction
    from excel_generator import ExcelGenerator
    from processor import ContributionProcessor
    from category_memory import CategoryMemory
    from member_registry import MemberRegistry
    from ocr_processor import OCRProcessor, ImagePreprocessor
    from settings_manager import SettingsManager
except ImportError as e:
    st.error(f"CRITICAL ERROR: A module is missing: {e}")

# ── Category Definitions ─────────────────────────────────────────────────
INCOME_CATEGORIES  = ["Contribution", "Benevolence", "Missions"]
EXPENSE_CATEGORIES = [
    "Worship & Admin & Welcome", "Facility Rental", "Transaction Charge",
    "Ministry/Assets", "Accommodation & Transport", "Mercy Day",
    "Lawyer/Regis", "Immigration & Flight",
    "Benevolence Expense", "Benevolence Expense Transaction Charge",
]
TRANSFER_CATEGORIES = ["Missions Transfer", "Benevolence Transfer", "Contribution Transfer"]

# ── Draft Save / Load Helpers ────────────────────────────────────────────
def _serialize_draft(session_state) -> bytes:
    """Serialize current session state to a downloadable JSON draft."""
    draft = {
        "version": 2,
        "step": session_state.get("step", 1),
        "report_date": session_state.get("report_date", datetime.now()).isoformat() if hasattr(session_state.get("report_date", datetime.now()), 'isoformat') else str(session_state.get("report_date", "")),
        "attendance": session_state.get("attendance", {}),
        "cash_breakdown": session_state.get("cash_breakdown", {}),
        "all_trans": [
            {
                "receipt_no": t.receipt_no,
                "date": t.date.isoformat() if hasattr(t.date, 'isoformat') else str(t.date),
                "details": t.details,
                "amount": float(t.amount),
                "transaction_type": t.transaction_type,
                "sender_name": t.sender_name,
                "sender_phone": getattr(t, 'sender_phone', None),
                "pre_category": getattr(t, 'pre_category', 'Contribution'),
            }
            for t in session_state.get("all_trans", [])
        ],
        "categorized_df": session_state.get("categorized_df").to_dict(orient="records") if session_state.get("categorized_df") is not None else None,
        "income_entries": _serialize_entries(session_state.get("income_entries", [])),
        "expense_entries": _serialize_entries(session_state.get("expense_entries", [])),
    }
    return json.dumps(draft, indent=2, default=str).encode("utf-8")

def _serialize_entries(entries):
    """Make entry dicts JSON-serializable by converting datetimes."""
    result = []
    for e in entries:
        d = dict(e)
        for k, v in d.items():
            if hasattr(v, 'isoformat'):
                d[k] = v.isoformat()
        result.append(d)
    return result

def _load_draft(draft_bytes, session_state):
    """Restore session state from a draft JSON file."""
    draft = json.loads(draft_bytes)
    
    # Restore step
    session_state.step = draft.get("step", 2)
    
    # Restore report date
    rd = draft.get("report_date", "")
    try:
        parsed_date = datetime.fromisoformat(rd)
        session_state.report_date = parsed_date.date() if hasattr(parsed_date, 'date') else parsed_date
    except Exception:
        session_state.report_date = datetime.now().date()
    
    session_state.attendance = draft.get("attendance", {})
    session_state.cash_breakdown = draft.get("cash_breakdown", {})
    
    # Restore transactions
    all_trans = []
    for td in draft.get("all_trans", []):
        try:
            t_date = datetime.fromisoformat(td["date"])
        except Exception:
            t_date = datetime.now()
        t = Transaction(
            receipt_no=td.get("receipt_no", ""),
            date=t_date,
            details=td.get("details", ""),
            amount=float(td.get("amount", 0)),
            transaction_type=td.get("transaction_type", "Paid In"),
            sender_name=td.get("sender_name"),
            sender_phone=td.get("sender_phone")
        )
        t.pre_category = td.get("pre_category", "Contribution")
        all_trans.append(t)
    session_state.all_trans = all_trans
    
    # Restore categorized dataframe
    cat_data = draft.get("categorized_df")
    if cat_data:
        session_state.categorized_df = pd.DataFrame(cat_data)
    
    # Restore income / expense entries
    def _parse_entries(entries_list):
        result = []
        for e in (entries_list or []):
            d = dict(e)
            if "date" in d and isinstance(d["date"], str):
                try:
                    d["date"] = datetime.fromisoformat(d["date"])
                except Exception:
                    pass
            result.append(d)
        return result
    
    session_state.income_entries = _parse_entries(draft.get("income_entries", []))
    session_state.expense_entries = _parse_entries(draft.get("expense_entries", []))

# ── Helper: Get church directory ────────────────────────────────────────
def get_church_dir(church_slug: str) -> Path:
    """Get the filesystem directory for a given church."""
    church_dir = Path("churches") / church_slug
    church_dir.mkdir(parents=True, exist_ok=True)
    return church_dir

def get_aliases_path(church_slug: str) -> Path:
    """Get the member_aliases.json path for a given church."""
    return get_church_dir(church_slug) / "member_aliases.json"

# ── Supabase Auth ────────────────────────────────────────────────────────
@st.cache_resource
def get_supabase():
    """Initialize and return a Supabase client using anon key."""
    import supabase
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return supabase.create_client(url, key)

@st.cache_resource
def get_supabase_admin():
    """Initialize and return a Supabase client using service_role key (admin only)."""
    import supabase
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_SERVICE_ROLE_KEY"]
    return supabase.create_client(url, key)

# ── Church helpers ──────────────────────────────────────────────────────
@st.cache_data(ttl=60)
def get_churches():
    """Fetch all churches from Supabase."""
    sup = get_supabase()
    resp = sup.table("churches").select("*").order("name").execute()
    return resp.data if resp.data else []

def get_church_by_id(church_id: str) -> dict | None:
    """Fetch a single church by ID."""
    sup = get_supabase()
    resp = sup.table("churches").select("*").eq("id", church_id).maybe_single().execute()
    return resp.data if resp.data else None

# ── Profile helpers ──────────────────────────────────────────────────────
def get_profile(user_id: str) -> dict | None:
    """Fetch the profile row for a given user id."""
    sup = get_supabase()
    resp = sup.table("profiles").select("*").eq("id", user_id).maybe_single().execute()
    if resp.data:
        return resp.data
    return None

# Streamlit Page Config
st.set_page_config(
    page_title="M-Pesa Contribution Tracker",
    layout="wide",
    page_icon="💰",
    initial_sidebar_state="expanded"
)

# Custom Sleek Styling
st.markdown("""
    <style>
    .main .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
        max-width: 1200px;
    }
    .stProgress > div > div > div > div {
        background-color: #27ae60;
    }
    div.stButton > button {
        border-radius: 6px;
    }
    </style>
""", unsafe_allow_html=True)

# ═════════════════════════════════════════════════════════════════════════
#  AUTHENTICATION CHECK
# ═════════════════════════════════════════════════════════════════════════

# Use session state for auth token management
if "supabase_session" not in st.session_state:
    st.session_state.supabase_session = None
if "user_profile" not in st.session_state:
    st.session_state.user_profile = None
if "church_info" not in st.session_state:
    st.session_state.church_info = None

# ── Render Auth Screen if not logged in ──────────────────────────────────
if st.session_state.supabase_session is None:
    st.markdown("<h1 style='text-align: center;'>💰 M-Pesa Tracker Pro</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: gray;'>Multi-Church Sunday Contribution Ledger</p>", unsafe_allow_html=True)

    tab1, tab2 = st.tabs(["🔐 Login", "📝 Sign Up"])

    with tab1:
        with st.form("login_form"):
            email = st.text_input("Email")
            password = st.text_input("Password", type="password")
            submitted = st.form_submit_button("Log In", use_container_width=True)
            if submitted:
                if not email or not password:
                    st.error("Please enter both email and password.")
                else:
                    try:
                        sup = get_supabase()
                        res = sup.auth.sign_in_with_password({"email": email, "password": password})
                        st.session_state.supabase_session = res.session
                        # Fetch profile
                        profile = get_profile(res.user.id)
                        st.session_state.user_profile = profile

                        # Fetch church info
                        if profile and profile.get("church_id"):
                            church = get_church_by_id(profile["church_id"])
                            st.session_state.church_info = church

                        st.rerun()
                    except Exception as e:
                        st.error(f"Login failed: {e}")

    with tab2:
        with st.form("signup_form"):
            new_email = st.text_input("Email")
            new_password = st.text_input("Password", type="password", help="At least 6 characters")
            
            # Church selection during signup
            churches = get_churches()
            church_options = {c["name"]: c["id"] for c in churches}
            church_names = list(church_options.keys())
            
            selected_church_name = st.selectbox(
                "Select Your Church",
                options=church_names,
                index=0,
                help="Choose the church you belong to. If your church isn't listed, contact the super admin."
            )
            
            submitted2 = st.form_submit_button("Create Account", use_container_width=True)
            if submitted2:
                if not new_email or not new_password:
                    st.error("Please fill in all fields.")
                elif len(new_password) < 6:
                    st.error("Password must be at least 6 characters.")
                else:
                    try:
                        sup = get_supabase()
                        # Sign up with metadata including church_id
                        res = sup.auth.sign_up({
                            "email": new_email,
                            "password": new_password,
                            "options": {
                                "data": {
                                    "church_id": church_options[selected_church_name]
                                }
                            }
                        })
                        
                        # If signup succeeded, update the profile's church_id
                        if res.user:
                            sup_admin = get_supabase_admin()
                            sup_admin.table("profiles").update({
                                "church_id": church_options[selected_church_name]
                            }).eq("id", res.user.id).execute()
                        
                        st.success(
                            "Account created! Please check your email for a confirmation link. "
                            "Once confirmed, ask your church admin to approve your account."
                        )
                    except Exception as e:
                        st.error(f"Sign up failed: {e}")

    st.stop()  # Don't render app content until authenticated

# ── Check profile status ─────────────────────────────────────────────────
profile = st.session_state.user_profile

if profile is None:
    # Try fetching again (e.g. after signup confirmation)
    sup = get_supabase()
    user = st.session_state.supabase_session.user
    profile = get_profile(user.id)
    st.session_state.user_profile = profile

if profile is None:
    st.warning("Your profile is not fully set up yet. Please contact the admin.")
    if st.button("Log Out"):
        st.session_state.supabase_session = None
        st.session_state.user_profile = None
        st.session_state.church_info = None
        st.rerun()
    st.stop()

# ── Ensure church_info is loaded ─────────────────────────────────────────
if st.session_state.church_info is None and profile.get("church_id"):
    church = get_church_by_id(profile["church_id"])
    st.session_state.church_info = church

church_info = st.session_state.church_info
church_name = church_info["name"] if church_info else "Your Church"
church_slug = church_info["slug"] if church_info else "nairobi-icc"
is_super_admin = (profile.get("role") == "admin")

# ── Show pending screen ──────────────────────────────────────────────────
if profile.get("status") == "pending":
    st.markdown("<h1 style='text-align: center;'>💰 M-Pesa Tracker Pro</h1>", unsafe_allow_html=True)
    st.info(f"⏳ Your account for **{church_name}** is pending admin approval. You'll be notified once approved.")
    if st.button("Log Out"):
        st.session_state.supabase_session = None
        st.session_state.user_profile = None
        st.session_state.church_info = None
        st.rerun()
    st.stop()

# ── Show rejected screen ─────────────────────────────────────────────────
if profile.get("status") == "rejected":
    st.markdown("<h1 style='text-align: center;'>💰 M-Pesa Tracker Pro</h1>", unsafe_allow_html=True)
    st.error("❌ Your account was rejected. Contact your church admin for more information.")
    if st.button("Log Out"):
        st.session_state.supabase_session = None
        st.session_state.user_profile = None
        st.session_state.church_info = None
        st.rerun()
    st.stop()

# ═════════════════════════════════════════════════════════════════════════
#  AUTHENTICATED – Render App
# ═════════════════════════════════════════════════════════════════════════

# Sidebar with user info
with st.sidebar:
    st.image("https://img.icons8.com/color/96/money-transfer.png", width=80)
    st.title("M-Pesa Tracker Pro")
    st.subheader(f"🏛️ {church_name}")
    st.markdown("---")
    st.markdown(f"**Logged in as:** {profile.get('email', '')}")
    st.markdown(f"**Role:** {'👑 Super Admin' if is_super_admin else '👤 User'}")
    st.markdown(f"**Church:** {church_name}")
    if st.button("🚪 Log Out"):
        st.session_state.supabase_session = None
        st.session_state.user_profile = None
        st.session_state.church_info = None
        st.rerun()
    st.markdown("---")

    # Member Management from Template
    st.markdown("### 👥 Manage Members")
    st.caption("Edit members directly from your uploaded template")

    if st.button("📋 Open Member Manager", use_container_width=True):
        st.session_state.show_member_manager = not st.session_state.get("show_member_manager", False)

    if st.session_state.get("show_member_manager", False):
        st.markdown("#### Current Members")
        # Load template if available
        template_path = st.session_state.get("temp_template_path", "")
        if template_path and Path(template_path).exists():
            try:
                processor.prepare_template(template_path)
                members = processor.generator._get_members_from_combined()
                if members:
                    member_data = []
                    for m in members:
                        member_data.append({
                            "Name": f"{m.get('first_name', '')} {m.get('last_name', '')}".strip(),
                            "Region": m.get('region', ''),
                            "Bible Talk": m.get('bible_talk', ''),
                            "Ministry": m.get('ministry', ''),
                            "Pledge": m.get('pledge', 0)
                        })
                    st.dataframe(pd.DataFrame(member_data), use_container_width=True, hide_index=True)
                else:
                    st.info("No members found in template.")
            except Exception as e:
                st.error(f"Error loading members: {e}")
        else:
            st.info("Upload a template in Step 1 to manage members.")

        st.markdown("#### Add New Member")
        with st.form("add_member_form"):
            col1, col2 = st.columns(2)
            with col1:
                new_first = st.text_input("First Name *")
                new_last = st.text_input("Last Name *")
                new_region = st.text_input("Region", value="")
            with col2:
                new_bible_talk = st.text_input("Bible Talk", value="")
                new_ministry = st.selectbox("Ministry", ["Campus", "Singles", "Marrieds", "Teens"])
                new_pledge = st.text_input("Pledge", value="0")

            if st.form_submit_button("➕ Add Member"):
                if new_first and new_last:
                    try:
                        processor.generator.load_template(template_path)
                        new_path = processor.generator.add_member(
                            new_region, new_bible_talk, new_ministry,
                            new_first, new_last, new_pledge
                        )
                        st.success(f"Added {new_first} {new_last}!")
                        # Update the template path in session state
                        if new_path:
                            st.session_state.temp_template_path = new_path
                        st.rerun()
                    except Exception as e:
                        st.error(f"Failed to add member: {e}")
                else:
                    st.warning("First Name and Last Name are required.")

        st.markdown("#### Remove Member")
        if members:
            member_names = [f"{m.get('first_name', '')} {m.get('last_name', '')}".strip() for m in members]
            member_to_remove = st.selectbox("Select member to remove", options=member_names)
            removal_type = st.radio("Removal Type", ["fallaway", "moveaway"], horizontal=True)
            if st.button("🗑️ Remove Selected Member"):
                if member_to_remove:
                    try:
                        processor.generator.load_template(template_path)
                        new_path = processor.generator.remove_member(member_to_remove.split()[0], removal_type)
                        st.success(f"Removed {member_to_remove} as {removal_type}")
                        if new_path:
                            st.session_state.temp_template_path = new_path
                        st.rerun()
                    except Exception as e:
                        st.error(f"Failed to remove member: {e}")

# ═════════════════════════════════════════════════════════════════════════
#  SUPER ADMIN PANEL (only for super admin users - role='admin')
# ═════════════════════════════════════════════════════════════════════════
if is_super_admin:
    with st.sidebar:
        st.markdown("---")
        st.markdown("### 🔧 Super Admin Panel")
        if st.button("👥 Admin: Manage Users"):
            st.session_state.show_admin_panel = not st.session_state.get("show_admin_panel", False)
        if st.button("🏛️ Admin: Manage Churches"):
            st.session_state.show_church_panel = not st.session_state.get("show_church_panel", False)

    # ── User Management Panel ────────────────────────────────────────────────
    if st.session_state.get("show_admin_panel", False):
        st.markdown("## 👥 Super Admin — User Management")
        st.markdown("Approve or reject pending user registrations across all churches.")

        try:
            sup_admin = get_supabase_admin()
            resp = sup_admin.table("profiles").select("*, churches(name)").order("created_at").execute()
            users = resp.data if resp.data else []

            if not users:
                st.info("No users registered yet.")
            else:
                data_rows = []
                for u in users:
                    church_name_row = u.get("churches", {}).get("name", "Unknown") if u.get("churches") else "Unknown"
                    data_rows.append({
                        "ID": u["id"][:8] + "...",
                        "Email": u["email"],
                        "Church": church_name_row,
                        "Role": u["role"],
                        "Status": u["status"],
                        "Created": u.get("created_at", "")[:10] if u.get("created_at") else ""
                    })
                df_users = pd.DataFrame(data_rows)
                st.dataframe(df_users, use_container_width=True, hide_index=True)

                st.markdown("#### Approve / Reject Pending Users")
                pending = [u for u in users if u["status"] == "pending"]
                if not pending:
                    st.success("No pending users!")
                else:
                    for u in pending:
                        church_name_row = u.get("churches", {}).get("name", "Unknown") if u.get("churches") else "Unknown"
                        col1, col2, col3, col4 = st.columns([2.5, 1.5, 1, 1])
                        with col1:
                            st.write(f"**{u['email']}**")
                        with col2:
                            st.write(f"Church: {church_name_row}")
                        with col3:
                            if st.button(f"✅ Approve", key=f"approve_{u['id']}"):
                                sup_admin.table("profiles").update({"status": "approved"}).eq("id", u["id"]).execute()
                                st.success(f"Approved {u['email']}")
                                st.rerun()
                        with col4:
                            if st.button(f"❌ Reject", key=f"reject_{u['id']}"):
                                sup_admin.table("profiles").update({"status": "rejected"}).eq("id", u["id"]).execute()
                                st.success(f"Rejected {u['email']}")
                                st.rerun()
        except Exception as e:
            st.error(f"Admin panel error: {e}")

        st.markdown("---")

    # ── Church Management Panel ────────────────────────────────────────────────
    if st.session_state.get("show_church_panel", False):
        st.markdown("## 🏛️ Super Admin — Church Management")
        st.markdown("Create new churches for the platform.")

        # List existing churches
        try:
            sup_admin = get_supabase_admin()
            churches_data = get_churches()
            
            if churches_data:
                st.markdown("### Existing Churches")
                church_rows = []
                for c in churches_data:
                    church_rows.append({
                        "Name": c["name"],
                        "Slug": c["slug"],
                        "Created": c.get("created_at", "")[:10] if c.get("created_at") else ""
                    })
                st.dataframe(pd.DataFrame(church_rows), use_container_width=True, hide_index=True)
            else:
                st.info("No churches configured yet.")
        except Exception as e:
            st.error(f"Error loading churches: {e}")

        st.markdown("### Create New Church")
        with st.form("create_church_form"):
            new_church_name = st.text_input("Church Name", help="e.g., 'Nairobi ICC', 'Mombasa ICC', etc.")
            new_church_slug = st.text_input(
                "Church Slug",
                help="URL-friendly identifier (e.g., 'nairobi-icc', 'mombasa-icc'). "
                     "Use lowercase letters, numbers, and hyphens only."
            )
            
            submitted_church = st.form_submit_button("Create Church", use_container_width=True)
            if submitted_church:
                if not new_church_name or not new_church_slug:
                    st.error("Please fill in both church name and slug.")
                else:
                    # Validate slug format
                    import re
                    if not re.match(r'^[a-z0-9-]+$', new_church_slug):
                        st.error("Slug must contain only lowercase letters, numbers, and hyphens.")
                    else:
                        try:
                            sup_admin.table("churches").insert({
                                "name": new_church_name,
                                "slug": new_church_slug
                            }).execute()
                            
                            # Create the church directory
                            get_church_dir(new_church_slug)
                            
                            # Clear cached churches
                            get_churches.clear()
                            
                            st.success(f"Church '{new_church_name}' created successfully!")
                            st.rerun()
                        except Exception as e:
                            error_msg = str(e)
                            if "duplicate key" in error_msg.lower():
                                st.error(f"A church with slug '{new_church_slug}' already exists.")
                            else:
                                st.error(f"Failed to create church: {e}")

        st.markdown("---")

# ═════════════════════════════════════════════════════════════════════════
#  MAIN APPLICATION (Approved Users Only)
# ═════════════════════════════════════════════════════════════════════════

# Initialize Session State
if "step" not in st.session_state:
    st.session_state.step = 1
if "all_trans" not in st.session_state:
    st.session_state.all_trans = []
if "cash_breakdown" not in st.session_state:
    st.session_state.cash_breakdown = {}
if "temp_template_path" not in st.session_state:
    st.session_state.temp_template_path = ""
if "report_date" not in st.session_state:
    st.session_state.report_date = datetime.now()
if "attendance" not in st.session_state:
    st.session_state.attendance = {"men": 0, "women": 0, "children": 0}
if "categorized_df" not in st.session_state:
    st.session_state.categorized_df = None

# Initialize Singletons (per-church)
@st.cache_resource
def get_processor(church_slug_param: str = "nairobi-icc"):
    return ContributionProcessor(church_slug=church_slug_param)

@st.cache_resource
def get_ocr_processor(ocr_backend: str = "tesseract"):
    return OCRProcessor(ocr_backend=ocr_backend)

@st.cache_resource
def get_mpesa_parser():
    return MpesaParser()

processor = get_processor(church_slug)
ocr_proc = get_ocr_processor(st.session_state.get("ocr_backend", "tesseract"))
parser = get_mpesa_parser()

# Steps Wizard Headers
cols = st.columns(3)
step_names = ["1. Upload Files", "2. Categorize Transactions", "3. Match Names & Download"]
for i, name in enumerate(step_names, 1):
    with cols[i-1]:
        if st.session_state.step == i:
            st.markdown(f"#### **🔵 {name}**")
        else:
            st.markdown(f"#### <span style='color:gray'>{name}</span>", unsafe_allow_html=True)
st.markdown("---")

# ═════════════════════════════════════════════════════════════════════════
#  STEP 1 — Upload Screen
# ═════════════════════════════════════════════════════════════════════════
if st.session_state.step == 1:
    st.markdown(f"### Upload M-Pesa PDF & Ledger Template — {church_name}")

    # Resume from Draft
    with st.expander("📤 Resume from a Saved Draft", expanded=False):
        draft_file = st.file_uploader("Upload a previously saved draft (.json)", type=["json"], key="draft_uploader")
        if draft_file is not None:
            try:
                _load_draft(draft_file.read(), st.session_state)
                st.success("✅ Draft loaded! Resuming your session...")
                st.rerun()
            except Exception as e:
                st.error(f"Failed to load draft: {e}")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("#### 1. Contribution Sources")
        uploaded_pdf = st.file_uploader("M-Pesa Statement (PDF or CSV)", type=["pdf", "csv"])
        pdf_password = st.text_input("M-Pesa PDF Password (if encrypted)", type="password", help="Usually your ID number")

        st.markdown("#### 2. Cash Record Note Images (Optional)")

        # OCR Engine selector
        ocr_choice = st.selectbox(
            "OCR Engine for Handwriting",
            options=["tesseract", "easyocr", "google_vision"],
            index=0,
            help="tesseract = good for printed text (free, local, works on Streamlit Cloud)\n"
                 "easyocr = best for handwriting (free, local, requires >1GB RAM)\n"
                 "google_vision = cloud-based (requires API key)"
        )
        if "ocr_backend" not in st.session_state or st.session_state.ocr_backend != ocr_choice:
            st.session_state.ocr_backend = ocr_choice
            get_ocr_processor.clear()
            st.rerun()

        img_contrib = st.file_uploader("Contribution Note Image", type=["png", "jpg", "jpeg", "bmp"])
        img_benev = st.file_uploader("Benevolence Note Image", type=["png", "jpg", "jpeg", "bmp"])
        img_missions = st.file_uploader("Missions Note Image", type=["png", "jpg", "jpeg", "bmp"])

    with col2:
        st.markdown("#### 3. Master Ledger Template (Required)")
        uploaded_template = st.file_uploader(f"Upload {church_name}'s Last Completed Report Template (.xlsx)", type="xlsx")

        st.markdown("#### 4. Settings & Metadata")
        report_date = st.date_input("Sunday Report Date", datetime.now())

        col_m, col_w, col_c = st.columns(3)
        with col_m:
            men_count = st.number_input("Men Attendance", min_value=0, value=0, step=1)
        with col_w:
            women_count = st.number_input("Women Attendance", min_value=0, value=0, step=1)
        with col_c:
            children_count = st.number_input("Children Attendance", min_value=0, value=0, step=1)

    if st.button("🚀 Process Uploads & Categorize", use_container_width=True):
        if not uploaded_template:
            st.error("Please upload last week's Excel template file.")
        elif not uploaded_pdf and not img_contrib and not img_benev and not img_missions:
            st.error("Please upload at least one M-Pesa statement or Cash note image.")
        else:
            with st.spinner("Processing statement and images..."):
                # Save ledger template locally temporarily
                temp_dir = tempfile.gettempdir()
                temp_template_path = os.path.join(temp_dir, f"{church_slug}_temp_template.xlsx")
                with open(temp_template_path, "wb") as f:
                    f.write(uploaded_template.getbuffer())

                st.session_state.temp_template_path = temp_template_path
                st.session_state.report_date = report_date
                st.session_state.attendance = {"men": men_count, "women": women_count, "children": children_count}

                all_trans = []
                combined_breakdown = {}

                # 1. Parse PDF / CSV
                if uploaded_pdf:
                    temp_pdf_path = os.path.join(temp_dir, f"{church_slug}_{uploaded_pdf.name}")
                    with open(temp_pdf_path, "wb") as f:
                        f.write(uploaded_pdf.getbuffer())

                    try:
                        parsed = parser.parse_pdf(temp_pdf_path, pdf_password)
                        for t in parsed:
                            d = (t.details or '').lower()
                            if t.transaction_type == 'Paid In':
                                if 'miss' in d:          t.pre_category = 'Missions'
                                elif 'benev' in d:       t.pre_category = 'Benevolence'
                                else:                    t.pre_category = 'Contribution'
                            else:
                                if 'charge' in d:        t.pre_category = 'Transaction Charge'
                                elif 'bundle' in d or 'data' in d: t.pre_category = 'Worship & Admin & Welcome'
                            all_trans.append(t)
                    except Exception as e:
                        st.error(f"Error parsing PDF: {e}")
                        st.stop()

                # 2. Process Cash Images
                img_paths = {
                    "Contribution": img_contrib,
                    "Benevolence": img_benev,
                    "Missions": img_missions
                }

                for cat, file_obj in img_paths.items():
                    if file_obj:
                        temp_img_path = os.path.join(temp_dir, f"{church_slug}_{file_obj.name}")
                        with open(temp_img_path, "wb") as f:
                            f.write(file_obj.getbuffer())

                        try:
                            entries, breakdown = ocr_proc.process_image(temp_img_path)
                            for e in entries:
                                t = Transaction("CASH", datetime.now(), f"Cash: {cat}", e.amount, "Paid In", sender_name=e.name)
                                t.pre_category = cat
                                all_trans.append(t)
                            for denom, total in breakdown.items():
                                combined_breakdown[denom] = combined_breakdown.get(denom, 0) + total
                        except Exception as e:
                            st.error(f"Error performing OCR on {cat} image: {e}")
                            st.stop()

                st.session_state.all_trans = all_trans
                st.session_state.cash_breakdown = combined_breakdown

                # Initialize DataFrame for Step 2
                df_data = []
                for idx, t in enumerate(all_trans):
                    sender = t.sender_name or ""
                    detail = t.details or ""
                    combined = f"{sender} ({detail})" if sender and detail else (sender or detail)

                    df_data.append({
                        "ID": idx + 1,
                        "Type": "IN" if t.transaction_type == "Paid In" else "OUT",
                        "Date": t.date.strftime("%Y-%m-%d %H:%M"),
                        "Amount": float(t.amount),
                        "Sender / Details": combined,
                        "Category": getattr(t, 'pre_category', 'Ignore' if t.transaction_type == "Paid Out" else 'Contribution'),
                        "Notes": ""
                    })

                st.session_state.categorized_df = pd.DataFrame(df_data)
                st.session_state.step = 2
                st.rerun()

# ═════════════════════════════════════════════════════════════════════════
#  STEP 2 — Categorize Transactions Screen
# ═════════════════════════════════════════════════════════════════════════
elif st.session_state.step == 2:
    st.markdown("### Step 2 of 3 — Categorize Transactions")
    st.write("Verify the system's categorization guesses. Mark items as 'Ignore' to exclude them from the ledger.")

    # Edit in table
    categories_list = INCOME_CATEGORIES + EXPENSE_CATEGORIES + TRANSFER_CATEGORIES + ["Ignore"]

    df = st.session_state.categorized_df

    with st.form("categorization_form"):
        edited_df = st.data_editor(
            df,
            column_config={
                "ID": st.column_config.NumberColumn("ID", disabled=True),
                "Type": st.column_config.TextColumn("Type", disabled=True),
                "Date": st.column_config.TextColumn("Date", disabled=True),
                "Amount": st.column_config.NumberColumn("Amount", format="Ksh %,.2f", disabled=True),
                "Sender / Details": st.column_config.TextColumn("Sender / Details", disabled=True),
                "Category": st.column_config.SelectboxColumn("Category", options=categories_list, required=True),
                "Notes": st.column_config.TextColumn("Notes (Expense Detail)", help="Notes for receipt details/expenses")
            },
            hide_index=True,
            use_container_width=True,
            num_rows="fixed"
        )

        col_back, col_save, col_next = st.columns([1, 1, 1])
        with col_back:
            back_btn = st.form_submit_button("← Back", use_container_width=True)
        with col_save:
            save_btn = st.form_submit_button("📥 Save Draft", use_container_width=True)
        with col_next:
            next_btn = st.form_submit_button("✓ Finish & Continue →", use_container_width=True)

    if back_btn:
        st.session_state.step = 1
        st.rerun()

    if save_btn:
        st.session_state.categorized_df = edited_df
        draft_data = _serialize_draft(st.session_state)
        report_dt = st.session_state.get("report_date", datetime.now())
        fname = f"draft_{church_slug}_{report_dt}.json" if not hasattr(report_dt, 'strftime') else f"draft_{church_slug}_{report_dt.strftime('%Y%m%d')}.json"
        st.download_button(
            label="⬇️ Download Draft File",
            data=draft_data,
            file_name=fname,
            mime="application/json",
            use_container_width=True
        )

    if next_btn:
        # Store changes
        st.session_state.categorized_df = edited_df

        # Parse edited categories
        income = []
        expense = []

        for idx, row in edited_df.iterrows():
            t = st.session_state.all_trans[idx]
            cat = row["Category"]
            note = str(row["Notes"]).strip()

            if cat == "Ignore":
                continue

            entry = {
                "receipt_no": t.receipt_no,
                "date": t.date,
                "details": note if note else (t.details or ""),
                "amount": t.amount,
                "type": t.transaction_type,
                "name": t.sender_name,
                "sender_name": t.sender_name,
                "sender_phone": t.sender_phone,
                "category": cat
            }

            if cat in INCOME_CATEGORIES:
                income.append(entry)
            else:
                expense.append(entry)

        st.session_state.income_entries = income
        st.session_state.expense_entries = expense
        st.session_state.step = 3
        st.rerun()

# ═════════════════════════════════════════════════════════════════════════
#  STEP 3 — Name Matching & Download
# ═════════════════════════════════════════════════════════════════════════
elif st.session_state.step == 3:
    st.markdown(f"### Step 3 of 3 — Member Matching & Final Generation — {church_name}")

    # Initialize/load members from template
    if "members" not in st.session_state or not st.session_state.members:
        with st.spinner("Extracting member list from Master Ledger template..."):
            try:
                members = processor.prepare_template(st.session_state.temp_template_path)
                st.session_state.members = members
            except Exception as e:
                st.error(f"Error loading members: {e}")
                st.stop()
    else:
        # Just ensure template is prepared for processor
        processor.prepare_template(st.session_state.temp_template_path)

    members = st.session_state.members
    income_entries = st.session_state.income_entries
    expense_entries = st.session_state.expense_entries

    # 1. Find matched and unmatched names
    matched_names = []
    unmatched_names = []
    seen_unmatched = set()

    for d in income_entries:
        sender_name = (d.get('sender_name') or d.get('name') or '').strip()
        if not sender_name:
            continue

        match, score = processor.matching_engine.find_match(sender_name, d.get('sender_phone'), d['amount'])
        if match:
            matched_names.append({
                "original_name": sender_name,
                "matched_to": f"{match.get('first_name', '')} {match.get('last_name', '')}".strip(),
                "amount": d['amount'],
                "confidence": score,
                "unmatch": False
            })
        else:
            if sender_name.lower() not in seen_unmatched:
                unmatched_names.append({"name": sender_name, "amount": d['amount']})
                seen_unmatched.add(sender_name.lower())

    st.markdown("#### 1. Review Member Matches")
    
    # Cache list for the callback
    st.session_state.matched_names_list = matched_names

    # 1a. View Matched Names
    with st.expander(f"✅ View Confirmed Matches ({len(matched_names)})", expanded=False):
        if not matched_names:
            st.info("No matches found yet.")
        else:
            def unmatch_callback():
                changes = st.session_state.get("matched_names_editor", {}).get("edited_rows", {})
                for row_idx, edits in changes.items():
                    if edits.get("unmatch") is True:
                        orig_name = st.session_state.matched_names_list[int(row_idx)]["original_name"]
                        processor.matching_engine.save_alias(orig_name, "--UNMATCHED--")
                        st.toast(f"Unmatched '{orig_name}'")

            match_df = pd.DataFrame(matched_names)
            st.data_editor(
                match_df,
                column_config={
                    "original_name": st.column_config.TextColumn("Receipt/Statement Name", disabled=True),
                    "matched_to": st.column_config.TextColumn("Matched Member", disabled=True),
                    "amount": st.column_config.NumberColumn("Amount", format="Ksh %,.2f", disabled=True),
                    "confidence": st.column_config.ProgressColumn("Confidence", min_value=0.0, max_value=1.0),
                    "unmatch": st.column_config.CheckboxColumn("Unmatch?", default=False)
                },
                hide_index=True,
                use_container_width=True,
                key="matched_names_editor",
                on_change=unmatch_callback
            )

    # 1b. Resolve Unmatched Names
    if not unmatched_names:
        st.success("🎉 All contribution names match members perfectly!")
    else:
        st.warning("The following names from M-Pesa or Cash could not be matched to an existing member. Map them to aliases below if they correspond to an existing member (they will be saved automatically).")

        # Build member list for dropdowns
        member_names = ["-- Select Member (Leave as Visitor) --"]
        for m in members:
            full_name = f"{m.get('first_name', '')} {m.get('last_name', '')}".strip()
            if full_name:
                member_names.append(full_name)

        def save_alias_callback(u_name, w_key):
            sel_member = st.session_state[w_key]
            if sel_member != "-- Select Member (Leave as Visitor) --":
                try:
                    processor.matching_engine.save_alias(u_name, sel_member)
                    st.toast(f"Saved mapping: {u_name} ➔ {sel_member}")
                except Exception as e:
                    st.error(f"Failed to save alias: {e}")

        # Render a form mapping names
        col_name, col_amt, col_member = st.columns([3, 1, 3])
        with col_name:
            st.markdown("**Unmatched Name**")
        with col_amt:
            st.markdown("**Amount**")
        with col_member:
            st.markdown("**Map to Existing Member**")

        for idx, item in enumerate(unmatched_names):
            c_name, c_amt, c_sel = st.columns([3, 1, 3])
            with c_name:
                st.write(item["name"])
            with c_amt:
                st.write(f"Ksh {item['amount']:,.2f}")
            with c_sel:
                # Use a unique key based on the name to prevent state leakage when items shift
                widget_key = f"alias_map_{idx}_{item['name']}"
                st.selectbox(
                    f"Map '{item['name']}'",
                    options=member_names,
                    key=widget_key,
                    label_visibility="collapsed",
                    on_change=save_alias_callback,
                    args=(item["name"], widget_key)
                )

        st.info("💡 Changes are saved automatically and the list will update immediately.")

    # 2. Excel Generation
    st.markdown("---")
    st.markdown("#### 2. Generate and Download Final Report")

    col_back, col_save3, col_gen = st.columns([1, 1, 1])
    with col_back:
        if st.button("← Back to Categorization", use_container_width=True):
            st.session_state.step = 2
            st.rerun()

    with col_save3:
        draft_data_3 = _serialize_draft(st.session_state)
        report_dt_3 = st.session_state.get("report_date", datetime.now())
        fname_3 = f"draft_{church_slug}_{report_dt_3}.json" if not hasattr(report_dt_3, 'strftime') else f"draft_{church_slug}_{report_dt_3.strftime('%Y%m%d')}.json"
        st.download_button(
            label="📥 Save Draft",
            data=draft_data_3,
            file_name=fname_3,
            mime="application/json",
            use_container_width=True
        )

    with col_gen:
        if st.button("✨ Compile Final Excel Report", use_container_width=True, type="primary"):
            with st.spinner("Generating Excel workbook..."):
                try:
                    # Save path locally temporarily
                    temp_dir = tempfile.gettempdir()
                    report_dt = st.session_state.report_date
                    report_dt_parsed = datetime.combine(report_dt, datetime.min.time())
                    
                    # Church-specific filename
                    temp_output_path = os.path.join(
                        temp_dir,
                        f"{church_slug}_Contribution_Report_{report_dt.strftime('%Y%m%d')}.xlsx"
                    )

                    # Set generator's cash breakdown
                    processor.generator.cash_breakdown = st.session_state.cash_breakdown

                    # Setup amount values for Bible Talk
                    for m in members:
                        m['amount'] = 0
                    for d in income_entries:
                        match, _ = processor.matching_engine.find_match(d['sender_name'], d.get('sender_phone'), d['amount'])
                        if match:
                            for m in members:
                                if m['row_index'] == match['row_index']:
                                    m['amount'] += d['amount']
                                    break

                    matched_mpesa, unmatched_mpesa = [], []
                    matched_cash, unmatched_cash = [], []

                    for d in income_entries:
                        entry = dict(d)
                        sender_name  = (d.get('sender_name') or d.get('name') or '').strip()
                        sender_phone = d.get('sender_phone')
                        match, _ = processor.matching_engine.find_match(sender_name, sender_phone, d['amount'])
                        if match:
                            entry['member_row'] = match['row_index']
                            first = str(match.get('first_name', '')).strip()
                            last  = str(match.get('last_name',  '')).strip()
                            entry['name'] = f"{first} {last}".title() if (first or last) else sender_name
                            if d.get('receipt_no') == "CASH":
                                matched_cash.append(entry)
                            else:
                                matched_mpesa.append(entry)
                        else:
                            if not entry.get('name'):
                                entry['name'] = sender_name
                            if d.get('receipt_no') == "CASH":
                                unmatched_cash.append(entry)
                            else:
                                unmatched_mpesa.append(entry)

                    # Run creator report with matches
                    processor.generator.create_report_with_matches(
                        report_dt_parsed, temp_output_path,
                        matched_mpesa, unmatched_mpesa, matched_cash, unmatched_cash, members
                    )

                    # Finalize names conversions for all other tabs
                    for entry in unmatched_mpesa + unmatched_cash + expense_entries:
                        if (entry.get('category') or '').strip().lower() in {'missions', 'benevolence', 'contribution', 'missions transfer', 'benevolence transfer', 'contribution transfer'}:
                            sender_name = (entry.get('sender_name') or entry.get('name') or '').strip()
                            match, _ = processor.matching_engine.find_match(sender_name, entry.get('sender_phone'), entry.get('amount'))
                            if match:
                                first = str(match.get('first_name', '')).strip()
                                last  = str(match.get('last_name',  '')).strip()
                                entry['name'] = f"{first} {last}".title() if (first or last) else sender_name

                    all_transactions_for_ledger = (
                        matched_mpesa + unmatched_mpesa +
                        matched_cash + unmatched_cash +
                        expense_entries
                    )

                    attn = st.session_state.attendance
                    processor.generator.finalize_report(
                        report_dt_parsed,
                        matched_mpesa + unmatched_mpesa,
                        matched_cash + unmatched_cash,
                        members,
                        temp_output_path,
                        attendance=attn,
                        all_transactions=all_transactions_for_ledger
                    )

                    # Read generated file bytes for download
                    with open(temp_output_path, "rb") as f:
                        report_bytes = f.read()

                    st.session_state.report_bytes = report_bytes
                    st.session_state.report_filename = f"{church_name}_Contribution_Report_{report_dt.strftime('%b_%d_%Y')}.xlsx"
                    st.success("🎉 Report compiled successfully!")

                except Exception as e:
                    st.error(f"Error compiling Excel report: {e}")
                    import traceback
                    st.text(traceback.format_exc())

    # Render download button if bytes are available
    if "report_bytes" in st.session_state:
        st.markdown("---")
        st.download_button(
            label="📥 Download Generated Excel Report",
            data=st.session_state.report_bytes,
            file_name=st.session_state.report_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

        if st.button("🔄 Start Over / Process New Report", use_container_width=True):
            st.session_state.clear()
            st.session_state.step = 1
            st.rerun()