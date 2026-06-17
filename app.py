import os
import sys
import tempfile
from datetime import datetime
from pathlib import Path
import io
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
    from ocr_processor import OCRProcessor
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

# Initialize Singletons
@st.cache_resource
def get_processor():
    return ContributionProcessor()

@st.cache_resource
def get_ocr_processor():
    return OCRProcessor()

@st.cache_resource
def get_mpesa_parser():
    return MpesaParser()

processor = get_processor()
ocr_proc = get_ocr_processor()
parser = get_mpesa_parser()

# Sidebar info
with st.sidebar:
    st.image("https://img.icons8.com/color/96/money-transfer.png", width=80)
    st.title("M-Pesa Tracker Pro")
    st.subheader("Nairobi ICC Sunday Ledger")
    st.write("A secure web dashboard to process M-Pesa PDFs and note images, matching contributors with church members.")
    st.markdown("---")
    
    # Help upload current member mappings
    st.markdown("### Aliases Configuration")
    aliases_file = Path("member_aliases.json")
    if aliases_file.exists():
        with open(aliases_file, "r") as f:
            st.download_button(
                label="📥 Download Current member_aliases.json",
                data=f.read(),
                file_name="member_aliases.json",
                mime="application/json"
            )
    
    uploaded_aliases = st.file_uploader("Upload existing member_aliases.json", type="json")
    if uploaded_aliases is not None:
        with open(aliases_file, "wb") as f:
            f.write(uploaded_aliases.getvalue())
        st.success("Successfully uploaded & loaded member_aliases.json!")
        # Force re-init of matching engine
        if hasattr(processor, "matching_engine") and processor.matching_engine:
            processor.matching_engine.load_aliases()

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
    st.markdown("### Upload M-Pesa PDF & Ledger Template")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("#### 1. Contribution Sources")
        uploaded_pdf = st.file_uploader("M-Pesa Statement (PDF or CSV)", type=["pdf", "csv"])
        pdf_password = st.text_input("M-Pesa PDF Password (if encrypted)", type="password", help="Usually your ID number")
        
        st.markdown("#### 2. Cash Record Note Images (Optional)")
        img_contrib = st.file_uploader("Contribution Note Image", type=["png", "jpg", "jpeg", "bmp"])
        img_benev = st.file_uploader("Benevolence Note Image", type=["png", "jpg", "jpeg", "bmp"])
        img_missions = st.file_uploader("Missions Note Image", type=["png", "jpg", "jpeg", "bmp"])
        
    with col2:
        st.markdown("#### 3. Master Ledger Template (Required)")
        uploaded_template = st.file_uploader("Upload Last Week's Completed Report Template (.xlsx)", type="xlsx")
        
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
                temp_template_path = os.path.join(temp_dir, "temp_template.xlsx")
                with open(temp_template_path, "wb") as f:
                    f.write(uploaded_template.getbuffer())
                
                st.session_state.temp_template_path = temp_template_path
                st.session_state.report_date = report_date
                st.session_state.attendance = {"men": men_count, "women": women_count, "children": children_count}
                
                all_trans = []
                combined_breakdown = {}
                
                # 1. Parse PDF / CSV
                if uploaded_pdf:
                    temp_pdf_path = os.path.join(temp_dir, uploaded_pdf.name)
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
                        temp_img_path = os.path.join(temp_dir, file_obj.name)
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
    
    # Store changes
    st.session_state.categorized_df = edited_df
    
    col_back, col_next = st.columns([1, 1])
    with col_back:
        if st.button("← Back", use_container_width=True):
            st.session_state.step = 1
            st.rerun()
            
    with col_next:
        if st.button("✓ Finish & Continue →", use_container_width=True):
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
    st.markdown("### Step 3 of 3 — Member Matching & Final Generation")
    
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
    
    # 1. Find unmatched names
    unmatched_names = []
    seen_unmatched = set()
    
    for d in income_entries:
        sender_name = (d.get('sender_name') or d.get('name') or '').strip()
        if not sender_name:
            continue
        
        match, _ = processor.matching_engine.find_match(sender_name, d.get('sender_phone'), d['amount'])
        if not match:
            if sender_name.lower() not in seen_unmatched:
                unmatched_names.append({"name": sender_name, "amount": d['amount']})
                seen_unmatched.add(sender_name.lower())
                
    st.markdown("#### 1. Review Member Matches")
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
                selected_member = st.selectbox(
                    f"Map '{item['name']}'",
                    options=member_names,
                    key=f"alias_map_{idx}",
                    label_visibility="collapsed"
                )
                if selected_member != "-- Select Member (Leave as Visitor) --":
                    try:
                        processor.matching_engine.save_alias(item["name"], selected_member)
                        st.toast(f"Saved mapping: {item['name']} ➔ {selected_member}")
                    except Exception as e:
                        st.error(f"Failed to save alias: {e}")
                        
        st.info("💡 Clicking 'Refresh Page' after selecting mappings will update the list.")
        if st.button("🔄 Refresh Matches"):
            st.rerun()

    # 2. Excel Generation
    st.markdown("---")
    st.markdown("#### 2. Generate and Download Final Report")
    
    col_back, col_gen = st.columns([1, 1])
    with col_back:
        if st.button("← Back to Categorization", use_container_width=True):
            st.session_state.step = 2
            st.rerun()
            
    with col_gen:
        if st.button("✨ Compile Final Excel Report", use_container_width=True, type="primary"):
            with st.spinner("Generating Excel workbook..."):
                try:
                    # Save path locally temporarily
                    temp_dir = tempfile.gettempdir()
                    report_dt = st.session_state.report_date
                    report_dt_parsed = datetime.combine(report_dt, datetime.min.time())
                    temp_output_path = os.path.join(temp_dir, f"Nairobi_Contribution_Report_{report_dt.strftime('%Y%m%d')}.xlsx")
                    
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
                    st.session_state.report_filename = f"Nairobi_Contribution_Report_{report_dt.strftime('%b_%d_%Y')}.xlsx"
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
