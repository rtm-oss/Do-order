import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import zipfile
import os
import shutil
from datetime import datetime
import subprocess
import platform

# --- 1. Page Configuration ---
st.set_page_config(
    page_title="Medical Auto-Docs Pro", 
    page_icon="🩺", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- 2. Advanced Custom CSS ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@300;400;600;700&family=Poppins:wght@300;400;600&display=swap');
    body { font-family: 'Poppins', 'Cairo', sans-serif; background-color: #f8fafc; }
    
    /* Animations */
    @keyframes pulse-blue {
        0% { transform: scale(1); box-shadow: 0 0 0 0 rgba(59, 130, 246, 0.7); }
        70% { transform: scale(1.05); box-shadow: 0 0 0 15px rgba(59, 130, 246, 0); }
        100% { transform: scale(1); box-shadow: 0 0 0 0 rgba(59, 130, 246, 0); }
    }
    @keyframes fadeInUp {
        from { opacity: 0; transform: translateY(20px); }
        to { opacity: 1; transform: translateY(0); }
    }

    /* Sidebar & Header */
    section[data-testid="stSidebar"] { background-color: #1e293b; color: white; }
    section[data-testid="stSidebar"] * { color: #e2e8f0 !important; }
    
    .header-container {
        background: linear-gradient(135deg, #0f172a 0%, #334155 100%);
        padding: 30px; border-radius: 16px; color: white;
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.1);
        display: flex; align-items: center; gap: 25px; margin-bottom: 40px;
        border: 1px solid rgba(255,255,255,0.1);
    }
    .header-icon {
        background: rgba(255,255,255,0.1); width: 70px; height: 70px;
        display: flex; align-items: center; justify-content: center;
        border-radius: 12px; font-size: 35px; backdrop-filter: blur(10px);
    }

    /* Stepper */
    .stepper-wrapper { display: flex; justify-content: space-between; margin-bottom: 40px; position: relative; padding: 0 40px; }
    .stepper-item::before, .stepper-item::after { position: absolute; content: ""; border-bottom: 4px solid #e2e8f0; width: 100%; top: 25px; z-index: 2; }
    .stepper-item::before { left: -50%; } .stepper-item::after { left: 50%; }
    .stepper-item { position: relative; display: flex; flex-direction: column; align-items: center; flex: 1; z-index: 5; }
    .stepper-item .step-counter {
        width: 54px; height: 54px; border-radius: 50%; background: #f1f5f9;
        border: 4px solid white; margin-bottom: 12px; font-weight: 800;
        color: #94a3b8; font-size: 18px; display: flex; justify-content: center; align-items: center;
        transition: all 0.4s;
    }
    .stepper-item.active .step-counter { background-color: #3b82f6; color: #fff; animation: pulse-blue 2s infinite; }
    .stepper-item.completed .step-counter { background-color: #10b981; color: #fff; }

    /* Cards */
    .file-card {
        background: white; border-radius: 16px; padding: 25px 20px;
        text-align: center; position: relative; overflow: hidden;
        transition: all 0.3s ease; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
        border: 1px solid #f1f5f9; height: 100%; animation: fadeInUp 0.6s ease-out;
    }
    .file-card::before { content: ""; position: absolute; top: 0; left: 0; right: 0; height: 6px; }
    .card-data::before { background: linear-gradient(90deg, #3b82f6, #0ea5e9); }
    .card-back::before { background: linear-gradient(90deg, #f59e0b, #d97706); }
    .card-knee::before { background: linear-gradient(90deg, #8b5cf6, #ec4899); }
    .file-card:hover { transform: translateY(-8px); }
    
    .icon-box {
        width: 70px; height: 70px; margin: 0 auto 15px; border-radius: 50%;
        display: flex; align-items: center; justify-content: center; font-size: 32px;
    }
    .card-data .icon-box { background: #eff6ff; color: #3b82f6; }
    .card-back .icon-box { background: #fffbeb; color: #f59e0b; }
    .card-knee .icon-box { background: #fbf8ff; color: #8b5cf6; }
    .card-title { font-weight: 700; font-size: 16px; color: #334155; }

    /* Buttons */
    .stButton > button { border-radius: 12px; height: 55px; font-weight: 600; border: none; box-shadow: 0 4px 6px rgba(0,0,0,0.05); font-size: 16px; }
    .btn-primary button { background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%); color: white; }
    .btn-secondary button { background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%); color: white; }
    .btn-success button { background: linear-gradient(135deg, #10b981 0%, #059669 100%); color: white; }

</style>
""", unsafe_allow_html=True)

# --- 3. Helper Functions ---
if 'step' not in st.session_state: st.session_state.step = 1
TEMP_FOLDER = "temp_gen_files"

def clean_number(value):
    val = str(value).strip()
    return val[:-2] if val.endswith('.0') else val

# --- دالة التحويل الذكية (Robust Conversion v2) ---
def convert_to_pdf_cross_platform(source_folder):
    abs_folder = os.path.abspath(source_folder)
    system_os = platform.system()
    
    if system_os == "Windows":
        try:
            from docx2pdf import convert
            import pythoncom
            pythoncom.CoInitialize()
            convert(abs_folder)
            return True, "Success"
        except Exception as e:
            return False, str(e)
            
    else: # Linux / Streamlit Cloud
        os.environ['HOME'] = '/tmp'
        try:
            check = subprocess.run(["which", "libreoffice"], capture_output=True, text=True)
            if check.returncode != 0:
                return False, "LibreOffice is MISSING. Ensure packages.txt exists."

            converted_count = 0
            errors = []
            files_to_convert = [f for f in os.listdir(abs_folder) if f.endswith(".docx")]
            
            if not files_to_convert: return False, "No DOCX files found."

            for filename in files_to_convert:
                input_path = os.path.join(abs_folder, filename)
                cmd = ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", abs_folder, input_path]
                result = subprocess.run(cmd, capture_output=True, text=True)
                
                if result.returncode == 0: converted_count += 1
                else: errors.append(f"{filename}: {result.stderr}")

            if converted_count > 0: return True, f"Converted {converted_count} files."
            else: return False, f"All failed. Errors: {errors[:2]}"
        except Exception as e:
            return False, f"System Error: {str(e)}"

# --- 4. Sidebar ---
with st.sidebar:
    if os.path.exists("sidebar_logo.png"): st.image("sidebar_logo.png", width=120)
    else: st.markdown("<div style='font-size: 50px; text-align:center'>🩺</div>", unsafe_allow_html=True)
    st.markdown("### ⚙️ Tools Menu")
    app_mode = st.radio("Choose Mode:", ["📝 Generator (Main)", "🔄 PDF Converter Tool"])
    # تم حذف سطر التنبيه حسب الطلب

# --- 5. Main Logic ---

# ==========================================
# MODE 1: GENERATOR (MAIN)
# ==========================================
if app_mode == "📝 Generator (Main)":
    st.markdown("""
    <div class="header-container">
        <div class="header-icon">🏥</div>
        <div class="header-text">
            <h1>Medical Docs Generator</h1>
            <p>Advanced Patient Document Processing System</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Stepper
    step1_class = "active" if st.session_state.step == 1 else "completed"
    step2_class = "active" if st.session_state.step == 2 else ("completed" if st.session_state.step > 2 else "")
    step3_class = "active" if st.session_state.step == 3 else ""
    
    st.markdown(f"""
    <div class="stepper-wrapper">
        <div class="stepper-item {step1_class}"><div class="step-counter">1</div><div class="step-name">Upload</div></div>
        <div class="stepper-item {step2_class}"><div class="step-counter">2</div><div class="step-name">Edit & Gen</div></div>
        <div class="stepper-item {step3_class}"><div class="step-counter">3</div><div class="step-name">Convert</div></div>
    </div>
    """, unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("<div class='file-card card-data'><div class='icon-box'>📊</div><div class='card-title'>Data Sheet</div></div>", unsafe_allow_html=True)
        uploaded_data = st.file_uploader(" ", type=['csv', 'xlsx'], key="u1", label_visibility="collapsed")
    with c2:
        st.markdown("<div class='file-card card-back'><div class='icon-box'>🦴</div><div class='card-title'>Back Template</div></div>", unsafe_allow_html=True)
        template_back = st.file_uploader(" ", type=['docx'], key="u2", label_visibility="collapsed")
    with c3:
        st.markdown("<div class='file-card card-knee'><div class='icon-box'>🦵</div><div class='card-title'>Knee Template</div></div>", unsafe_allow_html=True)
        template_knee = st.file_uploader(" ", type=['docx'], key="u3", label_visibility="collapsed")

    if uploaded_data and (template_back or template_knee):
        if uploaded_data.name.endswith('.csv'): df = pd.read_csv(uploaded_data, engine='python')
        else: df = pd.read_excel(uploaded_data)
        df.columns = df.columns.str.strip()
        
        # --- FIX: Ensure Height is treated as String/Object to allow free text editing ---
        if 'Height' in df.columns:
            # نحول العمود كله لنصوص عشان يقبل التعديل كنص
            df['Height'] = df['Height'].astype(str).replace('nan', '')
        # --------------------------------------------------------------------------

        df = df.fillna('')

        st.markdown("---")
        st.subheader("✏️ Step 1: Edit Data (Before Generation)")
        
        # استخدام st.column_config.TextColumn لقبول أي تنسيق
        edited_df = st.data_editor(
            df, 
            num_rows="dynamic", 
            use_container_width=True,
            column_config={
                "Height": st.column_config.TextColumn(
                    "Height",
                    help="Patient Height (Write exactly as you want, e.g. 5.8 or 5'8)"
                )
            }
        )

        st.markdown("---")
        col_left, col_right = st.columns(2)

        with col_left:
            st.markdown("### 🛠️ Step 2: Generate Drafts")
            st.markdown('<div class="btn-primary">', unsafe_allow_html=True)
            if st.button("🚀 Generate Word Files"):
                try:
                    if os.path.exists(TEMP_FOLDER): shutil.rmtree(TEMP_FOLDER)
                    os.makedirs(TEMP_FOLDER)

                    files_count = 0
                    bar = st.progress(0, "Processing...")

                    for i, row in edited_df.iterrows():
                        bar.progress((i + 1) / len(edited_df))
                        raw_phone = clean_number(row.get('Primary Phone', ''))
                        phone = raw_phone[1:] if raw_phone.startswith('0') else raw_phone
                        prod = str(row.get('Products', '')).upper().strip()
                        chk, emp = "☑", "☐"
                        mL = mR = emp
                        if 'LKB' in prod or 'LEFT' in prod: mL = chk
                        if 'RKB' in prod or 'RIGHT' in prod: mR = chk
                        if 'BKB' in prod: mL = mR = chk

                        ctx = {
                            'date': datetime.now().strftime("%d/%m/%Y"),
                            'first_name': str(row.get('Full Name', '')), 'last_name': str(row.get('Last Name', '')),
                            'dob': str(row.get('Date of Birth', '')), 'address': str(row.get('Address', '')),
                            'city': str(row.get('City', '')), 'state': str(row.get('State', '')),
                            'zip': clean_number(row.get('ZIP Code', '')), 'phone': phone,
                            'weight': clean_number(row.get('Weight', '')), 
                            
                            # --- هنا نستخدم clean_number العادية عشان تشيل .0 بس تسيب 5.8 زي ما هي ---
                            'height': clean_number(row.get('Height', '')),
                            
                            'insurance': str(row.get('Primary Insurance', '')), 'policy_num': clean_number(row.get('MCN', '')),
                            'dr_name': str(row.get('Dr Name', '')), 'dr_npi': clean_number(row.get('NPI', '')),
                            'dr_address': str(row.get('Dr Address', '')), 'dr_city': str(row.get('Dr City', '')),
                            'dr_state': str(row.get('Dr State', '')), 'dr_zip': clean_number(row.get('Dr ZIP Code', '')),
                            'dr_phone': clean_number(row.get('Dr Phone Number', '')), 'dr_fax': clean_number(row.get('Dr Fax', '')),
                            'L': mL, 'R': mR
                        }

                        tmpls = []
                        if template_back and ('BB' in prod or 'L0457' in prod): tmpls.append((template_back, "Back_Brace"))
                        if template_knee and ('KB' in prod or 'KNEE' in prod or 'L1833' in prod or 'BKB' in prod): tmpls.append((template_knee, "Knee_Brace"))

                        for t_file, suf in tmpls:
                            t_file.seek(0)
                            doc = DocxTemplate(t_file)
                            doc.render(ctx)
                            fn = f"{str(row.get('Full Name', 'Patient')).strip()}_{str(row.get('Last Name', '')).strip()}_{suf}.docx"
                            doc.save(os.path.join(TEMP_FOLDER, fn))
                            files_count += 1
                    
                    bar.empty()
                    if files_count > 0:
                        st.session_state.step = 2
                        st.success(f"Generated {files_count} drafts!")
                        
                        zip_buf = io.BytesIO()
                        with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED) as zf:
                            for fn in os.listdir(TEMP_FOLDER):
                                if fn.endswith(".docx"): zf.write(os.path.join(TEMP_FOLDER, fn), arcname=fn)
                        st.download_button("📥 Download Word Drafts", zip_buf.getvalue(), "Draft_Word_Docs.zip", "application/zip")
                    else:
                        st.warning("No files generated.")
                except Exception as e: st.error(f"Error: {e}")
            st.markdown('</div>', unsafe_allow_html=True)

        with col_right:
            st.markdown("### 👁️ Step 3: Convert to PDF")
            if st.session_state.step >= 2:
                st.markdown('<div class="btn-success">', unsafe_allow_html=True)
                if st.button("🔄 Convert All to PDF"):
                    if os.path.exists(TEMP_FOLDER):
                        with st.spinner("Converting on server (This might take a few seconds)..."):
                            
                            success, msg = convert_to_pdf_cross_platform(TEMP_FOLDER)

                            if success:
                                pdf_files = [f for f in os.listdir(TEMP_FOLDER) if f.endswith(".pdf")]
                                if len(pdf_files) > 0:
                                    zip_buf = io.BytesIO()
                                    with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED) as zf:
                                        for fn in pdf_files:
                                            zf.write(os.path.join(TEMP_FOLDER, fn), arcname=fn)
                                    
                                    st.session_state.step = 3
                                    st.success(f"✅ Converted {len(pdf_files)} files successfully!")
                                    st.download_button("📥 Download Final PDF ZIP", zip_buf.getvalue(), "Final_PDFs.zip", "application/zip")
                                else:
                                    st.error("⚠️ Conversion ran but produced 0 PDF files.")
                                    st.info("Debug Info: " + msg)
                            else:
                                st.error(f"❌ Conversion Failed: {msg}")
                    else:
                        st.warning("Generate Word files first.")
                st.markdown('</div>', unsafe_allow_html=True)
            else:
                 st.info("Complete Step 2 first.")

# ==========================================
# MODE 2: PDF CONVERTER TOOL
# ==========================================
elif app_mode == "🔄 PDF Converter Tool":
    st.markdown("""
    <div class="header-container" style="background: linear-gradient(135deg, #d35400 0%, #f39c12 100%);">
        <div style="font-size:40px;">🔄</div>
        <div>
            <h1 style="margin:0; font-size:24px;">Word to PDF Converter</h1>
            <p>Upload edited Word files -> Get PDFs</p>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    uploaded_docx = st.file_uploader("Upload Word Files (.docx)", type=['docx'], accept_multiple_files=True)
    
    if uploaded_docx:
        st.markdown('<div class="btn-success">', unsafe_allow_html=True)
        if st.button("Convert Uploaded Files"):
            conv_folder = "temp_convert_upload"
            if os.path.exists(conv_folder): shutil.rmtree(conv_folder)
            os.makedirs(conv_folder)
            
            for uf in uploaded_docx:
                with open(os.path.join(conv_folder, uf.name), "wb") as f:
                    f.write(uf.getbuffer())
            
            with st.spinner("Converting..."):
                success, msg = convert_to_pdf_cross_platform(conv_folder)
                if success:
                    pdf_files = [f for f in os.listdir(conv_folder) if f.endswith(".pdf")]
                    if len(pdf_files) > 0:
                        zip_buf = io.BytesIO()
                        with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED) as zf:
                            for fn in pdf_files:
                                zf.write(os.path.join(conv_folder, fn), arcname=fn)
                        st.success(f"Converted {len(pdf_files)} files!")
                        st.download_button("📥 Download PDFs", zip_buf.getvalue(), "Converted_PDFs.zip", "application/zip")
                    else:
                        st.error("No PDFs created.")
                        st.info("Debug: " + msg)
                else:
                    st.error(f"Conversion Failed: {msg}")
        st.markdown('</div>', unsafe_allow_html=True)

st.markdown("<div style='text-align: center; margin-top: 50px; color: #cbd5e1; font-size: 12px;'>Medical Docs Automation Tool © 2025</div>", unsafe_allow_html=True)
