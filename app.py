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

# --- 2. Advanced Custom CSS (The Pro Design) ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@300;400;600;700&family=Poppins:wght@300;400;600&display=swap');
    
    body {
        font-family: 'Poppins', 'Cairo', sans-serif;
        background-color: #f8fafc;
    }

    /* --- Animations --- */
    @keyframes pulse-blue {
        0% { transform: scale(1); box-shadow: 0 0 0 0 rgba(59, 130, 246, 0.7); }
        70% { transform: scale(1.05); box-shadow: 0 0 0 15px rgba(59, 130, 246, 0); }
        100% { transform: scale(1); box-shadow: 0 0 0 0 rgba(59, 130, 246, 0); }
    }
    @keyframes fadeInUp {
        from { opacity: 0; transform: translateY(20px); }
        to { opacity: 1; transform: translateY(0); }
    }

    /* --- Sidebar --- */
    section[data-testid="stSidebar"] {
        background-color: #1e293b;
        color: white;
    }
    section[data-testid="stSidebar"] * { color: #e2e8f0 !important; }
    
    /* --- Header --- */
    .header-container {
        background: linear-gradient(135deg, #0f172a 0%, #334155 100%);
        padding: 30px;
        border-radius: 16px;
        color: white;
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.1);
        display: flex;
        align-items: center;
        gap: 25px;
        margin-bottom: 40px;
        border: 1px solid rgba(255,255,255,0.1);
    }
    .header-icon {
        background: rgba(255,255,255,0.1);
        width: 70px;
        height: 70px;
        display: flex;
        align-items: center;
        justify-content: center;
        border-radius: 12px;
        font-size: 35px;
        backdrop-filter: blur(10px);
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .header-text h1 { margin: 0; font-size: 28px; font-weight: 700; letter-spacing: -0.5px; }

    /* --- Stepper --- */
    .stepper-wrapper {
        display: flex;
        justify-content: space-between;
        margin-bottom: 40px;
        position: relative;
        padding: 0 40px;
    }
    .stepper-item::before, .stepper-item::after {
        position: absolute;
        content: "";
        border-bottom: 4px solid #e2e8f0;
        width: 100%;
        top: 25px;
        z-index: 2;
    }
    .stepper-item::before { left: -50%; }
    .stepper-item::after { left: 50%; }
    
    .stepper-item {
        position: relative;
        display: flex;
        flex-direction: column;
        align-items: center;
        flex: 1;
        z-index: 5;
    }
    .stepper-item .step-counter {
        position: relative;
        z-index: 5;
        display: flex;
        justify-content: center;
        align-items: center;
        width: 54px;
        height: 54px;
        border-radius: 50%;
        background: #f1f5f9;
        border: 4px solid white;
        margin-bottom: 12px;
        font-weight: 800;
        color: #94a3b8;
        font-size: 18px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
        transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
    }
    .stepper-item.active .step-counter {
        background-color: #3b82f6;
        color: #fff;
        animation: pulse-blue 2s infinite;
        box-shadow: 0 0 0 4px rgba(59, 130, 246, 0.2);
    }
    .stepper-item.completed .step-counter {
        background-color: #10b981;
        color: #fff;
    }
    .stepper-item:first-child::before { content: none; }
    .stepper-item:last-child::after { content: none; }

    /* --- Cards --- */
    .file-card {
        background: white;
        border-radius: 16px;
        padding: 25px 20px;
        text-align: center;
        position: relative;
        overflow: hidden;
        transition: all 0.3s ease;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
        border: 1px solid #f1f5f9;
        height: 100%;
        animation: fadeInUp 0.6s ease-out;
    }
    .file-card::before {
        content: ""; position: absolute; top: 0; left: 0; right: 0; height: 6px;
    }
    .card-data::before { background: linear-gradient(90deg, #3b82f6, #0ea5e9); }
    .card-back::before { background: linear-gradient(90deg, #f59e0b, #d97706); }
    .card-knee::before { background: linear-gradient(90deg, #8b5cf6, #ec4899); }

    .file-card:hover { transform: translateY(-8px); box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.1); }

    .icon-box {
        width: 70px; height: 70px; margin: 0 auto 15px; border-radius: 50%;
        display: flex; align-items: center; justify-content: center; font-size: 32px;
        transition: all 0.3s ease;
    }
    .card-data .icon-box { background: #eff6ff; color: #3b82f6; }
    .card-back .icon-box { background: #fffbeb; color: #f59e0b; }
    .card-knee .icon-box { background: #fbf8ff; color: #8b5cf6; }

    .card-title { font-weight: 700; font-size: 16px; color: #334155; margin-bottom: 5px; }
    .card-subtitle { font-size: 12px; color: #94a3b8; font-weight: 400; }

    /* Buttons */
    .stButton > button {
        border-radius: 12px; height: 55px; font-weight: 600; border: none;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05); font-size: 16px; transition: all 0.2s;
    }
    .btn-primary button { background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%); color: white; }
    .btn-primary button:hover { box-shadow: 0 10px 15px -3px rgba(59, 130, 246, 0.4); transform: translateY(-2px); }
    
    .btn-success button { background: linear-gradient(135deg, #10b981 0%, #059669 100%); color: white; }
    .btn-success button:hover { box-shadow: 0 10px 15px -3px rgba(16, 185, 129, 0.4); transform: translateY(-2px); }

</style>
""", unsafe_allow_html=True)

# --- 3. Helper Functions & Server Logic ---
if 'step' not in st.session_state: st.session_state.step = 1
TEMP_FOLDER = "temp_gen_files"

def clean_number(value):
    val = str(value).strip()
    return val[:-2] if val.endswith('.0') else val

def format_as_float(value):
    try: return str(float(value)) if value and str(value).strip() else ""
    except: return str(value)

# دالة التحويل الذكية (تعمل على الويندوز واللينكس/سيرفر)
def convert_to_pdf_cross_platform(source_folder):
    system_os = platform.system()
    if system_os == "Windows":
        try:
            from docx2pdf import convert
            import pythoncom
            pythoncom.CoInitialize()
            convert(source_folder)
            return True
        except Exception as e:
            st.error(f"Windows Error: {e}")
            return False
    else: # Linux / Streamlit Cloud
        try:
            # استخدام LibreOffice الموجود على السيرفر
            cmd = f"libreoffice --headless --convert-to pdf --outdir \"{source_folder}\" \"{source_folder}/*.docx\""
            subprocess.run(cmd, shell=True, check=True)
            return True
        except Exception as e:
            st.error(f"Linux Conversion Error: {e}")
            return False

# --- 4. Sidebar ---
with st.sidebar:
    try: st.image("sidebar_logo.png", width=120)
    except: st.markdown("<div style='font-size: 50px; text-align:center'>🩺</div>", unsafe_allow_html=True)
    st.markdown("### ⚙️ Control Panel")
    st.info("System Ready for Deployment.")

# --- 5. Main Content ---

# Header
st.markdown("""
<div class="header-container">
    <div class="header-icon">🏥</div>
    <div class="header-text">
        <h1>Medical Docs Automation</h1>
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
    <div class="stepper-item {step1_class}">
        <div class="step-counter">1</div>
        <div class="step-name">Upload</div>
    </div>
    <div class="stepper-item {step2_class}">
        <div class="step-counter">2</div>
        <div class="step-name">Generate</div>
    </div>
    <div class="stepper-item {step3_class}">
        <div class="step-counter">3</div>
        <div class="step-name">Finish</div>
    </div>
</div>
""", unsafe_allow_html=True)

# --- Upload Section (The Nice Cards) ---
c1, c2, c3 = st.columns(3)

with c1:
    st.markdown("""
    <div class="file-card card-data">
        <div class="icon-box">📊</div>
        <div class="card-title">Patient Data</div>
        <div class="card-subtitle">CSV / Excel File</div>
    </div>
    """, unsafe_allow_html=True)
    uploaded_data = st.file_uploader(" ", type=['csv', 'xlsx'], key="u1", label_visibility="collapsed")

with c2:
    st.markdown("""
    <div class="file-card card-back">
        <div class="icon-box">🦴</div>
        <div class="card-title">Back Template</div>
        <div class="card-subtitle">Word Document (.docx)</div>
    </div>
    """, unsafe_allow_html=True)
    template_back = st.file_uploader(" ", type=['docx'], key="u2", label_visibility="collapsed")

with c3:
    st.markdown("""
    <div class="file-card card-knee">
        <div class="icon-box">🦵</div>
        <div class="card-title">Knee Template</div>
        <div class="card-subtitle">Word Document (.docx)</div>
    </div>
    """, unsafe_allow_html=True)
    template_knee = st.file_uploader(" ", type=['docx'], key="u3", label_visibility="collapsed")

# --- Logic ---
if uploaded_data and (template_back or template_knee):
    if uploaded_data.name.endswith('.csv'):
        df = pd.read_csv(uploaded_data, engine='python')
    else:
        df = pd.read_excel(uploaded_data)
    df.columns = df.columns.str.strip()
    df = df.fillna('')
    
    st.markdown("<br>", unsafe_allow_html=True)

    col_left, col_right = st.columns([1, 1], gap="large")

    with col_left:
        st.markdown("### 🛠️ Process Data")
        
        st.markdown('<div class="btn-primary">', unsafe_allow_html=True)
        if st.button("🚀 Start Generation"):
            try:
                if os.path.exists(TEMP_FOLDER): shutil.rmtree(TEMP_FOLDER)
                os.makedirs(TEMP_FOLDER)

                files_count = 0
                my_bar = st.progress(0, text="Processing...")

                for index, row in df.iterrows():
                    my_bar.progress((index + 1) / len(df), text=f"Processing: {row.get('Full Name', 'Patient')}")
                    
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
                        'first_name': str(row.get('Full Name', '')),
                        'last_name': str(row.get('Last Name', '')),
                        'dob': str(row.get('Date of Birth', '')),
                        'address': str(row.get('Address', '')),
                        'city': str(row.get('City', '')),
                        'state': str(row.get('State', '')),
                        'zip': clean_number(row.get('ZIP Code', '')),
                        'phone': phone,
                        'weight': clean_number(row.get('Weight', '')),
                        'height': format_as_float(row.get('Height', '')),
                        'insurance': str(row.get('Primary Insurance', '')),
                        'policy_num': clean_number(row.get('MCN', '')),
                        'dr_name': str(row.get('Dr Name', '')),
                        'dr_npi': clean_number(row.get('NPI', '')),
                        'dr_address': str(row.get('Dr Address', '')),
                        'dr_city': str(row.get('Dr City', '')),
                        'dr_state': str(row.get('Dr State', '')),
                        'dr_zip': clean_number(row.get('Dr ZIP Code', '')),
                        'dr_phone': clean_number(row.get('Dr Phone Number', '')),
                        'dr_fax': clean_number(row.get('Dr Fax', '')),
                        'L': mL, 'R': mR
                    }

                    tmpls = []
                    if template_back and ('BB' in prod or 'L0457' in prod):
                        tmpls.append((template_back, "Back_Brace"))
                    if template_knee and ('KB' in prod or 'KNEE' in prod or 'L1833' in prod or 'BKB' in prod):
                        tmpls.append((template_knee, "Knee_Brace"))
                    
                    for t_file, suf in tmpls:
                        t_file.seek(0)
                        doc = DocxTemplate(t_file)
                        doc.render(ctx)
                        fn = f"{str(row.get('Full Name', 'Patient')).strip()}_{str(row.get('Last Name', '')).strip()}_{suf}.docx"
                        doc.save(os.path.join(TEMP_FOLDER, fn))
                        files_count += 1

                my_bar.empty()
                if files_count > 0:
                    st.session_state.step = 2
                    st.success(f"Generated {files_count} files successfully!")
                else:
                    st.warning("No matching products found.")

            except Exception as e:
                st.error(f"Error: {e}")
        st.markdown('</div>', unsafe_allow_html=True)

    with col_right:
        st.markdown("### 👁️ Review & Finalize")
        
        if st.session_state.step >= 2:
            st.info("Files ready for conversion.")
            
            # زر فتح الفولدر يظهر فقط لو شغالين لوكال (Windows)
            if platform.system() == "Windows":
                st.markdown('<div class="btn-secondary">', unsafe_allow_html=True)
                def open_folder_local():
                    try: os.startfile(TEMP_FOLDER)
                    except: pass
                if st.button("📂 Open Folder (Local Only)"):
                    open_folder_local()
                st.markdown('</div>', unsafe_allow_html=True)
            
            st.markdown("<hr style='margin: 15px 0; opacity: 0.3;'>", unsafe_allow_html=True)
            
            st.markdown('<div class="btn-success">', unsafe_allow_html=True)
            if st.button("📥 Convert to PDF"):
                try:
                    with st.spinner("Converting documents to PDF (Server-side)..."):
                        # استدعاء دالة التحويل الذكية
                        success = convert_to_pdf_cross_platform(TEMP_FOLDER)

                        if success:
                            zip_buffer = io.BytesIO()
                            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zf:
                                for filename in os.listdir(TEMP_FOLDER):
                                    if filename.endswith(".pdf"):
                                        zf.write(os.path.join(TEMP_FOLDER, filename), arcname=filename)
                            
                            st.session_state.step = 3
                            st.success("Conversion Complete!")
                            st.download_button(
                                label="Download ZIP",
                                data=zip_buffer.getvalue(),
                                file_name="Completed_Docs.zip",
                                mime="application/zip"
                            )
                        else:
                            st.error("Conversion failed. Check server logs.")
                except Exception as e:
                    st.error(f"Error: {e}")
            st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.markdown("""
            <div style="background:#f8fafc; padding:20px; border-radius:10px; text-align:center; color:#94a3b8; border:2px dashed #e2e8f0;">
                Waiting for generation step...
            </div>
            """, unsafe_allow_html=True)

st.markdown("""
<div style='text-align: center; margin-top: 50px; color: #cbd5e1; font-size: 12px;'>
    Medical Docs Automation Tool © 2025
</div>
""", unsafe_allow_html=True)