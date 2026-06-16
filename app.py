"""
app.py – Giao diện Streamlit.
Chạy: streamlit run app.py
"""

import os
import io
import importlib
import tempfile
import zipfile

import streamlit as st
import pandas as pd

try:
    import pdfplumber, openpyxl, docx, docx2txt, pytesseract, cv2
    _MISSING = []
except ImportError as _e:
    _MISSING = [str(_e)]

from converter import (
    detect_and_convert,
    get_captured_logs,
    UnsupportedFormatError,
    ExtractionFailedError,
)

# ══════════════════════════════════════════════════════════════════════════════
# PAGE CONFIG
# ══════════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Table → Excel",
    page_icon="📊",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ══════════════════════════════════════════════════════════════════════════════
# CSS
# ══════════════════════════════════════════════════════════════════════════════
st.html("""
<link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;700&family=Inter:wght@300;400;500;600&display=swap" rel="stylesheet">
<style>
html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}
h1, h2, h3, .top-bar h1 {
    font-family: 'Outfit', sans-serif;
}

/* Set the overall background with a premium light blue gradient */
[data-testid="stAppViewContainer"] {
    background: linear-gradient(135deg, #F4F8FC 0%, #E5EFF9 100%) !important;
}

/* Moving gradient for the top bar header */
@keyframes gradientBG {
    0% { background-position: 0% 50%; }
    50% { background-position: 100% 50%; }
    100% { background-position: 0% 50%; }
}

.top-bar {
    background: linear-gradient(-45deg, #1F4E79, #2E75B6, #0F2B48, #2A5C8A) !important;
    background-size: 400% 400% !important;
    animation: gradientBG 12s ease infinite !important;
    border-radius: 16px; 
    padding: 35px 28px; 
    margin-bottom: 25px; 
    color: #ffffff;
    box-shadow: 0 10px 30px rgba(31, 78, 121, 0.15);
    border: 1px solid rgba(255, 255, 255, 0.12);
    text-align: center;
    transition: all 0.3s ease;
}
.top-bar h1 { margin: 0; font-size: 1.9rem; font-weight: 700; letter-spacing: -0.5px; }
.top-bar p  { margin: 10px 0 0; font-size: 0.9rem; opacity: 0.90; line-height: 1.5; }

.stat-row { display: flex; gap: 12px; flex-wrap: wrap; margin: 20px 0; }

/* Premium Glassmorphic Stat Cards */
.stat-card {
    background: rgba(255, 255, 255, 0.75) !important; 
    backdrop-filter: blur(10px);
    border: 1px solid rgba(31, 78, 121, 0.08) !important; 
    border-radius: 12px;
    padding: 16px 14px; 
    text-align: center; 
    flex: 1; 
    min-width: 100px;
    box-shadow: 0 6px 16px rgba(31, 78, 121, 0.03) !important;
    transition: all 0.3s cubic-bezier(0.25, 0.8, 0.25, 1);
}
.stat-card:hover {
    transform: translateY(-4px);
    box-shadow: 0 12px 24px rgba(31, 78, 121, 0.08) !important;
    border: 1px solid rgba(31, 78, 121, 0.2) !important;
    background: rgba(255, 255, 255, 0.9) !important; 
}
.stat-card .num { font-size: 1.7rem; font-weight: 700; color: #1F4E79; }
.stat-card .lbl { font-size: 0.72rem; color: #666666; text-transform: uppercase; letter-spacing: 0.8px; margin-top: 4px; font-weight: 500; }

/* Modern Primary Button */
.stButton>button {
    background: linear-gradient(135deg, #2E75B6 0%, #1F4E79 100%) !important;
    color: white !important;
    border-radius: 12px !important;
    border: none !important;
    font-weight: 600 !important;
    padding: 0.7rem 2.5rem !important;
    box-shadow: 0 4px 15px rgba(31, 78, 121, 0.2) !important;
    transition: all 0.3s cubic-bezier(0.25, 0.8, 0.25, 1) !important;
    font-family: 'Inter', sans-serif !important;
}
.stButton>button:hover {
    transform: translateY(-2px);
    box-shadow: 0 8px 22px rgba(31, 78, 121, 0.35) !important;
}
.stButton>button:active {
    transform: translateY(1px);
}

/* Modern Download Button */
.stDownloadButton>button {
    background: linear-gradient(135deg, #2E75B6 0%, #1F4E79 100%) !important;
    color: white !important;
    border-radius: 12px !important;
    border: none !important;
    font-weight: 600 !important;
    padding: 0.7rem 2.5rem !important;
    box-shadow: 0 4px 15px rgba(31, 78, 121, 0.2) !important;
    transition: all 0.3s cubic-bezier(0.25, 0.8, 0.25, 1) !important;
    font-family: 'Inter', sans-serif !important;
}
.stDownloadButton>button:hover {
    transform: translateY(-2px);
    box-shadow: 0 8px 25px rgba(31, 78, 121, 0.35) !important;
}
.stDownloadButton>button:active {
    transform: translateY(1px);
}

.log-box {
    background: #0d1117; color: #c9d1d9;
    font-family: 'Courier New', monospace; font-size: 0.78rem;
    padding: 12px 16px; border-radius: 10px;
    max-height: 180px; overflow-y: auto; line-height: 1.6;
    border: 1px solid #21262d; margin-top: 10px;
}
.log-box .ok   { color: #56d364; }
.log-box .warn { color: #e3b341; }
.log-box .err  { color: #f85149; }
.log-box .info { color: #58a6ff; }

.tab-meta { font-size: 0.82rem; color: #555555; margin: -4px 0 10px; font-weight: 500; }

/* Enhanced Glassmorphic Uploader */
[data-testid="stFileUploader"] {
    border: 2px dashed #2E75B6 !important; 
    border-radius: 16px !important; 
    padding: 24px !important;
    background-color: rgba(255, 255, 255, 0.6) !important;
    backdrop-filter: blur(8px);
    box-shadow: 0 4px 15px rgba(31, 78, 121, 0.02) !important;
    transition: all 0.3s ease !important;
}
[data-testid="stFileUploader"]:hover {
    background-color: rgba(255, 255, 255, 0.85) !important;
    border-color: #1F4E79 !important;
    box-shadow: 0 8px 25px rgba(31, 78, 121, 0.08) !important;
}

div.block-container {
    padding-top: 2.5rem !important;
    padding-bottom: 2.5rem !important;
    max-width: 800px !important;
}

/* Animations */
@keyframes fadeInUp {
    from {
        opacity: 0;
        transform: translateY(12px);
    }
    to {
        opacity: 1;
        transform: translateY(0);
    }
}
</style>
""")

if _MISSING:
    st.warning("⚠️ Hệ thống phát hiện thiếu một số thư viện sau: " + "; ".join(_MISSING))


# ══════════════════════════════════════════════════════════════════════════════
# HEADER
# ══════════════════════════════════════════════════════════════════════════════
st.html("""
<div class="top-bar">
  <h1>📊 Chuyển Đổi Bảng Dữ Liệu → Excel</h1>
  <p>PDF (text & scan) · Word (.doc / .docx) · Hình ảnh (.png / .jpg / .jpeg)
     · Tự động nhận diện cấu trúc bảng, không giới hạn số cột</p>
</div>
""")


# ══════════════════════════════════════════════════════════════════════════════
# UPLOAD
# ══════════════════════════════════════════════════════════════════════════════
# Initialize session state variables
if "conversion_results" not in st.session_state:
    st.session_state.conversion_results = {}
if "last_uploaded_names" not in st.session_state:
    st.session_state.last_uploaded_names = []

uploaded_files = st.file_uploader(
    "Kéo thả hoặc bấm để chọn một hoặc nhiều file",
    type=["pdf", "doc", "docx", "png", "jpg", "jpeg"],
    accept_multiple_files=True,
    help="Số cột nhận diện tự động theo từng file – không cố định.",
)

# Detect if the list of uploaded files changed
current_names = [f.name for f in uploaded_files] if uploaded_files else []
if current_names != st.session_state.last_uploaded_names:
    st.session_state.conversion_results = {}
    st.session_state.last_uploaded_names = current_names

if not uploaded_files:
    st.info("⬆️ Tải lên các file để bắt đầu.")
    st.stop()

# ── Nút + metric ────────────────────────────────────────────────────────────
total_size = sum(f.size for f in uploaded_files)
c1, c2 = st.columns([3, 1.2])
with c2:
    if total_size >= 1024 * 1024:
        size_str = f"{total_size / (1024 * 1024):.2f} MB"
    else:
        size_str = f"{total_size / 1024:.1f} KB"
    st.metric("Tổng kích thước", size_str)
with c1:
    run = st.button("▶ Bắt đầu chuyển đổi tất cả", type="primary", use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# PROCESSING
# ══════════════════════════════════════════════════════════════════════════════
if run:
    # Clear previous conversion results when starting a new run
    st.session_state.conversion_results = {}
    
    # Process each file sequentially
    for idx, uploaded in enumerate(uploaded_files):
        suffix = os.path.splitext(uploaded.name)[1].lower()
        excel_bytes = b""
        sheets_info = []
        convert_ok  = False
        err_msg     = ""
        logs        = []

        # Ghi file tạm (converter cần đường dẫn thực trên ổ đĩa)
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp_in:
            tmp_in.write(uploaded.getbuffer())
            tmp_in_path = tmp_in.name

        tmp_out_path = tmp_in_path + ".xlsx"

        try:
            status_placeholder = st.empty()
            with status_placeholder.container():
                with st.status(f"🔄 Đang xử lý ({idx+1}/{len(uploaded_files)}): {uploaded.name}...", expanded=True) as status:
                    # Chạy converter
                    try:
                        detect_and_convert(tmp_in_path, tmp_out_path)
                        convert_ok = True
                    except (UnsupportedFormatError, ExtractionFailedError,
                            FileNotFoundError) as exc:
                        err_msg = str(exc)
                    except Exception as exc:
                        err_msg = f"Lỗi không mong đợi: {exc}"

                    # Capture logs
                    logs = get_captured_logs()

                    if convert_ok:
                        # Đọc Excel vào RAM ngay (trước khi xóa file tạm)
                        with open(tmp_out_path, "rb") as f:
                            excel_bytes = f.read()

                        xl = pd.ExcelFile(io.BytesIO(excel_bytes))
                        for sheet in xl.sheet_names:
                            df = xl.parse(sheet)
                            sheets_info.append({
                                "name": sheet,
                                "rows": len(df),
                                "cols": len(df.columns),
                                "df":   df,
                            })
                        status.update(label=f"✅ Hoàn tất: {uploaded.name}", state="complete", expanded=False)
                    else:
                        status.update(label=f"❌ Lỗi xử lý: {uploaded.name}", state="error", expanded=True)
        finally:
            for p in (tmp_in_path, tmp_out_path):
                try:
                    if os.path.exists(p):
                        os.remove(p)
                except Exception:
                    pass

        # Save result for this specific file
        if convert_ok:
            st.session_state.conversion_results[uploaded.name] = {
                "ok": True,
                "excel_bytes": excel_bytes,
                "sheets_info": sheets_info,
                "uploaded_size": uploaded.size,
            }
        else:
            st.session_state.conversion_results[uploaded.name] = {
                "ok": False,
                "err_msg": err_msg,
                "logs": logs,
            }
        
        # Clear the status indicator widget to keep UI clean
        status_placeholder.empty()

# ══════════════════════════════════════════════════════════════════════════════
# RENDER RESULTS FROM SESSION STATE
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.conversion_results:
    results = st.session_state.conversion_results
    
    st.markdown("### 📊 Kết quả chuyển đổi")
    
    # Check if there are multiple successful conversions to show "Download All" ZIP button
    success_files = [f for f in results if results[f]["ok"]]
    if len(success_files) > 1:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for fname in success_files:
                out_name = os.path.splitext(fname)[0] + "_converted.xlsx"
                zip_file.writestr(out_name, results[fname]["excel_bytes"])
        
        st.download_button(
            label="🗂️ Tải xuống TẤT CẢ các file Excel (.zip)",
            data=zip_buffer.getvalue(),
            file_name="tat_ca_file_excel_chuyen_doi.zip",
            mime="application/zip",
            use_container_width=True,
            key="download_all_zip"
        )
        st.markdown("<div style='margin-top: -10px; margin-bottom: 25px;'></div>", unsafe_allow_html=True)
    
    # Create main tabs for each file
    file_tab_labels = []
    for fname in results:
        status_icon = "✅" if results[fname]["ok"] else "❌"
        file_tab_labels.append(f"{status_icon} {fname}")
        
    file_tabs = st.tabs(file_tab_labels)
    
    for tab, fname in zip(file_tabs, results):
        res = results[fname]
        with tab:
            if res["ok"]:
                # 1. Custom Blue success alert
                st.html(f"""
                <div style="background: rgba(235, 245, 255, 0.85); backdrop-filter: blur(10px); border: 1px solid #BEE3F8; color: #1E3A8A; padding: 18px; border-radius: 12px; font-weight: 600; display: flex; align-items: center; gap: 12px; margin-bottom: 20px; box-shadow: 0 6px 18px rgba(31, 78, 121, 0.05); animation: fadeInUp 0.6s ease;">
                  <span style="font-size: 1.3rem;">🎉</span> 
                  <span>Chuyển đổi thành công: <b>{fname}</b></span>
                </div>
                """)

                sheets_info = res["sheets_info"]
                excel_bytes = res["excel_bytes"]
                uploaded_size = res["uploaded_size"]

                # 2. STAT CARDS
                total_rows = sum(s["rows"] for s in sheets_info)
                max_cols   = max((s["cols"] for s in sheets_info), default=0)
                in_kb      = uploaded_size // 1024
                out_kb     = len(excel_bytes) // 1024

                st.html(f"""
                <div class="stat-row">
                  <div class="stat-card"><div class="num">{len(sheets_info)}</div><div class="lbl">Bảng (sheet)</div></div>
                  <div class="stat-card"><div class="num">{total_rows:,}</div><div class="lbl">Tổng số dòng</div></div>
                  <div class="stat-card"><div class="num">{max_cols}</div><div class="lbl">Số cột tối đa</div></div>
                  <div class="stat-card"><div class="num">{in_kb} KB</div><div class="lbl">File đầu vào</div></div>
                  <div class="stat-card"><div class="num">{out_kb} KB</div><div class="lbl">File Excel</div></div>
                </div>
                """)

                # 3. DOWNLOAD EXCEL
                out_name = os.path.splitext(fname)[0] + "_converted.xlsx"
                st.download_button(
                    label=f"📥 Tải xuống file Excel của {fname}",
                    data=excel_bytes,
                    file_name=out_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key=f"dl_{fname}"  # Unique key for Streamlit
                )

                st.markdown("---")

                # 4. PREVIEW
                st.subheader("👀 Xem trước dữ liệu")
                if not sheets_info:
                    st.warning("Không có sheet nào.")
                else:
                    tab_labels = [
                        f"📋 {s['name']}  ({s['rows']} dòng × {s['cols']} cột)"
                        for s in sheets_info
                    ]
                    tabs = st.tabs(tab_labels)

                    for sub_tab, info in zip(tabs, sheets_info):
                        with sub_tab:
                            df = info["df"]

                            st.html(
                                f'<div class="tab-meta">'
                                f'📌 <b>{info["rows"]}</b> dòng · <b>{info["cols"]}</b> cột'
                                f'</div>'
                            )

                            if df.empty:
                                st.info("Sheet này không có dữ liệu.")
                                continue

                            # Safe preview slider configurations
                            n_rows  = info["rows"]
                            s_max   = max(5, min(200, n_rows))
                            s_val   = min(50, n_rows)
                            s_min   = min(5, n_rows)
                            
                            if s_min < s_max:
                                preview_n = st.slider(
                                    "Số dòng xem trước",
                                    min_value=s_min,
                                    max_value=s_max,
                                    value=s_val,
                                    key=f"slider_{fname}_{info['name']}",
                                )
                            else:
                                preview_n = n_rows

                            st.dataframe(
                                df.head(preview_n),
                                use_container_width=True,
                                hide_index=True,
                            )
            else:
                # Show failure error box and captured logs for this file
                st.error(f"Lỗi khi xử lý file {fname}: {res['err_msg']}")
                if res["logs"]:
                    html_lines = []
                    for line in res["logs"]:
                        if   "[ERROR]"   in line: css = "err"
                        elif "[WARNING]" in line: css = "warn"
                        elif any(k in line.lower() for k in ("thành công", "lưu", "hoàn")):
                            css = "ok"
                        else:
                            css = "info"
                        html_lines.append(f'<div class="{css}">{line}</div>')
                    st.html(
                        f'<div class="log-box">{"".join(html_lines)}</div>'
                    )
else:
    # Prompt the user to run conversion if not already converted
    st.info("⬆️ Nhấn nút 'Bắt đầu chuyển đổi tất cả' ở trên để tiến hành xử lý.")