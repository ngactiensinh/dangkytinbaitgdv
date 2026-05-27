import streamlit as st
import pandas as pd
from supabase import create_client
from datetime import datetime
import io
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import plotly.express as px
import unicodedata
import re

# ─────────────────────────────────────────────
# 1. CẤU HÌNH TRANG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Hệ thống Quản lý Tin bài – Ban Tuyên giáo",
    page_icon="📰",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# 2. KẾT NỐI SUPABASE
# ─────────────────────────────────────────────
try:
    URL = st.secrets["SUPABASE_URL"]
    KEY = st.secrets["SUPABASE_KEY"]
    ADMIN_PASS = st.secrets.get("ADMIN_PASS", "141983")
    supabase = create_client(URL, KEY)
except Exception as e:
    st.error("⚠️ Lỗi kết nối cơ sở dữ liệu. Vui lòng kiểm tra cấu hình Secrets!")
    st.stop()

# ─────────────────────────────────────────────
# 3. CSS HIỆN ĐẠI - PHONG CÁCH HÀNH CHÍNH
# ─────────────────────────────────────────────
CUSTOM_CSS = """
<style>
    /* Reset & Font */
    html, body, [data-testid="stAppViewContainer"] {
        background-color: #F5F7FA;
        font-family: 'Inter', 'Segoe UI', 'Roboto', sans-serif;
    }

    /* Header Banner */
    .header-banner {
        background: linear-gradient(135deg, #1B3A5C 0%, #2C5282 100%);
        border-bottom: 3px solid #E2E8F0;
        padding: 24px 32px;
        border-radius: 12px;
        color: white;
        text-align: center;
        margin-bottom: 28px;
        box-shadow: 0 2px 12px rgba(0,0,0,0.08);
    }
    .header-banner h1 {
        font-size: 1.6rem;
        font-weight: 600;
        margin: 0 0 8px 0;
        letter-spacing: -0.02em;
    }
    .header-banner h2 {
        font-size: 1rem;
        font-weight: 400;
        color: #CBD5E0;
        margin: 0;
    }
    .header-banner .star {
        font-size: 1.2rem;
        color: #FBD38D;
        letter-spacing: 4px;
        margin-bottom: 8px;
    }

    /* Sidebar */
    [data-testid="stSidebar"] {
        background: #1A2C3E !important;
        border-right: 1px solid #2D4A6E;
    }
    [data-testid="stSidebar"] * {
        color: #E2E8F0 !important;
    }
    [data-testid="stSidebar"] .stTextInput input {
        background: white !important;
        color: #1A202C !important;
        border: 1px solid #CBD5E0 !important;
        border-radius: 8px;
    }
    [data-testid="stSidebar"] .stTextInput input::placeholder {
        color: #718096 !important;
    }
    [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 {
        color: #FBD38D !important;
        border-bottom: 1px solid #2D4A6E;
        padding-bottom: 8px;
    }

    /* Tabs */
    [data-testid="stTabs"] [data-baseweb="tab-list"] {
        background: white;
        border-radius: 10px 10px 0 0;
        padding: 6px 12px 0;
        gap: 8px;
        border-bottom: 2px solid #E2E8F0;
    }
    [data-testid="stTabs"] [data-baseweb="tab"] {
        font-weight: 500;
        color: #4A5568;
        border-radius: 8px 8px 0 0;
        padding: 10px 24px;
    }
    [data-testid="stTabs"] [aria-selected="true"] {
        background: #2C5282 !important;
        color: white !important;
    }

    /* Buttons */
    .stButton > button {
        background: #2C5282;
        color: white;
        border: none;
        border-radius: 8px;
        font-weight: 500;
        padding: 0.5rem 1.2rem;
        transition: all 0.2s;
    }
    .stButton > button:hover {
        background: #1A3A5C;
        transform: translateY(-1px);
        box-shadow: 0 2px 8px rgba(0,0,0,0.15);
    }

    /* Delete Button */
    .delete-btn > button {
        background: #C53030;
    }
    .delete-btn > button:hover {
        background: #9B2C2C;
    }

    /* Download Button */
    [data-testid="stDownloadButton"] > button {
        background: #38A169 !important;
    }
    [data-testid="stDownloadButton"] > button:hover {
        background: #2F855A !important;
    }

    /* Form Container */
    [data-testid="stForm"] {
        background: white;
        border-radius: 12px;
        padding: 28px;
        border: 1px solid #E2E8F0;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }

    /* Metric Cards */
    [data-testid="stMetric"] {
        background: white;
        border-radius: 12px;
        padding: 20px;
        border-left: 4px solid #2C5282;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    [data-testid="stMetricValue"] {
        color: #2C5282 !important;
        font-size: 1.8rem !important;
        font-weight: 700 !important;
    }

    /* Subheader */
    .stSubheader, h3 {
        color: #1A3A5C !important;
        font-weight: 600;
        border-left: 3px solid #FBD38D;
        padding-left: 12px;
    }

    /* Checkbox */
    .stCheckbox label {
        font-weight: 500;
        color: #2D3748;
    }

    /* Warning Box */
    .stAlert {
        border-radius: 10px;
        border-left: 4px solid;
    }

    /* Divider */
    .section-divider {
        height: 2px;
        background: linear-gradient(90deg, #CBD5E0, #2C5282, #CBD5E0);
        margin: 20px 0;
        border-radius: 2px;
    }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# ─────────────────────────────────────────────
# 4. HÀM TIỆN ÍCH
# ─────────────────────────────────────────────
def clean_filename(filename: str) -> str:
    filename = unicodedata.normalize("NFKD", filename).encode("ASCII", "ignore").decode("utf-8")
    filename = filename.replace(" ", "_")
    return re.sub(r"[^\w.\-]", "", filename)

def add_hyperlink_to_cell(cell, text: str, url: str):
    if not url or url.strip() in ("", "nan"):
        cell.text = text
        return
    paragraph = cell.paragraphs[0]
    paragraph.clear()
    part = paragraph.part
    r_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    new_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    rStyle = OxmlElement("w:rStyle")
    rStyle.set(qn("w:val"), "Hyperlink")
    rPr.append(rStyle)
    new_run.append(rPr)
    t = OxmlElement("w:t")
    t.text = text if text and text != "nan" else url
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)

def create_word_report(df: pd.DataFrame, date_str: str) -> bytes:
    doc = Document()
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width
    for attr in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(section, attr, Cm(1.5))

    heading = doc.add_heading("BẢNG TỔNG HỢP ĐĂNG KÝ TIN BÀI", 1)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in heading.runs:
        run.font.color.rgb = RGBColor(27, 58, 92)
        run.font.size = Pt(16)

    p_date = doc.add_paragraph(f"Ngày tổng hợp: {date_str}")
    p_date.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_date.runs[0].font.italic = True

    doc.add_paragraph()

    headers = ["STT", "Tiêu đề", "Người gửi", "Nguồn / File đính kèm", "Đề xuất MXH", "Người đăng"]
    col_widths = (Cm(1.2), Cm(8.5), Cm(3.5), Cm(5.0), Cm(4.0), Cm(4.3))
    table = doc.add_table(rows=1, cols=6)
    table.style = "Table Grid"
    table.autofit = False

    for i, col in enumerate(table.columns):
        col.width = col_widths[i]

    hdr_row = table.rows[0]
    hdr_row.height = Cm(0.9)
    for i, (cell, text) in enumerate(zip(hdr_row.cells, headers)):
        cell.text = text
        run = cell.paragraphs[0].runs[0]
        run.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
        run.font.size = Pt(11)
        tc_pr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), "2C5282")
        tc_pr.append(shd)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    for idx, row in enumerate(df.itertuples(), 1):
        row_cells = table.add_row().cells
        row_cells[0].text = str(idx)
        row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        row_cells[1].text = str(row.tieu_de)
        row_cells[2].text = str(row.nguoi_gui)

        source = str(row.nguon_tin) if pd.notna(row.nguon_tin) else ""
        link = str(getattr(row, "duong_dan", "")) if pd.notna(getattr(row, "duong_dan", None)) else ""
        is_shared = "sưu tầm" in source.lower() or "đăng lại" in source.lower()

        if is_shared and link and link.lower() not in ("", "nan"):
            first_link = link.splitlines()[0].strip()
            display_text = source if source and source != "nan" else first_link
            add_hyperlink_to_cell(row_cells[3], display_text, first_link)
        else:
            row_cells[3].text = source

        mxh = ["Đăng Web"]
        if getattr(row, "dang_facebook", False):
            mxh.append("Facebook")
        if getattr(row, "dang_zalo", False):
            mxh.append("Zalo OA")
        row_cells[4].text = ", ".join(mxh)
        row_cells[5].text = str(row.nguoi_dang) if pd.notna(getattr(row, "nguoi_dang", None)) else ""

        if idx % 2 == 0:
            for cell in row_cells:
                tc_pr = cell._tc.get_or_add_tcPr()
                shd = OxmlElement("w:shd")
                shd.set(qn("w:val"), "clear")
                shd.set(qn("w:color"), "auto")
                shd.set(qn("w:fill"), "F7FAFC")
                tc_pr.append(shd)

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def submit_article(sender, is_new, source_label, note, new_title, files, shared_title, shared_link):
    if not sender.strip():
        st.error("⛔ Vui lòng điền Họ và tên!")
        return

    sender_clean = sender.strip()
    note_clean = note.strip() if note else ""

    if is_new:
        if not new_title.strip():
            st.error("⛔ Vui lòng nhập Tiêu đề bài viết!")
            return
        if not files:
            st.error("⛔ Chưa tải file đính kèm!")
            return

        links = []
        for f in files:
            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
            safe_name = f"{timestamp}_{clean_filename(f.name)}"
            try:
                supabase.storage.from_("tin_bai").upload(safe_name, f.read())
                url = supabase.storage.from_("tin_bai").get_public_url(safe_name)
                links.append(url)
            except Exception as e:
                st.error(f"Lỗi upload {f.name}: {e}")

        if links:
            supabase.table("dang_ky_tin_bai").insert({
                "nguoi_gui": sender_clean,
                "tieu_de": new_title.strip(),
                "nguon_tin": source_label,
                "duong_dan": "\n".join(links),
                "ghi_chu": note_clean,
            }).execute()
            st.success(f"✅ Đã gửi thành công! ({len(files)} file đính kèm)")
    else:
        if not shared_title.strip():
            st.error("⛔ Vui lòng nhập tiêu đề bài sưu tầm!")
            return
        if not shared_link.strip():
            st.error("⛔ Vui lòng nhập đường dẫn bài sưu tầm!")
            return

        try:
            supabase.table("dang_ky_tin_bai").insert({
                "nguoi_gui": sender_clean,
                "tieu_de": shared_title.strip(),
                "nguon_tin": source_label,
                "duong_dan": shared_link.strip(),
                "ghi_chu": note_clean,
            }).execute()
            st.success("✅ Đã gửi đăng ký bài sưu tầm thành công!")
        except Exception as e:
            st.error(f"Lỗi: {e}")

def delete_articles(ids_to_delete):
    if not ids_to_delete:
        st.warning("Chưa chọn bài nào để xóa!")
        return False
    try:
        for article_id in ids_to_delete:
            supabase.table("dang_ky_tin_bai").delete().eq("id", article_id).execute()
        st.success(f"✅ Đã xóa {len(ids_to_delete)} bài viết!")
        return True
    except Exception as e:
        st.error(f"Lỗi xóa: {e}")
        return False

def update_article(article_id, new_title, new_sender):
    try:
        supabase.table("dang_ky_tin_bai").update({
            "tieu_de": new_title,
            "nguoi_gui": new_sender,
        }).eq("id", article_id).execute()
        st.success("✅ Đã cập nhật bài viết!")
        return True
    except Exception as e:
        st.error(f"Lỗi cập nhật: {e}")
        return False

# ─────────────────────────────────────────────
# 5. HEADER
# ─────────────────────────────────────────────
st.markdown("""
    <div class="header-banner">
        <span class="star">★ ★ ★ ★ ★</span>
        <h1>Hệ thống Quản lý Tin bài</h1>
        <h2>Ban Tuyên giáo và Dân vận Tỉnh ủy Tuyên Quang</h2>
    </div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# 6. SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🔐 Quản trị hệ thống")
    admin_password = st.text_input("Mật khẩu quản trị:", type="password", placeholder="Nhập mật khẩu...")
    is_admin = admin_password == ADMIN_PASS
    
    if is_admin:
        st.success("✓ Đã xác thực")
    elif admin_password:
        st.error("✗ Sai mật khẩu")

    st.markdown("---")
    st.markdown("""
        <div style='font-size:0.75rem; color:#94A3B8; line-height:1.6;'>
            <b>📌 Hướng dẫn:</b><br>
            • Tab 1: Đăng ký tin bài<br>
            • Tab 2: Duyệt & xuất Word<br>
            • Tab 3: Thống kê<br>
        </div>
    """, unsafe_allow_html=True)

# ─────────────────────────────────────────────
# 7. TABS
# ─────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs([
    "✍️ Đăng ký tin bài",
    "📋 Tổng hợp & Duyệt",
    "📊 Thống kê",
])

# ========== TAB 1: ĐĂNG KÝ ==========
with tab1:
    st.subheader("Gửi thông tin đăng ký bài viết")
    
    article_type = st.radio(
        "Loại tin bài:",
        ["✏️ Viết mới", "🔗 Đề nghị đăng lại (Sưu tầm)"],
        horizontal=True,
    )
    is_new_article = "Viết mới" in article_type

    st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)

    with st.form("form_dang_ky", clear_on_submit=True):
        sender_name = st.text_input("👤 Họ và tên người gửi *", placeholder="Nhập họ tên đầy đủ...")
        
        new_title = ""
        uploaded_files = []
        shared_title = ""
        shared_link = ""

        if is_new_article:
            st.markdown("**📄 Bài viết mới**")
            new_title = st.text_input("Tiêu đề bài viết *", placeholder="Nhập tiêu đề...")
            uploaded_files = st.file_uploader(
                "📎 Tải file đính kèm (nhiều file):",
                type=["doc", "docx", "pdf", "png", "jpg"],
                accept_multiple_files=True,
                help="Có thể chọn nhiều file cùng lúc"
            )
        else:
            st.markdown("**🔗 Bài sưu tầm**")
            shared_title = st.text_input("Tiêu đề bài sưu tầm *", placeholder="Nhập tiêu đề...")
            shared_link = st.text_input("Đường dẫn (URL) *", placeholder="https://...")

        note = st.text_area("💬 Ghi chú (không bắt buộc)", placeholder="Thông tin thêm...")

        submitted = st.form_submit_button("📨 Gửi đăng ký", use_container_width=True)
        
        if submitted:
            submit_article(
                sender=sender_name,
                is_new=is_new_article,
                source_label="Viết mới" if is_new_article else "Đề nghị đăng lại (Sưu tầm)",
                note=note,
                new_title=new_title,
                files=uploaded_files,
                shared_title=shared_title,
                shared_link=shared_link
            )

# ========== TAB 2: TỔNG HỢP & DUYỆT ==========
with tab2:
    if not is_admin:
        st.warning("🔒 Vui lòng nhập mật khẩu quản trị ở thanh bên trái để sử dụng chức năng này.")
    else:
        st.subheader("Quản lý tin bài")
        st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)

        # Session state cho edit modal
        if "edit_mode" not in st.session_state:
            st.session_state.edit_mode = False
            st.session_state.edit_id = None
            st.session_state.edit_title = ""
            st.session_state.edit_sender = ""

        # Chọn ngày
        selected_date = st.date_input("📅 Chọn ngày:", datetime.now().date())
        date_str = selected_date.strftime("%d/%m/%Y")
        
        # Lấy dữ liệu
        result = supabase.table("dang_ky_tin_bai").select("*").eq("ngay_dang_ky", selected_date.isoformat()).execute()
        
        if not result.data:
            st.info(f"ℹ️ Không có tin bài nào ngày {date_str}.")
        else:
            df = pd.DataFrame(result.data)
            for col in ("nguoi_dang", "dang_facebook", "dang_zalo"):
                if col not in df.columns:
                    df[col] = False if col.startswith("dang") else ""
            df["nguoi_dang"] = df["nguoi_dang"].fillna("")
            df["dang_facebook"] = df["dang_facebook"].fillna(False)
            df["dang_zalo"] = df["dang_zalo"].fillna(False)

            st.markdown(f"**📊 Tổng số bài: `{len(df)}` bài - Ngày {date_str}**")

            # Thêm checkbox cho xóa
            df_with_checkbox = df.copy()
            df_with_checkbox.insert(0, "Chọn", False)
            
            edited_df = st.data_editor(
                df_with_checkbox,
                column_config={
                    "Chọn": st.column_config.CheckboxColumn("☑️", default=False),
                    "dang_facebook": st.column_config.CheckboxColumn("📘 Facebook", default=False),
                    "dang_zalo": st.column_config.CheckboxColumn("💬 Zalo", default=False),
                    "duong_dan": st.column_config.TextColumn("🔗 Link / File"),
                    "nguoi_dang": st.column_config.TextColumn("👤 Người đăng"),
                    "id": st.column_config.NumberColumn("ID", disabled=True),
                },
                disabled=["nguoi_gui", "tieu_de", "nguon_tin", "id", "ghi_chu", "ngay_dang_ky"],
                hide_index=True,
                use_container_width=True,
            )
            
            col1, col2, col3, col4 = st.columns([1, 1, 1, 1])
            
            with col1:
                if st.button("💾 Lưu trạng thái MXH", use_container_width=True):
                    success_count = 0
                    for _, row in edited_df.iterrows():
                        row_id = row.get("id")
                        if row_id:
                            try:
                                supabase.table("dang_ky_tin_bai").update({
                                    "dang_facebook": bool(row["dang_facebook"]),
                                    "dang_zalo": bool(row["dang_zalo"]),
                                    "nguoi_dang": str(row.get("nguoi_dang", "")),
                                }).eq("id", row_id).execute()
                                success_count += 1
                            except Exception:
                                pass
                    if success_count > 0:
                        st.success(f"✅ Đã lưu {success_count} bài!")
                        st.rerun()
            
            with col2:
                word_bytes = create_word_report(edited_df, date_str)
                st.download_button(
                    "📥 Xuất Word",
                    data=word_bytes,
                    file_name=f"TinBai_{selected_date.strftime('%Y%m%d')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )
            
            # Nút xóa bài được chọn
            with col3:
                selected_ids = edited_df[edited_df["Chọn"] == True]["id"].tolist()
                if selected_ids:
                    st.markdown('<div class="delete-btn">', unsafe_allow_html=True)
                    if st.button(f"🗑️ Xóa {len(selected_ids)} bài đã chọn", use_container_width=True):
                        if delete_articles(selected_ids):
                            st.rerun()
                    st.markdown('</div>', unsafe_allow_html=True)
            
            # Nút sửa bài
            with col4:
                if len(selected_ids) == 1:
                    if st.button("✏️ Sửa bài đã chọn", use_container_width=True):
                        selected_row = edited_df[edited_df["Chọn"] == True].iloc[0]
                        st.session_state.edit_mode = True
                        st.session_state.edit_id = selected_row["id"]
                        st.session_state.edit_title = selected_row["tieu_de"]
                        st.session_state.edit_sender = selected_row["nguoi_gui"]
                        st.rerun()
                elif len(selected_ids) > 1:
                    st.info("⚠️ Chỉ được chọn 1 bài để sửa")
            
            # Modal sửa bài
            if st.session_state.edit_mode:
                with st.expander("✏️ Đang sửa bài viết", expanded=True):
                    new_title = st.text_input("Tiêu đề mới:", value=st.session_state.edit_title)
                    new_sender = st.text_input("Người gửi mới:", value=st.session_state.edit_sender)
                    
                    col_ok, col_cancel = st.columns(2)
                    with col_ok:
                        if st.button("✅ Cập nhật"):
                            if update_article(st.session_state.edit_id, new_title, new_sender):
                                st.session_state.edit_mode = False
                                st.rerun()
                    with col_cancel:
                        if st.button("❌ Hủy"):
                            st.session_state.edit_mode = False
                            st.rerun()

        # Khu vực nguy hiểm
        st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
        with st.expander("⚠️ Khu vực nguy hiểm", expanded=False):
            st.warning("Thao tác này sẽ XÓA TOÀN BỘ dữ liệu và không thể hoàn tác!")
            if st.button("🔥 Xóa toàn bộ tin bài", type="primary"):
                try:
                    supabase.table("dang_ky_tin_bai").delete().neq("id", 0).execute()
                    st.success("Đã xóa toàn bộ dữ liệu!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Lỗi: {e}")

# ========== TAB 3: THỐNG KÊ ==========
with tab3:
    st.subheader("Báo cáo thống kê")
    st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)

    all_data = supabase.table("dang_ky_tin_bai").select("*").execute()
    
    if not all_data.data:
        st.info("Chưa có dữ liệu.")
    else:
        df_all = pd.DataFrame(all_data.data)
        df_all["ngay_dang_ky"] = pd.to_datetime(df_all["ngay_dang_ky"], errors="coerce")
        df_all = df_all.dropna(subset=["ngay_dang_ky"])
        df_all["Năm"] = df_all["ngay_dang_ky"].dt.year

        col_year, col_type, _ = st.columns(3)
        
        with col_year:
            years = sorted(df_all["Năm"].unique().tolist(), reverse=True)
            selected_year = st.selectbox("📅 Năm:", years)
        
        df_year = df_all[df_all["Năm"] == selected_year].copy()
        
        with col_type:
            filter_type = st.selectbox("🔍 Lọc theo:", ["Cả năm", "Theo Quý", "Theo Tháng", "Theo Tuần", "Theo Ngày"])
        
        df_filtered = df_year.copy()
        if filter_type == "Theo Quý":
            df_year["Quý"] = df_year["ngay_dang_ky"].dt.quarter
            quarters = sorted(df_year["Quý"].unique())
            if quarters:
                selected = st.selectbox("Chọn:", [f"Quý {q}" for q in quarters])
                df_filtered = df_year[df_year["Quý"] == int(selected.split()[-1])]
        elif filter_type == "Theo Tháng":
            df_year["Tháng"] = df_year["ngay_dang_ky"].dt.month
            months = sorted(df_year["Tháng"].unique())
            if months:
                selected = st.selectbox("Chọn:", [f"Tháng {m}" for m in months])
                df_filtered = df_year[df_year["Tháng"] == int(selected.split()[-1])]
        elif filter_type == "Theo Tuần":
            df_year["Tuần"] = df_year["ngay_dang_ky"].dt.isocalendar().week
            weeks = sorted(df_year["Tuần"].unique())
            if weeks:
                selected = st.selectbox("Chọn:", [f"Tuần {w}" for w in weeks])
                df_filtered = df_year[df_year["Tuần"] == int(selected.split()[-1])]
        elif filter_type == "Theo Ngày":
            dates = sorted(df_year["ngay_dang_ky"].dt.date.unique(), reverse=True)
            if dates:
                selected = st.selectbox("Chọn:", dates, format_func=lambda x: x.strftime("%d/%m/%Y"))
                df_filtered = df_year[df_year["ngay_dang_ky"].dt.date == selected]

        if df_filtered.empty:
            st.warning("Không có dữ liệu trong khoảng thời gian này.")
        else:
            col_m1, col_m2, col_m3 = st.columns(3)
            total_articles = len(df_filtered)
            new_articles = len(df_filtered[df_filtered["nguon_tin"].str.contains("Viết mới", case=False, na=False)])
            shared_articles = total_articles - new_articles
            
            col_m1.metric("📰 Tổng số bài", total_articles)
            col_m2.metric("✏️ Bài viết mới", new_articles)
            col_m3.metric("🔗 Bài sưu tầm", shared_articles)

            # Biểu đồ
            chart_data = df_filtered.groupby(["nguoi_gui", "nguon_tin"]).size().reset_index(name="Số lượng")
            fig = px.bar(
                chart_data,
                x="nguoi_gui",
                y="Số lượng",
                color="nguon_tin",
                barmode="group",
                text="Số lượng",
                color_discrete_sequence=["#2C5282", "#C53030"],
                labels={"nguoi_gui": "Người gửi", "nguon_tin": "Loại bài"},
                title="Phân bổ tin bài theo cán bộ"
            )
            fig.update_traces(textposition="outside")
            fig.update_layout(
                plot_bgcolor="white",
                paper_bgcolor="white",
                title_font=dict(size=14, color="#1A3A5C"),
                font=dict(family="Inter, Segoe UI"),
                yaxis=dict(gridcolor="#E2E8F0")
            )
            st.plotly_chart(fig, use_container_width=True)
