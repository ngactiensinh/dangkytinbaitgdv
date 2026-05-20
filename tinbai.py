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
# 1. KẾT NỐI SUPABASE
# ─────────────────────────────────────────────
URL       = st.secrets["SUPABASE_URL"]
KEY       = st.secrets["SUPABASE_KEY"]
ADMIN_PASS = st.secrets.get("ADMIN_PASS", "141983")
supabase  = create_client(URL, KEY)

try:
    supabase.table("thong_ke_truy_cap").insert({"ten_app": "Đăng ký Tin bài"}).execute()
except Exception:
    pass


# ─────────────────────────────────────────────
# 2. CSS – PHONG CÁCH CHÍNH TRỊ / HIỆN ĐẠI
# ─────────────────────────────────────────────
CUSTOM_CSS = """
<style>
    /* ── Màu nền & font ── */
    html, body, [data-testid="stAppViewContainer"] {
        background-color: #F0F4F8;
        font-family: 'Segoe UI', 'Arial', sans-serif;
    }

    /* ── Header banner ── */
    .header-banner {
        background: linear-gradient(135deg, #003466 0%, #004B87 60%, #005C9E 100%);
        border-bottom: 4px solid #C0392B;
        padding: 22px 32px;
        border-radius: 10px;
        color: white;
        text-align: center;
        box-shadow: 0 4px 18px rgba(0,52,102,0.18);
        margin-bottom: 24px;
        position: relative;
        overflow: hidden;
    }
    .header-banner::before {
        content: "";
        position: absolute;
        top: -30px; left: -30px;
        width: 120px; height: 120px;
        border-radius: 50%;
        background: rgba(255,255,255,0.05);
    }
    .header-banner::after {
        content: "";
        position: absolute;
        bottom: -20px; right: -20px;
        width: 90px; height: 90px;
        border-radius: 50%;
        background: rgba(255,255,255,0.06);
    }
    .header-banner h1 {
        font-size: 1.65rem;
        font-weight: 800;
        letter-spacing: 0.05em;
        text-transform: uppercase;
        margin: 0 0 6px 0;
        text-shadow: 0 2px 6px rgba(0,0,0,0.25);
    }
    .header-banner h2 {
        font-size: 1.05rem;
        font-weight: 400;
        color: #FFD700;
        margin: 0;
        letter-spacing: 0.03em;
    }
    .header-banner .star {
        font-size: 1.4rem;
        color: #FFD700;
        letter-spacing: 6px;
        display: block;
        margin-bottom: 10px;
    }

    /* ── Sidebar ── */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #002147 0%, #003466 100%) !important;
        border-right: 3px solid #C0392B;
    }
    [data-testid="stSidebar"] * { color: #E8F0FE !important; }
    [data-testid="stSidebar"] .stTextInput input {
        background: rgba(255,255,255,0.1) !important;
        border: 1px solid rgba(255,215,0,0.4) !important;
        color: white !important;
        border-radius: 6px;
    }
    [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 {
        color: #FFD700 !important;
        border-bottom: 1px solid rgba(255,215,0,0.3);
        padding-bottom: 8px;
    }

    /* ── Tab ── */
    [data-testid="stTabs"] [data-baseweb="tab-list"] {
        background: white;
        border-radius: 8px 8px 0 0;
        padding: 4px 8px 0 8px;
        border-bottom: 2px solid #003466;
        gap: 4px;
    }
    [data-testid="stTabs"] [data-baseweb="tab"] {
        font-weight: 600;
        color: #555;
        border-radius: 6px 6px 0 0;
        padding: 10px 20px;
        font-size: 0.92rem;
    }
    [data-testid="stTabs"] [aria-selected="true"] {
        background: #003466 !important;
        color: white !important;
        border-bottom: 2px solid #FFD700 !important;
    }

    /* ── Nút bấm ── */
    .stButton > button {
        background: linear-gradient(135deg, #003466, #005C9E);
        color: white;
        border: none;
        border-radius: 6px;
        font-weight: 600;
        padding: 0.5rem 1.4rem;
        font-size: 0.9rem;
        transition: all 0.2s ease;
        letter-spacing: 0.02em;
        box-shadow: 0 2px 8px rgba(0,52,102,0.2);
    }
    .stButton > button:hover {
        background: linear-gradient(135deg, #C0392B, #E74C3C);
        box-shadow: 0 4px 14px rgba(192,57,43,0.3);
        transform: translateY(-1px);
    }

    /* ── Nút tải xuống ── */
    [data-testid="stDownloadButton"] > button {
        background: linear-gradient(135deg, #1a6e36, #27AE60) !important;
        color: white !important;
        border-radius: 6px !important;
        font-weight: 700 !important;
        border: none !important;
        box-shadow: 0 2px 8px rgba(26,110,54,0.25);
    }
    [data-testid="stDownloadButton"] > button:hover {
        background: linear-gradient(135deg, #145a2c, #1E8449) !important;
        transform: translateY(-1px);
    }

    /* ── Form container ── */
    [data-testid="stForm"] {
        background: white;
        border-radius: 10px;
        padding: 24px;
        border: 1px solid #D6E4F0;
        box-shadow: 0 2px 12px rgba(0,52,102,0.07);
    }

    /* ── Metric cards ── */
    [data-testid="stMetric"] {
        background: white;
        border-radius: 10px;
        padding: 18px 20px;
        border-left: 5px solid #003466;
        box-shadow: 0 2px 10px rgba(0,52,102,0.08);
    }
    [data-testid="stMetricValue"] {
        color: #C0392B !important;
        font-size: 2rem !important;
        font-weight: 800 !important;
    }

    /* ── Subheader ── */
    .stSubheader, h3 {
        color: #003466 !important;
        font-weight: 700;
        border-left: 4px solid #C0392B;
        padding-left: 12px;
    }

    /* ── Divider màu ── */
    hr { border-color: #D6E4F0; }

    /* ── Thông báo / cảnh báo ── */
    [data-testid="stAlert"] {
        border-radius: 8px;
        border-left: 4px solid;
    }

    /* ── Radio button ── */
    [data-testid="stRadio"] label {
        font-weight: 500;
        color: #003466;
    }

    /* ── Badge trạng thái ── */
    .badge-admin {
        display: inline-block;
        background: #FFD700;
        color: #003466;
        font-weight: 700;
        padding: 3px 12px;
        border-radius: 20px;
        font-size: 0.78rem;
        letter-spacing: 0.04em;
        margin-left: 8px;
    }

    /* ── Section divider ── */
    .section-divider {
        height: 3px;
        background: linear-gradient(90deg, #003466, #C0392B, #FFD700);
        border-radius: 2px;
        margin: 20px 0;
    }
</style>
"""


# ─────────────────────────────────────────────
# 3. TIỆN ÍCH
# ─────────────────────────────────────────────
def lam_sach_ten_file(ten_file: str) -> str:
    """Loại bỏ dấu và ký tự đặc biệt khỏi tên file trước khi upload."""
    ten_file = unicodedata.normalize("NFKD", ten_file).encode("ASCII", "ignore").decode("utf-8")
    ten_file = ten_file.replace(" ", "_")
    ten_file = re.sub(r"[^\w.\-]", "", ten_file)
    return ten_file


def them_hyperlink_vao_cell(cell, text: str, url: str):
    """
    Chèn văn bản có hyperlink vào một ô bảng Word.
    Nếu url rỗng thì chỉ in text thuần.
    """
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


# ─────────────────────────────────────────────
# 4. TẠO FILE WORD
# ─────────────────────────────────────────────
def tao_file_word(df: pd.DataFrame, ngay_thang: str) -> bytes:
    doc = Document()
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width
    for attr in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(section, attr, Cm(1.5))

    heading = doc.add_heading("BẢNG TỔNG HỢP ĐĂNG KÝ TIN BÀI", 1)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in heading.runs:
        run.font.color.rgb = RGBColor(0, 52, 102)
        run.font.size = Pt(16)

    p_ngay = doc.add_paragraph(f"Ngày tổng hợp: {ngay_thang}")
    p_ngay.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_ngay.runs[0].font.italic = True
    p_ngay.runs[0].font.color.rgb = RGBColor(80, 80, 80)

    doc.add_paragraph()

    headers  = ["STT", "Tiêu đề", "Người gửi", "Nguồn / File đính kèm", "Đề xuất MXH", "Người đăng"]
    col_widths = (Cm(1.2), Cm(8.5), Cm(3.5), Cm(5.0), Cm(4.0), Cm(4.3))

    table = doc.add_table(rows=1, cols=6)
    table.style = "Table Grid"
    table.autofit = False

    for i, col in enumerate(table.columns):
        col.width = col_widths[i]

    hdr_row = table.rows[0]
    hdr_row.height = Cm(0.9)
    for i, (cell, text) in enumerate(zip(hdr_row.cells, headers)):
        cell.width = col_widths[i]
        cell.text = text
        run = cell.paragraphs[0].runs[0]
        run.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
        run.font.size = Pt(11)
        tc_pr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), "003466")
        tc_pr.append(shd)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    for idx, row in enumerate(df.itertuples(), 1):
        row_cells = table.add_row().cells

        row_cells[0].text = str(idx)
        row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        row_cells[1].text = str(row.tieu_de)
        row_cells[2].text = str(row.nguoi_gui)

        nguon_str  = str(row.nguon_tin) if pd.notna(row.nguon_tin) else ""
        duong_dan  = str(row.duong_dan) if pd.notna(getattr(row, "duong_dan", None)) else ""
        is_suu_tam = "sưu tầm" in nguon_str.lower() or "đăng lại" in nguon_str.lower()

        if is_suu_tam and duong_dan and duong_dan.lower() not in ("", "nan"):
            first_link = duong_dan.splitlines()[0].strip()
            hien_thi   = nguon_str if nguon_str and nguon_str != "nan" else first_link
            them_hyperlink_vao_cell(row_cells[3], hien_thi, first_link)
        else:
            row_cells[3].text = nguon_str

        mxh = ["Đăng Web"]
        if getattr(row, "dang_facebook", False): mxh.append("Facebook")
        if getattr(row, "dang_zalo",    False): mxh.append("Zalo OA")
        row_cells[4].text = ", ".join(mxh)
        row_cells[5].text = str(row.nguoi_dang) if pd.notna(getattr(row, "nguoi_dang", None)) else ""

        for i, cell in enumerate(row_cells):
            cell.width = col_widths[i]
            for para in cell.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(10)

        if idx % 2 == 0:
            for cell in row_cells:
                tc_pr = cell._tc.get_or_add_tcPr()
                shd = OxmlElement("w:shd")
                shd.set(qn("w:val"),   "clear")
                shd.set(qn("w:color"), "auto")
                shd.set(qn("w:fill"),  "EAF2FF")
                tc_pr.append(shd)

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()


# ─────────────────────────────────────────────
# 5. GIAO DIỆN CHÍNH
# ─────────────────────────────────────────────
def main():
    st.set_page_config(
        page_title="Hệ thống Quản lý Tin bài – Ban Tuyên giáo",
        page_icon="📰",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

    # Header banner
    st.markdown("""
        <div class="header-banner">
            <span class="star">★ ★ ★ ★ ★</span>
            <h1>Hệ thống quản lý tin bài đăng trang Thông tin điện tử</h1>
            <h2>Ban Tuyên giáo và Dân vận Tỉnh ủy Tuyên Quang</h2>
        </div>
    """, unsafe_allow_html=True)

    # ── Sidebar ──
    with st.sidebar:
        st.markdown("### 🔐 Quản trị hệ thống")
        mat_khau = st.text_input("Mật khẩu quản trị:", type="password", placeholder="Nhập mật khẩu…")
        is_admin = mat_khau == ADMIN_PASS
        if is_admin:
            st.success("✔ Đã xác thực quản trị viên")
        elif mat_khau:
            st.error("✘ Mật khẩu không đúng")

        st.markdown("---")
        st.markdown("""
            <div style='font-size:0.78rem; color:#a0b4cc; line-height:1.7;'>
                <b style='color:#FFD700'>📌 Hướng dẫn nhanh:</b><br>
                • Tab 1: Cán bộ đăng ký tin bài<br>
                • Tab 2: Quản trị duyệt & xuất Word<br>
                • Tab 3: Thống kê biểu đồ<br>
            </div>
        """, unsafe_allow_html=True)

    # ── Tabs ──
    tab1, tab2, tab3 = st.tabs([
        "✍️  Đăng ký tin bài",
        "📋  Tổng hợp & Duyệt",
        "📊  Thống kê & Biểu đồ",
    ])

    # ══════════════════════════════════════════
    # TAB 1 – ĐĂNG KÝ TIN BÀI
    # ══════════════════════════════════════════
    with tab1:
        st.subheader("Gửi thông tin đăng ký bài viết")
        
        # CHÚ Ý: CHUYỂN NÚT RADIO RA KHỎI FORM ĐỂ UI UPDATE NGAY LẬP TỨC
        nguon = st.radio(
            "📂 Loại tin bài:",
            ["✏️ Viết mới", "🔗 Đề nghị đăng lại (Sưu tầm)"],
            horizontal=True,
        )
        la_viet_moi = "Viết mới" in nguon

        st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)

        with st.form("form_dang_ky", clear_on_submit=True):
            nguoi_gui = st.text_input(
                "👤 Họ và tên người gửi *",
                placeholder="Nhập họ và tên đầy đủ của đồng chí…"
            )

            # Khởi tạo các biến để tránh NameError khi lưu
            tieu_de_viet_moi = ""
            file_uploads = []
            tieu_de_suu_tam = ""
            link_suu_tam = ""

            if la_viet_moi:
                st.markdown("**📄 Thông tin bài viết mới**")
                tieu_de_viet_moi = st.text_input(
                    "Tiêu đề bài viết *",
                    placeholder="Ví dụ: Infographic về chuyển đổi số năm 2025…"
                )
                file_uploads = st.file_uploader(
                    "📎 Tải lên file đính kèm (chọn nhiều file):",
                    type=["doc", "docx", "pdf", "png", "jpg"],
                    accept_multiple_files=True,
                )
            else:
                st.markdown("**🔗 Thông tin bài sưu tầm / đăng lại**")
                tieu_de_suu_tam = st.text_input(
                    "Tiêu đề bài sưu tầm *", 
                    placeholder="Nhập tiêu đề bài viết..."
                )
                link_suu_tam = st.text_input(
                    "Đường dẫn (URL bài gốc) *", 
                    placeholder="Dán link bài gốc vào đây (VD: https://baotuyenquang.com.vn/...)"
                )

            ghi_chu = st.text_area("💬 Ghi chú thêm:", placeholder="Không bắt buộc…")

            btn_gui = st.form_submit_button("📨 Gửi đăng ký tin bài", use_container_width=True)

            if btn_gui:
                _xu_ly_gui(
                    nguoi_gui   = nguoi_gui,
                    la_viet_moi = la_viet_moi,
                    nguon_label = "Viết mới" if la_viet_moi else "Đề nghị đăng lại (Sưu tầm)",
                    ghi_chu     = ghi_chu,
                    tieu_de_viet_moi = tieu_de_viet_moi,
                    file_uploads     = file_uploads,
                    tieu_de_suu_tam  = tieu_de_suu_tam,
                    link_suu_tam     = link_suu_tam
                )

    # ══════════════════════════════════════════
    # TAB 2 – TỔNG HỢP & DUYỆT
    # ══════════════════════════════════════════
    with tab2:
        if not is_admin:
            st.warning("🔒 Vui lòng nhập mật khẩu Quản trị ở thanh bên trái để sử dụng chức năng này.")
            st.stop()

        st.subheader("Bảng tổng hợp & phê duyệt trình duyệt")
        st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)

        ngay_xem     = st.date_input("📅 Chọn ngày cần xem / xuất báo cáo:", datetime.now().date())
        ngay_xem_str = ngay_xem.strftime("%d/%m/%Y")

        res = supabase.table("dang_ky_tin_bai").select("*").eq("ngay_dang_ky", ngay_xem.isoformat()).execute()

        if not res.data:
            st.info(f"ℹ️ Không có tin bài nào được đăng ký trong ngày **{ngay_xem_str}**.")
        else:
            df_ngay = pd.DataFrame(res.data)
            for col in ("nguoi_dang", "dang_facebook", "dang_zalo"):
                if col not in df_ngay.columns:
                    df_ngay[col] = False if col.startswith("dang") else ""
            df_ngay["nguoi_dang"]   = df_ngay["nguoi_dang"].fillna("")
            df_ngay["dang_facebook"]= df_ngay["dang_facebook"].fillna(False)
            df_ngay["dang_zalo"]    = df_ngay["dang_zalo"].fillna(False)

            st.markdown(f"**Tổng số bài:** `{len(df_ngay)}` bài đăng ký ngày {ngay_xem_str}")

            edited_df = st.data_editor(
                df_ngay,
                column_config={
                    "dang_facebook": st.column_config.CheckboxColumn("📘 Đăng FB",   default=False),
                    "dang_zalo":     st.column_config.CheckboxColumn("💬 Đăng Zalo", default=False),
                    "duong_dan":     st.column_config.TextColumn("🔗 Link / File (bôi đen để copy)"),
                    "nguoi_dang":    st.column_config.TextColumn("👤 Người đăng"),
                },
                disabled=["nguoi_gui", "tieu_de", "nguon_tin"],
                hide_index=True,
                use_container_width=True,
            )

            c1, c2, c3 = st.columns([2, 2, 1])
            with c1:
                if st.button("💾 Lưu trạng thái Mạng xã hội", use_container_width=True):
                    for _, row in edited_df.iterrows():
                        supabase.table("dang_ky_tin_bai").update({
                            "dang_facebook": bool(row["dang_facebook"]),
                            "dang_zalo":     bool(row["dang_zalo"]),
                            "nguoi_dang":    str(row.get("nguoi_dang", "")),
                        }).eq("id", row["id"]).execute()
                    st.success("✅ Đã lưu thành công!")

            with c2:
                word_bytes = tao_file_word(edited_df, ngay_xem_str)
                st.download_button(
                    "📥 Xuất file Word trình duyệt",
                    data=word_bytes,
                    file_name=f"TinBai_{ngay_xem.strftime('%Y%m%d')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )

        st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
        with st.expander("⚠️ Khu vực nguy hiểm – Quản trị dữ liệu", expanded=False):
            st.warning("Thao tác dưới đây sẽ **xóa toàn bộ** dữ liệu và không thể hoàn tác!")
            if st.button("🗑️ Xóa sạch toàn bộ tin bài (Reset)"):
                try:
                    supabase.table("dang_ky_tin_bai").delete().neq("id", 0).execute()
                    st.success("Đã dọn sạch toàn bộ dữ liệu!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Lỗi: {e}")

    # ══════════════════════════════════════════
    # TAB 3 – THỐNG KÊ & BIỂU ĐỒ
    # ══════════════════════════════════════════
    with tab3:
        st.subheader("Báo cáo thống kê tin bài")
        st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)

        res_all = supabase.table("dang_ky_tin_bai").select("*").execute()
        if not res_all.data:
            st.info("Hệ thống chưa có dữ liệu tin bài.")
        else:
            df_all = pd.DataFrame(res_all.data)
            df_all["ngay_dang_ky"] = pd.to_datetime(df_all["ngay_dang_ky"], errors="coerce")
            df_all = df_all.dropna(subset=["ngay_dang_ky"])
            df_all["Năm"] = df_all["ngay_dang_ky"].dt.year

            col_nam, col_kieu, col_ct = st.columns(3)

            with col_nam:
                ds_nam  = sorted(df_all["Năm"].unique().tolist(), reverse=True)
                chon_nam = st.selectbox("📅 Năm:", ds_nam)

            df_nam = df_all[df_all["Năm"] == chon_nam].copy()

            with col_kieu:
                kieu_loc = st.selectbox("🔍 Lọc theo:", ["Cả năm", "Theo Quý", "Theo Tháng", "Theo Tuần", "Theo Ngày"])

            with col_ct:
                df_loc = df_nam.copy()
                if kieu_loc == "Theo Quý":
                    df_nam["Quý"] = df_nam["ngay_dang_ky"].dt.quarter
                    q_opts = sorted(df_nam["Quý"].unique())
                    if q_opts:
                        chon = st.selectbox("Chọn:", [f"Quý {q}" for q in q_opts])
                        df_loc = df_nam[df_nam["Quý"] == int(chon.split()[-1])]
                elif kieu_loc == "Theo Tháng":
                    df_nam["Tháng"] = df_nam["ngay_dang_ky"].dt.month
                    t_opts = sorted(df_nam["Tháng"].unique())
                    if t_opts:
                        chon = st.selectbox("Chọn:", [f"Tháng {t}" for t in t_opts])
                        df_loc = df_nam[df_nam["Tháng"] == int(chon.split()[-1])]
                elif kieu_loc == "Theo Tuần":
                    df_nam["Tuần"] = df_nam["ngay_dang_ky"].dt.isocalendar().week
                    tu_opts = sorted(df_nam["Tuần"].unique())
                    if tu_opts:
                        chon = st.selectbox("Chọn:", [f"Tuần thứ {t}" for t in tu_opts])
                        df_loc = df_nam[df_nam["Tuần"] == int(chon.split()[-1])]
                elif kieu_loc == "Theo Ngày":
                    ng_opts = sorted(df_nam["ngay_dang_ky"].dt.date.unique(), reverse=True)
                    if ng_opts:
                        chon = st.selectbox("Chọn:", ng_opts, format_func=lambda x: x.strftime("%d/%m/%Y"))
                        df_loc = df_nam[df_nam["ngay_dang_ky"].dt.date == chon]
                else:
                    st.info(f"Toàn bộ năm {chon_nam}")

            st.markdown("---")
            if df_loc.empty:
                st.warning("Không có bài viết nào trong khoảng thời gian đã chọn.")
            else:
                so_tu_viet  = len(df_loc[df_loc["nguon_tin"].str.contains("Viết mới|tự viết",  case=False, na=False)])
                so_suu_tam  = len(df_loc[df_loc["nguon_tin"].str.contains("Sưu tầm|đăng lại", case=False, na=False)])

                m1, m2, m3 = st.columns(3)
                m1.metric("📰 Tổng số bài",   len(df_loc))
                m2.metric("✏️ Bài tự viết",   so_tu_viet)
                m3.metric("🔗 Bài sưu tầm",   so_suu_tam)

                df_bd = df_loc.groupby(["nguoi_gui", "nguon_tin"]).size().reset_index(name="Số lượng")
                fig = px.bar(
                    df_bd,
                    x="nguoi_gui", y="Số lượng",
                    color="nguon_tin",
                    barmode="group",
                    text="Số lượng",
                    color_discrete_sequence=["#003466", "#C0392B"],
                    labels={"nguoi_gui": "Người gửi", "nguon_tin": "Loại bài"},
                )
                fig.update_traces(textposition="outside", marker_line_width=0)
                fig.update_layout(
                    title=dict(text="Biểu đồ phân bổ Tin bài theo Cán bộ", font=dict(size=15, color="#003466")),
                    plot_bgcolor="white",
                    paper_bgcolor="white",
                    yaxis=dict(gridcolor="#EEF2F7"),
                    font=dict(family="Segoe UI, Arial"),
                    legend_title="Loại tin bài",
                )
                st.plotly_chart(fig, use_container_width=True)


# ─────────────────────────────────────────────
# 6. XỬ LÝ GỬI ĐĂNG KÝ (tách ra hàm riêng)
# ─────────────────────────────────────────────
def _xu_ly_gui(nguoi_gui, la_viet_moi, nguon_label, ghi_chu,
               tieu_de_viet_moi, file_uploads, tieu_de_suu_tam, link_suu_tam):
    
    if not nguoi_gui.strip():
        st.error("⛔ Đồng chí vui lòng điền Họ và tên trước khi gửi!")
        return

    nguoi_gui_clean = nguoi_gui.strip()
    ghi_chu_clean   = ghi_chu.strip() if ghi_chu else ""

    if la_viet_moi:
        if not tieu_de_viet_moi.strip():
            st.error("⛔ Vui lòng nhập Tiêu đề bài viết!")
            return
        if not file_uploads:
            st.error("⛔ Đồng chí chọn 'Viết mới' nhưng chưa tải file nào lên!")
            return

        links = []
        for f in file_uploads:
            ts        = datetime.now().strftime("%Y%m%d%H%M%S")
            raw_name  = f"{ts}_{f.name}"
            safe_name = lam_sach_ten_file(raw_name)
            try:
                supabase.storage.from_("tin_bai").upload(safe_name, f.read())
                url = supabase.storage.from_("tin_bai").get_public_url(safe_name)
                links.append(url)
            except Exception as e:
                st.error(f"Lỗi tải file {f.name}: {e}")

        if links:
            supabase.table("dang_ky_tin_bai").insert({
                "nguoi_gui": nguoi_gui_clean,
                "tieu_de":   tieu_de_viet_moi.strip(),
                "nguon_tin": nguon_label,
                "duong_dan": "\n".join(links),
                "ghi_chu":   ghi_chu_clean,
            }).execute()
            st.success(f"🎉 Đã gửi thành công 1 bài viết mới (kèm {len(file_uploads)} file đính kèm)!")

    else:  # Sưu tầm
        if not tieu_de_suu_tam.strip():
            st.error("⛔ Vui lòng nhập tiêu đề bài sưu tầm!")
            return
        if not link_suu_tam.strip():
            st.error("⛔ Vui lòng nhập đường dẫn bài sưu tầm!")
            return

        try:
            supabase.table("dang_ky_tin_bai").insert({
                "nguoi_gui": nguoi_gui_clean,
                "tieu_de":   tieu_de_suu_tam.strip(),
                "nguon_tin": nguon_label,
                "duong_dan": link_suu_tam.strip(),
                "ghi_chu":   ghi_chu_clean,
            }).execute()
            st.success("🎉 Đã gửi đăng ký bài sưu tầm thành công!")
        except Exception as e:
            st.error(f"Lỗi khi gửi bài sưu tầm: {e}")

# ─────────────────────────────────────────────
if __name__ == "__main__":
    main()
