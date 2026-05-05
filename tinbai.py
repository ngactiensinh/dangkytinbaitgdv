import streamlit as st
import pandas as pd
from supabase import create_client
from datetime import datetime
import io
from docx import Document

# 1. KẾT NỐI
URL = st.secrets["SUPABASE_URL"]
KEY = st.secrets["SUPABASE_KEY"]
ADMIN_PASS = st.secrets.get("ADMIN_PASS", "123456") # Lấy mật khẩu từ Secrets
supabase = create_client(URL, KEY)

# Hàm tạo file Word báo cáo
def tao_file_word(df, ngay_thang):
    doc = Document()
    doc.add_heading('BẢNG TỔNG HỢP ĐĂNG KÝ TIN BÀI', 1)
    doc.add_paragraph(f'Ngày tổng hợp: {ngay_thang}\n')
    
    # Tạo bảng
    table = doc.add_table(rows=1, cols=4)
    table.style = 'Table Grid'
    hdr = table.rows[0].cells
    hdr[0].text = 'STT'
    hdr[1].text = 'Tiêu đề / Người gửi'
    hdr[2].text = 'Nguồn / File đính kèm'
    hdr[3].text = 'Đề xuất MXH'

    for idx, row in enumerate(df.itertuples(), 1):
        row_cells = table.add_row().cells
        row_cells[0].text = str(idx)
        row_cells[1].text = f"{row.tieu_de}\n(Gửi bởi: {row.nguoi_gui})"
        link_hien_thi = str(row.duong_dan) if row.duong_dan else "Không có file/link"
        row_cells[2].text = f"{row.nguon_tin}\nLink/File: {link_hien_thi}"
        
        mxh = []
        if row.dang_facebook: mxh.append("Facebook")
        if row.dang_zalo: mxh.append("Zalo OA")
        row_cells[3].text = ", ".join(mxh) if mxh else "Đăng Web"
    
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def main():
    st.set_page_config(page_title="Hệ thống Quản lý Tin bài", layout="wide")
    st.markdown('<div style="background-color:#004B87;padding:15px;border-radius:10px;color:white;text-align:center;"><h1>📝 HỆ THỐNG QUẢN LÝ TIN BÀI 4.0</h1></div>', unsafe_allow_html=True)

    # Sidebar để Quản trị viên đăng nhập
    with st.sidebar:
        st.header("🔐 Dành cho Quản trị")
        mat_khau_nhap = st.text_input("Nhập mật khẩu để mở bảng tổng hợp:", type="password")
        is_admin = (mat_khau_nhap == ADMIN_PASS)

    # --- KHU VỰC ĐĂNG KÝ (Ai cũng thấy) ---
    st.subheader("Gửi thông tin bài viết hàng ngày")
    with st.form("form_dang_ky", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            nguoi_gui = st.selectbox("Người gửi:", ["Lê Minh Tiến", "Trần Thị Thảo", "Nguyễn Thị Phương Thảo", "Phạm Thị Nga", "Lãnh đạo Phòng chuyên môn"])
            tieu_de = st.text_input("Tiêu đề bài viết:")
        with col2:
            nguon = st.radio("Nguồn tin:", ["Viết mới", "Đề nghị đăng lại"], horizontal=True)
            link_ngoai = st.text_input("Đường dẫn (Nếu lấy bài từ báo khác):")
        
        file_upload = st.file_uploader("Tải lên file bài viết (Word/PDF/Ảnh) nếu là tin tự viết:", type=["doc", "docx", "pdf", "png", "jpg"])
        ghi_chu = st.text_area("Ghi chú thêm:")
        btn_gui = st.form_submit_button("Gửi đăng ký") # Đã bỏ chữ Sếp Tuấn nha!
        
        if btn_gui:
            link_chinh = link_ngoai
            # Xử lý upload file lên Supabase Storage
            if file_upload is not None:
                file_name = f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{file_upload.name}"
                try:
                    # Upload file
                    supabase.storage.from_("tin_bai").upload(file_name, file_upload.read())
                    # Lấy link công khai
                    link_chinh = supabase.storage.from_("tin_bai").get_public_url(file_name)
                except Exception as e:
                    st.error(f"Lỗi tải file: {e}")
            
            # Lưu vào Database
            data = {"nguoi_gui": nguoi_gui, "tieu_de": tieu_de, "nguon_tin": nguon, "duong_dan": link_chinh, "ghi_chu": ghi_chu}
            supabase.table("dang_ky_tin_bai").insert(data).execute()
            st.success("Đã gửi bài thành công! Tổng hợp sẽ được xuất lúc 15h00.")

    st.divider()

    # --- KHU VỰC QUẢN TRỊ (Chỉ hiện khi nhập đúng pass) ---
    if is_admin:
        today_str = datetime.now().strftime('%d/%m/%Y')
        st.header(f"📊 Bảng tổng hợp trình duyệt (Ngày {today_str})")
        
        res = supabase.table("dang_ky_tin_bai").select("*").eq("ngay_dang_ky", datetime.now().date().isoformat()).execute()
        
        if res.data:
            df = pd.DataFrame(res.data)
            edited_df = st.data_editor(
                df,
                column_config={
                    "dang_facebook": st.column_config.CheckboxColumn("Đăng FB", default=False),
                    "dang_zalo": st.column_config.CheckboxColumn("Đăng Zalo", default=False),
                    "duong_dan": st.column_config.LinkColumn("Link/File")
                },
                disabled=["nguoi_gui", "tieu_de", "nguon_tin"],
                hide_index=True
            )
            
            col_a, col_b = st.columns(2)
            with col_a:
                if st.button("Lưu đánh dấu MXH"):
                    for _, row in edited_df.iterrows():
                        supabase.table("dang_ky_tin_bai").update({"dang_facebook": row["dang_facebook"], "dang_zalo": row["dang_zalo"]}).eq("id", row["id"]).execute()
                    st.success("Đã lưu trạng thái MXH!")
            
            with col_b:
                # Nút tải file Word
                word_file = tao_file_word(edited_df, today_str)
                st.download_button(
                    label="📥 Xuất file Word trình Ban Biên tập",
                    data=word_file,
                    file_name=f"TongHop_TinBai_{datetime.now().strftime('%Y%m%d')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        else:
            st.info("Chưa có tin bài nào được gửi lên trong hôm nay.")

if __name__ == "__main__":
    main()
