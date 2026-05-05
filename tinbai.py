import streamlit as st
import pandas as pd
from supabase import create_client
from datetime import datetime
import io
from docx import Document
from docx.shared import Cm
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH

# 1. KẾT NỐI
URL = st.secrets["SUPABASE_URL"]
KEY = st.secrets["SUPABASE_KEY"]
ADMIN_PASS = st.secrets.get("ADMIN_PASS", "tuyenquang2026")
supabase = create_client(URL, KEY)

# Hàm tạo file Word báo cáo (ĐÃ NÂNG CẤP GIAO DIỆN)
def tao_file_word(df, ngay_thang):
    doc = Document()
    
    # --- CÀI ĐẶT KHỔ GIẤY NGANG VÀ CĂN SÁT LỀ ---
    section = doc.sections[0]
    new_width, new_height = section.page_height, section.page_width
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = new_width
    section.page_height = new_height
    
    # Chỉnh lề cách mép giấy 1.5 cm cho rộng rãi
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)

    # --- TIÊU ĐỀ BÁO CÁO CĂN GIỮA ---
    heading = doc.add_heading('BẢNG TỔNG HỢP ĐĂNG KÝ TIN BÀI', 1)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_ngay = doc.add_paragraph(f'Ngày tổng hợp: {ngay_thang}\n')
    p_ngay.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # --- TẠO BẢNG VÀ CHIA KÍCH THƯỚC CỘT ---
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    
    # Set độ rộng chuẩn cho từng cột (Tổng khoảng 26cm khổ A4 ngang)
    widths = [Cm(1.5), Cm(8.0), Cm(4.5), Cm(9.0), Cm(3.5)]
    
    hdr = table.rows[0].cells
    headers = ['STT', 'Tiêu đề', 'Người gửi', 'Nguồn / File đính kèm', 'Đề xuất MXH']
    
    for i, h in enumerate(headers):
        hdr[i].text = h
        # In đậm tiêu đề bảng
        hdr[i].paragraphs[0].runs[0].bold = True

    # --- ĐỔ DỮ LIỆU VÀO BẢNG ---
    for idx, row in enumerate(df.itertuples(), 1):
        row_cells = table.add_row().cells
        row_cells[0].text = str(idx)
        row_cells[1].text = str(row.tieu_de)
        row_cells[2].text = str(row.nguoi_gui)
        
        link = str(row.duong_dan).strip()
        if link.lower() == 'nan' or link == "" or link.lower() == 'empty':
            link_hien_thi = "Không có"
        else:
            link_hien_thi = link
            
        row_cells[3].text = f"{row.nguon_tin}\nLink/File: {link_hien_thi}"
        
        mxh = []
        if row.dang_facebook: mxh.append("Facebook")
        if row.dang_zalo: mxh.append("Zalo OA")
        row_cells[4].text = ", ".join(mxh) if mxh else "Đăng Web"

    # Áp dụng lại độ rộng cho tất cả các dòng (Bắt buộc với python-docx)
    for row in table.rows:
        for idx, width in enumerate(widths):
            row.cells[idx].width = width

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def main():
    st.set_page_config(page_title="Hệ thống Quản lý Tin bài", layout="wide")
    st.markdown('<div style="background-color:#004B87;padding:15px;border-radius:10px;color:white;text-align:center;"><h1>📝 HỆ THỐNG QUẢN LÝ TIN BÀI 4.0</h1></div>', unsafe_allow_html=True)

    with st.sidebar:
        st.header("🔐 Dành cho Quản trị")
        mat_khau_nhap = st.text_input("Nhập mật khẩu để thao tác:", type="password")
        is_admin = (mat_khau_nhap == ADMIN_PASS)

    tab1, tab2, tab3 = st.tabs(["✍️ Đăng ký tin bài", "📊 Tổng hợp hàng ngày", "📈 Thống kê & Biểu đồ"])

    # --- TAB 1: ĐĂNG KÝ ---
    with tab1:
        st.subheader("Gửi thông tin bài viết")
        
        with st.form("form_dang_ky", clear_on_submit=True):
            nguoi_gui = st.text_input("Họ và tên người gửi (*Bắt buộc):", placeholder="Nhập họ và tên của đồng chí...")
            nguon = st.radio("Loại tin bài (Chọn 1 trong 2):", ["Viết mới", "Đề nghị đăng lại (Sưu tầm)"], horizontal=True)
            
            st.markdown("---")
            st.markdown("### 🔹 NẾU LÀ TIN VIẾT MỚI:")
            st.caption("Có thể chọn nhiều file cùng lúc. Tiêu đề bài sẽ tự động lấy theo tên file.")
            file_uploads = st.file_uploader("Tải lên các file (Word/PDF/Ảnh):", type=["doc", "docx", "pdf", "png", "jpg"], accept_multiple_files=True)
            
            st.markdown("### 🔹 NẾU LÀ TIN SƯU TẦM (Đăng lại):")
            st.caption("Gõ trực tiếp vào bảng dưới. Bấm dấu cộng (+) ở góc dưới bảng để thêm dòng mới.")
            df_links = pd.DataFrame([{"tieu_de": "", "link": ""}])
            edited_links = st.data_editor(df_links, num_rows="dynamic", column_config={"tieu_de": "Tiêu đề bài sưu tầm", "link": "Đường dẫn (Link bài gốc)"}, use_container_width=True)
            
            st.markdown("---")
            ghi_chu = st.text_area("Ghi chú chung:")
            btn_gui = st.form_submit_button("Gửi đăng ký tin bài")
            
            if btn_gui:
                if not nguoi_gui.strip():
                    st.error("Đồng chí vui lòng điền Họ và tên trước khi gửi!")
                else:
                    count = 0
                    nguoi_gui_clean = str(nguoi_gui).strip()
                    nguon_clean = str(nguon).strip()
                    ghi_chu_clean = str(ghi_chu).strip() if pd.notna(ghi_chu) else ""

                    if nguon == "Viết mới":
                        if not file_uploads:
                            st.error("Đồng chí chọn 'Viết mới' nhưng chưa tải file nào lên!")
                        else:
                            for f in file_uploads:
                                file_name = f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{f.name}"
                                link_chinh = ""
                                try:
                                    res_upload = supabase.storage.from_("tin_bai").upload(file_name, f.read())
                                    link_chinh = supabase.storage.from_("tin_bai").get_public_url(file_name)
                                except Exception as e:
                                    st.error(f"Lỗi khi tải file '{f.name}' lên hệ thống lưu trữ! Chi tiết lỗi: {e}")
                                
                                tieu_de_file = f.name.rsplit('.', 1)[0] 
                                data = {
                                    "nguoi_gui": nguoi_gui_clean, 
                                    "tieu_de": str(tieu_de_file).strip(), 
                                    "nguon_tin": nguon_clean, 
                                    "duong_dan": str(link_chinh).strip(), 
                                    "ghi_chu": ghi_chu_clean
                                }
                                supabase.table("dang_ky_tin_bai").insert(data).execute()
                                count += 1
                            st.success(f"🎉 Đã gửi thành công {count} bài viết mới!")
                    
                    else:
                        for index, row in edited_links.iterrows():
                            t_de = str(row["tieu_de"]) if pd.notna(row["tieu_de"]) else ""
                            l_ink = str(row["link"]) if pd.notna(row["link"]) else ""
                            
                            t_de = t_de.strip()
                            l_ink = l_ink.strip()
                            
                            if t_de and t_de.lower() != "nan":
                                if l_ink.lower() == "nan": l_ink = ""
                                data = {"nguoi_gui": nguoi_gui_clean, "tieu_de": t_de, "nguon_tin": nguon_clean, "duong_dan": l_ink, "ghi_chu": ghi_chu_clean}
                                supabase.table("dang_ky_tin_bai").insert(data).execute()
                                count += 1
                                
                        if count > 0:
                            st.success(f"🎉 Đã gửi thành công {count} bài sưu tầm!")
                        else:
                            st.error("Vui lòng điền ít nhất một tiêu đề bài sưu tầm hợp lệ!")

    # --- TAB 2: TỔNG HỢP TRÌNH DUYỆT ---
    with tab2:
        if is_admin:
            today_str = datetime.now().strftime('%d/%m/%Y')
            st.subheader(f"Bảng tổng hợp trình duyệt (Ngày {today_str})")
            
            res_today = supabase.table("dang_ky_tin_bai").select("*").eq("ngay_dang_ky", datetime.now().date().isoformat()).execute()
            
            if res_today.data:
                df_today = pd.DataFrame(res_today.data)
                
                df_today['duong_dan'] = df_today['duong_dan'].fillna("").astype(str)
                df_today.loc[df_today['duong_dan'].str.lower() == 'nan', 'duong_dan'] = ""

                edited_df = st.data_editor(
                    df_today,
                    column_config={
                        "dang_facebook": st.column_config.CheckboxColumn("Đăng FB", default=False),
                        "dang_zalo": st.column_config.CheckboxColumn("Đăng Zalo", default=False),
                        "duong_dan": st.column_config.LinkColumn("Link tải file/báo gốc", display_text="Bấm vào đây để xem/tải")
                    },
                    disabled=["nguoi_gui", "tieu_de", "nguon_tin"],
                    hide_index=True
                )
                
                c_luu, c_xuat = st.columns(2)
                with c_luu:
                    if st.button("Lưu trạng thái MXH"):
                        for _, row in edited_df.iterrows():
                            supabase.table("dang_ky_tin_bai").update({"dang_facebook": row["dang_facebook"], "dang_zalo": row["dang_zalo"]}).eq("id", row["id"]).execute()
                        st.success("Đã lưu đánh dấu Mạng xã hội!")
                with c_xuat:
                    word_data = tao_file_word(edited_df, today_str)
                    st.download_button("📥 Xuất file Word trình duyệt", data=word_data, file_name=f"TinBai_{datetime.now().strftime('%Y%m%d')}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            else:
                st.info("Hôm nay chưa có ai gửi bài.")

            st.divider()
            st.warning("⚠️ Khu vực dọn dẹp dữ liệu")
            if st.button("🗑️ Xóa sạch toàn bộ tin bài (Reset dữ liệu test)"):
                try:
                    supabase.table("dang_ky_tin_bai").delete().neq("id", 0).execute()
                    st.success("Đã dọn dẹp sạch sẽ toàn bộ dữ liệu!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Có lỗi khi xóa: {e}")
        else:
            st.warning("Vui lòng nhập mật khẩu Quản trị ở thanh bên trái để xem và thao tác phần này.")

    # --- TAB 3: THỐNG KÊ ---
    with tab3:
        st.subheader("Báo cáo số lượng tin bài")
        res_all = supabase.table("dang_ky_tin_bai").select("*").execute()
        
        if res_all.data:
            df_all = pd.DataFrame(res_all.data)
            df_all['ngay_dang_ky'] = pd.to_datetime(df_all['ngay_dang_ky'])
            df_all['Năm'] = df_all['ngay_dang_ky'].dt.year

            chon_nam = st.selectbox("Năm:", df_all['Năm'].unique().tolist())
            df_loc = df_all[df_all['Năm'] == chon_nam]

            if not df_loc.empty:
                m1, m2, m3 = st.columns(3)
                m1.metric("Tổng số bài", len(df_loc))
                m2.metric("Bài tự viết", len(df_loc[df_loc['nguon_tin'] == 'Viết mới']))
                m3.metric("Bài sưu tầm", len(df_loc[df_loc['nguon_tin'] == 'Đề nghị đăng lại']))

                df_bieu_do = df_loc.groupby(['nguoi_gui', 'nguon_tin']).size().reset_index(name='Số lượng')
                fig = px.bar(df_bieu_do, x='nguoi_gui', y='Số lượng', color='nguon_tin', barmode='group')
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Không có dữ liệu.")
        else:
            st.info("Chưa có dữ liệu.")

if __name__ == "__main__":
    main()
