import streamlit as st
import pandas as pd
from supabase import create_client
from datetime import datetime
import io
from docx import Document
import plotly.express as px

# 1. KẾT NỐI
URL = st.secrets["SUPABASE_URL"]
KEY = st.secrets["SUPABASE_KEY"]
ADMIN_PASS = st.secrets.get("ADMIN_PASS", "tuyenquang2026")
supabase = create_client(URL, KEY)

# Hàm tạo file Word báo cáo
def tao_file_word(df, ngay_thang):
    doc = Document()
    doc.add_heading('BẢNG TỔNG HỢP ĐĂNG KÝ TIN BÀI', 1)
    doc.add_paragraph(f'Ngày tổng hợp: {ngay_thang}\n')
    
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
        link_hien_thi = str(row.duong_dan) if row.duong_dan else "Không có"
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

    # Sidebar quản trị
    with st.sidebar:
        st.header("🔐 Dành cho Quản trị")
        mat_khau_nhap = st.text_input("Nhập mật khẩu để thao tác:", type="password")
        is_admin = (mat_khau_nhap == ADMIN_PASS)

    # TẠO 3 TAB
    tab1, tab2, tab3 = st.tabs(["✍️ Đăng ký tin bài", "📊 Tổng hợp hàng ngày", "📈 Thống kê & Biểu đồ"])

    # --- TAB 1: ĐĂNG KÝ (DÀNH CHO MỌI NGƯỜI) ---
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
            st.caption("Gõ trực tiếp vào bảng dưới. Bấm dấu cộng (+) ở góc dưới bảng để thêm dòng mới nếu gửi nhiều bài.")
            df_links = pd.DataFrame([{"tieu_de": "", "link": ""}])
            edited_links = st.data_editor(df_links, num_rows="dynamic", column_config={"tieu_de": "Tiêu đề bài sưu tầm", "link": "Đường dẫn (Link bài gốc)"}, use_container_width=True)
            
            st.markdown("---")
            ghi_chu = st.text_area("Ghi chú chung (Áp dụng cho tất cả các bài gửi trong lần này):")
            btn_gui = st.form_submit_button("Gửi đăng ký tin bài")
            
            if btn_gui:
                if not nguoi_gui.strip():
                    st.error("Đồng chí vui lòng điền Họ và tên trước khi gửi!")
                else:
                    count = 0
                    if nguon == "Viết mới":
                        if not file_uploads:
                            st.error("Đồng chí chọn 'Viết mới' nhưng chưa tải file nào lên!")
                        else:
                            for f in file_uploads:
                                file_name = f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{f.name}"
                                link_chinh = ""
                                try:
                                    supabase.storage.from_("tin_bai").upload(file_name, f.read())
                                    link_chinh = supabase.storage.from_("tin_bai").get_public_url(file_name)
                                except Exception:
                                    pass
                                
                                # Lấy tên file làm tiêu đề (cắt bỏ đuôi mở rộng)
                                tieu_de_file = f.name.rsplit('.', 1)[0] 
                                data = {"nguoi_gui": nguoi_gui, "tieu_de": tieu_de_file, "nguon_tin": nguon, "duong_dan": link_chinh, "ghi_chu": ghi_chu}
                                supabase.table("dang_ky_tin_bai").insert(data).execute()
                                count += 1
                            st.success(f"🎉 Đã gửi thành công {count} bài viết mới! Bộ phận tổng hợp sẽ xử lý lúc 15h00.")
                    
                    else: # Tin sưu tầm
                        valid_rows = [row for index, row in edited_links.iterrows() if str(row["tieu_de"]).strip() != ""]
                        if not valid_rows:
                            st.error("Đồng chí chọn 'Đăng lại' nhưng chưa điền Tiêu đề nào vào bảng!")
                        else:
                            for row in valid_rows:
                                data = {"nguoi_gui": nguoi_gui, "tieu_de": row["tieu_de"], "nguon_tin": nguon, "duong_dan": row["link"], "ghi_chu": ghi_chu}
                                supabase.table("dang_ky_tin_bai").insert(data).execute()
                                count += 1
                            st.success(f"🎉 Đã gửi thành công {count} bài sưu tầm! Bộ phận tổng hợp sẽ xử lý lúc 15h00.")

    # --- TAB 2: TỔNG HỢP TRÌNH DUYỆT (Chỉ Admin thấy nội dung) ---
    with tab2:
        if is_admin:
            today_str = datetime.now().strftime('%d/%m/%Y')
            st.subheader(f"Bảng tổng hợp trình duyệt (Ngày {today_str})")
            
            res_today = supabase.table("dang_ky_tin_bai").select("*").eq("ngay_dang_ky", datetime.now().date().isoformat()).execute()
            
            if res_today.data:
                df_today = pd.DataFrame(res_today.data)
                edited_df = st.data_editor(
                    df_today,
                    column_config={
                        "dang_facebook": st.column_config.CheckboxColumn("Đăng FB", default=False),
                        "dang_zalo": st.column_config.CheckboxColumn("Đăng Zalo", default=False),
                        "duong_dan": st.column_config.LinkColumn("Link/File")
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

            # Nút XÓA SẠCH DỮ LIỆU
            st.divider()
            st.warning("⚠️ Khu vực dọn dẹp dữ liệu")
            if st.button("🗑️ Xóa sạch toàn bộ tin bài (Reset dữ liệu test)"):
                try:
                    # Lệnh này sẽ xóa mọi dòng có id khác 0 (nghĩa là xóa tất cả)
                    supabase.table("dang_ky_tin_bai").delete().neq("id", 0).execute()
                    st.success("Đã dọn dẹp sạch sẽ toàn bộ dữ liệu test!")
                    st.rerun() # Refresh lại trang ngay lập tức
                except Exception as e:
                    st.error(f"Có lỗi khi xóa: {e}")
        else:
            st.warning("Vui lòng nhập mật khẩu Quản trị ở thanh bên trái để xem và thao tác phần này.")

    # --- TAB 3: THỐNG KÊ & BIỂU ĐỒ ---
    with tab3:
        st.subheader("Báo cáo số lượng tin bài")
        res_all = supabase.table("dang_ky_tin_bai").select("*").execute()
        
        if res_all.data:
            df_all = pd.DataFrame(res_all.data)
            df_all['ngay_dang_ky'] = pd.to_datetime(df_all['ngay_dang_ky'])
            df_all['Tháng'] = df_all['ngay_dang_ky'].dt.month
            df_all['Quý'] = df_all['ngay_dang_ky'].dt.quarter
            df_all['Năm'] = df_all['ngay_dang_ky'].dt.year

            f_col1, f_col2, f_col3 = st.columns(3)
            with f_col1:
                ds_nam = df_all['Năm'].unique().tolist()
                chon_nam = st.selectbox("Năm:", ds_nam, index=ds_nam.index(datetime.now().year) if datetime.now().year in ds_nam else 0)
            with f_col2:
                chon_quy = st.selectbox("Quý:", ["Tất cả", 1, 2, 3, 4])
            with f_col3:
                chon_thang = st.selectbox("Tháng:", ["Tất cả"] + list(range(1, 13)))

            df_loc = df_all[df_all['Năm'] == chon_nam]
            if chon_quy != "Tất cả": df_loc = df_loc[df_loc['Quý'] == chon_quy]
            if chon_thang != "Tất cả": df_loc = df_loc[df_loc['Tháng'] == chon_thang]

            if not df_loc.empty:
                st.write("---")
                m1, m2, m3 = st.columns(3)
                m1.metric("Tổng số bài", len(df_loc))
                m2.metric("Bài tự viết mới", len(df_loc[df_loc['nguon_tin'] == 'Viết mới']))
                m3.metric("Bài đề nghị đăng lại", len(df_loc[df_loc['nguon_tin'] == 'Đề nghị đăng lại']))

                st.markdown("##### 🏆 Thống kê năng suất cán bộ")
                df_bieu_do = df_loc.groupby(['nguoi_gui', 'nguon_tin']).size().reset_index(name='Số lượng')
                fig = px.bar(df_bieu_do, x='nguoi_gui', y='Số lượng', color='nguon_tin', 
                             title="Theo dõi chỉ tiêu tin bài", barmode='group',
                             labels={'nguoi_gui': 'Cán bộ', 'Số lượng': 'Số bài'})
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Không có dữ liệu trong khoảng thời gian này.")
        else:
            st.info("Hệ thống chưa có dữ liệu để thống kê.")

if __name__ == "__main__":
    main()
