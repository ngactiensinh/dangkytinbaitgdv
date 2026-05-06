import streamlit as st
import pandas as pd
from supabase import create_client
from datetime import datetime
import io
from docx import Document
from docx.shared import Cm
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
import plotly.express as px
import unicodedata
import re

# 1. KẾT NỐI
URL = st.secrets["SUPABASE_URL"]
KEY = st.secrets["SUPABASE_KEY"]
ADMIN_PASS = st.secrets.get("ADMIN_PASS", "141983")
supabase = create_client(URL, KEY)

# Hàm tạo file Word báo cáo (CHUẨN 6 CỘT THEO ẢNH, ẨN LINK)
def tao_file_word(df, ngay_thang):
    doc = Document()
    
    # Cài đặt khổ giấy ngang
    section = doc.sections[0]
    new_width, new_height = section.page_height, section.page_width
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = new_width
    section.page_height = new_height
    section.left_margin, section.right_margin = Cm(1.5), Cm(1.5)
    section.top_margin, section.bottom_margin = Cm(1.5), Cm(1.5)

    # Tiêu đề
    heading = doc.add_heading('BẢNG TỔNG HỢP ĐĂNG KÝ TIN BÀI', 1)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_ngay = doc.add_paragraph(f'Ngày tổng hợp: {ngay_thang}\n')
    p_ngay.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Tạo bảng 6 cột
    table = doc.add_table(rows=1, cols=6)
    table.style = 'Table Grid'
    table.autofit = False
    table.allow_autofit = False
    
    # Tỷ lệ 6 cột cho khổ A4 ngang (~26.5 cm)
    widths = (Cm(1.2), Cm(8.5), Cm(3.5), Cm(4.5), Cm(4.5), Cm(4.3))
    
    for i, col in enumerate(table.columns):
        col.width = widths[i]
        
    hdr = table.rows[0].cells
    headers = ['STT', 'Tiêu đề', 'Người gửi', 'Nguồn / File đính kèm', 'Đề xuất MXH', 'Người đăng']
    
    for i, h in enumerate(headers):
        hdr[i].text = h
        hdr[i].paragraphs[0].runs[0].bold = True

    # Đổ dữ liệu
    for idx, row in enumerate(df.itertuples(), 1):
        row_cells = table.add_row().cells
        row_cells[0].text = str(idx)
        row_cells[1].text = str(row.tieu_de)
        row_cells[2].text = str(row.nguoi_gui)
        
        # Ẩn link, chỉ hiện nguồn tin theo đúng yêu cầu
        row_cells[3].text = str(row.nguon_tin)
        
        # Xử lý MXH
        mxh = ["Đăng Web"] # Mặc định luôn có Đăng Web
        if row.dang_facebook: mxh.append("Facebook")
        if row.dang_zalo: mxh.append("Zalo OA")
        row_cells[4].text = ", ".join(mxh)
        
        # Cột Người đăng mới
        row_cells[5].text = str(row.nguoi_dang) if pd.notna(row.nguoi_dang) else ""
        
        # Ép lại chiều rộng
        for i, cell in enumerate(row_cells):
            cell.width = widths[i]

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def lam_sach_ten_file(ten_file):
    # 1. Bỏ dấu tiếng Việt
    ten_file = unicodedata.normalize('NFKD', ten_file).encode('ASCII', 'ignore').decode('utf-8')
    # 2. Thay khoảng trắng bằng dấu gạch dưới
    ten_file = ten_file.replace(' ', '_')
    # 3. Xóa sạch các ký tự đặc biệt (chỉ giữ lại chữ, số, dấu chấm, gạch dưới, gạch ngang)
    ten_file = re.sub(r'[^\w\.\-]', '', ten_file)
    return ten_file
    
def main():
    st.set_page_config(page_title="Hệ thống Quản lý Tin bài", layout="wide")
    
    st.markdown("""
        <div style="background-color:#004B87; padding:20px; border-radius:10px; color:white; text-align:center;">
            <h2 style="margin:0; font-size: 1.8rem; text-transform: uppercase;">Hệ thống quản lý tin bài đăng trang Thông tin điện tử</h2>
            <h3 style="margin:5px 0 0 0; font-size: 1.3rem; font-weight: normal; color: #FFD700;">Ban Tuyên giáo và Dân vận Tỉnh ủy Tuyên Quang</h3>
        </div>
        <br>
    """, unsafe_allow_html=True)

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
            tieu_de_viet_moi = st.text_input("Tiêu đề bài viết (*Bắt buộc nếu chọn Viết mới):", placeholder="Ví dụ: Infographic cuộc thi Tìm hiểu về chuyển đổi số...")
            file_uploads = st.file_uploader("Tải lên các file đính kèm (Có thể chọn nhiều file cùng lúc):", type=["doc", "docx", "pdf", "png", "jpg"], accept_multiple_files=True)
            
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
                    nguoi_gui_clean = str(nguoi_gui).strip()
                    nguon_clean = str(nguon).strip()
                    ghi_chu_clean = str(ghi_chu).strip() if pd.notna(ghi_chu) else ""

                    if nguon == "Viết mới":
                        if not tieu_de_viet_moi.strip():
                            st.error("Đồng chí vui lòng nhập Tiêu đề bài viết!")
                        elif not file_uploads:
                            st.error("Đồng chí chọn 'Viết mới' nhưng chưa tải file nào lên!")
                        else:
                            danh_sach_link = []
                            for f in file_uploads:
                                file_name = f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{f.name}"
                                # --- ĐOẠN CODE ĐÃ LẮP MÀNG LỌC ---
                                # 1. Chạy tên file qua hàm làm sạch trước khi upload
                                file_name_sach = lam_sach_ten_file(file_name)
                                
                                try:
                                    # 2. Upload bằng cái tên đã làm sạch (file_name_sach)
                                    supabase.storage.from_("tin_bai").upload(file_name_sach, f.read())
                                    
                                    # 3. Lấy link cũng phải dùng tên đã làm sạch
                                    link_chinh = supabase.storage.from_("tin_bai").get_public_url(file_name_sach)
                                    
                                    danh_sach_link.append(link_chinh)
                                except Exception as e:
                                    st.error(f"Lỗi tải file {f.name}: {e}")
                                # ---------------------------------
                            
                            # Gộp tất cả link vào chung 1 bài viết
                            link_tong_hop = "\n".join(danh_sach_link)
                            data = {
                                "nguoi_gui": nguoi_gui_clean, 
                                "tieu_de": tieu_de_viet_moi.strip(), 
                                "nguon_tin": nguon_clean, 
                                "duong_dan": link_tong_hop, 
                                "ghi_chu": ghi_chu_clean
                            }
                            supabase.table("dang_ky_tin_bai").insert(data).execute()
                            st.success(f"🎉 Đã gửi thành công 1 bài viết mới (kèm {len(file_uploads)} file đính kèm)!")
                    
                    else: # Xử lý tin sưu tầm
                        count = 0
                        for index, row in edited_links.iterrows():
                            t_de = str(row["tieu_de"]).strip() if pd.notna(row["tieu_de"]) else ""
                            l_ink = str(row["link"]).strip() if pd.notna(row["link"]) else ""
                            
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
            st.subheader("Bảng tổng hợp trình duyệt")
            
            # 🌟 TÍNH NĂNG MỚI: BỘ LỌC CHỌN NGÀY
            ngay_xem = st.date_input("Chọn ngày muốn xem và xuất báo cáo:", datetime.now().date())
            ngay_xem_str = ngay_xem.strftime('%d/%m/%Y')
            
            # Lọc dữ liệu theo ngày bạn chọn thay vì chốt cứng hôm nay
            res_ngay = supabase.table("dang_ky_tin_bai").select("*").eq("ngay_dang_ky", ngay_xem.isoformat()).execute()
            
            if res_ngay.data:
                df_ngay = pd.DataFrame(res_ngay.data)
                
                if 'nguoi_dang' not in df_ngay.columns:
                    df_ngay['nguoi_dang'] = ""
                else:
                    df_ngay['nguoi_dang'] = df_ngay['nguoi_dang'].fillna("")

                edited_df = st.data_editor(
                    df_ngay,
                    column_config={
                        "dang_facebook": st.column_config.CheckboxColumn("Đăng FB", default=False),
                        "dang_zalo": st.column_config.CheckboxColumn("Đăng Zalo", default=False),
                        "duong_dan": st.column_config.TextColumn("Danh sách Link (Bôi đen copy nếu cần)"),
                        "nguoi_dang": st.column_config.TextColumn("Người đăng (Gõ tên để xuất file)")
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
                    word_data = tao_file_word(edited_df, ngay_xem_str)
                    st.download_button("📥 Xuất file Word trình duyệt", data=word_data, file_name=f"TinBai_{ngay_xem.strftime('%Y%m%d')}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            else:
                st.info(f"Không có tin bài nào được đăng ký trong ngày {ngay_xem_str}.")

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
            # Ép kiểu ngày tháng cho chuẩn để tránh lỗi
            df_all['ngay_dang_ky'] = pd.to_datetime(df_all['ngay_dang_ky'], errors='coerce')
            df_all = df_all.dropna(subset=['ngay_dang_ky']) # Bỏ qua các dòng bị lỗi ngày
            df_all['Năm'] = df_all['ngay_dang_ky'].dt.year

            # 🌟 BỘ LỌC LIÊN HOÀN 3 TẦNG 🌟
            col_nam, col_kieu_loc, col_chi_tiet = st.columns(3)

            with col_nam:
                # Tự động lấy danh sách các năm có trong DB, xếp từ mới nhất
                danh_sach_nam = sorted(df_all['Năm'].unique().tolist(), reverse=True)
                chon_nam = st.selectbox("📅 Chọn Năm:", danh_sach_nam)

            # Lấy trước dữ liệu của năm đã chọn
            df_nam = df_all[df_all['Năm'] == chon_nam].copy()

            with col_kieu_loc:
                kieu_loc = st.selectbox("🔍 Lọc theo:", ["Cả năm", "Theo Quý", "Theo Tháng", "Theo Tuần", "Theo Ngày"])

            with col_chi_tiet:
                df_loc = df_nam.copy() # Mặc định ban đầu là lấy cả năm
                
                if kieu_loc == "Theo Quý":
                    df_nam['Quý'] = df_nam['ngay_dang_ky'].dt.quarter
                    quy_list = sorted(df_nam['Quý'].unique().tolist())
                    if quy_list:
                        chon_quy = st.selectbox("📌 Chọn Quý:", [f"Quý {q}" for q in quy_list])
                        q_val = int(chon_quy.split(" ")[1])
                        df_loc = df_nam[df_nam['Quý'] == q_val]
                        
                elif kieu_loc == "Theo Tháng":
                    df_nam['Tháng'] = df_nam['ngay_dang_ky'].dt.month
                    thang_list = sorted(df_nam['Tháng'].unique().tolist())
                    if thang_list:
                        chon_thang = st.selectbox("📌 Chọn Tháng:", [f"Tháng {t}" for t in thang_list])
                        t_val = int(chon_thang.split(" ")[1])
                        df_loc = df_nam[df_nam['Tháng'] == t_val]
                        
                elif kieu_loc == "Theo Tuần":
                    df_nam['Tuần'] = df_nam['ngay_dang_ky'].dt.isocalendar().week
                    tuan_list = sorted(df_nam['Tuần'].unique().tolist())
                    if tuan_list:
                        chon_tuan = st.selectbox("📌 Chọn Tuần:", [f"Tuần thứ {t}" for t in tuan_list])
                        t_val = int(chon_tuan.split(" ")[2])
                        df_loc = df_nam[df_nam['Tuần'] == t_val]
                        
                elif kieu_loc == "Theo Ngày":
                    ngay_list = sorted(df_nam['ngay_dang_ky'].dt.date.unique().tolist(), reverse=True)
                    if ngay_list:
                        # Format hiển thị ngày theo kiểu VN cho đẹp
                        chon_ngay = st.selectbox("📌 Chọn Ngày:", ngay_list, format_func=lambda x: x.strftime('%d/%m/%Y'))
                        df_loc = df_nam[df_nam['ngay_dang_ky'].dt.date == chon_ngay]
                        
                else:
                    # Nếu chọn "Cả năm" thì hiện cái dòng này cho đỡ trống
                    st.info(f"Đang hiển thị toàn bộ dữ liệu năm {chon_nam}")

            st.markdown("---")

            # --- HIỂN THỊ SỐ LIỆU ĐÃ LỌC ---
            if not df_loc.empty:
                m1, m2, m3 = st.columns(3)
                m1.metric("Tổng số bài", len(df_loc))
                
                # Vẫn giữ nguyên công thức "Bắt từ khóa" xịn sò lúc nãy
                so_bai_tu_viet = len(df_loc[df_loc['nguon_tin'].astype(str).str.contains('Viết mới|tự viết', case=False, na=False)])
                so_bai_suu_tam = len(df_loc[df_loc['nguon_tin'].astype(str).str.contains('Sưu tầm|đăng lại', case=False, na=False)])
                
                m2.metric("Bài tự viết", so_bai_tu_viet)
                m3.metric("Bài sưu tầm", so_bai_suu_tam)

                # Vẽ biểu đồ kèm hiện số liệu nổi trên cột
                df_bieu_do = df_loc.groupby(['nguoi_gui', 'nguon_tin']).size().reset_index(name='Số lượng')
                fig = px.bar(df_bieu_do, x='nguoi_gui', y='Số lượng', color='nguon_tin', barmode='group', text='Số lượng')
                fig.update_traces(textposition='outside')
                fig.update_layout(title="Biểu đồ phân bổ Tin bài theo Cán bộ", xaxis_title="Người gửi", yaxis_title="Số lượng")
                
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.warning("Không có bài viết nào được đăng ký trong khoảng thời gian này!")
        else:
            st.info("Hệ thống chưa có dữ liệu tin bài.")

if __name__ == "__main__":
    main()
