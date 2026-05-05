import streamlit as st
import pandas as pd
from supabase import create_client
from datetime import datetime

# 1. KẾT NỐI (Lấy an toàn từ Secrets)
URL = st.secrets["SUPABASE_URL"]
KEY = st.secrets["SUPABASE_KEY"]
supabase = create_client(URL, KEY)

def main():
    st.set_page_config(page_title="Hệ thống Quản lý Tin bài", layout="wide")
    
    # Giao diện Header
    st.markdown('<div style="background-color:#004B87;padding:15px;border-radius:10px;color:white;text-align:center;"><h1>📝 HỆ THỐNG QUẢN LÝ TIN BÀI 4.0</h1></div>', unsafe_allow_html=True)

    # TẠO TAB: Một bên cho anh em đăng ký, một bên cho sếp Tuấn tổng hợp
    tab1, tab2 = st.tabs(["✍️ Đăng ký tin bài", "📊 Bảng tổng hợp trình duyệt"])

    # --- TAB 1: DÀNH CHO ANH EM TRONG TỔ THƯ KÝ ---
    with tab1:
        st.subheader("Gửi thông tin bài viết hàng ngày")
        with st.form("form_dang_ky", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                # Danh sách anh em theo thông báo phân công
                nguoi_gui = st.selectbox("Người gửi:", ["Lê Minh Tiến", "Trần Thị Thảo", "Nguyễn Thị Phương Thảo", "Phạm Thị Nga", "Lãnh đạo Phòng chuyên môn"])
                tieu_de = st.text_input("Tiêu đề bài viết:", placeholder="Ví dụ: Tin hoạt động lãnh đạo Ban...")
            with col2:
                nguon = st.radio("Nguồn tin:", ["Viết mới", "Đề nghị đăng lại"], horizontal=True)
                link = st.text_input("Link bài viết/file nội dung:")
            
            ghi_chu = st.text_area("Ghi chú thêm:")
            btn_gui = st.form_submit_button("Gửi bài cho sếp Tuấn")
            
            if btn_gui:
                data = {"nguoi_gui": nguoi_gui, "tieu_de": tieu_de, "nguon_tin": nguon, "duong_dan": link, "ghi_chu": ghi_chu}
                supabase.table("dang_ky_tin_bai").insert(data).execute()
                st.success("Đã gửi thành công! Sếp Tuấn sẽ tổng hợp lúc 15h00.")

    # --- TAB 2: DÀNH RIÊNG CHO SẾP TUẤN TỔNG HỢP ---
    with tab2:
        st.subheader(f"Danh sách tin bài ngày {datetime.now().strftime('%d/%m/%Y')}")
        
        # Lấy dữ liệu bài trong ngày
        today = datetime.now().date().isoformat()
        res = supabase.table("dang_ky_tin_bai").select("*").eq("ngay_dang_ky", today).execute()
        
        if res.data:
            df = pd.DataFrame(res.data)
            
            # Giao diện đánh dấu FB/Zalo OA như sếp Lợi yêu cầu
            st.write("Sếp Tuấn đánh dấu bài đăng MXH tại đây:")
            edited_df = st.data_editor(
                df,
                column_config={
                    "dang_facebook": st.column_config.CheckboxColumn("Đăng FB", default=False),
                    "dang_zalo": st.column_config.CheckboxColumn("Đăng Zalo OA", default=False),
                    "duong_dan": st.column_config.LinkColumn("Link")
                },
                disabled=["nguoi_gui", "tieu_de", "nguon_tin"],
                hide_index=True
            )
            
            if st.button("Lưu đánh dấu & Chốt danh sách"):
                # Cập nhật từng dòng sếp đã tích chọn vào Database
                for _, row in edited_df.iterrows():
                    supabase.table("dang_ky_tin_bai").update({
                        "dang_facebook": row["dang_facebook"],
                        "dang_zalo": row["dang_zalo"]
                    }).eq("id", row["id"]).execute()
                st.success("Đã chốt danh sách đăng bài ngày hôm nay!")
                
                # Nút xuất file trình lãnh đạo duyệt
                st.download_button("📥 Tải danh sách trình duyệt (Excel)", edited_df.to_csv(index=False).encode('utf-8-sig'), f"trinh_duyet_{today}.csv", "text/csv")
        else:
            st.info("Chưa có tin bài nào được gửi lên trong hôm nay.")

if __name__ == "__main__":
    main()
