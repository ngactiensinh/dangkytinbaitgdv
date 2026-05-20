# (Giữ nguyên các đoạn Import, Cấu hình, CSS, Tiện ích ở trên...)
# ... (Phần này sếp vẫn giữ y nguyên như bản V6.1 mình gửi trước đó) ...

# ─────────────────────────────────────────────
# SỬA LẠI HÀM LƯU TRẠNG THÁI TRONG TAB 2
# ─────────────────────────────────────────────
# Sếp thay thế đoạn code xử lý nút "Lưu trạng thái Mạng xã hội" bằng đoạn này:

            with c1:
                if st.button("💾 Lưu trạng thái Mạng xã hội", use_container_width=True):
                    thanh_cong = 0
                    loi = 0
                    for _, row in edited_df.iterrows():
                        row_id = row.get("id")
                        if row_id:
                            try:
                                # Chỉ cập nhật những trường thực sự thay đổi
                                supabase.table("dang_ky_tin_bai").update({
                                    "dang_facebook": bool(row["dang_facebook"]),
                                    "dang_zalo":     bool(row["dang_zalo"]),
                                    "nguoi_dang":    str(row.get("nguoi_dang", "")),
                                }).eq("id", row_id).execute()
                                thanh_cong += 1
                            except Exception:
                                loi += 1
                    
                    if thanh_cong > 0:
                        st.success(f"✅ Đã lưu {thanh_cong} tin bài thành công!")
                    if loi > 0:
                        st.error(f"⚠️ Có {loi} bài gặp lỗi khi lưu.")
                    st.rerun()

# (Phần còn lại của code sếp giữ nguyên hoàn toàn)
