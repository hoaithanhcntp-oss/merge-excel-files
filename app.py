#Ứng dụng giao diện trực quan (Streamlit)
#Ứng dụng cho phép kéo thả nhiều file, hiển thị bảng xem trước (preview) kèm số thứ tự dòng Excel để người dùng quan sát và chọn đúng số dòng tiêu đề và số dòng dữ liệu.

import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Gộp File Excel", layout="wide")

st.title("Hệ thống gộp nhiều file Excel")

# ==========================================
# 1. KHU VỰC CẤU HÌNH CỐ ĐỊNH (LUÔN HIỂN THỊ TEXTBOX)
# ==========================================
st.sidebar.header("⚙️ Cấu hình mặc định")

# Textbox nhập vị trí hàng tiêu đề
default_header_str = st.sidebar.text_input(
    "1. Hàng tiêu đề (Header):",
    value="1",
    help="Nhập số thứ tự hàng chứa tên cột (theo số dòng trên Excel, ví dụ: 1, 3, 5...)"
)

# Textbox nhập vị trí hàng bắt đầu lấy dữ liệu
default_start_str = st.sidebar.text_input(
    "2. Hàng bắt đầu lấy dữ liệu:",
    value="2",
    help="Nhập số thứ tự hàng chứa dòng dữ liệu đầu tiên (ví dụ: 2, 4, 6...)"
)

# Tùy chọn bổ sung ở sidebar
sheet_name_input = st.sidebar.text_input("Tên Sheet (để trống = lấy sheet đầu tiên):", value="")
add_filename_col = st.sidebar.checkbox("Thêm cột tên file nguồn", value=True)
drop_blank_rows = st.sidebar.checkbox("Loại bỏ các dòng hoàn toàn trống", value=True)

# Chuyển đổi an toàn giá trị mặc định sang số nguyên
try:
    default_header_val = int(default_header_str.strip())
except ValueError:
    default_header_val = 1

try:
    default_start_val = int(default_start_str.strip())
except ValueError:
    default_start_val = 2


# ==========================================
# 2. KHU VỰC TẢI LÊN FILE
# ==========================================
st.subheader("📂 Bước 1: Tải các file Excel lên")

uploaded_files = st.file_uploader(
    "Kéo thả hoặc chọn các file Excel (.xlsx, .xls):", 
    type=["xlsx", "xls"], 
    accept_multiple_files=True
)

# ==========================================
# 3. DANH SÁCH TEXTBOX RIÊNG CHO MỖI FILE
# ==========================================
file_settings = {}

if uploaded_files:
    st.subheader("📝 Bước 2: Kiểm tra và chỉnh sửa hàng dữ liệu cho từng file")
    st.caption("Bạn có thể thay đổi trực tiếp giá trị trong ô Textbox của từng file nếu các file có vị trí dòng bắt đầu khác nhau.")

    # Hiển thị tiêu đề các cột nhập liệu
    col_head_name, col_head_header, col_head_start = st.columns([4, 3, 3])
    with col_head_name:
        st.markdown("**Tên file**")
    with col_head_header:
        st.markdown("**Hàng Header (Tiêu đề)**")
    with col_head_start:
        st.markdown("**Hàng bắt đầu lấy dữ liệu**")

    # Hiển thị từng file thành một hàng ngang có ô Textbox riêng
    for idx, f in enumerate(uploaded_files):
        c_name, c_hdr, c_start = st.columns([4, 3, 3])
        
        with c_name:
            st.write(f"📄 **{idx + 1}. {f.name}**")
        
        with c_hdr:
            h_input = st.text_input(
                label=f"Hàng Header của file {f.name}",
                value=str(default_header_val),
                key=f"txt_header_{idx}",
                label_visibility="collapsed"
            )
            
        with c_start:
            # ĐÂY LÀ TEXTBOX NHẬP HÀNG BẮT ĐẦU LẤY DỮ LIỆU CỦA TỪNG FILE
            s_input = st.text_input(
                label=f"Hàng bắt đầu dữ liệu của file {f.name}",
                value=str(default_start_val),
                key=f"txt_start_{idx}",
                label_visibility="collapsed"
            )
            
        # Parse giá trị từ Textbox
        try:
            h_num = int(h_input.strip())
        except ValueError:
            h_num = default_header_val

        try:
            s_num = int(s_input.strip())
        except ValueError:
            s_num = default_start_val
            
        file_settings[f.name] = {"header": h_num, "start": s_num}

    st.divider()

    # ==========================================
    # 4. TIẾN HÀNH GỘP FILE
    # ==========================================
    st.subheader("🚀 Bước 3: Thực hiện gộp file")
    
    if st.button("Bắt đầu gộp dữ liệu", type="primary"):
        merged_dfs = []
        errors = []
        
        target_sheet = sheet_name_input.strip() if sheet_name_input.strip() != "" else 0
        
        for file in uploaded_files:
            try:
                cfg = file_settings[file.name]
                header_row = cfg["header"]
                data_start_row = cfg["start"]
                
                # Đọc toàn bộ file thô không gán header trước
                df_raw = pd.read_excel(file, sheet_name=target_sheet, header=None)
                
                # Kiểm tra số dòng
                if len(df_raw) < data_start_row:
                    errors.append(f"⚠️ File '{file.name}': Tổng số dòng ({len(df_raw)}) nhỏ hơn hàng bắt đầu lấy dữ liệu ({data_start_row}).")
                    continue
                
                # 1. Trích xuất tiêu đề cột từ hàng header
                header_series = df_raw.iloc[header_row - 1]
                cols = []
                for c_idx, val in enumerate(header_series):
                    if pd.isna(val) or str(val).strip() == "":
                        cols.append(f"Cột_{c_idx + 1}")
                    else:
                        cols.append(str(val).strip())
                
                # 2. Cắt dữ liệu từ dòng data_start_row trở đi
                df_data = df_raw.iloc[data_start_row - 1:].copy()
                df_data.columns = cols
                
                # 3. Thêm cột tên file nguồn nếu được chọn
                if add_filename_col:
                    df_data.insert(0, "Tên_File_Gốc", file.name)
                
                # 4. Bỏ các dòng rỗng
                if drop_blank_rows:
                    df_data = df_data.dropna(how="all", subset=cols)
                    
                merged_dfs.append(df_data)
                
            except Exception as e:
                errors.append(f"❌ Lỗi khi đọc file '{file.name}': {str(e)}")
        
        if errors:
            for err in errors:
                st.warning(err)
                
        if merged_dfs:
            final_df = pd.concat(merged_dfs, ignore_index=True)
            st.success(f" Đã gộp thành công {len(merged_dfs)} file! Tổng cộng có **{len(final_df)}** dòng dữ liệu.")
            
            st.markdown("### Xem trước dữ liệu sau khi gộp:")
            st.dataframe(final_df.head(50), use_container_width=True)
            
            # Xuất file kết quả tải về
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                final_df.to_excel(writer, index=False, sheet_name="Du_Lieu_Tong_Hop")
            output.seek(0)
            
            st.download_button(
                label="📥 Tải file Excel tổng hợp (.xlsx)",
                data=output,
                file_name="Tong_Hop_Cac_File.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

else:
    st.info("Vui lòng tải lên ít nhất 1 file Excel để xuất hiện danh sách cấu hình chi tiết.")
