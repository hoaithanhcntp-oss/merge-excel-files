#Ứng dụng giao diện trực quan (Streamlit)
#Ứng dụng cho phép kéo thả nhiều file, hiển thị bảng xem trước (preview) kèm số thứ tự dòng Excel để người dùng quan sát và chọn đúng số dòng tiêu đề và số dòng dữ liệu.

import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Gộp File Excel", layout="wide")

st.title("Hệ thống gộp file Excel")
st.write("Tải lên nhiều file Excel có cùng cấu trúc để gộp thành một file duy nhất.")

# 1. Upload nhiều file
uploaded_files = st.file_uploader(
    "Chọn các file Excel (.xlsx, .xls)", 
    type=["xlsx", "xls"], 
    accept_multiple_files=True
)

if uploaded_files:
    st.info(f"Đã tải lên {len(uploaded_files)} file.")
    
    # Đọc nhanh file đầu tiên để xem cấu trúc và số sheet
    first_file = uploaded_files[0]
    excel_obj = pd.ExcelFile(first_file)
    sheet_names = excel_obj.sheet_names
    
    col_cfg1, col_cfg2 = st.columns(2)
    with col_cfg1:
        selected_sheet = st.selectbox("Chọn Sheet cần lấy dữ liệu:", sheet_names, index=0)
    with col_cfg2:
        add_source_col = st.checkbox("Thêm cột tên file nguồn vào kết quả", value=True)

    # Đọc bản xem trước thô của file đầu tiên (20 dòng đầu)
    raw_preview = pd.read_excel(first_file, sheet_name=selected_sheet, header=None, nrows=20)
    
    # Hiển thị số dòng tương ứng với số hàng trên Excel (1-based index)
    preview_display = raw_preview.copy()
    preview_display.index = [f"Dòng {i + 1}" for i in range(len(preview_display))]
    
    st.subheader("Xem trước file đầu tiên (Để xác định số dòng):")
    st.dataframe(preview_display, use_container_width=True)
    
    # 2. Tuỳ chọn dòng Header và dòng bắt đầu lấy dữ liệu
    col_input1, col_input2, col_input3 = st.columns(3)
    with col_input1:
        header_row = st.number_input(
            "Vị trí hàng là Header (tiêu đề cột):", 
            min_value=1, 
            max_value=100, 
            value=1, 
            step=1,
            help="Số thứ tự dòng chứa tên các cột (theo giao diện Excel, bắt đầu từ 1)"
        )
    with col_input2:
        data_start_row = st.number_input(
            "Vị trí hàng bắt đầu lấy dữ liệu:", 
            min_value=header_row + 1, 
            max_value=200, 
            value=header_row + 1, 
            step=1,
            help="Số thứ tự dòng chứa bản ghi đầu tiên (thường là ngay sau dòng header)"
        )
    with col_input3:
        drop_empty = st.checkbox("Loại bỏ các dòng hoàn toàn trống", value=True)

    # 3. Nút thực hiện gộp dữ liệu
    if st.button("Tiến hành gộp dữ liệu", type="primary"):
        merged_dfs = []
        errors = []
        
        progress_bar = st.progress(0)
        
        for idx, file in enumerate(uploaded_files):
            try:
                # Đọc file không dùng header để tự cắt theo chỉ số dòng
                df_raw = pd.read_excel(file, sheet_name=selected_sheet, header=None)
                
                # Kiểm tra số lượng dòng của file
                if len(df_raw) < data_start_row - 1:
                    errors.append(f"File '{file.name}' không đủ số dòng dữ liệu.")
                    continue
                
                # Lấy tiêu đề cột từ dòng header được chỉ định
                header_series = df_raw.iloc[header_row - 1]
                cols = []
                for col_idx, val in enumerate(header_series):
                    if pd.isna(val) or str(val).strip() == "":
                        cols.append(f"Cot_{col_idx + 1}")
                    else:
                        cols.append(str(val).strip())
                
                # Lấy phần dữ liệu từ data_start_row trở đi
                df_data = df_raw.iloc[data_start_row - 1:].copy()
                df_data.columns = cols
                
                # Tùy chọn thêm cột tên file
                if add_source_col:
                    df_data.insert(0, "Tên_File_Nguồn", file.name)
                
                if drop_empty:
                    # Bỏ các dòng rỗng
                    subset_check = cols
                    df_data = df_data.dropna(how="all", subset=subset_check)
                
                merged_dfs.append(df_data)
            except Exception as e:
                errors.append(f"Lỗi khi đọc file '{file.name}': {str(e)}")
            
            progress_bar.progress((idx + 1) / len(uploaded_files))
            
        if errors:
            for err in errors:
                st.warning(err)
                
        if merged_dfs:
            final_df = pd.concat(merged_dfs, ignore_index=True)
            st.success(f"Gộp thành công {len(merged_dfs)} file! Tổng cộng: {len(final_df)} dòng dữ liệu.")
            
            st.subheader("Kết quả dữ liệu sau khi gộp:")
            st.dataframe(final_df.head(100), use_container_width=True)
            
            # Xuất file Excel để tải xuống
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                final_df.to_excel(writer, index=False, sheet_name="Merged_Data")
            output.seek(0)
            
            st.download_button(
                label="📥 Tải file Excel đã gộp (.xlsx)",
                data=output,
                file_name="Ket_qua_gop.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
