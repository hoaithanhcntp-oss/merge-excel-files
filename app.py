#Ứng dụng giao diện trực quan (Streamlit)
#Ứng dụng cho phép kéo thả nhiều file, hiển thị bảng xem trước (preview) kèm số thứ tự dòng Excel để người dùng quan sát và chọn đúng số dòng tiêu đề và số dòng dữ liệu.

import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Gộp File Excel", layout="wide")

st.title("Hệ thống gộp nhiều file Excel")
st.write("Hỗ trợ cấu hình hàng tiêu đề (Header) và hàng bắt đầu lấy dữ liệu riêng cho từng file.")

# 1. Tải lên danh sách file
uploaded_files = st.file_uploader(
    "Chọn các file Excel cần gộp (.xlsx, .xls):", 
    type=["xlsx", "xls"], 
    accept_multiple_files=True
)

if uploaded_files:
    st.info(f"Đã tải lên {len(uploaded_files)} file.")
    
    # Đọc nhanh danh sách sheet từ file đầu tiên
    first_file = uploaded_files[0]
    excel_obj = pd.ExcelFile(first_file)
    sheet_names = excel_obj.sheet_names
    
    col_opt1, col_opt2 = st.columns(2)
    with col_opt1:
        selected_sheet = st.selectbox("Chọn Sheet cần lấy dữ liệu:", sheet_names, index=0)
    with col_opt2:
        add_source_col = st.checkbox("Thêm cột tên file nguồn vào kết quả", value=True)
    
    # 2. Lựa chọn chế độ cấu hình dòng
    st.subheader("Tùy chọn cấu hình dòng")
    config_mode = st.radio(
        "Chế độ cấu hình:",
        options=["Cấu hình riêng cho từng file", "Áp dụng chung cùng một vị trí dòng cho tất cả file"],
        index=0,
        horizontal=True
    )
    
    file_configs = {}

    if config_mode == "Áp dụng chung cùng một vị trí dòng cho tất cả file":
        # Xem trước file đầu tiên
        raw_preview = pd.read_excel(first_file, sheet_name=selected_sheet, header=None, nrows=15)
        raw_preview.index = [f"Dòng {i + 1}" for i in range(len(raw_preview))]
        st.write("Xem trước file mẫu đầu tiên:")
        st.dataframe(raw_preview, use_container_width=True)
        
        c1, c2 = st.columns(2)
        with c1:
            common_header = st.number_input("Hàng là Header (tiêu đề cột):", min_value=1, max_value=100, value=1, step=1)
        with c2:
            common_start = st.number_input("Hàng bắt đầu có dữ liệu:", min_value=common_header + 1, max_value=200, value=common_header + 1, step=1)
            
        for f in uploaded_files:
            file_configs[f.name] = {"header_row": common_header, "data_start_row": common_start}

    else:
        st.write("Thiết lập dòng tiêu đề và dòng bắt đầu dữ liệu cho từng file:")
        
        for idx, f in enumerate(uploaded_files):
            with st.expander(f"📄 File {idx + 1}: {f.name}", expanded=(idx == 0)):
                # Đọc 12 dòng đầu của chính file đó để người dùng nhìn trực quan
                try:
                    df_prev = pd.read_excel(f, sheet_name=selected_sheet, header=None, nrows=12)
                    df_prev.index = [f"Dòng {i + 1}" for i in range(len(df_prev))]
                    st.caption("12 dòng đầu của file này (theo số thứ tự dòng trên Excel):")
                    st.dataframe(df_prev, use_container_width=True)
                except Exception as ex:
                    st.warning(f"Không thể tải xem trước file này: {ex}")
                
                # Ô nhập riêng cho từng file
                col_h, col_d = st.columns(2)
                with col_h:
                    h_val = st.number_input(
                        f"Hàng Header (File: {f.name})",
                        min_value=1,
                        max_value=100,
                        value=1,
                        step=1,
                        key=f"header_{idx}"
                    )
                with col_d:
                    d_val = st.number_input(
                        f"Hàng bắt đầu lấy dữ liệu (File: {f.name})",
                        min_value=h_val + 1,
                        max_value=200,
                        value=h_val + 1,
                        step=1,
                        key=f"start_{idx}"
                    )
                
                file_configs[f.name] = {"header_row": h_val, "data_start_row": d_val}

    drop_empty = st.checkbox("Loại bỏ các dòng dữ liệu hoàn toàn trống", value=True)

    # 3. Tiến hành gộp
    if st.button("Tiến hành gộp dữ liệu", type="primary"):
        merged_dfs = []
        errors = []
        progress_bar = st.progress(0)
        
        for idx, file in enumerate(uploaded_files):
            try:
                cfg = file_configs[file.name]
                h_row = cfg["header_row"]
                d_row = cfg["data_start_row"]
                
                # Đọc toàn bộ file thô
                df_raw = pd.read_excel(file, sheet_name=selected_sheet, header=None)
                
                if len(df_raw) < d_row - 1:
                    errors.append(f"File '{file.name}' có số dòng ít hơn hàng bắt đầu dữ liệu ({d_row}).")
                    continue
                
                # Lấy tên cột từ dòng header đã chọn
                header_series = df_raw.iloc[h_row - 1]
                cols = []
                for c_idx, val in enumerate(header_series):
                    if pd.isna(val) or str(val).strip() == "":
                        cols.append(f"Cot_{c_idx + 1}")
                    else:
                        cols.append(str(val).strip())
                
                # Cắt phần dữ liệu từ d_row trở đi
                df_data = df_raw.iloc[d_row - 1:].copy()
                df_data.columns = cols
                
                if add_source_col:
                    df_data.insert(0, "Tên_File_Nguồn", file.name)
                
                if drop_empty:
                    df_data = df_data.dropna(how="all", subset=cols)
                    
                merged_dfs.append(df_data)
            except Exception as e:
                errors.append(f"Lỗi tại file '{file.name}': {str(e)}")
                
            progress_bar.progress((idx + 1) / len(uploaded_files))
            
        if errors:
            for err in errors:
                st.warning(err)
                
        if merged_dfs:
            final_df = pd.concat(merged_dfs, ignore_index=True)
            st.success(f"Đã gộp thành công {len(merged_dfs)}/{len(uploaded_files)} file. Tổng số dòng dữ liệu: {len(final_df)}.")
            
            st.subheader("Xem trước kết quả sau khi gộp:")
            st.dataframe(final_df.head(50), use_container_width=True)
            
            # Xuất file kết quả
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                final_df.to_excel(writer, index=False, sheet_name="Merged_Data")
            output.seek(0)
            
            st.download_button(
                label="📥 Tải file Excel đã gộp (.xlsx)",
                data=output,
                file_name="File_Gop_Ket_Qua.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
