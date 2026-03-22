import streamlit as st
import pandas as pd
import io

# ==========================================
# 1. 頁面基本設定與標題
# ==========================================
st.set_page_config(page_title="Program Items Generator", layout="wide")
st.title("🎃 萬聖節專案 Program Items 自動生成工具")
st.markdown("請上傳 Program Sheet，系統將自動萃取資料並生成標準格式的 Program Items 表單。")

# ==========================================
# 2. 建立檔案上傳區塊
# ==========================================
file_program = st.file_uploader("上傳 Program Sheet (Excel 格式)", type=["xlsx", "xls"])

st.divider() 

# ==========================================
# 3. 核心處理邏輯
# ==========================================
if st.button("生成 Program Items", type="primary"):
    
    if file_program: 
        with st.spinner("資料處理中，請稍候..."):
            try:
                # --- 步驟 A: 讀取檔案並尋找真實表頭 ---
                df_raw = pd.read_excel(file_program, header=None)
                
                header_idx = -1
                for i in range(min(20, len(df_raw))):
                    row_values = [str(val).strip().upper() for val in df_raw.iloc[i].values]
                    if 'DPCI' in row_values:
                        header_idx = i
                        break
                
                if header_idx != -1:
                    df_raw.columns = df_raw.iloc[header_idx]
                    df_raw = df_raw.iloc[header_idx + 1:].reset_index(drop=True)
                else:
                    st.warning("⚠️ 警告：在檔案前 20 列中找不到 'DPCI' 欄位，可能會導致抓取失敗。")

                # --- 步驟 B: 建立終極欄位比對字典 ---
                def normalize_col(col_name):
                    return str(col_name).replace('\n', '').replace('\r', '').replace(' ', '').upper()

                raw_columns_map = {normalize_col(c): c for c in df_raw.columns}

                # --- 步驟 C: 更新後的 22 個指定欄位 (已移除 Self PPT/PPT completed/Self BAH) ---
                target_columns = [
                    "DPCI", "CATEGORY", "ITEM_DESC", "PHOTO", "FRP Level", 
                    "Red Seal(Y/N)", "CF item( Y/N )", "Tollgate Exempt", 
                    "TPR Lite/Exempt", "Factory Name", "Factory ID", 
                    "Total SKU per Factory", "QTY", 
                    "Tollgate Date", "TPR Date", "Dupro Date", "Result", 
                    "TOP Result", "FRI plan", "Port of Export", 
                    "1st Ship window", "Inspection Office"
                ]
                
                df_main = pd.DataFrame()
                
                # --- 步驟 D: 抓取資料寫入主表 ---
                for col in target_columns:
                    norm_target = normalize_col(col)
                    
                    if norm_target in raw_columns_map:
                        original_col_name = raw_columns_map[norm_target]
                        df_main[col] = df_raw[original_col_name]
                        
                    elif norm_target == normalize_col("ITEM_DESC") and normalize_col("Product Description") in raw_columns_map:
                        df_main[col] = df_raw[raw_columns_map[normalize_col("Product Description")]]
                    elif norm_target == normalize_col("Factory ID") and normalize_col("Import Vendor ID") in raw_columns_map:
                        df_main[col] = df_raw[raw_columns_map[normalize_col("Import Vendor ID")]]
                    elif norm_target == normalize_col("Factory Name") and normalize_col("Import Vendor Name") in raw_columns_map:
                        df_main[col] = df_raw[raw_columns_map[normalize_col("Import Vendor Name")]]
                    
                    else:
                        df_main[col] = ""
                
                if "DPCI" in df_main.columns:
                    df_main = df_main.dropna(subset=['DPCI'], how='all')

                # --- 步驟 E: 匯出單一 Tab 的 Excel 並設定格式 ---
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_main.to_excel(writer, index=False, sheet_name='Program Items')
                    
                    workbook = writer.book
                    worksheet = writer.sheets['Program Items']
                    
                    # 1. 設定整份工作表的預設字體為 Arial (處理 pandas 預設寫入的資料格)
                    workbook.formats[0].set_font_name('Arial')
                    
                    # 2. 建立資料列的通用格式 (Arial)
                    cell_format = workbook.add_format({'font_name': 'Arial'})
                    
                    # 3. 建立表頭格式：Arial、粗體、黃底、加邊框
                    header_format = workbook.add_format({
                        'bold': True, 
                        'bg_color': '#FFD966', 
                        'border': 1,
                        'font_name': 'Arial'
                    }) 
                    
                    # 套用格式到表頭與資料直行
                    for col_num, value in enumerate(df_main.columns.values):
                        # 覆寫第一列 (表頭) 格式
                        worksheet.write(0, col_num, value, header_format)
                        # 將整行的寬度設為 15，並套用 Arial 字體格式
                        worksheet.set_column(col_num, col_num, 15, cell_format)

                processed_data = output.getvalue()
                
                st.success("✅ Program Items 處理完成！已移除指定欄位並套用 Arial 字體與粗體表頭。")
                
                st.download_button(
                    label="📥 下載 Program Items.xlsx",
                    data=processed_data,
                    file_name="Automated_Program_Items.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            except Exception as e:
                st.error(f"❌ 處理檔案時發生錯誤: {e}")

    else:
        st.warning("⚠️ 請先上傳 Program Sheet 檔案。")
