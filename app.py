import streamlit as st
import pandas as pd
import io

# ==========================================
# 1. 頁面基本設定與標題
# ==========================================
st.set_page_config(page_title="Program Items Generator", layout="wide")
st.title("🎃 萬聖節專案 Program Items 自動生成工具")
st.markdown("請分別上傳「主資料表」與「CATEGORY對照表」，系統將自動合併並生成標準格式的 Program Items 表單。")

# ==========================================
# 2. 建立雙檔案上傳區塊
# ==========================================
col1, col2 = st.columns(2)
with col1:
    file_main = st.file_uploader("📁 1. 上傳【主資料表】(Excel 格式)", type=["xlsx", "xls"])
with col2:
    file_cat = st.file_uploader("📁 2. 上傳【CATEGORY對照表】(Excel/CSV 格式)", type=["xlsx", "xls", "csv"])

st.divider() 

# ==========================================
# 3. 核心處理邏輯
# ==========================================
# 必須兩個檔案都上傳後，才會顯示後續的工作表選擇與執行按鈕
if file_main and file_cat:
    
    # --- 讀取主資料表的工作表 ---
    xls_main = pd.ExcelFile(file_main)
    sheet_names_main = xls_main.sheet_names
    default_idx_main = 0
    for i, s in enumerate(sheet_names_main):
        if 'ITEM' in s.upper() or 'HWLN' in s.upper() or 'PROGRAM' in s.upper():
            default_idx_main = i
            break
            
    # --- 讀取 CATEGORY對照表 的工作表 (若為 CSV 則略過) ---
    if file_cat.name.endswith(('.xlsx', '.xls')):
        xls_cat = pd.ExcelFile(file_cat)
        sheet_names_cat = xls_cat.sheet_names
        default_idx_cat = 0
        for i, s in enumerate(sheet_names_cat):
            if 'DATA' in s.upper():
                default_idx_cat = i
                break
    else:
        sheet_names_cat = ["CSV 檔案 (預設)"]
        default_idx_cat = 0

    # 顯示下拉選單讓使用者確認抓取的 Sheet
    col_sel1, col_sel2 = st.columns(2)
    with col_sel1:
        selected_sheet_main = st.selectbox("📄 請選擇【主資料表】的工作表：", sheet_names_main, index=default_idx_main)
    with col_sel2:
        selected_sheet_cat = st.selectbox("📄 請選擇【CATEGORY對照表】的工作表：", sheet_names_cat, index=default_idx_cat)

    # 執行按鈕
    if st.button("生成 Program Items", type="primary"):
        with st.spinner("資料比對與處理中，請稍候..."):
            try:
                # ---------------------------------------------------------
                # 步驟 A: 讀取 CATEGORY 對照表並建立 VLOOKUP 字典
                # ---------------------------------------------------------
                if file_cat.name.endswith('.csv'):
                    df_cat_raw = pd.read_csv(file_cat, header=None)
                else:
                    df_cat_raw = pd.read_excel(xls_cat, sheet_name=selected_sheet_cat, header=None)
                    
                # 找 CATEGORY 表的表頭 (往下掃 20 列找 DPCI)
                cat_header_idx = -1
                for i in range(min(20, len(df_cat_raw))):
                    if any('DPCI' in str(v).strip().upper() for v in df_cat_raw.iloc[i].values):
                        cat_header_idx = i
                        break
                        
                if cat_header_idx != -1:
                    df_cat_raw.columns = df_cat_raw.iloc[cat_header_idx]
                    df_cat_raw = df_cat_raw.iloc[cat_header_idx + 1:].reset_index(drop=True)
                else:
                    st.error("❌ 錯誤：在 CATEGORY 對照表中找不到 'DPCI' 欄位，無法進行比對。")
                    st.stop()

                # 建立欄位名稱正規化函數 (去空白、去換行、轉大寫)
                def normalize_col(col_name):
                    return str(col_name).replace('\n', '').replace('\r', '').replace(' ', '').upper()

                cat_cols_map = {normalize_col(c): c for c in df_cat_raw.columns}
                
                # 找出 DPCI 與 CATEGORY 的實際欄位名稱
                dpci_col_cat = cat_cols_map.get("DPCI", cat_cols_map.get("DPCI#"))
                cat_col_target = None
                
                # 模糊搜尋包含 CATEGORY 字眼的欄位 (例如 "AK Costumes Category" 也會中)
                for norm_c, orig_c in cat_cols_map.items():
                    if 'CATEGORY' in norm_c:
                        cat_col_target = orig_c
                        break
                
                # 建立 DPCI -> CATEGORY 的映射字典
                category_mapping = {}
                if dpci_col_cat and cat_col_target:
                    # 清理 DPCI 格式 (去除 - 與空白) 以確保能 100% Match
                    clean_cat_dpci = df_cat_raw[dpci_col_cat].astype(str).str.replace("-", "").str.strip()
                    category_mapping = dict(zip(clean_cat_dpci, df_cat_raw[cat_col_target]))
                else:
                    st.warning("⚠️ 警告：在 CATEGORY 對照表中找不到包含 'CATEGORY' 字眼的欄位。")

                # ---------------------------------------------------------
                # 步驟 B: 讀取主資料表並抓取內容
                # ---------------------------------------------------------
                df_raw = pd.read_excel(xls_main, sheet_name=selected_sheet_main, header=None)
                
                header_idx = -1
                for i in range(min(20, len(df_raw))):
                    if any('DPCI' in str(val).strip().upper() for val in df_raw.iloc[i].values):
                        header_idx = i
                        break
                
                if header_idx != -1:
                    df_raw.columns = df_raw.iloc[header_idx]
                    df_raw = df_raw.iloc[header_idx + 1:].reset_index(drop=True)
                else:
                    st.error(f"❌ 錯誤：在主資料表『{selected_sheet_main}』中找不到 'DPCI' 欄位。")
                    st.stop()

                raw_columns_map = {normalize_col(c): c for c in df_raw.columns}

                # 指定的 22 個欄位
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
                
                # 將資料寫入主表
                for col in target_columns:
                    norm_target = normalize_col(col)
                    
                    if norm_target == "DPCI":
                        if "DPCI" in raw_columns_map:
                            df_main[col] = df_raw[raw_columns_map["DPCI"]]
                        elif "DPCI#" in raw_columns_map:
                            df_main[col] = df_raw[raw_columns_map["DPCI#"]]
                        else:
                            df_main[col] = ""
                    
                    # 【特殊邏輯】：CATEGORY 欄位改為跨表 VLOOKUP 抓取
                    elif norm_target == "CATEGORY":
                        if "DPCI" in df_main.columns and category_mapping:
                            clean_main_dpci = df_main["DPCI"].astype(str).str.replace("-", "").str.strip()
                            df_main[col] = clean_main_dpci.map(category_mapping).fillna("")
                        else:
                            df_main[col] = ""

                    # 其他欄位維持原邏輯在主表中抓取
                    elif norm_target in raw_columns_map:
                        df_main[col] = df_raw[raw_columns_map[norm_target]]
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

                # ---------------------------------------------------------
                # 步驟 C: 匯出 Excel (維持全表 Arial、標題黃底粗體)
                # ---------------------------------------------------------
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_main.to_excel(writer, index=False, sheet_name='Program Items')
                    
                    workbook = writer.book
                    worksheet = writer.sheets['Program Items']
                    
                    workbook.formats[0].set_font_name('Arial')
                    cell_format = workbook.add_format({'font_name': 'Arial'})
                    header_format = workbook.add_format({
                        'bold': True, 
                        'bg_color': '#FFD966', 
                        'border': 1,
                        'font_name': 'Arial'
                    }) 
                    
                    for col_num, value in enumerate(df_main.columns.values):
                        worksheet.write(0, col_num, value, header_format)
                        worksheet.set_column(col_num, col_num, 15, cell_format)

                processed_data = output.getvalue()
                
                st.success("✅ 處理完成！已成功合併並跨檔抓取 CATEGORY 資料。")
                
                st.download_button(
                    label="📥 下載 Program Items.xlsx",
                    data=processed_data,
                    file_name="Automated_Program_Items.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            except Exception as e:
                st.error(f"❌ 處理檔案時發生錯誤: {e}")

else:
    st.info("💡 提示：請先在上方上傳「主資料表」與「CATEGORY對照表」，系統才會顯示生成按鈕。")
