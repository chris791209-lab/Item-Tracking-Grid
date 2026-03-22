import streamlit as st
import pandas as pd
import io
import re

# ==========================================
# 1. 頁面基本設定與標題
# ==========================================
st.set_page_config(page_title="Halloween Item Tracking Grid Generator", layout="wide")
st.title("🎃 萬聖節專案 Tracking Grid 自動生成工具")
st.markdown("請上傳 Program Sheet，系統將自動為您萃取資料並生成包含各子分頁的 Tracking Grid。")

# ==========================================
# 2. 建立檔案上傳區塊
# ==========================================
file_program = st.file_uploader("上傳 Program Sheet (Excel 格式)", type=["xlsx", "xls"])

st.divider() 

# ==========================================
# 3. 核心處理邏輯
# ==========================================
if st.button("生成 Tracking Grid", type="primary"):
    
    if file_program: 
        with st.spinner("資料處理中，請稍候..."):
            try:
                # --- 步驟 A: 智慧讀取 Program Sheet (解決表頭偏移問題) ---
                # 先讀取前 10 行來尋找真正的表頭 (尋找包含 DPCI 的那一列)
                df_preview = pd.read_excel(file_program, nrows=10)
                header_idx = 0
                
                if 'DPCI' not in df_preview.columns:
                    for i in range(len(df_preview)):
                        # 將該列的值轉為字串清單進行比對
                        row_values = [str(val).strip() for val in df_preview.iloc[i].values]
                        if 'DPCI' in row_values:
                            header_idx = i + 1 # 找到真正的表頭列索引
                            break
                
                # 將檔案指標歸零，並使用正確的 header_idx 重新讀取完整資料
                file_program.seek(0)
                df_raw = pd.read_excel(file_program, header=header_idx)
                
                # --- 步驟 B: 清理欄位名稱 (解決 Alt+Enter 換行與多餘空白問題) ---
                # 將欄位名稱的換行符號 \n 換成空格，並去除前後空白
                df_raw.columns = df_raw.columns.astype(str).str.replace(r'\n', ' ', regex=True)
                df_raw.columns = df_raw.columns.str.replace(r'\s+', ' ', regex=True).str.strip()

                # --- 步驟 C: 建立主表 (26C5 HWLN ITEMS) ---
                # 包含您指定的重點欄位與追蹤欄位
                final_main_cols = [
                    'DPCI', 'CATEGORY', 'ITEM_DESC', 'PHOTO', 'FRP Level', 
                    'Factory Name', 'Factory ID', 'Total SKU per Factory', 'QTY',
                    'Tollgate Exempt', 'TPR Lite/Exempt', 'Tollgate Date', 'TPR Date', 'Result'
                ]
                
                df_main = pd.DataFrame()
                
                # 欄位精準 Mapping
                for col in final_main_cols:
                    if col in df_raw.columns:
                        df_main[col] = df_raw[col]
                    # 容錯：處理常見的不同命名方式
                    elif col == 'ITEM_DESC' and 'Product Description' in df_raw.columns:
                        df_main[col] = df_raw['Product Description']
                    elif col == 'Factory ID' and 'Import Vendor ID' in df_raw.columns:
                        df_main[col] = df_raw['Import Vendor ID']
                    elif col == 'Factory Name' and 'Import Vendor Name' in df_raw.columns:
                        df_main[col] = df_raw['Import Vendor Name']
                    else:
                        df_main[col] = "" # 若來源檔案真的沒有此欄位，則留空維持版面
                        
                df_hwln_items = df_main
                
                # --- 步驟 D: 建立工廠清單 (Factory list) ---
                if 'Factory ID' in df_hwln_items.columns:
                    df_project_factories = df_hwln_items[['Factory ID', 'Factory Name']].drop_duplicates().dropna(subset=['Factory ID'])
                    df_project_factories.rename(columns={'Factory ID': 'Facility ID'}, inplace=True)
                else:
                    df_project_factories = pd.DataFrame(columns=['Facility ID', 'Factory Name'])
                    
                factory_template_cols = ['Year', 'Total Skus', 'Costume Skus', 'FA Audit Date', 'FA Score & Grade', 'FA Expired Date', 'Remark']
                for col in factory_template_cols:
                    df_project_factories[col] = ""
                
                cols_order = ['Year', 'Factory Name', 'Facility ID', 'Total Skus', 'Costume Skus', 'FA Audit Date', 'FA Score & Grade', 'FA Expired Date', 'Remark']
                df_project_factories = df_project_factories[[c for c in cols_order if c in df_project_factories.columns]]

                # --- 步驟 E: 建立豁免清單 (Exemption request) ---
                df_exemption = df_hwln_items.copy()
                df_exemption['Justification'] = "" 
                df_exemption['Item Status (New/CF)'] = "CF"
                
                # --- 步驟 F: 匯出多活頁簿 Excel ---
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_hwln_items.to_excel(writer, index=False, sheet_name='26C5 HWLN ITEMS')
                    df_exemption.to_excel(writer, index=False, sheet_name='Exemption request')
                    df_project_factories.to_excel(writer, index=False, sheet_name='Factory list')
                    
                    workbook = writer.book
                    worksheet = writer.sheets['26C5 HWLN ITEMS']
                    header_format = workbook.add_format({'bold': True, 'bg_color': '#FFD966'}) 
                    
                    for col_num, value in enumerate(df_hwln_items.columns.values):
                        worksheet.write(0, col_num, value, header_format)

                processed_data = output.getvalue()
                
                st.success("✅ Tracking Grid 處理完成！資料已成功抓取。")
                
                st.download_button(
                    label="📥 下載 Halloween Item Tracking Grid.xlsx",
                    data=processed_data,
                    file_name="Automated_Halloween_Tracking_Grid.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            except Exception as e:
                st.error(f"❌ 處理檔案時發生錯誤: {e}")

    else:
        st.warning("⚠️ 請先上傳 Program Sheet 檔案。")
