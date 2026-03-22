import streamlit as st
import pandas as pd
import io

# ==========================================
# 1. 頁面基本設定與標題
# ==========================================
st.set_page_config(page_title="Halloween Item Tracking Grid Generator", layout="wide")
st.title("🎃 萬聖節專案 Tracking Grid 自動生成工具")
st.markdown("請上傳 Program Sheet，系統將自動為您萃取資料並生成包含各子分頁的 Tracking Grid。")

# ==========================================
# 2. 建立檔案上傳區塊 (簡化為單一上傳)
# ==========================================
file_program = st.file_uploader("上傳 Program Sheet (Excel 格式)", type=["xlsx", "xls"])

st.divider() 

# ==========================================
# 3. 核心處理邏輯
# ==========================================
if st.button("生成 Tracking Grid", type="primary"):
    
    # 檢查是否已上傳檔案
    if file_program: 
        with st.spinner("資料處理中，請稍候..."):
            try:
                # --- 步驟 A: 讀取上傳的 Program Sheet ---
                df_raw = pd.read_excel(file_program)
                
                # --- 步驟 B: 建立主表 (26C5 HWLN ITEMS) ---
                # 預期最終版面需要的欄位
                final_main_cols = [
                    'DPCI', 'ITEM_DESC', 'FRP Level', 'Tollgate Exempt', 
                    'TPR Lite/Exempt', 'Factory Name', 'Factory ID', 
                    'Tollgate Date', 'TPR Date', 'Result'
                ]
                
                df_main = pd.DataFrame()
                
                # 嘗試對應來源欄位，若來源沒有該欄位，則自動建立空白欄位以「維持介面」
                for col in final_main_cols:
                    if col in df_raw.columns:
                        df_main[col] = df_raw[col]
                    # 容錯處理：若是常見的其他命名方式，自動 Mapping
                    elif col == 'ITEM_DESC' and 'Product Description' in df_raw.columns:
                        df_main[col] = df_raw['Product Description']
                    elif col == 'Factory ID' and 'Import Vendor ID' in df_raw.columns:
                        df_main[col] = df_raw['Import Vendor ID']
                    elif col == 'Factory Name' and 'Import Vendor Name' in df_raw.columns:
                        df_main[col] = df_raw['Import Vendor Name']
                    else:
                        df_main[col] = "" 
                        
                df_hwln_items = df_main
                
                # --- 步驟 C: 建立工廠清單 (Factory list) 維持相同介面 ---
                # 從主表自動萃取本次有參與的不重複工廠
                if 'Factory ID' in df_hwln_items.columns:
                    df_project_factories = df_hwln_items[['Factory ID', 'Factory Name']].drop_duplicates().dropna(subset=['Factory ID'])
                    df_project_factories.rename(columns={'Factory ID': 'Facility ID'}, inplace=True)
                else:
                    df_project_factories = pd.DataFrame(columns=['Facility ID', 'Factory Name'])
                    
                # 補上原本 Factory List 該有的追蹤欄位 (留空供後續手動填寫)
                factory_template_cols = ['Year', 'Total Skus', 'Costume Skus', 'FA Audit Date', 'FA Score & Grade', 'FA Expired Date', 'Remark']
                for col in factory_template_cols:
                    df_project_factories[col] = ""
                
                # 重新排列工廠清單的欄位順序
                cols_order = ['Year', 'Factory Name', 'Facility ID', 'Total Skus', 'Costume Skus', 'FA Audit Date', 'FA Score & Grade', 'FA Expired Date', 'Remark']
                df_project_factories = df_project_factories[[c for c in cols_order if c in df_project_factories.columns]]

                # --- 步驟 D: 建立豁免清單 (Exemption request) 維持相同介面 ---
                df_exemption = df_hwln_items.copy()
                df_exemption['Justification'] = "" 
                df_exemption['Item Status (New/CF)'] = "CF"
                
                # --- 步驟 E: 匯出為多活頁簿的 Excel 檔案 ---
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_hwln_items.to_excel(writer, index=False, sheet_name='26C5 HWLN ITEMS')
                    df_exemption.to_excel(writer, index=False, sheet_name='Exemption request')
                    df_project_factories.to_excel(writer, index=False, sheet_name='Factory list')
                    
                    # 主表表頭簡單格式化
                    workbook = writer.book
                    worksheet = writer.sheets['26C5 HWLN ITEMS']
                    header_format = workbook.add_format({'bold': True, 'bg_color': '#FFD966'}) 
                    
                    for col_num, value in enumerate(df_hwln_items.columns.values):
                        worksheet.write(0, col_num, value, header_format)

                processed_data = output.getvalue()
                
                st.success("✅ Tracking Grid 處理完成！已根據 Program Sheet 產出所有分頁。")
                
                # 提供下載
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
