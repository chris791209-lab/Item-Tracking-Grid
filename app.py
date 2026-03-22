import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Halloween Item Tracking Grid Generator", layout="wide")
st.title("🎃 萬聖節專案 Tracking Grid 自動生成工具")

st.markdown("請上傳系統匯出的原始資料與工廠清單，系統將自動為您生成 Tracking Grid。")

# 建立上傳區塊
col1, col2, col3 = st.columns(3)
with col1:
    file_data = st.file_uploader("上傳 Data (CSV/Excel)", type=["csv", "xlsx"])
with col2:
    file_data_cos = st.file_uploader("上傳 Data Cos (CSV/Excel)", type=["csv", "xlsx"])
with col3:
    file_factory = st.file_uploader("上傳 Factory List (CSV/Excel)", type=["csv", "xlsx"])if st.button("生成 Tracking Grid"):
    if file_data and file_factory: # 確保必要的檔案已上傳
        with st.spinner("資料處理中..."):
            
            # 1. 讀取檔案
            df_data = pd.read_csv(file_data) if file_data.name.endswith('.csv') else pd.read_excel(file_data)
            df_factory_raw = pd.read_csv(file_factory) if file_factory.name.endswith('.csv') else pd.read_excel(file_factory)
            
            # 2. 篩選 Data 所需欄位 (此處欄位名稱需對應您實際的 Raw Data 表頭)
            # 假設您的 Data 表中有 'DPCI', 'Item Description', 'Import Vendor ID' (對應工廠) 等
            core_columns = ['DPCI', 'Product Description', 'Import Vendor ID', 'Import Vendor Name']
            
            # 處理可能缺失的欄位，避免報錯
            available_cols = [col for col in core_columns if col in df_data.columns]
            df_main = df_data[available_cols].copy()
            
            # 重新命名欄位以符合 Tracking Grid 主表的格式
            df_main.rename(columns={
                'Product Description': 'ITEM_DESC',
                'Import Vendor ID': 'Factory ID',
                'Import Vendor Name': 'Factory Name'
            }, inplace=True)# --- 建立主表 (HWLN ITEMS) ---
            # 新增人工需要填寫的追蹤欄位 (給予空值)
            tracking_columns = ["FRP Level", "Tollgate Exempt", "TPR Lite/Exempt", "Tollgate Date", "TPR Date", "Result"]
            for col in tracking_columns:
                df_main[col] = "" 
                
            # 整理主表欄位順序 (依據您附件的結構)
            final_main_cols = ['DPCI', 'ITEM_DESC', 'FRP Level', 'Tollgate Exempt', 'TPR Lite/Exempt', 'Factory Name', 'Factory ID', 'Tollgate Date', 'TPR Date', 'Result']
            # 確保只保留存在的欄位
            df_hwln_items = df_main[[c for c in final_main_cols if c in df_main.columns]]


            # --- 建立工廠清單 (Factory List) ---
            # 透過對主表 Factory ID 去重複，自動抓出本次專案有參與的工廠
            unique_factories = df_hwln_items[['Factory ID', 'Factory Name']].drop_duplicates().dropna()
            
            # 將這些工廠與上傳的 Factory List 詳細資料(如合規分數、到期日)進行合併
            if 'Facility ID' in df_factory_raw.columns:
                df_project_factories = pd.merge(unique_factories, df_factory_raw, left_on='Factory ID', right_on='Facility ID', how='left')
            else:
                df_project_factories = df_factory_raw # 如果無法對應，就直接輸出原表


            # --- 建立豁免清單 (Exemption Request) ---
            # 範例邏輯：挑選特定工廠或特定 FRP Level 的品項放入此表 (可自行修改條件)
            df_exemption = df_hwln_items.copy()
            df_exemption['Justification'] = "" # 新增說明欄位
            df_exemption['Item Status (New/CF)'] = "CF"# 使用 BytesIO 在記憶體中建立 Excel，不需存入硬碟
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # 將各個 DataFrame 寫入不同的 Sheet
                df_hwln_items.to_excel(writer, index=False, sheet_name='26C5 HWLN ITEMS')
                df_exemption.to_excel(writer, index=False, sheet_name='Exemption request')
                df_project_factories.to_excel(writer, index=False, sheet_name='Factory list')
                
                # 若需要，可在此處加入 XlsxWriter 的格式化邏輯 (例如自動調整欄寬、上色)
                workbook = writer.book
                worksheet = writer.sheets['26C5 HWLN ITEMS']
                header_format = workbook.add_format({'bold': True, 'bg_color': '#FFD966'}) # 萬聖節黃/橘色表頭
                
                for col_num, value in enumerate(df_hwln_items.columns.values):
                    worksheet.write(0, col_num, value, header_format)

            processed_data = output.getvalue()
            
            st.success("✅ Tracking Grid 生成成功！")
            
            # 提供下載按鈕
            st.download_button(
                label="📥 下載 Halloween Item Tracking Grid.xlsx",
                data=processed_data,
                file_name="Automated_Halloween_Tracking_Grid.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("⚠️ 請先上傳 Data 與 Factory List 檔案才能執行。")