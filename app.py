import streamlit as st
import pandas as pd
import io

# ==========================================
# 1. 頁面基本設定與標題
# ==========================================
st.set_page_config(page_title="Halloween Item Tracking Grid Generator", layout="wide")
st.title("🎃 萬聖節專案 Tracking Grid 自動生成工具")
st.markdown("請上傳系統匯出的原始資料與工廠清單，系統將自動為您生成 Tracking Grid。")

# ==========================================
# 2. 建立檔案上傳區塊
# ==========================================
col1, col2, col3 = st.columns(3)

with col1:
    file_data = st.file_uploader("上傳 Data (CSV/Excel)", type=["csv", "xlsx"])
with col2:
    file_data_cos = st.file_uploader("上傳 Data Cos (CSV/Excel)", type=["csv", "xlsx"])
with col3:
    # 這裡已經修正：將 if st.button 移出並換行，避免語法錯誤
    file_factory = st.file_uploader("上傳 Factory List (CSV/Excel)", type=["csv", "xlsx"])

st.divider() # 加上一條分隔線讓版面更清楚

# ==========================================
# 3. 核心處理邏輯 (按下按鈕後執行)
# ==========================================
# 確保 if 判斷式獨立於上述的 with 區塊之外
if st.button("生成 Tracking Grid", type="primary"):
    
    # 檢查必要的檔案是否都已上傳 (目前預設需要 Data 和 Factory List)
    if file_data and file_factory: 
        with st.spinner("資料處理中，請稍候..."):
            
            try:
                # --- 步驟 A: 讀取上傳的檔案 ---
                df_data = pd.read_csv(file_data) if file_data.name.endswith('.csv') else pd.read_excel(file_data)
                df_factory_raw = pd.read_csv(file_factory) if file_factory.name.endswith('.csv') else pd.read_excel(file_factory)
                
                # --- 步驟 B: 篩選與重新命名 Data 所需欄位 ---
                # 這裡設定您預期原始資料中會有的欄位名稱
                core_columns = ['DPCI', 'Product Description', 'Import Vendor ID', 'Import Vendor Name']
                
                # 檢查欄位是否存在，避免程式找不到欄位而報錯
                available_cols = [col for col in core_columns if col in df_data.columns]
                df_main = df_data[available_cols].copy()
                
                # 將原始欄位名稱替換成 Tracking Grid 上需要的名稱
                df_main.rename(columns={
                    'Product Description': 'ITEM_DESC',
                    'Import Vendor ID': 'Factory ID',
                    'Import Vendor Name': 'Factory Name'
                }, inplace=True)
                
                # --- 步驟 C: 建立主表 (26C5 HWLN ITEMS) ---
                tracking_columns = ["FRP Level", "Tollgate Exempt", "TPR Lite/Exempt", "Tollgate Date", "TPR Date", "Result"]
                for col in tracking_columns:
                    df_main[col] = "" # 建立空欄位供後續手動填寫或未來進一步自動判斷
                    
                # 設定主表最終輸出的欄位順序
                final_main_cols = ['DPCI', 'ITEM_DESC', 'FRP Level', 'Tollgate Exempt', 'TPR Lite/Exempt', 'Factory Name', 'Factory ID', 'Tollgate Date', 'TPR Date', 'Result']
                df_hwln_items = df_main[[c for c in final_main_cols if c in df_main.columns]]
                
                # --- 步驟 D: 建立工廠清單 (Factory list) ---
                # 抓出本次有參與的工廠 ID 與名稱 (去重複)
                if 'Factory ID' in df_hwln_items.columns:
                    unique_factories = df_hwln_items[['Factory ID', 'Factory Name']].drop_duplicates().dropna()
                    
                    # 嘗試與上傳的 Factory List 進行合併 (假設 Factory List 裡面的 ID 欄位叫做 'Facility ID')
                    if 'Facility ID' in df_factory_raw.columns:
                        df_project_factories = pd.merge(unique_factories, df_factory_raw, left_on='Factory ID', right_on='Facility ID', how='left')
                    else:
                        df_project_factories = df_factory_raw
                else:
                    df_project_factories = df_factory_raw

                # --- 步驟 E: 建立豁免清單 (Exemption request) ---
                # 先複製主表，並加上特定欄位做為雛型
                df_exemption = df_hwln_items.copy()
                df_exemption['Justification'] = "" 
                df_exemption['Item Status (New/CF)'] = "CF"
                
                # --- 步驟 F: 匯出為多活頁簿的 Excel 檔案 ---
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    # 將各個 DataFrame 寫入不同的 Sheet
                    df_hwln_items.to_excel(writer, index=False, sheet_name='26C5 HWLN ITEMS')
                    df_exemption.to_excel(writer, index=False, sheet_name='Exemption request')
                    df_project_factories.to_excel(writer, index=False, sheet_name='Factory list')
                    
                    # 針對主表做簡單的格式化 (黃底粗體表頭)
                    workbook = writer.book
                    worksheet = writer.sheets['26C5 HWLN ITEMS']
                    header_format = workbook.add_format({'bold': True, 'bg_color': '#FFD966'}) 
                    
                    for col_num, value in enumerate(df_hwln_items.columns.values):
                        worksheet.write(0, col_num, value, header_format)

                processed_data = output.getvalue()
                
                st.success("✅ Tracking Grid 處理完成！請點擊下方按鈕下載。")
                
                # ==========================================
                # 4. 提供下載按鈕
                # ==========================================
                st.download_button(
                    label="📥 下載 Halloween Item Tracking Grid.xlsx",
                    data=processed_data,
                    file_name="Automated_Halloween_Tracking_Grid.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            except Exception as e:
                # 捕捉其他可能的資料處理錯誤並顯示在畫面上，方便除錯
                st.error(f"❌ 處理檔案時發生錯誤: {e}")
                st.info("提示：請確認上傳的原始檔案內是否包含預期的欄位名稱。")

    else:
        st.warning("⚠️ 請務必上傳『Data』與『Factory List』檔案後，再點擊生成。")
