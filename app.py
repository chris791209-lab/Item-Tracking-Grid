import streamlit as st
import pandas as pd
import io

# ==========================================
# 1. 頁面基本設定與標題
# ==========================================
st.set_page_config(page_title="Program Items Generator", layout="wide")
st.title("🎃 萬聖節專案 Program Items 自動生成工具")
st.markdown("請上傳專案檔案 (可一次選取/拖曳多個檔案)。系統將自動解析 Master Sheet 的卡片資料，並結合 Data 表的 Subclass Name。")

# ==========================================
# 2. 單一檔案上傳區塊 (支援多檔同時上傳)
# ==========================================
uploaded_files = st.file_uploader("📁 請上傳 Excel / CSV 檔案 (可同時上傳 Master Sheet 與 Data)", 
                                  type=["xlsx", "xls", "csv"], 
                                  accept_multiple_files=True)

st.divider() 

# 數值清理函數 (避免 Excel 將 34287 讀成 34287.0)
def clean_val(v):
    s = str(v).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s if s != 'nan' else ""

# ==========================================
# 3. 核心處理邏輯
# ==========================================
if uploaded_files:
    # 建立所有可用工作表的選單
    sheet_options = []
    df_dict = {}
    
    with st.spinner("讀取檔案結構中..."):
        for file in uploaded_files:
            if file.name.endswith('.csv'):
                df = pd.read_csv(file, header=None)
                name = f"[{file.name}] CSV"
                sheet_options.append(name)
                df_dict[name] = df
            else:
                xls = pd.ExcelFile(file)
                for sheet in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=sheet, header=None)
                    name = f"[{file.name}] {sheet}"
                    sheet_options.append(name)
                    df_dict[name] = df
                    
    # 自動預選最可能的工作表
    default_master_idx = 0
    default_data_idx = 0
    for i, name in enumerate(sheet_options):
        if 'MASTER' in name.upper() or 'PROGRAM' in name.upper():
            default_master_idx = i
        if 'DATA' in name.upper():
            default_data_idx = i

    col1, col2 = st.columns(2)
    with col1:
        selected_master = st.selectbox("📄 1. 請選擇【卡片主資料表 (Master Sheet)】：", sheet_options, index=default_master_idx)
    with col2:
        data_options = ["(不使用對照表)"] + sheet_options
        selected_data = st.selectbox("📄 2. 請選擇【CATEGORY 對照表 (尋找 Subclass Name)】：", data_options, index=default_data_idx + 1 if sheet_options else 0)

    if st.button("生成 Program Items", type="primary"):
        with st.spinner("解析卡片與資料比對中，請稍候..."):
            try:
                # ---------------------------------------------------------
                # 步驟 A: 建立 CATEGORY 對照字典 (DPCI -> Subclass Name)
                # ---------------------------------------------------------
                cat_mapping = {}
                if selected_data != "(不使用對照表)":
                    df_data = df_dict[selected_data]
                    header_idx = -1
                    for i in range(min(20, len(df_data))):
                        if any('DPCI' in str(v).strip().upper() for v in df_data.iloc[i].values):
                            header_idx = i
                            break
                            
                    if header_idx != -1:
                        df_data.columns = df_data.iloc[header_idx]
                        df_data = df_data.iloc[header_idx + 1:].reset_index(drop=True)
                        
                        def normalize_col(col_name):
                            return str(col_name).replace('\n', '').replace('\r', '').replace(' ', '').upper()
                            
                        cat_cols_map = {normalize_col(c): c for c in df_data.columns}
                        
                        dpci_col = cat_cols_map.get("DPCI", cat_cols_map.get("DPCI#"))
                        subclass_col = cat_cols_map.get("SUBCLASSNAME")
                        
                        if dpci_col and subclass_col:
                            # 清理 DPCI 格式確保 100% 吻合
                            clean_dpci = df_data[dpci_col].astype(str).str.replace("-", "").str.strip()
                            clean_dpci = clean_dpci.apply(lambda x: x[:-2] if x.endswith('.0') else x)
                            cat_mapping = dict(zip(clean_dpci, df_data[subclass_col]))
                        else:
                            st.warning("⚠️ 警告：在對照表中找不到 'DPCI' 或 'Subclass Name' 欄位。")

                # ---------------------------------------------------------
                # 步驟 B: 解析 Master Sheet 卡片資料
                # ---------------------------------------------------------
                df_master = df_dict[selected_master]
                parsed_items = []
                
                # 掃描整張表，尋找卡片的起點 "DPCI:"
                for r in range(len(df_master)):
                    for c in range(len(df_master.columns)):
                        val = str(df_master.iloc[r, c]).strip().upper()
                        if val == 'DPCI:':
                            # 1. 抓取 DPCI
                            dpci = clean_val(df_master.iloc[r, c+1] if c+1 < len(df_master.columns) else "")
                            
                            # 2. 往下掃描找 Description:
                            desc = ""
                            for i in range(15):
                                if r+i >= len(df_master): break
                                if str(df_master.iloc[r+i, c]).strip().upper() == 'DESCRIPTION:':
                                    desc = clean_val(df_master.iloc[r+i, c+1] if c+1 < len(df_master.columns) else "")
                                    break
                                    
                            # 3. 往下掃描找 QTY: (QTY 可能在同行稍微右邊的欄位)
                            qty = ""
                            for i in range(15):
                                if r+i >= len(df_master): break
                                found_qty = False
                                for j in range(c, min(c+6, len(df_master.columns))):
                                    if str(df_master.iloc[r+i, j]).strip().upper() == 'QTY:':
                                        qty = clean_val(df_master.iloc[r+i, j+1] if j+1 < len(df_master.columns) else "")
                                        found_qty = True
                                        break
                                if found_qty: break
                                
                            # 4. 往下掃描找 Factory (並切分 Name 與 ID)
                            factory_name = ""
                            factory_id = ""
                            for i in range(15):
                                if r+i >= len(df_master): break
                                cell_val = str(df_master.iloc[r+i, c]).strip()
                                if cell_val.upper().startswith('FACTORY:') or cell_val.upper().startswith('"FACTORY:'):
                                    factory_str = cell_val.replace('"', '')
                                    if ':' in factory_str:
                                        factory_str = factory_str.split(':', 1)[1].strip()
                                    
                                    # 利用斜線拆分資料段落
                                    parts = factory_str.split('/')
                                    if len(parts) >= 1:
                                        factory_name = parts[0].strip()
                                    if len(parts) >= 2:
                                        factory_id = clean_val(parts[1])
                                    break
                                    
                            # 將解析出來的單個產品寫入清單
                            parsed_items.append({
                                'DPCI': dpci,
                                'ITEM_DESC': desc,
                                'PHOTO': '', # 保留圖片空白欄位
                                'Factory Name': factory_name,
                                'Factory ID': factory_id,
                                'QTY': qty
                            })

                df_out = pd.DataFrame(parsed_items)

                # ---------------------------------------------------------
                # 步驟 C: 重組 22 項最終欄位並匯出
                # ---------------------------------------------------------
                target_columns = [
                    "DPCI", "CATEGORY", "ITEM_DESC", "PHOTO", "FRP Level", 
                    "Red Seal(Y/N)", "CF item( Y/N )", "Tollgate Exempt", 
                    "TPR Lite/Exempt", "Factory Name", "Factory ID", 
                    "Total SKU per Factory", "QTY", 
                    "Tollgate Date", "TPR Date", "Dupro Date", "Result", 
                    "TOP Result", "FRI plan", "Port of Export", 
                    "1st Ship window", "Inspection Office"
                ]

                if not df_out.empty:
                    # 進行 Subclass Name 替換 CATEGORY 的 VLOOKUP
                    if cat_mapping:
                        clean_main_dpci = df_out['DPCI'].astype(str).str.replace('-', '').str.strip()
                        df_out['CATEGORY'] = clean_main_dpci.map(cat_mapping).fillna('')
                    else:
                        df_out['CATEGORY'] = ""

                    # 補齊其他指定空欄位
                    for col in target_columns:
                        if col not in df_out.columns:
                            df_out[col] = ""
                    
                    df_out = df_out[target_columns] # 依照指定的 22 欄位排序

                else:
                    st.warning("⚠️ 在 Master Sheet 中未偵測到任何含有 'DPCI:' 的卡片。")
                    st.stop()

                # 格式設定 (Arial, 粗體黃底)
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_out.to_excel(writer, index=False, sheet_name='Program Items')
                    
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
                    
                    for col_num, value in enumerate(df_out.columns.values):
                        worksheet.write(0, col_num, value, header_format)
                        worksheet.set_column(col_num, col_num, 15, cell_format)

                processed_data = output.getvalue()
                
                st.success("✅ 處理完成！已成功從卡片式排版中精準萃取所需資料。")
                st.info("💡 提示：Excel 內的產品圖片因屬浮動物件，程式已為您預留 PHOTO 空白欄位，供您後續快速貼上圖片。")
                
                st.download_button(
                    label="📥 下載 Program Items.xlsx",
                    data=processed_data,
                    file_name="Automated_Program_Items.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            except Exception as e:
                st.error(f"❌ 處理檔案時發生錯誤: {e}")

else:
    st.info("💡 提示：請在上方上傳相關的 Excel/CSV 檔案，您可以一次把多個檔案全選並拖曳進來！")
