import streamlit as st
import pandas as pd
import io
import os
import tempfile
import zipfile
import openpyxl
from openpyxl_image_loader import SheetImageLoader
from openpyxl.utils import get_column_letter
from PIL import Image

# ==========================================
# 1. 頁面基本設定與標題
# ==========================================
st.set_page_config(page_title="Program Items Generator", layout="wide")
st.title("🎃 萬聖節專案 Program Items 自動生成工具")
st.markdown("請上傳專案檔案 (可一次選取/拖曳多個檔案)。系統將自動解析 Master Sheet，並結合 Data 表的 Subclass Name 與圖片。")

# ==========================================
# 2. 檔案上傳與選項區塊
# ==========================================
st.markdown("### 📄 步驟 1：上傳資料檔案")
uploaded_files = st.file_uploader("📁 請上傳 Excel / CSV 檔案 (Master Sheet 與 Data 表)", 
                                  type=["xlsx", "xls", "csv"], 
                                  accept_multiple_files=True)

st.markdown("### 🖼️ 步驟 2：圖片來源 (跟隨您的原始邏輯)")
img_option = st.radio("請選擇您的圖片提供方式：", [
    "1. 📊 從 Data 表自動萃取 (自動尋找 Thumbnail 或 Photo 欄位)",
    "2. 🗂️ 從 Master Sheet 卡片自動萃取 (抓取 DPCI 上方的圖片)",
    "3. 📁 上傳 ZIP 壓縮檔 (檔名需對應 DPCI)",
    "4. ❌ 不需要抓取圖片"
])

uploaded_zip = None
if img_option.startswith("3"):
    uploaded_zip = st.file_uploader("📁 請上傳 .zip 圖片壓縮檔", type=["zip"])

st.divider() 

# ==========================================
# 3. 核心處理邏輯
# ==========================================
if uploaded_files:
    sheet_options = []
    df_dict = {}
    file_bytes_dict = {} 
    
    with st.spinner("讀取檔案結構中..."):
        for file in uploaded_files:
            file_bytes = file.getvalue()
            file_bytes_dict[file.name] = file_bytes
            
            if file.name.endswith('.csv'):
                df = pd.read_csv(io.BytesIO(file_bytes), header=None)
                name = f"[{file.name}] CSV"
                sheet_options.append(name)
                df_dict[name] = df
            else:
                xls = pd.ExcelFile(io.BytesIO(file_bytes))
                for sheet in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=sheet, header=None)
                    name = f"[{file.name}] {sheet}"
                    sheet_options.append(name)
                    df_dict[name] = df
                    
    # 【修正點 1】嚴格比對「工作表名稱」而非檔案名稱，避免選到最後一個工廠 Sheet
    default_master_idx = 0
    default_data_idx = 0
    for i, name in enumerate(sheet_options):
        sheet_name_only = name.split("] ")[1].upper() # 只拿工作表名稱來比對
        if 'MASTER' in sheet_name_only or 'PROGRAM' in sheet_name_only:
            default_master_idx = i
            break # 找到就停止，確保是真正的 Master
            
    for i, name in enumerate(sheet_options):
        sheet_name_only = name.split("] ")[1].upper()
        if 'DATA' in sheet_name_only:
            default_data_idx = i
            break

    col1, col2 = st.columns(2)
    with col1:
        selected_master = st.selectbox("📄 1. 請選擇【卡片主資料表 (Master Sheet)】：", sheet_options, index=default_master_idx)
    with col2:
        data_options = ["(不使用對照表)"] + sheet_options
        selected_data = st.selectbox("📄 2. 請選擇【CATEGORY 對照表 (尋找 Subclass Name)】：", data_options, index=default_data_idx + 1 if sheet_options else 0)

    if st.button("✨ 生成 Program Items", type="primary"):
        with st.spinner("解析卡片與處理圖片中，這可能需要數十秒..."):
            
            # 使用 temp_dir 統一管理圖片 (與您的邏輯完全一致)
            with tempfile.TemporaryDirectory() as temp_dir:
                try:
                    # --- 圖片來源 3: 解壓縮 ZIP ---
                    if img_option.startswith("3") and uploaded_zip:
                        with zipfile.ZipFile(uploaded_zip, 'r') as zip_ref:
                            zip_ref.extractall(temp_dir)

                    # ---------------------------------------------------------
                    # 步驟 A: 建立 CATEGORY 對照字典 & 圖片來源 1 (Data 表)
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
                            dpci_col_cat = cat_cols_map.get("DPCI", cat_cols_map.get("DPCI#"))
                            subclass_col = cat_cols_map.get("SUBCLASSNAME")
                            
                            if dpci_col_cat and subclass_col:
                                clean_dpci = df_data[dpci_col_cat].astype(str).str.replace("-", "").str.strip()
                                clean_dpci = clean_dpci.apply(lambda x: x[:-2] if x.endswith('.0') else x)
                                cat_mapping = dict(zip(clean_dpci, df_data[subclass_col]))

                        # --- 圖片來源 1: 從 Data 表萃取 ---
                        if img_option.startswith("1") and not selected_data.endswith("CSV"):
                            data_file_name = selected_data.split("]")[0][1:]
                            data_sheet_name = selected_data.split("] ")[1]
                            wb_data = openpyxl.load_workbook(io.BytesIO(file_bytes_dict[data_file_name]), data_only=True)
                            sheet_data = wb_data[data_sheet_name]
                            
                            try:
                                image_loader_data = SheetImageLoader(sheet_data)
                                d_header_row, thumb_col, d_dpci_col = None, None, None
                                for r in range(1, min(20, sheet_data.max_row + 1)):
                                    for c in range(1, sheet_data.max_column + 1):
                                        val = str(sheet_data.cell(row=r, column=c).value).strip().lower()
                                        if val in ['thumbnail', 'photo', 'image']:
                                            thumb_col, d_header_row = c, r
                                        elif val in ['dpci', 'pid', 'spark pid']:
                                            d_dpci_col = c
                                    if thumb_col and d_dpci_col: break
                                
                                if thumb_col and d_dpci_col:
                                    thumb_letter = get_column_letter(thumb_col)
                                    for r in range(d_header_row + 1, sheet_data.max_row + 1):
                                        dpci_val = str(sheet_data.cell(row=r, column=d_dpci_col).value).strip()
                                        safe_name = "".join(x for x in dpci_val if x.isalnum() or x in "-_")
                                        if safe_name.endswith('.0'): safe_name = safe_name[:-2]
                                        if not safe_name: continue
                                        
                                        img_cell = f"{thumb_letter}{r}"
                                        if image_loader_data.image_in(img_cell):
                                            try:
                                                img = image_loader_data.get(img_cell)
                                                img.save(os.path.join(temp_dir, f"{safe_name}.png"), "PNG")
                                            except: pass
                            except Exception:
                                pass # 若無圖片則略過

                    # ---------------------------------------------------------
                    # 步驟 B: 解析 Master Sheet 卡片資料 & 圖片來源 2 (卡片)
                    # ---------------------------------------------------------
                    parsed_items = []
                    master_file_name = selected_master.split("]")[0][1:] 
                    master_sheet_name = selected_master.split("] ")[1]
                    
                    wb = openpyxl.load_workbook(io.BytesIO(file_bytes_dict[master_file_name]), data_only=True)
                    sheet = wb[master_sheet_name]
                    
                    image_loader = None
                    if img_option.startswith("2"):
                        try:
                            image_loader = SheetImageLoader(sheet)
                        except: pass
                    
                    for r in range(1, sheet.max_row + 1):
                        for c in range(1, sheet.max_column + 1):
                            val = str(sheet.cell(row=r, column=c).value).strip().upper()
                            
                            if val == 'DPCI:':
                                dpci = str(sheet.cell(row=r, column=c+1).value).strip()
                                if dpci.lower() == 'none': dpci = ""
                                
                                # --- 圖片來源 2: 卡片掃描 ---
                                if img_option.startswith("2") and image_loader:
                                    img_obj = None
                                    # 擴大搜索範圍：掃描 DPCI 正上方與周圍的格子
                                    for row_offset in [1, 2, 0]: 
                                        if r - row_offset < 1: continue
                                        for col_offset in range(5):
                                            img_cell = f"{get_column_letter(c + col_offset)}{r - row_offset}"
                                            if image_loader.image_in(img_cell):
                                                try:
                                                    img_obj = image_loader.get(img_cell)
                                                    break
                                                except: pass
                                        if img_obj: break
                                        
                                    if img_obj:
                                        safe_name = "".join(x for x in dpci if x.isalnum() or x in "-_")
                                        if safe_name.endswith('.0'): safe_name = safe_name[:-2]
                                        try:
                                            img_obj.save(os.path.join(temp_dir, f"{safe_name}.png"), "PNG")
                                        except: pass
                                        
                                desc = ""
                                for i in range(1, 15):
                                    if r+i <= sheet.max_row:
                                        if str(sheet.cell(row=r+i, column=c).value).strip().upper() == 'DESCRIPTION:':
                                            desc = str(sheet.cell(row=r+i, column=c+1).value).strip()
                                            if desc.lower() == 'none': desc = ""
                                            break
                                            
                                qty = ""
                                for i in range(1, 15):
                                    if r+i > sheet.max_row: break
                                    found_qty = False
                                    for j in range(c, c+6):
                                        if j <= sheet.max_column:
                                            if str(sheet.cell(row=r+i, column=j).value).strip().upper() == 'QTY:':
                                                qty = str(sheet.cell(row=r+i, column=j+1).value).strip()
                                                if qty.lower() == 'none': qty = ""
                                                if qty.endswith('.0'): qty = qty[:-2]
                                                found_qty = True
                                                break
                                    if found_qty: break
                                    
                                factory_name = ""
                                factory_id = ""
                                for i in range(1, 15):
                                    if r+i > sheet.max_row: break
                                    cell_val = str(sheet.cell(row=r+i, column=c).value).strip()
                                    if cell_val.upper().startswith('FACTORY:') or cell_val.upper().startswith('"FACTORY:'):
                                        factory_str = cell_val.replace('"', '')
                                        if ':' in factory_str:
                                            factory_str = factory_str.split(':', 1)[1].strip()
                                        parts = factory_str.split('/')
                                        if len(parts) >= 1: factory_name = parts[0].strip()
                                        if len(parts) >= 2: factory_id = parts[1].strip()
                                        break
                                        
                                parsed_items.append({
                                    'DPCI': dpci,
                                    'ITEM_DESC': desc,
                                    'Factory Name': factory_name,
                                    'Factory ID': factory_id,
                                    'QTY': qty
                                })

                    # ---------------------------------------------------------
                    # 步驟 C: 重組匯出與【置入圖片】
                    # ---------------------------------------------------------
                    if not parsed_items:
                        st.warning("⚠️ 在 Master Sheet 中未偵測到任何含有 'DPCI:' 的卡片。")
                        st.stop()
                    
                    df_out = pd.DataFrame(parsed_items)

                    if cat_mapping:
                        clean_main_dpci = df_out['DPCI'].astype(str).str.replace('-', '').str.strip()
                        df_out['CATEGORY'] = clean_main_dpci.map(cat_mapping).fillna('')
                    else:
                        df_out['CATEGORY'] = ""

                    target_columns = [
                        "DPCI", "CATEGORY", "ITEM_DESC", "PHOTO", "FRP Level", 
                        "Red Seal(Y/N)", "CF item( Y/N )", "Tollgate Exempt", 
                        "TPR Lite/Exempt", "Factory Name", "Factory ID", 
                        "Total SKU per Factory", "QTY", 
                        "Tollgate Date", "TPR Date", "Dupro Date", "Result", 
                        "TOP Result", "FRI plan", "Port of Export", 
                        "1st Ship window", "Inspection Office"
                    ]

                    for col in target_columns:
                        if col not in df_out.columns:
                            df_out[col] = ""
                    df_out = df_out[target_columns]

                    output = io.BytesIO()
                    img_insert_count = 0 # 計算成功插入的圖片數
                    
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        df_out.to_excel(writer, index=False, sheet_name='Program Items')
                        
                        workbook = writer.book
                        worksheet = writer.sheets['Program Items']
                        
                        workbook.formats[0].set_font_name('Arial')
                        cell_format = workbook.add_format({'font_name': 'Arial', 'valign': 'vcenter'})
                        header_format = workbook.add_format({'bold': True, 'bg_color': '#FFD966', 'border': 1, 'font_name': 'Arial'}) 
                        
                        for col_num, value in enumerate(df_out.columns.values):
                            worksheet.write(0, col_num, value, header_format)
                            worksheet.set_column(col_num, col_num, 15, cell_format)
                            
                        # 設定 PHOTO 欄位寬度
                        col_idx_photo = 3
                        worksheet.set_column(col_idx_photo, col_idx_photo, 16)
                        
                        # --- 將暫存資料夾內的圖片與 DPCI 比對並置入 ---
                        for row_num, item in enumerate(parsed_items):
                            excel_row = row_num + 1
                            worksheet.set_row(excel_row, 80) # 放大行高以容納圖片
                            
                            dpci_val = str(item['DPCI']).strip()
                            safe_name = "".join(x for x in dpci_val if x.isalnum() or x in "-_")
                            if safe_name.endswith('.0'): safe_name = safe_name[:-2]
                            
                            img_path = None
                            # 建立多種可能的檔名對應陣列
                            search_names = [f"{safe_name}.png", f"{safe_name}.jpg", f"{safe_name}.jpeg", 
                                            f"{dpci_val.lower()}.png", f"{dpci_val.lower()}.jpg"]
                            
                            for root, dirs, files in os.walk(temp_dir):
                                for file in files:
                                    if file.lower() in [s.lower() for s in search_names]:
                                        img_path = os.path.join(root, file)
                                        break
                                if img_path: break
                                
                            # 若有比對到圖片，調整大小後置入 Excel
                            if img_path:
                                try:
                                    with Image.open(img_path) as img:
                                        img.thumbnail((100, 100)) # 等比例縮小最大邊界
                                        resized_path = os.path.join(temp_dir, f"resized_{row_num}.png")
                                        img.save(resized_path, "PNG")
                                        worksheet.insert_image(excel_row, col_idx_photo, resized_path, {'x_offset': 5, 'y_offset': 5})
                                        img_insert_count += 1
                                except Exception:
                                    pass

                    processed_data = output.getvalue()
                    
                    st.success(f"✅ 處理完成！共解析出 **{len(parsed_items)}** 筆商品，並成功配對置入 **{img_insert_count}** 張圖片。")
                    
                    st.download_button(
                        label="📥 下載 Program Items.xlsx",
                        data=processed_data,
                        file_name="Automated_Program_Items.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                except Exception as e:
                    st.error(f"❌ 處理檔案時發生錯誤: {e}")

else:
    st.info("💡 提示：請在上方上傳相關的 Excel/CSV 檔案！")
