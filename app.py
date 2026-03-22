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
st.markdown("請上傳專案檔案。系統將自動讀取 Master Sheet 與 Data 進行解析、計數與智慧排版。")

# ==========================================
# 2. 檔案上傳與精簡選項
# ==========================================
st.markdown("### 📄 步驟 1：上傳資料檔案 (Excel / CSV)")
uploaded_files = st.file_uploader("📁 請將 [Program Sheet] 與 [Data 表] 一同拖曳至此", 
                                  type=["xlsx", "xls", "csv"], 
                                  accept_multiple_files=True)

st.markdown("### 🖼️ 步驟 2：選擇圖片來源")
img_option = st.radio("請選擇您的圖片提供方式：", [
    "1. 🗂️ 從 Master Sheet 卡片自動萃取 (抓取 DPCI 上方的圖片)",
    "2. 📁 上傳 ZIP 壓縮檔 (檔名需對應 DPCI)"
])

uploaded_zip = None
if img_option.startswith("2"):
    uploaded_zip = st.file_uploader("📁 請上傳 .zip 圖片壓縮檔", type=["zip"])

st.divider() 

# ==========================================
# 3. 核心處理邏輯
# ==========================================
if uploaded_files:
    # --- 自動分類檔案 ---
    master_file = None
    data_file = None
    
    for file in uploaded_files:
        name_upper = file.name.upper()
        if "PROGRAM" in name_upper or "MASTER" in name_upper:
            master_file = file
        elif "DATA" in name_upper or "2026" in name_upper or "68" in name_upper:
            data_file = file

    # 若無法完全自動辨識，先以使用者上傳順序盲猜
    if not master_file and len(uploaded_files) > 0: master_file = uploaded_files[0]
    if not data_file and len(uploaded_files) > 1: data_file = uploaded_files[1]

    if st.button("✨ 智慧生成 Program Items", type="primary"):
        with st.spinner("解析資料與自動排版中，請稍候..."):
            
            with tempfile.TemporaryDirectory() as temp_dir:
                try:
                    # --- 處理 ZIP 圖片包 ---
                    if img_option.startswith("2") and uploaded_zip:
                        with zipfile.ZipFile(uploaded_zip, 'r') as zip_ref:
                            zip_ref.extractall(temp_dir)

                    # ---------------------------------------------------------
                    # 步驟 A: 建立 CATEGORY 對照字典 (自動尋找 Data Sheet)
                    # ---------------------------------------------------------
                    cat_mapping = {}
                    if data_file:
                        if data_file.name.endswith('.csv'):
                            df_data = pd.read_csv(io.BytesIO(data_file.getvalue()), header=None)
                        else:
                            xls_data = pd.ExcelFile(io.BytesIO(data_file.getvalue()))
                            # 自動尋找 Data 分頁
                            target_sheet = xls_data.sheet_names[0]
                            for s in xls_data.sheet_names:
                                if "DATA" in s.upper(): target_sheet = s; break
                            df_data = pd.read_excel(xls_data, sheet_name=target_sheet, header=None)
                            
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

                    # ---------------------------------------------------------
                    # 步驟 B: 解析 Master Sheet 卡片資料
                    # ---------------------------------------------------------
                    parsed_items = []
                    
                    xls_master = pd.ExcelFile(io.BytesIO(master_file.getvalue()))
                    m_sheet = xls_master.sheet_names[-1] # 預設最後一頁
                    for s in xls_master.sheet_names:
                        if "MASTER" in s.upper(): m_sheet = s; break
                        
                    wb = openpyxl.load_workbook(io.BytesIO(master_file.getvalue()), data_only=True)
                    sheet = wb[m_sheet]
                    
                    image_loader = None
                    if img_option.startswith("1"):
                        try:
                            image_loader = SheetImageLoader(sheet)
                        except: pass
                    
                    for r in range(1, sheet.max_row + 1):
                        for c in range(1, sheet.max_column + 1):
                            val = str(sheet.cell(row=r, column=c).value).strip().upper()
                            
                            if val == 'DPCI:':
                                dpci = str(sheet.cell(row=r, column=c+1).value).strip()
                                if dpci.lower() == 'none': dpci = ""
                                
                                # 抓取圖片並以 DPCI 命名存入暫存區
                                if img_option.startswith("1") and image_loader:
                                    img_obj = None
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

                    if not parsed_items:
                        st.warning("⚠️ 在 Master Sheet 中未偵測到任何含有 'DPCI:' 的卡片。")
                        st.stop()
                    
                    df_out = pd.DataFrame(parsed_items)

                    # ---------------------------------------------------------
                    # 步驟 C: VLOOKUP 與群組化排序
                    # ---------------------------------------------------------
                    if cat_mapping:
                        clean_main_dpci = df_out['DPCI'].astype(str).str.replace('-', '').str.strip()
                        df_out['CATEGORY'] = clean_main_dpci.map(cat_mapping).fillna('')
                    else:
                        df_out['CATEGORY'] = ""

                    # 針對 Factory Name 進行排序，確保相同工廠的資料相連 (利於合併儲存格)
                    df_out.sort_values(by='Factory Name', inplace=True)
                    df_out.reset_index(drop=True, inplace=True)

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
                        if col not in df_out.columns: df_out[col] = ""

                    # ---------------------------------------------------------
                    # 步驟 D: XlsxWriter 進階排版 (合併儲存格、粗外框、頂部 L1 總數)
                    # ---------------------------------------------------------
                    output = io.BytesIO()
                    img_insert_count = 0
                    
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        workbook = writer.book
                        worksheet = workbook.add_worksheet('Program Items')
                        
                        # 定義動態格式產生器 (用於繪製粗外框)
                        format_cache = {}
                        def get_fmt(t=1, b=1, l=1, r=1, align=None):
                            key = (t, b, l, r, align)
                            if key not in format_cache:
                                props = {'font_name': 'Arial', 'valign': 'vcenter', 'text_wrap': True,
                                         'top': t, 'bottom': b, 'left': l, 'right': r}
                                if align: props['align'] = align
                                format_cache[key] = workbook.add_format(props)
                            return format_cache[key]

                        # 1. 寫入 L1 紫色粗體總數 (索引列0, 欄11)
                        fmt_total = workbook.add_format({'font_name': 'Arial', 'bold': True, 'font_color': '#800080', 'font_size': 12, 'align': 'center', 'valign': 'vcenter'})
                        worksheet.write(0, 11, len(df_out), fmt_total)
                        
                        # 2. 寫入標題列 (索引列1)
                        fmt_header = workbook.add_format({'bold': True, 'bg_color': '#FFD966', 'border': 1, 'font_name': 'Arial', 'valign': 'vcenter', 'text_wrap': True}) 
                        for c, col_name in enumerate(target_columns):
                            worksheet.write(1, c, col_name, fmt_header)
                            worksheet.set_column(c, c, 15) # 預設欄寬
                            
                        worksheet.set_column(3, 3, 16) # PHOTO 欄寬
                        worksheet.set_column(9, 9, 25) # Factory Name 欄寬加寬
                        
                        # 3. 找出每個工廠群組的起始與結束索引
                        factories = df_out['Factory Name'].tolist()
                        groups = []
                        start_idx = 0
                        for i in range(1, len(factories)):
                            if factories[i] != factories[start_idx]:
                                groups.append((start_idx, i - 1, factories[start_idx]))
                                start_idx = i
                        groups.append((start_idx, len(factories) - 1, factories[start_idx]))

                        # 4. 逐列寫入資料並設定外框
                        excel_row_offset = 2 # 資料從 Excel 第 3 列開始
                        for s_idx, e_idx, f_name in groups:
                            for i in range(s_idx, e_idx + 1):
                                excel_row = i + excel_row_offset
                                worksheet.set_row(excel_row, 80) # 放大行高容納圖片
                                
                                # 取得 DPCI 對應的圖片
                                dpci_val = str(df_out.iloc[i]['DPCI']).strip()
                                safe_name = "".join(x for x in dpci_val if x.isalnum() or x in "-_")
                                if safe_name.endswith('.0'): safe_name = safe_name[:-2]
                                
                                img_path = None
                                search_names = [f"{safe_name}.png", f"{safe_name}.jpg", f"{dpci_val.lower()}.png"]
                                for root, dirs, files in os.walk(temp_dir):
                                    for file in files:
                                        if file.lower() in [s.lower() for s in search_names]:
                                            img_path = os.path.join(root, file)
                                            break
                                    if img_path: break
                                
                                # 逐欄寫入
                                for c, col_name in enumerate(target_columns):
                                    if c == 9: continue # 跳過 Factory Name (稍後合併寫入)
                                    
                                    # 動態判斷粗外框 (群組頂部與底部為粗線 = 2)
                                    t_border = 2 if i == s_idx else 1
                                    b_border = 2 if i == e_idx else 1
                                    l_border = 2 if c == 0 else 1
                                    r_border = 2 if c == len(target_columns) - 1 else 1
                                    
                                    fmt = get_fmt(t_border, b_border, l_border, r_border)
                                    
                                    if c == 3: # PHOTO
                                        worksheet.write(excel_row, c, "", fmt)
                                        if img_path:
                                            try:
                                                with Image.open(img_path) as img:
                                                    img.thumbnail((100, 100))
                                                    resized_path = os.path.join(temp_dir, f"resized_{i}.png")
                                                    img.save(resized_path, "PNG")
                                                    worksheet.insert_image(excel_row, c, resized_path, {'x_offset': 5, 'y_offset': 5})
                                                    img_insert_count += 1
                                            except: pass
                                            
                                    elif c == 11: # Total SKU per Factory (僅首列顯示計數)
                                        if i == s_idx:
                                            worksheet.write(excel_row, c, e_idx - s_idx + 1, fmt)
                                        else:
                                            worksheet.write(excel_row, c, "", fmt)
                                            
                                    else:
                                        worksheet.write(excel_row, c, df_out.iloc[i][col_name], fmt)
                                        
                            # 單獨處理 Factory Name (垂直合併儲存格)
                            fmt_fact = get_fmt(t=2, b=2, l=1, r=1, align='center')
                            if s_idx == e_idx:
                                worksheet.write(s_idx + excel_row_offset, 9, f_name, fmt_fact)
                            else:
                                worksheet.merge_range(s_idx + excel_row_offset, 9, e_idx + excel_row_offset, 9, f_name, fmt_fact)

                    processed_data = output.getvalue()
                    
                    st.success(f"✅ 處理完成！共解析出 **{len(df_out)}** 筆商品，並成功置入 **{img_insert_count}** 張圖片。")
                    
                    st.download_button(
                        label="📥 下載 格式化 Program Items.xlsx",
                        data=processed_data,
                        file_name="Automated_Program_Items.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                except Exception as e:
                    st.error(f"❌ 處理檔案時發生錯誤: {e}")

else:
    st.info("💡 提示：請在上方直接拖曳上傳您的專案 Excel/CSV 檔案 (不限順序)。")
