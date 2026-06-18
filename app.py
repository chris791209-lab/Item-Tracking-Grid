import streamlit as st
import pandas as pd
import io
import os
import tempfile
import zipfile
import openpyxl
from openpyxl_image_loader import SheetImageLoader
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
from PIL import Image

# ==========================================
# 0. 頁面基本設定
# ==========================================
st.set_page_config(page_title="D240 Item Tracking Grid Generator", layout="wide")

# ==========================================
# 0. 密碼保護機制
# ==========================================
def check_password():
    """回傳 True 代表使用者輸入了正確的密碼"""
    def password_entered():
        if st.session_state["password"] == st.secrets["app_password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]  
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.title("🔒 系統登入")
        st.text_input(
            "🔒 請輸入 AE 部門共用密碼以啟用工具：", 
            type="password", 
            on_change=password_entered, 
            key="password"
        )
        return False
    
    elif not st.session_state["password_correct"]:
        st.title("🔒 系統登入")
        st.text_input(
            "🔒 請輸入 AE 部門共用密碼以啟用工具：", 
            type="password", 
            on_change=password_entered, 
            key="password"
        )
        st.error("❌ 密碼錯誤，請重新輸入。")
        return False
    else:
        return True

# 如果密碼驗證未通過，則阻擋往下執行
if not check_password():
    st.stop()

# ==========================================
# 1. 主程式標題 
# ==========================================
st.title("🎃 D240 Item Tracking Grid自動生成工具 (穩定適應版)")
st.markdown("請上傳專案檔案。系統將以 **Program Sheet** 為主視覺抓取圖片與排版，並自動從 **Data 表** 補齊缺失的工廠或數量資訊。")

# ==========================================
# 2. 檔案上傳與選項
# ==========================================
st.markdown("### 📄 步驟 1：上傳資料檔案")
uploaded_files = st.file_uploader("📁 請將 [Program Sheet] 與 [Data 表] 一同拖曳至此", 
                                  type=["xlsx", "xls", "csv"], 
                                  accept_multiple_files=True)

st.markdown("### 🖼️ 步驟 2：選擇圖片來源")
img_option = st.radio("請選擇您的圖片提供方式：", [
    "1. 🗂️ 從 Program / Master Sheet 卡片自動萃取 (包含大範圍智慧掃描)",
    "2. 📁 上傳 ZIP 壓縮檔 (檔名需對應 DPCI)"
])

uploaded_zip = None
if img_option.startswith("2"):
    uploaded_zip = st.file_uploader("📁 請上傳 .zip 圖片壓縮檔", type=["zip"])

st.divider() 

def clean_string(val):
    return str(val).replace(' ', '').replace(':', '').replace('#', '').upper()

# ==========================================
# 3. 核心處理邏輯
# ==========================================
if uploaded_files:
    master_file = None
    data_file = None
    
    for file in uploaded_files:
        fname = file.name.upper()
        if "TRACKING" in fname or "GRID" in fname or "AUTOMATED" in fname:
            continue
        if file.name.endswith('.csv') and ("DATA" in fname or "WRK" in fname):
            data_file = file
            continue
            
        try:
            xls_peek = pd.ExcelFile(io.BytesIO(file.getvalue()))
            snames = [s.upper() for s in xls_peek.sheet_names]
            if any("MASTER" in s for s in snames) or any("PROGRAM" in s for s in snames) or "PS" in snames:
                master_file = file
            if any("DATA" in s for s in snames) or any("PRODUCT" in s for s in snames):
                data_file = file
        except: pass

    if not master_file and len(uploaded_files) > 0: master_file = uploaded_files[0]
    if not data_file and len(uploaded_files) > 1: data_file = uploaded_files[1]

    if st.button("✨ 智慧生成 Item Tracking Grid", type="primary"):
        if not master_file:
            st.error("❌ 找不到 Program Sheet！請確認檔案是否正確。")
            st.stop()
            
        with st.spinner("解析卡片、萃取圖片與資料 VLOOKUP 中，請稍候..."):
            with tempfile.TemporaryDirectory() as temp_dir:
                try:
                    # --- 安全處理 ZIP ---
                    if img_option.startswith("2") and uploaded_zip:
                        zip_io = io.BytesIO(uploaded_zip.getvalue())
                        with zipfile.ZipFile(zip_io, 'r') as zip_ref:
                            zip_ref.extractall(temp_dir)

                    # ---------------------------------------------------------
                    # 步驟 A: 建立 Data 表字典
                    # ---------------------------------------------------------
                    cat_mapping = {}
                    fact_mapping = {}
                    qty_mapping = {}
                    
                    if data_file:
                        data_io = io.BytesIO(data_file.getvalue()) 
                        if data_file.name.endswith('.csv'):
                            df_data = pd.read_csv(data_io, header=None)
                        else:
                            xls_data = pd.ExcelFile(data_io)
                            target_sheet = xls_data.sheet_names[0]
                            for s in xls_data.sheet_names:
                                if "DATA" in s.upper() or "PRODUCT" in s.upper(): target_sheet = s; break
                            df_data = pd.read_excel(xls_data, sheet_name=target_sheet, header=None)
                            
                        header_idx = -1
                        for i in range(min(20, len(df_data))):
                            if any('DPCI' in str(v).strip().upper() for v in df_data.iloc[i].values):
                                header_idx = i; break
                                
                        if header_idx != -1:
                            df_data.columns = df_data.iloc[header_idx]
                            df_data = df_data.iloc[header_idx + 1:].reset_index(drop=True)
                            
                            def norm_c(col_name): return str(col_name).replace('\n', '').replace('\r', '').replace(' ', '').upper()
                            cat_cols_map = {norm_c(c): c for c in df_data.columns}
                            
                            dpci_col = cat_cols_map.get("DPCI", cat_cols_map.get("DPCI#"))
                            sub_col = cat_cols_map.get("SUBCLASSNAME", cat_cols_map.get("CATEGORY"))
                            fact_col = cat_cols_map.get("PRODUCTBUSINESSPARTNER", cat_cols_map.get("IMPORTVENDORNAME", cat_cols_map.get("VENDOR")))
                            qty_col = cat_cols_map.get("TOTALUNITS", cat_cols_map.get("QTY", cat_cols_map.get("ENTTTLRCPTU")))
                            
                            if dpci_col:
                                clean_dpci = df_data[dpci_col].astype(str).str.replace("-", "").str.strip()
                                clean_dpci = clean_dpci.apply(lambda x: x[:-2] if x.endswith('.0') else x)
                                
                                if sub_col: cat_mapping = dict(zip(clean_dpci, df_data[sub_col]))
                                if fact_col: fact_mapping = dict(zip(clean_dpci, df_data[fact_col]))
                                if qty_col: qty_mapping = dict(zip(clean_dpci, df_data[qty_col]))

                    # ---------------------------------------------------------
                    # 步驟 B: 解析 Program Sheet 卡片
                    # ---------------------------------------------------------
                    parsed_items = []
                    
                    master_io_pd = io.BytesIO(master_file.getvalue()) 
                    master_io_px = io.BytesIO(master_file.getvalue()) 
                    
                    xls_master = pd.ExcelFile(master_io_pd)
                    m_sheet = xls_master.sheet_names[-1]
                    for s in xls_master.sheet_names:
                        if "MASTER" in s.upper() or "PROGRAM" in s.upper() or "PS" in s.upper():
                            m_sheet = s; break
                            
                    wb = openpyxl.load_workbook(master_io_px, data_only=True)
                    sheet = wb[m_sheet]
                    
                    image_loader = None
                    if img_option.startswith("1"):
                        try: image_loader = SheetImageLoader(sheet)
                        except: pass
                    
                    for r in range(1, sheet.max_row + 1):
                        for c in range(1, sheet.max_column + 1):
                            cell_val = sheet.cell(row=r, column=c).value
                            if cell_val is None: continue
                            
                            val_clean = clean_string(cell_val)
                            
                            if val_clean == 'DPCI':
                                dpci = str(sheet.cell(row=r, column=c+1).value).strip()
                                if dpci.lower() == 'none' or dpci == "": 
                                    dpci = str(sheet.cell(row=r, column=c+2).value).strip()
                                if dpci.lower() == 'none': dpci = ""
                                if not dpci: continue
                                
                                # ==========================================
                                # 🖼️ 全新防死鎖圖片搜尋與儲存 
                                # ==========================================
                                img_obj = None
                                if img_option.startswith("1") and image_loader:
                                    found_img_cell = None
                                    
                                    # 第一層：雷達掃描
                                    for row_offset in range(-2, 17):
                                        for col_offset in range(-4, 9):
                                            c_idx = max(1, c + col_offset)
                                            r_idx = max(1, r - row_offset)
                                            img_cell = f"{get_column_letter(c_idx)}{r_idx}"
                                            if image_loader.image_in(img_cell):
                                                found_img_cell = img_cell
                                                break
                                        if found_img_cell: break
                                        
                                    if found_img_cell:
                                        try:
                                            raw_img = image_loader.get(found_img_cell)
                                            img_obj = raw_img.copy() 
                                        except: pass
                                    else:
                                        # 第二層：距離計算
                                        try:
                                            min_dist = 9999
                                            closest_cell = None
                                            for img_c in image_loader._images.keys():
                                                col_str, row_num = coordinate_from_string(img_c)
                                                img_col_idx = column_index_from_string(col_str)
                                                
                                                row_diff = r - row_num
                                                col_diff = abs(c - img_col_idx)
                                                
                                                if 0 <= row_diff <= 30 and col_diff <= 12:
                                                    dist = row_diff + col_diff
                                                    if dist < min_dist:
                                                        min_dist = dist
                                                        closest_cell = img_c
                                                        
                                            if closest_cell:
                                                raw_img = image_loader.get(closest_cell)
                                                img_obj = raw_img.copy() 
                                        except: pass
                                        
                                    if img_obj:
                                        safe_name = "".join(x for x in dpci if x.isalnum() or x in "-_")
                                        if safe_name.endswith('.0'): safe_name = safe_name[:-2]
                                        try: 
                                            img_obj.save(os.path.join(temp_dir, f"{safe_name}.png"), "PNG")
                                        except: pass
                                        
                                desc = ""
                                for i in range(1, 15):
                                    if r+i > sheet.max_row: break
                                    c_val = clean_string(sheet.cell(row=r+i, column=c).value)
                                    if 'DESCRIPTION' in c_val:
                                        desc = str(sheet.cell(row=r+i, column=c+1).value).strip()
                                        if desc.lower() == 'none' or desc == "": desc = str(sheet.cell(row=r+i, column=c+2).value).strip()
                                        if desc.lower() == 'none': desc = ""
                                        break
                                            
                                qty = ""
                                for i in range(1, 15):
                                    if r+i > sheet.max_row: break
                                    found_qty = False
                                    for j in range(c, min(c+8, sheet.max_column + 1)):
                                        q_val = clean_string(sheet.cell(row=r+i, column=j).value)
                                        if 'QTY' in q_val or 'CASEPACK' in q_val:
                                            qty = str(sheet.cell(row=r+i, column=j+1).value).strip()
                                            if qty.lower() == 'none' or qty == "": qty = str(sheet.cell(row=r+i, column=j+2).value).strip()
                                            if qty.lower() == 'none': qty = ""
                                            if qty.endswith('.0'): qty = qty[:-2]
                                            found_qty = True; break
                                    if found_qty: break
                                    
                                factory_name = ""
                                factory_id = ""
                                for i in range(1, 15):
                                    if r+i > sheet.max_row: break
                                    f_val = clean_string(sheet.cell(row=r+i, column=c).value)
                                    if 'FACTORY' in f_val or 'VENDOR' in f_val:
                                        raw_fact = str(sheet.cell(row=r+i, column=c).value).strip().replace('"', '')
                                        if ':' in raw_fact: raw_fact = raw_fact.split(':', 1)[1].strip()
                                        parts = raw_fact.split('/')
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
                        st.warning("⚠️ 在 Program Sheet 中未偵測到 DPCI。請確認檔案內容。")
                        st.stop()
                    
                    df_out = pd.DataFrame(parsed_items)

                    # ---------------------------------------------------------
                    # 步驟 C: 終極 VLOOKUP 補齊資料
                    # ---------------------------------------------------------
                    clean_main_dpci = df_out['DPCI'].astype(str).str.replace('-', '').str.strip()
                    
                    if cat_mapping: df_out['CATEGORY'] = clean_main_dpci.map(cat_mapping).fillna('')
                    else: df_out['CATEGORY'] = ""
                    
                    if fact_mapping:
                        df_out['Factory Name'] = df_out.apply(lambda row: fact_mapping.get(str(row['DPCI']).replace('-', ''), "") if not row['Factory Name'] else row['Factory Name'], axis=1)
                        
                    if qty_mapping:
                        df_out['QTY'] = df_out.apply(lambda row: qty_mapping.get(str(row['DPCI']).replace('-', ''), "") if not row['QTY'] else row['QTY'], axis=1)

                    df_out['Factory Name'] = df_out['Factory Name'].replace('', 'Unknown Factory')

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
                    # 步驟 D: XlsxWriter 進階排版
                    # ---------------------------------------------------------
                    output = io.BytesIO()
                    img_insert_count = 0
                    
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        workbook = writer.book
                        worksheet = workbook.add_worksheet('Program Items')
                        
                        format_cache = {}
                        def get_fmt(t=1, b=1, l=1, r=1, align=None):
                            key = (t, b, l, r, align)
                            if key not in format_cache:
                                props = {'font_name': 'Arial', 'valign': 'vcenter', 'text_wrap': True,
                                         'top': t, 'bottom': b, 'left': l, 'right': r}
                                if align: props['align'] = align
                                format_cache[key] = workbook.add_format(props)
                            return format_cache[key]

                        # 1. L1 總數
                        fmt_total = workbook.add_format({'font_name': 'Arial', 'bold': True, 'font_color': '#800080', 'font_size': 12, 'align': 'center', 'valign': 'vcenter'})
                        worksheet.write(0, 11, len(df_out), fmt_total)
                        
                        # 2. 標題列
                        fmt_header = workbook.add_format({'bold': True, 'bg_color': '#FFD966', 'border': 1, 'font_name': 'Arial', 'valign': 'vcenter', 'text_wrap': True}) 
                        for c, col_name in enumerate(target_columns):
                            worksheet.write(1, c, col_name, fmt_header)
                            worksheet.set_column(c, c, 15) 
                            
                        worksheet.set_column(3, 3, 16) # PHOTO
                        worksheet.set_column(9, 9, 25) # Factory Name
                        worksheet.set_column(10, 10, 15) # Factory ID
                        
                        # 3. 群組化
                        factories = df_out['Factory Name'].tolist()
                        groups = []
                        if factories:
                            start_idx = 0
                            for i in range(1, len(factories)):
                                if factories[i] != factories[start_idx]:
                                    groups.append((start_idx, i - 1, factories[start_idx]))
                                    start_idx = i
                            groups.append((start_idx, len(factories) - 1, factories[start_idx]))

                        # 4. 寫入資料
                        excel_row_offset = 2 
                        for s_idx, e_idx, f_name in groups:
                            f_id = str(df_out.iloc[s_idx]['Factory ID']).strip()
                            
                            for i in range(s_idx, e_idx + 1):
                                excel_row = i + excel_row_offset
                                worksheet.set_row(excel_row, 80) 
                                
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
                                
                                for c, col_name in enumerate(target_columns):
                                    if c in [9, 10]: continue 
                                    
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
                                            
                                    elif c == 11: # Total SKU
                                        if i == s_idx: worksheet.write(excel_row, c, e_idx - s_idx + 1, fmt)
                                        else: worksheet.write(excel_row, c, "", fmt)
                                    else:
                                        worksheet.write(excel_row, c, df_out.iloc[i][col_name], fmt)
                                        
                            # 垂直合併與外框
                            fmt_fact = get_fmt(t=2, b=2, l=1, r=1, align='center')
                            if s_idx == e_idx:
                                worksheet.write(s_idx + excel_row_offset, 9, f_name, fmt_fact)
                                worksheet.write(s_idx + excel_row_offset, 10, f_id, fmt_fact)
                            else:
                                worksheet.merge_range(s_idx + excel_row_offset, 9, e_idx + excel_row_offset, 9, f_name, fmt_fact)
                                worksheet.merge_range(s_idx + excel_row_offset, 10, e_idx + excel_row_offset, 10, f_id, fmt_fact)

                    processed_data = output.getvalue()
                    
                    st.success(f"✅ 處理完成！共解析出 **{len(df_out)}** 筆商品，並成功置入 **{img_insert_count}** 張圖片。")
                    
                    st.download_button(
                        label="📥 下載 Item Tracking Grid.xlsx",
                        data=processed_data,
                        file_name="Item_Tracking_Grid.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                except Exception as e:
                    st.error(f"❌ 處理檔案時發生錯誤: {e}")

else:
    st.info("💡 提示：請在上方直接拖曳上傳您的專案 Excel/CSV 檔案 (包含 Program Sheet 與 Data)。")
