import streamlit as st
import pandas as pd
import io
import os
import tempfile
import openpyxl
from openpyxl_image_loader import SheetImageLoader
from openpyxl.utils import get_column_letter
from PIL import Image

# ==========================================
# 1. 頁面基本設定與標題
# ==========================================
st.set_page_config(page_title="Program Items Generator", layout="wide")
st.title("🎃 萬聖節專案 Program Items 自動生成工具")
st.markdown("請上傳專案檔案 (可一次選取/拖曳多個檔案)。系統將自動解析 Master Sheet 的卡片資料、**包含圖片**，並結合 Data 表的 Subclass Name。")

# ==========================================
# 2. 檔案上傳區塊
# ==========================================
uploaded_files = st.file_uploader("📁 請上傳 Excel / CSV 檔案 (可同時上傳 Master Sheet 與 Data)", 
                                  type=["xlsx", "xls", "csv"], 
                                  accept_multiple_files=True)

st.divider() 

# ==========================================
# 3. 核心處理邏輯
# ==========================================
if uploaded_files:
    sheet_options = []
    df_dict = {}
    file_bytes_dict = {} # 儲存二進位檔案供 openpyxl 讀取
    
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
                    
    # 自動預選
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

    if st.button("✨ 生成 Program Items", type="primary"):
        with st.spinner("解析卡片與萃取圖片中，請稍候 (這可能需要幾十秒)..."):
            try:
                # ---------------------------------------------------------
                # 步驟 A: 建立 CATEGORY 對照字典
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
                            clean_dpci = df_data[dpci_col].astype(str).str.replace("-", "").str.strip()
                            clean_dpci = clean_dpci.apply(lambda x: x[:-2] if x.endswith('.0') else x)
                            cat_mapping = dict(zip(clean_dpci, df_data[subclass_col]))

                # ---------------------------------------------------------
                # 步驟 B: 使用 openpyxl 精準解析 Master Sheet 與圖片
                # ---------------------------------------------------------
                parsed_items = []
                
                # 取得檔名與工作表名
                master_file_name = selected_master.split("]")[0][1:] 
                master_sheet_name = selected_master.split("] ")[1]
                
                wb = openpyxl.load_workbook(io.BytesIO(file_bytes_dict[master_file_name]), data_only=True)
                sheet = wb[master_sheet_name]
                
                # 啟動圖片載入器
                try:
                    image_loader = SheetImageLoader(sheet)
                except Exception as e:
                    image_loader = None
                    st.warning("⚠️ 此檔案無法初始化圖片載入器，將略過圖片抓取。")
                
                for r in range(1, sheet.max_row + 1):
                    for c in range(1, sheet.max_column + 1):
                        val = str(sheet.cell(row=r, column=c).value).strip().upper()
                        
                        # 找到卡片起點
                        if val == 'DPCI:':
                            # 1. 抓取 DPCI
                            dpci = str(sheet.cell(row=r, column=c+1).value).strip()
                            dpci = dpci if dpci.lower() != 'none' else ""
                            
                            # 2. 抓取圖片 (在 DPCI 的正上方一格)
                            img_obj = None
                            if image_loader and r > 1:
                                img_cell = f"{get_column_letter(c)}{r - 1}"
                                try:
                                    if image_loader.image_in(img_cell):
                                        img_obj = image_loader.get(img_cell)
                                except Exception:
                                    pass # 忽略沒有圖片的錯誤
                                    
                            # 3. 往下掃描找 Description:
                            desc = ""
                            for i in range(1, 15):
                                if r+i <= sheet.max_row:
                                    if str(sheet.cell(row=r+i, column=c).value).strip().upper() == 'DESCRIPTION:':
                                        desc = str(sheet.cell(row=r+i, column=c+1).value).strip()
                                        desc = desc if desc.lower() != 'none' else ""
                                        break
                                        
                            # 4. 往下掃描找 QTY:
                            qty = ""
                            for i in range(1, 15):
                                if r+i > sheet.max_row: break
                                found_qty = False
                                for j in range(c, c+6):
                                    if j <= sheet.max_column:
                                        if str(sheet.cell(row=r+i, column=j).value).strip().upper() == 'QTY:':
                                            qty = str(sheet.cell(row=r+i, column=j+1).value).strip()
                                            qty = qty if qty.lower() != 'none' else ""
                                            if qty.endswith('.0'): qty = qty[:-2] # 清理小數點
                                            found_qty = True
                                            break
                                if found_qty: break
                                
                            # 5. 往下掃描找 Factory
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
                                'PHOTO_OBJ': img_obj, # 暫存 PIL 物件
                                'Factory Name': factory_name,
                                'Factory ID': factory_id,
                                'QTY': qty
                            })

                # ---------------------------------------------------------
                # 步驟 C: 重組 22 項最終欄位並匯出 Excel
                # ---------------------------------------------------------
                if not parsed_items:
                    st.warning("⚠️ 在 Master Sheet 中未偵測到任何含有 'DPCI:' 的卡片。")
                    st.stop()
                
                # 移除 PIL 圖片物件轉換為純文字 DataFrame
                df_out = pd.DataFrame([{k: v for k, v in item.items() if k != 'PHOTO_OBJ'} for item in parsed_items])

                # VLOOKUP CATEGORY
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

                # 補齊欄位並排序
                for col in target_columns:
                    if col not in df_out.columns:
                        df_out[col] = ""
                df_out = df_out[target_columns]

                # 格式設定與寫入圖片
                output = io.BytesIO()
                with tempfile.TemporaryDirectory() as temp_dir:
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
                            
                        # 設定 PHOTO 欄位特別寬度 (Index 3)
                        col_idx_photo = 3
                        worksheet.set_column(col_idx_photo, col_idx_photo, 16)
                        
                        # 迴圈插入每一列的圖片
                        for row_num, item in enumerate(parsed_items):
                            excel_row = row_num + 1
                            worksheet.set_row(excel_row, 80) # 放大行高以容納圖片
                            
                            img = item.get('PHOTO_OBJ')
                            if img:
                                safe_name = "".join(x for x in str(item['DPCI']) if x.isalnum() or x in "-_")
                                if not safe_name: safe_name = f"img_{row_num}"
                                img_path = os.path.join(temp_dir, f"{safe_name}.png")
                                
                                try:
                                    # 將圖片等比例縮小以完美塞入儲存格
                                    img.thumbnail((100, 100))
                                    img.save(img_path, format="PNG")
                                    # 插入圖片並給予 5px 的偏移讓畫面好看
                                    worksheet.insert_image(excel_row, col_idx_photo, img_path, {'x_offset': 5, 'y_offset': 5})
                                except Exception:
                                    pass

                processed_data = output.getvalue()
                
                st.success(f"✅ 處理完成！已成功萃取 **{len(parsed_items)}** 筆資料與對應圖片。")
                
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
