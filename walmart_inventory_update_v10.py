import streamlit as st
import openpyxl
from openpyxl.styles import Alignment, Font
import io
import re
import warnings

# 忽略警告
warnings.filterwarnings('ignore')

# ==========================================
# 核心逻辑函数 (基于 v11 修复版)
# ==========================================

def find_sheet_name(sheets, keywords):
    """根据关键词查找sheet名称"""
    for sheet_name in sheets:
        if any(keyword in sheet_name for keyword in keywords):
            return sheet_name
    return None

def get_cell_value(cell):
    """获取单元格值，处理空值"""
    if cell.value is None:
        return 0
    val = str(cell.value).strip()
    if val == '' or val.lower() in ('nan', '#n/a', '#na', 'none', ''):
        return 0
    try:
        if val.startswith('='):
            return 0
        return float(val)
    except (ValueError, TypeError):
        return val

def get_numeric_value(cell):
    """获取数值"""
    if cell is None or cell.value is None:
        return 0
    val = str(cell.value).strip()
    if val == '' or val.lower() in ('nan', '#n/a', '#na', 'none', ''):
        return 0
    try:
        if val.startswith('='):
            return 0
        return float(val)
    except (ValueError, TypeError):
        return 0

def load_product_reference_from_obj(product_file_obj):
    """从上传的产品资料表对象中加载sku集合和映射"""
    sku_set = set()
    sku_to_name = {}
    
    if product_file_obj is None:
        return sku_set, sku_to_name

    try:
        wb_product = openpyxl.load_workbook(product_file_obj, read_only=True)
        ws = wb_product[wb_product.sheetnames[0]]

        sku_col_idx = None
        name_col_idx = None
        
        # 智能查找列头
        for col_idx in range(1, 10):
            header = ws.cell(row=1, column=col_idx).value
            if header:
                header_str = str(header).strip()
                if header_str.lower() == 'sku' and sku_col_idx is None:
                    sku_col_idx = col_idx
                elif header_str in ['品名', '名称', 'name', 'Name', '品名英'] and name_col_idx is None:
                    name_col_idx = col_idx

        if sku_col_idx:
            row_idx = 2
            while True:
                try:
                    sku_cell = ws.cell(row=row_idx, column=sku_col_idx)
                    if not sku_cell: break
                    sku_val = str(sku_cell.value).strip() if sku_cell.value else ''
                except:
                    break

                if not sku_val:
                    break

                sku_set.add(sku_val)
                if name_col_idx:
                    name_cell = ws.cell(row=row_idx, column=name_col_idx)
                    name_val = str(name_cell.value).strip() if name_cell.value else ''
                    if name_val:
                        sku_to_name[sku_val] = name_val
                row_idx += 1
        wb_product.close()
        return sku_set, sku_to_name
    except Exception as e:
        st.error(f"读取产品资料表失败: {e}")
        return set(), {}

def extract_sku_smart(msku, sku_set):
    """智能SKU提取逻辑"""
    if not msku: return '', False
    if not sku_set:
        parts = msku.split('-')
        if len(parts) >= 2: return parts[1], False
        elif len(parts) >= 1: return parts[0], False
        return msku, False

    parts = msku.split('-')
    parts = [p.strip() for p in parts if p.strip()]

    # 1. 精确匹配
    for part in parts:
        if part in sku_set: return part, True

    # 2. 去除特殊字符匹配
    cleaned_parts = []
    for part in parts:
        cleaned = part.replace('"', '').replace("'", '').replace(' ', '')
        if cleaned: cleaned_parts.append(cleaned)
    for part in cleaned_parts:
        if part in sku_set: return part, True

    # 3. 模糊/子串匹配
    for part in cleaned_parts:
        if len(part) >= 4:
            for sku in sku_set:
                if part in sku or sku in part:
                    if len(sku) > 0 and (len(part) / len(sku) >= 0.5):
                        return sku, True
    
    # 4. 字母数字组合回退
    for part in parts:
        if re.search(r'[a-zA-Z]', part) and re.search(r'\d', part):
            cleaned = part.replace('"', '').replace("'", '').replace(' ', '')
            if cleaned in sku_set: return cleaned, True

    # 5. 最长部分回退
    if cleaned_parts:
        longest = max(cleaned_parts, key=len)
        return longest, False

    return '', False

def process_inventory(inventory_file, product_file):
    """主处理函数"""
    
    # 1. 加载产品资料
    sku_set, sku_to_name = load_product_reference_from_obj(product_file)
    if product_file:
        st.info(f"已加载产品资料：包含 {len(sku_set)} 个SKU，{len(sku_to_name)} 个品名映射")
    else:
        st.warning("未上传产品资料表，将使用简易模式提取SKU且无法自动填充品名。")

    # 2. 加载主文件
    wb = openpyxl.load_workbook(inventory_file)
    sheets = wb.sheetnames
    
    # 查找 Sheet
    inventory_sheet_name = sheets[1] if len(sheets) > 1 else None
    sz_stock_sheet_name = find_sheet_name(sheets, ['深圳仓', '深圳', '仓库'])
    wfs_stock_sheet_name = find_sheet_name(sheets, ['WFS库存', 'WFS'])
    sales_sheet_name = sheets[4] if len(sheets) > 4 else None
    po_sheet_name = find_sheet_name(sheets, ['采购订单', '采购', '在途'])

    if not all([inventory_sheet_name, wfs_stock_sheet_name, sales_sheet_name]):
        st.error("无法识别必要的Sheet，请检查文件格式（需包含WFS库存、销量明细、库存明细表）。")
        return None

    # === 第0步：记录原有记录 (历史保护) ===
    inventory_sheet = wb[inventory_sheet_name]
    existing_keys = set()
    
    # 扫描确定原有记录范围
    original_max_row = inventory_sheet.max_row
    st.write(f"原有记录保护范围：前 {original_max_row} 行 (含表头)")
    
    # 预加载现有Key以避免重复
    for r in range(3, original_max_row + 1):
        s_val = inventory_sheet.cell(r, 1).value
        m_val = inventory_sheet.cell(r, 2).value
        store = str(s_val).strip() if s_val else ''
        msku = str(m_val).strip() if m_val else ''
        if store or msku:
            existing_keys.add(f"{store}{msku}")

    # === 第1步：处理 WFS 库存 ===
    wfs_sheet = wb[wfs_stock_sheet_name]
    wfs_dict = {}
    
    # 动态映射列 (严格匹配逻辑修正)
    wfs_header = [c.value for c in wfs_sheet[1]]
    wfs_map = {}
    for idx, col_name in enumerate(wfs_header, 1):
        if not col_name: continue
        c = str(col_name).strip()
        if '仓库' in c: wfs_map['仓库'] = idx
        elif 'msku' in c: wfs_map['msku'] = idx
        elif c == 'GTIN码': wfs_map['GTIN码'] = idx
        elif '平台' in c and '商品' in c and 'ID' in c: wfs_map['平台商品ID'] = idx # 严格区分ID
        elif '品名' in c and 'ID' not in c: wfs_map['品名'] = idx
        elif 'sku' == c: wfs_map['sku'] = idx
        elif '商品状态' in c: wfs_map['商品状态'] = idx
        # 严格匹配数值列，必须包含"数量"
        elif 'WFS可售' in c and '新' in c and '数量' in c: wfs_map['可售'] = idx
        elif '无法入库' in c and '数量' in c: wfs_map['无法入库'] = idx
        elif '标发在途' in c and '数量' in c: wfs_map['标发'] = idx 
    
    if '标发' not in wfs_map:
        st.warning("⚠️ 警告：在WFS库存表中未找到'标发在途'且包含'数量'的列，相关数据可能为0。")

    for row in range(2, wfs_sheet.max_row + 1):
        wh = str(wfs_sheet.cell(row, wfs_map.get('仓库', 1)).value or '').strip()
        msku = str(wfs_sheet.cell(row, wfs_map.get('msku', 2)).value or '').strip()
        if wh and msku:
            key = f"{wh}{msku}"
            wfs_dict[key] = {
                '仓库': wh, 'msku': msku,
                'GTIN码': str(wfs_sheet.cell(row, wfs_map.get('GTIN码', 10)).value or ''),
                '品名': wfs_sheet.cell(row, wfs_map.get('品名', 12)).value,
                'sku': wfs_sheet.cell(row, wfs_map.get('sku', 13)).value,
                '商品状态': wfs_sheet.cell(row, wfs_map.get('商品状态', 14)).value,
                # 使用安全获取，如果没有映射到列则默认为0
                '可售': get_numeric_value(wfs_sheet.cell(row, wfs_map.get('可售', 999))) if '可售' in wfs_map else 0,
                '无法入库': get_numeric_value(wfs_sheet.cell(row, wfs_map.get('无法入库', 999))) if '无法入库' in wfs_map else 0,
                '标发': get_numeric_value(wfs_sheet.cell(row, wfs_map.get('标发', 999))) if '标发' in wfs_map else 0
            }

    # === 第2步：处理销量明细 (v11: 精确店铺匹配) ===
    sales_sheet = wb[sales_sheet_name]
    sales_dict = {}
    sales_header = [c.value for c in sales_sheet[1]]
    sales_map = {}
    for idx, col_name in enumerate(sales_header, 1):
        if not col_name: continue
        c = str(col_name).strip()
        if c == 'MSKU': sales_map['MSKU'] = idx
        elif c == 'SKU': sales_map['SKU'] = idx
        elif c == '店铺': sales_map['店铺'] = idx
        elif c == '品名': sales_map['品名'] = idx
        elif c == '小计': sales_map['小计'] = idx

    for row in range(2, sales_sheet.max_row + 1):
        msku = str(sales_sheet.cell(row, sales_map.get('MSKU', 4)).value or '').strip()
        store = str(sales_sheet.cell(row, sales_map.get('店铺', 3)).value or '').strip()
        
        if not store and '-' in msku: store = msku.split('-')[0]
             
        if msku:
            key = f"{store}{msku}"
            sku_val = str(sales_sheet.cell(row, sales_map.get('SKU', 6)).value or '').strip()
            if not sku_val: sku_val, _ = extract_sku_smart(msku, sku_set)
                
            sales_dict[key] = {
                '店铺': store, 'msku': msku,
                '销量': get_numeric_value(sales_sheet.cell(row, sales_map.get('小计', 13))),
                'SKU': sku_val,
                '品名': sales_sheet.cell(row, sales_map.get('品名', 7)).value
            }

    # === 第3步 & 4步：深圳仓 & 采购在途 ===
    sz_stock_dict = {}
    if sz_stock_sheet_name:
        sz_sheet = wb[sz_stock_sheet_name]
        sz_qty_col = 10 
        for col in range(1, 20):
            val = str(sz_sheet.cell(1, col).value or '')
            if '可用' in val: sz_qty_col = col; break
        
        for row in range(2, sz_sheet.max_row + 1):
            sku = str(sz_sheet.cell(row, 1).value or '').strip()
            if sku:
                qty = get_numeric_value(sz_sheet.cell(row, sz_qty_col))
                sz_stock_dict[sku] = sz_stock_dict.get(sku, 0) + qty

    po_dict = {}
    if po_sheet_name:
        po_sheet = wb[po_sheet_name]
        po_sku_col = 7; po_qty_col = 19
        for col in range(1, 40):
            val = str(po_sheet.cell(1, col).value or '')
            if 'SKU' == val.upper(): po_sku_col = col
            if '未入库' in val: po_qty_col = col
            
        for row in range(2, po_sheet.max_row + 1):
            sku = str(po_sheet.cell(row, po_sku_col).value or '').strip()
            if sku:
                qty = get_numeric_value(po_sheet.cell(row, po_qty_col))
                po_dict[sku] = po_dict.get(sku, 0) + qty

    # === 第5步：更新与新增 ===
    inv_header_row = 2
    inv_map = {}
    for col in range(1, 100):
        val = inventory_sheet.cell(inv_header_row, col).value
        if val:
            c = str(val).strip()
            if c == '店铺': inv_map['店铺'] = col
            elif c == 'msku': inv_map['msku'] = col
            elif c == '店铺&MSKU': inv_map['店铺&MSKU'] = col
            elif c == 'GTIN码': inv_map['GTIN'] = col
            elif c == '品名': inv_map['品名'] = col
            elif c == 'sku': inv_map['sku'] = col
            elif c == '商品状态': inv_map['状态'] = col
            # 同样严格匹配目标表列名
            elif 'WFS可售' in c and '数量' in c: inv_map['WFS可售'] = col
            elif '无法入库' in c and '数量' in c: inv_map['无法入库'] = col
            elif '标发' in c and '数量' in c: inv_map['标发'] = col
            elif '深圳仓' in c: inv_map['深圳仓'] = col
            elif '采购' in c: inv_map['采购'] = col
            elif '总库存' in c: inv_map['总库存'] = col
            elif '总周转' in c: inv_map['总周转'] = col
            elif c == sales_sheet_name: inv_map['销量'] = col
    
    if '销量' not in inv_map:
        new_col = inventory_sheet.max_column + 1
        inventory_sheet.cell(inv_header_row, new_col, value=sales_sheet_name)
        inv_map['销量'] = new_col

    # 5.1 更新现有记录
    for r in range(3, original_max_row + 1):
        s_val = str(inventory_sheet.cell(r, inv_map.get('店铺', 1)).value or '').strip()
        m_val = str(inventory_sheet.cell(r, inv_map.get('msku', 2)).value or '').strip()
        sku_val = str(inventory_sheet.cell(r, inv_map.get('sku', 8)).value or '').strip()
        
        if not s_val and not m_val: continue

        key = f"{s_val}{m_val}"
        
        if key in wfs_dict:
            d = wfs_dict[key]
            if 'GTIN' in inv_map: inventory_sheet.cell(r, inv_map['GTIN'], d['GTIN码'])
            if '品名' in inv_map and d['品名']: inventory_sheet.cell(r, inv_map['品名'], d['品名'])
            if 'sku' in inv_map and d['sku']: inventory_sheet.cell(r, inv_map['sku'], d['sku'])
            if '状态' in inv_map: inventory_sheet.cell(r, inv_map['状态'], d['商品状态'])
            if 'WFS可售' in inv_map: inventory_sheet.cell(r, inv_map['WFS可售'], d['可售'])
            if '无法入库' in inv_map: inventory_sheet.cell(r, inv_map['无法入库'], d['无法入库'])
            if '标发' in inv_map: inventory_sheet.cell(r, inv_map['标发'], d['标发'])
            if d['sku']: sku_val = str(d['sku']).strip()
        
        if key in sales_dict:
            inventory_sheet.cell(r, inv_map['销量'], sales_dict[key]['销量'])
        
        if sku_val:
            if sku_val in sz_stock_dict and '深圳仓' in inv_map:
                inventory_sheet.cell(r, inv_map['深圳仓'], sz_stock_dict[sku_val])
            if sku_val in po_dict and '采购' in inv_map:
                inventory_sheet.cell(r, inv_map['采购'], po_dict[sku_val])

    # 5.2 添加新记录
    all_keys = set(wfs_dict.keys()) | set(sales_dict.keys())
    new_rows_data = []
    
    for key in all_keys:
        if key in existing_keys: continue
        
        row_data = {
            '店铺': '', 'msku': '', '店铺&MSKU': key, 'sku': '', '品名': '',
            'WFS可售': 0, '无法入库': 0, '标发': 0, '销量': 0
        }
        
        if key in wfs_dict:
            src = wfs_dict[key]
            row_data.update({
                '店铺': src['仓库'], 'msku': src['msku'],
                'GTIN': src['GTIN码'], '品名': src['品名'],
                'sku': src['sku'], '状态': src['商品状态'],
                'WFS可售': src['可售'], '无法入库': src['无法入库'], '标发': src['标发']
            })
        elif key in sales_dict:
            src = sales_dict[key]
            row_data.update({
                '店铺': src['店铺'], 'msku': src['msku'],
                '销量': src['销量'], 'sku': src['SKU'], '品名': src['品名']
            })
            
        current_sku = str(row_data.get('sku', '') or '').strip()
        if not current_sku:
            extracted_sku, found = extract_sku_smart(row_data['msku'], sku_set)
            if extracted_sku:
                current_sku = extracted_sku
                row_data['sku'] = current_sku
        
        if not row_data.get('品名') and current_sku in sku_to_name:
            row_data['品名'] = sku_to_name[current_sku]

        new_rows_data.append(row_data)

    current_row = original_max_row + 1
    
    for data in new_rows_data:
        if '店铺' in inv_map: inventory_sheet.cell(current_row, inv_map['店铺'], data['店铺'])
        if 'msku' in inv_map: inventory_sheet.cell(current_row, inv_map['msku'], data['msku'])
        if '店铺&MSKU' in inv_map: inventory_sheet.cell(current_row, inv_map['店铺&MSKU'], data['店铺&MSKU'])
        if 'GTIN' in inv_map and 'GTIN' in data: inventory_sheet.cell(current_row, inv_map['GTIN'], data['GTIN'])
        if 'sku' in inv_map: inventory_sheet.cell(current_row, inv_map['sku'], data['sku'])
        if '品名' in inv_map: inventory_sheet.cell(current_row, inv_map['品名'], data['品名'])
        if '状态' in inv_map and '状态' in data: inventory_sheet.cell(current_row, inv_map['状态'], data['状态'])
        
        if 'WFS可售' in inv_map: inventory_sheet.cell(current_row, inv_map['WFS可售'], data['WFS可售'])
        if '无法入库' in inv_map: inventory_sheet.cell(current_row, inv_map['无法入库'], data['无法入库'])
        if '标发' in inv_map: inventory_sheet.cell(current_row, inv_map['标发'], data['标发'])
        if '销量' in inv_map: inventory_sheet.cell(current_row, inv_map['销量'], data['销量'])
        
        sku = str(data.get('sku', '')).strip()
        if sku:
            if sku in sz_stock_dict and '深圳仓' in inv_map:
                inventory_sheet.cell(current_row, inv_map['深圳仓'], sz_stock_dict[sku])
            if sku in po_dict and '采购' in inv_map:
                inventory_sheet.cell(current_row, inv_map['采购'], po_dict[sku])
                
        current_row += 1

    # === 第6步：计算公式 & 清理 ===
    for r in range(3, current_row):
        wfs = get_numeric_value(inventory_sheet.cell(r, inv_map.get('WFS可售')))
        unable = get_numeric_value(inventory_sheet.cell(r, inv_map.get('无法入库')))
        transit = get_numeric_value(inventory_sheet.cell(r, inv_map.get('标发')))
        sz = get_numeric_value(inventory_sheet.cell(r, inv_map.get('深圳仓')))
        sales = get_numeric_value(inventory_sheet.cell(r, inv_map.get('销量')))
        
        total = wfs + unable + transit + sz
        if '总库存' in inv_map:
            inventory_sheet.cell(r, inv_map['总库存'], total)
            
        if '总周转' in inv_map:
            if sales > 0:
                turnover = round((wfs + transit + sz) / sales * 30, 2)
                inventory_sheet.cell(r, inv_map['总周转'], turnover)
            else:
                inventory_sheet.cell(r, inv_map['总周转'], "")

    rows_to_delete = []
    for r in range(original_max_row + 1, current_row):
        is_zero = True
        vals = [
            get_numeric_value(inventory_sheet.cell(r, inv_map.get('WFS可售'))),
            get_numeric_value(inventory_sheet.cell(r, inv_map.get('无法入库'))),
            get_numeric_value(inventory_sheet.cell(r, inv_map.get('标发'))),
            get_numeric_value(inventory_sheet.cell(r, inv_map.get('深圳仓'))),
            get_numeric_value(inventory_sheet.cell(r, inv_map.get('采购'))),
            get_numeric_value(inventory_sheet.cell(r, inv_map.get('销量')))
        ]
        if any(v != 0 for v in vals): is_zero = False
            
        store_val = str(inventory_sheet.cell(r, inv_map.get('店铺', 1)).value or '').strip()
        
        if is_zero and not store_val: rows_to_delete.append(r)
        elif is_zero: rows_to_delete.append(r)

    for r in sorted(rows_to_delete, reverse=True):
        inventory_sheet.delete_rows(r, 1)

    st.success(f"处理完成！原有记录 {original_max_row} 行，新增记录 {len(new_rows_data) - len(rows_to_delete)} 条。")
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ==========================================
# Streamlit UI
# ==========================================

st.set_page_config(page_title="沃尔玛库存更新工具 v11", layout="wide")

st.title("🛒 沃尔玛呆滞库存更新工具 (v11 修复版)")
st.markdown("""
**功能说明：**
1. 自动合并 WFS库存、销量明细、深圳仓库存、采购在途数据。
2. **修复说明**：严格匹配“标发在途”与“数量”列，防止误读取 ID 或其他无关列。
3. **隐私安全**：数据仅在内存中处理，刷新页面即销毁。
""")

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. 上传库存数据表 (必选)")
    inventory_file = st.file_uploader("上传 .xlsx 文件", type=['xlsx'], key="inv")

with col2:
    st.subheader("2. 上传产品资料表 (推荐)")
    st.markdown("用于智能匹配 SKU 和自动填充品名")
    product_file = st.file_uploader("上传 .xlsx 文件", type=['xlsx'], key="prod")

if inventory_file:
    if st.button("🚀 开始处理", type="primary"):
        with st.spinner("正在分析数据，请稍候..."):
            try:
                processed_data = process_inventory(inventory_file, product_file)
                
                if processed_data:
                    st.success("✅ 处理成功！请点击下方按钮下载。")
                    st.download_button(
                        label="📥 下载更新后的库存表",
                        data=processed_data,
                        file_name=f"updated_{inventory_file.name}",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"处理过程中发生错误: {str(e)}")
                st.exception(e)
