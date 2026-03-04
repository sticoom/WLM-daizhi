import streamlit as st
import openpyxl
from openpyxl.styles import Alignment, Font
import io
import re
import warnings

# 忽略警告
warnings.filterwarnings('ignore')

# ==========================================
# 核心逻辑函数 (基于 v12 深度修复版)
# ==========================================

def clean_header(header_value):
    """清理表头：转字符串、去空格、统一括号、转小写"""
    if not header_value:
        return ""
    # 转字符串并去空格
    s = str(header_value).strip().lower()
    # 统一括号：将全角括号转为半角
    s = s.replace('（', '(').replace('）', ')')
    # 去除多余空格
    s = s.replace(' ', '')
    return s

def find_col_index_smart(sheet, keywords_must, keywords_exclude=None, header_row=1):
    """
    智能查找列索引
    :param keywords_must: 必须包含的关键词列表 (AND关系)
    :param keywords_exclude: 不能包含的关键词列表
    :return: 列索引 (1-based) 或 None
    """
    if keywords_exclude is None:
        keywords_exclude = []
        
    best_col = None
    best_score = 0 # 用于区分"标发在途"和"标发在途(数量)"，优先匹配更长的/带'数量'的
    
    # 遍历前50列
    for col in range(1, 51):
        val = sheet.cell(row=header_row, column=col).value
        if not val: continue
        
        header_clean = clean_header(val)
        
        # 1. 检查必须包含的词
        if not all(k.lower() in header_clean for k in keywords_must):
            continue
            
        # 2. 检查必须排除的词
        if any(k.lower() in header_clean for k in keywords_exclude):
            continue
            
        # 3. 评分机制：如果包含"数量"或"qty"，优先级更高
        current_score = 1
        if '数量' in header_clean or 'qty' in header_clean or '件数' in header_clean:
            current_score += 2
            
        # 如果是第一次找到，或者当前列分数更高，则更新
        if best_col is None or current_score > best_score:
            best_col = col
            best_score = current_score
            
    return best_col

def get_numeric_value(cell):
    """获取数值，强制转float，异常返回0"""
    if cell is None or cell.value is None:
        return 0
    val = str(cell.value).strip()
    if val == '' or val.lower() in ('nan', '#n/a', '#na', 'none', ''):
        return 0
    try:
        if val.startswith('='): return 0
        # 去除千分位逗号等
        val = val.replace(',', '')
        return float(val)
    except (ValueError, TypeError):
        return 0

def load_product_reference_from_obj(product_file_obj):
    """加载产品资料表"""
    sku_set = set()
    sku_to_name = {}
    if product_file_obj is None: return sku_set, sku_to_name

    try:
        wb_product = openpyxl.load_workbook(product_file_obj, read_only=True)
        ws = wb_product[wb_product.sheetnames[0]]

        sku_col = find_col_index_smart(ws, ['sku'])
        name_col = find_col_index_smart(ws, ['品名']) or find_col_index_smart(ws, ['名称']) or find_col_index_smart(ws, ['name'])

        if sku_col:
            row_idx = 2
            while True:
                try:
                    sku_cell = ws.cell(row=row_idx, column=sku_col)
                    if not sku_cell: break
                    sku_val = str(sku_cell.value).strip() if sku_cell.value else ''
                except: break

                if not sku_val: break
                sku_set.add(sku_val)
                if name_col:
                    name_cell = ws.cell(row=row_idx, column=name_col)
                    name_val = str(name_cell.value).strip() if name_cell.value else ''
                    if name_val: sku_to_name[sku_val] = name_val
                row_idx += 1
        wb_product.close()
        return sku_set, sku_to_name
    except Exception as e:
        st.error(f"读取产品资料表失败: {e}")
        return set(), {}

def extract_sku_smart(msku, sku_set):
    """智能SKU提取"""
    if not msku: return '', False
    if not sku_set:
        parts = msku.split('-')
        return (parts[1] if len(parts)>=2 else parts[0]), False

    parts = msku.split('-')
    parts = [p.strip() for p in parts if p.strip()]

    # 1. 精确匹配
    for part in parts:
        if part in sku_set: return part, True
    
    # 2. 去符号匹配
    cleaned_parts = [p.replace('"', '').replace("'", '').replace(' ', '') for p in parts]
    for part in cleaned_parts:
        if part in sku_set: return part, True

    # 3. 模糊匹配 (SKU通常由字母数字组成且长度>4)
    for part in cleaned_parts:
        if len(part) >= 4 and re.search(r'\d', part) and re.search(r'[a-zA-Z]', part):
            if part in sku_set: return part, True
            # 子串检查
            for sku in sku_set:
                if part in sku or sku in part:
                    if len(sku) > 0 and len(part)/len(sku) >= 0.6:
                        return sku, True
    return '', False

def process_inventory(inventory_file, product_file):
    # 1. 加载资料
    sku_set, sku_to_name = load_product_reference_from_obj(product_file)
    if product_file:
        st.info(f"📚 已加载产品资料：{len(sku_set)} 个SKU")

    # 2. 加载主文件
    wb = openpyxl.load_workbook(inventory_file)
    sheets = wb.sheetnames
    
    # Sheet 查找逻辑
    def find_sheet(keywords):
        for s in sheets:
            if any(k in s for k in keywords): return s
        return None

    inventory_sheet_name = sheets[1] if len(sheets) > 1 else None
    sz_stock_sheet_name = find_sheet(['深圳仓', '深圳', '仓库'])
    wfs_stock_sheet_name = find_sheet(['WFS库存', 'WFS'])
    sales_sheet_name = sheets[4] if len(sheets) > 4 else None
    po_sheet_name = find_sheet(['采购订单', '采购', '在途'])

    if not all([inventory_sheet_name, wfs_stock_sheet_name, sales_sheet_name]):
        st.error("❌ 无法识别必要的Sheet。请确保文件包含：第2个Sheet为库存表，第5个为销量表，以及名为'WFS...'的Sheet。")
        return None

    # === 第0步：保护原有记录 ===
    inventory_sheet = wb[inventory_sheet_name]
    original_max_row = inventory_sheet.max_row
    st.write(f"🛡️ 原有记录保护范围：前 {original_max_row} 行")
    
    existing_keys = set()
    for r in range(3, original_max_row + 1):
        s = str(inventory_sheet.cell(r, 1).value or '').strip()
        m = str(inventory_sheet.cell(r, 2).value or '').strip()
        if s or m: existing_keys.add(f"{s}{m}")

    # === 第1步：WFS 库存 (增强匹配) ===
    wfs_sheet = wb[wfs_stock_sheet_name]
    wfs_dict = {}
    
    # 智能查找列 - 排除ID列，优先找带'数量'的列
    w_wh = find_col_index_smart(wfs_sheet, ['仓库']) or 1
    w_msku = find_col_index_smart(wfs_sheet, ['msku']) or 2
    w_gtin = find_col_index_smart(wfs_sheet, ['gtin'])
    w_sku = find_col_index_smart(wfs_sheet, ['sku'])
    w_name = find_col_index_smart(wfs_sheet, ['品名'], keywords_exclude=['id'])
    w_status = find_col_index_smart(wfs_sheet, ['商品状态'])
    
    # 关键数值列 (排除ID)
    w_avail = find_col_index_smart(wfs_sheet, ['wfs', '可售', '新'], keywords_exclude=['id', 'code'])
    w_unable = find_col_index_smart(wfs_sheet, ['无法', '入库'], keywords_exclude=['id', 'code'])
    w_transit = find_col_index_smart(wfs_sheet, ['标发', '在途'], keywords_exclude=['id', 'code', '货件'])
    
    # 调试信息：显示找到的列名
    if w_transit:
        col_name = wfs_sheet.cell(1, w_transit).value
        st.write(f"✅ WFS '标发在途' 匹配到列: {col_name} (第{w_transit}列)")
    else:
        st.warning("⚠️ 未找到 WFS '标发在途' 列，该项数据将为 0")

    for row in range(2, wfs_sheet.max_row + 1):
        wh = str(wfs_sheet.cell(row, w_wh).value or '').strip()
        msku = str(wfs_sheet.cell(row, w_msku).value or '').strip()
        if wh and msku:
            key = f"{wh}{msku}"
            wfs_dict[key] = {
                '仓库': wh, 'msku': msku,
                'GTIN码': str(wfs_sheet.cell(row, w_gtin).value or '') if w_gtin else '',
                'sku': wfs_sheet.cell(row, w_sku).value if w_sku else '',
                '品名': wfs_sheet.cell(row, w_name).value if w_name else '',
                '商品状态': wfs_sheet.cell(row, w_status).value if w_status else '',
                # 数值获取
                'WFS可售': get_numeric_value(wfs_sheet.cell(row, w_avail)) if w_avail else 0,
                '无法入库': get_numeric_value(wfs_sheet.cell(row, w_unable)) if w_unable else 0,
                '标发在途': get_numeric_value(wfs_sheet.cell(row, w_transit)) if w_transit else 0
            }

    # === 第2步：销量明细 (精准列名匹配) ===
    sales_sheet = wb[sales_sheet_name]
    sales_dict = {}
    
    s_msku = find_col_index_smart(sales_sheet, ['msku'])
    s_store = find_col_index_smart(sales_sheet, ['店铺'])
    s_subtotal = find_col_index_smart(sales_sheet, ['小计']) or find_col_index_smart(sales_sheet, ['sales'])
    s_sku = find_col_index_smart(sales_sheet, ['sku'])
    s_name = find_col_index_smart(sales_sheet, ['品名'])

    if s_subtotal:
        st.write(f"✅ 销量 '小计' 匹配到列: {sales_sheet.cell(1, s_subtotal).value}")

    for row in range(2, sales_sheet.max_row + 1):
        msku = str(sales_sheet.cell(row, s_msku).value or '').strip() if s_msku else ''
        store = str(sales_sheet.cell(row, s_store).value or '').strip() if s_store else ''
        
        if not store and '-' in msku: store = msku.split('-')[0]
             
        if msku:
            key = f"{store}{msku}"
            sku_val = str(sales_sheet.cell(row, s_sku).value or '').strip() if s_sku else ''
            if not sku_val: sku_val, _ = extract_sku_smart(msku, sku_set)
                
            sales_dict[key] = {
                '店铺': store, 'msku': msku,
                '销量': get_numeric_value(sales_sheet.cell(row, s_subtotal)) if s_subtotal else 0,
                'SKU': sku_val,
                '品名': sales_sheet.cell(row, s_name).value if s_name else ''
            }

    # === 第3 & 4步：深圳仓 & 采购 ===
    sz_stock_dict = {}
    if sz_stock_sheet_name:
        sz_sheet = wb[sz_stock_sheet_name]
        sz_sku = find_col_index_smart(sz_sheet, ['sku']) or 1
        sz_qty = find_col_index_smart(sz_sheet, ['可用']) or find_col_index_smart(sz_sheet, ['实际', '可用']) or 10
        for row in range(2, sz_sheet.max_row + 1):
            sku = str(sz_sheet.cell(row, sz_sku).value or '').strip()
            if sku:
                qty = get_numeric_value(sz_sheet.cell(row, sz_qty))
                sz_stock_dict[sku] = sz_stock_dict.get(sku, 0) + qty

    po_dict = {}
    if po_sheet_name:
        po_sheet = wb[po_sheet_name]
        po_sku = find_col_index_smart(po_sheet, ['sku']) or 7
        po_qty = find_col_index_smart(po_sheet, ['未入库']) or 19
        for row in range(2, po_sheet.max_row + 1):
            sku = str(po_sheet.cell(row, po_sku).value or '').strip()
            if sku:
                qty = get_numeric_value(po_sheet.cell(row, po_qty))
                po_dict[sku] = po_dict.get(sku, 0) + qty

    # === 第5步：更新与新增 (目标表映射) ===
    inv_header_row = 2
    # 目标表映射：严格匹配列名
    i_store = find_col_index_smart(inventory_sheet, ['店铺'], keywords_exclude=['msku'], header_row=2)
    i_msku = find_col_index_smart(inventory_sheet, ['msku'], keywords_exclude=['店铺'], header_row=2)
    i_store_msku = find_col_index_smart(inventory_sheet, ['店铺', 'msku'], header_row=2)
    i_gtin = find_col_index_smart(inventory_sheet, ['gtin'], header_row=2)
    i_name = find_col_index_smart(inventory_sheet, ['品名'], header_row=2)
    i_sku = find_col_index_smart(inventory_sheet, ['sku'], header_row=2)
    i_status = find_col_index_smart(inventory_sheet, ['状态'], header_row=2)
    
    # 关键目标数值列
    i_avail = find_col_index_smart(inventory_sheet, ['wfs', '可售', '新'], header_row=2)
    i_unable = find_col_index_smart(inventory_sheet, ['无法', '入库'], header_row=2)
    i_transit = find_col_index_smart(inventory_sheet, ['标发'], header_row=2)
    i_sz = find_col_index_smart(inventory_sheet, ['深圳仓'], header_row=2)
    i_po = find_col_index_smart(inventory_sheet, ['采购'], header_row=2)
    i_total = find_col_index_smart(inventory_sheet, ['总库存'], header_row=2)
    i_turnover = find_col_index_smart(inventory_sheet, ['总周转'], header_row=2)
    
    # 销量列 (动态)
    i_sales = find_col_index_smart(inventory_sheet, [clean_header(sales_sheet_name)], header_row=2)
    if not i_sales:
        i_sales = inventory_sheet.max_column + 1
        inventory_sheet.cell(row=2, column=i_sales, value=sales_sheet_name)
        st.write(f"➕ 新增销量列: {sales_sheet_name}")

    # --- 5.1 更新现有 ---
    for r in range(3, original_max_row + 1):
        s = str(inventory_sheet.cell(r, i_store).value or '').strip() if i_store else ''
        m = str(inventory_sheet.cell(r, i_msku).value or '').strip() if i_msku else ''
        if not s and not m: continue
        key = f"{s}{m}"
        
        # 尝试获取当前行的SKU用于备用匹配
        curr_sku = str(inventory_sheet.cell(r, i_sku).value or '').strip() if i_sku else ''

        if key in wfs_dict:
            d = wfs_dict[key]
            if i_gtin: inventory_sheet.cell(r, i_gtin, d['GTIN码'])
            if i_name and d['品名']: inventory_sheet.cell(r, i_name, d['品名'])
            if i_sku and d['sku']: 
                inventory_sheet.cell(r, i_sku, d['sku'])
                curr_sku = d['sku']
            if i_status: inventory_sheet.cell(r, i_status, d['商品状态'])
            if i_avail: inventory_sheet.cell(r, i_avail, d['WFS可售'])
            if i_unable: inventory_sheet.cell(r, i_unable, d['无法入库'])
            if i_transit: inventory_sheet.cell(r, i_transit, d['标发在途'])
            
        if key in sales_dict:
            if i_sales: inventory_sheet.cell(r, i_sales, sales_dict[key]['销量'])
            
        if curr_sku:
            if curr_sku in sz_stock_dict and i_sz: inventory_sheet.cell(r, i_sz, sz_stock_dict[curr_sku])
            if curr_sku in po_dict and i_po: inventory_sheet.cell(r, i_po, po_dict[curr_sku])

    # --- 5.2 添加新行 ---
    all_keys = set(wfs_dict.keys()) | set(sales_dict.keys())
    new_rows_data = []
    
    for key in all_keys:
        if key in existing_keys: continue
        
        # 基础数据结构
        d = {
            '店铺': '', 'msku': '', '店铺&MSKU': key, 'sku': '', '品名': '',
            'WFS可售': 0, '无法入库': 0, '标发在途': 0, '销量': 0
        }
        
        # 优先取WFS数据
        if key in wfs_dict:
            src = wfs_dict[key]
            d.update(src) # 覆盖 WFS可售, 无法入库, 标发在途, sku, 品名 等
            d['店铺'] = src['仓库']
            d['msku'] = src['msku']
        
        # 补全销量数据
        if key in sales_dict:
            src = sales_dict[key]
            d['销量'] = src['销量']
            if not d['店铺']: d['店铺'] = src['店铺']
            if not d['msku']: d['msku'] = src['msku']
            if not d['sku']: d['sku'] = src['SKU']
            if not d['品名']: d['品名'] = src['品名']

        # SKU补全
        if not d['sku']:
            extracted, _ = extract_sku_smart(d['msku'], sku_set)
            if extracted: d['sku'] = extracted
            
        # 品名补全
        if not d['品名'] and d['sku'] in sku_to_name:
            d['品名'] = sku_to_name[d['sku']]

        new_rows_data.append(d)

    # --- 写入新行 ---
    curr_row = original_max_row + 1
    for data in new_rows_data:
        if i_store: inventory_sheet.cell(curr_row, i_store, data['店铺'])
        if i_msku: inventory_sheet.cell(curr_row, i_msku, data['msku'])
        if i_store_msku: inventory_sheet.cell(curr_row, i_store_msku, data['店铺&MSKU'])
        if i_gtin and 'GTIN码' in data: inventory_sheet.cell(curr_row, i_gtin, data['GTIN码'])
        if i_sku: inventory_sheet.cell(curr_row, i_sku, data['sku'])
        if i_name: inventory_sheet.cell(curr_row, i_name, data['品名'])
        if i_status and '商品状态' in data: inventory_sheet.cell(curr_row, i_status, data['商品状态'])
        
        # 写入关键数值
        if i_avail: inventory_sheet.cell(curr_row, i_avail, data['WFS可售'])
        if i_unable: inventory_sheet.cell(curr_row, i_unable, data['无法入库'])
        if i_transit: inventory_sheet.cell(curr_row, i_transit, data['标发在途'])
        if i_sales: inventory_sheet.cell(curr_row, i_sales, data['销量'])
        
        # 深圳仓 & 采购
        sku = data['sku']
        if sku:
            if i_sz and sku in sz_stock_dict: inventory_sheet.cell(curr_row, i_sz, sz_stock_dict[sku])
            if i_po and sku in po_dict: inventory_sheet.cell(curr_row, i_po, po_dict[sku])
            
        curr_row += 1

    # === 第6步：计算与清理 ===
    # 6.1 计算公式
    for r in range(3, curr_row):
        v_wfs = get_numeric_value(inventory_sheet.cell(r, i_avail)) if i_avail else 0
        v_unable = get_numeric_value(inventory_sheet.cell(r, i_unable)) if i_unable else 0
        v_transit = get_numeric_value(inventory_sheet.cell(r, i_transit)) if i_transit else 0
        v_sz = get_numeric_value(inventory_sheet.cell(r, i_sz)) if i_sz else 0
        v_sales = get_numeric_value(inventory_sheet.cell(r, i_sales)) if i_sales else 0
        
        total = v_wfs + v_unable + v_transit + v_sz
        if i_total: inventory_sheet.cell(r, i_total, total)
        
        if i_turnover:
            if v_sales > 0:
                turnover = round((v_wfs + v_transit + v_sz) / v_sales * 30, 2)
                inventory_sheet.cell(r, i_turnover, turnover)
            else:
                inventory_sheet.cell(r, i_turnover, "")

    # 6.2 严格删除逻辑 (只针对新增行)
    # 规则：WFS匹配的三个数量 + 销量表匹配的数量 均为0/空
    rows_to_del = []
    for r in range(original_max_row + 1, curr_row):
        # 检查 WFS 来源的数值
        val_wfs = get_numeric_value(inventory_sheet.cell(r, i_avail)) if i_avail else 0
        val_unable = get_numeric_value(inventory_sheet.cell(r, i_unable)) if i_unable else 0
        val_transit = get_numeric_value(inventory_sheet.cell(r, i_transit)) if i_transit else 0
        # 检查 销量
        val_sales = get_numeric_value(inventory_sheet.cell(r, i_sales)) if i_sales else 0
        
        # 如果这4个关键值都为0，则删除
        if (val_wfs == 0 and val_unable == 0 and val_transit == 0 and val_sales == 0):
            rows_to_del.append(r)

    for r in sorted(rows_to_del, reverse=True):
        inventory_sheet.delete_rows(r, 1)

    st.success(f"✅ 处理完成！原有 {original_max_row} 行，新增 {len(new_rows_data) - len(rows_to_del)} 条有效记录 (已自动过滤无数据行)。")
    
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out

# UI 部分
st.set_page_config(page_title="沃尔玛库存工具 v12", layout="wide")
st.title("🛒 沃尔玛库存更新工具 (v12 深度修复版)")
st.markdown("""
**本次更新重点 (v12)：**
1. **智能列名匹配**：兼容 `标发在途(数量)`、`标发在途`、`In Transit` 等多种表头写法，并排除 `ID/代码` 列。
2. **数据补全**：强制抓取 `WFS可售`、`无法入库`、`标发在途` 和 `销量小计`。
3. **严格过滤**：新增记录中，如果上述 4 个关键数值均为 0，将自动删除该行。
""")

c1, c2 = st.columns(2)
with c1:
    f_inv = st.file_uploader("上传库存明细表 (必选)", type=['xlsx'], key="inv")
with c2:
    f_prod = st.file_uploader("上传产品资料表 (推荐)", type=['xlsx'], key="prod")

if f_inv and st.button("🚀 开始处理"):
    try:
        data = process_inventory(f_inv, f_prod)
        if data:
            st.download_button("📥 下载结果", data, f"Updated_{f_inv.name}")
    except Exception as e:
        st.error(f"发生错误: {e}")
