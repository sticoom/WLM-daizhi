# step2_fill.py
import openpyxl
import io
import re
import warnings

warnings.filterwarnings('ignore')

def clean_header(val):
    if not val: return ""
    return str(val).strip().replace('（', '(').replace('）', ')').replace(' ', '')

def find_exact_col(sheet, exact_name):
    target = clean_header(exact_name)
    for col in range(1, 100):
        val1 = sheet.cell(row=1, column=col).value
        val2 = sheet.cell(row=2, column=col).value
        if val1 and target in clean_header(val1): return col
        if val2 and target in clean_header(val2): return col
    return None

def get_numeric_value(cell):
    if cell is None or cell.value is None: return 0
    val = str(cell.value).strip()
    if val == '' or val.lower() in ('nan', '#n/a', 'none'): return 0
    try:
        if val.startswith('='): return 0
        return float(val.replace(',', ''))
    except: return 0

def load_product_reference(product_file_obj):
    sku_set, sku_to_name = set(), {}
    if product_file_obj is None: return sku_set, sku_to_name
    try:
        wb = openpyxl.load_workbook(product_file_obj, read_only=True)
        ws = wb[wb.sheetnames[0]]
        sku_col, name_col = None, None
        for col in range(1, 20):
            val = str(ws.cell(1, col).value or '').lower()
            if 'sku' in val: sku_col = col
            if val in ['品名', '名称', 'name', '品名英']: name_col = col

        if sku_col:
            empty_count = 0
            for r in range(2, ws.max_row + 1):
                sku_val = str(ws.cell(r, sku_col).value or '').strip()
                if not sku_val:
                    empty_count += 1
                    if empty_count > 50: break
                    continue
                empty_count = 0
                sku_set.add(sku_val)
                if name_col:
                    name_val = str(ws.cell(r, name_col).value or '').strip()
                    if name_val: sku_to_name[sku_val] = name_val
        wb.close()
        return sku_set, sku_to_name
    except: return set(), {}

def extract_sku_smart(msku, sku_set):
    if not msku: return ''
    if not sku_set:
        parts = msku.split('-')
        return parts[1] if len(parts) >= 2 else parts[0]
    parts = [p.strip() for p in msku.split('-') if p.strip()]
    for p in parts:
        if p in sku_set: return p
    cleaned = [p.replace('"', '').replace("'", '').replace(' ', '') for p in parts]
    for p in cleaned:
        if p in sku_set: return p
    for p in cleaned:
        if len(p) >= 4 and re.search(r'\d', p) and re.search(r'[a-zA-Z]', p):
            if p in sku_set: return p
            for sku in sku_set:
                if p in sku or sku in p:
                    if len(sku) > 0 and len(p)/len(sku) >= 0.6: return sku
    return ''

def step2_fill_and_calculate(intermediate_file, product_file_obj):
    """
    模块二：遍历所有行，拆分SKU，回填所有细分数量指标，计算公式。
    """
    sku_set, sku_to_name = load_product_reference(product_file_obj)
    wb = openpyxl.load_workbook(intermediate_file)
    sheets = wb.sheetnames
    
    def find_sheet(kws):
        for s in sheets:
            if any(k in s for k in kws): return s
        return None

    inv_sheet_name = sheets[1]
    wfs_sheet_name = find_sheet(['WFS库存', 'WFS'])
    sz_sheet_name = find_sheet(['深圳仓', '深圳'])
    po_sheet_name = find_sheet(['采购订单', '采购', '在途'])
    sales_sheet_name = sheets[4] if len(sheets) > 4 else None

    # --- 获取全量数据 ---
    wfs_full = {}
    if wfs_sheet_name:
        w_sh = wb[wfs_sheet_name]
        c_wh = find_exact_col(w_sh, '仓库') or 1
        c_msku = find_exact_col(w_sh, 'msku') or 2
        c_gtin = find_exact_col(w_sh, 'GTIN码')
        c_status = find_exact_col(w_sh, '商品状态')
        c_avail = find_exact_col(w_sh, 'WFS可售(新)(数量)')
        c_unable = find_exact_col(w_sh, '无法入库(数量)')
        c_transit = find_exact_col(w_sh, '标发在途(数量)')
        c_a3 = find_exact_col(w_sh, '3个月内库龄(数量)')
        c_a6 = find_exact_col(w_sh, '3-6个月库龄(数量)')
        c_a12 = find_exact_col(w_sh, '6个月以上库龄(数量)')
        c_a12p = find_exact_col(w_sh, '12个月以上库龄(数量)')

        for r in range(2, w_sh.max_row + 1):
            wh = str(w_sh.cell(r, c_wh).value or '').strip()
            msku = str(w_sh.cell(r, c_msku).value or '').strip()
            if wh and msku:
                wfs_full[f"{wh}{msku}"] = {
                    'GTIN': str(w_sh.cell(r, c_gtin).value or '') if c_gtin else '',
                    '状态': w_sh.cell(r, c_status).value if c_status else '',
                    '可售': get_numeric_value(w_sh.cell(r, c_avail)) if c_avail else 0,
                    '无法': get_numeric_value(w_sh.cell(r, c_unable)) if c_unable else 0,
                    '在途': get_numeric_value(w_sh.cell(r, c_transit)) if c_transit else 0,
                    '库龄3': get_numeric_value(w_sh.cell(r, c_a3)) if c_a3 else 0,
                    '库龄6': get_numeric_value(w_sh.cell(r, c_a6)) if c_a6 else 0,
                    '库龄12': get_numeric_value(w_sh.cell(r, c_a12)) if c_a12 else 0,
                    '库龄超12': get_numeric_value(w_sh.cell(r, c_a12p)) if c_a12p else 0,
                }

    sales_full = {}
    if sales_sheet_name:
        s_sh = wb[sales_sheet_name]
        s_msku = find_exact_col(s_sh, 'MSKU') or 4
        s_store = find_exact_col(s_sh, '店铺') or 3
        s_subtotal = find_exact_col(s_sh, '小计') or 13
        for r in range(2, s_sh.max_row + 1):
            msku = str(s_sh.cell(r, s_msku).value or '').strip()
            store = str(s_sh.cell(r, s_store).value or '').strip()
            if not store and '-' in msku: store = msku.split('-')[0]
            if msku:
                sales_full[f"{store}{msku}"] = get_numeric_value(s_sh.cell(r, s_subtotal))

    sz_full, po_full = {}, {}
    if sz_sheet_name:
        sz_sh = wb[sz_sheet_name]
        sz_sku = find_exact_col(sz_sh, 'SKU') or 1
        sz_qty = find_exact_col(sz_sh, '实际可用') or find_exact_col(sz_sh, '可用') or 10
        for r in range(2, sz_sh.max_row + 1):
            sku = str(sz_sh.cell(r, sz_sku).value or '').strip()
            if sku: sz_full[sku] = sz_full.get(sku, 0) + get_numeric_value(sz_sh.cell(r, sz_qty))

    if po_sheet_name:
        po_sh = wb[po_sheet_name]
        po_sku = find_exact_col(po_sh, 'SKU') or 7
        po_qty = find_exact_col(po_sh, '未入库量') or find_exact_col(po_sh, '未入库') or 19
        for r in range(2, po_sh.max_row + 1):
            sku = str(po_sh.cell(r, po_sku).value or '').strip()
            if sku: po_full[sku] = po_full.get(sku, 0) + get_numeric_value(po_sh.cell(r, po_qty))

    # --- 注入主表与计算 ---
    inv_sh = wb[inv_sheet_name]
    
    i_store = find_exact_col(inv_sh, '店铺') or 1
    i_msku = find_exact_col(inv_sh, 'msku') or 2
    i_sku = find_exact_col(inv_sh, 'sku')
    i_name = find_exact_col(inv_sh, '品名')
    i_gtin = find_exact_col(inv_sh, 'GTIN码')
    i_status = find_exact_col(inv_sh, '商品状态')
    i_avail = find_exact_col(inv_sh, 'WFS可售(新)(数量)')
    i_unable = find_exact_col(inv_sh, '无法入库(数量)')
    i_transit = find_exact_col(inv_sh, '标发在途(数量)')
    i_a3 = find_exact_col(inv_sh, '3个月内库龄(数量)')
    i_a6 = find_exact_col(inv_sh, '3-6个月库龄(数量)')
    i_a12 = find_exact_col(inv_sh, '6个月以上库龄(数量)')
    i_a12p = find_exact_col(inv_sh, '12个月以上库龄(数量)')
    i_sz = find_exact_col(inv_sh, '深圳仓库存') or find_exact_col(inv_sh, '深圳仓')
    i_po = find_exact_col(inv_sh, '采购订单在途') or find_exact_col(inv_sh, '采购')
    
    i_total = find_exact_col(inv_sh, '总库存（不含采购订单）') or find_exact_col(inv_sh, '总库存')
    i_turn_wfs = find_exact_col(inv_sh, 'WFS在库周转')
    i_turn_transit = find_exact_col(inv_sh, 'WFS在途+在库周转')
    i_turn_total = find_exact_col(inv_sh, '总周转天数（不含采购订单）') or find_exact_col(inv_sh, '总周转')
    
    i_sales = find_exact_col(inv_sh, sales_sheet_name)
    if not i_sales:
        i_sales = inv_sh.max_column + 1
        inv_sh.cell(2, i_sales, sales_sheet_name)

    empty_count = 0
    for r in range(3, inv_sh.max_row + 1):
        store = str(inv_sh.cell(r, i_store).value or '').strip()
        msku = str(inv_sh.cell(r, i_msku).value or '').strip()
        
        if not store and not msku:
            empty_count += 1
            if empty_count > 50: break
            continue
        empty_count = 0
        
        curr_sku = str(inv_sh.cell(r, i_sku).value or '').strip() if i_sku else ''
        if not curr_sku and msku:
            curr_sku = extract_sku_smart(msku, sku_set)
            if i_sku and curr_sku: inv_sh.cell(r, i_sku, curr_sku)
                
        curr_name = str(inv_sh.cell(r, i_name).value or '').strip() if i_name else ''
        if not curr_name and curr_sku in sku_to_name and i_name:
            inv_sh.cell(r, i_name, sku_to_name[curr_sku])
                
        key = f"{store}{msku}"
        v_wfs = v_unable = v_transit = v_sz = v_po = v_sales = 0
        
        if key in wfs_full:
            d = wfs_full[key]
            v_wfs, v_unable, v_transit = d['可售'], d['无法'], d['在途']
            if i_avail: inv_sh.cell(r, i_avail, v_wfs)
            if i_unable: inv_sh.cell(r, i_unable, v_unable)
            if i_transit: inv_sh.cell(r, i_transit, v_transit)
            if i_a3: inv_sh.cell(r, i_a3, d['库龄3'])
            if i_a6: inv_sh.cell(r, i_a6, d['库龄6'])
            if i_a12: inv_sh.cell(r, i_a12, d['库龄12'])
            if i_a12p: inv_sh.cell(r, i_a12p, d['库龄超12'])
            if i_gtin and d['GTIN']: inv_sh.cell(r, i_gtin, d['GTIN'])
            if i_status and d['状态']: inv_sh.cell(r, i_status, d['状态'])
            
        if key in sales_full:
            v_sales = sales_full[key]
            if i_sales: inv_sh.cell(r, i_sales, v_sales)
            
        if curr_sku:
            if curr_sku in sz_full:
                v_sz = sz_full[curr_sku]
                if i_sz: inv_sh.cell(r, i_sz, v_sz)
            if curr_sku in po_full:
                v_po = po_full[curr_sku]
                if i_po: inv_sh.cell(r, i_po, v_po)
                
        if i_total:
            inv_sh.cell(r, i_total, v_wfs + v_unable + v_transit + v_sz)
            
        if v_sales > 0:
            if i_turn_wfs: inv_sh.cell(r, i_turn_wfs, round(v_wfs / v_sales * 30, 2))
            if i_turn_transit: inv_sh.cell(r, i_turn_transit, round((v_wfs + v_transit) / v_sales * 30, 2))
            if i_turn_total: inv_sh.cell(r, i_turn_total, round((v_wfs + v_transit + v_sz) / v_sales * 30, 2))
        else:
            if i_turn_wfs: inv_sh.cell(r, i_turn_wfs, "")
            if i_turn_transit: inv_sh.cell(r, i_turn_transit, "")
            if i_turn_total: inv_sh.cell(r, i_turn_total, "")

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out
