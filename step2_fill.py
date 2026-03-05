import openpyxl
import io
import re
import warnings

warnings.filterwarnings('ignore')

def get_numeric_value(cell):
    """安全获取数值"""
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
    """SKU 智能拆分"""
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
    模块二：采用 V10 级别的强硬列表头映射逻辑，绝对精准注入所有细分数据
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

    # ================= 1. 解析 WFS 数据 =================
    wfs_full = {}
    if wfs_sheet_name:
        w_sh = wb[wfs_sheet_name]
        wmap = {}
        for col in range(1, 40):
            val = w_sh.cell(row=1, column=col).value
            if not val: continue
            name = str(val).strip()
            if '仓库' in name: wmap['仓库'] = col
            elif 'msku' in name: wmap['msku'] = col
            elif name == 'GTIN码': wmap['GTIN'] = col
            elif '平台' in name and '商品' in name and 'ID' in name: wmap['平台商品ID'] = col
            elif '品名' in name and 'ID' not in name: wmap['品名'] = col
            elif name == 'sku': wmap['sku'] = col
            elif '商品状态' in name and '库龄' not in name: wmap['状态'] = col
            elif 'WFS可售(新)' in name and '数量' in name: wmap['WFS可售'] = col
            elif '无法入库' in name and '数量' in name: wmap['无法入库'] = col
            elif '标发在途' in name and '数量' in name: wmap['标发在途'] = col
            elif '3个月内库龄' in name and '数量' in name: wmap['3内'] = col
            elif '3-6个月库龄' in name and '数量' in name: wmap['3-6内'] = col
            elif '6个月以上库龄' in name and '数量' in name: wmap['6以上'] = col
            elif '12个月以上库龄' in name and '数量' in name: wmap['12以上'] = col
            elif '6-9个月' in name and '数量' in name: wmap['6-9'] = col
            elif '9-12个月' in name and '数量' in name: wmap['9-12'] = col

        for r in range(2, w_sh.max_row + 1):
            wh = str(w_sh.cell(r, wmap.get('仓库', 1)).value or '').strip()
            msku = str(w_sh.cell(r, wmap.get('msku', 2)).value or '').strip()
            if wh and msku:
                # 兼容 6个月以上的库龄被拆分成 6-9 和 9-12 的情况
                v_6_9 = get_numeric_value(w_sh.cell(r, wmap['6-9'])) if '6-9' in wmap else 0
                v_9_12 = get_numeric_value(w_sh.cell(r, wmap['9-12'])) if '9-12' in wmap else 0
                v_12p = get_numeric_value(w_sh.cell(r, wmap['12以上'])) if '12以上' in wmap else 0
                v_6plus = get_numeric_value(w_sh.cell(r, wmap['6以上'])) if '6以上' in wmap else (v_6_9 + v_9_12 + v_12p)

                wfs_full[f"{wh}{msku}"] = {
                    '平台商品ID': str(w_sh.cell(r, wmap['平台商品ID']).value or '') if '平台商品ID' in wmap else '',
                    'GTIN': str(w_sh.cell(r, wmap['GTIN']).value or '') if 'GTIN' in wmap else '',
                    '品名': str(w_sh.cell(r, wmap['品名']).value or '').strip() if '品名' in wmap else '',
                    'SKU': str(w_sh.cell(r, wmap['sku']).value or '').strip() if 'sku' in wmap else '',
                    '状态': w_sh.cell(r, wmap['状态']).value if '状态' in wmap else '',
                    '可售': get_numeric_value(w_sh.cell(r, wmap['WFS可售'])) if 'WFS可售' in wmap else 0,
                    '无法': get_numeric_value(w_sh.cell(r, wmap['无法入库'])) if '无法入库' in wmap else 0,
                    '在途': get_numeric_value(w_sh.cell(r, wmap['标发在途'])) if '标发在途' in wmap else 0,
                    '库龄3': get_numeric_value(w_sh.cell(r, wmap['3内'])) if '3内' in wmap else 0,
                    '库龄6': get_numeric_value(w_sh.cell(r, wmap['3-6内'])) if '3-6内' in wmap else 0,
                    '库龄12': v_6plus,
                    '库龄超12': v_12p,
                }

    # ================= 2. 解析 销量 数据 =================
    sales_full = {}
    if sales_sheet_name:
        s_sh = wb[sales_sheet_name]
        smap = {}
        for col in range(1, 50):
            val = s_sh.cell(row=1, column=col).value
            if not val: continue
            name = str(val).strip()
            if name == 'MSKU': smap['MSKU'] = col
            elif name == '店铺': smap['店铺'] = col
            elif name == '小计': smap['小计'] = col
            
        for r in range(2, s_sh.max_row + 1):
            msku = str(s_sh.cell(r, smap.get('MSKU', 4)).value or '').strip()
            store = str(s_sh.cell(r, smap.get('店铺', 3)).value or '').strip()
            if not store and '-' in msku: store = msku.split('-')[0]
            if msku:
                sales_full[f"{store}{msku}"] = get_numeric_value(s_sh.cell(r, smap.get('小计', 13)))

    # ================= 3. 解析 深圳与采购 数据 =================
    sz_full, po_full = {}, {}
    if sz_sheet_name:
        sz_sh = wb[sz_sheet_name]
        szmap = {}
        for col in range(1, 15):
            val = sz_sh.cell(row=1, column=col).value
            if not val: continue
            name = str(val).strip()
            if name == 'SKU': szmap['SKU'] = col
            elif '可用' in name: szmap['可用'] = col
        for r in range(2, sz_sh.max_row + 1):
            sku = str(sz_sh.cell(r, szmap.get('SKU', 1)).value or '').strip()
            if sku: sz_full[sku] = sz_full.get(sku, 0) + get_numeric_value(sz_sh.cell(r, szmap.get('可用', 10)))

    if po_sheet_name:
        po_sh = wb[po_sheet_name]
        pmap = {}
        for col in range(1, 35):
            val = po_sh.cell(row=1, column=col).value
            if not val: continue
            name = str(val).strip()
            if name == 'SKU': pmap['SKU'] = col
            elif '未入库' in name: pmap['未入库'] = col
        for r in range(2, po_sh.max_row + 1):
            sku = str(po_sh.cell(r, pmap.get('SKU', 7)).value or '').strip()
            if sku: po_full[sku] = po_full.get(sku, 0) + get_numeric_value(po_sh.cell(r, pmap.get('未入库', 19)))

    # ================= 4. 锁定主表列位，注回数据 =================
    inv_sh = wb[inv_sheet_name]
    imap = {}
    for col in range(1, 100):
        val = inv_sh.cell(row=2, column=col).value
        if not val: continue
        name = str(val).strip()
        if name == '店铺' and 'MSKU' not in name: imap['店铺'] = col
        elif name == 'msku' and '店铺' not in name: imap['msku'] = col
        elif name == 'GTIN码': imap['GTIN'] = col
        elif '平台' in name and '商品' in name and 'ID' in name: imap['平台商品ID'] = col
        elif name == '品名' and 'ID' not in name: imap['品名'] = col
        elif name == 'sku': imap['sku'] = col
        elif name == '商品状态': imap['状态'] = col
        elif 'WFS可售(新)' in name and '数量' in name: imap['WFS可售'] = col
        elif '无法入库' in name and '数量' in name: imap['无法入库'] = col
        elif '标发在途' in name and '数量' in name: imap['标发在途'] = col
        elif '深圳仓库存' in name: imap['深圳仓'] = col
        elif '采购订单在途' in name: imap['采购'] = col
        elif '总库存' in name: imap['总库存'] = col
        elif 'WFS在库周转' in name: imap['WFS在库周转'] = col
        elif 'WFS在途+在库周转' in name: imap['WFS在途周转'] = col
        elif '总周转天数（不含采购订单）' in name: imap['总周转'] = col
        elif '3个月内库龄' in name and '数量' in name: imap['3内'] = col
        elif '3-6个月库龄' in name and '数量' in name: imap['3-6内'] = col
        elif '6个月以上库龄' in name and '数量' in name: imap['6以上'] = col
        elif '12个月以上库龄' in name and '数量' in name: imap['12以上'] = col
        elif name == sales_sheet_name: imap['销量'] = col
        
    if '销量' not in imap:
        imap['销量'] = inv_sh.max_column + 1
        inv_sh.cell(2, imap['销量'], sales_sheet_name)

    empty_count = 0
    for r in range(3, inv_sh.max_row + 1):
        store = str(inv_sh.cell(r, imap.get('店铺', 1)).value or '').strip()
        msku = str(inv_sh.cell(r, imap.get('msku', 2)).value or '').strip()
        
        if not store and not msku:
            empty_count += 1
            if empty_count > 50: break
            continue
        empty_count = 0
        
        key = f"{store}{msku}"
        
        # --- 补全 SKU 和品名 ---
        curr_sku = str(inv_sh.cell(r, imap.get('sku', -1)).value or '').strip() if 'sku' in imap else ''
        if not curr_sku:
            if key in wfs_full and wfs_full[key]['SKU']:
                curr_sku = wfs_full[key]['SKU']
            elif msku:
                curr_sku = extract_sku_smart(msku, sku_set)
            if 'sku' in imap and curr_sku: 
                inv_sh.cell(r, imap['sku'], curr_sku)
        
        curr_name = str(inv_sh.cell(r, imap.get('品名', -1)).value or '').strip() if '品名' in imap else ''
        if not curr_name:
            if key in wfs_full and wfs_full[key]['品名']: curr_name = wfs_full[key]['品名']
            elif curr_sku in sku_to_name: curr_name = sku_to_name[curr_sku]
            if '品名' in imap and curr_name: inv_sh.cell(r, imap['品名'], curr_name)
                
        # --- 注入数量和详情 ---
        v_wfs = v_unable = v_transit = v_sz = v_po = v_sales = 0
        
        if key in wfs_full:
            d = wfs_full[key]
            v_wfs, v_unable, v_transit = d['可售'], d['无法'], d['在途']
            
            if '平台商品ID' in imap and d['平台商品ID']: inv_sh.cell(r, imap['平台商品ID'], d['平台商品ID'])
            if 'GTIN' in imap and d['GTIN']: inv_sh.cell(r, imap['GTIN'], d['GTIN'])
            if '状态' in imap and d['状态']: inv_sh.cell(r, imap['状态'], d['状态'])
            
            if 'WFS可售' in imap: inv_sh.cell(r, imap['WFS可售'], v_wfs)
            if '无法入库' in imap: inv_sh.cell(r, imap['无法入库'], v_unable)
            if '标发在途' in imap: inv_sh.cell(r, imap['标发在途'], v_transit)
            if '3内' in imap: inv_sh.cell(r, imap['3内'], d['库龄3'])
            if '3-6内' in imap: inv_sh.cell(r, imap['3-6内'], d['库龄6'])
            if '6以上' in imap: inv_sh.cell(r, imap['6以上'], d['库龄12'])
            if '12以上' in imap: inv_sh.cell(r, imap['12以上'], d['库龄超12'])
            
        if key in sales_full:
            v_sales = sales_full[key]
            if '销量' in imap: inv_sh.cell(r, imap['销量'], v_sales)
            
        if curr_sku:
            if curr_sku in sz_full:
                v_sz = sz_full[curr_sku]
                if '深圳仓' in imap: inv_sh.cell(r, imap['深圳仓'], v_sz)
            if curr_sku in po_full:
                v_po = po_full[curr_sku]
                if '采购' in imap: inv_sh.cell(r, imap['采购'], v_po)
                
        # --- 计算公式 ---
        if '总库存' in imap:
            inv_sh.cell(r, imap['总库存'], v_wfs + v_unable + v_transit + v_sz)
            
        if v_sales > 0:
            if 'WFS在库周转' in imap: inv_sh.cell(r, imap['WFS在库周转'], round(v_wfs / v_sales * 30, 2))
            if 'WFS在途周转' in imap: inv_sh.cell(r, imap['WFS在途周转'], round((v_wfs + v_transit) / v_sales * 30, 2))
            if '总周转' in imap: inv_sh.cell(r, imap['总周转'], round((v_wfs + v_transit + v_sz) / v_sales * 30, 2))
        else:
            if 'WFS在库周转' in imap: inv_sh.cell(r, imap['WFS在库周转'], "")
            if 'WFS在途周转' in imap: inv_sh.cell(r, imap['WFS在途周转'], "")
            if '总周转' in imap: inv_sh.cell(r, imap['总周转'], "")

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out
