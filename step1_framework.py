# step1_framework.py
import openpyxl
import io
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

def get_real_max_row(sheet, start_row=3):
    real_max = start_row - 1
    empty_count = 0
    for r in range(start_row, sheet.max_row + 1):
        v1 = str(sheet.cell(r, 1).value or '').strip()
        v2 = str(sheet.cell(r, 2).value or '').strip()
        if not v1 and not v2:
            empty_count += 1
            if empty_count > 50: break
        else:
            real_max = r
            empty_count = 0
    return real_max

def get_numeric_value(cell):
    if cell is None or cell.value is None: return 0
    val = str(cell.value).strip()
    if val == '' or val.lower() in ('nan', '#n/a', 'none'): return 0
    try:
        if val.startswith('='): return 0
        return float(val.replace(',', ''))
    except: return 0

def step1_add_new_rows(inventory_file):
    """
    模块一：只负责比对 WFS 和销量表，找出新记录，
    并过滤掉 4 大核心指标全为 0 的无效行，最终构建出基础框架。
    """
    wb = openpyxl.load_workbook(inventory_file)
    sheets = wb.sheetnames
    
    def find_sheet(kws):
        for s in sheets:
            if any(k in s for k in kws): return s
        return None

    inv_sheet_name = sheets[1] if len(sheets) > 1 else None
    wfs_sheet_name = find_sheet(['WFS库存', 'WFS'])
    sales_sheet_name = sheets[4] if len(sheets) > 4 else None

    if not all([inv_sheet_name, wfs_sheet_name, sales_sheet_name]):
        raise ValueError("缺少必要的Sheet (库存明细、WFS库存、销量明细)")

    inv_sh = wb[inv_sheet_name]
    original_max_row = get_real_max_row(inv_sh, start_row=3)
    
    # 记录历史组合
    existing_keys = set()
    for r in range(3, original_max_row + 1):
        s = str(inv_sh.cell(r, 1).value or '').strip()
        m = str(inv_sh.cell(r, 2).value or '').strip()
        if s or m: existing_keys.add(f"{s}{m}")

    # 读取 WFS 基础验证数据
    wfs_sh = wb[wfs_sheet_name]
    wfs_dict = {}
    w_wh = find_exact_col(wfs_sh, '仓库') or 1
    w_msku = find_exact_col(wfs_sh, 'msku') or 2
    w_avail = find_exact_col(wfs_sh, 'WFS可售(新)(数量)')
    w_unable = find_exact_col(wfs_sh, '无法入库(数量)')
    w_transit = find_exact_col(wfs_sh, '标发在途(数量)')

    empty_count = 0
    for r in range(2, wfs_sh.max_row + 1):
        wh = str(wfs_sh.cell(r, w_wh).value or '').strip()
        msku = str(wfs_sh.cell(r, w_msku).value or '').strip()
        if not wh and not msku:
            empty_count += 1
            if empty_count > 50: break
            continue
        empty_count = 0
        if wh and msku:
            wfs_dict[f"{wh}{msku}"] = {
                '仓库': wh, 'msku': msku,
                '可售': get_numeric_value(wfs_sh.cell(r, w_avail)) if w_avail else 0,
                '无法': get_numeric_value(wfs_sh.cell(r, w_unable)) if w_unable else 0,
                '在途': get_numeric_value(wfs_sh.cell(r, w_transit)) if w_transit else 0
            }

    # 读取销量验证数据
    sales_sh = wb[sales_sheet_name]
    sales_dict = {}
    s_msku = find_exact_col(sales_sh, 'MSKU') or 4
    s_store = find_exact_col(sales_sh, '店铺') or 3
    s_subtotal = find_exact_col(sales_sh, '小计') or 13

    empty_count = 0
    for r in range(2, sales_sh.max_row + 1):
        msku = str(sales_sh.cell(r, s_msku).value or '').strip()
        store = str(sales_sh.cell(r, s_store).value or '').strip()
        if not msku and not store:
            empty_count += 1
            if empty_count > 50: break
            continue
        empty_count = 0
        if not store and '-' in msku: store = msku.split('-')[0]
        if msku:
            sales_dict[f"{store}{msku}"] = {
                '店铺': store, 'msku': msku, 
                '销量': get_numeric_value(sales_sh.cell(r, s_subtotal)) if s_subtotal else 0
            }

    # 找出新行并执行 4 大核心过滤
    all_keys = set(wfs_dict.keys()) | set(sales_dict.keys())
    new_rows = []
    
    for key in all_keys:
        if key in existing_keys: continue
        
        v_wfs = wfs_dict[key]['可售'] if key in wfs_dict else 0
        v_unable = wfs_dict[key]['无法'] if key in wfs_dict else 0
        v_transit = wfs_dict[key]['在途'] if key in wfs_dict else 0
        v_sales = sales_dict[key]['销量'] if key in sales_dict else 0
        
        # 严格过滤：4 项全 0 则抛弃
        if v_wfs == 0 and v_unable == 0 and v_transit == 0 and v_sales == 0:
            continue
            
        d = {'店铺&MSKU': key}
        if key in wfs_dict:
            d['店铺'] = wfs_dict[key]['仓库']
            d['msku'] = wfs_dict[key]['msku']
        elif key in sales_dict:
            d['店铺'] = sales_dict[key]['店铺']
            d['msku'] = sales_dict[key]['msku']
            
        new_rows.append(d)

    # 将新框架追加到末尾
    i_store = find_exact_col(inv_sh, '店铺') or 1
    i_msku = find_exact_col(inv_sh, 'msku') or 2
    i_store_msku = find_exact_col(inv_sh, '店铺&MSKU')
    
    curr_row = original_max_row + 1
    for d in new_rows:
        if i_store: inv_sh.cell(curr_row, i_store, d.get('店铺', ''))
        if i_msku: inv_sh.cell(curr_row, i_msku, d.get('msku', ''))
        if i_store_msku: inv_sh.cell(curr_row, i_store_msku, d.get('店铺&MSKU', ''))
        curr_row += 1

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out, len(new_rows)
