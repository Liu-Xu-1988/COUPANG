import streamlit as st
import pandas as pd
import io
import re

# ==========================================
# 1. 页面配置 (宽屏)
# ==========================================
st.set_page_config(layout="wide", page_title="Coupang 经营看板 Pro (最终版)")
st.title("📊 Coupang 经营分析看板 (最终版·店铺增强)")

# --- 列号配置 (请根据实际Excel列号修改) ---
IDX_M_CODE   = 0    # A列: 内部编码
IDX_M_SHOP   = 1    # B列: 登品店铺 (新增) <--- 请确认您的店铺名是否在B列(索引1)
IDX_M_SKU    = 3    # D列: SKU ID
IDX_M_COST   = 6    # G列: 采购价格
IDX_M_PROFIT = 10   # K列: 单品毛利
IDX_M_BAR    = 12   # M列: ID号码

IDX_S_ID     = 0    # A列
IDX_S_QTY    = 8    # I列

IDX_A_CAMPAIGN = 5  # F列
IDX_A_GROUP    = 6  # G列
IDX_A_SPEND    = 15 # P列
IDX_A_SALES    = 29 # AD列

IDX_I_R_ID   = 2    # C列
IDX_I_R_QTY  = 7    # H列

IDX_I_J_BAR  = 2    # C列
IDX_I_J_QTY  = 10   # K列

# ==========================================
# 2. 侧边栏
# ==========================================
with st.sidebar:
    st.header("🔍 数据筛选")
    filter_code = st.text_input("输入产品编号 (如 C123)", placeholder="留空则显示全部...").strip().upper()
    
    st.write("") 
    filter_profit = st.radio(
        "💰 利润筛选 (最终净利润)",
        ("全部显示", "只看盈利 (>0)", "只看亏损 (<0)"),
        index=0
    )
    
    st.divider()
    
    st.header("👁️ 视图设置")
    table_height = st.slider("表格显示高度 (像素)", 600, 3000, 1500, step=100)

    st.divider()
    
    st.header("📂 数据源上传")
    st.info("⚠️ 前3个为必传项，后2个为选传项")
    
    file_master = st.file_uploader("1. 基础信息表 (Master) *必传", type=['csv', 'xlsx', 'xlsm'])
    files_sales = st.file_uploader("2. 销售表 (Sales) *必传", type=['csv', 'xlsx', 'xlsm'], accept_multiple_files=True)
    files_ads = st.file_uploader("3. 广告表 (Ads) *必传", type=['csv', 'xlsx', 'xlsm'], accept_multiple_files=True)
    files_inv = st.file_uploader("4. 库存信息表 (火箭仓 Rocket)", type=['csv', 'xlsx', 'xlsm'], accept_multiple_files=True)
    files_inv_j = st.file_uploader("5. 库存信息表 (极风OMS)", type=['csv', 'xlsx', 'xlsm'], accept_multiple_files=True)

# ==========================================
# 3. 工具函数
# ==========================================
def clean_for_match(series):
    return series.astype(str).str.replace(r'\.0$', '', regex=True).str.replace('"', '').str.strip().str.upper()

def clean_num(series):
    return pd.to_numeric(series.astype(str).str.replace(',', ''), errors='coerce').fillna(0)

def extract_code_from_text(text):
    if pd.isna(text): return None
    match = re.search(r'([Cc]\d+)', str(text))
    if match: return match.group(1).upper()
    return None

def read_file_strict(file):
    try:
        file.seek(0)
        if file.name.endswith('.csv'):
            return pd.read_csv(file, dtype=str)
        else:
            return pd.read_excel(file, dtype=str, engine='openpyxl')
    except:
        file.seek(0)
        return pd.read_csv(file, dtype=str, encoding='gbk')

# ==========================================
# 4. 主逻辑
# ==========================================
st.divider()

missing_files = []
if not file_master: missing_files.append("1.基础信息表")
if not files_sales: missing_files.append("2.销售表")
if not files_ads: missing_files.append("3.广告表")

if missing_files:
    st.warning(f"👉 请在左侧上传必要文件后开始分析。当前缺失：{'、'.join(missing_files)}")
else:
    btn_label = "🚀 开始生成报表"
    filters_applied = []
    if filter_code: filters_applied.append(f"编号:{filter_code}")
    if filter_profit != "全部显示": filters_applied.append(f"{filter_profit}")
    
    if filters_applied:
        btn_label += f" (筛选: {' + '.join(filters_applied)})"
    
    if st.button(btn_label, type="primary", use_container_width=True):
        try:
            with st.spinner("正在全速计算中..."):
                # --- Step 1: 读取基础表 ---
                df_master = read_file_strict(file_master)
                col_code_name = df_master.columns[IDX_M_CODE]

                df_calc = df_master.copy()
                df_calc['_MATCH_SKU'] = clean_for_match(df_calc.iloc[:, IDX_M_SKU])
                df_calc['_MATCH_BAR'] = clean_for_match(df_calc.iloc[:, IDX_M_BAR])
                df_calc['_MATCH_CODE'] = clean_for_match(df_calc.iloc[:, IDX_M_CODE])
                df_calc['_VAL_PROFIT'] = clean_num(df_calc.iloc[:, IDX_M_PROFIT])
                df_calc['_VAL_COST'] = clean_num(df_calc.iloc[:, IDX_M_COST])
                # 【新增】读取店铺名
                df_calc['_MATCH_SHOP'] = df_calc.iloc[:, IDX_M_SHOP].astype(str).str.strip()

                # --- Step 2: 销售表 ---
                sales_list = [read_file_strict(f) for f in files_sales]
                df_sales_all = pd.concat(sales_list, ignore_index=True)
                df_sales_all['_MATCH_SKU'] = clean_for_match(df_sales_all.iloc[:, IDX_S_ID])
                df_sales_all['销量'] = clean_num(df_sales_all.iloc[:, IDX_S_QTY])
                sales_agg = df_sales_all.groupby('_MATCH_SKU')['销量'].sum().reset_index()
                sales_agg.rename(columns={'销量': 'SKU销量'}, inplace=True) 

                # --- Step 3: 广告表 ---
                ads_list = [read_file_strict(f) for f in files_ads]
                df_ads_all = pd.concat(ads_list, ignore_index=True)
                df_ads_all['含税广告费'] = clean_num(df_ads_all.iloc[:, IDX_A_SPEND]) * 1.1
                df_ads_all['广告销量'] = clean_num(df_ads_all.iloc[:, IDX_A_SALES])
                df_ads_all['Code_Group'] = df_ads_all.iloc[:, IDX_A_GROUP].apply(extract_code_from_text)
                df_ads_all['Code_Campaign'] = df_ads_all.iloc[:, IDX_A_CAMPAIGN].apply(extract_code_from_text)
                df_ads_all['_MATCH_CODE'] = df_ads_all['Code_Group'].fillna(df_ads_all['Code_Campaign'])
                valid_ads = df_ads_all.dropna(subset=['_MATCH_CODE'])
                ads_agg = valid_ads.groupby('_MATCH_CODE')[['含税广告费', '广告销量']].sum().reset_index()
                ads_agg.rename(columns={'含税广告费': 'R列_产品总广告费', '广告销量': '产品广告销量'}, inplace=True)

                # --- Step 4: 库存表 ---
                if files_inv:
                    inv_list = [read_file_strict(f) for f in files_inv]
                    df_inv_all = pd.concat(inv_list, ignore_index=True)
                    df_inv_all['_MATCH_SKU'] = clean_for_match(df_inv_all.iloc[:, IDX_I_R_ID])
                    df_inv_all['火箭仓库存'] = clean_num(df_inv_all.iloc[:, IDX_I_R_QTY])
                    inv_agg = df_inv_all.groupby('_MATCH_SKU')['火箭仓库存'].sum().reset_index()
                else:
                    inv_agg = pd.DataFrame(columns=['_MATCH_SKU', '火箭仓库存'])

                if files_inv_j:
                    inv_j_list = [read_file_strict(f) for f in files_inv_j]
                    df_inv_j_all = pd.concat(inv_j_list, ignore_index=True)
                    df_inv_j_all['_MATCH_BAR'] = clean_for_match(df_inv_j_all.iloc[:, IDX_I_J_BAR])
                    df_inv_j_all['极风库存'] = clean_num(df_inv_j_all.iloc[:, IDX_I_J_QTY])
                    inv_j_agg = df_inv_j_all.groupby('_MATCH_BAR')['极风库存'].sum().reset_index()
                else:
                    inv_j_agg = pd.DataFrame(columns=['_MATCH_BAR', '极风库存'])

                # --- Step 5: 计算 ---
                df_final = pd.merge(df_calc, sales_agg, on='_MATCH_SKU', how='left', sort=False)
                df_final['SKU销量'] = df_final['SKU销量'].fillna(0).astype(int)
                df_final = pd.merge(df_final, inv_agg, on='_MATCH_SKU', how='left', sort=False)
                df_final['火箭仓库存'] = df_final['火箭仓库存'].fillna(0).astype(int)
                df_final = pd.merge(df_final, inv_j_agg, on='_MATCH_BAR', how='left', sort=False)
                df_final['极风库存'] = df_final['极风库存'].fillna(0).astype(int)

                df_final['P列_SKU总毛利'] = df_final['SKU销量'] * df_final['_VAL_PROFIT']
                df_final['Q列_产品总利润'] = df_final.groupby('_MATCH_CODE', sort=False)['P列_SKU总毛利'].transform('sum')
                df_final['产品总销量'] = df_final.groupby('_MATCH_CODE', sort=False)['SKU销量'].transform('sum')
                
                df_final = pd.merge(df_final, ads_agg, on='_MATCH_CODE', how='left', sort=False)
                
                # 广告费 int
                df_final['R列_产品总广告费'] = df_final['R列_产品总广告费'].fillna(0).round(0).astype(int)
                
                df_final['产品广告销量'] = df_final['产品广告销量'].fillna(0)
                df_final['S列_最终净利润'] = df_final['Q列_产品总利润'] - df_final['R列_产品总广告费']

                # --- Step 6: 报表构造 ---
                
                # Sheet2 (业务报表 - 产品维度)
                df_final['产品_火箭仓库存'] = df_final.groupby('_MATCH_CODE', sort=False)['火箭仓库存'].transform('sum')
                df_final['产品_极风库存'] = df_final.groupby('_MATCH_CODE', sort=False)['极风库存'].transform('sum')
                df_final['产品_总库存'] = df_final['产品_火箭仓库存'] + df_final['产品_极风库存']

                # 【修改点】加入了 _MATCH_SHOP (店铺名)
                df_sheet2 = df_final[[col_code_name, '_MATCH_SHOP', 'Q列_产品总利润', 'R列_产品总广告费', 'S列_最终净利润', '产品总销量', '产品广告销量', '产品_火箭仓库存', '产品_极风库存', '产品_总库存']].copy()
                df_sheet2 = df_sheet2.drop_duplicates(subset=[col_code_name], keep='first')
                
                df_sheet2.rename(columns={
                    '_MATCH_SHOP': '登品店铺', # 重命名
                    '产品_火箭仓库存': '火箭仓库存', 
                    '产品_极风库存': '极风库存',
                    '产品_总库存': '总库存'
                }, inplace=True)

                df_sheet2['广告费占比'] = df_sheet2.apply(
                    lambda x: x['R列_产品总广告费'] / x['Q列_产品总利润'] if x['Q列_产品总利润'] != 0 else 0, axis=1
                )
                df_sheet2['自然销量'] = df_sheet2['产品总销量'] - df_sheet2['产品广告销量']
                df_sheet2['自然销量占比'] = df_sheet2.apply(
                    lambda x: x['自然销量'] / x['产品总销量'] if x['产品总销量'] != 0 else 0, axis=1
                )
                
                # 【修改点】把 '登品店铺' 放在第2位 (Code之后)
                cols_order_s2 = [
                    col_code_name, '登品店铺', 
                    'Q列_产品总利润', 'R列_产品总广告费', 'S列_最终净利润', 
                    '广告费占比', '自然销量占比', 
                    '总库存', 
                    '产品总销量', '产品广告销量', '自然销量', '自然销量占比',
                    '火箭仓库存', '极风库存'
                ]
                cols_order_s2 = list(dict.fromkeys(cols_order_s2))
                df_sheet2 = df_sheet2[cols_order_s2]

                # Sheet3 (库存分析 - SKU维度)
                df_final['火箭仓库存数量'] = df_final['火箭仓库存']
                df_final['总库存'] = df_final['火箭仓库存数量'] + df_final['极风库存']
                df_final['库存货值'] = df_final['总库存'] * df_final['_VAL_COST'] * 1.2
                df_final['安全库存'] = df_final['SKU销量'] * 3
                df_final['冗余标准'] = df_final['SKU销量'] * 8
                
                df_final['待补数量'] = df_final.apply(
                    lambda x: (x['安全库存'] - x['总库存']) if x['总库存'] < x['安全库存'] else 0,
                    axis=1
                )

                def calc_dead_stock_value(row):
                    total = row['总库存']
                    redundant_std = row['冗余标准']
                    if total == 0 and redundant_std == 0: return 0
                    if total >= redundant_std: return row['库存货值']
                    return 0
                df_final['滞销库存货值'] = df_final.apply(calc_dead_stock_value, axis=1)

                cols_master_AM = df_master.columns[:13].tolist()
                
                # Sheet1 (利润分析 - SKU维度)
                cols_s1_final = cols_master_AM + [
                    'SKU销量', 'P列_SKU总毛利', 'Q列_产品总利润', 'R列_产品总广告费', 'S列_最终净利润'
                ]
                df_final_clean = df_final[cols_s1_final].copy()

                cols_inv_final = cols_master_AM + [
                    '火箭仓库存数量', '极风库存', '总库存', 
                    '库存货值', '滞销库存货值', 
                    '待补数量', 
                    'SKU销量', '安全库存', '冗余标准'
                ]
                df_sheet3 = df_final[cols_inv_final].copy()

                # 重命名
                rename_dict = {
                    'P列_SKU总毛利': 'SKU总毛利',
                    'Q列_产品总利润': '产品总利润',
                    'R列_产品总广告费': '产品总广告费',
                    'S列_最终净利润': '最终净利润'
                }
                df_final_clean.rename(columns=rename_dict, inplace=True)
                df_sheet2.rename(columns=rename_dict, inplace=True)

                # --- 筛选 ---
                if filter_code:
                    df_final_clean = df_final_clean[df_final_clean[col_code_name].astype(str).str.contains(filter_code, na=False)]
                    df_sheet2 = df_sheet2[df_sheet2[col_code_name].astype(str).str.contains(filter_code, na=False)]
                    df_sheet3 = df_sheet3[df_sheet3[col_code_name].astype(str).str.contains(filter_code, na=False)]

                if filter_profit == "只看盈利 (>0)":
                    df_final_clean = df_final_clean[df_final_clean['最终净利润'] > 0]
                    valid_indices = df_final_clean.index
                    df_sheet3 = df_sheet3.loc[df_sheet3.index.isin(valid_indices)]
                    df_sheet2 = df_sheet2[df_sheet2['最终净利润'] > 0]
                elif filter_profit == "只看亏损 (<0)":
                    df_final_clean = df_final_clean[df_final_clean['最终净利润'] < 0]
                    valid_indices = df_final_clean.index
                    df_sheet3 = df_sheet3.loc[df_sheet3.index.isin(valid_indices)]
                    df_sheet2 = df_sheet2[df_sheet2['最终净利润'] < 0]

                # 插入序号列 (注意：中文版方括号)
                df_sheet2.reset_index(drop=True, inplace=True)
                # 【修改点】使用中文方括号
                idx_col_name_s2 = f"产品总数【{len(df_sheet2)}】"
                df_sheet2.insert(0, idx_col_name_s2, range(1, len(df_sheet2) + 1))

                df_final_clean.reset_index(drop=True, inplace=True)
                idx_col_name_s1 = f"SKU总数【{len(df_final_clean)}】"
                df_final_clean.insert(0, idx_col_name_s1, range(1, len(df_final_clean) + 1))

                df_sheet3.reset_index(drop=True, inplace=True)
                idx_col_name_s3 = f"SKU总数【{len(df_sheet3)}】"
                df_sheet3.insert(0, idx_col_name_s3, range(1, len(df_sheet3) + 1))

                # ==========================================
                # 🔥 看板展示
                # ==========================================
                if df_sheet2.empty:
                    st.warning("⚠️ 筛选结果为空")
                else:
                    net_profit = df_sheet2['最终净利润'].sum()
                    inv_val = df_sheet3['库存货值'].sum()
                    dead_val = df_sheet3['滞销库存货值'].sum()
                    restock = df_sheet3['待补数量'].sum()
                    total_qty = df_sheet2['产品总销量'].sum()
                    
                    st.subheader("📈 经营概览")
                    k1, k2, k3, k4, k5 = st.columns(5)
                    k1.metric("💰 最终净利润", f"{net_profit:,.0f}", delta_color="normal" if net_profit>0 else "inverse")
                    k2.metric("📦 总销售数量", f"{total_qty:,.0f}")
                    k3.metric("🏭 库存总货值", f"¥ {inv_val:,.0f}")
                    k4.metric("🔴 滞销资金", f"¥ {dead_val:,.0f}", delta="风险", delta_color="inverse")
                    k5.metric("🚨 待补数量", f"{restock:,.0f}")

                    st.divider()

                    # === 样式函数 ===
                    def safe_fmt_int(x):
                        try:
                            if pd.isna(x) or x == '': return ""
                            return "{:,.0f}".format(float(x))
                        except: return str(x)

                    def safe_fmt_pct(x):
                        try:
                            if pd.isna(x) or x == '': return ""
                            return "{:.1%}".format(float(x))
                        except: return str(x)

                    def get_format_dict(df):
                        format_dict = {}
                        for col in df.columns:
                            c_str = str(col)
                            if any(x in c_str for x in ['比', '率', '占比']):
                                format_dict[col] = safe_fmt_pct
                            elif any(x in c_str for x in ['利润', '费用', '货值', '金额', '毛利', '销量', '库存', '数量', '标准', '待补', '序号', '广告费']):
                                format_dict[col] = safe_fmt_int
                        return format_dict

                    def apply_visual_style(df, cols_to_color, is_sheet2=False):
                        try:
                            styler = df.style.format(get_format_dict(df))
                            def zebra_rows(x):
                                # 如果是Sheet2，前2列都是文本(序号+店铺)，用第3列(Code)做斑马纹
                                col_idx = 2 if is_sheet2 else 1
                                codes = x.iloc[:, col_idx].astype(str)
                                groups = (codes != codes.shift()).cumsum()
                                is_odd = groups % 2 != 0
                                styles = pd.DataFrame('', index=x.index, columns=x.columns)
                                styles.loc[is_odd, :] = 'background-color: #f0f2f6' 
                                return styles
                            styler = styler.apply(zebra_rows, axis=None)
                            
                            def highlight_cells(x):
                                styles = []
                                for col in x.index:
                                    style = ''
                                    if col in ['自然销量占比', '总库存']:
                                        style += 'font-weight: bold;'
                                    if col == '广告费占比':
                                        try:
                                            if x[col] > 0.5: style += 'color: #d32f2f; font-weight: bold;'
                                        except: pass
                                    styles.append(style)
                                return styles
                            styler = styler.apply(highlight_cells, axis=1)

                            valid_cols = [c for c in cols_to_color if c in df.columns]
                            if valid_cols:
                                styler = styler.background_gradient(subset=valid_cols, cmap='RdYlGn', vmin=-10000, vmax=10000)
                            return styler
                        except: return df
                    
                    def apply_inventory_style(df):
                        try:
                            styler = df.style.format(get_format_dict(df))
                            def zebra_rows(x):
                                codes = x.iloc[:, 1].astype(str)
                                groups = (codes != codes.shift()).cumsum()
                                is_odd = groups % 2 != 0
                                styles = pd.DataFrame('', index=x.index, columns=x.columns)
                                styles.loc[is_odd, :] = 'background-color: #f0f2f6' 
                                return styles
                            styler = styler.apply(zebra_rows, axis=None)

                            def highlight_logic(x):
                                styles = []
                                for col in x.index:
                                    style = ''
                                    if col == '待补数量' and x['待补数量'] > 0:
                                        style += 'background-color: #fff3cd; color: #e65100; font-weight: bold;'
                                    if col == '滞销库存货值' and x['滞销库存货值'] > 0:
                                        style += 'color: #880e4f; font-weight: bold;'
                                    if col == '总库存':
                                        try:
                                            total = x['总库存']
                                            safe = x['安全库存']
                                            redundant = x['冗余标准']
                                            if total == 0 and redundant == 0: pass 
                                            elif total < safe: style += 'background-color: #ffcccc; color: #cc0000; font-weight: bold;'
                                            elif total >= redundant: style += 'background-color: #e1bee7; color: #4a148c; font-weight: bold;'
                                        except: pass
                                    styles.append(style)
                                return styles
                            styler = styler.apply(highlight_logic, axis=1)
                            return styler
                        except: return df

                    # 默认使用瀑布流
                    st.markdown("### 📝 1. 利润分析")
                    st.dataframe(apply_visual_style(df_final_clean, ['最终净利润']), use_container_width=True, height=table_height, hide_index=True)
                    
                    st.markdown("### 📊 2. 业务报表")
                    st.dataframe(apply_visual_style(df_sheet2, ['最终净利润'], True), use_container_width=True, height=table_height, hide_index=True)
                    
                    st.markdown("### 🏭 3. 库存分析")
                    try:
                        st_inv = apply_inventory_style(df_sheet3)
                        st_inv = st_inv.bar(subset=['总库存'], color='#800080')\
                                       .bar(subset=['库存货值'], color='#2ca02c')\
                                       .bar(subset=['滞销库存货值'], color='#880e4f')
                        st.dataframe(st_inv, use_container_width=True, height=table_height, hide_index=True)
                    except:
                        st.dataframe(df_sheet3, use_container_width=True, hide_index=True)

                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        df_final_clean.to_excel(writer, index=False, sheet_name='利润分析')
                        df_sheet2.to_excel(writer, index=False, sheet_name='业务报表')
                        df_sheet3.to_excel(writer, index=False, sheet_name='库存分析')
                        
                        wb = writer.book
                        fmt_header = wb.add_format({'bold': True, 'bg_color': '#4472C4', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
                        fmt_int = wb.add_format({'num_format': '#,##0', 'align': 'center'})
                        fmt_pct = wb.add_format({'num_format': '0.0%', 'align': 'center'})
                        fmt_pct_bold = wb.add_format({'num_format': '0.0%', 'align': 'center', 'bold': True})
                        fmt_int_bold = wb.add_format({'num_format': '#,##0', 'align': 'center', 'bold': True})
                        fmt_red_alert = wb.add_format({'num_format': '0.0%', 'align': 'center', 'bold': True, 'font_color': '#9C0006', 'bg_color': '#FFC7CE'})
                        fmt_grey = wb.add_format({'bg_color': '#BFBFBF', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
                        fmt_white = wb.add_format({'bg_color': '#FFFFFF', 'border': 1, 'align': 'center', 'valign': 'vcenter'})

                        def set_sheet_format(sheet_name, df_obj, group_col_idx):
                            ws = writer.sheets[sheet_name]
                            # 业务报表有 序号+店铺，所以组列在第3位(idx=2)
                            # 其他表有 序号，所以组列在第2位(idx=1)
                            actual_group_col = 2 if sheet_name == '业务报表' else 1
                            
                            raw_codes = df_obj.iloc[:, actual_group_col].astype(str).tolist()
                            clean_codes = [str(x).replace('.0','').replace('"','').strip().upper() for x in raw_codes]
                            is_grey = False
                            for i in range(len(raw_codes)):
                                if i > 0 and clean_codes[i] != clean_codes[i-1]: is_grey = not is_grey
                                ws.set_row(i + 1, None, fmt_grey if is_grey else fmt_white)
                            
                            for i, col in enumerate(df_obj.columns):
                                c_str = str(col)
                                width = 12
                                cell_fmt = None
                                is_bold_col = col in ['自然销量占比', '总库存']
                                
                                if any(x in c_str for x in ['比', '率', '占比']):
                                    cell_fmt = fmt_pct_bold if is_bold_col else fmt_pct
                                    width = 12
                                elif any(x in c_str for x in ['利润', '费用', '货值', '金额', '毛利', '销量', '库存', '数量', '标准', '待补', '序号', '广告费', '总数']):
                                    cell_fmt = fmt_int_bold if is_bold_col else fmt_int
                                    width = 15
                                
                                if cell_fmt: ws.set_column(i, i, width, cell_fmt)
                                else: ws.set_column(i, i, width)
                                ws.write(0, i, col, fmt_header)
                                
                                if col == '广告费占比':
                                    ws.conditional_format(1, i, len(df_obj), i, {'type': 'cell', 'criteria': '>', 'value': 0.5, 'format': fmt_red_alert})

                        set_sheet_format('利润分析', df_final_clean, IDX_M_CODE)
                        set_sheet_format('业务报表', df_sheet2, IDX_M_CODE)
                        set_sheet_format('库存分析', df_sheet3, IDX_M_CODE)

                    st.download_button(
                        label="📥 下载 Excel",
                        data=output.getvalue(),
                        file_name=f"Coupang_Report_{filter_code if filter_code else 'All'}.xlsx",
                        mime="application/vnd.ms-excel",
                        type="primary",
                        use_container_width=True
                    )

        except Exception as e:
            st.error(f"❌ 运行出错: {e}")
