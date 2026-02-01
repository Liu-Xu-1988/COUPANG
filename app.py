import streamlit as st
import pandas as pd
import io
import re

# ==========================================
# 1. 页面配置 (宽屏)
# ==========================================
st.set_page_config(layout="wide", page_title="Coupang 经营看板 Pro (最终版)")
st.title("📊 Coupang 经营分析看板 (完美表头版)")

# --- 列号配置 ---
# Master表 (基础表)
IDX_M_CODE   = 0    # A列: 内部编码
IDX_M_SKU    = 3    # D列: SKU ID (用于匹配火箭仓)
IDX_M_COST   = 6    # G列: 采购价格 (RMB)
IDX_M_PROFIT = 10   # K列: 单品毛利
IDX_M_BAR    = 12   # M列: ID号码 (用于匹配极风库存)

# Sales表 (销售表)
IDX_S_ID     = 0    # A列
IDX_S_QTY    = 8    # I列

# Ads表 (广告表)
IDX_A_CAMPAIGN = 5  # F列
IDX_A_GROUP    = 6  # G列
IDX_A_SPEND    = 15 # P列
IDX_A_SALES    = 29 # AD列 (30列)

# Inventory Rocket (火箭仓)
IDX_I_R_ID   = 2    # C列: ID
IDX_I_R_QTY  = 7    # H列: 库存数量

# Inventory Jifeng (极风)
IDX_I_J_BAR  = 2    # C列: 产品条码
IDX_I_J_QTY  = 10   # K列: 数值
# -----------------

# ==========================================
# 2. 侧边栏 (含筛选 & 上传)
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
    
    st.header("📂 数据源上传")
    st.info("请按顺序上传以下文件：")
    
    file_master = st.file_uploader("1. 基础信息表 (Master)", type=['csv', 'xlsx', 'xlsm'])
    files_sales = st.file_uploader("2. 销售表 (Sales - 近1周数据)", type=['csv', 'xlsx', 'xlsm'], accept_multiple_files=True)
    files_ads = st.file_uploader("3. 广告表 (Ads)", type=['csv', 'xlsx', 'xlsm'], accept_multiple_files=True)
    files_inv = st.file_uploader("4. 库存信息表 (火箭仓 Rocket)", type=['csv', 'xlsx', 'xlsm'], accept_multiple_files=True)
    files_inv_j = st.file_uploader("5. 库存信息表 (极风OMS)", type=['csv', 'xlsx', 'xlsm'], accept_multiple_files=True)

# ==========================================
# 3. 清洗工具函数
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
if file_master and files_sales and files_ads:
    st.divider()
    
    btn_label = "🚀 生成完美报表"
    filters_applied = []
    if filter_code: filters_applied.append(f"编号:{filter_code}")
    if filter_profit != "全部显示": filters_applied.append(f"{filter_profit}")
    
    if filters_applied:
        btn_label += f" (筛选: {' + '.join(filters_applied)})"
    
    if st.button(btn_label, type="primary", use_container_width=True):
        try:
            with st.spinner("正在进行多维数据计算..."):
                
                # --- Step 1: 基础表 ---
                df_master = read_file_strict(file_master)
                # 强制只取前13列 (A-M)
                df_master_clean = df_master.iloc[:, :13].copy()
                col_code_name = df_master.columns[IDX_M_CODE]

                # 辅助计算表
                df_calc = df_master.copy()
                df_calc['_MATCH_SKU'] = clean_for_match(df_calc.iloc[:, IDX_M_SKU])
                df_calc['_MATCH_BAR'] = clean_for_match(df_calc.iloc[:, IDX_M_BAR])
                df_calc['_MATCH_CODE'] = clean_for_match(df_calc.iloc[:, IDX_M_CODE])
                df_calc['_VAL_PROFIT'] = clean_num(df_calc.iloc[:, IDX_M_PROFIT])
                df_calc['_VAL_COST'] = clean_num(df_calc.iloc[:, IDX_M_COST])

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

                # --- Step 5: 关联 & 计算 ---
                df_final = pd.merge(df_calc, sales_agg, on='_MATCH_SKU', how='left', sort=False)
                df_final['SKU销量'] = df_final['SKU销量'].fillna(0).astype(int)
                
                df_final = pd.merge(df_final, inv_agg, on='_MATCH_SKU', how='left', sort=False)
                df_final['火箭仓库存'] = df_final['火箭仓库存'].fillna(0).astype(int)
                
                df_final = pd.merge(df_final, inv_j_agg, on='_MATCH_BAR', how='left', sort=False)
                df_final['极风库存'] = df_final['极风库存'].fillna(0).astype(int)

                # 5.3 利润
                df_final['P列_SKU总毛利'] = df_final['SKU销量'] * df_final['_VAL_PROFIT']
                df_final['Q列_产品总利润'] = df_final.groupby('_MATCH_CODE', sort=False)['P列_SKU总毛利'].transform('sum')
                df_final['产品总销量'] = df_final.groupby('_MATCH_CODE', sort=False)['SKU销量'].transform('sum')
                
                # 5.4 广告
                df_final = pd.merge(df_final, ads_agg, on='_MATCH_CODE', how='left', sort=False)
                df_final['R列_产品总广告费'] = df_final['R列_产品总广告费'].fillna(0)
                df_final['产品广告销量'] = df_final['产品广告销量'].fillna(0)
                
                # 5.5 净利
                df_final['S列_最终净利润'] = df_final['Q列_产品总利润'] - df_final['R列_产品总广告费']

                # --- Step 6: 报表生成 ---
                
                # Sheet2: 业务报表
                df_final['产品_火箭仓库存'] = df_final.groupby('_MATCH_CODE', sort=False)['火箭仓库存'].transform('sum')
                df_final['产品_极风库存'] = df_final.groupby('_MATCH_CODE', sort=False)['极风库存'].transform('sum')
                df_final['产品_总库存'] = df_final['产品_火箭仓库存'] + df_final['产品_极风库存']

                df_sheet2 = df_final[[col_code_name, 'Q列_产品总利润', 'R列_产品总广告费', 'S列_最终净利润', '产品总销量', '产品广告销量', '产品_火箭仓库存', '产品_极风库存', '产品_总库存']].copy()
                df_sheet2 = df_sheet2.drop_duplicates(subset=[col_code_name], keep='first')
                
                df_sheet2.rename(columns={
                    '产品_火箭仓库存': '火箭仓库存', 
                    '产品_极风库存': '极风库存',
                    '产品_总库存': '总库存'
                }, inplace=True)

                df_sheet2['广告/毛利比'] = df_sheet2.apply(
                    lambda x: x['R列_产品总广告费'] / x['Q列_产品总利润'] if x['Q列_产品总利润'] != 0 else 0, axis=1
                )
                df_sheet2['自然销量'] = df_sheet2['产品总销量'] - df_sheet2['产品广告销量']
                df_sheet2['自然销量占比'] = df_sheet2.apply(
                    lambda x: x['自然销量'] / x['产品总销量'] if x['产品总销量'] != 0 else 0, axis=1
                )
                
                cols_order_s2 = [
                    col_code_name, 'Q列_产品总利润', 'R列_产品总广告费', 'S列_最终净利润', 
                    '广告/毛利比', '产品总销量', '产品广告销量', '自然销量', '自然销量占比',
                    '火箭仓库存', '极风库存', '总库存'
                ]
                df_sheet2 = df_sheet2[cols_order_s2]

                # --- Step 7: 库存分析表 (Sheet3) ---
                
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

                # --- Step 8: 深度清洗与重命名 (关键修改) ---
                cols_master_AM = df_master.columns[:13].tolist()
                
                # 构造 Sheet1 列
                cols_s1_final = cols_master_AM + [
                    'SKU销量', 'P列_SKU总毛利', 'Q列_产品总利润', 'R列_产品总广告费', 'S列_最终净利润'
                ]
                df_final_clean = df_final[cols_s1_final].copy()

                # 构造 Sheet3 列
                cols_inv_final = cols_master_AM + [
                    '火箭仓库存数量', '极风库存', '总库存', 
                    '库存货值', '滞销库存货值', 
                    '待补数量', 
                    'SKU销量', '安全库存', '冗余标准'
                ]
                df_sheet3 = df_final[cols_inv_final].copy()

                # 【核心】统一重命名：去除 P/Q/R/S 列前缀
                rename_dict = {
                    'P列_SKU总毛利': 'SKU总毛利',
                    'Q列_产品总利润': '产品总利润',
                    'R列_产品总广告费': '产品总广告费',
                    'S列_最终净利润': '最终净利润'
                }
                
                df_final_clean.rename(columns=rename_dict, inplace=True)
                df_sheet2.rename(columns=rename_dict, inplace=True)
                # df_sheet3 不需要重命名，因为它用的都是直接的字段名

                # ==========================================
                # 🔍 Step 9: 执行筛选
                # ==========================================
                
                if filter_code:
                    df_final_clean = df_final_clean[df_final_clean[col_code_name].astype(str).str.contains(filter_code, na=False)]
                    df_sheet2 = df_sheet2[df_sheet2[col_code_name].astype(str).str.contains(filter_code, na=False)]
                    df_sheet3 = df_sheet3[df_sheet3[col_code_name].astype(str).str.contains(filter_code, na=False)]

                if filter_profit == "只看盈利 (>0)":
                    # 使用 df_final 的原始索引进行同步
                    mask = df_final['S列_最终净利润'] > 0
                    # 注意：df_final_clean 已经改名了，所以用 '最终净利润'
                    df_final_clean = df_final_clean[df_final_clean['最终净利润'] > 0]
                    # df_sheet3 还没改名，但行是对应的，可以用原始 mask 筛选吗？
                    # 更安全的做法：重新计算 mask 或者根据 df_final_clean 的 index 筛选
                    valid_indices = df_final_clean.index
                    df_sheet3 = df_sheet3.loc[df_sheet3.index.isin(valid_indices)]
                    
                    df_sheet2 = df_sheet2[df_sheet2['最终净利润'] > 0]
                    
                elif filter_profit == "只看亏损 (<0)":
                    df_final_clean = df_final_clean[df_final_clean['最终净利润'] < 0]
                    valid_indices = df_final_clean.index
                    df_sheet3 = df_sheet3.loc[df_sheet3.index.isin(valid_indices)]
                    
                    df_sheet2 = df_sheet2[df_sheet2['最终净利润'] < 0]

                # ==========================================
                # 🔥 看板展示
                # ==========================================
                
                if df_sheet2.empty:
                    st.warning(f"⚠️ 筛选结果为空。")
                else:
                    total_qty = df_sheet2['产品总销量'].sum()
                    net_profit = df_sheet2['最终净利润'].sum() # 已改名
                    inv_value_total = df_sheet3['库存货值'].sum()
                    dead_stock_value = df_sheet3['滞销库存货值'].sum()
                    total_restock = df_sheet3['待补数量'].sum()
                    
                    title_suffix = ""
                    if filter_code: title_suffix += f" [编号:{filter_code}]"
                    if filter_profit != "全部显示": title_suffix += f" [{filter_profit}]"
                    
                    st.subheader(f"📈 经营概览 {title_suffix}")
                    k1, k2, k3, k4, k5 = st.columns(5)
                    profit_color = "normal" if net_profit > 0 else "inverse"
                    
                    k1.metric("💰 最终净利润", f"{net_profit:,.0f}", delta_color=profit_color)
                    k2.metric("📦 总销售数量", f"{total_qty:,.0f}") 
                    k3.metric("🏭 库存总货值", f"¥ {inv_value_total:,.0f}")
                    k4.metric("🔴 滞销资金占用", f"¥ {dead_stock_value:,.0f}", delta="需重点清理", delta_color="inverse")
                    k5.metric("🚨 建议补货量", f"{total_restock:,.0f}")

                    st.divider()

                    tab1, tab2, tab3 = st.tabs(["📝 1. 利润分析", "📊 2. 业务报表", "🏭 3. 库存分析"])
                    
                    # === 格式化函数 ===
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
                            if any(x in c_str for x in ['利润', '费用', '货值', '金额', '毛利', '销量', '库存', '数量', '标准', '待补']):
                                if '率' not in c_str and '比' not in c_str:
                                    format_dict[col] = safe_fmt_int
                            elif any(x in c_str for x in ['比', '率', '占比']):
                                format_dict[col] = safe_fmt_pct
                        return format_dict

                    def apply_visual_style(df, cols_to_color, is_sheet2=False):
                        try:
                            styler = df.style.format(get_format_dict(df))
                            def zebra_rows(x):
                                codes = x.iloc[:, 0].astype(str)
                                groups = (codes != codes.shift()).cumsum()
                                is_odd = groups % 2 != 0
                                styles = pd.DataFrame('', index=x.index, columns=x.columns)
                                styles.loc[is_odd, :] = 'background-color: #f0f2f6' 
                                return styles
                            styler = styler.apply(zebra_rows, axis=None)
                            # 确保列存在再应用高亮
                            valid_cols = [c for c in cols_to_color if c in df.columns]
                            if valid_cols:
                                styler = styler.background_gradient(subset=valid_cols, cmap='RdYlGn', vmin=-10000, vmax=10000)
                            return styler
                        except: return df
                    
                    def apply_inventory_style(df):
                        try:
                            styler = df.style.format(get_format_dict(df))
                            def zebra_rows(x):
                                codes = x.iloc[:, 0].astype(str)
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

                    with tab1:
                        st.caption("利润明细 (Sheet1)")
                        # 注意：列名已改，这里传新列名
                        st.dataframe(apply_visual_style(df_final_clean, ['最终净利润']), use_container_width=True, height=600)
                    
                    with tab2:
                        st.caption("业务汇总 (Sheet2)")
                        st.dataframe(apply_visual_style(df_sheet2, ['最终净利润'], is_sheet2=True), use_container_width=True, height=600)
                    
                    with tab3:
                        st.caption("库存分析 (Sheet3)")
                        try:
                            st_inv = apply_inventory_style(df_sheet3)
                            st_inv = st_inv.bar(subset=['总库存'], color='#800080')\
                                           .bar(subset=['库存货值'], color='#2ca02c')\
                                           .bar(subset=['滞销库存货值'], color='#880e4f')
                            st.dataframe(st_inv, use_container_width=True, height=600)
                        except:
                            st.dataframe(df_sheet3, use_container_width=True)

                    # ==========================================
                    # 📥 下载逻辑 (Excel 格式精细化)
                    # ==========================================
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        df_final_clean.to_excel(writer, index=False, sheet_name='利润分析')
                        df_sheet2.to_excel(writer, index=False, sheet_name='业务报表')
                        df_sheet3.to_excel(writer, index=False, sheet_name='库存分析')
                        
                        wb = writer.book
                        # 表头样式：加粗、蓝底、白字、边框
                        fmt_header = wb.add_format({
                            'bold': True, 
                            'bg_color': '#4472C4', 
                            'font_color': 'white', 
                            'border': 1, 
                            'align': 'center',
                            'valign': 'vcenter'
                        })
                        
                        fmt_int = wb.add_format({'num_format': '#,##0', 'align': 'center'})
                        fmt_pct = wb.add_format({'num_format': '0.0%', 'align': 'center'})
                        
                        base_font = {'font_name': 'Microsoft YaHei', 'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter'}
                        fmt_grey = wb.add_format(dict(base_font, bg_color='#BFBFBF'))
                        fmt_white = wb.add_format(dict(base_font, bg_color='#FFFFFF'))

                        def set_sheet_format(sheet_name, df_obj, group_col_idx):
                            ws = writer.sheets[sheet_name]
                            raw_codes = df_obj.iloc[:, group_col_idx].astype(str).tolist()
                            clean_codes = [str(x).replace('.0','').replace('"','').strip().upper() for x in raw_codes]
                            is_grey = False
                            for i in range(len(raw_codes)):
                                if i > 0 and clean_codes[i] != clean_codes[i-1]:
                                    is_grey = not is_grey
                                ws.set_row(i + 1, None, fmt_grey if is_grey else fmt_white)

                            for i, col in enumerate(df_obj.columns):
                                c_str = str(col)
                                width = 12
                                cell_fmt = None
                                if any(x in c_str for x in ['利润', '费用', '货值', '金额', '毛利', '销量', '库存', '数量', '标准', '待补']):
                                    if '率' not in c_str and '比' not in c_str:
                                        cell_fmt = fmt_int
                                        width = 15
                                elif any(x in c_str for x in ['比', '率', '占比']):
                                    cell_fmt = fmt_pct
                                    width = 12
                                
                                if cell_fmt:
                                    ws.set_column(i, i, width, cell_fmt)
                                else:
                                    ws.set_column(i, i, width)
                                
                                # 强制写入表头（覆盖默认）
                                ws.write(0, i, col, fmt_header)

                        set_sheet_format('利润分析', df_final_clean, IDX_M_CODE)
                        set_sheet_format('业务报表', df_sheet2, IDX_M_CODE)
                        set_sheet_format('库存分析', df_sheet3, IDX_M_CODE)

                    st.divider()
                    st.success(f"✅ 报表生成完毕！")
                    
                    st.download_button(
                        label="📥 下载 Excel (含利润/业务/库存 3个Sheet)",
                        data=output.getvalue(),
                        file_name=f"Coupang_Report_Final_{filter_code if filter_code else 'All'}.xlsx",
                        mime="application/vnd.ms-excel",
                        type="primary",
                        use_container_width=True
                    )

        except Exception as e:
            st.error(f"❌ 运行出错: {e}")
else:
    st.info("👈 请上传文件")
