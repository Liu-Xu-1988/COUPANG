import streamlit as st
import pandas as pd
import io
import re

# ==========================================
# 1. 页面配置 (宽屏)
# ==========================================
st.set_page_config(layout="wide", page_title="Coupang 经营看板 Pro (最终版)")
st.title("📊 Coupang 经营分析看板 (最终版·精简交互)")

# --- 列号配置 ---
IDX_M_CODE   = 0    # A列
IDX_M_SKU    = 3    # D列
IDX_M_COST   = 6    # G列
IDX_M_PROFIT = 10   # K列
IDX_M_BAR    = 12   # M列

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
    # 【修改点】移除了模式选择，仅保留高度调节
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
                # --- Step 1-5: 数据处理 ---
                df_master = read_file_strict(file_master)
                col_code_name = df_master.columns[IDX_M_CODE]

                df_calc = df_master.copy()
                df_calc['_MATCH_SKU'] = clean_for_match(df_calc.iloc[:, IDX_M_SKU])
                df_calc['_MATCH_BAR'] = clean_for_match(df_calc.iloc[:, IDX_M_BAR])
                df_calc['_MATCH_CODE'] = clean_for_match(df_calc.iloc[:, IDX_M_CODE])
                df_calc['_VAL_PROFIT'] = clean_num(df_calc.iloc[:, IDX_M_PROFIT])
                df_calc['_VAL_COST'] = clean_num(df_calc.iloc[:, IDX_M_COST])

                sales_list = [read_file_strict(f) for f in files_sales]
                df_sales_all = pd.concat(sales_list, ignore_index=True)
                df_sales_all['_MATCH_SKU'] = clean_for_match(df_sales_all.iloc[:, IDX_S_ID])
                df_sales_all['销量'] = clean_num(df_sales_all.iloc[:, IDX_S_QTY])
                sales_agg = df_sales_all.groupby('_MATCH_SKU')['销量'].sum().reset_index()
                sales_agg.rename(columns={'销量': 'SKU销量'}, inplace=True) 

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
                df_final['R列_产品总广告费'] = df_final['R列_产品总广告费'].fillna(0).round(0).astype(int)
                df_final['产品广告销量'] = df_final['产品广告销量'].fillna(0)
                df_final['S列_最终净利润'] = df_final['Q列_产品总利润'] - df_final['R列_产品总广告费']

                # --- Step 6: 报表构造 ---
                
                # Sheet2 (业务报表 - 产品维度)
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

                df_sheet2['广告费占比'] = df_sheet2.apply(
                    lambda x: x['R列_产品总广告费'] / x['Q列_产品总利润'] if x['Q列_产品总利润'] != 0 else 0, axis=1
                )
                df_sheet2['自然销量'] = df_sheet2['产品总销量'] - df_sheet2['产品广告销量']
                df_sheet2['自然销量占比'] = df_sheet2.apply(
                    lambda x: x['自然销量'] / x['产品总销量'] if x['产品总销量'] != 0 else 0, axis=1
                )
                
                cols_order_s2 = [
                    col_code_name, 'Q列_产品总利润', 'R列_产品总广告费', 'S列_最终净利润', 
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
