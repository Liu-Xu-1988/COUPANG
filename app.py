import streamlit as st
import pandas as pd
import io
import re

# ==========================================
# 1. 页面配置 (宽屏)
# ==========================================
st.set_page_config(layout="wide", page_title="Coupang 经营看板 Pro (含库存)")
st.title("📊 Coupang 经营分析看板 (全功能版+库存)")

# --- 列号配置 ---
# Master表 (基础表)
IDX_M_CODE   = 0    # A列: 内部编码
IDX_M_SKU    = 3    # D列: SKU ID (用于匹配库存)
IDX_M_PROFIT = 10   # K列: 单品毛利

# Sales表 (销售表)
IDX_S_ID     = 0    # A列
IDX_S_QTY    = 8    # I列

# Ads表 (广告表)
IDX_A_CAMPAIGN = 5  # F列
IDX_A_GROUP    = 6  # G列
IDX_A_SPEND    = 15 # P列
IDX_A_SALES    = 29 # AD列 (30列)

# Inventory表 (库存表 - 新增!)
IDX_I_ID     = 2    # C列: ID (用于匹配基础表D列)
IDX_I_QTY    = 7    # H列: 库存数量
# -----------------

# ==========================================
# 2. 侧边栏上传
# ==========================================
with st.sidebar:
    st.header("📂 数据源上传")
    st.info("请按顺序上传以下文件：")
    
    file_master = st.file_uploader("1. 基础信息表 (Master)", type=['csv', 'xlsx', 'xlsm'])
    files_sales = st.file_uploader("2. 销售表 (Sales)", type=['csv', 'xlsx', 'xlsm'], accept_multiple_files=True)
    files_ads = st.file_uploader("3. 广告表 (Ads)", type=['csv', 'xlsx', 'xlsm'], accept_multiple_files=True)
    # 新增库存上传
    files_inv = st.file_uploader("4. 库存信息表 (Inventory)", type=['csv', 'xlsx', 'xlsm'], accept_multiple_files=True)

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
# 只要前三个表还在，就可以跑主流程，库存表是可选的（但为了完整性最好都有）
if file_master and files_sales and files_ads:
    st.divider()
    
    if st.button("🚀 生成全套报表 (含库存)", type="primary", use_container_width=True):
        try:
            with st.spinner("正在全速处理数据..."):
                
                # --- Step 1: 基础表 ---
                df_master = read_file_strict(file_master)
                col_code_name = df_master.columns[IDX_M_CODE]

                df_master['_MATCH_SKU'] = clean_for_match(df_master.iloc[:, IDX_M_SKU])
                df_master['_MATCH_CODE'] = clean_for_match(df_master.iloc[:, IDX_M_CODE])
                df_master['_VAL_PROFIT'] = clean_num(df_master.iloc[:, IDX_M_PROFIT])

                # --- Step 2: 销售表 ---
                sales_list = [read_file_strict(f) for f in files_sales]
                df_sales_all = pd.concat(sales_list, ignore_index=True)
                
                df_sales_all['_MATCH_SKU'] = clean_for_match(df_sales_all.iloc[:, IDX_S_ID])
                df_sales_all['销量'] = clean_num(df_sales_all.iloc[:, IDX_S_QTY])
                
                sales_agg = df_sales_all.groupby('_MATCH_SKU')['销量'].sum().reset_index()
                sales_agg.rename(columns={'销量': 'O列_合并销量'}, inplace=True)

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

                # --- Step 4: 库存表处理 (新增!) ---
                if files_inv:
                    inv_list = [read_file_strict(f) for f in files_inv]
                    df_inv_all = pd.concat(inv_list, ignore_index=True)
                    
                    # 匹配逻辑：库存C列(ID) 对 基础表D列(SKU)
                    df_inv_all['_MATCH_SKU'] = clean_for_match(df_inv_all.iloc[:, IDX_I_ID])
                    df_inv_all['库存数量'] = clean_num(df_inv_all.iloc[:, IDX_I_QTY])
                    
                    # 聚合库存 (以防同一个SKU在多行出现)
                    inv_agg = df_inv_all.groupby('_MATCH_SKU')['库存数量'].sum().reset_index()
                else:
                    # 如果没上传库存表，就给空表
                    inv_agg = pd.DataFrame(columns=['_MATCH_SKU', '库存数量'])

                # --- Step 5: 关联 & 计算 ---
                # 5.1 基础 + 销售
                df_final = pd.merge(df_master, sales_agg, on='_MATCH_SKU', how='left', sort=False)
                df_final['O列_合并销量'] = df_final['O列_合并销量'].fillna(0).astype(int)
                
                # 5.2 关联库存 (新增!) -> 这个用于生成 Sheet3
                df_final = pd.merge(df_final, inv_agg, on='_MATCH_SKU', how='left', sort=False)
                df_final['库存数量'] = df_final['库存数量'].fillna(0).astype(int)

                # 5.3 利润计算
                df_final['P列_SKU总毛利'] = df_final['O列_合并销量'] * df_final['_VAL_PROFIT']
                df_final['Q列_产品总利润'] = df_final.groupby('_MATCH_CODE', sort=False)['P列_SKU总毛利'].transform('sum')
                df_final['产品总销量'] = df_final.groupby('_MATCH_CODE', sort=False)['O列_合并销量'].transform('sum')
                
                # 5.4 关联广告
                df_final = pd.merge(df_final, ads_agg, on='_MATCH_CODE', how='left', sort=False)
                df_final['R列_产品总广告费'] = df_final['R列_产品总广告费'].fillna(0)
                df_final['产品广告销量'] = df_final['产品广告销量'].fillna(0)
                
                # 5.5 净利计算
                df_final['S列_最终净利润'] = df_final['Q列_产品总利润'] - df_final['R列_产品总广告费']

                # --- Step 6: 报表数据生成 ---
                
                # Sheet2: 业务报表
                df_sheet2 = df_final[[col_code_name, 'Q列_产品总利润', 'R列_产品总广告费', 'S列_最终净利润', '产品总销量', '产品广告销量']].copy()
                df_sheet2 = df_sheet2.drop_duplicates(subset=[col_code_name], keep='first')
                
                df_sheet2['广告/毛利比'] = df_sheet2.apply(
                    lambda x: x['R列_产品总广告费'] / x['Q列_产品总利润'] if x['Q列_产品总利润'] != 0 else 0, axis=1
                )
                df_sheet2['自然销量'] = df_sheet2['产品总销量'] - df_sheet2['产品广告销量']
                df_sheet2['自然销量占比'] = df_sheet2.apply(
                    lambda x: x['自然销量'] / x['产品总销量'] if x['产品总销量'] != 0 else 0, axis=1
                )
                
                # Sheet2 列顺序
                cols_order_s2 = [
                    col_code_name, 'Q列_产品总利润', 'R列_产品总广告费', 'S列_最终净利润', 
                    '广告/毛利比', '产品总销量', '产品广告销量', '自然销量', '自然销量占比'
                ]
                df_sheet2 = df_sheet2[cols_order_s2]

                # Sheet3: 库存分析 (新增!)
                # 逻辑：A-M列 (即 Master 的前13列) + N列 (库存数量)
                # 确保 df_final 保留了 Master 的列顺序
                # 我们取 df_final 的前13列 (0-12) 和 '库存数量'
                cols_master_AM = df_final.columns[:13].tolist() 
                df_sheet3 = df_final[cols_master_AM + ['库存数量']].copy()
                # 确保 Sheet3 也按 A-M 列去重？不，Profit Analysis 是明细表，Inventory Analysis 应该也是明细表
                # 既然是 "格式和利润分析完全一样"，那么应该保留 SKU 级明细

                # --- Step 7: 清理辅助列 ---
                cols_to_drop = [c for c in df_final.columns if str(c).startswith('_') or str(c).startswith('Code_')]
                df_final.drop(columns=cols_to_drop, inplace=True)

                # ==========================================
                # 🔥 看板展示
                # ==========================================
                
                total_qty = df_sheet2['产品总销量'].sum()
                net_profit = df_sheet2['S列_最终净利润'].sum()
                total_inv = df_final['库存数量'].sum() # 总库存
                
                st.subheader("📈 经营概览")
                k1, k2, k3, k4 = st.columns(4)
                k1.metric("💰 最终净利润", f"{net_profit:,.0f}")
                k2.metric("📦 总销售数量", f"{total_qty:,.0f}") 
                k3.metric("🏭 总库存量", f"{total_inv:,.0f}")
                k4.metric("📊 动销率", f"{(total_qty/total_inv if total_inv else 0):.1%}")

                st.divider()

                tab1, tab2, tab3 = st.tabs(["📝 1. 利润分析 (明细)", "📊 2. 业务报表 (汇总)", "🏭 3. 库存分析 (新增)"])
                
                def try_style(df, cols, is_sheet2=False):
                    try:
                        styler = df.style.format(precision=0)
                        if is_sheet2:
                            styler = styler.format({
                                '广告/毛利比': '{:.1%}', '自然销量占比': '{:.1%}',
                                '产品总销量': '{:,.0f}', '产品广告销量': '{:,.0f}', '自然销量': '{:,.0f}'
                            })
                        return styler.background_gradient(subset=cols, cmap='RdYlGn', vmin=-10000, vmax=10000)
                    except: return df

                with tab1:
                    st.caption("利润明细 (Sheet1)")
                    st.dataframe(try_style(df_final, ['S列_最终净利润']), use_container_width=True, height=600)
                
                with tab2:
                    st.caption("业务汇总 (Sheet2)")
                    st.dataframe(try_style(df_sheet2, ['S列_最终净利润'], is_sheet2=True), use_container_width=True, height=600)
                
                with tab3:
                    st.caption("库存分析 (Sheet3) - 结构与利润分析一致，N列为库存")
                    # 库存不一定需要红绿配色，用蓝色条显示数量
                    try:
                        st.dataframe(
                            df_sheet3.style.format(precision=0).bar(subset=['库存数量'], color='#5fba7d'),
                            use_container_width=True, height=600
                        )
                    except:
                        st.dataframe(df_sheet3, use_container_width=True)

                # ==========================================
                # 📥 下载逻辑 (3个 Sheet)
                # ==========================================
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='利润分析')
                    df_sheet2.to_excel(writer, index=False, sheet_name='业务报表')
                    df_sheet3.to_excel(writer, index=False, sheet_name='库存分析')
                    
                    # 通用样式
                    wb = writer.book
                    fmt_header = wb.add_format({'bold': True, 'bg_color': '#4472C4', 'font_color': 'white', 'border': 1, 'align': 'center'})
                    fmt_money = wb.add_format({'num_format': '#,##0', 'align': 'center'})
                    fmt_pct = wb.add_format({'num_format': '0.0%', 'align': 'center'})
                    
                    # --- Sheet1 & Sheet3 斑马纹样式 (A-M列通用) ---
                    # 定义斑马纹格式
                    base_font = {'font_name': 'Microsoft YaHei', 'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter'}
                    fmt_grey = wb.add_format(dict(base_font, bg_color='#BFBFBF'))
                    fmt_white = wb.add_format(dict(base_font, bg_color='#FFFFFF'))

                    # 辅助函数：应用斑马纹
                    def apply_zebra(sheet_name, df_obj, target_col_idx_for_group=0):
                        ws = writer.sheets[sheet_name]
                        # 自动列宽
                        for i, col in enumerate(df_obj.columns):
                            str_len = max(df_obj[col].astype(str).map(len).max(), len(str(col))) * 1.5
                            ws.set_column(i, i, min(max(str_len, 10), 40))
                        
                        # 斑马纹逻辑
                        raw_codes = df_obj.iloc[:, target_col_idx_for_group].astype(str).tolist()
                        clean_codes = [str(x).replace('.0','').replace('"','').strip().upper() for x in raw_codes]
                        is_grey = False
                        for i in range(len(raw_codes)):
                            if i > 0 and clean_codes[i] != clean_codes
