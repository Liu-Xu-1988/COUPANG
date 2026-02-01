import streamlit as st
import pandas as pd
import io
import re

# ==========================================
# 1. 页面配置
# ==========================================
st.set_page_config(layout="wide", page_title="Coupang 经营看板 Pro (最终修复版)")
st.title("📊 Coupang 经营分析看板 (库存逻辑修复版)")
st.markdown("""
### 🔧 修复说明：
1. **库存汇总修复**：业务报表 (Sheet2) 现已正确包含该产品下**所有 SKU 的库存总和**，不再只取第一行。
2. **条码匹配增强**：增加抗科学计数法处理，解决极风库存匹配失败的问题。
3. **数据自检**：如果库存显示为 0，请查看下方的“数据匹配自检”区域。
""")

# --- 列号配置 ---
# Master表 (基础表)
IDX_M_CODE    = 0    # A列: 内部编码
IDX_M_SKU     = 3    # D列: SKU ID (匹配火箭仓)
IDX_M_PROFIT  = 10   # K列: 单品毛利
IDX_M_BARCODE = 12   # M列: ID号码 (匹配极风库存)

# Sales表 (销售表)
IDX_S_ID      = 0    # A列
IDX_S_QTY     = 8    # I列

# Ads表 (广告表)
IDX_A_CAMPAIGN = 5  # F列
IDX_A_GROUP    = 6  # G列
IDX_A_SPEND    = 15 # P列
IDX_A_SALES    = 29 # AD列

# Inventory Rocket (火箭仓)
IDX_I_R_ID    = 2    # C列: ID
IDX_I_R_QTY   = 7    # H列: 库存

# Inventory Jifeng (极风)
IDX_I_J_BAR   = 2    # C列: 产品条码
IDX_I_J_QTY   = 10   # K列: 数值
# -----------------

# ==========================================
# 2. 侧边栏上传
# ==========================================
with st.sidebar:
    st.header("📂 数据源上传")
    st.info("请按顺序上传文件：")
    
    file_master = st.file_uploader("1. 基础信息表 (Master)", type=['csv', 'xlsx', 'xlsm'])
    files_sales = st.file_uploader("2. 销售表 (Sales)", type=['csv', 'xlsx', 'xlsm'], accept_multiple_files=True)
    files_ads = st.file_uploader("3. 广告表 (Ads)", type=['csv', 'xlsx', 'xlsm'], accept_multiple_files=True)
    files_inv_r = st.file_uploader("4. 火箭仓库存表 (Rocket)", type=['csv', 'xlsx', 'xlsm'], accept_multiple_files=True)
    files_inv_j = st.file_uploader("5. 极风库存表 (Jifeng)", type=['csv', 'xlsx', 'xlsm'], accept_multiple_files=True)

# ==========================================
# 3. 清洗工具 (增强版)
# ==========================================
def clean_for_match(series):
    """通用清洗：去空格、转大写、去.0"""
    return series.astype(str).str.replace(r'\.0$', '', regex=True).str.replace('"', '').str.strip().str.upper()

def clean_barcode(series):
    """条码专用清洗：防止科学计数法 (8.8E+12 -> 88...)"""
    # 先转为字符串，如果包含 'E+' 则尝试还原
    def _fix_sci(x):
        try:
            s = str(x).upper()
            if 'E' in s:
                return str(int(float(x))) # 尝试还原科学计数法
            return s.replace('.0', '').strip()
        except:
            return str(x).strip()
    
    return series.apply(_fix_sci)

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
    
    if st.button("🚀 开始修复计算", type="primary", use_container_width=True):
        try:
            with st.spinner("正在执行增强匹配逻辑..."):
                
                # --- Step 1: 基础表 ---
                df_master = read_file_strict(file_master)
                col_code_name = df_master.columns[IDX_M_CODE]

                # 关键：生成匹配键
                df_master['_MATCH_SKU'] = clean_for_match(df_master.iloc[:, IDX_M_SKU])
                # 条码清洗：使用增强版清洗函数
                df_master['_MATCH_BAR'] = clean_barcode(df_master.iloc[:, IDX_M_BARCODE])
                
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

                # --- Step 4.1: 火箭仓库存 (Rocket) ---
                if files_inv_r:
                    inv_r_list = [read_file_strict(f) for f in files_inv_r]
                    df_inv_r = pd.concat(inv_r_list, ignore_index=True)
                    # 匹配逻辑：C列 ID -> 基础表 SKU (D列)
                    df_inv_r['_MATCH_SKU'] = clean_for_match(df_inv_r.iloc[:, IDX_I_R_ID])
                    df_inv_r['火箭仓库存'] = clean_num(df_inv_r.iloc[:, IDX_I_R_QTY])
                    inv_r_agg = df_inv_r.groupby('_MATCH_SKU')['火箭仓库存'].sum().reset_index()
                else:
                    inv_r_agg = pd.DataFrame(columns=['_MATCH_SKU', '火箭仓库存'])

                # --- Step 4.2: 极风库存 (Jifeng) ---
                if files_inv_j:
                    inv_j_list = [read_file_strict(f) for f in files_inv_j]
                    df_inv_j = pd.concat(inv_j_list, ignore_index=True)
                    # 匹配逻辑：C列 条码 -> 基础表 ID号码 (M列)
                    # 使用增强清洗
                    df_inv_j['_MATCH_BAR'] = clean_barcode(df_inv_j.iloc[:, IDX_I_J_BAR])
                    df_inv_j['极风库存'] = clean_num(df_inv_j.iloc[:, IDX_I_J_QTY])
                    inv_j_agg = df_inv_j.groupby('_MATCH_BAR')['极风库存'].sum().reset_index()
                else:
                    inv_j_agg = pd.DataFrame(columns=['_MATCH_BAR', '极风库存'])

                # --- 🔍 调试自检区 ---
                with st.expander("🛠️ 数据匹配自检 (如果库存为0请看这里)", expanded=False):
                    st.write(f"**基础表行数**: {len(df_master)}")
                    
                    c1, c2 = st.columns(2)
                    with c1:
                        st.write(f"**🚀 火箭仓匹配状态**")
                        if not inv_r_agg.empty:
                            st.write(f"提取到库存记录: {len(inv_r_agg)} 条")
                            st.write("样例数据 (SKU | 库存):")
                            st.dataframe(inv_r_agg.head(3), height=100)
                        else:
                            st.error("未提取到火箭仓数据，请检查上传文件或C列/H列位置。")
                            
                    with c2:
                        st.write(f"**🌪️ 极风库存匹配状态**")
                        if not inv_j_agg.empty:
                            st.write(f"提取到库存记录: {len(inv_j_agg)} 条")
                            st.write("样例数据 (条码 | 库存):")
                            st.dataframe(inv_j_agg.head(3), height=100)
                        else:
                            st.error("未提取到极风数据，请检查上传文件或C列/K列位置。")

                # --- Step 5: 关联 & 计算 ---
                # 5.1 基础 + 销售
                df_final = pd.merge(df_master, sales_agg, on='_MATCH_SKU', how='left', sort=False)
                df_final['O列_合并销量'] = df_final['O列_合并销量'].fillna(0).astype(int)
                
                # 5.2 关联 火箭仓库存 (SKU级)
                df_final = pd.merge(df_final, inv_r_agg, on='_MATCH_SKU', how='left', sort=False)
                df_final['火箭仓库存'] = df_final['火箭仓库存'].fillna(0).astype(int)
                
                # 5.3 关联 极风库存 (条码级)
                df_final = pd.merge(df_final, inv_j_agg, on='_MATCH_BAR', how='left', sort=False)
                df_final['极风库存'] = df_final['极风库存'].fillna(0).astype(int)

                # 5.4 利润与汇总计算
                df_final['P列_SKU总毛利'] = df_final['O列_合并销量'] * df_final['_VAL_PROFIT']
                
                # === 关键修复：计算【产品级别】的汇总数据 (解决Sheet2库存为0的问题) ===
                df_final['Q列_产品总利润'] = df_final.groupby('_MATCH_CODE', sort=False)['P列_SKU总毛利'].transform('sum')
                df_final['产品总销量'] = df_final.groupby('_MATCH_CODE', sort=False)['O列_合并销量'].transform('sum')
                
                # 新增：汇总库存到产品级
                df_final['产品_火箭仓库存'] = df_final.groupby('_MATCH_CODE', sort=False)['火箭仓库存'].transform('sum')
                df_final['产品_极风库存'] = df_final.groupby('_MATCH_CODE', sort=False)['极风库存'].transform('sum')
                # ==============================================================

                # 5.5 关联广告
                df_final = pd.merge(df_final, ads_agg, on='_MATCH_CODE', how='left', sort=False)
                df_final['R列_产品总广告费'] = df_final['R列_产品总广告费'].fillna(0)
                df_final['产品广告销量'] = df_final['产品广告销量'].fillna(0)
                
                # 5.6 净利计算
                df_final['S列_最终净利润'] = df_final['Q列_产品总利润'] - df_final['R列_产品总广告费']

                # --- Step 6: 报表生成 ---
                
                # Sheet2: 业务报表 (使用汇总后的列)
                # 注意：这里我们取 '产品_火箭仓库存' 而不是 '火箭仓库存'
                df_sheet2 = df_final[[col_code_name, 'Q列_产品总利润', 'R列_产品总广告费', 'S列_最终净利润', '产品总销量', '产品广告销量', '产品_火箭仓库存', '产品_极风库存']].copy()
                df_sheet2 = df_sheet2.drop_duplicates(subset=[col_code_name], keep='first')
                
                # 重命名以便展示
                df_sheet2.rename(columns={'产品_火箭仓库存': '火箭仓库存', '产品_极风库存': '极风库存'}, inplace=True)

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
                    '广告/毛利比', '产品总销量', '产品广告销量', '自然销量', '自然销量占比',
                    '火箭仓库存', '极风库存'
                ]
                df_sheet2 = df_sheet2[cols_order_s2]

                # Sheet3: 库存分析 (保持SKU级明细)
                cols_master_AM = df_final.columns[:13].tolist() 
                df_sheet3 = df_final[cols_master_AM + ['火箭仓库存', '极风库存']].copy()
                df_sheet3.rename(columns={'火箭仓库存': '火箭仓库存数量'}, inplace=True)

                # --- Step 7: 清理 ---
                cols_to_drop = [c for c in df_final.columns if str(c).startswith('_') or str(c).startswith('Code_') or c.startswith('产品_')]
                df_final.drop(columns=cols_to_drop, inplace=True)

                # ==========================================
                # 🔥 看板展示
                # ==========================================
                
                total_qty = df_sheet2['产品总销量'].sum()
                net_profit = df_sheet2['S列_最终净利润'].sum()
                inv_rocket = df_sheet2['火箭仓库存'].sum()
                inv_jifeng = df_sheet2['极风库存'].sum()
                total_inv = inv_rocket + inv_jifeng
                
                st.subheader("📈 经营概览")
                k1, k2, k3, k4 = st.columns(4)
                k1.metric("💰 最终净利润", f"{net_profit:,.0f}")
                k2.metric("📦 总销售数量", f"{total_qty:,.0f}") 
                k3.metric("🏭 总库存", f"{total_inv:,.0f}", help=f"火箭仓:{inv_rocket:,.0f} | 极风:{inv_jifeng:,.0f}")
                k4.metric("📊 整体动销率", f"{(total_qty/total_inv if total_inv else 0):.1%}")

                st.divider()

                tab1, tab2, tab3 = st.tabs(["📝 1. 利润分析", "📊 2. 业务报表 (已汇总)", "🏭 3. 库存分析 (SKU明细)"])
                
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
                    st.caption("业务汇总 (Sheet2) - ⚠️已修复库存汇总逻辑")
                    st.dataframe(try_style(df_sheet2, ['S列_最终净利润'], is_sheet2=True), use_container_width=True, height=600)
                
                with tab3:
                    st.caption("库存分析 (Sheet3)")
                    try:
                        st.dataframe(
                            df_sheet3.style.format(precision=0)
                            .bar(subset=['火箭仓库存数量'], color='#5fba7d')
                            .bar(subset=['极风库存'], color='#4472c4'),
                            use_container_width=True, height=600
                        )
                    except:
                        st.dataframe(df_sheet3, use_container_width=True)

                # ==========================================
                # 📥 下载逻辑
                # ==========================================
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='利润分析')
                    df_sheet2.to_excel(writer, index=False, sheet_name='业务报表')
                    df_sheet3.to_excel(writer, index=False, sheet_name='库存分析')
                    
                    wb = writer.book
                    fmt_header = wb.add_format({'bold': True, 'bg_color': '#4472C4', 'font_color': 'white', 'border': 1, 'align': 'center'})
                    fmt_money = wb.add_format({'num_format': '#,##0', 'align': 'center'})
                    fmt_pct = wb.add_format({'num_format': '0.0%', 'align': 'center'})
                    
                    base_font = {'font_name': 'Microsoft YaHei', 'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter'}
                    fmt_grey = wb.add_format(dict(base_font, bg_color='#BFBFBF'))
                    fmt_white = wb.add_format(dict(base_font, bg_color='#FFFFFF'))

                    def apply_zebra(sheet_name, df_obj, target_col_idx_for_group=0):
                        ws = writer.sheets[sheet_name]
                        for i, col in enumerate(df_obj.columns):
                            str_len = max(df_obj[col].astype(str).map(len).max(), len(str(col))) * 1.5
                            ws.set_column(i, i, min(max(str_len, 10), 40))
                        
                        raw_codes = df_obj.iloc[:, target_col_idx_for_group].astype(str).tolist()
                        clean_codes = [str(x).replace('.0','').replace('"','').strip().upper() for x in raw_codes]
                        is_grey = False
                        for i in range(len(raw_codes)):
                            if i > 0 and clean_codes[i] != clean_codes[i-1]:
                                is_grey = not is_grey
                            ws.set_row(i + 1, None, fmt_grey if is_grey else fmt_white)
                    
                    apply_zebra('利润分析', df_final, IDX_M_CODE)
                    apply_zebra('库存分析', df_sheet3, IDX_M_CODE)

                    ws2 = writer.sheets['业务报表']
                    for i, val in enumerate(df_sheet2.columns): ws2.write(0, i, val, fmt_header)
                    ws2.set_column(0, 0, 20)
                    ws2.set_column(1, 3, 15, fmt_money)
                    ws2.set_column(4, 4, 15, fmt_pct)
                    ws2.set_column(5, 7, 15, fmt_money)
                    ws2.set_column(8, 8, 15, fmt_pct)
                    ws2.set_column(9, 10, 15, fmt_money)

                st.divider()
                st.success("✅ 全套报表生成完毕！")
                
                st.download_button(
                    label="📥 下载 Excel (含利润/业务/库存 3个Sheet)",
                    data=output.getvalue(),
                    file_name="Coupang_Full_Report_Fixed_v6.xlsx",
                    mime="application/vnd.ms-excel",
                    type="primary",
                    use_container_width=True
                )

        except Exception as e:
            st.error(f"❌ 运行出错: {e}")
else:
    st.info("👈 请上传文件 (库存表可选)")
