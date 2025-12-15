import streamlit as st
import pandas as pd
import io
import re

# ==========================================
# 1. 页面配置
# ==========================================
st.set_page_config(layout="wide", page_title="Coupang 利润核算 (最终定稿版)")
st.title("📊 最终定稿：双表输出 (Sheet1保留原样 + Sheet2看板)")
st.markdown("""
### 📝 输出说明：
1.  **Sheet 1 (利润分析)**：完全保持之前的格式、顺序和样式（斑马纹、自动列宽）。
2.  **Sheet 2 (业务看板)**：
    * 仅提取 **A列, Q列, R列, S列**。
    * **顺序严格跟随 Sheet1** (即基础表顺序)，不做额外排序。
    * 样式：大字体 + 净利润数据条。
""")

# --- 列号配置 ---
IDX_M_CODE   = 0    # A列
IDX_M_SKU    = 3    # D列
IDX_M_PROFIT = 10   # K列
IDX_S_ID     = 0    # A列
IDX_S_QTY    = 8    # I列
IDX_A_NAME   = 5    # F列
IDX_A_SPEND  = 15   # P列
# -----------------

# ==========================================
# 2. 上传区域
# ==========================================
with st.sidebar:
    st.header("📂 文件上传")
    file_master = st.file_uploader("1. 基础信息表 (Master)", type=['csv', 'xlsx'])
    file_sales = st.file_uploader("3. 销售表 (Sales)", type=['csv', 'xlsx'])
    file_ads = st.file_uploader("4. 广告表 (Ads)", type=['csv', 'xlsx'])

# ==========================================
# 3. 清洗工具
# ==========================================
def clean_for_match(series):
    return series.astype(str).str.replace(r'\.0$', '', regex=True).str.replace('"', '').str.strip().str.upper()

def clean_num(series):
    return pd.to_numeric(series, errors='coerce').fillna(0)

def extract_code_from_ad(text):
    if pd.isna(text): return None
    match = re.search(r'([Cc]\d+)', str(text))
    if match: return match.group(1).upper()
    return None

def read_file_strict(file):
    try:
        if file.name.endswith('.csv'):
            return pd.read_csv(file, dtype=str)
        else:
            return pd.read_excel(file, dtype=str)
    except:
        file.seek(0)
        return pd.read_csv(file, dtype=str, encoding='gbk')

def get_col_width(series):
    max_len = series.astype(str).map(len).max()
    return max_len

# ==========================================
# 4. 主逻辑
# ==========================================
if file_master and file_sales and file_ads:
    st.divider()
    if st.button("🚀 生成最终报表", type="primary", use_container_width=True):
        try:
            with st.status("🔄 正在计算...", expanded=True):
                # --------------------------------------------
                # 计算逻辑 (完全保持原样)
                # --------------------------------------------
                df_master = read_file_strict(file_master)
                # 记录一下 A 列的原始列名，后面 Sheet2 要用
                col_code_name = df_master.columns[IDX_M_CODE]

                df_master['_MATCH_SKU'] = clean_for_match(df_master.iloc[:, IDX_M_SKU])
                df_master['_MATCH_CODE'] = clean_for_match(df_master.iloc[:, IDX_M_CODE])
                df_master['_VAL_PROFIT'] = clean_num(df_master.iloc[:, IDX_M_PROFIT])

                df_sales = read_file_strict(file_sales)
                df_sales['_MATCH_SKU'] = clean_for_match(df_sales.iloc[:, IDX_S_ID])
                df_sales['销量'] = clean_num(df_sales.iloc[:, IDX_S_QTY])
                sales_agg = df_sales.groupby('_MATCH_SKU')['销量'].sum().reset_index()
                sales_agg.rename(columns={'销量': 'O列_合并销量'}, inplace=True)

                df_ads = read_file_strict(file_ads)
                df_ads['提取编号'] = df_ads.iloc[:, IDX_A_NAME].apply(extract_code_from_ad)
                df_ads['含税广告费'] = clean_num(df_ads.iloc[:, IDX_A_SPEND]) * 1.1
                valid_ads = df_ads.dropna(subset=['提取编号'])
                ads_agg = valid_ads.groupby('提取编号')['含税广告费'].sum().reset_index()
                ads_agg.rename(columns={'提取编号': '_MATCH_CODE', '含税广告费': 'R列_产品总广告费'}, inplace=True)

                # 合并
                df_final = pd.merge(df_master, sales_agg, on='_MATCH_SKU', how='left', sort=False)
                df_final['O列_合并销量'] = df_final['O列_合并销量'].fillna(0).astype(int)
                df_final['P列_SKU总毛利'] = df_final['O列_合并销量'] * df_final['_VAL_PROFIT']
                df_final['Q列_产品总利润'] = df_final.groupby('_MATCH_CODE', sort=False)['P列_SKU总毛利'].transform('sum')
                df_final = pd.merge(df_final, ads_agg, on='_MATCH_CODE', how='left', sort=False)
                df_final['R列_产品总广告费'] = df_final['R列_产品总广告费'].fillna(0)
                df_final['S列_最终净利润'] = df_final['Q列_产品总利润'] - df_final['R列_产品总广告费']

                # --------------------------------------------
                # 关键步骤：在删除辅助列之前，提取 Sheet2 数据
                # --------------------------------------------
                # 1. 提取需要的 4 列：A列(原始名), Q列, R列, S列
                # 注意：keep='first' 确保了顺序严格跟随 Sheet1 (Master表) 的顺序
                df_sheet2 = df_final[[col_code_name, 'Q列_产品总利润', 'R列_产品总广告费', 'S列_最终净利润']].copy()
                df_sheet2 = df_sheet2.drop_duplicates(subset=[col_code_name], keep='first')
                
                # 2. 清理 Sheet1 的辅助列 (保持原代码逻辑)
                cols_to_drop = [c for c in df_final.columns if c.startswith('_')]
                df_final.drop(columns=cols_to_drop, inplace=True)

                # --------------------------------------------
                # Step E: 输出 Excel
                # --------------------------------------------
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    
                    # ========================================
                    # Sheet 1: 利润分析 (代码完全保留，不做修改)
                    # ========================================
                    df_final.to_excel(writer, index=False, sheet_name='利润分析')
                    wb = writer.book
                    ws = writer.sheets['利润分析']
                    
                    # 样式对象 (保持原样)
                    base_font = {'font_name': 'Microsoft YaHei', 'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter'}
                    fmt_row_grey = wb.add_format(dict(base_font, bg_color='#BFBFBF'))
                    fmt_row_white = wb.add_format(dict(base_font, bg_color='#FFFFFF'))
                    fmt_s_profit = wb.add_format(dict(base_font, bg_color='#C6EFCE'))
                    fmt_s_loss = wb.add_format(dict(base_font, bg_color='#FFC7CE'))

                    # 自动列宽循环 (保持原样)
                    for i, col in enumerate(df_final.columns):
                        max_len = get_col_width(df_final[col])
                        header_len = len(str(col)) * 1.5
                        final_width = max(max_len, header_len) + 2
                        if final_width > 50: final_width = 50
                        if final_width < 10: final_width = 10
                        ws.set_column(i, i, final_width)

                    ws.freeze_panes(1, 0)

                    # 智能着色循环 (保持原样)
                    # 需要重新获取 Code 列和 Profit 列的索引，因为 drop 之后位置可能变了，但逻辑不变
                    # 原逻辑是依赖 col_code_idx = IDX_M_CODE (0)
                    col_code_idx = IDX_M_CODE 
                    cols_list = df_final.columns.tolist()
                    col_profit_idx = cols_list.index('S列_最终净利润') if 'S列_最终净利润' in cols_list else -1

                    raw_codes = df_final.iloc[:, col_code_idx].astype(str).tolist()
                    clean_codes = [str(x).replace('.0','').replace('"','').strip().upper() for x in raw_codes]
                    
                    is_grey = False
                    for i in range(len(raw_codes)):
                        excel_row = i + 1
                        if i > 0 and clean_codes[i] != clean_codes[i-1]:
                            is_grey = not is_grey
                        
                        ws.set_row(excel_row, None, fmt_row_grey if is_grey else fmt_row_white)
                        
                        if col_profit_idx != -1:
                            val = df_final.iloc[i, col_profit_idx]
                            try:
                                num_val = float(val)
                            except:
                                num_val = 0
                            
                            if num_val > 0:
                                ws.write(excel_row, col_profit_idx, val, fmt_s_profit)
                            elif num_val < 0:
                                ws.write(excel_row, col_profit_idx, val, fmt_s_loss)
                            else:
                                ws.write(excel_row, col_profit_idx, val, fmt_row_grey if is_grey else fmt_row_white)

                    # ========================================
                    # Sheet 2: 业务报表 (新增)
                    # ========================================
                    df_sheet2.to_excel(writer, index=False, sheet_name='业务报表')
                    ws2 = writer.sheets['业务报表']
                    
                    # 样式设置
                    fmt_header2 = wb.add_format({'font_name': 'Microsoft YaHei', 'bold': True, 'font_size': 12, 'bg_color': '#4472C4', 'font_color': 'white', 'border': 1, 'align': 'center'})
                    fmt_body2 = wb.add_format({'font_name': 'Microsoft YaHei', 'font_size': 11, 'border': 1, 'align': 'center', 'valign': 'vcenter'})
                    fmt_money2 = wb.add_format({'font_name': 'Microsoft YaHei', 'font_size': 11, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '#,##0'})
                    
                    # 设置表头
                    for col_num, value in enumerate(df_sheet2.columns.values):
                        ws2.write(0, col_num, value, fmt_header2)
                    
                    # 设置列宽
                    ws2.set_column(0, 0, 25, fmt_body2) # A列 产品编号
                    ws2.set_column(1, 3, 18, fmt_money2) # 钱列
                    ws2.freeze_panes(1, 0)
                    
                    # 数据条 (Data Bar) - 仅给净利润 (第4列, 索引3)
                    (max_r2, max_c2) = df_sheet2.shape
                    ws2.conditional_format(1, 3, max_r2, 3, {
                        'type': 'data_bar',
                        'bar_color': '#63C384',
                        'bar_negative_color': '#FF0000',
                        'bar_axis_position': 'middle'
                    })

            st.success("✅ 报表生成成功！Sheet1 保持原样，Sheet2 已按顺序生成。")
            st.download_button("📥 下载最终报表", output.getvalue(), "Coupang_Final_Report_v2.xlsx")

        except Exception as e:
            st.error(f"❌ 错误: {e}")
else:
    st.info("👈 请在左侧上传文件")