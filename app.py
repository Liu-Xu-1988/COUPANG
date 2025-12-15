import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(layout="wide", page_title="Coupang 利润核算 (表头匹配版)")
st.title("🔘 步骤五：多店铺利润核算 (表头匹配版)")
st.markdown("### 操作流程：上传文件 -> 确认就绪 -> **点击按钮** -> 生成报表")
st.caption("✅ 此版本逻辑：通过**表头名称**识别数据，不依赖列的顺序。")

# ==========================================
# 0. 【配置区】请确保你的表格里包含这些表头(列名)
# 如果Coupang改了表头名字，请在这里修改
# ==========================================

# A. 基础信息表 (Master)
# 需要包含: 注册商品ID, 单件毛利, 注册商品名称(或编号)
KEY_M_SKU = '注册商品ID'     # 用于关联销售表
KEY_M_PROFIT = '单件毛利'    # 用于计算利润
KEY_M_CODE = '注册商品名称'  # 用于提取 C01 这种编号 (如果是其他列名请修改这里)

# B. 销售表 (Sales)
# 需要包含: 注册商品ID, 销售数量
KEY_S_ID = '注册商品ID'      # 必须和Master里的ID能对上
KEY_S_QTY = '销售数量'       # 或者是 '销量', 'Quantity'

# C. 广告表 (Ads)
# 需要包含: 广告活动名称, 执行金额
KEY_A_NAME = '广告活动名称'  # 用于提取产品编号
KEY_A_SPEND = '执行金额'     # 或者是 '总花费', 'Spend'

# ==========================================
# 1. 上传区域
# ==========================================
with st.sidebar:
    st.header("1. 文件上传区")
    file_master = st.file_uploader("基础信息表 (Master)", type=['csv', 'xlsx'])
    files_sales = st.file_uploader("销售表 (Sales)", type=['csv', 'xlsx'], accept_multiple_files=True)
    files_ads = st.file_uploader("广告表 (Ads)", type=['csv', 'xlsx'], accept_multiple_files=True)

    st.markdown("---")
    if file_master and files_sales and files_ads:
        st.success("✅ 文件已就绪，请去右侧开始。")
    else:
        st.info("⏳ 等待文件上传...")

# ==========================================
# 2. 工具函数
# ==========================================
def clean_id(series):
    """清洗ID：转字符串，去小数，去空格"""
    return series.astype(str).str.replace(r'\.0$', '', regex=True).str.replace('"', '').str.strip()

def clean_num(series):
    """清洗数值：转数字，无法转换的变0"""
    return pd.to_numeric(series, errors='coerce').fillna(0)

def extract_product_code(text):
    """
    从广告名称中提取 C01, c12 这种编号
    正则逻辑：寻找 C (大小写均可) + 数字
    """
    if pd.isna(text): return None
    match = re.search(r'([Cc]\d+)', str(text))
    if match: return match.group(1).upper()
    return None

def read_file(file):
    """读取文件的通用函数"""
    try:
        file.seek(0)
        if file.name.endswith('.csv'):
            try: return pd.read_csv(file)
            except: file.seek(0); return pd.read_csv(file, encoding='gbk')
        else:
            return pd.read_excel(file)
    except Exception as e:
        st.error(f"❌ 读取失败: {file.name} - {e}")
        return pd.DataFrame()

# ==========================================
# 3. 主程序
# ==========================================

if file_master and files_sales and files_ads:
    st.divider()
    
    if st.button("🚀 点击开始计算", type="primary", use_container_width=True):
        st.divider()
        with st.status("🔄 正在计算中...", expanded=True):
            try:
                # -------------------------------------------------------
                # A. 处理 Master (基础表)
                # -------------------------------------------------------
                st.write("1. 读取基础表...")
                df_master = read_file(file_master)
                
                # 检查列名是否存在
                missing_cols = [col for col in [KEY_M_SKU, KEY_M_PROFIT, KEY_M_CODE] if col not in df_master.columns]
                if missing_cols:
                    st.error(f"❌ 基础表中找不到这些列名: {missing_cols}")
                    st.stop()
                
                df_master['__ORDER__'] = range(len(df_master))
                df_master['关联ID'] = clean_id(df_master[KEY_M_SKU])
                df_master['单件毛利'] = clean_num(df_master[KEY_M_PROFIT])
                df_master['产品编号_清洗'] = clean_id(df_master[KEY_M_CODE]).str.upper() # 这里如果是"注册商品名称"，通常里面包含了C01

                # 如果编号在"注册商品名称"里混着，尝试提取一下
                # 如果你的基础表有一列专门叫"产品编号"，可以不用这一步
                if '产品编号_清洗' not in df_master.columns or df_master['产品编号_清洗'].iloc[0] == '':
                     df_master['产品编号_清洗'] = df_master[KEY_M_CODE].apply(extract_product_code)

                # -------------------------------------------------------
                # B. 处理 Sales (销售表)
                # -------------------------------------------------------
                st.write("2. 合并销售数据...")
                all_sales = []
                for f in files_sales:
                    df = read_file(f)
                    if not df.empty:
                        # 兼容不同列名 (如果有时候是 '销量', 有时候是 '销售数量')
                        if KEY_S_QTY not in df.columns and '销量' in df.columns:
                            df.rename(columns={'销量': KEY_S_QTY}, inplace=True)
                            
                        if KEY_S_ID in df.columns and KEY_S_QTY in df.columns:
                            all_sales.append(df)
                        else:
                            st.warning(f"⚠️ 文件 {f.name} 缺少 '{KEY_S_ID}' 或 '{KEY_S_QTY}' 列，已跳过")
                
                if all_sales:
                    df_sales_all = pd.concat(all_sales, ignore_index=True)
                    df_sales_all['关联ID'] = clean_id(df_sales_all[KEY_S_ID])
                    df_sales_all['销量'] = clean_num(df_sales_all[KEY_S_QTY])
                    sales_agg = df_sales_all.groupby('关联ID')['销量'].sum().reset_index()
                    sales_agg.rename(columns={'销量': 'O列_合并销量'}, inplace=True)
                else:
                    sales_agg = pd.DataFrame(columns=['关联ID', 'O列_合并销量'])

                # -------------------------------------------------------
                # C. 处理 Ads (广告表)
                # -------------------------------------------------------
                st.write("3. 匹配广告花费...")
                all_ads = []
                for f in files_ads:
                    df = read_file(f)
                    if not df.empty:
                        # 检查列名
                        if KEY_A_NAME in df.columns and KEY_A_SPEND in df.columns:
                            all_ads.append(df)
                        else:
                            st.warning(f"⚠️ 广告表 {f.name} 缺少 '{KEY_A_NAME}' 或 '{KEY_A_SPEND}'，已跳过")

                if all_ads:
                    df_ads_all = pd.concat(all_ads, ignore_index=True)
                    # 提取编号
                    df_ads_all['提取编号'] = df_ads_all[KEY_A_NAME].apply(extract_product_code)
                    # 计算含税 (10%)
                    df_ads_all['含税广告费'] = clean_num(df_ads_all[KEY_A_SPEND]) * 1.1
                    
                    ads_agg = df_ads_all.groupby('提取编号')['含税广告费'].sum().reset_index()
                    ads_agg.rename(columns={'提取编号': '产品编号_清洗', '含税广告费': 'R列_产品总广告费'}, inplace=True)
                else:
                    ads_agg = pd.DataFrame(columns=['产品编号_清洗', 'R列_产品总广告费'])

                # -------------------------------------------------------
                # D. 合并计算
                # -------------------------------------------------------
                st.write("4. 生成最终报表...")
                
                # 1. 基础表 + 销量
                df_final = pd.merge(df_master, sales_agg, on='关联ID', how='left')
                df_final['O列_合并销量'] = df_final['O列_合并销量'].fillna(0).astype(int)
                
                # 2. 算SKU毛利
                df_final['P列_SKU总毛利'] = df_final['O列_合并销量'] * df_final['单件毛利']
                
                # 3. 算产品总利润 (按清洗后的编号汇总)
                # 注意：如果提取不到编号，这里会是空的
                df_final['Q列_产品总利润'] = df_final.groupby('产品编号_清洗')['P列_SKU总毛利'].transform('sum')
                
                # 4. 减去广告费 (按清洗后的编号匹配)
                df_final = pd.merge(df_final, ads_agg, on='产品编号_清洗', how='left')
                df_final['R列_产品总广告费'] = df_final['R列_产品总广告费'].fillna(0)
                
                # 5. 最终净利
                df_final['S列_最终净利润'] = df_final['Q列_产品总利润'] - df_final['R列_产品总广告费']

                # 清理
                df_final.sort_values(by=['__ORDER__'], inplace=True)
                keep_cols = [c for c in df_final.columns if c not in ['__ORDER__', '关联ID', '单件毛利', '提取编号']]
                df_final = df_final[keep_cols]

                # -------------------------------------------------------
                # E. 导出 Excel
                # -------------------------------------------------------
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    wb = writer.book
                    
                    # Sheet 1
                    df_final.to_excel(writer, index=False, sheet_name='Result')
                    ws = writer.sheets['Result']
                    
                    # 简单样式
                    fmt_header = wb.add_format({'bold': True, 'bg_color': '#DDEBF7', 'border': 1})
                    for col_num, value in enumerate(df_final.columns.values):
                        ws.write(0, col_num, str(value), fmt_header)
                    ws.set_column(0, len(df_final.columns)-1, 15)

                st.success("✅ 计算成功！")
                st.download_button("📥 下载结果报表", output.getvalue(), "Coupang_Result.xlsx", "application/vnd.ms-excel", type='primary')

            except Exception as e:
                st.error(f"❌ 运行出错: {e}")
                st.info("💡 建议检查：上传的表格里，表头名字是不是改了？请看代码最上面的【配置区】。")

else:
    st.info("👈 请上传文件")