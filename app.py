import streamlit as st
import pandas as pd
import numpy as np
import io
import datetime
import pytz

# ==========================================
# 1. 核心算法逻辑 (Core Logic Engine)
# ==========================================

FEMALE_SIZES = ["160/80", "160/84", "160/88", "160/92", "160/96", "165/84", "165/88", "165/92", "165/96", "165/100", "165/104", "170/88", "170/92", "170/96", "170/100", "170/104"]
MALE_SIZES = ["165/88", "165/92", "165/96", "165/100", "170/88", "170/92", "170/96", "170/100", "170/104", "170/108", "175/88", "175/92", "175/96", "175/100", "175/104", "175/108", "175/112", "180/92", "180/96", "180/100", "180/104", "180/108", "180/112", "180/116", "180/120", "185/100", "185/104", "185/108", "185/112", "185/116", "185/120"]

# 尝试加载历史经验库 (Track B)
try:
    history_df = pd.read_csv("history_data.csv")
    history_loaded = True
except FileNotFoundError:
    history_df = None
    history_loaded = False

def estimate_chest_by_theory(gender, height, weight):
    """基于服装人体工程学的纯理论胸围估算 (Track A 专用)"""
    if gender == "女":
        std_weight = (height - 70) * 0.6
        base_chest = height * 0.52
        return base_chest + (weight - std_weight) * 0.85
    else:
        std_weight = (height - 80) * 0.7
        base_chest = height * 0.50
        return base_chest + (weight - std_weight) * 0.8

def get_track_a_size(gender, height, chest):
    """Track A: 标准规则匹配 (物理算法)"""
    if gender == "女":
        hao = 160 if height <= 162 else (165 if height <= 167 else 170)
        available_xings = [int(s.split('/')[1]) for s in FEMALE_SIZES if s.startswith(str(hao))]
        if not available_xings: return "无匹配标准码"
        return f"{hao}/{min(available_xings, key=lambda x:abs(x-chest))}"
    else:
        hao = 165 if height <= 167 else (170 if height <= 172 else (175 if height <= 177 else (180 if height <= 182 else 185)))
        available_xings = [int(s.split('/')[1]) for s in MALE_SIZES if s.startswith(str(hao))]
        if not available_xings: return "无匹配标准码"
        return f"{hao}/{min(available_xings, key=lambda x:abs(x-chest))}"

def get_track_b_size(gender, height, weight):
    """Track B: 真实历史经验检索 (经验算法)"""
    if not history_loaded: return None
    similar_cases = history_df[
        (history_df['性别'] == gender) &
        (history_df['身高(cm)'] >= height - 2) & (history_df['身高(cm)'] <= height + 2) &
        (history_df['体重(kg)'] >= weight - 3) & (history_df['体重(kg)'] <= weight + 3)
    ]
    if not similar_cases.empty:
        return similar_cases['推荐尺码'].mode()[0]
    return None

def dual_track_match(row):
    """双轨合并逻辑 (彻底解耦版)"""
    gender = row.get('性别')
    h_val = row.get('身高')
    w_val = row.get('体重')
    c_val = row.get('胸围')
    
    if pd.isna(gender) or pd.isna(h_val) or pd.isna(w_val):
        return "数据不全", "缺失必填项(性别/身高/体重)"

    try:
        h, w = float(h_val), float(w_val)
    except (ValueError, TypeError):
        return "数据错误", "身高或体重包含非数字"

    # --- 核心逻辑分离 ---
    status_flags = []
    
    # 1. 确定用于 Track A 的胸围数据
    if pd.isna(c_val) or str(c_val).strip() == "":
        c_theory = estimate_chest_by_theory(gender, h, w)
        status_flags.append("[理论估算胸围]")
    else:
        try:
            c_theory = float(c_val)
        except (ValueError, TypeError):
            c_theory = estimate_chest_by_theory(gender, h, w)
            status_flags.append("[理论估算胸围]")

    # 2. 分别运行两条独立轨道
    a_size = get_track_a_size(gender, h, c_theory) # 理论派
    b_size = get_track_b_size(gender, h, w)        # 经验派
    
    # 3. 冲突对齐逻辑
    if b_size and b_size != a_size:
        status_flags.append(f"[人工复核] 历史高频选 {b_size}")
    else:
        if not status_flags: status_flags.append("通过")
            
    return a_size, " | ".join(status_flags)

# ==========================================
# 2. 智能表头预处理 (Smart Header Parsing)
# ==========================================
def process_uploaded_file(uploaded_file):
    if uploaded_file.name.endswith('.csv'):
        raw_df = pd.read_csv(uploaded_file, header=None)
    else:
        raw_df = pd.read_excel(uploaded_file, header=None)
        
    header_idx = 0
    for i in range(min(10, len(raw_df))):
        row_str = "".join([str(x) for x in raw_df.iloc[i].fillna("")]).replace(" ", "")
        if "身高" in row_str and "体重" in row_str:
            header_idx = i
            break
            
    df = raw_df.iloc[header_idx+1:].reset_index(drop=True)
    df.columns = raw_df.iloc[header_idx].values
    
    col_mapping = {}
    for col in df.columns:
        c_str = str(col).replace(" ", "").replace("\n", "")
        if "性别" in c_str: col_mapping[col] = "性别"
        elif "身高" in c_str: col_mapping[col] = "身高"
        elif "体重" in c_str: col_mapping[col] = "体重"
        elif "胸围" in c_str: col_mapping[col] = "胸围"
        elif "姓名" in c_str: col_mapping[col] = "姓名"
        
    df = df.rename(columns=col_mapping)
    if "胸围" not in df.columns:
        df["胸围"] = np.nan
        
    df = df.dropna(subset=['性别', '身高', '体重'], how='all')
    return df

# ==========================================
# 3. 网页端交互界面 (Streamlit UI)
# ==========================================
st.set_page_config(page_title="被装智能匹配系统", layout="wide", page_icon="👔")
st.title("👔 执勤服智能规格匹配系统 (双轨解耦版)")

with st.sidebar:
    st.header("⚙️ 配置中心")
    cst_tz = pytz.timezone('Asia/Shanghai')
    cst_time = datetime.datetime.now(cst_tz).strftime('%Y-%m-%d %H:%M:%S')
    st.info(f"系统运行时间：\n{cst_time} (CST)")
    st.markdown("---")
    st.markdown("""
    **双轨匹配架构：**
    - **Track A (理论)**: 基于人机工程学比例推算。
    - **Track B (经验)**: 检索历史数据库相似案例。
    - **自动冲突识别**: 两轨结果不一致时自动报警。
    """)

uploaded_file = st.file_uploader("📂 上传单位人员体征表格", type=["xlsx", "xls", "csv"])

if uploaded_file:
    try:
        df = process_uploaded_file(uploaded_file)
        st.write("### 🧹 数据清洗预览")
        st.dataframe(df.head())

        required_columns = ['性别', '身高', '体重']
        if not all(col in df.columns for col in required_columns):
            st.error(f"❌ 核心数据缺失，识别到的列有：{', '.join(df.columns)}")
        else:
            if st.button("🚀 执行批量智能匹配", type="primary"):
                results = df.apply(dual_track_match, axis=1, result_type='expand')
                df['推荐尺码'] = results[0]
                df['系统状态'] = results[1]
                
                st.write("### 📊 匹配结果详情")
                
                # 定义高亮样式函数
                def highlight_status(val):
                    if '[人工复核]' in str(val):
                        return 'color: #D32F2F; font-weight: bold'
                    elif '[理论估算胸围]' in str(val):
                        return 'color: #1976D2'
                    return 'color: green'

                # --- 关键修复：将 .applymap 替换为 .map ---
                st.dataframe(df.style.map(highlight_status, subset=['系统状态']))

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='匹配结果')
                    worksheet = writer.sheets['匹配结果']
                    for i, col in enumerate(df.columns):
                        worksheet.set_column(i, i, max(len(str(col)), 12))
                        
                st.download_button(
                    label="📥 下载完成匹配的 Excel 表格",
                    data=output.getvalue(),
                    file_name="执勤服批量匹配结果_双轨版.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    except Exception as e:
         st.error(f"⚠️ 发生错误: {str(e)}")
