import streamlit as st
import pandas as pd
import numpy as np
import io
import datetime
import pytz

# ==========================================
# 1. 智能列名映射字典 (Fuzzy Mapping Dictionary)
# ==========================================
COLUMN_MAPS = {
    '性别': ['性别', '男/女', 'gender', '人员属性'],
    '身高': ['身高', '高度', 'height', 'cm'],
    '体重': ['体重', '重量', 'weight', 'kg', '斤'],
    '胸围': ['胸围', '胸宽', 'chest'],
    '腰围': ['腰围', '腰宽', 'waist'],
    '尺码': ['春秋执勤', '春秋', '秋执勤', '执勤服', '冬执勤', '衣服号', '服装型号', '尺码', '规格', 'size', '型号']
}

def fuzzy_match_columns(df_columns):
    """根据字典自动识别并映射列名"""
    new_mapping = {}
    for col in df_columns:
        c_clean = str(col).lower().strip()
        for target, keywords in COLUMN_MAPS.items():
            if any(key in c_clean for key in keywords):
                new_mapping[col] = target
                break
    return new_mapping

# ==========================================
# 2. 核心算法逻辑 (Core Engine)
# ==========================================

FEMALE_SIZES = ["160/80", "160/84", "160/88", "160/92", "160/96", "165/84", "165/88", "165/92", "165/96", "165/100", "165/104", "170/88", "170/92", "170/96", "170/100", "170/104"]
MALE_SIZES = ["165/88", "165/92", "165/96", "165/100", "170/88", "170/92", "170/96", "170/100", "170/104", "170/108", "175/88", "175/92", "175/96", "175/100", "175/104", "175/108", "175/112", "180/92", "180/96", "180/100", "180/104", "180/108", "180/112", "180/116", "180/120", "185/100", "185/104", "185/108", "185/112", "185/116", "185/120"]

# 加载历史库
try:
    try:
        h_df = pd.read_csv("history_data.csv", encoding='utf-8-sig')
    except UnicodeDecodeError:
        h_df = pd.read_csv("history_data.csv", encoding='gbk')
    
    h_df.columns = h_df.columns.str.strip()
    h_df = h_df.rename(columns=fuzzy_match_columns(h_df.columns))
    history_df = h_df
    history_loaded = True
except Exception as e:
    st.sidebar.error(f"历史库识别失败: {e}")
    history_loaded = False

def estimate_chest_by_theory(gender, height, weight):
    """基于人机工程学的理论估算"""
    if gender == "女":
        std_weight = (height - 70) * 0.6
        return (height * 0.52) + (weight - std_weight) * 0.85
    else:
        std_weight = (height - 80) * 0.7
        return (height * 0.50) + (weight - std_weight) * 0.8

def get_track_a_size(gender, height, chest):
    """Track A: 标准规则"""
    if gender == "女":
        hao = 160 if height <= 162 else (165 if height <= 167 else 170)
        xings = [int(s.split('/')[1]) for s in FEMALE_SIZES if s.startswith(str(hao))]
        return f"{hao}/{min(xings, key=lambda x:abs(x-chest))}" if xings else "无匹配"
    else:
        hao = 165 if height <= 167 else (170 if height <= 172 else (175 if height <= 177 else (180 if height <= 182 else 185)))
        xings = [int(s.split('/')[1]) for s in MALE_SIZES if s.startswith(str(hao))]
        return f"{hao}/{min(xings, key=lambda x:abs(x-chest))}" if xings else "无匹配"

def get_track_b_size(gender, height, weight):
    """Track B: 经验检索"""
    if not history_loaded: return None
    similar = history_df[
        (history_df['性别'] == gender) &
        (history_df['身高'] >= height - 2) & (history_df['身高'] <= height + 2) &
        (history_df['体重'] >= weight - 3) & (history_df['体重'] <= weight + 3)
    ]
    return similar['尺码'].mode()[0] if not similar.empty else None

def dual_track_match(row):
    """双轨合并逻辑"""
    gender = row.get('性别')
    h_raw, w_raw = row.get('身高'), row.get('体重')
    c_raw = row.get('胸围', np.nan)
    
    if pd.isna(gender) or pd.isna(h_raw) or pd.isna(w_raw):
        return "数据不全", "缺失核心项"
    
    try:
        h, w = float(h_raw), float(w_raw)
        # 统一性别表述
        gender = '女' if '女' in str(gender) else '男'
    except:
        return "格式错误", "包含非法字符"

    # Track A 理论胸围
    try:
        c_theory = float(c_raw) if pd.notna(c_raw) and str(c_raw).strip() != "" else estimate_chest_by_theory(gender, h, w)
    except:
        c_theory = estimate_chest_by_theory(gender, h, w)

    a_size = get_track_a_size(gender, h, c_theory)
    b_size = get_track_b_size(gender, h, w)
    
    # 经验优先
    if b_size:
        return b_size, ("通过" if b_size == a_size else f"[人工复核] 标准规则建议 {a_size}")
    return a_size, "通过 (参考理论值)"

# ==========================================
# 3. 文件处理与网页 UI
# ==========================================
st.set_page_config(page_title="被装匹配系统", layout="wide")
st.title("👔 智能规格匹配系统 (智能表头版)")

with st.sidebar:
    cst_time = datetime.datetime.now(pytz.timezone('Asia/Shanghai')).strftime('%Y-%m-%d %H:%M:%S')
    st.info(f"系统运行时间：{cst_time} (CST)")
    st.success("历史库状态: 已挂载" if history_loaded else "历史库状态: 未就绪")

uploaded_file = st.file_uploader("📂 上传人员体征表格", type=["xlsx", "xls", "csv"])

if uploaded_file:
    try:
        # 自动识别表头行
        raw_df = pd.read_csv(uploaded_file, header=None) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file, header=None)
        header_idx = 0
        for i in range(min(10, len(raw_df))):
            r_str = "".join([str(x) for x in raw_df.iloc[i].fillna("")])
            if "身高" in r_str or "体重" in r_str or "cm" in r_str.lower():
                header_idx = i
                break
        
        df = raw_df.iloc[header_idx+1:].reset_index(drop=True)
        df.columns = raw_df.iloc[header_idx].values
        
        # 应用模糊映射
        df = df.rename(columns=fuzzy_match_columns(df.columns))
        st.write("### 🧹 智能列名识别预览", df.head(3))

        if not all(col in df.columns for col in ['性别', '身高', '体重']):
            st.error("❌ 无法识别关键列（性别/身高/体重），请检查表头。")
        else:
            if st.button("🚀 执行批量匹配", type="primary"):
                results = df.apply(dual_track_match, axis=1, result_type='expand')
                df['推荐尺码'] = results[0]
                df['匹配状态'] = results[1]
                
                st.write("### 📊 匹配结果")
                st.dataframe(df.style.map(lambda x: 'color: red' if '[人工复核]' in str(x) else '', subset=['匹配状态']))

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='结果')
                st.download_button("📥 下载结果表格", data=output.getvalue(), file_name="匹配结果.xlsx")
    except Exception as e:
        st.error(f"⚠️ 处理出错: {e}")
