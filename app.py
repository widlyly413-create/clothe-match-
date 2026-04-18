import streamlit as st
import pandas as pd
import numpy as np
import io
import datetime
import pytz

# ==========================================
# 1. 智能列名映射与配置
# ==========================================
COLUMN_MAPS = {
    '性别': ['性别', '男/女', 'gender', '人员属性'],
    '身高': ['身高', '高度', 'height', 'cm'],
    '体重': ['体重', '重量', 'weight', 'kg', '斤'],
    '胸围': ['胸围', '胸宽', 'chest'],
    '腰围': ['腰围', '腰宽', 'waist'],
    '尺码': ['春秋执勤', '春秋', '秋执勤', '执勤服', '冬执勤', '衣服号', '服装型号', '尺码', '规格', 'size', '型号']
}

FEMALE_SIZES = ["160/80", "160/84", "160/88", "160/92", "160/96", "165/84", "165/88", "165/92", "165/96", "165/100", "165/104", "170/88", "170/92", "170/96", "170/100", "170/104"]
MALE_SIZES = ["165/88", "165/92", "165/96", "165/100", "170/88", "170/92", "170/96", "170/100", "170/104", "170/108", "175/88", "175/92", "175/96", "175/100", "175/104", "175/108", "175/112", "180/92", "180/96", "180/100", "180/104", "180/108", "180/112", "180/116", "180/120", "185/100", "185/104", "185/108", "185/112", "185/116", "185/120"]

# ==========================================
# 2. 核心辅助算法
# ==========================================
def fuzzy_match_columns(df_columns):
    new_mapping = {}
    for col in df_columns:
        c_clean = str(col).lower().strip()
        for target, keywords in COLUMN_MAPS.items():
            if any(key in c_clean for key in keywords):
                new_mapping[col] = target
                break
    return new_mapping

def is_adjacent(size_a, size_b):
    """判定是否为邻近尺码 (身高±5cm 或 胸围±4cm)"""
    if size_a == size_b: return True
    try:
        h_a, x_a = map(int, size_a.split('/'))
        h_b, x_b = map(int, size_b.split('/'))
        # 同号异型 (胸围差4) 或 同型异号 (身高差5)
        if (h_a == h_b and abs(x_a - x_b) <= 4) or (x_a == x_b and abs(h_a - h_b) <= 5):
            return True
        return False
    except: return False

def estimate_chest_by_theory(gender, height, weight):
    if gender == "女":
        std_w = (height - 70) * 0.6
        return (height * 0.52) + (weight - std_w) * 0.85
    else:
        std_w = (height - 80) * 0.7
        return (height * 0.50) + (weight - std_w) * 0.8

def get_track_a_size(gender, height, chest):
    if gender == "女":
        hao = 160 if height <= 162 else (165 if height <= 167 else 170)
        xings = [int(s.split('/')[1]) for s in FEMALE_SIZES if s.startswith(str(hao))]
        return f"{hao}/{min(xings, key=lambda x:abs(x-chest))}" if xings else "160/88"
    else:
        hao = 165 if height <= 167 else (170 if height <= 172 else (175 if height <= 177 else (180 if height <= 182 else 185)))
        xings = [int(s.split('/')[1]) for s in MALE_SIZES if s.startswith(str(hao))]
        return f"{hao}/{min(xings, key=lambda x:abs(x-chest))}" if xings else "175/96"

# 加载历史库
try:
    try: h_df = pd.read_csv("history_data.csv", encoding='utf-8-sig')
    except: h_df = pd.read_csv("history_data.csv", encoding='gbk')
    h_df.columns = h_df.columns.str.strip()
    h_df = h_df.rename(columns=fuzzy_match_columns(h_df.columns))
    history_df = h_df
    history_loaded = True
except: history_loaded = False

def get_track_b_size(gender, height, weight):
    if not history_loaded: return None
    similar = history_df[
        (history_df['性别'] == gender) &
        (history_df['身高'] >= height - 2) & (history_df['身高'] <= height + 2) &
        (history_df['体重'] >= weight - 3) & (history_df['体重'] <= weight + 3)
    ]
    return similar['尺码'].mode()[0] if not similar.empty else None

def dual_track_match(row):
    gender = row.get('性别')
    h_raw, w_raw = row.get('身高'), row.get('体重')
    c_raw = row.get('胸围', np.nan)
    
    if pd.isna(gender) or pd.isna(h_raw) or pd.isna(w_raw):
        return "数据不全", "缺失核心项"
    
    try:
        h, w = float(h_raw), float(w_raw)
        gender = '女' if '女' in str(gender) else '男'
    except: return "格式错误", "包含非法字符"

    try:
        c_theory = float(c_raw) if pd.notna(c_raw) and str(c_raw).strip() != "" else estimate_chest_by_theory(gender, h, w)
    except:
        c_theory = estimate_chest_by_theory(gender, h, w)

    a_size = get_track_a_size(gender, h, c_theory)
    b_size = get_track_b_size(gender, h, w)
    
    if b_size:
        if is_adjacent(a_size, b_size):
            return b_size, "通过"
        else:
            return b_size, f"[人工复核] 建议{a_size}"
    return a_size, "通过 (参考理论值)"

# ==========================================
# 3. 界面逻辑
# ==========================================
st.set_page_config(page_title="服装规格匹配系统", layout="wide")
st.title("👔 智能规格匹配系统")

uploaded_file = st.file_uploader("📂 上传人员信息表格", type=["xlsx", "xls", "csv"])

if uploaded_file:
    try:
        raw_df = pd.read_csv(uploaded_file, header=None) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file, header=None)
        # 寻找表头
        header_idx = 0
        for i in range(min(10, len(raw_df))):
            r_str = "".join([str(x) for x in raw_df.iloc[i].fillna("")])
            if any(k in r_str for k in ["身高", "体重", "性别"]):
                header_idx = i
                break
        
        df = raw_df.iloc[header_idx+1:].reset_index(drop=True)
        df.columns = raw_df.iloc[header_idx].values
        
        # 映射并识别，但保持原始列（不强行加胸围）
        col_map = fuzzy_match_columns(df.columns)
        process_df = df.rename(columns=col_map)

        if st.button("🚀 执行批量智能匹配", type="primary"):
            results = process_df.apply(dual_track_match, axis=1, result_type='expand')
            df['推荐尺码'] = results[0]
            df['系统状态'] = results[1]
            
            st.write("### 📊 匹配分析结果")
            st.dataframe(df.style.map(lambda x: 'color: red; font-weight: bold' if '[人工复核]' in str(x) else 'color: green', subset=['系统状态']))

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='匹配结果')
            st.download_button("📥 下载结果表格", data=output.getvalue(), file_name="匹配结果.xlsx")
    except Exception as e:
        st.error(f"⚠️ 处理出错: {e}")
