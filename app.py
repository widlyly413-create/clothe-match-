import streamlit as st
import pandas as pd
import io

# ==========================================
# 1. 核心算法逻辑 (基于提供的数据资产)
# ==========================================

# 加载标准清单 
FEMALE_SIZES = ["160/80", "160/84", "160/88", "160/92", "160/96", "165/84", "165/88", "165/92", "165/96", "165/100", "165/104", "170/88", "170/92", "170/96", "170/100", "170/104"]
MALE_SIZES = ["165/88", "165/92", "165/96", "165/100", "170/88", "170/92", "170/96", "170/100", "170/104", "170/108", "175/88", "175/92", "175/96", "175/100", "175/104", "175/108", "175/112", "180/92", "180/96", "180/100", "180/104", "180/108", "180/112", "180/116", "180/120", "185/100", "185/104", "185/108", "185/112", "185/116", "185/120"]

def get_track_a_size(gender, height, chest):
    """Track A: 标准规则匹配 """
    if gender == "女":
        hao = 160 if height <= 162 else (165 if height <= 167 else 170)
        # 寻找最接近的型
        return f"{hao}/{min([int(s.split('/')[1]) for s in FEMALE_SIZES if s.startswith(str(hao))], key=lambda x:abs(x-chest))}"
    else:
        hao = 165 if height <= 167 else (170 if height <= 172 else (175 if height <= 177 else (180 if height <= 182 else 185)))
        return f"{hao}/{min([int(s.split('/')[1]) for s in MALE_SIZES if s.startswith(str(hao))], key=lambda x:abs(x-chest))}"

def get_track_b_size(gender, height, weight):
    """Track B: 历史经验模拟 """
    # 模拟历史数据中针对矮个女性的特殊尺码建议
    if gender == "女" and height <= 155:
        return "155/88" if weight <= 55 else "160/88"
    return None

def dual_track_match(row):
    """双轨合并逻辑"""
    gender = row['性别']
    h = row['身高']
    w = row['体重']
    c = row['胸围']
    
    a_size = get_track_a_size(gender, h, c)
    b_size = get_track_b_size(gender, h, w)
    
    if b_size and b_size != a_size:
        return a_size, f"[人工复核] 经验提示选{b_size}"
    return a_size, "通过"

# ==========================================
# 2. 网页端交互界面
# ==========================================
st.set_page_config(page_title="被装智能匹配系统", layout="wide")
st.title("👔 执勤服智能规格匹配系统 (批量处理版)")

# 侧边栏：配置与说明
with st.sidebar:
    st.header("⚙️ 配置中心")
    st.info("系统当前运行时间：中国标准时间 (CST)")
    st.markdown("""
    **逻辑说明：**
    1. **Track A:** 基于标准清单计算 
    2. **Track B:** 基于历史数据修正 
    3. **冲突对齐:** 自动标记差异数据
    """)

uploaded_file = st.file_uploader("📂 上传单位人员体征表格 (Excel/CSV)", type=["xlsx", "xls", "csv"])

if uploaded_file:
    # 读取数据
    if uploaded_file.name.endswith('.csv'):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)
    
    st.write("### 原始数据预览")
    st.dataframe(df.head())

    if st.button("🚀 执行批量智能匹配"):
        # 进度条
        progress_bar = st.progress(0)
        
        # 核心匹配过程
        results = df.apply(dual_track_match, axis=1, result_type='expand')
        df['推荐尺码'] = results[0]
        df['匹配状态'] = results[1]
        
        progress_bar.progress(100)
        
        # 结果展示
        st.write("### 匹配结果详情")
        
        # 定义红色标记样式
        def highlight_conflict(val):
            color = 'red' if '[人工复核]' in str(val) else 'black'
            return f'color: {color}'

        styled_df = df.style.applymap(highlight_conflict, subset=['匹配状态'])
        st.dataframe(styled_df)

        # 准备下载
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='匹配结果')
            # 自动调整列宽
            worksheet = writer.sheets['匹配结果']
            for i, col in enumerate(df.columns):
                worksheet.set_column(i, i, max(len(col), 15))
                
        st.download_button(
            label="📥 下载处理后的 Excel 表格",
            data=output.getvalue(),
            file_name="执勤服尺码匹配结果_V2.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )