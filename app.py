import streamlit as st
import pandas as pd
import io
import datetime
import pytz

# ==========================================
# 1. 核心算法逻辑 (Core Logic Engine)
# ==========================================

# 标准清单 (Track A)
FEMALE_SIZES = ["160/80", "160/84", "160/88", "160/92", "160/96", "165/84", "165/88", "165/92", "165/96", "165/100", "165/104", "170/88", "170/92", "170/96", "170/100", "170/104"]
MALE_SIZES = ["165/88", "165/92", "165/96", "165/100", "170/88", "170/92", "170/96", "170/100", "170/104", "170/108", "175/88", "175/92", "175/96", "175/100", "175/104", "175/108", "175/112", "180/92", "180/96", "180/100", "180/104", "180/108", "180/112", "180/116", "180/120", "185/100", "185/104", "185/108", "185/112", "185/116", "185/120"]

# 尝试加载历史经验库 (Track B)
try:
    history_df = pd.read_csv("history_data.csv")
    history_loaded = True
except FileNotFoundError:
    history_df = None
    history_loaded = False

def get_track_a_size(gender, height, chest):
    """Track A: 标准规则匹配"""
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
    """Track B: 真实历史经验检索"""
    if not history_loaded:
        return None
    
    # 设定相似体征的宽容度：身高上下2cm，体重上下3kg
    similar_cases = history_df[
        (history_df['性别'] == gender) &
        (history_df['身高(cm)'] >= height - 2) & (history_df['身高(cm)'] <= height + 2) &
        (history_df['体重(kg)'] >= weight - 3) & (history_df['体重(kg)'] <= weight + 3)
    ]
    
    if not similar_cases.empty:
        # 获取这群相似体型的人选择最多的尺码 (众数)
        most_frequent_size = similar_cases['推荐尺码'].mode()[0]
        return most_frequent_size
    
    return None

def dual_track_match(row):
    """双轨合并逻辑"""
    # 提取数据并进行基础清理
    gender = row.get('性别')
    h = row.get('身高')
    w = row.get('体重')
    c = row.get('胸围')
    
    # 异常值过滤
    if pd.isna(gender) or pd.isna(h) or pd.isna(w) or pd.isna(c):
        return "数据不全", "数据缺失"

    try:
        h, w, c = float(h), float(w), float(c)
    except ValueError:
         return "数据格式错误", "包含非数字字符"

    # 执行双轨匹配
    a_size = get_track_a_size(gender, h, c)
    b_size = get_track_b_size(gender, h, w)
    
    # 冲突对齐逻辑
    if b_size and b_size != a_size:
        return a_size, f"[人工复核] 经验提示选 {b_size}"
    return a_size, "通过"

# ==========================================
# 2. 网页端交互界面 (Streamlit UI)
# ==========================================
st.set_page_config(page_title="智能被装匹配系统", layout="wide", page_icon="👔")
st.title("👔 执勤服智能规格匹配系统 (批量处理版)")

# 侧边栏：配置与说明
with st.sidebar:
    st.header("⚙️ 配置中心")
    
    # 按要求设置并显示中国标准时间 (CST)
    cst_tz = pytz.timezone('Asia/Shanghai')
    cst_time = datetime.datetime.now(cst_tz).strftime('%Y-%m-%d %H:%M:%S')
    st.info(f"系统当前运行时间：\n{cst_time} (CST)")
    
    st.markdown("---")
    st.markdown("""
    **逻辑说明：**
    1. **Track A:** 基于标准清单物理计算
    2. **Track B:** 基于历史经验库(CSV)修正
    3. **冲突对齐:** 自动标记 Track A 与 B 的差异数据
    """)
    
    if history_loaded:
        st.success(f"✅ 历史经验库已挂载 (共 {len(history_df)} 条记录)")
    else:
        st.error("⚠️ 未检测到 history_data.csv，Track B 经验检索功能失效。请确保该文件在根目录。")

uploaded_file = st.file_uploader("📂 上传单位人员体征表格 (Excel/CSV)", type=["xlsx", "xls", "csv"])

if uploaded_file:
    try:
        # 读取数据
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
            
        # 【关键修复】清除表头前后的隐藏空格，防止 KeyError
        df.columns = df.columns.str.strip()
        
        st.write("### 原始数据预览")
        st.dataframe(df.head())

        # 【关键修复】核心字段校验预警
        required_columns = ['性别', '身高', '体重', '胸围']
        missing_columns = [col for col in required_columns if col not in df.columns]

        if missing_columns:
            st.error(f"❌ 解析失败！上传的表格缺少必要的表头：**{', '.join(missing_columns)}**")
            st.warning(f"👉 系统实际读取到的表头是：{', '.join(df.columns)}")
            st.info("请检查 Excel 文件，确保第一行是正确的表头名称（不要包含多余的隐藏空格）。")
        else:
            # 字段齐全，显示运行按钮
            if st.button("🚀 执行批量智能匹配", type="primary"):
                progress_bar = st.progress(0)
                
                # 核心匹配过程
                results = df.apply(dual_track_match, axis=1, result_type='expand')
                df['推荐尺码'] = results[0]
                df['匹配状态'] = results[1]
                
                progress_bar.progress(100)
                
                st.write("### 匹配结果详情")
                
                # 定义红色高亮样式
                def highlight_conflict(val):
                    color = 'red' if '[人工复核]' in str(val) else 'black'
                    return f'color: {color}'

                styled_df = df.style.applymap(highlight_conflict, subset=['匹配状态'])
                st.dataframe(styled_df)

                # 准备 Excel 文件下载 (输出到内存 BytesIO)
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='匹配结果')
                    # 自动调整列宽以增加可读性
                    worksheet = writer.sheets['匹配结果']
                    for i, col in enumerate(df.columns):
                        worksheet.set_column(i, i, max(len(col), 15))
                        
                st.download_button(
                    label="📥 下载处理后的 Excel 表格",
                    data=output.getvalue(),
                    file_name="执勤服尺码匹配结果_处理完成.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    except Exception as e:
         st.error(f"⚠️ 文件处理过程中发生未知错误: {str(e)}")