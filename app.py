import streamlit as st
import pandas as pd
import io
import os

# ===========================
# 配置：固定B模板存储文件夹
# ===========================
TEMPLATE_FOLDER = "b_templates"
if not os.path.exists(TEMPLATE_FOLDER):
    os.makedirs(TEMPLATE_FOLDER)

# ===========================
# 核心：强制文本读取（彻底解决ID尾数变0000）
# ===========================
@st.cache_data(ttl=3600)
def read_excel(file):
    return pd.read_excel(
        file,
        engine="openpyxl",
        dtype=str
    )

# ===========================
# B模板管理函数
# ===========================
def get_b_templates():
    return [f for f in os.listdir(TEMPLATE_FOLDER) if f.endswith((".xlsx", ".xls"))]

def save_b_template(uploaded_file):
    path = os.path.join(TEMPLATE_FOLDER, uploaded_file.name)
    with open(path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return path

# ===========================
# 三表匹配引擎（C→A→B）
# ===========================
class DataMatcher:
    def __init__(self, df_a):
        self.df_a = df_a.copy()
        for col in self.df_a.columns:
            self.df_a[col] = self.df_a[col].astype(str).str.strip()

    def match_c_to_a(self, df_c, key_c, key_a):
        df_c_clean = df_c.copy()
        for col in df_c_clean.columns:
            df_c_clean[col] = df_c_clean[col].astype(str).str.strip()
        
        result = pd.merge(
            df_c_clean[[key_c]],
            self.df_a,
            left_on=key_c,
            right_on=key_a,
            how="left"
        )
        return result

    def fill_b_template(self, df_matched, df_b_template, mapping):
        df_b = df_b_template.copy()
        for col in df_b.columns:
            df_b[col] = df_b[col].astype(str).str.strip()

        for b_col, a_col in mapping.items():
            if b_col in df_b.columns and a_col in df_matched.columns:
                df_b[b_col] = df_matched[a_col]
        return df_b

# ===========================
# 网页界面
# ===========================
st.set_page_config(page_title="三表Excel自动生成器", layout="wide")
st.title("📊 A+B+C三表Excel一键生成器（公网版）")

# ===========================
# 侧边栏
# ===========================
with st.sidebar:
    st.header("1. B固定模板管理")
    with st.expander("📤 上传新B模板"):
        up_b = st.file_uploader("上传B模板", type=['xlsx','xls'], key="up_b")
        if up_b and st.button("保存模板"):
            save_b_template(up_b)
            st.success("✅ 模板保存成功！")

    template_list = get_b_templates()
    selected_b = st.selectbox("选择B模板", template_list)

    st.header("2. 上传数据文件")
    up_a = st.file_uploader("上传A表（总数据源）", type=['xlsx','xls'], key="up_a")
    up_c = st.file_uploader("上传C表（仅主键）", type=['xlsx','xls'], key="up_c")

# ===========================
# 主界面
# ===========================
if up_a and up_c and selected_b:
    df_a = read_excel(up_a)
    df_c = read_excel(up_c)
    df_b = read_excel(os.path.join(TEMPLATE_FOLDER, selected_b))

    col1, col2, col3 = st.columns(3)
    with col1: st.subheader("A表"); st.dataframe(df_a.head(3), use_container_width=True)
    with col2: st.subheader("C表"); st.dataframe(df_c.head(3), use_container_width=True)
    with col3: st.subheader("B表"); st.dataframe(df_b.head(3), use_container_width=True)

    matcher = DataMatcher(df_a)
    st.divider()
    st.subheader("⚙️ 匹配配置")

    col_key1, col_key2 = st.columns(2)
    with col_key1: key_c = st.selectbox("C表主键列", df_c.columns)
    with col_key2: key_a = st.selectbox("A表对应主键列", df_a.columns)

    st.write("🔗 B模板 ↔ A表 字段映射")
    mapping = {}
    map_cols = st.columns(3)
    for i, b_col in enumerate(df_b.columns):
        a_options = ['--- 不填充 ---'] + list(df_a.columns)
        default_idx = a_options.index(b_col) if b_col in df_a.columns else 0
        with map_cols[i % 3]:
            sel = st.selectbox(f"B列：{b_col}", a_options, index=default_idx)
            if sel != '--- 不填充 ---':
                mapping[b_col] = sel

    st.divider()
    if st.button("🚀 一键生成最终表格", type="primary"):
        with st.spinner("执行：C→A匹配 → 填充B模板..."):
            df_c_a = matcher.match_c_to_a(df_c, key_c, key_a)
            final_result = matcher.fill_b_template(df_c_a, df_b, mapping)
            
            st.success("✅ 生成完成！")
            st.dataframe(final_result, use_container_width=True)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as w:
                final_result.to_excel(w, index=False)
            
            st.download_button("📥 下载结果", data=output.getvalue(), file_name=f"结果_{selected_b}")

else:
    st.info("""
    📌 使用流程：
    1. 选择B固定模板
    2. 上传A表（数据源）+ C表（主键列表）
    3. 配置映射 → 一键生成下载
    """)