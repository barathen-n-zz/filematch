import streamlit as st
import pandas as pd
import json
import io
import os

# ======================== 【密码保护】 ========================
def check_password():
    if "password_ok" not in st.session_state:
        st.session_state.password_ok = False
    if not st.session_state.password_ok:
        st.title("🔒 请输入工具密码")
        password = st.text_input("密码", type="password")
        if st.button("登录"):
            if password == "123456":
                st.session_state.password_ok = True
                st.success("登录成功！")
                st.rerun()
            else:
                st.error("密码错误！")
        st.stop()
check_password()
# =============================================================

# ===========================
# 配置文件夹（自动创建）
# ===========================
TEMPLATE_FOLDER = "b_templates"
MAPPING_FOLDER = "template_mappings"  # 存储字段映射
for folder in [TEMPLATE_FOLDER, MAPPING_FOLDER]:
    if not os.path.exists(folder):
        os.makedirs(folder)

# ===========================
# 核心：文本读取（防ID变0）
# ===========================
@st.cache_data(ttl=3600)
def read_excel(file):
    return pd.read_excel(file, engine="openpyxl", dtype=str)

# ===========================
# B模板管理（上传+删除）
# ===========================
def get_b_templates():
    return [f for f in os.listdir(TEMPLATE_FOLDER) if f.endswith((".xlsx", ".xls"))]

def save_b_template(uploaded_file):
    path = os.path.join(TEMPLATE_FOLDER, uploaded_file.name)
    with open(path, "wb") as f:
        f.write(uploaded_file.getbuffer())

def delete_b_template(template_name):
    # 删除模板 + 同步删除映射
    file_path = os.path.join(TEMPLATE_FOLDER, template_name)
    map_path = os.path.join(MAPPING_FOLDER, f"{template_name}.json")
    if os.path.exists(file_path):
        os.remove(file_path)
    if os.path.exists(map_path):
        os.remove(map_path)
    return True

# ===========================
# 【核心新功能】字段映射 保存/加载
# ===========================
def get_mapping_file(template_name):
    return os.path.join(MAPPING_FOLDER, f"{template_name}.json")

def save_mapping(template_name, mapping):
    with open(get_mapping_file(template_name), "w", encoding="utf-8") as f:
        json.dump(mapping, f, ensure_ascii=False, indent=2)

def load_mapping(template_name):
    path = get_mapping_file(template_name)
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

# ===========================
# 三表匹配引擎
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
        return pd.merge(df_c_clean[[key_c]], self.df_a, left_on=key_c, right_on=key_a, how="left")

    def fill_b_template(self, df_matched, df_b_template, mapping):
        df_b = df_b_template.copy()
        for col in df_b.columns:
            df_b[col] = df_b[col].astype(str).str.strip()
        for b_col, a_col in mapping.items():
            if b_col in df_b.columns and a_col in df_matched.columns:
                df_b[b_col] = merged[a_col]
        return df_b

# ===========================
# 界面
# ===========================
st.set_page_config(page_title="三表Excel自动生成器", layout="wide")
st.title("📊 A+B+C三表一键生成（映射自动保存版）")

# ===========================
# 侧边栏
# ===========================
with st.sidebar:
    st.header("1. B模板管理")
    with st.expander("📤 上传新B模板"):
        up_b = st.file_uploader("上传B模板", type=['xlsx','xls'], key="up_b")
        if up_b and st.button("保存模板"):
            save_b_template(up_b)
            st.success("✅ 模板上传成功！")

    template_list = get_b_templates()
    selected_b = None
    if template_list:
        selected_b = st.selectbox("选择B模板", template_list)
        # 删除模板
        if st.button("🗑️ 删除选中模板"):
            if st.checkbox(f"⚠️ 确认删除 {selected_b}？"):
                delete_b_template(selected_b)
                st.success("✅ 删除成功！")
                st.rerun()
    else:
        st.info("请先上传B模板")

    st.header("2. 上传数据")
    up_a = st.file_uploader("A表（总数据源）", type=['xlsx','xls'])
    up_c = st.file_uploader("C表（主键列表）", type=['xlsx','xls'])

# ===========================
# 主界面：映射自动加载 + 生成
# ===========================
if up_a and up_c and selected_b:
    df_a = read_excel(up_a)
    df_c = read_excel(up_c)
    df_b = read_excel(os.path.join(TEMPLATE_FOLDER, selected_b))

    # 表格预览
    col1, col2, col3 = st.columns(3)
    with col1: st.subheader("A表"); st.dataframe(df_a.head(3), use_container_width=True)
    with col2: st.subheader("C表"); st.dataframe(df_c.head(3), use_container_width=True)
    with col3: st.subheader("B模板"); st.dataframe(df_b.head(3), use_container_width=True)

    matcher = DataMatcher(df_a)
    st.divider()
    st.subheader("⚙️ 主键匹配")
    col_key1, col_key2 = st.columns(2)
    with col_key1: key_c = st.selectbox("C表主键", df_c.columns)
    with col_key2: key_a = st.selectbox("A表对应主键", df_a.columns)

    # ===========================
    # 【核心】自动加载映射 + 配置
    # ===========================
    st.subheader("🔗 字段映射（自动保存/加载）")
    saved_mapping = load_mapping(selected_b)  # 自动加载历史映射
    current_mapping = {}

    map_cols = st.columns(3)
    for i, b_col in enumerate(df_b.columns):
        a_options = ['--- 不填充 ---'] + list(df_a.columns)
        # 自动填充已保存的映射
        default_val = saved_mapping.get(b_col, '--- 不填充 ---')
        default_idx = a_options.index(default_val) if default_val in a_options else 0
        
        with map_cols[i % 3]:
            sel = st.selectbox(f"B→{b_col}", a_options, index=default_idx)
            if sel != '--- 不填充 ---':
                current_mapping[b_col] = sel

    # 保存映射按钮
    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        if st.button("💾 保存当前映射（永久生效）", type="primary"):
            save_mapping(selected_b, current_mapping)
            st.success(f"✅ 【{selected_b}】映射已保存！")
    with col_btn2:
        if st.button("🚀 一键生成表格"):
            with st.spinner("生成中..."):
                df_c_a = matcher.match_c_to_a(df_c, key_c, key_a)
                final = matcher.fill_b_template(df_c_a, df_b, current_mapping)
                st.success("✅ 生成完成！")
                st.dataframe(final, use_container_width=True)

                # 下载
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as w:
                    final.to_excel(w, index=False)
                st.download_button("📥 下载结果", output.getvalue(), f"结果_{selected_b}")

else:
    st.info("""
    📌 使用流程：
    1. 选择B模板 → 配置字段映射 → 点击【保存映射】
    2. 下次选择该模板 → 映射自动加载！
    3. 上传A+C表 → 一键生成
    """)