import streamlit as st
import pandas as pd
import json
import io
import os

# ======================== 【双密码独立配置】 ========================
# 1. 登录密码：所有用户登录系统使用（只能生成文件，不能改模板）
LOGIN_PASSWORD = "888888"
# 2. 管理员密码：仅用于上传/删除B模板（独立密码，不通用）
ADMIN_PASSWORD = "666666"
# ==================================================================

# 系统登录校验（通用密码）
def check_system_login():
    if "is_logged_in" not in st.session_state:
        st.session_state.is_logged_in = False
    if not st.session_state.is_logged_in:
        st.title("🔒 系统登录")
        pwd = st.text_input("请输入登录密码", type="password")
        if st.button("登录系统"):
            if pwd == LOGIN_PASSWORD:
                st.session_state.is_logged_in = True
                st.success("登录成功！")
                st.rerun()
            else:
                st.error("登录密码错误！")
        st.stop()

check_system_login()

# ===========================
# 自动创建文件夹
# ===========================
TEMPLATE_FOLDER = "b_templates"
MAPPING_FOLDER = "template_mappings"
for folder in [TEMPLATE_FOLDER, MAPPING_FOLDER]:
    if not os.path.exists(folder):
        os.makedirs(folder)

# ===========================
# 文本读取（防ID尾数变0）
# ===========================
@st.cache_data(ttl=3600)
def read_excel(file):
    return pd.read_excel(file, engine="openpyxl", dtype=str)

# ===========================
# B模板管理函数
# ===========================
def get_b_templates():
    return [f for f in os.listdir(TEMPLATE_FOLDER) if f.endswith((".xlsx", ".xls"))]

def save_b_template(uploaded_file):
    path = os.path.join(TEMPLATE_FOLDER, uploaded_file.name)
    with open(path, "wb") as f:
        f.write(uploaded_file.getbuffer())

def delete_b_template(template_name):
    file_path = os.path.join(TEMPLATE_FOLDER, template_name)
    map_path = os.path.join(MAPPING_FOLDER, f"{template_name}.json")
    if os.path.exists(file_path):
        os.remove(file_path)
    if os.path.exists(map_path):
        os.remove(map_path)

# ===========================
# 字段映射 保存/加载
# ===========================
def load_mapping(template_name):
    path = os.path.join(MAPPING_FOLDER, f"{template_name}.json")
    return json.load(open(path, "r", encoding="utf-8")) if os.path.exists(path) else {}

def save_mapping(template_name, mapping):
    with open(os.path.join(MAPPING_FOLDER, f"{template_name}.json"), "w", encoding="utf-8") as f:
        json.dump(mapping, f, ensure_ascii=False, indent=2)

# ===========================
# 三表匹配引擎
# ===========================
class DataMatcher:
    def __init__(self, df_a):
        self.df_a = df_a.copy().astype(str).apply(lambda x: x.str.strip())

    def match_c_to_a(self, df_c, key_c, key_a):
        df_c_clean = df_c.copy().astype(str).apply(lambda x: x.str.strip())
        return pd.merge(df_c_clean[[key_c]], self.df_a, left_on=key_c, right_on=key_a, how="left")

    def fill_b_template(self, df_matched, df_b_template, mapping):
        df_b = df_b_template.copy().astype(str).apply(lambda x: x.str.strip())
        for b_col, a_col in mapping.items():
            if b_col in df_b.columns and a_col in df_matched.columns:
                df_b[b_col] = df_matched[a_col]
        return df_b

# ===========================
# 界面布局
# ===========================
st.set_page_config(page_title="三表Excel生成器", layout="wide")
st.title("📊一键生成B数据表模版")

# ===========================
# 侧边栏：双密码权限控制
# ===========================
with st.sidebar:
    st.header("1. B模板管理")
    # 管理员验证：独立密码，验证通过才可管理模板
    admin_auth = False
    with st.expander("🔑 管理员操作（上传/删除模板）"):
        admin_pwd = st.text_input("请输入管理员密码", type="password", key="admin")
        if admin_pwd == ADMIN_PASSWORD:
            admin_auth = True
            st.success("✅ 管理员已授权")
        elif admin_pwd:
            st.error("❌ 管理员密码错误")

    template_list = get_b_templates()
    selected_b = None

    # 管理员功能区
    if admin_auth:
        with st.expander("📤 上传新B模板"):
            up_b = st.file_uploader("选择模板文件", type=['xlsx','xls'])
            if up_b and st.button("保存模板"):
                save_b_template(up_b)
                st.success("上传成功！")
                st.rerun()

        if template_list:
            selected_b = st.selectbox("选择B模板", template_list)
            # 无确认直接删除（仅管理员）
            if st.button("🗑️ 删除选中模板"):
                delete_b_template(selected_b)
                st.success(f"已删除：{selected_b}")
                st.rerun()
    else:
        # 普通用户：仅选择模板，无管理权限
        if template_list:
            selected_b = st.selectbox("选择B模板", template_list)
        else:
            st.info("暂无可用模板")

    st.header("2. 数据上传")
    up_a = st.file_uploader("A表（总数据源）", type=['xlsx','xls'])
    up_c = st.file_uploader("C表（主键列表）", type=['xlsx','xls'])

# ===========================
# 主功能区
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

    # 主键配置
    st.subheader("⚙️ 主键匹配")
    key_c = st.selectbox("C表主键", df_c.columns)
    key_a = st.selectbox("A表对应主键", df_a.columns)

    # 字段映射（自动加载）
    st.subheader("🔗 字段映射（自动保存/加载）")
    saved_map = load_mapping(selected_b)
    current_map = {}

    map_cols = st.columns(3)
    for i, b_col in enumerate(df_b.columns):
        opts = ['--- 不填充 ---'] + list(df_a.columns)
        default_idx = opts.index(saved_map.get(b_col, '--- 不填充 ---')) if saved_map.get(b_col) in opts else 0
        with map_cols[i % 3]:
            sel = st.selectbox(f"B→{b_col}", opts, index=default_idx)
            if sel != '--- 不填充 ---':
                current_map[b_col] = sel

    # 操作按钮
    st.divider()
    col_save, col_gen = st.columns(2)
    with col_save:
        if st.button("💾 保存当前映射"):
            save_mapping(selected_b, current_map)
            st.success("映射已永久保存！")
    with col_gen:
        start_gen = st.button("🚀 一键生成表格", type="primary")

    # ===========================
    # 生成进度条
    # ===========================
    if start_gen:
        bar = st.progress(0)
        text = st.empty()

        text.text("步骤1：C表匹配A表数据")
        bar.progress(30)
        matched = matcher.match_c_to_a(df_c, key_c, key_a)

        text.text("步骤2：填充数据到B模板")
        bar.progress(70)
        result = matcher.fill_b_template(matched, df_b, current_map)

        bar.progress(100)
        text.text("✅ 生成完成！")
        st.dataframe(result, use_container_width=True)

        # 下载
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as w:
            result.to_excel(w, index=False)
        st.download_button("📥 下载结果", out.getvalue(), f"结果_{selected_b}")

else:
    st.info("""
    📌 使用说明：
    1. 输入登录密码 → 登录系统
    2. 需管理模板 → 输入管理员密码
    3. 选择模板 → 配置映射 → 保存
    4. 上传A+C表 → 一键生成
    """)