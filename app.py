import streamlit as st
import pandas as pd
import json
import io
import os
import zipfile
from datetime import datetime

# ======================== 【双密码独立配置】 ========================
LOGIN_PASSWORD = "123456"       # 普通用户登录密码
ADMIN_PASSWORD = "888888"     # 管理员模板管理密码
# ==================================================================

# 系统登录校验
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
    try:
        return pd.read_excel(file, engine="openpyxl", dtype=str)
    except:
        st.error(f"文件读取失败：{file.name}，请检查文件格式！")
        return None

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
# 字段映射 保存/加载 + 【优化1：同名字段自动映射】
# ===========================
def auto_map_columns(df_b, df_a):
    """自动映射B模板和A表相同列名的字段"""
    auto_mapping = {}
    b_cols = list(df_b.columns)
    a_cols = list(df_a.columns)
    for col in b_cols:
        if col in a_cols:
            auto_mapping[col] = col
    return auto_mapping

def load_mapping(template_name, df_b, df_a):
    """优先加载保存的映射，无则自动匹配同名列"""
    path = os.path.join(MAPPING_FOLDER, f"{template_name}.json")
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    # 无保存映射 → 自动匹配同名列
    return auto_map_columns(df_b, df_a)

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
# 【优化2：批量生成打包ZIP】
# ===========================
def create_zip(files_data):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for filename, data in files_data.items():
            zf.writestr(filename, data)
    zip_buffer.seek(0)
    return zip_buffer

# ===========================
# 界面布局
# ===========================
st.set_page_config(page_title="三表Excel生成器", layout="wide")
st.title("📊 A+B+C三表一键生成工具（自动映射+批量版）")

# ===========================
# 侧边栏：双密码权限控制
# ===========================
with st.sidebar:
    st.header("1. B模板管理")
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
    selected_b_list = []

    if admin_auth:
        with st.expander("📤 上传新B模板"):
            up_b = st.file_uploader("选择模板文件", type=['xlsx','xls'])
            if up_b and st.button("保存模板"):
                save_b_template(up_b)
                st.success("上传成功！")
                st.rerun()

        if template_list:
            selected_b = st.selectbox("选择B模板（单生成）", template_list)
            selected_b_list = st.multiselect("多选B模板（批量生成）", template_list)
            if st.button("🗑️ 删除选中模板"):
                delete_b_template(selected_b)
                st.success(f"已删除：{selected_b}")
                st.rerun()
    else:
        if template_list:
            selected_b = st.selectbox("选择B模板（单生成）", template_list)
            selected_b_list = st.multiselect("多选B模板（批量生成）", template_list)
        else:
            st.info("暂无可用模板")

    st.header("2. 数据上传")
    up_a = st.file_uploader("A表（总数据源）", type=['xlsx','xls'])
    # 支持上传多个C表
    up_c_list = st.file_uploader("C表（主键列表，可多选）", type=['xlsx','xls'], accept_multiple_files=True)

# ===========================
# 主功能区
# ===========================
if up_a and up_c_list:
    df_a = read_excel(up_a)
    if df_a is None:
        st.stop()

    # --------------------
    # 单模板生成
    # --------------------
    if selected_b and len(up_c_list) == 1:
        df_c = read_excel(up_c_list[0])
        df_b = read_excel(os.path.join(TEMPLATE_FOLDER, selected_b))
        if df_c is None or df_b is None:
            st.stop()

        col1, col2, col3 = st.columns(3)
        with col1: st.subheader("A表"); st.dataframe(df_a.head(3), use_container_width=True)
        with col2: st.subheader("C表"); st.dataframe(df_c.head(3), use_container_width=True)
        with col3: st.subheader("B模板"); st.dataframe(df_b.head(3), use_container_width=True)

        matcher = DataMatcher(df_a)
        st.divider()

        st.subheader("⚙️ 主键匹配")
        key_c = st.selectbox("C表主键", df_c.columns)
        key_a = st.selectbox("A表对应主键", df_a.columns)

        # 自动映射字段
        st.subheader("🔗 字段映射（自动匹配同名列+可修改）")
        saved_map = load_mapping(selected_b, df_b, df_a)
        current_map = {}

        map_cols = st.columns(3)
        for i, b_col in enumerate(df_b.columns):
            opts = ['--- 不填充 ---'] + list(df_a.columns)
            default_val = saved_map.get(b_col, '--- 不填充 ---')
            default_idx = opts.index(default_val) if default_val in opts else 0
            with map_cols[i % 3]:
                sel = st.selectbox(f"B→{b_col}", opts, index=default_idx)
                if sel != '--- 不填充 ---':
                    current_map[b_col] = sel

        st.divider()
        col_save, col_gen = st.columns(2)
        with col_save:
            if st.button("💾 保存当前映射"):
                save_mapping(selected_b, current_map)
                st.success("映射已永久保存！")
        with col_gen:
            start_gen = st.button("🚀 一键生成表格", type="primary")

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

            out = io.BytesIO()
            with pd.ExcelWriter(out, engine="openpyxl") as w:
                result.to_excel(w, index=False)
            st.download_button("📥 下载结果", out.getvalue(), f"结果_{selected_b}")

    # --------------------
    # 【优化2：批量生成（核心）】
    # --------------------
    if selected_b_list and len(up_c_list) > 0:
        st.divider()
        st.subheader("🚀 批量生成模式（一键生成所有模板+对应C表）")
        st.info(f"已选 {len(selected_b_list)} 个B模板 | 已上传 {len(up_c_list)} 个C表\n✅ 程序将自动按顺序匹配：模板1→C表1，模板2→C表2...")

        key_c = st.text_input("C表统一主键列名", value="订单编号", help="所有C表的主键列名必须一致")
        key_a = st.selectbox("A表对应主键列", df_a.columns)

        if st.button("✅ 一键批量生成并打包下载", type="primary"):
            matcher = DataMatcher(df_a)
            bar = st.progress(0)
            status = st.empty()
            files_data = {}

            # 循环批量生成
            total = len(selected_b_list)
            for idx, (b_file, c_file) in enumerate(zip(selected_b_list, up_c_list)):
                progress = int((idx+1)/total * 100)
                status.text(f"正在生成：{b_file} → {c_file.name}")
                bar.progress(progress)

                # 读取文件
                df_c = read_excel(c_file)
                df_b = read_excel(os.path.join(TEMPLATE_FOLDER, b_file))
                if df_c is None or df_b is None:
                    continue

                # 加载映射+生成
                mapping = load_mapping(b_file, df_b, df_a)
                matched = matcher.match_c_to_a(df_c, key_c, key_a)
                result = matcher.fill_b_template(matched, df_b, mapping)

                # 写入内存
                out = io.BytesIO()
                with pd.ExcelWriter(out, engine="openpyxl") as w:
                    result.to_excel(w, index=False)
                files_data[f"结果_{b_file}_{c_file.name}"] = out.getvalue()

            # 打包ZIP
            zip_file = create_zip(files_data)
            bar.progress(100)
            status.text("✅ 全部生成完成！")
            st.success(f"成功生成 {len(files_data)} 个文件")
            st.download_button(
                "📦 下载全部文件(ZIP)",
                zip_file,
                f"批量生成结果_{datetime.now().strftime('%Y%m%d%H%M')}.zip"
            )

else:
    st.info("""
    🎯 新功能说明：
    1. 🚀 自动映射：B/A表同名列自动绑定，不用手动配置
    2. 📦 批量生成：多选B模板+上传多个C表，一键生成所有文件
    3. 🔐 双密码：登录密码（使用）/管理员密码（配置模板）
    """)