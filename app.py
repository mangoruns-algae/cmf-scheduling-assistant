import base64
import hmac
from pathlib import Path

import streamlit as st

from excel_io import export_schedule_to_excel, read_input_excel
from scheduler import generate_schedule


st.set_page_config(page_title="CMF 排班助手", page_icon="🤖", layout="wide")

logo_base64 = base64.b64encode(Path("assets/henlius-logo.png").read_bytes()).decode("ascii")
st.markdown(
    f"""
    <style>
        .cmf-header {{
            display: flex;
            align-items: center;
            gap: 32px;
            min-height: 76px;
            margin: 0 0 14px 0;
        }}
        .cmf-header img {{
            display: block;
            width: 330px;
            height: auto;
            flex: 0 0 auto;
        }}
        .cmf-header h1 {{
            margin: 0;
            padding: 0;
            color: #262730;
            font-size: 34px;
            line-height: 1;
            font-weight: 700;
            white-space: nowrap;
        }}
        .cmf-header .cmf-latin {{
            display: inline-block;
            font-size: 39px;
            letter-spacing: 0.3px;
            transform: translateY(1px);
        }}
        @media (max-width: 700px) {{
            .cmf-header {{ gap: 18px; }}
            .cmf-header img {{ width: 230px; }}
            .cmf-header h1 {{ font-size: 27px; }}
            .cmf-header .cmf-latin {{ font-size: 31px; }}
        }}
        .cmf-footer {{
            position: fixed;
            left: 0;
            right: 0;
            bottom: 0;
            z-index: 999;
            padding: 8px 16px;
            border-top: 1px solid rgba(49, 51, 63, 0.10);
            background: rgba(255, 255, 255, 0.92);
            color: #8b8f99;
            font-size: 12px;
            line-height: 1.4;
            text-align: center;
            backdrop-filter: blur(6px);
        }}
        .block-container {{ padding-bottom: 52px; }}
    </style>
    <div class="cmf-header">
        <img src="data:image/png;base64,{logo_base64}" alt="Henlius 复宏汉霖 Logo">
        <h1><span class="cmf-latin">CMF</span> 排班助手</h1>
    </div>
    <div class="cmf-footer">
        版本 0.1&nbsp;&nbsp;·&nbsp;&nbsp;更新日期 2026.08.21&nbsp;&nbsp;·&nbsp;&nbsp;Powered by Qingyuan
    </div>
    """,
    unsafe_allow_html=True,
)

if not st.session_state.get("authenticated", False):
    left_space, login_col, right_space = st.columns([1, 1.15, 1])
    with login_col:
        st.subheader("账户登录")
        with st.form("login_form", clear_on_submit=False):
            username = st.text_input("账户", placeholder="请输入账户")
            password = st.text_input("密码", type="password", placeholder="请输入密码")
            submitted = st.form_submit_button("登录", type="primary", use_container_width=True)

        if submitted:
            username_ok = hmac.compare_digest(username, st.secrets["auth"]["username"])
            password_ok = hmac.compare_digest(password, st.secrets["auth"]["password"])
            if username_ok and password_ok:
                st.session_state["authenticated"] = True
                st.rerun()
            else:
                st.error("账户或密码错误")
    st.stop()

with st.sidebar:
    st.caption(f"当前账户：{st.secrets['auth']['username']}")
    if st.button("退出登录", use_container_width=True):
        st.session_state.clear()
        st.rerun()

st.caption("MVP：人员数据 + 可用性 + 生产用工需求 + 排班规则 → 自动生成班表")

uploaded = st.file_uploader("上传排班数据 Excel", type=["xlsx"])
if uploaded is None:
    st.info("请先上传 staff_scheduling_template.xlsx")
    st.stop()

try:
    data = read_input_excel(uploaded)
except Exception as exc:
    st.error(str(exc))
    st.stop()

staff = data["人员信息"]
availability = data["人员可用性"]
demand = data["用工需求"]
rules = data["排班规则"]

tab1, tab2, tab3, tab4 = st.tabs(["人员资源", "用工需求", "排班规则", "排班结果"])

with tab1:
    c1, c2, c3 = st.columns(3)
    c1.metric("启用人员", int((staff["enabled"].astype(str).str.upper() == "Y").sum()))
    c2.metric("所属产线", staff["home_line"].nunique())
    c3.metric("可跨线人员", int((staff["can_cross_line"].astype(str).str.upper() == "Y").sum()))
    st.dataframe(staff, use_container_width=True, hide_index=True)

with tab2:
    st.dataframe(demand, use_container_width=True, hide_index=True)

with tab3:
    st.dataframe(rules, use_container_width=True, hide_index=True)

with tab4:
    if st.button("🚀 生成排班", type="primary", use_container_width=True):
        st.session_state["result"] = generate_schedule(staff, availability, demand, rules)

    result = st.session_state.get("result")
    if result is None:
        st.info("点击“生成排班”开始计算。")
    else:
        assigned = int((result["status"] == "已排班").sum()) if not result.empty else 0
        borrowed = int((result["assignment_type"] == "跨产线借调").sum()) if not result.empty else 0
        shortage = int((result["status"] == "缺员").sum()) if not result.empty else 0

        c1, c2, c3 = st.columns(3)
        c1.metric("已完成排班", assigned)
        c2.metric("跨产线借调", borrowed)
        c3.metric("缺员", shortage)
        st.dataframe(result, use_container_width=True, hide_index=True)

        view = result[result["status"] == "已排班"].copy()
        if not view.empty:
            view["安排"] = (
                view["production_line"].astype(str)
                + " / " + view["shift"].astype(str)
                + " / " + view["task_type"].astype(str)
            )
            matrix = view.pivot_table(
                index="employee_name", columns="date", values="安排", aggfunc=lambda values: "；".join(values)
            )
            st.subheader("人员 × 日期 班表视图")
            st.dataframe(matrix, use_container_width=True)

        output = export_schedule_to_excel(result, data)
        st.download_button(
            "下载排班结果 Excel",
            data=output,
            file_name="排班结果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
