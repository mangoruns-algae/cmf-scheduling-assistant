import base64
import hmac
from datetime import datetime
from pathlib import Path

import streamlit as st

from excel_io import export_schedule_to_excel, read_input_excel
from schedule_viewer import (
    available_people,
    export_records_pdf,
    export_schedule_pdf,
    generated_schedule_records,
    load_schedule_workbook,
    matching_records,
    people_from_records,
    render_records_html,
    render_schedule_html,
)
from scheduler import generate_schedule


st.set_page_config(page_title="CMF 排班助手", page_icon="🤖", layout="wide")


COLUMN_LABELS = {
    "employee_id": "员工编号",
    "employee_name": "员工姓名",
    "home_line": "所属产线",
    "role": "岗位角色",
    "skills": "技能",
    "qualified_lines": "资质产线",
    "can_night_shift": "可上夜班",
    "can_cross_line": "可跨线",
    "max_consecutive_days": "最大连续工作天数",
    "weekly_max_hours": "每周工时上限",
    "enabled": "是否启用",
    "remark": "备注",
    "date": "日期",
    "availability_status": "可用状态",
    "available_shift": "可用班次",
    "fixed_line": "固定产线",
    "reason": "原因",
    "production_line": "生产线",
    "shift": "班次",
    "batch_id": "批次号",
    "task_type": "任务类型",
    "required_role": "所需角色",
    "required_skill": "所需技能",
    "required_headcount": "需求人数",
    "priority": "优先级",
    "allow_borrowing": "允许借调",
    "rule_code": "规则编码",
    "rule_name": "规则名称",
    "rule_value": "规则值",
    "value_type": "值类型",
    "scope": "适用范围",
    "assignment_type": "安排类型",
    "status": "排班状态",
    "conflict_message": "冲突说明",
}

CREATOR_USERS = {"qingyuan_qin", "admin"}


@st.cache_resource
def locked_schedule_store():
    """Process-wide published schedule shared by all active Streamlit sessions."""
    return {}


def can_create(username):
    return str(username).casefold() in CREATOR_USERS


def table_config(frame):
    """Return display-only column labels without changing the underlying data."""
    return {column: COLUMN_LABELS[column] for column in frame.columns if column in COLUMN_LABELS}


def table_height(frame, maximum=520):
    return min(maximum, max(180, 44 + len(frame) * 35))


def render_footer():
    st.markdown(
        """
        <div class="cmf-footer">
            版本 0.1&nbsp;&nbsp;·&nbsp;&nbsp;更新日期 2026.08.21&nbsp;&nbsp;·&nbsp;&nbsp;Powered by Qingyuan
        </div>
        """,
        unsafe_allow_html=True,
    )


def authenticate_user(username, password):
    """Return the configured username on success; username matching is case-insensitive."""
    normalized_username = username.strip().casefold().encode("utf-8")
    accounts = st.secrets.get("accounts", {})

    for configured_username, configured_password in accounts.items():
        username_ok = hmac.compare_digest(
            normalized_username,
            str(configured_username).casefold().encode("utf-8"),
        )
        password_ok = hmac.compare_digest(password, str(configured_password))
        if username_ok and password_ok:
            return str(configured_username)

    return None

logo_base64 = base64.b64encode(Path("assets/henlius-logo.png").read_bytes()).decode("ascii")
st.markdown(
    f"""
    <style>
        .cmf-header {{
            display: flex;
            align-items: center;
            flex-wrap: wrap;
            gap: 24px;
            min-height: 94px;
            margin: 0 0 12px 0;
            padding: 8px 0;
            overflow: visible;
        }}
        .cmf-header img {{
            display: block;
            width: clamp(300px, 27vw, 370px);
            max-width: 100%;
            height: auto;
            flex: 0 0 auto;
            object-fit: contain;
        }}
        .cmf-header h1 {{
            margin: 0;
            padding: 4px 0 6px;
            color: #262730;
            font-size: clamp(36px, 3.2vw, 44px);
            line-height: 1.25;
            font-weight: 700;
            flex: 0 1 auto;
            white-space: normal;
            overflow: visible;
        }}
        .cmf-intro {{
            margin: 0 0 24px 0;
            color: #68707d;
            font-size: 14px;
            line-height: 1.6;
        }}
        .cmf-steps {{
            display: grid;
            grid-template-columns: repeat(4, minmax(0, 1fr));
            gap: 10px;
            margin: 4px 0 24px 0;
        }}
        .cmf-step {{
            padding: 12px 14px;
            border: 1px solid #e4e8ef;
            border-radius: 10px;
            background: #f8fafc;
            color: #414957;
            font-size: 14px;
            font-weight: 600;
            line-height: 1.35;
        }}
        .cmf-step strong {{
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 22px;
            height: 22px;
            margin-right: 7px;
            border-radius: 50%;
            background: #e8f1ff;
            color: #1f5fae;
            font-size: 12px;
        }}
        @media (max-width: 700px) {{
            .cmf-header {{ align-items: flex-start; flex-direction: column; gap: 8px; min-height: auto; }}
            .cmf-header img {{ width: min(320px, 92vw); }}
            .cmf-header h1 {{ font-size: 34px; line-height: 1.25; }}
            .cmf-steps {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
        }}
        .cmf-footer {{
            margin-top: 36px;
            padding: 16px 12px 4px;
            border-top: 1px solid rgba(49, 51, 63, 0.10);
            color: #8b8f99;
            font-size: 12px;
            line-height: 1.4;
            text-align: center;
        }}
        .block-container {{ max-width: 1320px; padding-top: 2rem; padding-bottom: 2rem; }}
        .schedule-wrap {{
            max-height: 680px;
            overflow: auto;
            border: 1px solid #dfe5ed;
            border-radius: 12px;
            background: #ffffff;
        }}
        .schedule-table {{
            width: 100%;
            min-width: 1050px;
            border-collapse: separate;
            border-spacing: 0;
            table-layout: fixed;
            font-size: 13px;
            line-height: 1.45;
        }}
        .schedule-table .date-col {{ width: 11%; }}
        .schedule-table .time-col {{ width: 12%; }}
        .schedule-table .work-col {{ width: 36%; }}
        .schedule-table .lead-col {{ width: 14%; }}
        .schedule-table .people-col {{ width: 27%; }}
        .schedule-cell {{
            padding: 9px 10px;
            border-right: 1px solid #dfe5ed;
            border-bottom: 1px solid #dfe5ed;
            vertical-align: middle;
            overflow-wrap: anywhere;
        }}
        .schedule-header {{
            position: sticky;
            top: 0;
            z-index: 2;
            background: #d9eaf7 !important;
            color: #17365d !important;
        }}
        .schedule-match {{
            background: #fff2b2 !important;
            box-shadow: inset 0 2px #e0a800, inset 0 -2px #e0a800;
        }}
        .schedule-leader-match {{
            background: #dcebff !important;
            box-shadow: inset 0 0 0 2px #2f6fed;
            color: #174a9c !important;
            font-weight: 700 !important;
        }}
        div[data-testid="stMetric"] {{
            min-height: 104px;
            padding: 16px 18px;
            border: 1px solid #e5e9f0;
            border-radius: 12px;
            background: #ffffff;
        }}
        div[data-testid="stButton"] button,
        div[data-testid="stFormSubmitButton"] button,
        div[data-testid="stDownloadButton"] button {{
            transition: transform 140ms cubic-bezier(0.23, 1, 0.32, 1);
        }}
        div[data-testid="stButton"] button:active,
        div[data-testid="stFormSubmitButton"] button:active,
        div[data-testid="stDownloadButton"] button:active {{
            transform: scale(0.98);
        }}
        @media (prefers-reduced-motion: reduce) {{
            div[data-testid="stButton"] button,
            div[data-testid="stFormSubmitButton"] button,
            div[data-testid="stDownloadButton"] button {{ transition: none; }}
        }}
    </style>
    <div class="cmf-header">
        <img src="data:image/png;base64,{logo_base64}" alt="Henlius 复宏汉霖 Logo">
        <h1>CMF 排班助手</h1>
    </div>
    """,
    unsafe_allow_html=True,
)

if not st.session_state.get("authenticated", False):
    st.markdown('<p class="cmf-intro">安全登录后即可上传排班数据并生成班表。</p>', unsafe_allow_html=True)
    left_space, login_col, right_space = st.columns([1, 0.9, 1])
    with login_col:
        with st.container(border=True):
            st.subheader("账户登录")
            st.caption("请输入已授权的账户信息")
            with st.form("login_form", clear_on_submit=False):
                username = st.text_input("账户", placeholder="请输入账户")
                password = st.text_input("密码", type="password", placeholder="请输入密码")
                submitted = st.form_submit_button("登录", type="primary", use_container_width=True)

            if submitted:
                authenticated_username = authenticate_user(username, password)
                if authenticated_username is not None:
                    st.session_state["authenticated"] = True
                    st.session_state["username"] = authenticated_username
                    st.rerun()
                else:
                    st.error("账户或密码错误，请检查后重试。")
    render_footer()
    st.stop()

with st.sidebar:
    current_username = st.session_state.get("username", "已登录用户")
    creator_access = can_create(current_username)
    st.caption(f"当前账户：{current_username}")
    st.markdown("#### 权限清单")
    st.markdown("✅ 排班预览")
    st.markdown("✅ PDF 导出")
    st.markdown("✅ 个人班次定位")
    st.markdown("✅ 创建与发布排班" if creator_access else "🔒 创建与发布排班")
    published = locked_schedule_store()
    if published:
        st.caption(f"已发布：{published.get('name', '当前班表')}")
    else:
        st.caption("当前暂无已发布班表")
    st.divider()
    if st.button("退出登录", use_container_width=True):
        st.session_state.clear()
        st.rerun()

if creator_access:
    mode = st.radio(
        "工作模式",
        ["排班创建者", "排班预览者"],
        horizontal=True,
        key="work_mode",
        help="创建者用于生成并发布新班表；预览者用于只读查阅、个人定位和 PDF 导出。",
    )
else:
    mode = "排班预览者"
    st.markdown("#### 排班预览者")
    st.caption("当前账户为只读权限，已自动进入排班预览。")


def render_creator_mode():
    st.markdown(
        """
        <p class="cmf-intro">上传排班数据，依次核对人员、需求和规则后生成班表。</p>
        <div class="cmf-steps">
            <div class="cmf-step"><strong>1</strong>上传数据</div>
            <div class="cmf-step"><strong>2</strong>检查基础信息</div>
            <div class="cmf-step"><strong>3</strong>生成排班</div>
            <div class="cmf-step"><strong>4</strong>下载结果</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.subheader("上传排班数据")
    st.caption("请选择基于 staff_scheduling_template.xlsx 填写的 Excel 文件。")
    uploaded = st.file_uploader("上传排班数据 Excel", type=["xlsx"], label_visibility="collapsed", key="creator_upload")
    if uploaded is None:
        st.info("等待上传排班数据。上传成功后可检查人员、需求和规则。")
        return

    try:
        data = read_input_excel(uploaded)
    except Exception as exc:
        st.error(str(exc))
        return

    st.success(f"已读取 {uploaded.name}，请核对数据后生成排班。")
    staff = data["人员信息"]
    availability = data["人员可用性"]
    demand = data["用工需求"]
    rules = data["排班规则"]
    tab1, tab2, tab3, tab4 = st.tabs(["1  人员资源", "2  用工需求", "3  排班规则", "4  排班结果"])

    with tab1:
        c1, c2, c3 = st.columns(3)
        c1.metric("启用人员", int((staff["enabled"].astype(str).str.upper() == "Y").sum()))
        c2.metric("所属产线", staff["home_line"].nunique())
        c3.metric("可跨线人员", int((staff["can_cross_line"].astype(str).str.upper() == "Y").sum()))
        st.markdown("#### 人员信息")
        st.dataframe(staff, use_container_width=True, hide_index=True, height=table_height(staff), column_config=table_config(staff))
        with st.expander(f"人员可用性（{len(availability)} 条）"):
            st.dataframe(availability, use_container_width=True, hide_index=True, height=table_height(availability, maximum=420), column_config=table_config(availability))

    with tab2:
        st.caption(f"共 {len(demand)} 条用工需求")
        st.dataframe(demand, use_container_width=True, hide_index=True, height=table_height(demand), column_config=table_config(demand))

    with tab3:
        st.caption(f"共 {len(rules)} 条排班规则")
        st.dataframe(rules, use_container_width=True, hide_index=True, height=table_height(rules), column_config=table_config(rules))

    with tab4:
        action_col, _ = st.columns([1, 3])
        with action_col:
            if st.button("生成排班", type="primary", use_container_width=True):
                st.session_state["result"] = generate_schedule(staff, availability, demand, rules)
        result = st.session_state.get("result")
        if result is None:
            st.info("点击“生成排班”开始计算。")
            return
        assigned = int((result["status"] == "已排班").sum()) if not result.empty else 0
        borrowed = int((result["assignment_type"] == "跨产线借调").sum()) if not result.empty else 0
        shortage = int((result["status"] == "缺员").sum()) if not result.empty else 0
        c1, c2, c3 = st.columns(3)
        c1.metric("已完成排班", assigned)
        c2.metric("跨产线借调", borrowed)
        c3.metric("缺员", shortage)
        st.markdown("#### 排班明细")
        st.dataframe(result, use_container_width=True, hide_index=True, height=table_height(result), column_config=table_config(result))
        view = result[result["status"] == "已排班"].copy()
        if not view.empty:
            view["安排"] = view["production_line"].astype(str) + " / " + view["shift"].astype(str) + " / " + view["task_type"].astype(str)
            matrix = view.pivot_table(index="employee_name", columns="date", values="安排", aggfunc=lambda values: "；".join(values))
            st.subheader("人员 × 日期 班表视图")
            st.caption("横向滚动查看全部日期；每个单元格依次显示产线、班次和任务。")
            st.dataframe(matrix, use_container_width=True, height=table_height(matrix, maximum=460))
        output = export_schedule_to_excel(result, data)
        download_col, publish_col, _ = st.columns([1, 1, 2])
        with download_col:
            st.download_button("下载排班结果 Excel", data=output, file_name="排班结果.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        with publish_col:
            if st.button("锁定并发布班表", type="primary", use_container_width=True):
                store = locked_schedule_store()
                store.clear()
                store.update({
                    "kind": "generated",
                    "name": "排班结果",
                    "records": generated_schedule_records(result),
                    "excel": output,
                    "locked_by": current_username,
                    "locked_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
                })
                st.success("当前班表已锁定并发布，预览者现在可以直接查看。")


def render_viewer_mode():
    st.markdown('<p class="cmf-intro">只读查阅已排好的生产班表，按姓名定位个人任务，并导出便于传阅的 PDF。</p>', unsafe_allow_html=True)
    store = locked_schedule_store()

    if creator_access:
        with st.expander("发布已有排班 Excel", expanded=not bool(store)):
            st.caption("上传现有生产排班模板，确认预览无误后可锁定为所有预览者看到的版本。")
            uploaded = st.file_uploader("上传已排好的生产排班 Excel", type=["xlsx"], label_visibility="collapsed", key="viewer_upload")
            if uploaded is not None:
                try:
                    candidate = load_schedule_workbook(uploaded.getvalue())
                    candidate_sheet_name = st.selectbox("待发布工作表", candidate.schedule_sheets, key="candidate_sheet")
                    if st.button("锁定并发布这个排班", type="primary", key="publish_uploaded"):
                        store.clear()
                        store.update({
                            "kind": "workbook",
                            "name": uploaded.name,
                            "bytes": uploaded.getvalue(),
                            "sheet": candidate_sheet_name,
                            "locked_by": current_username,
                            "locked_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
                        })
                        st.success("排班已锁定并发布。")
                        st.rerun()
                except Exception as exc:
                    st.error(f"无法读取排班文件：{exc}")

    if not store:
        st.info("当前还没有创建者锁定并发布的排班，请稍后再查看。")
        return

    st.subheader("已发布班次")
    st.caption(f"{store.get('name')} · 发布人 {store.get('locked_by')} · 发布时间 {store.get('locked_at')}")

    if store.get("kind") == "workbook":
        try:
            schedule = load_schedule_workbook(store["bytes"])
            selected_sheet = store.get("sheet") if store.get("sheet") in schedule.schedule_sheets else schedule.schedule_sheets[0]
            sheet = schedule.workbook[selected_sheet]
        except Exception as exc:
            st.error(f"已发布排班无法读取：{exc}")
            return
        people = available_people(sheet)
        selected_person = st.selectbox("人员定位", [""] + people, format_func=lambda value: value or "全部人员（不高亮）", key="published_person")
        matches = matching_records(sheet, selected_person)
        st.markdown("#### 排班预览")
        if selected_person:
            st.caption("黄色整行表示本人参与；蓝色日期表示本人担任现场负责人。未参与的同日工序不会高亮。")
        st.markdown(render_schedule_html(sheet, selected_person), unsafe_allow_html=True)
        try:
            pdf_bytes = export_schedule_pdf(sheet, selected_person)
            filename_suffix = f"-{selected_person}" if selected_person else ""
            st.download_button("导出单页排班 PDF", data=pdf_bytes, file_name=f"{selected_sheet}{filename_suffix}-排班预览.pdf", mime="application/pdf", type="primary")
        except Exception as exc:
            st.error(f"PDF 生成失败：{exc}")
    else:
        records = store.get("records", [])
        people = people_from_records(records)
        selected_person = st.selectbox("人员定位", [""] + people, format_func=lambda value: value or "全部人员（不高亮）", key="generated_person")
        st.markdown("#### 排班预览")
        if selected_person:
            st.caption("黄色整行表示本人参与；蓝色日期表示本人担任现场负责人。未参与的同日工序不会高亮。")
        st.markdown(render_records_html(records, selected_person), unsafe_allow_html=True)
        try:
            pdf_bytes = export_records_pdf(records, store.get("name", "排班结果"), selected_person)
            filename_suffix = f"-{selected_person}" if selected_person else ""
            download_pdf, download_excel, _ = st.columns([1, 1, 2])
            with download_pdf:
                st.download_button("导出单页排班 PDF", data=pdf_bytes, file_name=f"排班结果{filename_suffix}.pdf", mime="application/pdf", type="primary", use_container_width=True)
            with download_excel:
                st.download_button("下载排班 Excel", data=store.get("excel", b""), file_name="排班结果.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        except Exception as exc:
            st.error(f"PDF 生成失败：{exc}")


if mode == "排班创建者":
    render_creator_mode()
else:
    render_viewer_mode()

render_footer()
