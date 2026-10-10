import io
import pandas as pd
import streamlit as st
import requests
import threading
from datetime import datetime
import pytz

# ============ 钉钉机器人配置：从streamlit secrets读取，不在代码写死token ============
def ding_notify_user_action():
    """后台线程发送钉钉通知，使用国内北京时间，不阻塞主业务流程"""
    try:
        # 从secrets读取webhook地址
        DING_WEBHOOK_URL = st.secrets["ding_webhook"]
        if not DING_WEBHOOK_URL:
            return
        # 获取上海时区(北京时间)
        tz_shanghai = pytz.timezone("Asia/Shanghai")
        now_time = datetime.now(tz_shanghai).strftime("%Y-%m-%d %H:%M:%S")
        msg_content = f"streamlit：中奖拆分，有用户执行了拆分功能，触发时间：{now_time}"
        payload = {
            "msgtype": "text",
            "text": {
                "content": msg_content
            }
        }
        requests.post(DING_WEBHOOK_URL, json=payload, timeout=3)
    except Exception:
        # 没有配置secrets / 钉钉发送失败，全部静默，不影响工具使用
        pass

# 1. 页面设置：使用居中布局
st.set_page_config(
    page_title="Excel 中奖金额拆分工具",
    page_icon="⚡",
    layout="centered"
)

# 2. 核心拆分逻辑
def split_logic(df, split_amount, target_col="中奖金额"):
    if target_col not in df.columns:
        raise ValueError(f"表格中未找到名称为「{target_col}」的列，请检查表头。")
    new_rows = []
    for _, row in df.iterrows():
        try:
            val = float(row[target_col])
        except (ValueError, TypeError):
            new_rows.append(row.to_dict())
            continue
        if val > split_amount:
            remaining = val
            while remaining > 0:
                current_split = min(remaining, split_amount)
                new_row = row.to_dict()
                new_row[target_col] = round(current_split, 2)
                new_rows.append(new_row)
                remaining -= current_split
                remaining = round(remaining, 2)
        else:
            new_rows.append(row.to_dict())
    return pd.DataFrame(new_rows)

# 3. 头部标题
st.title("⚡ Excel 中奖金额拆分")
st.caption("自动识别「中奖金额」列并按指定阈值切分，其余列数据自动复制。")

# 4. 重要声明 (醒目的警告框)
st.warning(
    "⚠️ **重要免责声明**：本工具仅供辅助处理。**拆分完成后，请务必自行严格核对拆分结果与数据总额，确认无误后再进行业务使用！**"
)

# 5. 原生卡片容器包裹核心操作区
with st.container(border=True):
    # 第一步：参数与文件配置
    st.subheader("1. 配置与上传", divider="blue")

    split_amount = st.number_input(
        "拆分阈值金额 (元)",
        min_value=0.01,
        value=200.0,
        step=50.0,
        format="%.2f",
        help="金额大于此数值的行会被切分成多行"
    )
    uploaded_file = st.file_uploader(
        "上传 Excel 文件 (.xlsx / .xls)",
        type=["xlsx", "xls"]
    )
    # 第二步：触发处理
    if uploaded_file is not None:
        st.write("")
        if st.button("开始拆分处理", type="primary", use_container_width=True):
            # 开启后台线程发送钉钉通知，真正执行功能时触发
            t = threading.Thread(target=ding_notify_user_action, daemon=True)
            t.start()

            with st.spinner("正在拆分处理中..."):
                try:
                    if uploaded_file.name.lower().endswith('.xls'):
                        df = pd.read_excel(uploaded_file, sheet_name=0, engine='xlrd')
                    else:
                        df = pd.read_excel(uploaded_file, sheet_name=0, engine='openpyxl')
                    result_df = split_logic(df, split_amount, target_col="中奖金额")
                    output_stream = io.BytesIO()
                    with pd.ExcelWriter(output_stream, engine='openpyxl') as writer:
                        result_df.to_excel(writer, index=False)
                    output_data = output_stream.getvalue()
                    # 保存结果到 session
                    st.session_state['result_df'] = result_df
                    st.session_state['output_data'] = output_data
                    st.session_state['file_name'] = f"拆分后_{uploaded_file.name.rsplit('.', 1)[0]}.xlsx"
                except Exception as e:
                    st.error(f"处理失败: {str(e)}")

# 6. 结果展示卡片
if 'output_data' in st.session_state and uploaded_file is not None:
    st.write("")
    with st.container(border=True):
        st.subheader("2. 处理结果", divider="green")
        st.success("已成功将表格拆分完成！")

        # 结果预览
        st.dataframe(st.session_state['result_df'].head(10), use_container_width=True)
        # 醒目的全宽下载按钮
        st.download_button(
            label="📥 下载拆分后的 Excel",
            data=st.session_state['output_data'],
            file_name=st.session_state['file_name'],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
