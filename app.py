import io
import pandas as pd
import streamlit as st

# 1. 页面基本配置
st.set_page_config(
    page_title="Excel 中奖金额拆分工具", 
    page_icon="✂️",
    layout="centered"  # 居中紧凑布局
)

# 2. 注入自定义 CSS 样式（提升 UI 美观度）
st.markdown("""
    <style>
    /* 限制整体容器宽度并居中，增加卡片阴影感 */
    .block-container {
        max-width: 720px;
        padding-top: 2rem;
        padding-bottom: 3rem;
    }
    
    /* 标题与副标题样式 */
    .main-title {
        font-size: 2rem;
        font-weight: 700;
        color: #1E293B;
        margin-bottom: 0.2rem;
    }
    .sub-title {
        font-size: 0.95rem;
        color: #64748B;
        margin-bottom: 1.5rem;
    }
    
    /* 说明卡片 */
    .info-card {
        background-color: #F8FAFC;
        border-left: 4px solid #3B82F6;
        padding: 1rem 1.2rem;
        border-radius: 6px;
        margin-bottom: 1.5rem;
        font-size: 0.9rem;
        color: #334155;
    }
    
    /* 自定义按钮全宽与样式 */
    div.stButton > button {
        width: 100%;
        border-radius: 8px;
        height: 2.8rem;
        font-size: 1rem;
        font-weight: 600;
    }
    
    /* 下载按钮高亮 */
    div.stDownloadButton > button {
        width: 100%;
        background-color: #10B981 !important;
        color: white !important;
        border: none !important;
        border-radius: 8px;
        height: 3rem;
        font-size: 1.05rem;
        font-weight: 700;
        box-shadow: 0 4px 6px -1px rgba(16, 185, 129, 0.2);
    }
    div.stDownloadButton > button:hover {
        background-color: #059669 !important;
    }
    </style>
""", unsafe_allow_html=True)

# 3. 头部区域
st.markdown('<div class="main-title">✂️ Excel 中奖金额拆分工具</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">自动识别「中奖金额」列并按设定阈值切分行，其余列数据自动复用</div>', unsafe_allow_html=True)

# 4. 使用说明卡片
st.markdown("""
<div class="info-card">
    <b>💡 使用说明：</b><br>
    1. 上传包含 <b>「中奖金额」</b> 列的 Excel 表格（只处理第一个 Sheet）<br>
    2. 设置单笔拆分上限（如 200），系统自动将 >200 的金额拆分为多行（如 414 拆为 200、200、14）<br>
    3. 处理完成后可实时预览并一键下载结果文件。
</div>
""", unsafe_allow_html=True)

# 5. 核心拆分逻辑
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

# 6. 表单交互区（分两列设置参数，更加精致）
col1, col2 = st.columns([1, 1])

with col1:
    split_amount = st.number_input(
        "⚙️ 设置拆分阈值金额 (元)", 
        min_value=0.01, 
        value=200.0, 
        step=50.0,
        format="%.2f",
        help="超过此金额的数据行将被拆分"
    )

with col2:
    uploaded_file = st.file_uploader(
        "📁 上传 Excel 文件 (.xlsx / .xls)", 
        type=["xlsx", "xls"]
    )

# 7. 处理与下载逻辑
if uploaded_file is not None and split_amount > 0:
    st.divider()
    
    if st.button("🚀 开始拆分处理", type="primary"):
        with st.spinner("数据拆分处理中..."):
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

                # 保存到 Session 中以保持下载状态
                st.session_state['result_df'] = result_df
                st.session_state['output_data'] = output_data
                st.session_state['file_name'] = f"拆分后_{uploaded_file.name.rsplit('.', 1)[0]}.xlsx"

            except Exception as e:
                st.error(f"❌ 处理失败: {str(e)}")

# 8. 结果展示与下载（如果处理成功）
if 'output_data' in st.session_state and uploaded_file is not None:
    st.success("🎉 拆分成功！")
    
    # 结果预览卡片
    with st.expander("👀 点击查看拆分后数据预览（前 15 行）", expanded=True):
        st.dataframe(st.session_state['result_df'].head(15), use_container_width=True)

    # 绿色高亮一键下载按钮
    st.download_button(
        label="📥 立即下载拆分后的 Excel 文件",
        data=st.session_state['output_data'],
        file_name=st.session_state['file_name'],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
