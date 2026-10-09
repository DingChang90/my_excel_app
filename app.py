import io
import pandas as pd
import streamlit as st

# 1. 页面基本设置
st.set_page_config(page_title="Excel 中奖金额拆分工具", page_icon="📊")
st.title("📊 Excel 中奖金额拆分工具")

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

        # 切分逻辑，不留 0 元
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

# 3. 前端界面交互
split_amount = st.number_input(
    "1. 请输入拆分金额 (元):", 
    min_value=0.01, 
    value=200.0, 
    step=10.0,
    format="%.2f"
)

uploaded_file = st.file_uploader(
    "2. 请上传 Excel 文件 (.xlsx / .xls):", 
    type=["xlsx", "xls"]
)

if uploaded_file is not None and split_amount > 0:
    if st.button("🚀 开始拆分并生成文件", type="primary"):
        with st.spinner("处理中，请稍候..."):
            try:
                # 读取上传的文件（仅读取第一个 sheet）
                if uploaded_file.name.lower().endswith('.xls'):
                    df = pd.read_excel(uploaded_file, sheet_name=0, engine='xlrd')
                else:
                    df = pd.read_excel(uploaded_file, sheet_name=0, engine='openpyxl')

                # 执行拆分
                result_df = split_logic(df, split_amount, target_col="中奖金额")

                # 生成 Excel 内存流
                output_stream = io.BytesIO()
                with pd.ExcelWriter(output_stream, engine='openpyxl') as writer:
                    result_df.to_excel(writer, index=False)
                output_data = output_stream.getvalue()

                # 提示成功并展示预览
                st.success("拆分完成！")
                st.write("▼ 处理后数据预览（前 10 行）：")
                st.dataframe(result_df.head(10))

                # 下载按钮
                base_name = uploaded_file.name.rsplit('.', 1)[0]
                download_filename = f"拆分后_{base_name}.xlsx"

                st.download_button(
                    label="📥 点击下载拆分后的 Excel",
                    data=output_data,
                    file_name=download_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"处理失败，原因: {str(e)}")
