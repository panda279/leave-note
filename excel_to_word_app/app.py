import streamlit as st
import pandas as pd
from docx import Document
import io

# 设置页面标题
st.title("📊 Excel转Word工具")
st.write("最简单的Excel数据转Word表格工具")

# 第一步：上传Excel文件
st.header("第一步：上传Excel文件")
excel_file = st.file_uploader("选择Excel文件", type=['xlsx'])

if excel_file is not None:
    # 读取Excel文件
    df = pd.read_excel(excel_file)
    
    # 显示数据预览
    st.subheader("数据预览")
    st.write(f"总共有 {len(df)} 行数据")
    st.dataframe(df)
    
    # 第二步：选择列
    st.header("第二步：选择要导出的列")
    all_columns = df.columns.tolist()
    selected_columns = st.multiselect(
        "选择要添加到Word的列",
        all_columns,
        default=all_columns[:4] if len(all_columns) >= 4 else all_columns
    )
    
    # 第三步：生成Word文档
    st.header("第三步：生成Word文档")
    
    if st.button("生成Word文档") and selected_columns:
        # 创建Word文档
        doc = Document()
        
        # 添加标题
        doc.add_heading('Excel数据表格', 0)
        
        # 创建表格
        table = doc.add_table(rows=1, cols=len(selected_columns))
        
        # 设置表头
        header_cells = table.rows[0].cells
        for i, col in enumerate(selected_columns):
            header_cells[i].text = str(col)
        
        # 添加数据行
        for index, row in df.iterrows():
            row_cells = table.add_row().cells
            for i, col in enumerate(selected_columns):
                value = row[col]
                row_cells[i].text = str(value) if pd.notna(value) else ""
        
        # 保存到内存
        file_stream = io.BytesIO()
        doc.save(file_stream)
        file_stream.seek(0)
        
        # 提供下载
        st.success("文档生成成功！")
        st.download_button(
            label="📥 下载Word文档",
            data=file_stream,
            file_name="生成的表格.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

else:
    st.info("请先上传Excel文件")

