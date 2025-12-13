import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import io
from datetime import datetime

# 设置页面标题
st.title("📄 公假单/抵晚单生成工具")
st.write("自动处理Excel数据，按学院排序生成格式规范的请假单")

# 上传Excel文件
excel_file = st.file_uploader("选择Excel文件", type=['xlsx', 'xls'])

def create_document(df, selected_columns, doc_type, activity_info):
    """创建Word文档"""
    doc = Document()
    
    # 设置全局字体和段落间距
    normal_style = doc.styles['Normal']
    normal_style.font.name = '宋体'
    normal_style._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    normal_style.font.size = Pt(10.5)
    normal_style.paragraph_format.space_before = Pt(0)
    normal_style.paragraph_format.space_after = Pt(0)
    normal_style.paragraph_format.line_spacing = 1.0
    
    # 标题
    title_paragraph = doc.add_paragraph()
    title = '公假单' if doc_type == "公假单" else '抵晚自习请假单'
    title_run = title_paragraph.add_run(title)
    title_run.font.name = '黑体'
    title_run._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
    title_run.font.size = Pt(22)
    title_run.font.bold = True
    title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_paragraph.paragraph_format.space_after = Pt(12)
    
    # 学院称呼
    college_paragraph = doc.add_paragraph()
    college_run = college_paragraph.add_run('各二级学院：')
    college_run.font.name = '宋体'
    college_run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    college_run.font.size = Pt(12)
    college_run.font.bold = True
    
    # 正文内容
    if doc_type == "公假单":
        text1 = f'兹定于{activity_info["activity_date"]}举办"{activity_info["activity_name"]}"活动。以下同学因参与活动组织工作，将于{activity_info["work_date"]} {activity_info["work_time"]}协助相关会务工作，无法参加该时间段课程。'
        text2 = f'特此申请为以下同学办理 {activity_info["work_date"]} {activity_info["work_time"]} 的公假手续，恳请贵学院予以批准，谢谢！'
    else:
        text1 = f'以下同学因参与{activity_info["work_date"]}的"{activity_info["activity_name"]}"活动，无法参加当晚晚自习。'
        text2 = '特此申请为以下同学办理晚自习请假手续，恳请贵学院予以批准，谢谢！'
    
    # 第一段
    para1 = doc.add_paragraph()
    para1.paragraph_format.first_line_indent = Pt(21)
    run1 = para1.add_run(text1)
    run1.font.name = '宋体'
    run1._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    run1.font.size = Pt(10.5)
    
    # 第二段
    para2 = doc.add_paragraph()
    para2.paragraph_format.first_line_indent = Pt(21)
    run2 = para2.add_run(text2)
    run2.font.name = '宋体'
    run2._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    run2.font.size = Pt(10.5)
    
    # 表格前间距
    doc.add_paragraph()
    
    # 创建表格
    table = doc.add_table(rows=1, cols=len(selected_columns))
    table.style = "Table Grid"
    
    # 表头
    header_cells = table.rows[0].cells
    for i, col in enumerate(selected_columns):
        para = header_cells[i].paragraphs[0]
        run = para.add_run(str(col))
        run.font.name = '宋体'
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
        run.font.size = Pt(11)
        run.font.bold = True
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # 数据行
    for _, row in df.iterrows():
        row_cells = table.add_row().cells
        for i, col in enumerate(selected_columns):
            value = row[col] if pd.notna(row[col]) else ""
            para = row_cells[i].paragraphs[0]
            run = para.add_run(str(value))
            run.font.name = '宋体'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
            run.font.size = Pt(10.5)
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # 表格后间距
    doc.add_paragraph()
    
    # 落款
    signature_paragraph = doc.add_paragraph()
    signature_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    run1 = signature_paragraph.add_run('共青团温州理工学院委员会\n')
    run1.font.name = '宋体'
    run1._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    run1.font.size = Pt(10.5)
    run1.font.bold = True
    
    run2 = signature_paragraph.add_run(activity_info['signature_date'])
    run2.font.name = '宋体'
    run2._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    run2.font.size = Pt(10.5)
    
    return doc

# 主程序
if excel_file is not None:
    try:
        # 读取Excel
        df = pd.read_excel(excel_file, header=0)
        df.columns = df.columns.str.strip()
        
        st.write("**数据预览：**")
        st.dataframe(df.head(5))
        
        # 选择文档类型
        doc_type = st.radio("选择文档类型", ["公假单", "抵晚单"])
        
        # 填写活动信息
        st.write("**第二步：填写活动信息**")
        col1, col2 = st.columns(2)
        
        with col1:
            activity_name = st.text_input("活动名称", "学术讲座")
            work_date = st.text_input("工作日期（如：X月X日）", "X月X日")
            signature_date = st.text_input("落款日期", "xx年xx月xx日")
        
        with col2:
            if doc_type == "公假单":
                activity_date = st.text_input("活动举办日期", "X年X月X日")
                work_time = st.selectbox("工作时间段", ["上午", "下午", "全天"])
            else:
                activity_date = ""
                work_time = ""
        
        # 选择列
        st.write("**第三步：选择表格列**")
        selected_columns = st.multiselect(
            "选择要显示的列",
            df.columns.tolist(),
            default=df.columns[:min(4, len(df.columns))].tolist()
        )
        
        if selected_columns:
            activity_info = {
                'activity_name': activity_name,
                'activity_date': activity_date,
                'work_date': work_date,
                'work_time': work_time,
                'signature_date': signature_date
            }
            
            if st.button("生成文档"):
                with st.spinner("正在生成..."):
                    doc = create_document(df, selected_columns, doc_type, activity_info)
                    file_stream = io.BytesIO()
                    doc.save(file_stream)
                    file_stream.seek(0)
                    
                    st.success("✅ 文档生成成功！")
                    st.download_button(
                        label="📥 下载Word文档",
                        data=file_stream,
                        file_name=f"{doc_type}_{activity_name}_{datetime.now().strftime('%Y%m%d')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
    
    except Exception as e:
        st.error(f"处理文件失败: {str(e)}")
else:
    st.info("请先上传Excel文件")
