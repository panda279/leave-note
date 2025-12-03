import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml.ns import qn
import io

# 设置页面标题
st.title("📊 Excel转Word工具 (学院精确排序版)")
st.write("自动清理空格后，按指定顺序严格排序学院数据")

# 第一步：上传Excel文件
st.header("第一步：上传Excel文件")
excel_file = st.file_uploader("选择Excel文件", type=['xlsx'])

# 定义学院排序顺序
COLLEGE_ORDER = [
    "经济与管理学院",
    "法学院",
    "文学与传媒学院", 
    "数据科学与人工智能学院",
    "电子与电气学院",
    "机器人工程学院",
    "建筑与能源工程学院",
    "设计艺术学院",
    "外国语学院",
    "创新创业学院"
]

def set_font(run, font_name='宋体', font_size=Pt(10.5), bold=False):
    """统一设置字体，确保中文字体生效"""
    run.font.name = font_name
    # 关键：设置中文字体
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.size = font_size
    run.font.bold = bold
    return run

def create_word_document(df, selected_columns):
    # 创建文档
    doc = Document()
    
    # ========== 第一部分：全局字体设置 ==========
    # 1. 设置文档默认字体（最基础保障）
    normal_style = doc.styles['Normal']
    normal_style.font.name = '宋体'
    normal_style._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    normal_style.font.size = Pt(10.5)
    
    # 2. 安全设置其他关键样式
    key_style_names = ['Normal', 'Default Paragraph Font', 'Body Text']
    for style_name in key_style_names:
        try:
            style = doc.styles[style_name]
            if hasattr(style, 'font'):
                style.font.name = '宋体'
                # 确保中文字体设置
                rpr = style.element.get_or_add_rPr()
                rfonts = rpr.get_or_add_rFonts()
                rfonts.set(qn('w:eastAsia'), '宋体')
                style.font.size = Pt(10.5)
        except (KeyError, AttributeError):
            continue

    # ========== 第二部分：强化字体设置函数 ==========
    def set_font_robust(run, font_name='宋体', font_size=Pt(10.5), bold=False):
        # 设置英文字体
        run.font.name = font_name
        
        # 关键：确保中文字体设置（双重保障）
        run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
        
        # 额外保障：直接操作XML
        try:
            rpr = run._element.get_or_add_rPr()
            rfonts = rpr.get_or_add_rFonts()
            rfonts.set(qn('w:eastAsia'), font_name)
        except:
            pass  # 如果上面的方法失败，使用默认方法
        
        run.font.size = font_size
        run.font.bold = bold
        return run
    
    # ========== 修正：文档大标题（无默认下划线） ==========
    title_paragraph = doc.add_paragraph()
    title_run = title_paragraph.add_run('公假单')
    # 设置标题字体：黑体、小二、加粗、居中
    title_run.font.name = '黑体'
    title_run._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
    title_run.font.size = Pt(22)
    title_run.font.bold = True
    title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    # 在大标题后添加一个空行，使排版更美观
    doc.add_paragraph()
    # ========== 大标题添加结束 ==========
    
    # --- 请假说明 ---
    title_paragraph = doc.add_paragraph()
    title_run = title_paragraph.add_run('各二级学院：')
    set_font_robust(title_run, '宋体', Pt(12), bold=True)

    # 添加第一段文字，并设置缩进：使其首字与上一行的“各”字对齐
    text_paragraph1 = doc.add_paragraph()
    # 关键：设置整个段落的左侧缩进，抵消默认的段落缩进，使首字顶到最左
    text_paragraph1.paragraph_format.left_indent = Pt(0)  # 确保左侧无额外缩进
    text_paragraph1.paragraph_format.first_line_indent = Pt(24) # 首行不额外缩进
    text_paragraph1.paragraph_format.space_after = Pt(0)  # 与下段无间距
    text_content1 = '兹定于X年X月X日举办"XXX（填活动名称）"活动。以下同学因参与活动组织工作，将于X月X日 上午/下午/全天（根据实际时间选择）协助相关会务工作，无法参加该时间段课程。'
    text_run1 = text_paragraph1.add_run(text_content1)
    set_font_robust(text_run1, '宋体', Pt(10.5))

    # 添加第二段文字，缩进设置与第一段完全相同
    text_paragraph2 = doc.add_paragraph()
    text_paragraph2.paragraph_format.left_indent = Pt(0)  # 左侧无额外缩进
    text_paragraph2.paragraph_format.first_line_indent = Pt(24) # 首行不额外缩进
    # 第二段后可以留一点间距，或设为0与表格紧贴
    text_paragraph2.paragraph_format.space_after = Pt(12)
    text_content2 = '特此申请为以下同学办理 X月X日 上午/下午/全天 的公假手续，恳请贵学院予以批准，谢谢！'
    text_run2 = text_paragraph2.add_run(text_content2)
    set_font_robust(text_run2, '宋体', Pt(10.5))

    # 在说明文字和表格之间添加一个空行
    doc.add_paragraph()
    
    table=doc.add_table(rows=1,cols=len(selected_columns))
    # 设置宽度
    for i, col in enumerate(selected_columns):
        base_width = Inches(2.0)
        extra_per_char = Inches(0.08)
        col_width = base_width + (len(str(col))) * extra_per_char
        table.columns[i].width = min(col_width, Inches(3.5))

    # 添加边框函数（内部函数）
    def add_table_borders(table_obj):
        """手动为表格添加边框，不影响字体"""
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        
        tbl = table_obj._tbl
        # 为表格添加边框属性
        tblPr = tbl.get_or_add_tblPr()
        
        # 创建边框元素
        borders = OxmlElement('w:tblBorders')
        
        # 定义各边边框
        border_types = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
        for border_type in border_types:
            border = OxmlElement(f'w:{border_type}')
            border.set(qn('w:val'), 'single')      # 单线边框
            border.set(qn('w:sz'), '4')            # 边框粗细（4=0.5磅）
            border.set(qn('w:space'), '0')         # 边框间距
            border.set(qn('w:color'), '000000')    # 黑色边框
            borders.append(border)
        
        tblPr.append(borders)
    
    # 调用函数添加边框
    table.style="Table Grid"                      
    
    # 表头
    header_cells = table.rows[0].cells
    for i, col in enumerate(selected_columns):
        header_cells[i].text = ''
        paragraph = header_cells[i].paragraphs[0]
        paragraph.clear()  # 清空所有内容
        
        run = paragraph.add_run(str(col))
        set_font_robust(run, '宋体', Pt(11), bold=True)
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # 数据行
    for index, row in df.iterrows():
        row_cells = table.add_row().cells
        for i, col in enumerate(selected_columns):
            value = row[col]
            row_cells[i].text = ''
            paragraph = row_cells[i].paragraphs[0]
            paragraph.clear()  # 清空所有内容
            
            text_content = str(value) if pd.notna(value) else ""
            run = paragraph.add_run(text_content)
            set_font_robust(run, '宋体', Pt(10.5))
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # ========== 第五部分：落款 ==========
    doc.add_paragraph()
    signature_paragraph = doc.add_paragraph()
    signature_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    run1 = signature_paragraph.add_run('共青团温州理工学院委员会')
    set_font_robust(run1, '宋体', Pt(10.5), bold=True)
    signature_paragraph.add_run('\n')
    
    run2 = signature_paragraph.add_run('xx年xx月xx日')
    set_font_robust(run2, '宋体', Pt(10.5))
    
    return doc

if excel_file is not None:
    # 读取Excel文件
    df = pd.read_excel(excel_file)
    
    # 显示原始数据预览
    st.subheader("数据预览 (原始)")
    st.write(f"总共有 {len(df)} 行数据")
    st.dataframe(df)
    
    # 第二步：检查并处理"学院"列
    st.header("第二步：处理学院排序")
    
    # 检查是否存在"学院"列
    if '学院' not in df.columns:
        st.error("错误：在Excel文件中未找到名为'学院'的列。请检查列名。")
        st.stop()
    
    # 核心步骤1：自动删除空格
    st.info("正在清理'学院'列中的空格...")
    df['学院'] = df['学院'].astype(str).str.strip()
    st.info("正在规范化学院名称")
    college_name_mapping={
        "经管学院":"经济与管理学院",
        "文传学院":"文学与传媒学院",
        "电电学院":"电子与电气工程学院",
        "建工学院":"建筑与能源工程学院",
        "外院":"外国语学院",
        "设艺学院":"设计艺术学院",
        "创业学院":"创新与创业学院",
        "数智学院":"数据科学与人工智能学院"}
    def normalize_college_name(name):
        name_clean=str(name).strip()
        return college_name_mapping.get(name_clean,name_clean)
    df["学院"]=df["学院"].apply(normalize_college_name)      
    # 显示清理后的唯一值
    unique_colleges = df['学院'].unique()
    st.write("**清理空格后，'学院'列的唯一值有：**", unique_colleges.tolist())
    
    # 核心步骤2：按指定顺序重组数据
    st.info("正在按指定顺序重组数据...")
    
    # 创建一个空的DataFrame来存放排序后的结果
    sorted_dfs = []
    
    # 按照指定顺序，逐个学院提取数据
    for college in COLLEGE_ORDER:
        # 筛选出当前学院的行
        college_data = df[df['学院'] == college]
        if not college_data.empty:
            sorted_dfs.append(college_data)
            st.write(f"  ✓ 已提取: {college} ({len(college_data)}行)")
        else:
            st.write(f"  ⚠ 未找到: {college} (0行)")
    
    # 合并所有排序后的数据
    if sorted_dfs:
        df_sorted = pd.concat(sorted_dfs, ignore_index=True)
        
        # 处理不在指定顺序中的其他学院
        other_colleges = set(df['学院'].unique()) - set(COLLEGE_ORDER)
        if other_colleges:
            st.warning(f"发现以下未在排序列表中的学院，它们将被放在最后：{list(other_colleges)}")
            other_data = df[df['学院'].isin(other_colleges)]
            df_sorted = pd.concat([df_sorted, other_data], ignore_index=True)
        
        # 显示排序后的数据
        st.subheader("数据预览 (按学院排序后)")
        st.dataframe(df_sorted)
        
        # 更新df为排序后的数据
        df = df_sorted
    else:
        st.error("未匹配到任何指定学院的数据。请检查'学院'列的值。")
        st.stop()
    
    # 第三步：选择列
    st.header("第三步：选择要导出的列")
    all_columns = df.columns.tolist()
    selected_columns = st.multiselect(
        "选择要添加到Word的列",
        all_columns,
        default=all_columns[:4] if len(all_columns) >= 4 else all_columns)
    
    # 第四步：生成Word文档
    st.header("第四步：生成Word文档")
    
    if st.button("生成Word文档") and selected_columns:
        with st.spinner("正在生成Word文档..."):
            # 创建Word文档
            doc = create_word_document(df, selected_columns)
            
            # 保存到内存
            file_stream = io.BytesIO()
            doc.save(file_stream)
            file_stream.seek(0)
            
            # 提供下载
            st.success("文档生成成功！")
            st.download_button(
                label="📥 下载Word文档",
                data=file_stream,
                file_name="按学院排序的表格.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

else:
    st.info("请先上传Excel文件")