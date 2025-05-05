from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

# 打开现有的 Word 文档
doc = Document("output.docx")
# 遍历文档中的所有表格
for table in doc.tables:
    # 设置表格属性
    tblPr = table._element.xpath(".//w:tblPr")[0]  # 获取表格属性节点

    # 设置表格宽度为自动
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:w"), "0")
    tblW.set(qn("w:type"), "auto")
    tblPr.append(tblW)

    # 设置表格边框
    tblBorders = OxmlElement("w:tblBorders")
    for border in ["top", "bottom"]:
        borderElement = OxmlElement(f"w:{border}")
        borderElement.set(qn("w:val"), "single")  # 单线
        borderElement.set(qn("w:sz"), "12")  # 边框大小
        borderElement.set(qn("w:space"), "0")  # 间距
        borderElement.set(qn("w:color"), "000000")  # 黑色
        tblBorders.append(borderElement)
    tblPr.append(tblBorders)

    # 设置表格样式
    tblLook = OxmlElement("w:tblLook")
    tblLook.set(qn("w:val"), "04A0")  # 样式值
    tblLook.set(qn("w:firstRow"), "1")  # 首行样式
    tblLook.set(qn("w:lastRow"), "0")  # 末行无样式
    tblLook.set(qn("w:firstColumn"), "1")  # 首列样式
    tblLook.set(qn("w:lastColumn"), "0")  # 末列无样式
    tblLook.set(qn("w:noHBand"), "0")  # 启用水平带状样式
    tblLook.set(qn("w:noVBand"), "1")  # 禁用垂直带状样式
    tblPr.append(tblLook)

    # 设置表格居中
    tblAlignment = OxmlElement("w:jc")
    tblAlignment.set(qn("w:val"), "center")  # 表格居中
    tblPr.append(tblAlignment)

    # 计算每列宽度
    total_width_twips = 9286  # 表格总宽度
    num_columns = len(table.columns)  # 表格列数
    column_width_twips = total_width_twips // num_columns  # 每列宽度

    # 为表格的第一行设置样式
    first_row = table.rows[0]  # 获取第一行
    for cell in first_row.cells:
        tcPr = cell._element.get_or_add_tcPr()  # 获取或创建单元格属性节点

        # 设置单元格底部边框
        tcBorders = OxmlElement("w:tcBorders")
        bottomBorder = OxmlElement("w:bottom")
        bottomBorder.set(qn("w:val"), "single")  # 单线
        bottomBorder.set(qn("w:sz"), "6")  # 边框大小
        bottomBorder.set(qn("w:space"), "0")  # 间距
        bottomBorder.set(qn("w:color"), "000000")  # 黑色
        tcBorders.append(bottomBorder)
        tcPr.append(tcBorders)

        # 设置单元格填充样式
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")  # 填充类型
        shd.set(qn("w:color"), "auto")  # 自动颜色
        shd.set(qn("w:fill"), "auto")  # 自动填充
        tcPr.append(shd)

    # 设置每列宽度并为单元格添加段落和文本样式
    for column in table.columns:
        for cell in column.cells:
            cell.width = Pt(column_width_twips / 20)  # 转换为点（1 点 = 20 twips）

            # 设置单元格段落属性
            pPr = OxmlElement("w:pPr")
            snapToGrid = OxmlElement("w:snapToGrid")
            snapToGrid.set(qn("w:val"), "0")  # 不对齐到网格
            pPr.append(snapToGrid)

            ind = OxmlElement("w:ind")
            ind.set(qn("w:firstLine"), "480")  # 首行缩进 480 twips
            pPr.append(ind)

            jc = OxmlElement("w:jc")
            jc.set(qn("w:val"), "center")  # 居中对齐
            pPr.append(jc)

            rPr = OxmlElement("w:rPr")
            rFonts = OxmlElement("w:rFonts")
            sz = OxmlElement("w:sz")
            sz.set(qn("w:val"), "21")
            rPr.append(sz)

            # 将段落属性和文本样式添加到单元格段落
            for paragraph in cell.paragraphs:
                paragraph._p.insert(0, pPr)  # 插入段落属性
                for run in paragraph.runs:
                    run._r.insert(0, rPr)  # 确保样式应用到每个运行

# 遍历每个节并插入页眉
for section in doc.sections:
    # 获取节的页眉
    header = section.header
    # 确保页眉中有段落
    if not header.paragraphs:
        paragraph = header.add_paragraph()
    else:
        paragraph = header.paragraphs[0]

    # 清空段落内容
    for run in paragraph.runs:
        run.clear()

    # 添加页眉内容
    run = paragraph.add_run("这是页眉内容 ")

    # 创建页码字段
    fldChar1 = OxmlElement("w:fldChar")
    fldChar1.set(qn("w:fldCharType"), "begin")
    run._r.append(fldChar1)

    instrText = OxmlElement("w:instrText")
    instrText.set(qn("xml:space"), "preserve")  # 避免字段指令被截断
    instrText.text = "PAGE \\* MERGEFORMAT"
    run._r.append(instrText)

    fldChar2 = OxmlElement("w:fldChar")
    fldChar2.set(qn("w:fldCharType"), "end")
    run._r.append(fldChar2)

    # 设置段落样式
    paragraph.style = "Header"  # 确保样式名称正确
    section.header_distance = Pt(56.7)  # 页眉距离顶部2cm


def is_abstract_paragraph(paragraph):
    return paragraph.style.name.lower() == 'abstract'

def apply_simsun_font(run):
    run.font.name = 'SimSun'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'SimSun')

for para in doc.paragraphs:
    if is_abstract_paragraph(para):
        text = para.text.strip()
        if text.startswith("关键词："):
            prefix = "关键词："
        elif text.lower().startswith("keywords:"):
            prefix = "Keywords:"
        else:
            raise AssertionError(f"Abstract paragraph does not start with expected prefix: {text}")

        # 清空原段落的所有 run，并重新构建
        para.clear()

        # 添加前缀部分（不改字体）
        run_prefix = para.add_run(prefix)

        # 添加其余内容，应用宋体字体
        remaining = text[len(prefix):]
        run_rest = para.add_run(remaining)
        apply_simsun_font(run_rest)



# 保存修改后的文档
doc.save("output.docx")
