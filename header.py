import os
import re
import shutil

from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

TITLE = "面向优秀论文标准的研究"
statistics_data = {}


# --- Helper Functions ---
def create_element(name):
    """Create an OXML element."""
    return OxmlElement(name)


def create_attribute(element, name, value):
    """Set an attribute for an OXML element."""
    element.set(qn(name), value)


def delete_paragraph(paragraph):
    """Remove a paragraph element."""
    p = paragraph._element
    if p is not None and p.getparent() is not None:
        p.getparent().remove(p)


def add_page_number_field(paragraph):
    """Add a PAGE field to a paragraph."""
    run = paragraph.add_run()
    fldSimple = create_element("w:fldSimple")
    create_attribute(fldSimple, "w:instr", r" PAGE \* MERGEFORMAT ")
    run._r.append(fldSimple)


def set_page_number_style(section, fmt="decimal", start=None):
    """Set page number format and starting number for a section."""
    sectPr = section._sectPr
    pgNumType = sectPr.find(qn("w:pgNumType")) or create_element("w:pgNumType")
    create_attribute(pgNumType, "w:fmt", fmt)
    if start is not None:
        create_attribute(pgNumType, "w:start", str(start))
    if pgNumType not in sectPr:
        sectPr.append(pgNumType)


def apply_simsun_tnr_font(run):
    """Apply SimSun font to a run."""
    run.font.name = "SimSun"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "SimSun")
    run._element.rPr.rFonts.set(qn("w:ascii"), "Times New Roman")
    run._element.rPr.rFonts.set(qn("w:hAnsi"), "Times New Roman")


def copy_sectPr_properties(source_sectPr, target_sectPr):
    """Copy key page settings from source sectPr to target_sectPr."""
    print("  Starting property copying...")
    properties_to_copy = ["pgSz", "pgMar", "cols", "docGrid"]

    for prop_tag in properties_to_copy:
        source_element = source_sectPr.find(qn(f"w:{prop_tag}"))
        if source_element is not None:
            # Create a new element with the same tag
            target_element = create_element(f"w:{prop_tag}")

            # Copy all attributes
            for name, value in source_element.attrib.items():
                ns, localname = name.split("}") if "}" in name else ("", name)
                localname = localname.split(":")[-1] if ":" in localname else localname
                create_attribute(target_element, f"w:{localname}", value)

            # Replace existing or append
            existing = target_sectPr.find(qn(f"w:{prop_tag}"))
            if existing is not None:
                target_sectPr.replace(existing, target_element)
            else:
                target_sectPr.append(target_element)

            print(f"  ✓ Copied {prop_tag} settings")

    print("  Property copying complete")


def add_next_page_section_break(paragraph):
    """Add a next-page section break to a paragraph."""
    assert paragraph is not None, "Paragraph cannot be None"

    p_element = paragraph._p
    pPr = p_element.get_or_add_pPr()

    assert pPr.find(qn("w:sectPr")) is None, "Paragraph already has a sectPr"
    sectPr = create_element("w:sectPr")
    pPr.append(sectPr)

    # Set section break type to next page
    type_element = sectPr.find(qn("w:type"))
    assert type_element is None, "Paragraph already has a type element"
    type_element = create_element("w:type")
    sectPr.append(type_element)
    create_attribute(type_element, "w:val", "nextPage")

    # Clear paragraph content (as it's just a marker)
    paragraph.text = ""
    assert len(paragraph.runs) == 1, "Paragraph must have exactly one run"
    p_element.remove(paragraph.runs[0]._r)

    print("Section break added successfully with appropriate properties")
    return True


def is_abstract_paragraph(paragraph):
    """Check if a paragraph is an abstract paragraph."""
    return paragraph.style.name.lower() == "abstract"


def add_toc(document):
    """
    Insert a Table of Contents (TOC) at the location of a marker text.
    """
    # find all the paragraph with style toc Heading
    target_paragraphs = [
        p for p in document.paragraphs if p.style.name.lower() == "toc heading"
    ]
    assert len(target_paragraphs) == 1, (
        "The document must contain exactly one paragraph with the style 'TOC Heading'."
    )
    target_paragraph = target_paragraphs[0]
    try:
        toc_title_paragraph = target_paragraph.insert_paragraph_before(
            "目录", style="TOC Heading"
        )
        toc_field_paragraph = target_paragraph.insert_paragraph_before("")
        run = toc_field_paragraph.add_run()

        fldChar_begin = create_element("w:fldChar")
        create_attribute(fldChar_begin, "w:fldCharType", "begin")

        instrText = create_element("w:instrText")
        create_attribute(instrText, "xml:space", "preserve")
        instrText.text = ' TOC \\o "1-3" \\h \\z \\u '

        fldChar_end = create_element("w:fldChar")
        create_attribute(fldChar_end, "w:fldCharType", "end")

        run._r.extend([fldChar_begin, instrText, fldChar_end])
        delete_paragraph(target_paragraph)
    except Exception as e:
        print(f"Error during TOC insertion: {e}")


# --- Main Processing ---
def process_document(doc):
    """Process the Word document."""

    for table in doc.tables:
        process_table(table)

    add_toc(doc)

    insert_section_breaks(doc)

    set_page_number_for_all_secs(doc)

    set_headers(doc)

    set_abstract_font(doc)

    process_math_equations(doc)

    style_superscript_hyperlinks(doc)

    replace_figure_format_in_doc(doc)

    force_update_fields(doc)


def insert_section_breaks(doc):
    """Insert section breaks at specified markers."""

    # find paragraphs with style "myBreak"
    break_paragraphs = [p for p in doc.paragraphs if p.style.name.lower() == "mybreak"]
    assert len(break_paragraphs) == 2, "目前只使用了2个分页标记，如果需要更多请修改代码"
    for p in break_paragraphs:
        assert add_next_page_section_break(p)


def process_table(table):
    """Apply formatting to a table."""
    tblPr = table._element.xpath(".//w:tblPr")[0]

    # Set table width to auto
    tblW = create_element("w:tblW")
    create_attribute(tblW, "w:w", "0")
    create_attribute(tblW, "w:type", "auto")
    tblPr.append(tblW)

    # Set table borders
    tblBorders = create_element("w:tblBorders")
    for border in ["top", "bottom"]:
        borderElement = create_element(f"w:{border}")
        create_attribute(borderElement, "w:val", "single")
        create_attribute(borderElement, "w:sz", "12")
        create_attribute(borderElement, "w:space", "0")
        create_attribute(borderElement, "w:color", "000000")
        tblBorders.append(borderElement)
    tblPr.append(tblBorders)

    # Set table style
    tblLook = create_element("w:tblLook")
    create_attribute(tblLook, "w:val", "04A0")
    create_attribute(tblLook, "w:firstRow", "1")
    create_attribute(tblLook, "w:lastRow", "0")
    create_attribute(tblLook, "w:firstColumn", "1")
    create_attribute(tblLook, "w:lastColumn", "0")
    create_attribute(tblLook, "w:noHBand", "0")
    create_attribute(tblLook, "w:noVBand", "1")
    tblPr.append(tblLook)

    # Center align the table
    tblAlignment = create_element("w:jc")
    create_attribute(tblAlignment, "w:val", "center")
    tblPr.append(tblAlignment)

    # Calculate column width
    TOTAL_WIDTH_TWIPS = 9286  # Total table width
    num_columns = len(table.columns)  # Number of columns
    column_width_twips = TOTAL_WIDTH_TWIPS // num_columns  # Column width

    # Apply style to the first row
    first_row = table.rows[0]  # Get the first row
    for cell in first_row.cells:
        tcPr = cell._element.get_or_add_tcPr()  # Get or create cell properties

        # Set bottom border
        tcBorders = create_element("w:tcBorders")
        bottomBorder = create_element("w:bottom")
        create_attribute(bottomBorder, "w:val", "single")
        create_attribute(bottomBorder, "w:sz", "6")
        create_attribute(bottomBorder, "w:space", "0")
        create_attribute(bottomBorder, "w:color", "000000")
        tcBorders.append(bottomBorder)
        tcPr.append(tcBorders)

    # Set column width and apply paragraph and text style
    for column in table.columns:
        for cell in column.cells:
            cell.width = Pt(column_width_twips / 20)  # Convert to points


def set_page_number_for_all_secs(doc):
    """Apply formatting to document sections."""
    num_sections = len(doc.sections)
    print(f"Document contains {num_sections} sections.")

    assert num_sections == 3, "Document must have at least 2 sections."

    # Section 1: Before TOC
    section1 = doc.sections[0]
    section1.footer_distance = Pt(56.7)  # Footer distance from the bottom (2 cm)
    set_page_number_style(section1, fmt="upperRoman", start=1)
    add_page_number_to_footer(section1)

    # Section 2: TOC
    section2 = doc.sections[1]
    section2.footer_distance = Pt(56.7)  # Footer distance from the bottom (2 cm)
    section2.footer.is_linked_to_previous = False
    set_page_number_style(section2, fmt="upperRoman", start=1)
    add_page_number_to_footer(section2)

    # Section 3: After TOC
    section3 = doc.sections[2]
    section3.footer_distance = Pt(56.7)  # Footer distance from the bottom (2 cm)
    section3.footer.is_linked_to_previous = False
    set_page_number_style(section3, fmt="decimal", start=1)


def set_headers(doc):
    """
    Set headers for all sections in the document.
    - First two sections: Header without page numbers.
    - Third section: Header with page numbers.
    """
    for i, section in enumerate(doc.sections):
        section.header.is_linked_to_previous = False
        # Get the header of the section
        header = section.header

        assert len(header.paragraphs) == 1, "Header must contain one paragraph"
        paragraph = header.paragraphs[0]

        # Clear existing content in the paragraph
        for run in paragraph.runs:
            run.clear()

        # Add header content
        if i < 2:  # First two sections
            paragraph.add_run(TITLE + " ")
        else:  # Third section and beyond
            run = paragraph.add_run(TITLE + " ")

            # Add page number field
            fldChar1 = OxmlElement("w:fldChar")
            fldChar1.set(qn("w:fldCharType"), "begin")
            run._r.append(fldChar1)

            instrText = OxmlElement("w:instrText")
            instrText.set(
                qn("xml:space"), "preserve"
            )  # Avoid truncation of field instructions
            instrText.text = "PAGE \\* MERGEFORMAT"
            run._r.append(instrText)

            fldChar2 = OxmlElement("w:fldChar")
            fldChar2.set(qn("w:fldCharType"), "end")
            run._r.append(fldChar2)

        # Set paragraph style
        paragraph.style = "header"  # Ensure the style name is correct

        # Set header distance
        section.header_distance = Pt(56.7)  # Header distance from the top (2 cm)


def add_page_number_to_footer(section):
    """Add page number to the footer of a section."""
    footer = section.footer
    paragraph = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    paragraph.style = "footer"
    add_page_number_field(paragraph)


def set_abstract_font(doc):
    """Set the font for all abstract paragraphs."""
    print("\n=== Processing abstract paragraphs ===")
    for paragraph in doc.paragraphs:
        if is_abstract_paragraph(paragraph):
            text = paragraph.text.strip()
            if text.startswith("关键词："):
                prefix = "关键词："
            elif text.lower().startswith("keywords:"):
                prefix = "Keywords:"
            else:
                print(
                    f"Warning: Abstract paragraph with unexpected format: '{text[:20]}...'"
                )
                continue

            print(f"Processing abstract paragraph: '{text[:30]}...'")
            paragraph.clear()

            # Add prefix with bold formatting
            run_prefix = paragraph.add_run(prefix)
            if prefix == "Keywords:":
                run_prefix.bold = True  # Apply bold to the prefix

            # Add the remaining text with appropriate font
            remaining = text[len(prefix) :]
            run_rest = paragraph.add_run(remaining)
            apply_simsun_tnr_font(run_rest)

    print("=== Completed processing abstract paragraphs ===\n")


def replace_figure_format_in_doc(doc):
    """
    Replace all occurrences of '图%d.%d' with '图%d-%d' in the entire Word document
    while preserving oMath elements and other formatting.
    """
    figure_pattern = re.compile(r"图(\d+)\.(\d+)")

    for paragraph in doc.paragraphs:
        # Process the paragraph run by run to preserve formatting
        for run in paragraph.runs:
            if run.text and "图" in run.text:
                run.text = figure_pattern.sub(r"图\1-\2", run.text)


def process_math_equations(doc):
    """
    Locate and format all oMathParagraph elements in the document.
    - Add right-aligned tab stop
    - Structure with equation, tab, and number
    - Detect and extract equation numbers in format (d.d)
    """
    print("\n=== STARTING MATH PARAGRAPH DETECTION AND FORMATTING ===")
    math_para_count = 0

    for i, paragraph in enumerate(doc.paragraphs):
        # Check if paragraph contains math paragraph elements
        if "<m:oMathPara" in paragraph._p.xml:
            math_para_count += 1
            # print(f"\n--- Math Paragraph #{math_para_count} found in paragraph {i} ---")
            all_math_t = paragraph._element.xpath(".//m:t")

            assert len(all_math_t) >= 3, "数学段落必须包含至少3个数学文本节点"
            last_4_text = all_math_t[-4].text
            assert last_4_text.strip("\u2001\u2003\u2000\u2002 ") == "", (
                "数学段落的倒数第四个数学文本节点必须是空格"
            )
            last_3_texts = [t.text for t in all_math_t[-3:]]
            assert last_3_texts[0] == "(" and last_3_texts[-1] == ")", (
                "数学段落的最后三个数学文本节点必须中有'('和')'"
            )
            equation_number = last_3_texts[1]
            assert re.match(r"^\d+\.\d+$", equation_number), (
                "数学段落的倒数第三个数学文本节点必须是数字编号，格式为d.d"
            )
            equation_number = equation_number.replace(".", "-")

            # Remove the last 4 m:t nodes
            for t_node in all_math_t[-4:]:
                r_node = t_node.getparent()

                if r_node is not None:
                    parent = r_node.getparent()
                    if parent is not None:
                        parent.remove(r_node)
            # Format the paragraph with oMathPara, passing the discovered number
            format_math_paragraph(paragraph, equation_number)

    print(
        f"\n=== COMPLETED MATH ELEMENT FORMATTING: {math_para_count} oMathPara elements formatted ===\n"
    )


def format_math_paragraph(paragraph, equation_number="replace_me"):
    """
    Format a paragraph containing oMathPara with proper style and equation numbering.

    Args:
        paragraph: The paragraph containing the math element
        equation_number: The sequential number to assign to this equation
    """
    try:
        # Set paragraph style to FormulaEquationNumbered
        paragraph.style = "FormulaEquationNumbered"

        # Add the equation number in parentheses after a tab
        # First get the XML element
        p_xml = paragraph._element

        # Create a run element for the tab and equation number
        run_xml = create_element("w:r")

        # Add a tab character
        tab_xml = create_element("w:tab")
        run_xml.append(tab_xml)

        # Add the equation number in parentheses
        text_xml = create_element("w:t")
        text_xml.text = f"（{equation_number}）"
        run_xml.append(text_xml)

        # Append the new run to the paragraph
        p_xml.append(run_xml)

        # print(
        #     f"Applied 'FormulaEquationNumbered' style with equation number {equation_number}"
        # )
        return True
    except Exception as e:
        print(f"Error formatting math paragraph: {e}")
        return False


def style_superscript_hyperlinks(doc, style_name="ae"):
    """
    Find all hyperlinked text that is also superscript and apply a specific style.

    Args:
        doc: The Word document
        style_name: The style to apply to superscript hyperlinks
    """
    print("\n=== IDENTIFYING AND STYLING SUPERSCRIPT HYPERLINKS ===")
    count = 0

    # Iterate through all paragraphs in the document
    for paragraph in doc.paragraphs:
        # Get the XML of the paragraph
        p_xml = paragraph._element

        # Find all hyperlinks in this paragraph
        hyperlinks = p_xml.findall(
            ".//w:hyperlink",
            namespaces={
                "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            },
        )

        for hyperlink in hyperlinks:
            # Find runs within this hyperlink
            runs_in_link = hyperlink.findall(
                ".//w:r",
                namespaces={
                    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                },
            )

            for run_xml in runs_in_link:
                # Check if this run has superscript formatting
                rPr = run_xml.find(
                    ".//w:rPr",
                    namespaces={
                        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    },
                )
                if rPr is not None:
                    vert_align = rPr.find(
                        ".//w:vertAlign[@w:val='superscript']",
                        namespaces={
                            "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                        },
                    )

                    if vert_align is not None:
                        # This run is both hyperlinked and superscript!
                        count += 1

                        # Get the text content for logging
                        text_elements = run_xml.findall(
                            ".//w:t",
                            namespaces={
                                "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            },
                        )
                        text_content = "".join(
                            [t.text for t in text_elements if t.text]
                        )

                        # Apply the specified style
                        style_element = create_element("w:rStyle")
                        create_attribute(style_element, "w:val", style_name)

                        # Add or replace the style in the run properties
                        existing_style = rPr.find(
                            ".//w:rStyle",
                            namespaces={
                                "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            },
                        )
                        assert existing_style is not None, (
                            "Run properties must contain rStyle"
                        )
                        rPr.replace(existing_style, style_element)

                        statistics_data[text_content] = (
                            statistics_data.get(text_content, 0) + 1
                        )

    print(
        f"\n=== COMPLETED: Styled {count} superscript hyperlinks with '{style_name}' style ===\n"
    )


def force_update_fields(doc):
    """强制 Word 在打开时提示更新域（包括目录）"""
    element = doc.settings.element
    update_fields = OxmlElement("w:updateFields")
    update_fields.set(qn("w:val"), "true")
    element.append(update_fields)


# --- Entry Point ---
if __name__ == "__main__":
    INPUT = "demo.md"
    REF_FILE = "cppref.bib"
    output = f"{TITLE}-generated.docx"

    shutil.make_archive("reference", "zip", "reference")
    if os.path.exists("reference.docx"):
        os.remove("reference.docx")
    os.rename("reference.zip", "reference.docx")
    assert (
        os.system(
            f"pandoc {INPUT} -o {output} --filter pandoc-crossref --reference-doc reference.docx --citeproc --csl GB-T-7714—2015（顺序编码，双语，姓名不大写，无URL、DOI，引注有页码）.csl --bibliography {REF_FILE}"
        )
        == 0
    ), "pandoc execution failed"

    doc = Document(output)
    assert doc is not None, "Failed to load the document"
    process_document(doc)
    doc.save(output)
    print("Output file saved as:", output)
