import re

from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt


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


def apply_simsun_font(run):
    """Apply SimSun font to a run."""
    run.font.name = "SimSun"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "SimSun")


def copy_sectPr_properties(source_sectPr, target_sectPr):
    """Copy key page settings from source sectPr to target sectPr."""
    properties_to_copy = ["pgSz", "pgMar"]
    for prop_tag in properties_to_copy:
        source_element = source_sectPr.find(qn(f"w:{prop_tag}"))
        if source_element is not None:
            target_element = target_sectPr.find(qn(f"w:{prop_tag}"))
            if target_element is not None:
                target_sectPr.replace(target_element, source_element)
            else:
                target_sectPr.append(source_element)


def add_next_page_section_break(paragraph):
    """Add a next-page section break to a paragraph."""
    if paragraph is None:
        print("Error: Specified paragraph not found.")
        return False

    print(f"Adding section break to paragraph: '{paragraph.text[:30]}...'")
    p_element = paragraph._p
    pPr = p_element.get_or_add_pPr()

    sectPr = pPr.find(qn("w:sectPr"))
    if sectPr is None:
        sectPr = create_element("w:sectPr")
        pPr.append(sectPr)
        try:
            last_section_sectPr = paragraph._parent.part.document.sections[-1]._sectPr
            copy_sectPr_properties(last_section_sectPr, sectPr)
        except Exception as e:
            print(f"Warning: Unable to copy page settings: {e}")

    type_element = sectPr.find(qn("w:type"))
    if type_element is None:
        type_element = create_element("w:type")
        sectPr.append(type_element)
    create_attribute(type_element, "w:val", "nextPage")

    paragraph.text = ""
    for run in paragraph.runs:
        p_element.remove(run._r)
    return True


def is_abstract_paragraph(paragraph):
    """Check if a paragraph is an abstract paragraph."""
    return paragraph.style.name.lower() == "abstract"


def add_toc(document, marker_text="PLACE_TOC_HERE"):
    """
    Insert a Table of Contents (TOC) at the location of a marker text.
    """
    target_paragraph = next(
        (p for p in document.paragraphs if marker_text in p.text), None
    )
    if not target_paragraph:
        print(f"Warning: Marker '{marker_text}' not found in the document.")
        return

    try:
        toc_title_paragraph = target_paragraph.insert_paragraph_before(
            "Table of Contents", style="TOC Heading"
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
        print(f"TOC field inserted at the location of '{marker_text}'.")
    except Exception as e:
        print(f"Error during TOC insertion: {e}")


# --- Main Processing ---
def process_document(file_path, output_path, marker1_text=None, marker2_text=None):
    """Process the Word document."""
    doc = Document(file_path)

    # Process tables
    for table in doc.tables:
        process_table(table)

    # Add TOC
    add_toc(doc)

    # Insert section breaks if markers are provided
    if marker1_text or marker2_text:
        insert_section_breaks(doc, marker1_text, marker2_text)

    # Process sections
    set_page_number_for_all_secs(doc)

    #
    set_headers(doc)

    #
    set_abstract_font(doc)

    #
    replace_figure_format_in_doc(doc)

    # Save the modified document
    doc.save(output_path)
    print(f"Document saved: {output_path}")


def insert_section_breaks(doc, marker1_text, marker2_text):
    """Insert section breaks at specified markers."""
    marker1_paragraph = None
    marker2_paragraph = None

    for para in doc.paragraphs:
        if marker1_text and marker1_text in para.text:
            marker1_paragraph = para
        if marker2_text and marker2_text in para.text:
            marker2_paragraph = para

    success1 = False
    success2 = False

    if marker1_paragraph:
        success1 = add_next_page_section_break(marker1_paragraph)
    else:
        print(f"Error: Marker '{marker1_text}' not found.")

    if marker2_paragraph and marker2_paragraph != marker1_paragraph:
        success2 = add_next_page_section_break(marker2_paragraph)
    elif marker2_text:
        print(f"Error: Marker '{marker2_text}' not found or overlaps with marker 1.")

    if not (success1 or success2):
        print("No section breaks were successfully inserted.")


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
    total_width_twips = 9286  # Total table width
    num_columns = len(table.columns)  # Number of columns
    column_width_twips = total_width_twips // num_columns  # Column width

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

    if num_sections < 3:
        print(
            "Warning: Document has fewer than 3 sections. Ensure proper section breaks."
        )

    # Section 1: Before TOC
    if num_sections >= 1:
        section1 = doc.sections[0]
        section1.footer_distance = Pt(56.7)  # Footer distance from the bottom (2 cm)
        set_page_number_style(section1, fmt="upperRoman", start=1)
        add_page_number_to_footer(section1)

    # Section 2: TOC
    if num_sections >= 2:
        section2 = doc.sections[1]
        section2.footer_distance = Pt(56.7)  # Footer distance from the bottom (2 cm)
        section2.footer.is_linked_to_previous = False
        set_page_number_style(section2, fmt="upperRoman", start=1)
        add_page_number_to_footer(section2)

    # Section 3: After TOC
    if num_sections >= 3:
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

        # Ensure the header has at least one paragraph
        if not header.paragraphs:
            paragraph = header.add_paragraph()
        else:
            paragraph = header.paragraphs[0]

        # Clear existing content in the paragraph
        for run in paragraph.runs:
            run.clear()

        # Add header content
        if i < 2:  # First two sections
            paragraph.add_run("这是页眉内容 ")
        else:  # Third section and beyond
            run = paragraph.add_run("这是页眉内容 ")

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
    for paragraph in doc.paragraphs:
        if is_abstract_paragraph(paragraph):
            text = paragraph.text.strip()
            if text.startswith("关键词："):
                prefix = "关键词："
            elif text.lower().startswith("keywords:"):
                prefix = "Keywords:"
            else:
                raise ValueError("Invalid abstract paragraph format.")
            paragraph.clear()
            run_prefix = paragraph.add_run(prefix)
            remaining = text[len(prefix) :]
            run_rest = paragraph.add_run(remaining)
            apply_simsun_font(run_rest)

def replace_figure_format_in_doc(doc):
    """
    Replace all occurrences of '图%d.%d' with '图%d-%d' in the entire Word document.
    """
    # Iterate through all paragraphs in the document
    for paragraph in doc.paragraphs:
        if paragraph.text:
            # Replace text in the paragraph
            paragraph.text = re.sub(r"图(\d+)\.(\d+)", r"图\1-\2", paragraph.text)

    # Iterate through all tables in the document
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if paragraph.text:
                        # Replace text in the table cell
                        paragraph.text = re.sub(
                            r"图(\d+)\.(\d+)", r"图\1-\2", paragraph.text
                        )


# --- Entry Point ---
if __name__ == "__main__":
    input_docx_path = "open.docx"
    output_docx_path = "open.docx"
    marker1_text = "%%%SECTION_BREAK_1%%%"
    marker2_text = "%%%SECTION_BREAK_2%%%"
    process_document(input_docx_path, output_docx_path, marker1_text, marker2_text)
