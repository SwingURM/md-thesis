import os
from collections import deque

from docx import Document

from header import (
    fix_reference_format,
    insert_section_breaks,
    process_hyperlink,
    process_math_equations,
    process_table,
    replace_ref_format_in_doc,
    set_abstract_font,
    set_all_secs_other,
    set_headers_other,
    update_reference_doc,
    title
)


# --- Main Processing ---
def process_document(doc):
    """Process the Word document."""

    deque(map(process_table, doc.tables))

    insert_section_breaks(doc, 1)

    set_all_secs_other(doc)

    set_headers_other(doc)

    process_math_equations(doc)

    set_abstract_font(doc)

    process_hyperlink(doc)

    replace_ref_format_in_doc(doc)

    fix_reference_format(doc)

    #修改Abstract样式，段前距为0
    style = doc.styles['Abstract']
    style.paragraph_format.space_before = 0


# --- Entry Point ---
if __name__ == "__main__":
    INPUT = "open.md"
    REF_FILE = "cppref.bib"
    output = f"{title}-开题报告-generated.docx"

    update_reference_doc()
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
