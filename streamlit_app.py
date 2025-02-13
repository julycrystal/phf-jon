import streamlit as st
from docx import Document
from docx.shared import Inches
import io


def docx_to_bytes(doc_object):
    bytes_io = io.BytesIO()
    doc_object.save(bytes_io)
    bytes_io.seek(0)
    return bytes_io.getvalue()


# Show title and description.
st.title("📄 PHF Editor for JC")
st.write(
    "Upload a document below and ask a question about it – GPT will answer! "
    "To use this app, you need to provide an OpenAI API key, which you can get [here](https://platform.openai.com/account/api-keys). "
)

# cover_file = st.file_uploader(
#     "Upload Cover page document (.doc, .docx)", type=("doc", "docx")
# )

phf_file = st.file_uploader(
    "Upload PHF document (.doc, .docx)", type=("doc", "docx")
)

# Ask the user for a question via `st.text_area`.
# question = st.text_area(
#     "Now ask a question about the document!",
#     placeholder="Can you give me a short summary?",
#     disabled=not uploaded_file,
# )

if phf_file:
    # cover_doc = Document(cover_file)
    phf_doc = Document(phf_file)

    # Format tables (widen table width to full)
    # for table_index, table in enumerate(phf_doc.tables):
    #     table.columns[0].cells[0].width = Inches(6)

    for _, table in enumerate(phf_doc.tables):
        table.autofit = True
        table.cell(0, 0).tables[0].autofit = False
        # table.cell(0, 0).tables[0]._tbl.tblGrid.gridCol_lst[0].w = Inches(4)
        # table.cell(0, 0).tables[0]._tbl.tr_lst[0].tc_lst[0].tcPr.tcW.w = Inches(4)
        table.cell(0, 0).tables[0]._tbl.tblGrid.gridCol_lst[1].w = Inches(10)

    # Fix lines on the bottom of page

    # Replace "Instructions" to "Recommendations"

    # Extract data to excel sheet with patient info and medications list

    # Let Jon add notes to Recommendations under each patient

    # Save as password protected PDF
    st.download_button(
        label = "Download",
        data = docx_to_bytes(phf_doc),
        file_name = "test.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
