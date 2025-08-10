import streamlit as st
from io import BytesIO
from docx import Document

st.set_page_config(page_title="DOCX Superscript → Inline", page_icon="📝")

st.title("DOCX: Superscript references → Inline")
st.caption("Upload a .docx. I’ll replace superscript citation numbers with inline text.")

style = st.radio(
    "Inline style",
    options=["n", "[n]", "(n)"],
    index=1,
    help="How the number should look in body text."
)

uploaded = st.file_uploader("Upload a .docx file", type=["docx"])

def iter_paragraphs_in_table(table):
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                yield p
            for t in cell.tables:
                yield from iter_paragraphs_in_table(t)

def all_paragraphs(doc: Document):
    for p in doc.paragraphs:
        yield p
    for t in doc.tables:
        yield from iter_paragraphs_in_table(t)

def render_number(txt: str, style: str) -> str:
    n = txt.strip()
    if not n.isdigit():
        return txt  # don't touch things like "a" or mixed text
    if style == "n":
        return n
    if style == "[n]":
        return f"[{n}]"
    return f"({n})"

def transform_docx(file_bytes: bytes, style: str):
    doc = Document(BytesIO(file_bytes))
    changed = 0

    for p in all_paragraphs(doc):
        # Optional: avoid touching the Notes sections (case-insensitive exact match)
        if p.text.strip().lower() == "notes":
            continue

        for r in p.runs:
            # python-docx: True/False/None — treat anything truthy as superscript
            if getattr(r.font, "superscript", None):
                original = r.text
                r.font.superscript = None  # make it baseline
                r.text = render_number(original, style)
                changed += 1

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio, changed

if uploaded:
    out, changed = transform_docx(uploaded.read(), style)
    st.success(f"Converted {changed} superscript references to inline.")
    st.download_button(
        label="Download converted DOCX",
        data=out.getvalue(),
        file_name="converted_inline_refs.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
