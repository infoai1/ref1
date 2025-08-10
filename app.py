import re
from io import BytesIO
import streamlit as st
from docx import Document

st.set_page_config(page_title="DOCX: Inline full refs for whole book", page_icon="📚")
st.title("Inline full references for the whole DOCX (all chapters)")

fmt = st.selectbox(
    "Inline format",
    ["— 1. Reference text", "[1. Reference text]", "(Reference text)"],
    index=0,
)
delete_notes = st.checkbox("Delete Notes/References sections after inlining", value=False)
uploaded = st.file_uploader("Upload the full book .docx", type=["docx"])

# ---- helpers ----
HEADING_RE = re.compile(r"^\s*(notes?|references|endnotes?|sources)\s*:?\s*$", re.I)

def para_iter(doc):
    # paragraphs in body + tables (recursive)
    for p in doc.paragraphs:
        yield p
    for t in doc.tables:
        yield from table_para_iter(t)

def table_para_iter(table):
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                yield p
            for t2 in cell.tables:
                yield from table_para_iter(t2)

def is_notes_heading(p):
    return bool(HEADING_RE.match(p.text or ""))

def is_heading_style(p):
    name = (getattr(p.style, "name", "") or "").lower()
    return name.startswith("heading")

def next_block_end(paragraphs, start):
    blanks = 0
    for j in range(start, len(paragraphs)):
        pj = paragraphs[j]
        if is_notes_heading(pj):
            return j
        if is_heading_style(pj) and pj.text.strip():
            return j
        if pj.text.strip() == "":
            blanks += 1
            if blanks >= 2:
                return j
        else:
            blanks = 0
    return len(paragraphs)

def strip_leading_num(s):
    m = re.match(r"^\s*(\d+)[\.\)\]\-–—:]?\s*(.*)$", s.strip())
    if m:
        return int(m.group(1)), m.group(2).strip()
    return None, s.strip()

def parse_notes(paragraphs, start, end):
    refmap, expected = {}, 1
    for p in paragraphs[start:end]:
        txt = (p.text or "").strip()
        if not txt:
            continue
        n, rest = strip_leading_num(txt)
        if n is None:
            n, rest = expected, txt
        refmap[n] = rest
        expected = n + 1
    return refmap

def expand_nums(token_str):
    s = token_str.replace("–", "-").replace("—", "-")
    out = []
    for part in re.split(r"[,\s]+", s):
        if not part:
            continue
        if "-" in part:
            a,b = part.split("-",1)
            if a.isdigit() and b.isdigit():
                out.extend(range(int(a), int(b)+1))
        elif part.isdigit():
            out.append(int(part))
    return out

def render(num, text, style):
    if style.startswith("—"):
        return f" — {num}. {text}"
    if style.startswith("["):
        return f" [{num}. {text}]"
    return f" ({text})"

def replace_in_body(paragraphs, start, end, refmap, style):
    changed = 0
    for i in range(start, end):
        p = paragraphs[i]
        if is_notes_heading(p):
            continue
        # superscript runs
        for r in p.runs:
            if getattr(r.font, "superscript", None):
                nums = expand_nums(r.text)
                if nums:
                    pieces = []
                    for n in nums:
                        txt = refmap.get(n)
                        pieces.append(render(n, txt if txt else f"[{n}]", style).lstrip())
                    r.font.superscript = None
                    r.text = "; ".join(pieces)
                    changed += 1
        # baseline bracket or paren numbers like [1], (1-3)
        orig = p.text
        def repl(m):
            nums = expand_nums(m.group(1))
            pieces = []
            for n in nums:
                txt = refmap.get(n)
                pieces.append(render(n, txt if txt else f"[{n}]", style))
            return "; ".join(pieces)
        new = re.sub(r"[\[\(]\s*([0-9,\-\–—\s]+)\s*[\]\)]", repl, p.text)
        if new != orig:
            p.text = new
            changed += 1
    return changed

def delete_block(paragraphs, start, end):
    # remove from bottom to top to keep indices valid
    for idx in range(end-1, start-1, -1):
        p = paragraphs[idx]
        p._element.getparent().remove(p._element)

# ---- main ----
if uploaded:
    doc = Document(uploaded)
    paragraphs = list(para_iter(doc))

    # find ALL notes blocks first
    note_heads = [i for i,p in enumerate(paragraphs) if is_notes_heading(p)]
    if not note_heads:
        st.error("No Notes/References sections found.")
        st.stop()

    # Build segments: body before first Notes, then Notes, etc.
    blocks = []
    prev_end = 0
    for i in note_heads:
        notes_start = i + 1
        notes_end = next_block_end(paragraphs, notes_start)
        refmap = parse_notes(paragraphs, notes_start, notes_end)
        blocks.append((prev_end, i, notes_start, notes_end, refmap))
        prev_end = notes_end

    total = 0
    for (body_start, head_i, notes_start, notes_end, refmap) in blocks:
        total += replace_in_body(paragraphs, body_start, head_i, refmap, fmt)

    if delete_notes:
        # delete after replacements; recalc paragraph list each deletion batch
        for (body_start, head_i, notes_start, notes_end, _) in reversed(blocks):
            paragraphs = list(para_iter(doc))
            # find the current indices again by matching text pointers (fallback by proximity)
            # simplest robust: delete a span around notes_start..notes_end using previous texts
            delete_block(paragraphs, head_i, notes_end)

    bio = BytesIO()
    doc.save(bio); bio.seek(0)
    st.success(f"Inlined {total} citation spot(s) across the whole document.")
    st.download_button(
        "Download DOCX",
        data=bio.getvalue(),
        file_name="book_inlined_references.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
