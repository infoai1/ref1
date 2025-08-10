import re
from io import BytesIO
import streamlit as st
from docx import Document

st.set_page_config(page_title="DOCX: Inline full references from Notes", page_icon="📚")
st.title("DOCX → Replace superscripts with full reference text")

fmt = st.selectbox(
    "Inline format",
    [
        "(Reference text)",
        "— 1. Reference text",
        "[1. Reference text]"
    ],
    index=1,
    help="How the inserted reference should look. For 2+ citations, items are joined with '; '.",
)

delete_notes = st.checkbox("Delete Notes sections after inlining", value=False)
uploaded = st.file_uploader("Upload a .docx chapter (with a 'Notes' section)", type=["docx"])

# -------- helpers --------
def all_paragraphs(doc):
    # linear order: body paragraphs only (tables rarely used for narrative text)
    for p in doc.paragraphs:
        yield p

def is_notes_heading(p):
    return p.text.strip().lower() == "notes"

def next_heading_index(paragraphs, start_i):
    # stop Notes when we hit a Heading or another 'Notes' or two consecutive blanks
    blanks = 0
    for j in range(start_i, len(paragraphs)):
        pj = paragraphs[j]
        if is_notes_heading(pj):
            return j
        name = getattr(pj.style, "name", "") or ""
        if name.lower().startswith("heading") and pj.text.strip() != "":
            return j
        if pj.text.strip() == "":
            blanks += 1
            if blanks >= 2:
                return j
        else:
            blanks = 0
    return len(paragraphs)

def strip_leading_number(s):
    # handles "1. ", "2) ", "[3] ", "3 – ", "3- "
    m = re.match(r"^\s*(\d+)[\.\)\]\-–—:]?\s*(.*)$", s.strip())
    if m:
        return int(m.group(1)), m.group(2).strip()
    return None, s.strip()

def parse_notes_block(paragraphs, start, end):
    """Return dict {num: full_text} from Notes block paragraphs[start:end]."""
    refmap = {}
    expected = 1
    for p in paragraphs[start:end]:
        text = p.text.strip()
        if not text:
            continue
        n, rest = strip_leading_number(text)
        if n is None:
            # numbering might be Word automatic; assign sequential number
            n = expected
            rest = text
        refmap[n] = rest
        expected = n + 1
    return refmap

def build_all_blocks(paragraphs):
    """Find each Notes block and return list of tuples:
       (body_start, body_end, notes_start, notes_end, refmap)
       body range is the text to transform using that block's refs."""
    idxs = [i for i,p in enumerate(paragraphs) if is_notes_heading(p)]
    blocks = []
    prev_end = 0
    for i in idxs:
        notes_start = i + 1
        notes_end = next_heading_index(paragraphs, notes_start)
        refmap = parse_notes_block(paragraphs, notes_start, notes_end)
        blocks.append((prev_end, i, notes_start, notes_end, refmap))
        prev_end = notes_end
    # trailing body after last Notes (rare) -> no mapping
    return blocks

def render_inline(num, text, style):
    if style.startswith("("):
        return f" ({text})"
    if style.startswith("—"):
        return f" — {num}. {text}"
    # bracket style
    return f" [{num}. {text}]"

def replace_superscripts_in_range(paragraphs, start, end, refmap, style):
    changed = 0
    for pi in range(start, end):
        p = paragraphs[pi]
        if is_notes_heading(p):
            continue
        # Replace true superscripts first (preserves other formatting)
        for r in p.runs:
            if getattr(r.font, "superscript", None):
                nums = re.findall(r"\d+", r.text)
                if not nums:
                    continue
                parts = []
                for nstr in nums:
                    n = int(nstr)
                    txt = refmap.get(n)
                    if txt:
                        parts.append(render_inline(n, txt, style).lstrip())
                    else:
                        # fallback: keep original number if not found
                        parts.append(f"[{n}]")
                r.font.superscript = None
                r.text = "; ".join(parts)
                changed += 1
        # Optional: catch bracketed numbers that were not superscripted
        # Example: "... text [1]" (baseline). Do a safe whole-paragraph replace.
        if "[" in p.text and "]" in p.text:
            orig = p.text
            def repl(m):
                n = int(m.group(1))
                txt = refmap.get(n)
                return render_inline(n, txt, style) if txt else m.group(0)
            new_text = re.sub(r"\[(\d+)\]", repl, p.text)
            if new_text != orig:
                # WARNING: setting p.text flattens runs; ok for reference-bearing lines
                p.text = new_text
                changed += 1
    return changed

def delete_notes_block(paragraphs, start, end):
    # python-docx can't delete paragraphs directly; use element removal
    for i in range(start-1, end):  # also remove the "Notes" heading above it
        if 0 <= i < len(paragraphs):
            p = paragraphs[i]
            p._element.getparent().remove(p._element)

# -------- main --------
if uploaded:
    doc = Document(uploaded)
    paragraphs = list(all_paragraphs(doc))
    blocks = build_all_blocks(paragraphs)

    if not blocks:
        st.error("No 'Notes' section found. Ensure each chapter ends with a 'Notes' heading.")
        st.stop()

    total_changes = 0
    for (body_start, body_end, notes_start, notes_end, refmap) in blocks:
        total_changes += replace_superscripts_in_range(
            paragraphs, body_start, body_end, refmap, fmt
        )
        if delete_notes and refmap:
            delete_notes_block(paragraphs, notes_start, notes_end)

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    st.success(f"Inlined {total_changes} citation spot(s).")
    st.download_button(
        "Download DOCX",
        data=bio.getvalue(),
        file_name="inlined_references.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
