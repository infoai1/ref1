import re
from io import BytesIO
import streamlit as st
from docx import Document

st.set_page_config(page_title="DOCX: Inline full refs (safe)", page_icon="📚")
st.title("Inline full references for the whole DOCX (all chapters) – SAFE MODE")

fmt = st.selectbox("Inline format", ["— 1. Reference text", "[1. Reference text]", "(Reference text)"], index=0)
delete_notes = st.checkbox("Delete Notes/References sections after inlining", value=False)
also_replace_parentheses = st.checkbox("Also convert (n)/(1–3). Risky near years; keep OFF.", value=False)
uploaded = st.file_uploader("Upload the full book .docx", type=["docx"])

HEADING_RE = re.compile(r"^\s*(notes?|references|endnotes?|sources)\s*:?\s*$", re.I)

def para_iter(doc):
    for p in doc.paragraphs:
        yield p
    for t in doc.tables:
        for r in t.rows:
            for c in r.cells:
                for p in c.paragraphs:
                    yield p

def is_notes_heading(p):
    return bool(HEADING_RE.match(p.text or ""))

def is_heading_style(p):
    return (getattr(p.style, "name", "") or "").lower().startswith("heading")

def next_block_end(paragraphs, start):
    blanks = 0
    for j in range(start, len(paragraphs)):
        pj = paragraphs[j]
        if is_notes_heading(pj): return j
        if is_heading_style(pj) and pj.text.strip(): return j
        if pj.text.strip() == "":
            blanks += 1
            if blanks >= 2: return j
        else:
            blanks = 0
    return len(paragraphs)

def strip_leading_num(s):
    m = re.match(r"^\s*(\d+)[\.\)\]\-–—:]?\s*(.*)$", s.strip())
    if m: return int(m.group(1)), m.group(2).strip()
    return None, s.strip()

def parse_notes(paragraphs, start, end):
    refmap, expected = {}, 1
    for p in paragraphs[start:end]:
        txt = (p.text or "").strip()
        if not txt: continue
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
        if not part: continue
        if "-" in part:
            a,b = part.split("-",1)
            if a.isdigit() and b.isdigit():
                a,b = int(a), int(b)
                if a <= b and (b - a) <= 20:  # prevent huge explosions
                    out.extend(range(a, b+1))
        elif part.isdigit():
            out.append(int(part))
    return out

def render(num, text, style):
    if style.startswith("—"): return f" — {num}. {text}"
    if style.startswith("["): return f" [{num}. {text}]"
    return f" ({text})"

def replace_in_body(paragraphs, start, end, refmap, style, allow_paren=False):
    changed = 0
    if not refmap: return 0
    max_key = max(refmap.keys())

    # Helper: only replace if all numbers exist in refmap, none look like years
    def safe_to_replace(nums):
        if not nums: return False
        if any(n >= 1000 for n in nums):  # treat 4-digit as year
            return False
        return all(1 <= n <= max_key and n in refmap for n in nums)

    for i in range(start, end):
        p = paragraphs[i]
        if is_notes_heading(p): continue

        # superscripts
        for r in p.runs:
            if getattr(r.font, "superscript", None):
                nums = [int(x) for x in re.findall(r"\d+", r.text)]
                if safe_to_replace(nums):
                    parts = [render(n, refmap[n], style).lstrip() for n in nums]
                    r.font.superscript = None
                    r.text = "; ".join(parts)
                    changed += 1

        # square brackets ONLY (common citation form)
        orig = p.text
        def repl_sq(m):
            nums = expand_nums(m.group(1))
            if not safe_to_replace(nums): return m.group(0)
            return "; ".join(render(n, refmap[n], style) for n in nums)

        new = re.sub(r"\[\s*([0-9,\-\–—\s]+)\s*\]", repl_sq, p.text)

        # optional: parentheses, off by default because of (1801–1876)
        if allow_paren:
            def repl_paren(m):
                nums = expand_nums(m.group(1))
                if not safe_to_replace(nums): return m.group(0)
                return "; ".join(render(n, refmap[n], style) for n in nums)
            new2 = re.sub(r"\(\s*([0-9,\-\–—\s]+)\s*\)", repl_paren, new)
        else:
            new2 = new

        if new2 != orig:
            p.text = new2
            changed += 1
    return changed

def delete_block(paragraphs, start, end):
    for idx in range(end-1, start-1, -1):
        p = paragraphs[idx]
        p._element.getparent().remove(p._element)

if uploaded:
    doc = Document(uploaded)
    paragraphs = list(para_iter(doc))
    heads = [i for i,p in enumerate(paragraphs) if is_notes_heading(p)]
    if not heads:
        st.error("No Notes/References sections found.")
        st.stop()

    blocks, prev_end = [], 0
    for i in heads:
        notes_start = i + 1
        notes_end = next_block_end(paragraphs, notes_start)
        refmap = parse_notes(paragraphs, notes_start, notes_end)
        blocks.append((prev_end, i, notes_start, notes_end, refmap))
        prev_end = notes_end

    total = 0
    for (body_start, head_i, notes_start, notes_end, refmap) in blocks:
        total += replace_in_body(paragraphs, body_start, head_i, refmap, fmt, allow_paren=also_replace_parentheses)

    if delete_notes:
        paragraphs = list(para_iter(doc))
        for (body_start, head_i, notes_start, notes_end, _) in reversed(blocks):
            delete_block(paragraphs, head_i, notes_end)

    bio = BytesIO()
    doc.save(bio); bio.seek(0)
    st.success(f"Inlined {total} citation spot(s) across the document (years preserved).")
    st.download_button("Download DOCX", data=bio.getvalue(),
        file_name="book_inlined_references_SAFE.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
