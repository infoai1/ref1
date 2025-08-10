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
    """Iterate through all paragraphs in document including those in tables"""
    for p in doc.paragraphs:
        yield p
    for t in doc.tables:
        for r in t.rows:
            for c in r.cells:
                for p in c.paragraphs:
                    yield p

def is_notes_heading(p):
    """Check if paragraph is a Notes/References heading"""
    return bool(HEADING_RE.match(p.text or ""))

def is_heading_style(p):
    """Check if paragraph has heading style"""
    style_name = getattr(p.style, "name", "") or ""
    return style_name.lower().startswith("heading")

def next_block_end(paragraphs, start):
    """Find the end of the notes block"""
    blanks = 0
    for j in range(start, len(paragraphs)):
        pj = paragraphs[j]
        if is_notes_heading(pj): 
            return j
        if is_heading_style(pj) and pj.text.strip(): 
            return j
        if pj.text.strip() == "":
            blanks += 1
            if blanks >= 3:  # Increased threshold for better detection
                return j
        else:
            blanks = 0
    return len(paragraphs)

def strip_leading_num(s):
    """Extract leading number from reference text"""
    # More flexible regex to handle various numbering formats
    m = re.match(r"^\s*(\d+)[\.\)\]\-–—:\s]*(.*)$", s.strip())
    if m: 
        return int(m.group(1)), m.group(2).strip()
    return None, s.strip()

def parse_notes(paragraphs, start, end):
    """Parse notes section to create reference mapping"""
    refmap = {}
    expected = 1
    current_ref_text = ""
    current_num = None
    
    for i in range(start, end):
        p = paragraphs[i]
        txt = (p.text or "").strip()
        if not txt: 
            continue
            
        n, rest = strip_leading_num(txt)
        
        if n is not None:
            # Save previous reference if exists
            if current_num is not None and current_ref_text:
                refmap[current_num] = current_ref_text.strip()
            
            # Start new reference
            current_num = n
            current_ref_text = rest
            expected = n + 1
        else:
            # Continuation of previous reference
            if current_num is not None:
                current_ref_text += " " + txt
    
    # Save the last reference
    if current_num is not None and current_ref_text:
        refmap[current_num] = current_ref_text.strip()
    
    return refmap

def expand_nums(token_str):
    """Expand number ranges like '1-3' to [1, 2, 3]"""
    s = token_str.replace("–", "-").replace("—", "-")
    out = []
    for part in re.split(r"[,\s]+", s):
        if not part: 
            continue
        if "-" in part:
            try:
                a, b = part.split("-", 1)
                if a.isdigit() and b.isdigit():
                    a, b = int(a), int(b)
                    if a <= b and (b - a) <= 20:  # Prevent huge expansions
                        out.extend(range(a, b + 1))
            except:
                pass
        elif part.isdigit():
            out.append(int(part))
    return out

def render(num, text, style):
    """Format the inline reference according to selected style"""
    if style.startswith("—"): 
        return f" — {num}. {text}"
    elif style.startswith("["): 
        return f" [{num}. {text}]"
    else:
        return f" ({text})"

def replace_in_runs(paragraph, refmap, style, allow_paren=False):
    """Replace citations while preserving document formatting"""
    if not refmap:
        return 0
    
    changed = 0
    max_key = max(refmap.keys()) if refmap else 0

    def safe_to_replace(nums):
        """Check if numbers are safe to replace (not years, exist in refmap)"""
        if not nums: 
            return False
        if any(n >= 1000 for n in nums):  # Treat 4-digit as years
            return False
        return all(1 <= n <= max_key and n in refmap for n in nums)

    # Handle superscript citations
    for run in paragraph.runs:
        if getattr(run.font, "superscript", None):
            nums = [int(x) for x in re.findall(r"\d+", run.text)]
            if safe_to_replace(nums):
                parts = [render(n, refmap[n], style).lstrip() for n in nums]
                run.font.superscript = None
                run.text = "; ".join(parts)
                changed += 1

    # Handle bracketed citations and optionally parenthetical ones
    full_text = paragraph.text
    original_text = full_text
    
    # Square brackets (common citation format)
    def repl_sq(m):
        nums = expand_nums(m.group(1))
        if not safe_to_replace(nums): 
            return m.group(0)
        return "; ".join(render(n, refmap[n], style) for n in nums)

    full_text = re.sub(r"\[\s*([0-9,\-\–—\s]+)\s*\]", repl_sq, full_text)

    # Parentheses (optional, risky near years)
    if allow_paren:
        def repl_paren(m):
            nums = expand_nums(m.group(1))
            if not safe_to_replace(nums): 
                return m.group(0)
            return "; ".join(render(n, refmap[n], style) for n in nums)
        
        full_text = re.sub(r"\(\s*([0-9,\-\–—\s]+)\s*\)", repl_paren, full_text)

    # Only update if text changed
    if full_text != original_text:
        # Clear existing runs and add new text
        for run in paragraph.runs:
            run.text = ""
        if paragraph.runs:
            paragraph.runs[0].text = full_text
        else:
            paragraph.add_run(full_text)
        changed += 1

    return changed

def replace_in_body(paragraphs, start, end, refmap, style, allow_paren=False):
    """Replace citations in the body text"""
    changed = 0
    for i in range(start, end):
        p = paragraphs[i]
        if is_notes_heading(p): 
            continue
        changed += replace_in_runs(p, refmap, style, allow_paren)
    return changed

def delete_block(doc, paragraphs, start, end):
    """Delete a block of paragraphs"""
    to_delete = []
    for idx in range(start, end):
        if idx < len(paragraphs):
            to_delete.append(paragraphs[idx])
    
    for p in to_delete:
        if p._element.getparent() is not None:
            p._element.getparent().remove(p._element)

if uploaded:
    try:
        doc = Document(uploaded)
        paragraphs = list(para_iter(doc))
        
        # Find all Notes/References sections
        heads = [i for i, p in enumerate(paragraphs) if is_notes_heading(p)]
        
        if not heads:
            st.error("No Notes/References sections found in the document.")
            st.stop()

        # Process each notes section
        blocks = []
        prev_end = 0
        
        for i in heads:
            notes_start = i + 1
            notes_end = next_block_end(paragraphs, notes_start)
            refmap = parse_notes(paragraphs, notes_start, notes_end)
            blocks.append((prev_end, i, notes_start, notes_end, refmap))
            prev_end = notes_end
        
        # Add final block if document continues after last notes section
        if prev_end < len(paragraphs):
            blocks.append((prev_end, len(paragraphs), len(paragraphs), len(paragraphs), {}))

        # Replace citations in body text
        total_replaced = 0
        for (body_start, head_i, notes_start, notes_end, refmap) in blocks:
            if refmap:  # Only process if we found references
                replaced = replace_in_body(paragraphs, body_start, head_i, refmap, fmt, 
                                         allow_paren=also_replace_parentheses)
                total_replaced += replaced
                
                st.write(f"Found {len(refmap)} references, replaced {replaced} citations in section")

        # Delete notes sections if requested
        if delete_notes and blocks:
            # Refresh paragraph list after modifications
            paragraphs = list(para_iter(doc))
            heads = [i for i, p in enumerate(paragraphs) if is_notes_heading(p)]
            
            # Delete in reverse order to maintain indices
            for head_idx in reversed(heads):
                notes_start = head_idx + 1
                notes_end = next_block_end(paragraphs, notes_start)
                delete_block(doc, paragraphs, head_idx, notes_end)

        # Save the modified document
        bio = BytesIO()
        doc.save(bio)
        bio.seek(0)
        
        st.success(f"Successfully inlined {total_replaced} citation(s) across the document!")
        st.download_button(
            "Download Modified DOCX",
            data=bio.getvalue(),
            file_name="book_inlined_references_SAFE.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        
    except Exception as e:
        st.error(f"Error processing document: {str(e)}")
        st.write("Please check that the uploaded file is a valid DOCX document.")

else:
    st.info("Please upload a DOCX file to begin processing.")
