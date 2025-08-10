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

# More comprehensive heading detection
HEADING_RE = re.compile(r"^\s*(notes?|references?|endnotes?|sources?|bibliography|citations?)\s*:?\s*$", re.I)

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
    text = (p.text or "").strip()
    if not text:
        return False
    return bool(HEADING_RE.match(text))

def is_heading_style(p):
    """Check if paragraph has heading style"""
    style_name = getattr(p.style, "name", "") or ""
    return style_name.lower().startswith("heading")

def find_notes_sections(paragraphs):
    """Find all notes sections in the document"""
    sections = []
    i = 0
    while i < len(paragraphs):
        p = paragraphs[i]
        if is_notes_heading(p) or (is_heading_style(p) and HEADING_RE.match(p.text or "")):
            # Found a notes section
            start = i + 1
            end = find_section_end(paragraphs, start)
            sections.append((i, start, end))
            i = end
        else:
            i += 1
    return sections

def find_section_end(paragraphs, start):
    """Find the end of a section (notes, chapter, etc.)"""
    consecutive_blanks = 0
    
    for i in range(start, len(paragraphs)):
        p = paragraphs[i]
        text = (p.text or "").strip()
        
        # Stop at next heading
        if is_notes_heading(p) or is_heading_style(p):
            return i
        
        # Count blank lines
        if not text:
            consecutive_blanks += 1
            if consecutive_blanks >= 2:
                return i
        else:
            consecutive_blanks = 0
    
    return len(paragraphs)

def parse_all_references(paragraphs, sections):
    """Parse all references from all notes sections"""
    all_refs = {}
    
    for section_head, section_start, section_end in sections:
        st.write(f"Processing notes section: {paragraphs[section_head].text}")
        refs = parse_references_advanced(paragraphs, section_start, section_end)
        st.write(f"Found {len(refs)} references in this section")
        all_refs.update(refs)
    
    return all_refs

def parse_references_advanced(paragraphs, start, end):
    """Advanced reference parsing with better multi-line handling"""
    refs = {}
    current_num = None
    current_text = ""
    
    for i in range(start, end):
        if i >= len(paragraphs):
            break
            
        p = paragraphs[i]
        text = (p.text or "").strip()
        
        if not text:
            continue
        
        # Try multiple patterns for numbered references
        patterns = [
            r"^(\d+)\.\s*(.+)$",          # "1. Reference text"
            r"^(\d+)\)\s*(.+)$",          # "1) Reference text"  
            r"^(\d+)\]\s*(.+)$",          # "1] Reference text"
            r"^(\d+)\s+(.+)$",            # "1 Reference text"
            r"^(\d+)[\-–—:]\s*(.+)$",     # "1— Reference text"
        ]
        
        found_match = False
        for pattern in patterns:
            match = re.match(pattern, text)
            if match:
                # Save previous reference
                if current_num is not None and current_text.strip():
                    refs[current_num] = current_text.strip()
                
                # Start new reference
                current_num = int(match.group(1))
                current_text = match.group(2)
                found_match = True
                break
        
        if not found_match and current_num is not None:
            # This is a continuation line
            current_text += " " + text
    
    # Save the last reference
    if current_num is not None and current_text.strip():
        refs[current_num] = current_text.strip()
    
    return refs

def find_all_citations(text):
    """Find all possible citation patterns in text"""
    citations = []
    
    # Pattern 1: [1], [2,3], [1-3], [1,2-5,7]
    for match in re.finditer(r'\[([0-9,\-–—\s]+)\]', text):
        citations.append(('square', match.group(1), match.span()))
    
    # Pattern 2: Superscript numbers (harder to detect in plain text)
    # We'll handle these in the run-level processing
    
    # Pattern 3: (1), (2,3) - only if enabled and not years
    if also_replace_parentheses:
        for match in re.finditer(r'\(([0-9,\-–—\s]+)\)', text):
            # Skip if looks like years
            nums = re.findall(r'\d+', match.group(1))
            if not any(int(n) >= 1000 for n in nums):
                citations.append(('paren', match.group(1), match.span()))
    
    # Pattern 4: Standalone numbers at end of sentences
    for match in re.finditer(r'[.!?]\s*(\d+)\s*$', text):
        num = int(match.group(1))
        if num < 1000:  # Not a year
            citations.append(('standalone', match.group(1), match.span()))
    
    return citations

def expand_number_range(num_str):
    """Expand number ranges and lists"""
    nums = []
    parts = re.split(r'[,\s]+', num_str.replace('–', '-').replace('—', '-'))
    
    for part in parts:
        part = part.strip()
        if not part:
            continue
            
        if '-' in part:
            try:
                start, end = part.split('-', 1)
                start, end = int(start), int(end)
                if 1 <= start <= end <= 999 and (end - start) < 50:
                    nums.extend(range(start, end + 1))
            except:
                pass
        elif part.isdigit():
            nums.append(int(part))
    
    return sorted(set(nums))  # Remove duplicates and sort

def format_reference(num, text, style):
    """Format reference according to selected style"""
    if style.startswith("—"):
        return f" — {num}. {text}"
    elif style.startswith("["):
        return f" [{num}. {text}]"
    else:
        return f" ({text})"

def replace_citations_in_paragraph(paragraph, all_refs, style, allow_paren=False):
    """Replace citations in a single paragraph"""
    if not all_refs:
        return 0
    
    changes = 0
    max_ref_num = max(all_refs.keys())
    
    # Handle superscript runs first
    for run in paragraph.runs:
        if getattr(run.font, 'superscript', None):
            original_text = run.text
            nums = [int(x) for x in re.findall(r'\d+', original_text)]
            
            # Check if all numbers are valid references and not years
            if (nums and all(1 <= n <= max_ref_num and n in all_refs for n in nums) 
                and not any(n >= 1000 for n in nums)):
                
                # Replace with inline references
                replacements = [format_reference(n, all_refs[n], style).strip() for n in nums]
                run.font.superscript = None
                run.text = "; ".join(replacements)
                changes += 1
    
    # Handle text-based citations
    original_text = paragraph.text
    modified_text = original_text
    
    # Process square brackets [1], [2,3], etc.
    def replace_square(match):
        nums = expand_number_range(match.group(1))
        if (nums and all(1 <= n <= max_ref_num and n in all_refs for n in nums) 
            and not any(n >= 1000 for n in nums)):
            replacements = [format_reference(n, all_refs[n], style) for n in nums]
            return "; ".join(replacements)
        return match.group(0)  # Keep original if not valid
    
    modified_text = re.sub(r'\[([0-9,\-–—\s]+)\]', replace_square, modified_text)
    
    # Process parentheses if enabled
    if allow_paren:
        def replace_paren(match):
            nums = expand_number_range(match.group(1))
            if (nums and all(1 <= n <= max_ref_num and n in all_refs for n in nums) 
                and not any(n >= 1000 for n in nums)):
                replacements = [format_reference(n, all_refs[n], style) for n in nums]
                return "; ".join(replacements)
            return match.group(0)
        
        modified_text = re.sub(r'\(([0-9,\-–—\s]+)\)', replace_paren, modified_text)
    
    # Update paragraph text if changed
    if modified_text != original_text:
        # Clear all runs and set new text
        for run in paragraph.runs:
            run.text = ""
        if paragraph.runs:
            paragraph.runs[0].text = modified_text
        else:
            paragraph.add_run(modified_text)
        changes += 1
    
    return changes

def process_document(doc, style, allow_paren=False, delete_notes_sections=False):
    """Main document processing function"""
    paragraphs = list(para_iter(doc))
    
    # Find all notes sections
    notes_sections = find_notes_sections(paragraphs)
    
    if not notes_sections:
        st.error("No Notes/References sections found!")
        return 0, 0
    
    st.write(f"Found {len(notes_sections)} notes sections")
    
    # Parse all references
    all_refs = parse_all_references(paragraphs, notes_sections)
    st.write(f"Total references parsed: {len(all_refs)}")
    
    if not all_refs:
        st.error("No references were successfully parsed!")
        return len(all_refs), 0
    
    # Replace citations in body text (everything except notes sections)
    total_replacements = 0
    notes_ranges = set()
    
    # Mark all notes section ranges
    for section_head, section_start, section_end in notes_sections:
        for i in range(section_head, section_end):
            notes_ranges.add(i)
    
    # Process body paragraphs
    for i, paragraph in enumerate(paragraphs):
        if i not in notes_ranges and paragraph.text and paragraph.text.strip():
            replacements = replace_citations_in_paragraph(paragraph, all_refs, style, allow_paren)
            total_replacements += replacements
    
    # Delete notes sections if requested
    if delete_notes_sections:
        # Delete in reverse order to maintain indices
        for section_head, section_start, section_end in reversed(notes_sections):
            for idx in range(section_end - 1, section_head - 1, -1):
                if idx < len(paragraphs):
                    p = paragraphs[idx]
                    if p._element.getparent() is not None:
                        p._element.getparent().remove(p._element)
    
    return len(all_refs), total_replacements

# Main Streamlit app
if uploaded:
    try:
        doc = Document(uploaded)
        
        if st.button("Process Document"):
            with st.spinner("Processing document..."):
                refs_found, citations_replaced = process_document(
                    doc, fmt, also_replace_parentheses, delete_notes
                )
            
            st.success(f"Found {refs_found} references and replaced {citations_replaced} citations!")
            
            # Save document
            bio = BytesIO()
            doc.save(bio)
            bio.seek(0)
            
            st.download_button(
                "Download Processed DOCX",
                data=bio.getvalue(),
                file_name="book_inlined_references_SAFE.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    
    except Exception as e:
        st.error(f"Error processing document: {str(e)}")
        st.exception(e)

else:
    st.info("Please upload a DOCX file to begin processing.")
    st.markdown("""
    ### Instructions:
    1. Upload your DOCX file with numbered citations
    2. Choose your preferred inline format
    3. Optionally enable parentheses conversion (risky near years)
    4. Optionally choose to delete Notes sections after inlining
    5. Click "Process Document"
    """)
