import re
import os
import zipfile
from io import BytesIO
from pathlib import Path
import streamlit as st
from docx import Document
from docx.shared import Pt

st.set_page_config(page_title="DOCX: Chapter-wise Citation Processor", page_icon="📚")
st.title("Chapter-wise DOCX Citation Processor")

st.markdown("""
### Workflow:
1. **Split**: Divide DOCX into chapters (based on font size 26)
2. **Process**: Handle citations in each chapter individually  
3. **Rejoin**: Combine all processed chapters back into one document
""")

fmt = st.selectbox("Inline format", ["— 1. Reference text", "[1. Reference text]", "(Reference text)"], index=0)
delete_notes = st.checkbox("Delete Notes/References sections after inlining", value=False)
also_replace_parentheses = st.checkbox("Also convert (n)/(1–3). Risky near years; keep OFF.", value=False)
uploaded = st.file_uploader("Upload the full book .docx", type=["docx"])

# Configuration
CHAPTER_FONT_SIZE = 26  # Font size that indicates chapter headers
HEADING_RE = re.compile(r"^\s*(notes?|references?|endnotes?|sources?|bibliography|citations?)\s*:?\s*$", re.I)

def get_font_size(run):
    """Get font size of a run in points"""
    if run.font.size:
        return run.font.size.pt
    return None

def is_chapter_header(paragraph):
    """Check if paragraph is a chapter header (font size 26)"""
    for run in paragraph.runs:
        if get_font_size(run) == CHAPTER_FONT_SIZE:
            return True
    return False

def find_chapter_boundaries(doc):
    """Find all chapter boundaries in the document"""
    chapters = []
    current_chapter_start = 0
    
    for i, paragraph in enumerate(doc.paragraphs):
        if is_chapter_header(paragraph):
            # If we're not at the beginning, save the previous chapter
            if i > 0:
                chapters.append((current_chapter_start, i - 1, get_chapter_title(doc.paragraphs, current_chapter_start, i - 1)))
            current_chapter_start = i
    
    # Add the last chapter
    if current_chapter_start < len(doc.paragraphs):
        chapters.append((current_chapter_start, len(doc.paragraphs) - 1, get_chapter_title(doc.paragraphs, current_chapter_start, len(doc.paragraphs) - 1)))
    
    return chapters

def get_chapter_title(paragraphs, start, end):
    """Extract chapter title from the first few paragraphs"""
    for i in range(start, min(start + 3, end + 1)):
        if i < len(paragraphs):
            text = paragraphs[i].text.strip()
            if text and len(text) < 100:  # Reasonable title length
                # Clean up the title for filename
                clean_title = re.sub(r'[^\w\s-]', '', text)
                clean_title = re.sub(r'\s+', '_', clean_title)
                return clean_title[:50]  # Limit length
    return f"Chapter_{start}"

def create_chapter_document(original_doc, start_para, end_para):
    """Create a new document with paragraphs from start_para to end_para"""
    new_doc = Document()
    
    # Copy paragraphs
    for i in range(start_para, end_para + 1):
        if i < len(original_doc.paragraphs):
            old_para = original_doc.paragraphs[i]
            new_para = new_doc.add_paragraph()
            
            # Copy paragraph style
            new_para.style = old_para.style
            
            # Copy runs with formatting
            for run in old_para.runs:
                new_run = new_para.add_run(run.text)
                new_run.font.name = run.font.name
                new_run.font.size = run.font.size
                new_run.font.bold = run.font.bold
                new_run.font.italic = run.font.italic
                new_run.font.underline = run.font.underline
                if hasattr(run.font, 'superscript'):
                    new_run.font.superscript = run.font.superscript
    
    # Copy tables that fall within the range
    for table in original_doc.tables:
        # This is a simplified approach - in practice, you'd need to check
        # if the table falls within the paragraph range
        pass
    
    return new_doc

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
    return bool(HEADING_RE.match(text)) if text else False

def find_notes_sections(paragraphs):
    """Find all notes sections in the document"""
    sections = []
    for i, p in enumerate(paragraphs):
        if is_notes_heading(p):
            start = i + 1
            end = find_section_end(paragraphs, start)
            sections.append((i, start, end))
    return sections

def find_section_end(paragraphs, start):
    """Find the end of a section"""
    consecutive_blanks = 0
    for i in range(start, len(paragraphs)):
        p = paragraphs[i]
        text = (p.text or "").strip()
        
        if is_notes_heading(p) or (hasattr(p, 'style') and 
                                  getattr(p.style, 'name', '').lower().startswith('heading')):
            return i
        
        if not text:
            consecutive_blanks += 1
            if consecutive_blanks >= 2:
                return i
        else:
            consecutive_blanks = 0
    return len(paragraphs)

def parse_references_advanced(paragraphs, start, end):
    """Advanced reference parsing"""
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
        
        patterns = [
            r"^(\d+)\.\s*(.+)$",
            r"^(\d+)\)\s*(.+)$",
            r"^(\d+)\]\s*(.+)$",
            r"^(\d+)\s+(.+)$",
            r"^(\d+)[\-–—:]\s*(.+)$",
        ]
        
        found_match = False
        for pattern in patterns:
            match = re.match(pattern, text)
            if match:
                if current_num is not None and current_text.strip():
                    refs[current_num] = current_text.strip()
                current_num = int(match.group(1))
                current_text = match.group(2)
                found_match = True
                break
        
        if not found_match and current_num is not None:
            current_text += " " + text
    
    if current_num is not None and current_text.strip():
        refs[current_num] = current_text.strip()
    
    return refs

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
    return sorted(set(nums))

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
    
    # Handle superscript runs
    for run in paragraph.runs:
        if getattr(run.font, 'superscript', None):
            nums = [int(x) for x in re.findall(r'\d+', run.text)]
            if (nums and all(1 <= n <= max_ref_num and n in all_refs for n in nums) 
                and not any(n >= 1000 for n in nums)):
                replacements = [format_reference(n, all_refs[n], style).strip() for n in nums]
                run.font.superscript = None
                run.text = "; ".join(replacements)
                changes += 1
    
    # Handle text-based citations
    original_text = paragraph.text
    modified_text = original_text
    
    def replace_square(match):
        nums = expand_number_range(match.group(1))
        if (nums and all(1 <= n <= max_ref_num and n in all_refs for n in nums) 
            and not any(n >= 1000 for n in nums)):
            replacements = [format_reference(n, all_refs[n], style) for n in nums]
            return "; ".join(replacements)
        return match.group(0)
    
    modified_text = re.sub(r'\[([0-9,\-–—\s]+)\]', replace_square, modified_text)
    
    if allow_paren:
        modified_text = re.sub(r'\(([0-9,\-–—\s]+)\)', replace_square, modified_text)
    
    if modified_text != original_text:
        for run in paragraph.runs:
            run.text = ""
        if paragraph.runs:
            paragraph.runs[0].text = modified_text
        else:
            paragraph.add_run(modified_text)
        changes += 1
    
    return changes

def process_single_chapter(doc, style, allow_paren=False, delete_notes_sections=False):
    """Process a single chapter document"""
    paragraphs = list(para_iter(doc))
    notes_sections = find_notes_sections(paragraphs)
    
    if not notes_sections:
        return 0, 0
    
    # Parse all references
    all_refs = {}
    for section_head, section_start, section_end in notes_sections:
        refs = parse_references_advanced(paragraphs, section_start, section_end)
        all_refs.update(refs)
    
    if not all_refs:
        return 0, 0
    
    # Replace citations
    total_replacements = 0
    notes_ranges = set()
    for section_head, section_start, section_end in notes_sections:
        for i in range(section_head, section_end):
            notes_ranges.add(i)
    
    for i, paragraph in enumerate(paragraphs):
        if i not in notes_ranges and paragraph.text and paragraph.text.strip():
            replacements = replace_citations_in_paragraph(paragraph, all_refs, style, allow_paren)
            total_replacements += replacements
    
    # Delete notes sections if requested
    if delete_notes_sections:
        for section_head, section_start, section_end in reversed(notes_sections):
            for idx in range(section_end - 1, section_head - 1, -1):
                if idx < len(paragraphs):
                    p = paragraphs[idx]
                    if p._element.getparent() is not None:
                        p._element.getparent().remove(p._element)
    
    return len(all_refs), total_replacements

def rejoin_chapters(chapter_docs):
    """Rejoin all processed chapters into one document"""
    final_doc = Document()
    
    for i, chapter_doc in enumerate(chapter_docs):
        if i > 0:
            # Add a page break between chapters
            final_doc.add_page_break()
        
        # Copy all paragraphs from chapter
        for para in chapter_doc.paragraphs:
            new_para = final_doc.add_paragraph()
            new_para.style = para.style
            
            for run in para.runs:
                new_run = new_para.add_run(run.text)
                new_run.font.name = run.font.name
                new_run.font.size = run.font.size
                new_run.font.bold = run.font.bold
                new_run.font.italic = run.font.italic
                new_run.font.underline = run.font.underline
                if hasattr(run.font, 'superscript'):
                    new_run.font.superscript = run.font.superscript
        
        # Copy tables
        for table in chapter_doc.tables:
            new_table = final_doc.add_table(rows=len(table.rows), cols=len(table.columns))
            for i, row in enumerate(table.rows):
                for j, cell in enumerate(row.cells):
                    new_table.cell(i, j).text = cell.text
    
    return final_doc

# Main Streamlit app
if uploaded:
    try:
        doc = Document(uploaded)
        
        # Step 1: Analyze and split into chapters
        st.subheader("📖 Step 1: Chapter Analysis")
        chapters = find_chapter_boundaries(doc)
        
        if not chapters:
            st.warning("No chapters found (no font size 26 text detected). Processing as single document...")
            chapters = [(0, len(doc.paragraphs) - 1, "Full_Document")]
        
        st.write(f"Found **{len(chapters)}** chapters:")
        for i, (start, end, title) in enumerate(chapters):
            st.write(f"- Chapter {i+1:02d}: {title} (paragraphs {start}-{end})")
        
        if st.button("🚀 Process All Chapters"):
            chapter_docs = []
            total_refs = 0
            total_replacements = 0
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            # Step 2: Process each chapter
            st.subheader("⚙️ Step 2: Processing Chapters")
            
            for i, (start, end, title) in enumerate(chapters):
                status_text.text(f"Processing Chapter {i+1}/{len(chapters)}: {title}")
                
                # Create chapter document
                chapter_doc = create_chapter_document(doc, start, end)
                
                # Process chapter
                refs_found, citations_replaced = process_single_chapter(
                    chapter_doc, fmt, also_replace_parentheses, delete_notes
                )
                
                chapter_docs.append(chapter_doc)
                total_refs += refs_found
                total_replacements += citations_replaced
                
                st.write(f"✅ Chapter {i+1:02d} - {title}: {refs_found} refs, {citations_replaced} replacements")
                progress_bar.progress((i + 1) / len(chapters))
            
            # Step 3: Rejoin all chapters
            st.subheader("🔗 Step 3: Rejoining Chapters")
            status_text.text("Rejoining all processed chapters...")
            
            final_doc = rejoin_chapters(chapter_docs)
            
            # Save final document
            bio = BytesIO()
            final_doc.save(bio)
            bio.seek(0)
            
            status_text.text("✅ Processing complete!")
            st.success(f"**Final Results**: {total_refs} total references found, {total_replacements} total citations replaced across {len(chapters)} chapters!")
            
            st.download_button(
                "📥 Download Processed Document",
                data=bio.getvalue(),
                file_name="book_processed_citations.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
            # Optional: Download individual chapters
            with st.expander("📁 Download Individual Chapters"):
                for i, (chapter_doc, (start, end, title)) in enumerate(zip(chapter_docs, chapters)):
                    chapter_bio = BytesIO()
                    chapter_doc.save(chapter_bio)
                    chapter_bio.seek(0)
                    
                    st.download_button(
                        f"Chapter {i+1:02d}: {title}",
                        data=chapter_bio.getvalue(),
                        file_name=f"chapter_{i+1:02d}_{title}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"chapter_{i}"
                    )
    
    except Exception as e:
        st.error(f"Error processing document: {str(e)}")
        st.exception(e)

else:
    st.info("📤 Please upload a DOCX file to begin processing.")
    st.markdown("""
    ### How it works:
    1. **Chapter Detection**: Automatically detects chapters based on font size 26
    2. **Individual Processing**: Each chapter is processed separately for better accuracy
    3. **Citation Replacement**: Converts numbered citations to full inline references
    4. **Rejoining**: Combines all processed chapters back into one document
    
    ### Benefits:
    - ✅ Better accuracy per chapter
    - ✅ Easier debugging and troubleshooting  
    - ✅ Memory efficient for large documents
    - ✅ Serial numbering for organization
    """)
