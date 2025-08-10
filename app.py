import re
from io import BytesIO
import streamlit as st
from docx import Document

st.set_page_config(page_title="DOCX: Font Size 26 Chapter Processor", page_icon="📚")
st.title("Strict Font Size 26 Chapter Detection & Citation Processing")

st.markdown("""
### Strict Rules:
- **Only font size 26** text will be considered as chapter headers
- **No other font sizes** or detection methods used
- If no font size 26 found, entire document processed as one chapter
""")

fmt = st.selectbox("Inline format", ["— 1. Reference text", "[1. Reference text]", "(Reference text)"], index=0)
delete_notes = st.checkbox("Delete Notes/References sections after inlining", value=False)
also_replace_parentheses = st.checkbox("Also convert (n)/(1–3). Risky near years; keep OFF.", value=False)
uploaded = st.file_uploader("Upload the full book .docx", type=["docx"])

HEADING_RE = re.compile(r"^\s*(notes?|references?|endnotes?|sources?|bibliography|citations?)\s*:?\s*$", re.I)

def get_font_size(run):
    """Get font size of a run in points"""
    if run.font.size:
        return run.font.size.pt
    return None

def find_font26_chapters(doc):
    """Find chapters ONLY based on font size 26"""
    chapters = []
    
    for i, paragraph in enumerate(doc.paragraphs):
        text = paragraph.text.strip()
        if not text:
            continue
            
        # Check if ANY run in paragraph has font size 26
        has_font26 = False
        for run in paragraph.runs:
            if get_font_size(run) == 26:
                has_font26 = True
                break
        
        if has_font26:
            # Clean title for filename
            clean_title = re.sub(r'[^\w\s-]', '', text)
            clean_title = re.sub(r'\s+', '_', clean_title)[:50]
            if not clean_title:
                clean_title = f"Chapter_{len(chapters)+1}"
                
            chapters.append({
                'index': i,
                'title': clean_title,
                'text': text
            })
    
    return chapters

def create_chapter_boundaries(chapters, total_paragraphs):
    """Create chapter boundaries for splitting"""
    if not chapters:
        st.warning("No font size 26 text found. Processing entire document as one chapter.")
        return [(0, total_paragraphs - 1, "Full_Document")]
    
    boundaries = []
    for i, chapter in enumerate(chapters):
        start = chapter['index']
        end = chapters[i + 1]['index'] - 1 if i + 1 < len(chapters) else total_paragraphs - 1
        boundaries.append((start, end, chapter['title']))
    
    return boundaries

def create_chapter_document(original_doc, start_para, end_para):
    """Create a new document with specified paragraph range"""
    new_doc = Document()
    
    # Copy paragraphs from start to end
    for i in range(start_para, min(end_para + 1, len(original_doc.paragraphs))):
        old_para = original_doc.paragraphs[i]
        new_para = new_doc.add_paragraph()
        
        # Copy paragraph style
        try:
            new_para.style = old_para.style
        except:
            pass
        
        # Copy all runs with formatting
        for run in old_para.runs:
            new_run = new_para.add_run(run.text)
            try:
                new_run.font.name = run.font.name
                new_run.font.size = run.font.size
                new_run.font.bold = run.font.bold
                new_run.font.italic = run.font.italic
                new_run.font.underline = run.font.underline
                if hasattr(run.font, 'superscript'):
                    new_run.font.superscript = run.font.superscript
                if hasattr(run.font, 'subscript'):
                    new_run.font.subscript = run.font.subscript
            except:
                pass
    
    return new_doc

def para_iter(doc):
    """Iterate through all paragraphs including tables"""
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
    """Find the end of a notes section"""
    consecutive_blanks = 0
    for i in range(start, len(paragraphs)):
        p = paragraphs[i]
        text = (p.text or "").strip()
        
        if is_notes_heading(p):
            return i
        
        if not text:
            consecutive_blanks += 1
            if consecutive_blanks >= 2:
                return i
        else:
            consecutive_blanks = 0
    return len(paragraphs)

def parse_references_advanced(paragraphs, start, end):
    """Parse references from notes section"""
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
        
        # Multiple patterns for numbered references
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
            # Continuation of previous reference
            current_text += " " + text
    
    # Save the last reference
    if current_num is not None and current_text.strip():
        refs[current_num] = current_text.strip()
    
    return refs

def expand_number_range(num_str):
    """Expand number ranges like '1-3' to [1, 2, 3]"""
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
    
    # Replace [1], [2,3], [1-3] etc.
    def replace_square(match):
        nums = expand_number_range(match.group(1))
        if (nums and all(1 <= n <= max_ref_num and n in all_refs for n in nums) 
            and not any(n >= 1000 for n in nums)):
            replacements = [format_reference(n, all_refs[n], style) for n in nums]
            return "; ".join(replacements)
        return match.group(0)
    
    modified_text = re.sub(r'\[([0-9,\-–—\s]+)\]', replace_square, modified_text)
    
    # Replace (1), (2,3) if enabled
    if allow_paren:
        modified_text = re.sub(r'\(([0-9,\-–—\s]+)\)', replace_square, modified_text)
    
    # Update paragraph if changed
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
    """Process citations in a single chapter"""
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
    
    # Replace citations in body text
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
            final_doc.add_page_break()
        
        # Copy all paragraphs from chapter
        for para in chapter_doc.paragraphs:
            new_para = final_doc.add_paragraph()
            try:
                new_para.style = para.style
            except:
                pass
            
            for run in para.runs:
                new_run = new_para.add_run(run.text)
                try:
                    new_run.font.name = run.font.name
                    new_run.font.size = run.font.size
                    new_run.font.bold = run.font.bold
                    new_run.font.italic = run.font.italic
                    new_run.font.underline = run.font.underline
                    if hasattr(run.font, 'superscript'):
                        new_run.font.superscript = run.font.superscript
                except:
                    pass
    
    return final_doc

# Main Streamlit app
if uploaded:
    try:
        doc = Document(uploaded)
        
        # Step 1: Find chapters with font size 26 ONLY
        st.subheader("🔍 Step 1: Font Size 26 Detection")
        chapters = find_font26_chapters(doc)
        
        if chapters:
            st.success(f"Found **{len(chapters)}** chapters with font size 26:")
            for i, chapter in enumerate(chapters):
                st.write(f"📖 **{i+1:02d}.** {chapter['text'][:80]}... (Paragraph {chapter['index']})")
        else:
            st.warning("❌ No font size 26 text found. Will process entire document as one chapter.")
        
        # Step 2: Create boundaries
        boundaries = create_chapter_boundaries(chapters, len(doc.paragraphs))
        
        st.subheader("📋 Step 2: Chapter Boundaries")
        for i, (start, end, title) in enumerate(boundaries):
            st.write(f"**{i+1:02d}. {title}** - Paragraphs {start} to {end} ({end-start+1} paragraphs)")
        
        # Step 3: Process
        if st.button("🚀 Process All Chapters"):
            chapter_docs = []
            total_refs = 0
            total_replacements = 0
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            st.subheader("⚙️ Step 3: Processing Results")
            
            for i, (start, end, title) in enumerate(boundaries):
                status_text.text(f"Processing Chapter {i+1}/{len(boundaries)}: {title}")
                
                # Create and process chapter
                chapter_doc = create_chapter_document(doc, start, end)
                refs_found, citations_replaced = process_single_chapter(
                    chapter_doc, fmt, also_replace_parentheses, delete_notes
                )
                
                chapter_docs.append(chapter_doc)
                total_refs += refs_found
                total_replacements += citations_replaced
                
                st.write(f"✅ **{i+1:02d}. {title}**: {refs_found} references, {citations_replaced} citations replaced")
                progress_bar.progress((i + 1) / len(boundaries))
            
            # Step 4: Rejoin
            status_text.text("🔗 Rejoining all processed chapters...")
            final_doc = rejoin_chapters(chapter_docs)
            
            # Save final document
            bio = BytesIO()
            final_doc.save(bio)
            bio.seek(0)
            
            status_text.text("✅ Processing complete!")
            st.success(f"🎉 **Final Results**: {total_refs} total references, {total_replacements} total citations replaced across {len(boundaries)} chapters!")
            
            st.download_button(
                "📥 Download Processed Document",
                data=bio.getvalue(),
                file_name="book_processed_font26_chapters.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
            # Optional: Download individual chapters
            with st.expander("📁 Download Individual Chapters"):
                for i, (chapter_doc, (start, end, title)) in enumerate(zip(chapter_docs, boundaries)):
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
    ### Strict Font Size 26 Rules:
    - ✅ **Only font size 26** text will be detected as chapter headers
    - ❌ **No other font sizes** considered
    - ❌ **No style-based detection**
    - ❌ **No pattern-based detection**
    - 📄 If no font size 26 found, entire document processed as one chapter
    """)
