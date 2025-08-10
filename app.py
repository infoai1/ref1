import re
from io import BytesIO
import streamlit as st
from docx import Document
from docx.shared import Pt
from collections import Counter, defaultdict

st.set_page_config(page_title="DOCX: Smart Chapter Detector & Citation Processor", page_icon="📚")
st.title("Smart Chapter Detection & Citation Processing")

st.markdown("""
### Workflow:
1. **Analyze**: Examine entire document structure
2. **Detect**: Find chapter boundaries using multiple criteria  
3. **Split**: Divide into logical chapters
4. **Process**: Handle citations in each chapter
5. **Rejoin**: Combine processed chapters
""")

uploaded = st.file_uploader("Upload the full book .docx", type=["docx"])

def analyze_document_structure(doc):
    """Comprehensive analysis of document structure"""
    analysis = {
        'total_paragraphs': len(doc.paragraphs),
        'font_sizes': Counter(),
        'styles': Counter(),
        'potential_headers': [],
        'page_breaks': [],
        'heading_patterns': [],
        'empty_paragraphs': 0
    }
    
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        
        # Count empty paragraphs
        if not text:
            analysis['empty_paragraphs'] += 1
            continue
            
        # Analyze styles
        style_name = getattr(para.style, 'name', 'Normal')
        analysis['styles'][style_name] += 1
        
        # Analyze font sizes
        font_sizes_in_para = []
        for run in para.runs:
            if run.font.size:
                size_pt = run.font.size.pt
                font_sizes_in_para.append(size_pt)
                analysis['font_sizes'][size_pt] += 1
        
        # Look for potential headers
        if font_sizes_in_para:
            max_font = max(font_sizes_in_para)
            avg_font = sum(font_sizes_in_para) / len(font_sizes_in_para)
            
            # Potential header criteria
            is_potential_header = (
                max_font > 14 or  # Large font
                style_name.lower().startswith('heading') or  # Heading style
                (len(text) < 100 and max_font >= 12) or  # Short text with decent font
                re.match(r'^(chapter|part|section)\s+\d+', text, re.I) or  # Chapter/Part pattern
                re.match(r'^\d+[\.\)]\s+[A-Z]', text) or  # Numbered sections
                (text.isupper() and len(text.split()) <= 10)  # All caps short text
            )
            
            if is_potential_header:
                analysis['potential_headers'].append({
                    'index': i,
                    'text': text[:100],
                    'font_size': max_font,
                    'style': style_name,
                    'criteria': []
                })
        
        # Look for common heading patterns
        patterns = [
            r'^(chapter|ch\.?)\s+\d+',
            r'^(part|section)\s+\d+',
            r'^\d+[\.\)]\s+[A-Z][A-Z\s]+$',
            r'^[A-Z][A-Z\s]{5,50}$',
            r'^(introduction|conclusion|foreword|preface|epilogue)$'
        ]
        
        for pattern in patterns:
            if re.match(pattern, text, re.I):
                analysis['heading_patterns'].append({
                    'index': i,
                    'text': text,
                    'pattern': pattern
                })
    
    return analysis

def smart_chapter_detection(doc, analysis):
    """Smart chapter detection using multiple criteria"""
    potential_chapters = []
    
    # Strategy 1: Use heading styles
    for i, para in enumerate(doc.paragraphs):
        style_name = getattr(para.style, 'name', 'Normal')
        if style_name.lower().startswith('heading 1'):
            potential_chapters.append({
                'index': i,
                'text': para.text.strip(),
                'method': 'heading_style',
                'confidence': 0.9
            })
    
    # Strategy 2: Use font size analysis
    if analysis['font_sizes']:
        # Find the largest font sizes
        sorted_fonts = sorted(analysis['font_sizes'].items(), key=lambda x: x[1], reverse=True)
        large_fonts = [size for size, count in sorted_fonts[:3] if size > 14]
        
        for i, para in enumerate(doc.paragraphs):
            text = para.text.strip()
            if not text or len(text) > 200:  # Skip empty or very long paragraphs
                continue
                
            max_font_in_para = 0
            for run in para.runs:
                if run.font.size:
                    max_font_in_para = max(max_font_in_para, run.font.size.pt)
            
            if max_font_in_para in large_fonts:
                potential_chapters.append({
                    'index': i,
                    'text': text,
                    'method': 'large_font',
                    'confidence': 0.7,
                    'font_size': max_font_in_para
                })
    
    # Strategy 3: Pattern-based detection
    patterns = [
        (r'^(chapter|ch\.?)\s+\d+', 0.95),
        (r'^(part|section)\s+\d+', 0.8),
        (r'^\d+[\.\)]\s+[A-Z][A-Z\s]+$', 0.7),
        (r'^[A-Z][A-Z\s]{10,50}$', 0.6),
        (r'^(introduction|conclusion|foreword|preface|epilogue|bibliography|references|notes)$', 0.8)
    ]
    
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue
            
        for pattern, confidence in patterns:
            if re.match(pattern, text, re.I):
                potential_chapters.append({
                    'index': i,
                    'text': text,
                    'method': 'pattern',
                    'confidence': confidence,
                    'pattern': pattern
                })
    
    # Strategy 4: All caps detection
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if (text and text.isupper() and 5 <= len(text.split()) <= 15 
            and not any(char.isdigit() for char in text)):
            potential_chapters.append({
                'index': i,
                'text': text,
                'method': 'all_caps',
                'confidence': 0.6
            })
    
    # Remove duplicates and sort by index
    seen_indices = set()
    unique_chapters = []
    for chapter in potential_chapters:
        if chapter['index'] not in seen_indices:
            seen_indices.add(chapter['index'])
            unique_chapters.append(chapter)
    
    unique_chapters.sort(key=lambda x: x['index'])
    
    # Merge nearby detections (within 3 paragraphs)
    merged_chapters = []
    for chapter in unique_chapters:
        if not merged_chapters or chapter['index'] - merged_chapters[-1]['index'] > 3:
            merged_chapters.append(chapter)
        else:
            # Merge with previous if higher confidence
            if chapter['confidence'] > merged_chapters[-1]['confidence']:
                merged_chapters[-1] = chapter
    
    return merged_chapters

def create_chapter_boundaries(chapters, total_paragraphs):
    """Create chapter boundaries for splitting"""
    if not chapters:
        return [(0, total_paragraphs - 1, "Full_Document")]
    
    boundaries = []
    for i, chapter in enumerate(chapters):
        start = chapter['index']
        end = chapters[i + 1]['index'] - 1 if i + 1 < len(chapters) else total_paragraphs - 1
        
        # Clean title for filename
        title = re.sub(r'[^\w\s-]', '', chapter['text'])
        title = re.sub(r'\s+', '_', title)[:50]
        if not title:
            title = f"Chapter_{i+1}"
        
        boundaries.append((start, end, title, chapter))
    
    return boundaries

def para_iter(doc):
    """Iterate through all paragraphs including tables"""
    for p in doc.paragraphs:
        yield p
    for t in doc.tables:
        for r in t.rows:
            for c in r.cells:
                for p in c.paragraphs:
                    yield p

def create_chapter_document(original_doc, start_para, end_para):
    """Create a new document with specified paragraph range"""
    new_doc = Document()
    
    # Copy core document properties
    new_doc.core_properties.title = f"Chapter {start_para}-{end_para}"
    
    # Copy paragraphs
    for i in range(start_para, min(end_para + 1, len(original_doc.paragraphs))):
        old_para = original_doc.paragraphs[i]
        new_para = new_doc.add_paragraph()
        
        # Copy paragraph style
        try:
            new_para.style = old_para.style
        except:
            pass
        
        # Copy runs with formatting
        for run in old_para.runs:
            new_run = new_para.add_run(run.text)
            # Copy font properties
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

# Citation processing functions (same as before)
HEADING_RE = re.compile(r"^\s*(notes?|references?|endnotes?|sources?|bibliography|citations?)\s*:?\s*$", re.I)

def is_notes_heading(p):
    text = (p.text or "").strip()
    return bool(HEADING_RE.match(text)) if text else False

def find_notes_sections(paragraphs):
    sections = []
    for i, p in enumerate(paragraphs):
        if is_notes_heading(p):
            start = i + 1
            end = find_section_end(paragraphs, start)
            sections.append((i, start, end))
    return sections

def find_section_end(paragraphs, start):
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

def process_single_chapter(doc, fmt="[1. Reference text]", allow_paren=False, delete_notes_sections=False):
    """Process citations in a single chapter"""
    paragraphs = list(para_iter(doc))
    notes_sections = find_notes_sections(paragraphs)
    
    if not notes_sections:
        return 0, 0
    
    # Parse references
    all_refs = {}
    for section_head, section_start, section_end in notes_sections:
        refs = parse_references_advanced(paragraphs, section_start, section_end)
        all_refs.update(refs)
    
    if not all_refs:
        return 0, 0
    
    # Replace citations (simplified version)
    total_replacements = 0
    notes_ranges = set()
    for section_head, section_start, section_end in notes_sections:
        for i in range(section_head, section_end):
            notes_ranges.add(i)
    
    # Simple citation replacement
    for i, paragraph in enumerate(paragraphs):
        if i not in notes_ranges and paragraph.text:
            original_text = paragraph.text
            modified_text = original_text
            
            # Replace [1], [1], etc.
            def replace_citation(match):
                try:
                    num = int(match.group(1))
                    if num in all_refs:
                        if fmt.startswith("—"):
                            return f" — {num}. {all_refs[num]}"
                        elif fmt.startswith("["):
                            return f" [{num}. {all_refs[num]}]"
                        else:
                            return f" ({all_refs[num]})"
                except:
                    pass
                return match.group(0)
            
            modified_text = re.sub(r'\[(\d+)\]', replace_citation, modified_text)
            
            if modified_text != original_text:
                # Update paragraph text
                for run in paragraph.runs:
                    run.text = ""
                if paragraph.runs:
                    paragraph.runs[0].text = modified_text
                else:
                    paragraph.add_run(modified_text)
                total_replacements += 1
    
    return len(all_refs), total_replacements

def rejoin_chapters(chapter_docs):
    """Rejoin processed chapters"""
    final_doc = Document()
    
    for i, chapter_doc in enumerate(chapter_docs):
        if i > 0:
            final_doc.add_page_break()
        
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
                except:
                    pass
    
    return final_doc

# Main app
if uploaded:
    try:
        doc = Document(uploaded)
        
        # Step 1: Analyze document structure
        st.subheader("📊 Step 1: Document Analysis")
        with st.spinner("Analyzing document structure..."):
            analysis = analyze_document_structure(doc)
        
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Total Paragraphs", analysis['total_paragraphs'])
            st.metric("Empty Paragraphs", analysis['empty_paragraphs'])
        with col2:
            st.metric("Unique Font Sizes", len(analysis['font_sizes']))
            st.metric("Unique Styles", len(analysis['styles']))
        
        # Show font size distribution
        st.write("**Font Size Distribution:**")
        font_data = dict(analysis['font_sizes'].most_common(10))
        st.bar_chart(font_data)
        
        # Show style distribution  
        st.write("**Style Distribution:**")
        style_data = dict(analysis['styles'].most_common(10))
        st.bar_chart(style_data)
        
        # Step 2: Smart chapter detection
        st.subheader("🔍 Step 2: Smart Chapter Detection")
        chapters = smart_chapter_detection(doc, analysis)
        
        st.write(f"**Detected {len(chapters)} potential chapters:**")
        for i, chapter in enumerate(chapters):
            confidence_color = "🟢" if chapter['confidence'] > 0.8 else "🟡" if chapter['confidence'] > 0.6 else "🔴"
            st.write(f"{confidence_color} **Chapter {i+1}** (Para {chapter['index']}): {chapter['text'][:80]}...")
            st.caption(f"Method: {chapter['method']}, Confidence: {chapter['confidence']:.1%}")
        
        # Step 3: Chapter boundaries
        boundaries = create_chapter_boundaries(chapters, len(doc.paragraphs))
        
        st.write(f"**Chapter Boundaries:**")
        for i, (start, end, title, chapter_info) in enumerate(boundaries):
            st.write(f"📖 **{i+1:02d}. {title}** - Paragraphs {start} to {end} ({end-start+1} paragraphs)")
        
        # Processing options
        st.subheader("⚙️ Step 3: Processing Options")
        fmt = st.selectbox("Citation format", ["[1. Reference text]", "— 1. Reference text", "(Reference text)"])
        delete_notes = st.checkbox("Delete Notes sections after processing")
        
        if st.button("🚀 Process All Chapters"):
            chapter_docs = []
            total_refs = 0
            total_replacements = 0
            
            progress_bar = st.progress(0)
            status = st.empty()
            
            for i, (start, end, title, chapter_info) in enumerate(boundaries):
                status.text(f"Processing chapter {i+1}/{len(boundaries)}: {title}")
                
                # Create and process chapter
                chapter_doc = create_chapter_document(doc, start, end)
                refs_found, citations_replaced = process_single_chapter(chapter_doc, fmt, False, delete_notes)
                
                chapter_docs.append(chapter_doc)
                total_refs += refs_found
                total_replacements += citations_replaced
                
                st.write(f"✅ **{i+1:02d}. {title}**: {refs_found} refs, {citations_replaced} citations replaced")
                progress_bar.progress((i + 1) / len(boundaries))
            
            # Rejoin chapters
            status.text("Rejoining processed chapters...")
            final_doc = rejoin_chapters(chapter_docs)
            
            # Save result
            bio = BytesIO()
            final_doc.save(bio)
            bio.seek(0)
            
            st.success(f"🎉 **Processing Complete!** {total_refs} references found, {total_replacements} citations replaced across {len(boundaries)} chapters!")
            
            st.download_button(
                "📥 Download Processed Document",
                data=bio.getvalue(),
                file_name="book_processed_complete.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

    except Exception as e:
        st.error(f"Error: {str(e)}")
        st.exception(e)
else:
    st.info("📤 Upload a DOCX file to begin comprehensive analysis and processing")
