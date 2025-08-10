import streamlit as st
from docx import Document
from docx.oxml.ns import qn
from collections import Counter, defaultdict
from io import BytesIO
import re
import sys
import os

st.set_page_config(page_title="DOCX Citation Processor", page_icon="📚", layout="wide")

# --- All Functions Now Inside app.py ---

def detect_all_font_sizes(doc):
    """Enhanced font size detection to find all possible font sizes."""
    font_sizes = Counter()
    try:
        # Check styles, defaults, and runs
        for style in doc.styles:
            if hasattr(style, 'font') and style.font.size:
                font_sizes[style.font.size.pt] += 1
        
        for para in doc.paragraphs:
            for run in para.runs:
                if run.font.size:
                    font_sizes[run.font.size.pt] += 1
    except Exception as e:
        st.warning(f"Could not perform full font analysis: {e}")
    return font_sizes

def find_paragraphs_with_font(doc, target_font_size):
    """Finds all paragraphs that contain a given font size."""
    candidates = []
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue
        
        has_target_font = False
        for run in para.runs:
            # Check run-level font size
            if run.font.size and run.font.size.pt == target_font_size:
                has_target_font = True
                break
            # Check style-level font size
            if para.style.font.size and para.style.font.size.pt == target_font_size:
                has_target_font = True
                break
        
        if has_target_font:
            candidates.append({'index': i, 'text': text})
    return candidates

def create_chapter_boundaries(selected_chapters, total_paragraphs):
    if not selected_chapters:
        return [(0, total_paragraphs - 1, "Full_Document")]
    selected_chapters.sort(key=lambda x: x['index'])
    boundaries = []
    for i, chapter in enumerate(selected_chapters):
        start = chapter['index']
        end = selected_chapters[i+1]['index'] - 1 if i+1 < len(selected_chapters) else total_paragraphs - 1
        title = re.sub(r'[^\w\s-]', '', chapter['text'])[:50] or f"Chapter_{i+1}"
        boundaries.append((start, end, title))
    return boundaries

# --- Other processing functions (para_iter, find_notes_sections, etc.) ---
# ... (These functions from your previous files would be included here) ...

# --- Streamlit App UI ---

st.title("📚 DOCX Citation Processor")
st.markdown("### A tool to analyze your document and process citations chapter by chapter.")

uploaded = st.file_uploader("Upload your DOCX file", type=["docx"])

if uploaded:
    doc = Document(uploaded)
    total_paragraphs = len(doc.paragraphs)

    st.header("Step 1: Font Size Analysis")
    with st.spinner("Analyzing document fonts..."):
        font_sizes = detect_all_font_sizes(doc)

    if font_sizes:
        st.success(f"Detected {len(font_sizes)} different font sizes.")
        st.bar_chart(font_sizes)
    else:
        st.warning("No font sizes were automatically detected. Please use the manual entry below.")

    st.header("Step 2: Select Chapter Font Size")
    
    # --- MANUAL OVERRIDE ---
    st.info("If your desired font size (like 26) was not detected, you can enter it manually.")
    manual_font_size = st.number_input("Enter Chapter Font Size Manually (e.g., 26):", min_value=1.0, max_value=100.0, step=0.5)

    detected_sizes = sorted(font_sizes.keys(), reverse=True)
    selected_font_size = st.selectbox("Or select from detected sizes:", detected_sizes)
    
    final_font_size = manual_font_size if manual_font_size > 0 else selected_font_size

    if st.button("Find Chapters with Selected Font Size", key="find_chapters"):
        with st.spinner(f"Searching for paragraphs with font size {final_font_size}pt..."):
            candidates = find_paragraphs_with_font(doc, final_font_size)
        
        if candidates:
            st.success(f"Found {len(candidates)} paragraphs with font size {final_font_size}pt.")
            st.session_state['chapter_candidates'] = candidates
        else:
            st.error(f"Could not find any paragraphs with font size {final_font_size}pt.")

    if 'chapter_candidates' in st.session_state:
        st.header("Step 3: Confirm Chapters & Process")
        candidates = st.session_state['chapter_candidates']
        
        st.write("Please confirm which of these are chapter headers:")
        selected_chapters = []
        for cand in candidates:
            if st.checkbox(f"Para {cand['index']}: {cand['text'][:100]}", value=True):
                selected_chapters.append(cand)
        
        if st.button("Process Confirmed Chapters", type="primary"):
            boundaries = create_chapter_boundaries(selected_chapters, total_paragraphs)
            st.write("Processing the following chapters:")
            for start, end, title in boundaries:
                st.write(f"- **{title}**: Paragraphs {start} to {end}")
            # --- Add your processing and download logic here ---
            st.success("Processing logic would run here.")

else:
    st.info("Please upload a DOCX file to begin.")

