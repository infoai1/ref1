import streamlit as st
import re

st.set_page_config(page_title="Step 3: Select Chapters", page_icon="✅")
st.title("✅ Step 3: Select & Preview Chapters")

def get_chapter_candidates():
    """Get chapter candidates from previous step"""
    if 'chapter_candidates' in st.session_state:
        return st.session_state['chapter_candidates']
    return None

def create_chapter_boundaries(selected_chapters, total_paragraphs):
    """Create chapter boundaries"""
    if not selected_chapters:
        return [(0, total_paragraphs - 1, "Full_Document")]
    
    # Sort by paragraph index
    selected_chapters.sort(key=lambda x: x['index'])
    
    boundaries = []
    for i, chapter in enumerate(selected_chapters):
        start = chapter['index']
        end = selected_chapters[i + 1]['index'] - 1 if i + 1 < len(selected_chapters) else total_paragraphs - 1
        
        # Clean title
        title = re.sub(r'[^\w\s-]', '', chapter['text'])
        title = re.sub(r'\s+', '_', title)[:40]
        if not title:
            title = f"Chapter_{i+1}"
        
        boundaries.append((start, end, title))
    
    return boundaries

uploaded = st.file_uploader("Upload your DOCX file", type=["docx"])
candidates_data = get_chapter_candidates()

if uploaded and candidates_data:
    from docx import Document
    doc = Document(uploaded)
    total_paragraphs = len(doc.paragraphs)
    
    st.subheader(f"📋 Chapters with Font Size {candidates_data['font_size']}pt")
    
    chapters = candidates_data['chapters']
    
    st.write(f"Found **{len(chapters)}** potential chapters. Select which ones to use:")
    
    # Let user select chapters
    selected_indices = []
    for i, chapter in enumerate(chapters):
        is_selected = st.checkbox(
            f"**Para {chapter['index']}**: {chapter['preview']}",
            key=f"ch_{i}",
            help=f"Full text: {chapter['text']}"
        )
        if is_selected:
            selected_indices.append(i)
    
    if selected_indices:
        selected_chapters = [chapters[i] for i in selected_indices]
        
        # Show chapter boundaries
        st.subheader("📖 Chapter Boundaries Preview")
        
        boundaries = create_chapter_boundaries(selected_chapters, total_paragraphs)
        
        for i, (start, end, title) in enumerate(boundaries):
            st.write(f"**{i+1:02d}. {title}**")
            st.caption(f"Paragraphs {start} to {end} ({end-start+1} paragraphs)")
        
        # Save for processing
        if st.button("✅ Confirm Chapter Selection"):
            st.session_state['final_chapters'] = {
                'selected_chapters': selected_chapters,
                'boundaries': boundaries,
                'font_size': candidates_data['font_size']
            }
            st.success("✅ Chapters confirmed! Proceed to Step 4 for processing.")
            st.info("Run: `streamlit run step4_citation_processing.py`")
    else:
        st.warning("Please select at least one chapter header.")
elif not uploaded:
    st.info("📤 Upload your DOCX file")
else:
    st.warning("⚠️ Please complete Step 2 first (Choose Font Size)")
    st.info("Run: `streamlit run step2_font_selection.py`")

