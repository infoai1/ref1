import streamlit as st
from docx import Document
from docx.oxml.ns import qn

st.set_page_config(page_title="Step 2: Choose Chapter Font", page_icon="🎯")
st.title("🎯 Step 2: Choose Chapter Font Size")

uploaded = st.file_uploader("Upload your DOCX file", type=["docx"])

def get_font_analysis():
    """Get font analysis from previous step"""
    if 'font_analysis' in st.session_state:
        return st.session_state['font_analysis']
    return None

def find_paragraphs_with_font(doc, target_font_size):
    """Find all paragraphs using the target font size"""
    candidates = []
    
    # Get style fonts
    style_fonts = {}
    for style in doc.styles:
        try:
            if hasattr(style, 'font') and style.font and style.font.size:
                style_fonts[style.name] = style.font.size.pt
        except:
            pass
    
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue
        
        has_target_font = False
        
        # Check style font
        try:
            style_name = para.style.name
            if style_fonts.get(style_name) == target_font_size:
                has_target_font = True
        except:
            pass
        
        # Check run fonts
        if not has_target_font:
            for run in para.runs:
                font_size = None
                try:
                    if run.font.size:
                        font_size = run.font.size.pt
                except:
                    pass
                
                try:
                    if run.element.rPr is not None:
                        sz = run.element.rPr.find(qn('w:sz'))
                        if sz is not None:
                            font_size = float(sz.get(qn('w:val'))) / 2
                except:
                    pass
                
                if font_size == target_font_size:
                    has_target_font = True
                    break
        
        if has_target_font:
            candidates.append({
                'index': i,
                'text': text,
                'preview': text[:60] + ('...' if len(text) > 60 else '')
            })
    
    return candidates

if uploaded:
    doc = Document(uploaded)
    analysis = get_font_analysis()
    
    if analysis:
        st.subheader("📊 Available Font Sizes")
        font_sizes = analysis['font_sizes']
        
        # Show available font sizes
        for size, count in sorted(font_sizes.items(), reverse=True):
            st.write(f"**{size}pt** - {count} occurrences")
        
        # Let user choose
        st.subheader("🎯 Select Chapter Header Font Size")
        available_sizes = sorted(font_sizes.keys(), reverse=True)
        selected_font = st.selectbox(
            "Choose font size for chapter headers:",
            available_sizes,
            format_func=lambda x: f"{x}pt ({font_sizes[x]} occurrences)"
        )
        
        if st.button("🔍 Find Chapters with This Font Size"):
            chapters = find_paragraphs_with_font(doc, selected_font)
            
            if chapters:
                st.success(f"Found {len(chapters)} potential chapter headers!")
                
                # Save for next step
                st.session_state['chapter_candidates'] = {
                    'font_size': selected_font,
                    'chapters': chapters
                }
                
                st.subheader("📋 Preview of Potential Chapters")
                for i, chapter in enumerate(chapters):
                    st.write(f"**{i+1:02d}.** Para {chapter['index']}: {chapter['preview']}")
                
                st.success("✅ Chapters found! Proceed to Step 3.")
                st.info("Run: `streamlit run step3_chapter_selection.py`")
            else:
                st.warning(f"No paragraphs found with font size {selected_font}pt")
    else:
        st.warning("⚠️ Please complete Step 1 first (Font Analysis)")
        st.info("Run: `streamlit run step1_font_analysis.py`")
else:
    st.info("📤 Upload the same DOCX file from Step 1")

