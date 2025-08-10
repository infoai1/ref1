import streamlit as st
from docx import Document
from docx.oxml.ns import qn

def find_paragraphs_with_font(doc, target_font_size):
    """Enhanced paragraph detection for specific font sizes including 26pt"""
    candidates = []
    
    # Get style fonts (this is where 26pt often hides)
    style_fonts = {}
    for style in doc.styles:
        try:
            if hasattr(style, 'font') and style.font and style.font.size:
                style_fonts[style.name] = style.font.size.pt
        except:
            pass
    
    # Check each paragraph with enhanced detection
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue
        
        has_target_font = False
        
        # Method 1: Check paragraph style font
        try:
            style_name = para.style.name
            if style_fonts.get(style_name) == target_font_size:
                has_target_font = True
        except:
            pass
        
        # Method 2: Check run fonts with multiple approaches
        if not has_target_font:
            for run in para.runs:
                font_size = None
                
                # Direct font.size check
                try:
                    if run.font.size:
                        font_size = run.font.size.pt
                except:
                    pass
                
                # XML-level font size check
                try:
                    if run.element.rPr is not None:
                        sz = run.element.rPr.find(qn('w:sz'))
                        if sz is not None:
                            font_size = float(sz.get(qn('w:val'))) / 2
                        
                        # Also check complex script font size
                        szCs = run.element.rPr.find(qn('w:szCs'))
                        if szCs is not None:
                            font_size = max(font_size or 0, float(szCs.get(qn('w:val'))) / 2)
                except:
                    pass
                
                if font_size == target_font_size:
                    has_target_font = True
                    break
        
        # Method 3: XML text search for specific font size
        if not has_target_font and target_font_size == 26:
            try:
                # Search for 52 (26*2) in the paragraph XML
                para_xml = para._element.xml
                if f'w:val="{int(target_font_size * 2)}"' in para_xml:
                    has_target_font = True
            except:
                pass
        
        if has_target_font:
            candidates.append({
                'index': i,
                'text': text,
                'preview': text[:60] + ('...' if len(text) > 60 else '')
            })
    
    return candidates

def get_font_analysis():
    """Get font analysis from previous step"""
    if 'font_analysis' in st.session_state:
        return st.session_state['font_analysis']
    return None

# Only run UI code if this file is run directly
if __name__ == "__main__":
    st.set_page_config(page_title="Step 2: Enhanced Chapter Font Selection", page_icon="🎯")
    st.title("🎯 Step 2: Enhanced Chapter Font Size Selection")
    
    uploaded = st.file_uploader("Upload your DOCX file", type=["docx"])
    
    if uploaded:
        doc = Document(uploaded)
        analysis = get_font_analysis()
        
        if analysis:
            st.subheader("📊 Available Font Sizes")
            font_sizes = analysis['font_sizes']
            
            # Highlight if 26pt was found
            if 26.0 in font_sizes:
                st.success("🎉 **Font size 26pt is available for selection!**")
            
            # Show available font sizes
            for size, count in sorted(font_sizes.items(), reverse=True):
                color = "🔴" if size >= 24 else "🟡" if size >= 18 else ""
                st.write(f"{color}**{size}pt** - {count} occurrences")
            
            # Let user choose
            st.subheader("🎯 Select Chapter Header Font Size")
            available_sizes = sorted(font_sizes.keys(), reverse=True)
            selected_font = st.selectbox(
                "Choose font size for chapter headers:",
                available_sizes,
                format_func=lambda x: f"{x}pt ({font_sizes[x]} occurrences)" + (" ⭐" if x == 26 else "")
            )
            
            if st.button("🔍 Find Chapters with This Font Size"):
                with st.spinner(f"Searching for paragraphs with {selected_font}pt font..."):
                    chapters = find_paragraphs_with_font(doc, selected_font)
                
                if chapters:
                    st.success(f"Found {len(chapters)} potential chapter headers with {selected_font}pt font!")
                    
                    # Save for next step
                    st.session_state['chapter_candidates'] = {
                        'font_size': selected_font,
                        'chapters': chapters
                    }
                    
                    st.subheader("📋 Preview of Potential Chapters")
                    for i, chapter in enumerate(chapters):
                        st.write(f"**{i+1:02d}.** Para {chapter['index']}: {chapter['preview']}")
                    
                    st.success("✅ Chapters found! Proceed to Step 3.")
                else:
                    st.warning(f"No paragraphs found with font size {selected_font}pt")
                    st.info("💡 Try a different font size or check if the document structure is complex")
        else:
            st.warning("⚠️ Please complete Step 1 first (Enhanced Font Analysis)")
    else:
        st.info("📤 Upload the same DOCX file from Step 1")
