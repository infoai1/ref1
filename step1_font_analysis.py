import streamlit as st
from docx import Document
from docx.oxml.ns import qn
from collections import Counter, defaultdict

st.set_page_config(page_title="Step 1: Font Size Analysis", page_icon="🔍")
st.title("🔍 Step 1: Document Font Size Analysis")

uploaded = st.file_uploader("Upload your DOCX file", type=["docx"])

def detect_all_font_sizes(doc):
    """Detect all font sizes in the document"""
    font_sizes = Counter()
    font_examples = defaultdict(list)
    
    # Check document styles
    style_fonts = {}
    for style in doc.styles:
        try:
            if hasattr(style, 'font') and style.font and style.font.size:
                style_fonts[style.name] = style.font.size.pt
        except:
            pass
    
    # Check each paragraph
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue
        
        # Get style font size
        style_font_size = None
        try:
            style_name = para.style.name
            style_font_size = style_fonts.get(style_name)
        except:
            pass
        
        # Check runs
        para_font_sizes = []
        for run in para.runs:
            font_size = style_font_size  # Default to style
            
            # Check run font size
            try:
                if run.font.size:
                    font_size = run.font.size.pt
            except:
                pass
            
            # Check XML
            try:
                if run.element.rPr is not None:
                    sz = run.element.rPr.find(qn('w:sz'))
                    if sz is not None:
                        font_size = float(sz.get(qn('w:val'))) / 2
            except:
                pass
            
            if font_size:
                para_font_sizes.append(font_size)
        
        # Use largest font in paragraph
        if para_font_sizes:
            max_font = max(para_font_sizes)
            font_sizes[max_font] += 1
            
            # Store examples
            if len(font_examples[max_font]) < 5:
                font_examples[max_font].append({
                    'para_index': i,
                    'text': text[:80] + ('...' if len(text) > 80 else ''),
                    'full_text': text
                })
    
    return font_sizes, font_examples

if uploaded:
    doc = Document(uploaded)
    
    st.subheader("📊 Analyzing Font Sizes...")
    font_sizes, font_examples = detect_all_font_sizes(doc)
    
    if font_sizes:
        st.success(f"Found {len(font_sizes)} different font sizes!")
        
        # Display results
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Total Paragraphs", len(doc.paragraphs))
        with col2:
            st.metric("Font Sizes Found", len(font_sizes))
        
        # Chart
        st.subheader("📈 Font Size Distribution")
        chart_data = dict(font_sizes.most_common())
        st.bar_chart(chart_data)
        
        # Examples
        st.subheader("📝 Font Size Examples")
        for font_size, count in sorted(font_sizes.items(), reverse=True):
            with st.expander(f"Font Size {font_size}pt ({count} occurrences)"):
                examples = font_examples[font_size]
                for ex in examples:
                    st.write(f"• **Para {ex['para_index']}**: {ex['text']}")
        
        # Save results for next step
        st.session_state['font_analysis'] = {
            'font_sizes': dict(font_sizes),
            'font_examples': dict(font_examples),
            'doc_path': uploaded.name
        }
        
        st.success("✅ Analysis complete! Proceed to Step 2.")
        st.info("Run: `streamlit run step2_font_selection.py`")
    else:
        st.error("No font sizes detected!")
else:
    st.info("📤 Upload a DOCX file to analyze font sizes")

