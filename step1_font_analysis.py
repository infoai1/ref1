import streamlit as st
from docx import Document
from docx.oxml.ns import qn
from collections import Counter, defaultdict

def detect_all_font_sizes(doc):
    """Enhanced font size detection - finds ALL font sizes including hidden ones like 26pt"""
    font_sizes = Counter()
    font_examples = defaultdict(list)
    
    # Method 1: Check document styles (where 26pt often hides)
    style_fonts = {}
    for style in doc.styles:
        try:
            if hasattr(style, 'font') and style.font and style.font.size:
                size_pt = style.font.size.pt
                style_fonts[style.name] = size_pt
                font_sizes[size_pt] += 1
        except:
            pass
    
    # Method 2: Check document defaults and themes
    try:
        # Document defaults
        doc_defaults = doc.element.xpath('//w:docDefaults')
        for d in doc_defaults:
            sz = d.xpath('.//w:sz')
            if sz:
                val = sz[0].get(qn('w:val'))
                if val:
                    font_sizes[float(val)/2] += 1
    except:
        pass
    
    # Method 3: Deep paragraph and run analysis
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue
        
        # Get style font size for this paragraph
        style_font_size = None
        try:
            style_name = para.style.name
            style_font_size = style_fonts.get(style_name)
        except:
            pass
        
        # Check each run with multiple methods
        para_font_sizes = []
        for run in para.runs:
            font_size = style_font_size  # Start with style font
            
            # Direct font.size check
            try:
                if run.font.size:
                    font_size = run.font.size.pt
            except:
                pass
            
            # XML-level font size check
            try:
                if run.element.rPr is not None:
                    # Standard font size
                    sz = run.element.rPr.find(qn('w:sz'))
                    if sz is not None:
                        font_size = float(sz.get(qn('w:val'))) / 2
                    
                    # Complex script font size
                    szCs = run.element.rPr.find(qn('w:szCs'))
                    if szCs is not None:
                        font_size = max(font_size or 0, float(szCs.get(qn('w:val'))) / 2)
            except:
                pass
            
            if font_size:
                para_font_sizes.append(font_size)
        
        # Use the largest font size found
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
    
    # Method 4: Global XML search for font sizes (catches hidden 26pt)
    try:
        from lxml import etree
        root = doc.element
        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        sz_elements = root.xpath('//w:sz', namespaces=ns)
        for sz in sz_elements:
            val = sz.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
            if val:
                font_sizes[float(val)/2] += 1
    except ImportError:
        st.warning("lxml not installed, skipping global XML search")
    except:
        pass
    
    return font_sizes, font_examples

# Only run UI code if this file is run directly
if __name__ == "__main__":
    st.set_page_config(page_title="Step 1: Font Size Analysis", page_icon="🔍")
    st.title("🔍 Step 1: Document Font Size Analysis")
    
    uploaded = st.file_uploader("Upload your DOCX file", type=["docx"])
    
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
