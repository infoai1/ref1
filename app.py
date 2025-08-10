import streamlit as st
from docx import Document
from io import BytesIO
import sys
import os

st.set_page_config(page_title="DOCX Citation Processor", page_icon="📚", layout="wide")

# Add current directory to path for imports
sys.path.append(os.path.dirname(__file__))

# Import functions from other modules
try:
    from step1_font_analysis import detect_all_font_sizes
    from step2_font_selection import find_paragraphs_with_font
    from step3_chapter_selection import create_chapter_boundaries
    from step4_citation_processing import (
        create_chapter_document,
        process_chapter_citations,
        para_iter,
        find_notes_sections
    )
    from step5_rejoin_chapters import rejoin_chapters_with_formatting
except ImportError as e:
    st.error(f"Error importing modules: {e}")
    st.info("Make sure all step files are in the same directory as app.py")
    st.stop()

# Sidebar for navigation
st.sidebar.title("📚 DOCX Citation Processor")
mode = st.sidebar.radio("Choose Mode:", ["🚀 Auto Process", "📋 Step by Step"])

st.title("📚 DOCX Citation Processor")
st.markdown("### Transform numbered citations to full inline references")

uploaded = st.file_uploader("Upload your DOCX file", type=["docx"])

if uploaded:
    doc = Document(uploaded)
    
    if mode == "🚀 Auto Process":
        st.header("🚀 Automated Processing")
        
        col1, col2 = st.columns(2)
        with col1:
            citation_format = st.selectbox(
                "Citation Format:",
                ["[1. Reference text]", "— 1. Reference text", "(Reference text)"]
            )
        with col2:
            delete_notes = st.checkbox("Delete Notes sections after processing")
        
        # 🆕 NEW: Manual font size option for auto mode
        st.subheader("📏 Chapter Detection Settings")
        use_manual_font = st.checkbox("Use manual font size for chapters (if auto-detection fails)")
        manual_font_size = 26.0
        
        if use_manual_font:
            manual_font_size = st.number_input("Chapter font size:", min_value=1.0, max_value=100.0, value=26.0, step=0.5)
        
        if st.button("🚀 Process Document", type="primary"):
            progress_bar = st.progress(0)
            status = st.empty()
            
            # Step 1: Analyze fonts
            status.text("📊 Step 1: Analyzing font sizes...")
            progress_bar.progress(0.2)
            
            font_sizes, font_examples = detect_all_font_sizes(doc)
            
            if not font_sizes:
                st.error("No font sizes detected in document!")
                st.stop()
            
            # Auto-select font size
            if use_manual_font:
                selected_font = manual_font_size
                st.info(f"🎯 Using manual font size: **{selected_font}pt** for chapter detection")
            else:
                selected_font = max(font_sizes.keys())
                st.success(f"✅ Auto-selected font size: **{selected_font}pt** for chapter detection")
            
            # Step 2: Find chapters
            status.text("🔍 Step 2: Finding chapters...")
            progress_bar.progress(0.4)
            
            chapters = find_paragraphs_with_font(doc, selected_font)
            
            if not chapters:
                st.warning(f"No chapters found with font size {selected_font}pt. Processing as single document.")
                chapters = [{'index': 0, 'text': 'Full Document', 'preview': 'Full Document'}]
            
            st.success(f"✅ Found **{len(chapters)}** chapters")
            
            # Step 3: Create boundaries
            status.text("📋 Step 3: Creating chapter boundaries...")
            progress_bar.progress(0.6)
            
            boundaries = create_chapter_boundaries(chapters, len(doc.paragraphs))
            
            # Step 4: Process each chapter
            status.text("⚙️ Step 4: Processing citations...")
            progress_bar.progress(0.8)
            
            chapter_docs = []
            total_refs = 0
            total_replacements = 0
            
            chapter_progress = st.progress(0)
            
            for i, (start, end, title) in enumerate(boundaries):
                chapter_progress.progress((i + 1) / len(boundaries))
                
                # Create chapter document
                chapter_doc = create_chapter_document(doc, start, end)
                
                # Process citations
                refs_found, citations_replaced = process_chapter_citations(
                    chapter_doc, citation_format, delete_notes
                )
                
                chapter_docs.append(chapter_doc)
                total_refs += refs_found
                total_replacements += citations_replaced
                
                st.write(f"✅ **{title}**: {refs_found} references, {citations_replaced} citations replaced")
            
            # Step 5: Rejoin chapters
            status.text("🔗 Step 5: Rejoining chapters...")
            progress_bar.progress(1.0)
            
            final_doc = rejoin_chapters_with_formatting(chapter_docs)
            
            # Create download
            bio = BytesIO()
            final_doc.save(bio)
            bio.seek(0)
            
            status.text("✅ Processing complete!")
            
            # Results
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Chapters Processed", len(boundaries))
            with col2:
                st.metric("References Found", total_refs)
            with col3:
                st.metric("Citations Replaced", total_replacements)
            
            st.success("🎉 Document processing complete!")
            
            # Download button
            st.download_button(
                "📥 Download Processed Document",
                data=bio.getvalue(),
                file_name="document_processed.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    
    else:  # Step by Step Mode
        st.header("📋 Step-by-Step Processing")
        
        # Initialize session state
        if 'step_data' not in st.session_state:
            st.session_state.step_data = {}
        
        # Step 1: Font Analysis
        with st.expander("📊 Step 1: Font Analysis", expanded=True):
            if st.button("🔍 Analyze Font Sizes"):
                font_sizes, font_examples = detect_all_font_sizes(doc)
                st.session_state.step_data['font_analysis'] = {
                    'font_sizes': dict(font_sizes),
                    'font_examples': dict(font_examples)
                }
                
                if font_sizes:
                    st.success(f"Found {len(font_sizes)} different font sizes!")
                    
                    # Display chart
                    chart_data = dict(font_sizes.most_common())
                    st.bar_chart(chart_data)
                    
                    # Show examples
                    for font_size, count in sorted(font_sizes.items(), reverse=True):
                        st.write(f"**{font_size}pt** - {count} occurrences")
                else:
                    st.error("No font sizes detected!")
            
            # 🆕 NEW: Manual font size addition in step-by-step mode
            if 'font_analysis' in st.session_state.step_data:
                st.subheader("➕ Add Missing Font Sizes")
                st.info("If your chapter font size (like 26pt) wasn't detected, add it manually:")
                
                manual_font = st.number_input("Enter font size:", min_value=1.0, max_value=100.0, value=26.0, step=0.5)
                
                if st.button("Add This Font Size"):
                    font_sizes = st.session_state.step_data['font_analysis']['font_sizes']
                    if manual_font not in font_sizes:
                        font_sizes[manual_font] = 1
                        st.session_state.step_data['font_analysis']['font_sizes'] = font_sizes
                        st.success(f"✅ Added {manual_font}pt to available font sizes!")
                        st.experimental_rerun()
                    else:
                        st.info(f"Font size {manual_font}pt already exists!")
        
        # Step 2: Font Selection
        if 'font_analysis' in st.session_state.step_data:
            with st.expander("🎯 Step 2: Select Chapter Font Size", expanded=True):
                font_sizes = st.session_state.step_data['font_analysis']['font_sizes']
                available_sizes = sorted(font_sizes.keys(), reverse=True)
                
                selected_font = st.selectbox(
                    "Choose font size for chapter headers:",
                    available_sizes,
                    format_func=lambda x: f"{x}pt ({font_sizes[x]} occurrences)"
                )
                
                if st.button("🔍 Find Chapters"):
                    chapters = find_paragraphs_with_font(doc, selected_font)
                    st.session_state.step_data['chapters'] = {
                        'font_size': selected_font,
                        'candidates': chapters
                    }
                    
                    if chapters:
                        st.success(f"Found {len(chapters)} potential chapters!")
                        for i, chapter in enumerate(chapters):
                            st.write(f"**{i+1}.** Para {chapter['index']}: {chapter['preview']}")
                    else:
                        st.warning(f"No chapters found with font size {selected_font}pt")
        
        # Step 3: Chapter Selection
        if 'chapters' in st.session_state.step_data:
            with st.expander("✅ Step 3: Select Specific Chapters", expanded=True):
                chapters = st.session_state.step_data['chapters']['candidates']
                
                st.write("Select which paragraphs should be chapter headers:")
                selected_indices = []
                
                for i, chapter in enumerate(chapters):
                    if st.checkbox(f"Para {chapter['index']}: {chapter['preview']}", key=f"ch_{i}"):
                        selected_indices.append(i)
                
                if selected_indices and st.button("✅ Confirm Selection"):
                    selected_chapters = [chapters[i] for i in selected_indices]
                    boundaries = create_chapter_boundaries(selected_chapters, len(doc.paragraphs))
                    
                    st.session_state.step_data['boundaries'] = boundaries
                    
                    st.success("Chapter selection confirmed!")
                    for i, (start, end, title) in enumerate(boundaries):
                        st.write(f"**{i+1}. {title}** - Paragraphs {start} to {end}")
        
        # Step 4: Processing
        if 'boundaries' in st.session_state.step_data:
            with st.expander("⚙️ Step 4: Process Citations", expanded=True):
                col1, col2 = st.columns(2)
                with col1:
                    fmt = st.selectbox("Citation format:", ["[1. Reference text]", "— 1. Reference text", "(Reference text)"])
                with col2:
                    delete_notes = st.checkbox("Delete Notes sections")
                
                if st.button("🚀 Process Citations"):
                    boundaries = st.session_state.step_data['boundaries']
                    chapter_docs = []
                    total_refs = 0
                    total_replacements = 0
                    
                    for start, end, title in boundaries:
                        chapter_doc = create_chapter_document(doc, start, end)
                        refs, citations = process_chapter_citations(chapter_doc, fmt, delete_notes)
                        chapter_docs.append(chapter_doc)
                        total_refs += refs
                        total_replacements += citations
                        st.write(f"✅ {title}: {refs} refs, {citations} citations")
                    
                    st.session_state.step_data['processed'] = {
                        'chapter_docs': chapter_docs,
                        'stats': {'refs': total_refs, 'replacements': total_replacements}
                    }
                    
                    st.success(f"Processing complete! {total_refs} refs, {total_replacements} citations replaced")
        
        # Step 5: Download
        if 'processed' in st.session_state.step_data:
            with st.expander("📥 Step 5: Download Results", expanded=True):
                if st.button("🔗 Rejoin & Download"):
                    chapter_docs = st.session_state.step_data['processed']['chapter_docs']
                    final_doc = rejoin_chapters_with_formatting(chapter_docs)
                    
                    bio = BytesIO()
                    final_doc.save(bio)
                    bio.seek(0)
                    
                    st.download_button(
                        "📥 Download Final Document",
                        data=bio.getvalue(),
                        file_name="document_processed_stepwise.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

else:
    st.info("📤 Upload a DOCX file to begin processing")
    
    # Instructions
    st.markdown("""
    ### 📋 How This Works:
    
    **🚀 Auto Process Mode:**
    - Automatically detects largest font size as chapter headers
    - NEW: Option to manually set chapter font size (e.g., 26pt)
    - Processes the entire document in one go
    - Best for quick processing
    
    **📋 Step by Step Mode:**
    - Manual control over each step
    - Choose specific font sizes and chapters
    - NEW: Add missing font sizes manually in Step 1
    - Best for precise control
    
    ### ✨ Features:
    - ✅ Complete font detection (styles + runs + XML)
    - ✅ Manual font size override for missing sizes
    - ✅ Interactive chapter selection
    - ✅ Multiple citation formats
    - ✅ Original formatting preservation
    - ✅ Chapter-wise processing for accuracy
    """)
