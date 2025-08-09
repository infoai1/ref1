import io
import os
import zipfile
import tempfile
from pathlib import Path

import streamlit as st

# pypandoc is a thin wrapper around pandoc. We'll auto-download pandoc if missing.
try:
    import pypandoc
    _ = pypandoc.get_pandoc_version()
except Exception:
    import pypandoc
    with st.spinner("Downloading Pandoc (first run takes a few minutes)..."):
        pypandoc.download_pandoc()

DEFAULT_CSL = "https://www.zotero.org/styles/vancouver-superscript"

st.set_page_config(page_title="MD → DOCX (Numbered Superscript Citations)", page_icon="📄")

st.title("📄 Markdown → DOCX with Superscript Citations & References")
st.write(
    "Upload **Markdown file(s)** containing citations like `[@smith2020]` and a **.bib** file. "
    "Choose output: one combined DOCX or one DOCX per chapter (per uploaded file)."
)

# --- Inputs ---
md_files = st.file_uploader(
    "Upload Markdown file(s) — one per chapter (order matters)",
    type=["md", "markdown"],
    accept_multiple_files=True,
)

bib_file = st.file_uploader(
    "Upload bibliography (.bib or CSL-JSON .json)",
    type=["bib", "json"],
    accept_multiple_files=False,
)

st.markdown("**Citation style** (CSL): default is Vancouver superscript.)")
use_default_csl = st.checkbox("Use default Vancouver superscript style", value=True)
custom_csl = st.text_input(
    "…or paste a CSL URL (e.g., https://www.zotero.org/styles/nature)", value=""
)

output_mode = st.radio(
    "Output mode",
    ["One combined DOCX (single References at end)", "One DOCX per chapter (zipped)"]
)

book_title = st.text_input("Optional: Title for the combined DOCX", value="My Book")

if st.button("Build output"):
    if not md_files:
        st.error("Please upload at least one Markdown file.")
        st.stop()
    if not bib_file:
        st.error("Please upload a .bib or .json bibliography file.")
        st.stop()

    csl = DEFAULT_CSL if use_default_csl or not custom_csl.strip() else custom_csl.strip()

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)

        # Save bibliography
        bib_path = tmp / bib_file.name
        bib_path.write_bytes(bib_file.read())

        # Save markdown files in order
        saved_md_paths = []
        for idx, f in enumerate(md_files, start=1):
            p = tmp / f"chapter_{idx:02d}_" + Path(f.name).name
            p.write_bytes(f.read())
            saved_md_paths.append(p)

        extra_args = [
            "--citeproc",
            f"--csl={csl}",
            f"--bibliography={str(bib_path)}",
            "--metadata", "link-citations=true",
            "--metadata", "reference-section-title=References",
        ]

        # We use Pandoc's markdown reader with citations enabled
        from_format = "markdown+yaml_metadata_block+citations"

        if output_mode.startswith("One combined"):
            # Build a single temporary combined Markdown with a small YAML header
            combined_md = tmp / "_combined.md"
            parts = ["---", f"title: {book_title}", "link-citations: true", "---", "\n"]

            for i, p in enumerate(saved_md_paths, start=1):
                # Chapter separator (optional). Pandoc doesn't force page-breaks for DOCX here.
                if i > 1:
                    parts.append("\n\n***\n\n")  # visual separator only
                parts.append(p.read_text(encoding="utf-8", errors="ignore"))

            combined_md.write_text("\n".join(parts), encoding="utf-8")

            out_docx = tmp / "output.docx"
            try:
                pypandoc.convert_file(
                    str(combined_md),
                    to="docx",
                    format=from_format,
                    outputfile=str(out_docx),
                    extra_args=extra_args,
                )
                st.success("Done! Download your DOCX below.")
                st.download_button(
                    "⬇️ Download DOCX",
                    data=out_docx.read_bytes(),
                    file_name="book_with_citations.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
            except Exception as e:
                st.exception(e)
        else:
            # One DOCX per chapter → ZIP
            buffer = io.BytesIO()
            with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for i, p in enumerate(saved_md_paths, start=1):
                    out_path = tmp / f"chapter_{i:02d}.docx"
                    try:
                        pypandoc.convert_file(
                            str(p),
                            to="docx",
                            format=from_format,
                            outputfile=str(out_path),
                            extra_args=extra_args,
                        )
                        zf.write(out_path, arcname=out_path.name)
                    except Exception as e:
                        st.error(f"Conversion failed for {p.name}: {e}")
                        st.stop()

            buffer.seek(0)
            st.success("Done! Download your ZIP of per-chapter DOCX files below.")
            st.download_button(
                "⬇️ Download ZIP",
                data=buffer.getvalue(),
                file_name="chapters_with_citations.zip",
                mime="application/zip",
            )

st.markdown(
    """
---
### How to cite in Markdown
Use citation keys from your `.bib` file: `Lorem ipsum [@doe2020; @smith1999, p. 12]`.

**Tip:** Each chapter can be its own `.md` file. Upload in the desired order.

**Note:** Page breaks between chapters in DOCX depend on your Word template. This app inserts a simple visual separator only.
"""
)
