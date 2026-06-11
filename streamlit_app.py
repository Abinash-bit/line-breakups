"""
Streamlit app: Fix DOCX line-break patterns using a reference DOCX.

Upload:
  1. The DOCX with WRONG line breaks (e.g. Bharat_ek_Khoj_21_Nikita.docx)
  2. The reference DOCX with CORRECT line breaks (e.g. BEK_EP_021_srt.docx)

Output: the first document's own text, re-broken to match the reference's
line-break pattern. Extra lines that exist only in the reference are left
as they are (skipped). No words of the first document are ever lost.

Run with:
    pip install streamlit python-docx
    streamlit run streamlit_app.py

NOTE: keep fix_linebreaks.py in the same folder as this file.
"""

import io

import streamlit as st

from fix_linebreaks import fix_linebreaks, write_docx, read_lines

st.set_page_config(page_title="DOCX Line-Break Fixer", page_icon="📄")

st.title("📄 DOCX Line-Break Fixer")
st.write(
    "Upload the document with **wrong line breaks** and the **reference** "
    "document with the correct line-break pattern. You'll get back the first "
    "document's text, re-broken to match the reference. Extra lines present "
    "only in the reference are kept as they are (skipped), and no words from "
    "the first document are lost."
)

col1, col2 = st.columns(2)
with col1:
    f1 = st.file_uploader("1️⃣ DOCX with wrong line breaks", type=["docx"], key="doc1")
with col2:
    f2 = st.file_uploader("2️⃣ Reference DOCX (correct breaks)", type=["docx"], key="doc2")

font_name = st.text_input("Output font", value="Courier New")

if f1 and f2:
    if st.button("🔧 Fix line breaks", type="primary"):
        with st.spinner("Aligning the two documents and re-breaking lines…"):
            try:
                lines = fix_linebreaks(io.BytesIO(f1.getvalue()), io.BytesIO(f2.getvalue()))

                buf = io.BytesIO()
                write_docx(lines, buf, font_name=font_name)
                buf.seek(0)

                n_in = len(read_lines(io.BytesIO(f1.getvalue())))
                n_ref = len(read_lines(io.BytesIO(f2.getvalue())))

                st.success(
                    f"Done! Input had {n_in} lines, reference has {n_ref} lines, "
                    f"fixed output has {len(lines)} lines."
                )

                out_name = f1.name.rsplit(".", 1)[0] + "_FIXED.docx"
                st.download_button(
                    "⬇️ Download fixed DOCX",
                    data=buf,
                    file_name=out_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )

                with st.expander("Preview (first 40 lines)"):
                    for text, has_note in lines[:40]:
                        st.text(text)
            except Exception as e:
                st.error(f"Something went wrong: {e}")
else:
    st.info("Upload both files to continue.")
