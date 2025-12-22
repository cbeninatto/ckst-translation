import io
import os
import zipfile
from datetime import datetime

import streamlit as st

from src.openai_translate import OpenAITranslator
from src.text_utils import parse_glossary_lines
from src.pdf_translate import translate_pdf_bytes
from src.pptx_translate import translate_pptx_bytes
from src.xlsm_translate import translate_xlsm_bytes


st.set_page_config(page_title="CKST Techpack Translator", layout="wide")

st.title("Techpack Translator (PT-BR → EN) — PDF / PPTX / XLSM")

with st.sidebar:
    st.header("OpenAI")
    api_key = st.text_input("OpenAI API key", type="password", value=os.getenv("OPENAI_API_KEY", ""))
    model = st.selectbox(
        "Model",
        options=[
            "gpt-4.1",
            "gpt-4.1-mini",
            "gpt-4o",
            "o4-mini",
        ],
        index=0,
        help="If your org is not verified for GPT-5 yet, use GPT-4.1 options.",
    )
    st.caption("Tip: On Streamlit Cloud, set OPENAI_API_KEY in Secrets.")

st.divider()

default_glossary = """\
alça => strap
alça de ombro => shoulder strap
alça transversal => crossbody strap
forro => lining
corpo => body
ferragem => hardware
rebite => rivet
mosquetão => swivel clasp
argola => ring
meia argola => D-ring
zíper => zipper
cursor => zipper puller
bolso interno => inner pocket
bolso externo => outer pocket
etiqueta => label
acabamento => finish
costura => stitching
pesponto => topstitching
espessura => thickness
largura => width
altura => height
comprimento => length
"""

col1, col2 = st.columns([1, 1])
with col1:
    glossary_text = st.text_area(
        "Glossary (optional) — one per line: `portuguese => english`",
        value=default_glossary,
        height=220,
    )
with col2:
    st.markdown(
        """
**How it works**
- Upload one or more **PDF / PPTX / XLSM**.
- The app replaces text **in place** (same page/slide/cell).
- For **XLSM**, only the **print area** is kept; everything outside it is removed.
        """
    )

uploaded = st.file_uploader(
    "Upload files",
    type=["pdf", "pptx", "xlsm"],
    accept_multiple_files=True,
)

run = st.button("Translate to English", type="primary", disabled=(not uploaded or not api_key))

if run:
    glossary = parse_glossary_lines(glossary_text)
    translator = OpenAITranslator(api_key=api_key, model=model, glossary=glossary)

    results = []
    overall = st.progress(0.0, text="Starting...")
    status = st.empty()

    total_files = len(uploaded)
    for file_idx, uf in enumerate(uploaded, start=1):
        filename = uf.name
        ext = filename.split(".")[-1].lower()
        data = uf.read()

        status.info(f"Translating **{filename}** ({file_idx}/{total_files})")

        per_file_bar = st.progress(0.0, text=f"{filename}: preparing...")

        def on_progress(label: str, done: int, total: int):
            pct = 0.0 if total == 0 else done / total
            per_file_bar.progress(pct, text=f"{filename}: {label} ({done}/{total})")

        try:
            if ext == "pdf":
                out_bytes = translate_pdf_bytes(data, translator, on_progress=on_progress)
                out_name = filename[:-4] + "_EN.pdf"
                mime = "application/pdf"
            elif ext == "pptx":
                out_bytes = translate_pptx_bytes(data, translator, on_progress=on_progress)
                out_name = filename[:-5] + "_EN.pptx"
                mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            elif ext == "xlsm":
                out_bytes = translate_xlsm_bytes(data, translator, on_progress=on_progress)
                out_name = filename[:-5] + "_EN.xlsm"
                mime = "application/vnd.ms-excel.sheet.macroEnabled.12"
            else:
                raise ValueError(f"Unsupported file type: {ext}")

            results.append((out_name, out_bytes, mime))
            st.success(f"✅ Done: {out_name}")
            st.download_button(f"Download {out_name}", data=out_bytes, file_name=out_name, mime=mime)
        except Exception as e:
            st.error(f"❌ Error translating {filename}: {e}")
        finally:
            overall.progress(file_idx / total_files, text=f"Processed {file_idx}/{total_files} file(s)")

    if results:
        zbuf = io.BytesIO()
        with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as zf:
            for out_name, out_bytes, _ in results:
                zf.writestr(out_name, out_bytes)
        zname = f"translations_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        st.download_button("Download ALL as ZIP", data=zbuf.getvalue(), file_name=zname, mime="application/zip")
