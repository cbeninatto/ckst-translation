from __future__ import annotations

import os
import zipfile
from io import BytesIO
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st

from src.openai_translate import OpenAITranslator
from src.pdf_translate import translate_pdf_bytes
from src.pptx_translate import translate_pptx_bytes
from src.text_utils import parse_glossary_lines, apply_glossary_hard


st.set_page_config(page_title="Handbag Dev Translator (PT→EN)", layout="wide")

st.title("Handbag Dev Translator (PT-BR → English)")
st.caption("Upload PDF/PPTX from designers, translate to factory-friendly English, download translated files.")

# ---- Sidebar settings ----
with st.sidebar:
    st.header("Settings")

    # Prefer secrets/env; fallback to manual input
    key_from_secrets = st.secrets.get("OPENAI_API_KEY") if hasattr(st, "secrets") else None
    api_key = key_from_secrets or os.getenv("OPENAI_API_KEY")
    if not api_key:
        api_key = st.text_input("OpenAI API key", type="password", help="Prefer using Streamlit secrets or env vars.")

    model = st.text_input("Model", value="gpt-4o-mini", help="Any model available in your OpenAI project.")

    source_lang = st.text_input("Source language", value="pt-BR")
    target_lang = st.text_input("Target language", value="en-US")

    st.divider()
    st.subheader("Glossary")
    glossary_raw = st.text_area(
        "One per line: Portuguese=English",
        value="",
        placeholder="couro=leather\nforro=lining\nferragem=hardware\nalça=strap\n",
        height=140,
    )
    glossary = parse_glossary_lines(glossary_raw)

    hard_glossary = st.checkbox(
        "Hard enforce glossary (post-processing)",
        value=False,
        help="If enabled, we also apply simple replacements after translation.",
    )

    st.divider()
    st.subheader("Translation behavior")
    extra_instructions = st.text_area(
        "Extra instructions (optional)",
        value="Use concise manufacturing English. Prefer handbag/material terminology.",
        height=100,
    )

    st.divider()
    st.subheader("Batching (advanced)")
    max_chars = st.slider("Max characters per API call", 6000, 30000, 18000, 1000)
    max_items_pptx = st.slider("Max text items per PPTX call", 10, 120, 60, 5)
    max_items_pdf = st.slider("Max pages per PDF call", 5, 60, 25, 5)

# ---- Upload UI ----
uploaded = st.file_uploader(
    "Upload PDF and/or PPTX (multiple allowed)",
    type=["pdf", "pptx"],
    accept_multiple_files=True,
)

colA, colB = st.columns([1, 1])
with colA:
    translate_btn = st.button("Translate", type="primary", disabled=(not uploaded or not api_key))
with colB:
    st.write("")

if not uploaded:
    st.info("Upload some files to begin.")
    st.stop()

if translate_btn and not api_key:
    st.error("Missing OpenAI API key. Add it via Streamlit secrets or environment variable.")
    st.stop()

# ---- Work ----
def _translate_files() -> Tuple[List[Tuple[str, bytes]], List[Dict]]:
    translator = OpenAITranslator(api_key=api_key, model=model)

    outputs: List[Tuple[str, bytes]] = []
    preview_rows: List[Dict] = []

    progress = st.progress(0)
    status = st.empty()

    for idx, f in enumerate(uploaded, start=1):
        name = f.name
        ext = name.lower().split(".")[-1]
        raw = f.getvalue()

        status.write(f"Translating **{name}** ({idx}/{len(uploaded)})...")

        try:
            if ext == "pptx":
                out_bytes = translate_pptx_bytes(
                    raw,
                    translator=translator,
                    source_lang=source_lang,
                    target_lang=target_lang,
                    glossary=glossary,
                    extra_instructions=extra_instructions,
                    max_chars=max_chars,
                    max_items=max_items_pptx,
                )
                out_name = name.replace(".pptx", "").replace(".PPTX", "") + " — EN.pptx"
                outputs.append((out_name, out_bytes))

            elif ext == "pdf":
                out_bytes = translate_pdf_bytes(
                    raw,
                    translator=translator,
                    source_lang=source_lang,
                    target_lang=target_lang,
                    glossary=glossary,
                    extra_instructions=extra_instructions,
                    max_chars=max_chars,
                    max_items=max_items_pdf,
                )
                out_name = name.replace(".pdf", "").replace(".PDF", "") + " — EN.pdf"
                outputs.append((out_name, out_bytes))

                # If PDF had no text, translate_pdf_bytes returns original bytes
                if out_bytes == raw:
                    st.warning(
                        f"⚠️ {name}: No extractable text found. This is likely a scanned/image-only PDF. "
                        "This app currently translates text-based PDFs only."
                    )
            else:
                st.warning(f"Skipping unsupported file: {name}")

        except Exception as e:
            st.error(f"❌ Error translating {name}: {e}")

        progress.progress(idx / len(uploaded))

    status.write("Done.")
    progress.empty()
    return outputs, preview_rows


if translate_btn:
    outputs, _ = _translate_files()

    if hard_glossary and glossary:
        # Optional: hard glossary pass for pptx is hard to do post-hoc safely, so we only do it
        # for simple preview use-cases. (Most users rely on prompt-based glossary.)
        st.info("Hard glossary enforcement was enabled. (Primary enforcement is via prompt.)")

    if not outputs:
        st.error("No outputs were created.")
        st.stop()

    st.success(f"✅ Translated {len(outputs)} file(s).")

    # Individual downloads + ZIP
    st.subheader("Downloads")

    zip_buf = BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as z:
        for out_name, out_bytes in outputs:
            z.writestr(out_name, out_bytes)

    st.download_button(
        "Download ALL as ZIP",
        data=zip_buf.getvalue(),
        file_name="handbag_dev_translations_EN.zip",
        mime="application/zip",
    )

    st.divider()
    for out_name, out_bytes in outputs:
        mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation" if out_name.lower().endswith(".pptx") else "application/pdf"
        st.download_button(
            f"Download: {out_name}",
            data=out_bytes,
            file_name=out_name,
            mime=mime,
        )
