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

st.set_page_config(page_title="CKST Translator", layout="wide")

st.title("CKST Techpack Translator (PT-BR ➜ EN)")
st.caption("PDF / PPTX / XLSM — handbag terminology focused")


def get_api_key() -> str:
    # Streamlit Cloud secrets
    try:
        v = st.secrets.get("OPENAI_API_KEY", "")
        if v:
            return str(v)
    except Exception:
        pass
    # Local environment variable
    return os.getenv("OPENAI_API_KEY", "") or ""


api_key = get_api_key()

with st.sidebar:
    st.header("OpenAI")

    model = st.selectbox(
        "Model",
        options=[
            "gpt-5.2-pro",
            "gpt-5.2",
            "gpt-5.1",
            "gpt-5",
            "gpt-5-mini",
            "gpt-5-nano",
            "gpt-4.1",
            "gpt-4.1-mini",
            "gpt-4.1-nano",
            "gpt-4o",
            "o4-mini",
        ],
        index=0,
    )

    reasoning_effort = st.selectbox(
        "Reasoning effort (if supported)",
        options=["none", "low", "medium", "high", "xhigh"],
        index=2,
        help="If the selected model rejects reasoning controls, set to 'none'.",
    )

# IMPORTANT: do not print anything like “OPENAI_API_KEY loaded (hidden)”
if not api_key:
    st.warning(
        "OpenAI key is not configured.\n\n"
        "• Streamlit Cloud: add `OPENAI_API_KEY` in Secrets.\n"
        "• Local: set environment variable `OPENAI_API_KEY`.\n"
        "Then reload."
    )

st.divider()

colA, colB = st.columns([1, 1])

with colA:
    glossary_text = st.text_area(
        "Glossary (optional) — one per line: `pt => en`",
        value=(
            "alça => strap\n"
            "alça de ombro => shoulder strap\n"
            "alça transversal => crossbody strap\n"
            "forro => lining\n"
            "corpo => body\n"
            "ferragem => hardware\n"
            "rebite => rivet\n"
            "mosquetão => swivel clasp\n"
            "argola => ring\n"
            "meia argola => D-ring\n"
            "zíper => zipper\n"
            "cursor => zipper puller\n"
            "bolso interno => inner pocket\n"
            "bolso externo => outer pocket\n"
            "etiqueta => label\n"
            "acabamento => finish\n"
            "costura => stitching\n"
            "pesponto => topstitching\n"
            "reforço => reinforcement\n"
            "espuma => foam\n"
            "entretela => interlining\n"
            "vivo => piping\n"
            "viés => binding tape\n"
            "ouro batido => brushed gold\n"
        ),
        height=220,
    )

with colB:
    extra_instructions = st.text_area(
        "Extra instructions (optional)",
        value=(
            "Use handbag / softgoods manufacturing terminology.\n"
            "Keep measurements, codes, SKUs, and numbers unchanged.\n"
            "Keep the structure and be factory-friendly."
        ),
        height=220,
    )

glossary = parse_glossary_lines(glossary_text)

st.divider()

uploaded_files = st.file_uploader(
    "Upload your files",
    type=["pdf", "pptx", "xlsm"],
    accept_multiple_files=True,
)

run = st.button("Translate to English", type="primary", disabled=not (uploaded_files and api_key))


def build_translator():
    try:
        return OpenAITranslator(api_key=api_key, model=model, reasoning_effort=reasoning_effort)
    except TypeError:
        pass
    try:
        return OpenAITranslator(api_key=api_key, model=model)
    except TypeError:
        pass
    return OpenAITranslator(api_key, model)


def call_translate(func, data: bytes, translator, on_progress):
    attempts = [
        lambda: func(
            data,
            translator,
            source_lang="pt-BR",
            target_lang="en",
            glossary=glossary,
            extra_instructions=extra_instructions,
            on_progress=on_progress,
        ),
        lambda: func(data, translator, glossary=glossary, extra_instructions=extra_instructions, on_progress=on_progress),
        lambda: func(data, translator, on_progress=on_progress),
        lambda: func(data, translator),
        lambda: func(data),
    ]
    last_err = None
    for a in attempts:
        try:
            return a()
        except TypeError as e:
            last_err = e
    raise last_err


if run:
    translator = build_translator()

    results = []
    overall = st.progress(0.0, text="Starting...")
    status = st.empty()

    total_files = len(uploaded_files)

    for idx, uf in enumerate(uploaded_files, start=1):
        filename = uf.name
        ext = filename.split(".")[-1].lower()
        data = uf.read()

        status.info(f"Processing **{filename}** ({idx}/{total_files})")

        # Main per-file bar (we'll show "pages" for XLSM, normal labels for others)
        main_bar = st.progress(0.0, text="starting...")
        # Secondary bar only for XLSM (cells/batches)
        sub_ph = st.empty()

        def on_progress_generic(label: str, done: int, total: int):
            total = max(1, int(total))
            done = max(0, int(done))
            pct = min(1.0, done / total)
            main_bar.progress(pct, text=f"{label} ({done}/{total})")

        def on_progress_xlsm(label: str, done: int, total: int):
            """
            XLSM uses two progress indicators:
              - label 'pages' (tabs) -> main_bar
              - label 'batches' -> sub bar
            """
            total = max(1, int(total))
            done = max(0, int(done))
            pct = min(1.0, done / total)

            if label in ("pages", "sheets", "tabs"):
                main_bar.progress(pct, text=f"pages ({done}/{total})")
            elif label in ("batches", "cells"):
                # create/refresh sub progress bar
                pb = sub_ph.progress(pct, text=f"cells ({done}/{total})")
            else:
                main_bar.progress(pct, text=f"{label} ({done}/{total})")

        try:
            if ext == "pdf":
                out_bytes = call_translate(translate_pdf_bytes, data, translator, on_progress_generic)
                out_name = filename[:-4] + "_EN.pdf"
                mime = "application/pdf"

            elif ext == "pptx":
                out_bytes = call_translate(translate_pptx_bytes, data, translator, on_progress_generic)
                out_name = filename[:-5] + "_EN.pptx"
                mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

            elif ext == "xlsm":
                out_bytes = translate_xlsm_bytes(
                    data,
                    translator,
                    source_lang="pt-BR",
                    target_lang="en",
                    glossary=glossary,
                    extra_instructions=extra_instructions,
                    on_progress=on_progress_xlsm,
                )
                out_name = filename[:-5] + "_EN.xlsm"
                mime = "application/vnd.ms-excel.sheet.macroEnabled.12"

                # Remove the sub bar after finishing
                sub_ph.empty()

            else:
                raise ValueError(f"Unsupported file type: {ext}")

            results.append((out_name, out_bytes, mime))
            st.success(f"✅ Done: {out_name}")
            st.download_button(
                label=f"Download {out_name}",
                data=out_bytes,
                file_name=out_name,
                mime=mime,
            )

        except Exception as e:
            st.error(f"❌ Error translating {filename}: {e}")

        overall.progress(idx / total_files, text=f"Processed {idx}/{total_files} file(s)")

    if results:
        zbuf = io.BytesIO()
        with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as zf:
            for out_name, out_bytes, _ in results:
                zf.writestr(out_name, out_bytes)

        zip_name = f"translations_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        st.download_button(
            "Download ALL as ZIP",
            data=zbuf.getvalue(),
            file_name=zip_name,
            mime="application/zip",
        )
