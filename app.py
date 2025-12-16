from __future__ import annotations

import os
import zipfile
from io import BytesIO
from typing import Dict, List, Tuple

import streamlit as st

from src.openai_translate import OpenAITranslator
from src.pdf_translate import translate_pdf_bytes
from src.pptx_translate import translate_pptx_bytes
from src.text_utils import parse_glossary_lines


st.set_page_config(page_title="Handbag Dev Translator (PT→EN)", layout="wide")

st.title("Handbag Dev Translator (PT-BR → English)")
st.caption("Translate designer PDF/PPTX to factory-friendly English (handbag terminology).")

DEFAULT_HANDBAG_GLOSSARY = """\
couro=leather
couro legítimo=genuine leather
couro sintético=synthetic leather
napa=nappa (leather/PU)
camurça=suede
microfibra=microfiber
lona=canvas
sarja=twill
forro=lining
forração=lining
entretela=interlining
espuma=foam
reforço=reinforcement
fita=webbing tape
viés=bias tape
pesponto=topstitching
costura=stitching
alça=strap
alça de mão=top handle
alça tiracolo=crossbody strap
regulador=adjuster buckle
fivela=buckle
mosquetão=swivel snap hook
argola=D-ring
meia argola=half D-ring
rebite=rivet
ilhós=grommet
zíper=zipper
cursor=zipper slider
puxador=zipper puller
ferragem=hardware
banho=plating
níquel=nickel
ouro=gold
ouro velho=antique gold
grafite=graphite
gunmetal=gunmetal
"""

# ---- Sidebar ----
with st.sidebar:
    st.header("Settings")

    key_from_secrets = st.secrets.get("OPENAI_API_KEY") if hasattr(st, "secrets") else None
    api_key = key_from_secrets or os.getenv("OPENAI_API_KEY")
    if not api_key:
        api_key = st.text_input("OpenAI API key", type="password")

    st.subheader("Model (latest)")
    model = st.selectbox(
        "Choose model",
        options=["gpt-5.2", "gpt-5.2-pro", "gpt-5.2-chat-latest"],
        index=0,
        help="gpt-5.2 = best default. gpt-5.2-pro = max quality. gpt-5.2-chat-latest = faster/cheaper style.",
    )

    reasoning_effort = st.selectbox(
        "reasoning.effort",
        options=["none", "minimal", "low", "medium", "high", "xhigh"],
        index=2,
        help="Higher can help with ambiguous technical wording (but uses more tokens).",
    )

    source_lang = st.text_input("Source language", value="pt-BR")
    target_lang = st.text_input("Target language", value="en-US")

    st.divider()
    st.subheader("Handbag terminology pack")
    use_default_pack = st.checkbox("Use built-in handbag glossary", value=True)

    glossary_raw = st.text_area(
        "Glossary (Portuguese=English), one per line",
        value=(DEFAULT_HANDBAG_GLOSSARY if use_default_pack else ""),
        height=220,
    )
    glossary = parse_glossary_lines(glossary_raw)

    extra_instructions = st.text_area(
        "Extra instructions (optional)",
        value=(
            "Write like a handbag tech pack for Chinese factories.\n"
            "Prefer terms: lining, reinforcement, piping, topstitching, hardware.\n"
            "If a word could be 'handle' vs 'strap', pick based on handbag context.\n"
            "Keep bullet lists as bullet lists.\n"
        ),
        height=140,
    )

# ---- Upload ----
uploaded = st.file_uploader(
    "Upload PDF and/or PPTX (multiple allowed)",
    type=["pdf", "pptx"],
    accept_multiple_files=True,
)

translate_btn = st.button("Translate", type="primary", disabled=(not uploaded or not api_key))

if not uploaded:
    st.info("Upload files to begin.")
    st.stop()

if translate_btn and not api_key:
    st.error("Missing OpenAI API key.")
    st.stop()


def run_translation() -> List[Tuple[str, bytes]]:
    translator = OpenAITranslator(api_key=api_key, model=model, reasoning_effort=reasoning_effort)

    outputs: List[Tuple[str, bytes]] = []

    overall = st.progress(0)
    overall_status = st.empty()

    for f_i, f in enumerate(uploaded, start=1):
        name = f.name
        ext = name.lower().split(".")[-1]
        raw = f.getvalue()

        overall_status.write(f"Processing **{name}** ({f_i}/{len(uploaded)})...")
        file_status = st.empty()
        file_progress = st.progress(0)

        try:
            if ext == "pdf":

                def on_pdf_progress(page_idx: int, total_pages: int):
                    if total_pages <= 0:
                        return
                    if page_idx < total_pages:
                        file_status.write(f"Translating PDF page **{page_idx+1}/{total_pages}** …")
                        file_progress.progress((page_idx + 1) / total_pages)
                    else:
                        file_status.write("PDF done.")
                        file_progress.progress(1.0)

                out_bytes = translate_pdf_bytes(
                    raw,
                    translator=translator,
                    source_lang=source_lang,
                    target_lang=target_lang,
                    glossary=glossary,
                    extra_instructions=extra_instructions,
                    on_progress=on_pdf_progress,
                )

                out_name = name.rsplit(".", 1)[0] + " — EN.pdf"
                outputs.append((out_name, out_bytes))

                if out_bytes == raw:
                    st.warning(
                        f"⚠️ {name}: No extractable text found (likely scanned/image-only PDF). "
                        "This version translates text-based PDFs."
                    )

            elif ext == "pptx":

                def on_pptx_progress(slide_1based: int, total_slides: int):
                    if total_slides <= 0:
                        return
                    file_status.write(f"Translating PPTX slide **{slide_1based}/{total_slides}** …")
                    file_progress.progress(slide_1based / total_slides)

                out_bytes = translate_pptx_bytes(
                    raw,
                    translator=translator,
                    source_lang=source_lang,
                    target_lang=target_lang,
                    glossary=glossary,
                    extra_instructions=extra_instructions,
                    on_progress=on_pptx_progress,
                )

                out_name = name.rsplit(".", 1)[0] + " — EN.pptx"
                outputs.append((out_name, out_bytes))

            else:
                st.warning(f"Skipping unsupported file: {name}")

        except Exception as e:
            st.error(f"❌ Error translating {name}: {e}")

        file_progress.empty()
        file_status.empty()

        overall.progress(f_i / len(uploaded))

    overall_status.write("Done.")
    overall.empty()
    return outputs


if translate_btn:
    outputs = run_translation()

    if not outputs:
        st.error("No outputs created.")
        st.stop()

    st.success(f"✅ Translated {len(outputs)} file(s).")

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
        mime = (
            "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            if out_name.lower().endswith(".pptx")
            else "application/pdf"
        )
        st.download_button(
            f"Download: {out_name}",
            data=out_bytes,
            file_name=out_name,
            mime=mime,
        )
