import io
import os
import zipfile
from datetime import datetime

import streamlit as st

from src.openai_translate import OpenAITranslator
from src.text_utils import parse_glossary_lines
from src.image_translate import translate_image_bytes
from src.pdf_translate import translate_pdf_bytes
from src.pptx_translate import translate_pptx_bytes
from src.xlsm_translate import translate_excel_to_xls_bytes  # MUST exist

st.set_page_config(page_title="CKST Translator", layout="wide")


def get_api_key() -> str:
    # Never show the key or any "loaded/hidden" message in the UI.
    try:
        v = st.secrets.get("OPENAI_API_KEY", "")
        if v:
            return str(v)
    except Exception:
        pass
    return os.getenv("OPENAI_API_KEY", "") or ""


def suffix_for_lang(lang: str) -> str:
    l = (lang or "").lower()
    if l.startswith("en"):
        return "EN"
    if l.startswith("pt"):
        return "PTBR"
    return l.upper().replace("-", "").replace("_", "")


api_key = get_api_key()

with st.sidebar:
    st.header("OpenAI")

    direction = st.selectbox(
        "Direction",
        options=["PT-BR → EN", "EN → PT-BR"],
        index=0,
    )

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
    )

if direction == "PT-BR → EN":
    source_lang = "pt-BR"
    target_lang = "en"
else:
    source_lang = "en"
    target_lang = "pt-BR"

st.title(f"CKST Techpack Translator ({source_lang} ➜ {target_lang})")
st.caption("PDF / PPTX / XLSM / XLS / Images — handbag terminology focused")

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
        f"Glossary (optional) — one per line: `{source_lang} => {target_lang}`",
        value=(
            "OURO BATIDO => BRUSH GOLD\n"
            "alça => strap\n"
            "forro => lining\n"
            "ferragem => hardware\n"
            "zíper => zipper\n"
            "cursor => zipper puller\n"
        ),
        height=240,
        help="The glossary follows the direction you selected above.",
    )

with colB:
    extra_instructions = st.text_area(
        "Extra instructions (optional)",
        value=(
            "Use handbag / softgoods manufacturing terminology.\n"
            "Keep measurements, codes, SKUs, and numbers unchanged.\n"
            "Be clear and factory-friendly."
        ),
        height=240,
    )

glossary = parse_glossary_lines(glossary_text)

st.divider()

uploaded_files = st.file_uploader(
    "Upload your files",
    type=["pdf", "pptx", "xlsm", "xls", "png", "jpg", "jpeg", "webp"],
    accept_multiple_files=True,
)

run = st.button("Translate", type="primary", disabled=not (uploaded_files and api_key))


def build_translator():
    # Tolerant to different OpenAITranslator signatures
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
    """
    Canonical signature:
      func(data, translator, source_lang=..., target_lang=..., glossary=..., extra_instructions=..., on_progress=...)
    With fallbacks.
    """
    attempts = [
        lambda: func(
            data,
            translator,
            source_lang=source_lang,
            target_lang=target_lang,
            glossary=glossary,
            extra_instructions=extra_instructions,
            on_progress=on_progress,
        ),
        lambda: func(
            data,
            translator,
            glossary=glossary,
            extra_instructions=extra_instructions,
            on_progress=on_progress,
        ),
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
    suffix = suffix_for_lang(target_lang)

    for idx, uf in enumerate(uploaded_files, start=1):
        filename = uf.name
        ext = filename.split(".")[-1].lower()
        data = uf.read()

        status.info(f"Processing **{filename}** ({idx}/{total_files})")

        main_bar = st.progress(0.0, text="starting…")
        batch_bar = st.progress(0.0, text="")  # used for Excel batches

        def on_progress_generic(label: str, done: int, total: int):
            total = max(1, int(total))
            done = max(0, int(done))
            pct = min(1.0, done / total)
            main_bar.progress(pct, text=f"{label} ({done}/{total})")

        def on_progress_excel(label: str, done: int, total: int):
            total = max(1, int(total))
            done = max(0, int(done))
            pct = min(1.0, done / total)
            if label in ("pages", "sheets", "tabs"):
                main_bar.progress(pct, text=f"pages ({done}/{total})")
            elif label in ("batches",):
                batch_bar.progress(pct, text=f"batches ({done}/{total})")
            else:
                main_bar.progress(pct, text=f"{label} ({done}/{total})")

        try:
            if ext == "pdf":
                out_bytes = call_translate(translate_pdf_bytes, data, translator, on_progress_generic)
                out_name = filename[:-4] + f"_{suffix}.pdf"
                mime = "application/pdf"

            elif ext == "pptx":
                out_bytes = call_translate(translate_pptx_bytes, data, translator, on_progress_generic)
                out_name = filename[:-5] + f"_{suffix}.pptx"
                mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

            elif ext in ("png", "jpg", "jpeg", "webp"):

                def on_progress_image(pct: float, msg: str):
                    pct = max(0.0, min(1.0, float(pct)))
                    main_bar.progress(pct, text=msg)

                out_bytes = translate_image_bytes(
                    image_bytes=data,
                    filename=filename,
                    api_key=api_key,
                    model=model,
                    source_lang=source_lang,
                    target_lang=target_lang,
                    glossary=glossary,
                    extra_instructions=extra_instructions,
                    progress_cb=on_progress_image,
                )
                out_name = filename.rsplit(".", 1)[0] + f"_{suffix}.png"
                mime = "image/png"
                st.image(out_bytes, caption=out_name, use_container_width=True)

            elif ext in ("xlsm", "xls"):
                out_bytes = translate_excel_to_xls_bytes(
                    excel_bytes=data,
                    input_ext=ext,
                    translator=translator,
                    source_lang=source_lang,
                    target_lang=target_lang,
                    glossary=glossary,
                    extra_instructions=extra_instructions,
                    on_progress=on_progress_excel,
                    batch_size=25,
                )
                out_name = filename.rsplit(".", 1)[0] + f"_{suffix}.xls"
                mime = "application/vnd.ms-excel"
                batch_bar.empty()

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

        zip_name = f"translations_{suffix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        st.download_button(
            "Download ALL as ZIP",
            data=zbuf.getvalue(),
            file_name=zip_name,
            mime="application/zip",
        )
