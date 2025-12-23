import shutil
import subprocess
import tempfile
from pathlib import Path


def soffice_available() -> bool:
    return shutil.which("soffice") is not None


def convert_office_bytes(input_bytes: bytes, input_ext: str, output_ext: str) -> bytes:
    """
    Convert Office docs using LibreOffice (soffice) headless.

    input_ext/output_ext examples: "xls", "xlsx", "xlsm"
    Returns converted file bytes.
    """
    if not soffice_available():
        raise RuntimeError(
            "LibreOffice (soffice) is not available. Install LibreOffice and ensure 'soffice' is in PATH. "
            "On Streamlit Cloud add packages.txt with 'libreoffice'."
        )

    input_ext = input_ext.lower().lstrip(".")
    output_ext = output_ext.lower().lstrip(".")

    with tempfile.TemporaryDirectory() as td:
        td_path = Path(td)
        in_path = td_path / f"input.{input_ext}"
        in_path.write_bytes(input_bytes)

        # Filters (xls needs a better hint sometimes)
        convert_to_candidates = []
        if output_ext == "xls":
            convert_to_candidates = ["xls", 'xls:"MS Excel 97"']
        else:
            convert_to_candidates = [output_ext]

        last_err = None
        for conv in convert_to_candidates:
            try:
                cmd = [
                    "soffice",
                    "--headless",
                    "--nologo",
                    "--norestore",
                    "--nolockcheck",
                    "--nodefault",
                    "--convert-to",
                    conv,
                    "--outdir",
                    str(td_path),
                    str(in_path),
                ]
                subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

                outs = list(td_path.glob(f"*.{output_ext}"))
                if not outs:
                    raise RuntimeError("Converted file not found after LibreOffice conversion.")
                return outs[0].read_bytes()
            except Exception as e:
                last_err = e

        raise RuntimeError(f"LibreOffice conversion failed: {last_err}")
