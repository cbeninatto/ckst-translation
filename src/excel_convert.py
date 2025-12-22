import os
import shutil
import subprocess
import tempfile
from pathlib import Path


def soffice_available() -> bool:
    return shutil.which("soffice") is not None


def convert_office_bytes(input_bytes: bytes, input_ext: str, output_ext: str) -> bytes:
    """
    Converts office files using LibreOffice (soffice) headless.

    input_ext/output_ext examples: "xls", "xlsx", "xlsm"
    Returns converted file bytes.

    Raises RuntimeError if soffice is missing or conversion fails.
    """
    if not soffice_available():
        raise RuntimeError(
            "LibreOffice (soffice) is not available on this system. "
            "To convert XLS/XLSM to XLS, install LibreOffice and ensure 'soffice' is in PATH."
        )

    input_ext = input_ext.lower().lstrip(".")
    output_ext = output_ext.lower().lstrip(".")

    with tempfile.TemporaryDirectory() as td:
        td_path = Path(td)
        in_path = td_path / f"input.{input_ext}"
        in_path.write_bytes(input_bytes)

        # Try a couple filter strings for better compatibility
        convert_to_candidates = []
        if output_ext == "xls":
            convert_to_candidates = ["xls", 'xls:"MS Excel 97"']
        elif output_ext == "xlsx":
            convert_to_candidates = ["xlsx"]
        elif output_ext == "xlsm":
            convert_to_candidates = ["xlsm"]
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

                # Find output
                outs = list(td_path.glob(f"input*.{output_ext}"))
                if not outs:
                    # Sometimes LibreOffice changes the basename slightly; fallback:
                    outs = list(td_path.glob(f"*.{output_ext}"))
                if not outs:
                    raise RuntimeError("Converted file not found after conversion.")
                return outs[0].read_bytes()

            except Exception as e:
                last_err = e

        raise RuntimeError(f"LibreOffice conversion failed: {last_err}")
