import os
import tempfile
from typing import List, Optional
from docx import Document


def convert_doc_to_docx(paths: List[str]) -> List[str]:
    out: List[str] = []
    try:
        import win32com.client  # type: ignore
    except Exception:
        return out
    try:
        app = win32com.client.Dispatch("Word.Application")
        app.Visible = False
        tmpdir = tempfile.mkdtemp(prefix="doc2docx_")
        for p in paths:
            try:
                doc = app.Documents.Open(p)
                base = os.path.splitext(os.path.basename(p))[0] + ".docx"
                dest = os.path.join(tmpdir, base)
                doc.SaveAs(dest, FileFormat=16)
                doc.Close(False)
                out.append(dest)
            except Exception:
                pass
        app.Quit()
    except Exception:
        return []
    return out


def read_docx(path: str) -> Document:
    return Document(path)
