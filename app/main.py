from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware

import traceback
import tempfile
import os
import shutil
import subprocess

from app.checker import check_docx


app = FastAPI(title="Word Phrase Checker")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.mount("/web", StaticFiles(directory="web", html=True), name="web")


@app.get("/")
def root():
    return FileResponse("web/index.html")


def _find_soffice() -> str | None:
    # 1) In PATH
    for name in ("soffice", "soffice.exe", "libreoffice", "libreoffice.exe"):
        p = shutil.which(name)
        if p:
            return p

    # 2) Common Windows paths
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return None


def convert_doc_to_docx(doc_bytes: bytes) -> bytes:
    """
    Convert .doc -> .docx using LibreOffice (soffice) headless.
    Requires LibreOffice installed on the server machine.
    """
    soffice = _find_soffice()
    if not soffice:
        raise RuntimeError(
            "Không tìm thấy LibreOffice (soffice). "
            "Hãy cài LibreOffice hoặc thêm soffice vào PATH để đọc file .doc."
        )

    with tempfile.TemporaryDirectory() as td:
        in_path = os.path.join(td, "input.doc")
        out_dir = td

        with open(in_path, "wb") as f:
            f.write(doc_bytes)

        # Convert
        # soffice --headless --convert-to docx --outdir <dir> <file>
        cmd = [soffice, "--headless", "--nologo", "--nolockcheck", "--convert-to", "docx", "--outdir", out_dir, in_path]
        proc = subprocess.run(cmd, capture_output=True, text=True)

        if proc.returncode != 0:
            raise RuntimeError(
                "LibreOffice convert lỗi. "
                f"stdout={proc.stdout[-300:]}, stderr={proc.stderr[-300:]}"
            )

        out_path = os.path.join(td, "input.docx")
        if not os.path.exists(out_path):
            # đôi khi LibreOffice đổi tên khác, tìm file .docx trong folder
            outs = [p for p in os.listdir(td) if p.lower().endswith(".docx")]
            if not outs:
                raise RuntimeError("Convert xong nhưng không thấy file .docx output.")
            out_path = os.path.join(td, outs[0])

        with open(out_path, "rb") as f:
            return f.read()


@app.post("/api/check")
async def api_check(
    file: UploadFile = File(...),
    phrases: str = Form(...),

    case_sensitive: bool = Form(False),
    whole_word: bool = Form(False),
    scan_headers_footers: bool = Form(True),
    check_du_toan_rule: bool = Form(True),

    spellcheck_vi: bool = Form(False),
    spell_max_distance: int = Form(2),
):
    try:
        filename = (file.filename or "").lower().strip()
        data = await file.read()
        if not data:
            return JSONResponse(status_code=400, content={"error": "File rỗng hoặc đọc file thất bại."})

        phrase_list = [line.strip() for line in (phrases or "").splitlines() if line.strip()]
        if not phrase_list:
            return JSONResponse(status_code=400, content={"error": "Bạn chưa nhập danh sách cụm từ (mỗi dòng 1 cụm)."})

        # Support .docx + .doc
        if filename.endswith(".docx"):
            docx_bytes = data
        elif filename.endswith(".doc"):
            docx_bytes = convert_doc_to_docx(data)
        else:
            return JSONResponse(status_code=400, content={"error": "Chỉ hỗ trợ .docx và .doc"})

        # clamp
        if spell_max_distance < 1:
            spell_max_distance = 1
        if spell_max_distance > 3:
            spell_max_distance = 3

        result = check_docx(
            docx_bytes,
            phrase_list,
            case_sensitive=case_sensitive,
            whole_word=whole_word,
            scan_headers_footers=scan_headers_footers,
            spellcheck_vi=spellcheck_vi,
            spell_max_distance=spell_max_distance,
              check_du_toan_rule=check_du_toan_rule,
        )
        return result

    except Exception as e:
        tb = traceback.format_exc()
        print(tb)
        return JSONResponse(
            status_code=500,
            content={"error": "Backend exception", "detail": str(e)},
        )
