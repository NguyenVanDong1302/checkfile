from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import JSONResponse
from app.checker import check_docx

app = FastAPI()

@app.get("/api/health")
def health():
    return {"ok": True}

@app.post("/api/check")
async def api_check(
    file: UploadFile = File(...),
    phrases: str = Form(...),
    case_sensitive: bool = Form(False),
    whole_word: bool = Form(False),
    scan_headers_footers: bool = Form(True),
    spellcheck_vi: bool = Form(False),
    spell_max_distance: int = Form(2),
    check_du_toan_rule: bool = Form(True),
):
    filename = (file.filename or "").lower().strip()
    data = await file.read()

    if not filename.endswith(".docx"):
        return JSONResponse(status_code=400, content={"error": "Trên Vercel chỉ hỗ trợ .docx. (.doc cần LibreOffice, không phù hợp serverless)"})

    phrase_list = [line.strip() for line in (phrases or "").splitlines() if line.strip()]
    if not phrase_list:
        return JSONResponse(status_code=400, content={"error": "Bạn chưa nhập danh sách cụm từ."})

    return check_docx(
        data,
        phrase_list,
        case_sensitive=case_sensitive,
        whole_word=whole_word,
        scan_headers_footers=scan_headers_footers,
        spellcheck_vi=spellcheck_vi,
        spell_max_distance=spell_max_distance,
        check_du_toan_rule=check_du_toan_rule,
    )
