# Word Phrase Checker (DOCX)

## Chức năng
- Nhập danh sách cụm từ (mỗi dòng 1 cụm)
- Upload file .docx
- Trả về:
  - cụm từ thiếu
  - cụm từ tìm thấy + vị trí (ước lượng trang/dòng), paragraph/table/header/footer, ngữ cảnh

## Chạy project
```bash
cd word-checker-app
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS/Linux:
# source .venv/bin/activate

pip install -r requirements.txt
uvicorn app.main:app --reload

python -m uvicorn app.main:app --reload