"""
Contract Review Agent — Public REST API
========================================
FastAPI wrapper cho agent.py.

Endpoints:
    POST /analyze   — Upload docx (+ eml tuỳ chọn), nhận file Excel
    GET  /health    — Health check
    GET  /docs      — Swagger UI (tự động)

Chạy local:
    uvicorn api:app --reload --port 8000

Biến môi trường:
    ANTHROPIC_API_KEY   (bắt buộc)
    API_SECRET_KEY      (tuỳ chọn — bearer token bảo vệ endpoint)
"""

from __future__ import annotations

import os
import tempfile
import uuid
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, Security, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.security import HTTPAuthorizationCredentials, HTTPBearer

# Import các hàm từ agent.py
from agent import collect_files, extract_contract, compare_email, export_excel


# ─── App setup ────────────────────────────────────────────────────────────────

app = FastAPI(
    title="Contract Review Agent — Zalopay FP&A",
    description=(
        "Upload hợp đồng Word (.docx) và email alignment (.eml), "
        "nhận file Excel với đầy đủ thông tin đã được trích xuất và phân tích."
    ),
    version="1.0.0",
    docs_url="/docs",
    redoc_url="/redoc",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["POST", "GET"],
    allow_headers=["*"],
)

security = HTTPBearer(auto_error=False)


# ─── Auth helper ──────────────────────────────────────────────────────────────

def verify_auth(credentials: HTTPAuthorizationCredentials | None) -> None:
    """Kiểm tra bearer token nếu API_SECRET_KEY được set."""
    secret = os.environ.get("API_SECRET_KEY", "")
    if not secret:
        return  # Không bắt buộc auth nếu chưa cấu hình
    if not credentials or credentials.credentials != secret:
        raise HTTPException(status_code=401, detail="Invalid or missing API key.")


# ─── Endpoints ────────────────────────────────────────────────────────────────

@app.get("/health", tags=["System"])
def health_check():
    """Kiểm tra trạng thái service."""
    api_key_ok = bool(os.environ.get("ANTHROPIC_API_KEY", ""))
    return {
        "status": "ok",
        "anthropic_key_configured": api_key_ok,
    }


@app.post(
    "/analyze",
    tags=["Contract Review"],
    summary="Trích xuất thông tin hợp đồng",
    response_description="File Excel chứa thông tin hợp đồng đã trích xuất",
)
async def analyze(
    contracts: list[UploadFile] = File(
        ...,
        description="Hợp đồng chính và các phụ lục (.docx). Có thể upload nhiều file.",
    ),
    emails: list[UploadFile] = File(
        default=[],
        description="(Tuỳ chọn) Email alignment (.eml) để so sánh với hợp đồng.",
    ),
    credentials: HTTPAuthorizationCredentials | None = Security(security),
):
    """
    Upload hợp đồng Word và (tuỳ chọn) email alignment để nhận file Excel phân tích.

    **Quy trình:**
    1. Đọc nội dung tất cả file .docx (hợp đồng + phụ lục)
    2. Gọi Claude API để trích xuất thông tin có cấu trúc
    3. Nếu có file .eml: so sánh email alignment với hợp đồng
    4. Xuất file Excel 4 sheet và trả về để download

    **Output Excel gồm:**
    - 📋 Tóm tắt HĐ — Loại HĐ, Các bên, Thời hạn, Dịch vụ, Tổng giá trị
    - 💰 Commercial Terms — Bảng phí, Ngân sách KM, Lãi & Phạt
    - 🏦 Payment & Recon — Cơ chế TT, Công nợ, Đối soát
    - 📧 Email vs HĐ — So sánh điểm khác biệt *(chỉ khi có file .eml)*
    """
    verify_auth(credentials)

    # Validate file types
    for f in contracts:
        if not (f.filename or "").lower().endswith(".docx"):
            raise HTTPException(400, f"File '{f.filename}' phải là .docx")
    for f in emails:
        if not (f.filename or "").lower().endswith(".eml"):
            raise HTTPException(400, f"File '{f.filename}' phải là .eml")

    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        raise HTTPException(500, "ANTHROPIC_API_KEY chưa được cấu hình trên server.")

    # Lưu upload vào temp dir
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)

        contract_paths: list[str] = []
        for upload in contracts:
            dest = tmp / (upload.filename or f"contract_{uuid.uuid4().hex}.docx")
            dest.write_bytes(await upload.read())
            contract_paths.append(str(dest))

        email_paths: list[str] = []
        for upload in emails:
            dest = tmp / (upload.filename or f"email_{uuid.uuid4().hex}.eml")
            dest.write_bytes(await upload.read())
            email_paths.append(str(dest))

        # Đọc nội dung
        contract_texts, email_texts = collect_files(contract_paths, email_paths, None)
        if not contract_texts:
            raise HTTPException(422, "Không đọc được nội dung từ các file hợp đồng.")

        # Phân tích
        try:
            data = extract_contract(contract_texts, api_key)
        except Exception as exc:
            raise HTTPException(502, f"Lỗi Claude API (extraction): {exc}") from exc

        comparison = None
        if email_texts:
            try:
                comparison = compare_email(email_texts, data, api_key)
            except Exception as exc:
                raise HTTPException(502, f"Lỗi Claude API (comparison): {exc}") from exc

        # Xuất Excel ra file tạm ngoài tmpdir (để không bị xoá trước khi gửi)
        output_path = tmp / "ContractReview.xlsx"
        export_excel(data, comparison, str(output_path), contract_paths)

        # Đọc vào memory trước khi tmpdir bị dọn
        excel_bytes = output_path.read_bytes()

    # Ghi ra file tạm tồn tại độc lập để FileResponse đọc được
    result_file = tempfile.NamedTemporaryFile(
        delete=False, suffix=".xlsx", prefix="ContractReview_"
    )
    result_file.write(excel_bytes)
    result_file.close()

    return FileResponse(
        path=result_file.name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="ContractReview.xlsx",
        background=None,
    )
