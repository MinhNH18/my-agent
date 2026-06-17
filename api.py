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
    LLM_API_KEY   (bắt buộc)
    API_SECRET_KEY      (tuỳ chọn — bearer token bảo vệ endpoint)
"""

from __future__ import annotations

import asyncio
import os
import tempfile
import uuid
from pathlib import Path

from fastapi import FastAPI, File, Header, HTTPException, Security, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, RedirectResponse
from fastapi.security import HTTPAuthorizationCredentials, HTTPBearer

# Import các hàm từ agent.py
from agent import collect_files, extract_contract, compare_email, export_excel

try:
    from greennode_agentbase.memory import MemoryClient
    from greennode_agentbase.memory.models import ChatMessage, EventCreateRequest
    _MEMORY_SDK_AVAILABLE = True
except ImportError:
    _MEMORY_SDK_AVAILABLE = False


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

# In-memory job store: job_id → {status, bytes?, error?, result_data?}
_jobs: dict[str, dict] = {}


# ─── Background analysis task ─────────────────────────────────────────────────

async def _run_analysis(
    job_id: str,
    contract_data: list[tuple[str, bytes]],
    email_data: list[tuple[str, bytes]],
    api_key: str,
    x_user_id: str | None,
    x_session_id: str | None,
) -> None:
    loop = asyncio.get_running_loop()
    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)

            contract_paths: list[str] = []
            for fname, data in contract_data:
                p = tmp / (fname or f"contract_{uuid.uuid4().hex}")
                p.write_bytes(data)
                contract_paths.append(str(p))

            email_paths: list[str] = []
            for fname, data in email_data:
                p = tmp / (fname or f"email_{uuid.uuid4().hex}")
                p.write_bytes(data)
                email_paths.append(str(p))

            contract_texts, email_texts = collect_files(contract_paths, email_paths, None)
            if not contract_texts:
                _jobs[job_id] = {"status": "error", "error": "Không đọc được nội dung từ các file hợp đồng."}
                return

            result_data = await loop.run_in_executor(None, extract_contract, contract_texts, api_key)

            comparison = None
            if email_texts:
                comparison = await loop.run_in_executor(None, compare_email, email_texts, result_data, api_key)

            output_path = tmp / "ContractReview.xlsx"
            export_excel(result_data, comparison, str(output_path), contract_paths)
            excel_bytes = output_path.read_bytes()

        _jobs[job_id] = {
            "status": "done",
            "bytes": excel_bytes,
            "result_data": result_data,
            "x_user_id": x_user_id,
            "x_session_id": x_session_id,
            "contract_fnames": [fn for fn, _ in contract_data],
        }
    except Exception as exc:
        _jobs[job_id] = {"status": "error", "error": str(exc)}


# ─── Memory helper ────────────────────────────────────────────────────────────

async def _log_to_memory(
    user_id: str,
    session_id: str,
    user_msg: str,
    assistant_msg: str,
) -> None:
    """Log a user/assistant event pair to AgentBase Memory (best-effort)."""
    memory_id = os.environ.get("MEMORY_ID", "")
    if not memory_id or not _MEMORY_SDK_AVAILABLE:
        return
    try:
        client = MemoryClient()
        import asyncio
        await asyncio.gather(
            client.create_event_async(
                id=memory_id,
                actorId=user_id,
                sessionId=session_id,
                request=EventCreateRequest(
                    payload=ChatMessage(role="user", content=user_msg)
                ),
            ),
            client.create_event_async(
                id=memory_id,
                actorId=user_id,
                sessionId=session_id,
                request=EventCreateRequest(
                    payload=ChatMessage(role="assistant", content=assistant_msg)
                ),
            ),
        )
    except Exception:
        pass  # best-effort — memory failure must not break the response


# ─── Auth helper ──────────────────────────────────────────────────────────────

def verify_auth(credentials: HTTPAuthorizationCredentials | None) -> None:
    """Kiểm tra bearer token nếu API_SECRET_KEY được set."""
    secret = os.environ.get("API_SECRET_KEY", "")
    if not secret:
        return  # Không bắt buộc auth nếu chưa cấu hình
    if not credentials or credentials.credentials != secret:
        raise HTTPException(status_code=401, detail="Invalid or missing API key.")


# ─── Endpoints ────────────────────────────────────────────────────────────────

_ZP_LOGO_SVG = '<svg width="108" height="24" viewBox="0 0 125 28" fill="none" xmlns="http://www.w3.org/2000/svg"><path fill-rule="evenodd" clip-rule="evenodd" d="M104.314 4.30811V17.2269H106.462V21.5445H104.935C103.203 21.5445 101.684 20.6532 100.8 19.3065C99.3434 20.6897 97.4172 21.5421 95.2961 21.5421C90.7666 21.5421 87.0919 17.6896 87.0919 12.936C87.0919 8.18252 90.7642 4.33003 95.2961 4.33003C97.0446 4.33003 98.6616 4.90717 99.9912 5.88368V4.30811H104.314ZM91.4046 12.9336C91.4046 15.3055 93.326 17.2269 95.6979 17.2269C98.0698 17.2269 99.9912 15.3055 99.9912 12.9336C99.9912 10.5617 98.0698 8.64034 95.6979 8.64034C93.326 8.64034 91.4046 10.5617 91.4046 12.9336ZM76.864 4.30822C72.1105 4.30822 68.2458 8.17532 68.2458 12.9264V28.0003H72.561V20.383C73.8298 21.1184 75.2958 21.5446 76.864 21.5446C81.6176 21.5446 85.4822 17.6775 85.4822 12.9264C85.4822 8.17532 81.6151 4.30822 76.864 4.30822ZM76.864 17.2294C74.4921 17.2294 72.561 15.2983 72.561 12.9264C72.561 10.5545 74.4921 8.6234 76.864 8.6234C79.2359 8.6234 81.1671 10.5545 81.1671 12.9264C81.1671 15.2983 79.2384 17.2294 76.864 17.2294ZM119.563 4.31303L114.973 15.0206L110.614 4.31303H105.953L112.584 20.5997L109.411 28.0003H114.108L124.26 4.31303H119.563Z" fill="#00CF6A"/><path fill-rule="evenodd" clip-rule="evenodd" d="M17.5863 1.22491C17.9637 1.97982 17.8834 2.86867 17.3768 3.54322L7.11487 17.2266H17.8712V21.5418H2.90927C2.06669 21.5418 1.30691 21.0743 0.92945 20.3194C0.551993 19.5644 0.632355 18.6756 1.13888 18.001L11.4033 4.31518H0.644531V0H15.604C16.449 0 17.2088 0.469994 17.5863 1.22491ZM47.8023 12.9262C47.8023 8.1751 51.6694 4.30799 56.4205 4.30799C61.1716 4.30799 65.0387 8.1751 65.0387 12.9262C65.0387 17.6773 61.1716 21.5444 56.4205 21.5444C51.6694 21.5444 47.8023 17.6773 47.8023 12.9262ZM52.1175 12.9262C52.1175 15.2981 54.0486 17.2292 56.4205 17.2292C58.7948 17.2292 60.7235 15.2981 60.7235 12.9262C60.7235 10.5543 58.7948 8.62318 56.4205 8.62318C54.0461 8.62318 52.1175 10.5543 52.1175 12.9262ZM36.7075 17.2268V4.30799H32.385V5.88357C31.0554 4.90705 29.4384 4.32991 27.6899 4.32991C23.158 4.32991 19.4857 8.18241 19.4857 12.9359C19.4857 17.6895 23.1605 21.5419 27.6899 21.5419C29.811 21.5419 31.7373 20.6896 33.1935 19.3064C34.0775 20.6531 35.5971 21.5444 37.3285 21.5444H38.8554V17.2268H36.7075ZM28.0917 17.2268C25.7199 17.2268 23.7985 15.3054 23.7985 12.9335C23.7985 10.5616 25.7199 8.64023 28.0917 8.64023C30.4636 8.64023 32.385 10.5616 32.385 12.9335C32.385 15.3054 30.4636 17.2268 28.0917 17.2268ZM44.8167 17.2218V0H40.4918V16.5935C40.4918 19.3258 42.7054 21.5418 45.4377 21.5418H46.9792V17.2218H44.8167Z" fill="#0033C9"/></svg>'

_UI_HTML = """<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Contract Intelligence — Zalopay FP&amp;A</title>
<style>
:root {
  --zp-blue:    #0033C9;
  --zp-blue-d:  #0029A3;
  --zp-blue-l:  #EBF0FF;
  --zp-blue-m:  #B3C3F5;
  --zp-green:   #00CF6A;
  --zp-green-d: #00A855;
  --zp-green-l: #E6FBF1;
  --zp-navy:    #001A6E;
  --zp-gray:    #6B7280;
  --zp-border:  #E5E7EB;
  --zp-bg:      #F4F6FB;
  --zp-success: #00875A;
  --zp-error:   #DE350B;
  --radius:     10px;
}
* { box-sizing: border-box; margin: 0; padding: 0; }
body {
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Inter", sans-serif;
  background: var(--zp-bg);
  min-height: 100vh;
  color: #111827;
}

/* ── Header ── */
.header {
  background: var(--zp-navy);
  padding: 0 32px;
  height: 60px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  position: sticky; top: 0; z-index: 10;
  border-bottom: 2px solid var(--zp-blue);
}
.header-brand { display: flex; align-items: center; gap: 16px; }
.header-logo { display: flex; align-items: center; }
.header-divider {
  width: 1px; height: 24px;
  background: rgba(255,255,255,.2);
}
.header-title {
  color: rgba(255,255,255,.75);
  font-size: 13px;
  font-weight: 500;
  letter-spacing: .2px;
}
.header-badge {
  background: rgba(0,207,106,.15);
  color: var(--zp-green);
  font-size: 10px;
  font-weight: 700;
  padding: 3px 8px;
  border-radius: 4px;
  letter-spacing: .6px;
  text-transform: uppercase;
  border: 1px solid rgba(0,207,106,.3);
}
.header-right { display: flex; align-items: center; gap: 20px; }
.header-link {
  color: rgba(255,255,255,.45);
  font-size: 13px;
  text-decoration: none;
  transition: color .15s;
}
.header-link:hover { color: rgba(255,255,255,.9); }

/* ── Layout ── */
.layout {
  max-width: 1020px;
  margin: 0 auto;
  padding: 36px 24px 60px;
  display: grid;
  grid-template-columns: 1fr 320px;
  gap: 24px;
  align-items: start;
}
@media (max-width: 720px) {
  .layout { grid-template-columns: 1fr; }
}

/* ── Hero ── */
.hero { margin-bottom: 24px; }
.hero h1 {
  font-size: 26px;
  font-weight: 700;
  color: var(--zp-navy);
  line-height: 1.3;
  margin-bottom: 8px;
}
.hero h1 span { color: var(--zp-blue); }
.hero p {
  color: var(--zp-gray);
  font-size: 15px;
  line-height: 1.6;
}

/* ── Card ── */
.card {
  background: #fff;
  border: 1px solid var(--zp-border);
  border-radius: var(--radius);
  padding: 28px;
  box-shadow: 0 1px 4px rgba(0,0,0,.06);
}
.card + .card { margin-top: 16px; }
.card-title {
  font-size: 13px;
  font-weight: 700;
  color: var(--zp-navy);
  text-transform: uppercase;
  letter-spacing: .6px;
  margin-bottom: 18px;
  padding-bottom: 12px;
  border-bottom: 2px solid var(--zp-blue-l);
  display: flex;
  align-items: center;
  gap: 8px;
}

/* ── Drop zone ── */
.field-label {
  font-size: 13px;
  font-weight: 600;
  color: #374151;
  margin-bottom: 6px;
  display: flex;
  align-items: center;
  gap: 6px;
}
.badge-opt {
  font-size: 11px;
  font-weight: 500;
  color: var(--zp-gray);
  background: #F3F4F6;
  border-radius: 4px;
  padding: 1px 6px;
}
.drop {
  border: 1.5px dashed var(--zp-blue-m);
  border-radius: var(--radius);
  padding: 22px 16px;
  text-align: center;
  cursor: pointer;
  background: var(--zp-blue-l);
  transition: border-color .15s, background .15s;
  position: relative;
}
.drop:hover, .drop.over {
  border-color: var(--zp-blue);
  background: #DCEbFF;
}
.drop input[type=file] { display: none; }
.drop-icon { font-size: 26px; margin-bottom: 6px; line-height: 1; }
.drop-text { font-size: 13px; color: var(--zp-gray); }
.drop-text b { color: var(--zp-blue); font-weight: 600; cursor: pointer; }
.drop-ext {
  margin-top: 4px;
  font-size: 11px;
  color: #9CA3AF;
  font-family: monospace;
  letter-spacing: .3px;
}
.file-list {
  list-style: none;
  margin-top: 8px;
}
.file-list li {
  display: flex;
  align-items: center;
  gap: 7px;
  padding: 6px 10px;
  background: #F9FAFB;
  border: 1px solid var(--zp-border);
  border-radius: 6px;
  margin-top: 4px;
  font-size: 13px;
  color: #374151;
}
.file-list li .fname { flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.file-list li .fsize { color: #9CA3AF; font-size: 12px; flex-shrink: 0; }
.file-icon-pdf  { color: #DE350B; }
.file-icon-docx { color: var(--zp-blue); }
.file-icon-eml  { color: #00875A; }
.file-icon-msg  { color: #7C3AED; }
.del-btn {
  flex-shrink: 0; background: none; border: none; cursor: pointer;
  font-size: 16px; line-height: 1; color: #9CA3AF; padding: 0 2px;
  border-radius: 4px; transition: color .15s, background .15s;
}
.del-btn:hover { color: #EF4444; background: #FEE2E2; }

/* ── Separator ── */
.sep { height: 1px; background: var(--zp-border); margin: 20px 0; }

/* ── Action button ── */
.btn-analyze {
  width: 100%;
  padding: 14px;
  background: linear-gradient(135deg, var(--zp-blue) 0%, #0044E0 100%);
  color: #fff;
  border: none;
  border-radius: var(--radius);
  font-size: 15px;
  font-weight: 700;
  cursor: pointer;
  transition: opacity .15s, box-shadow .15s;
  display: flex;
  align-items: center;
  justify-content: center;
  gap: 8px;
  letter-spacing: .2px;
  margin-top: 4px;
  box-shadow: 0 2px 8px rgba(0,51,201,.25);
}
.btn-analyze:hover:not(:disabled) {
  opacity: .9;
  box-shadow: 0 4px 16px rgba(0,51,201,.4);
}
.btn-analyze:disabled { background: #9CA3AF; cursor: not-allowed; box-shadow: none; }

/* ── Status ── */
.status-box {
  margin-top: 14px;
  border-radius: var(--radius);
  padding: 13px 16px;
  font-size: 14px;
  display: none;
}
.status-box.loading {
  display: flex;
  align-items: center;
  gap: 12px;
  background: var(--zp-blue-l);
  color: var(--zp-blue-d);
  border: 1px solid var(--zp-blue-m);
}
.status-box.error {
  display: block;
  background: #FFF5F3;
  color: var(--zp-error);
  border: 1px solid #FFBDAD;
}
.status-box.success {
  display: block;
  background: #E3FCEF;
  color: var(--zp-success);
  border: 1px solid #ABF5D1;
}
.spinner {
  width: 18px; height: 18px;
  border: 2.5px solid var(--zp-blue-m);
  border-top-color: var(--zp-blue);
  border-radius: 50%;
  animation: spin .75s linear infinite;
  flex-shrink: 0;
}
@keyframes spin { to { transform: rotate(360deg); } }
.progress-steps {
  margin-top: 10px;
  display: none;
}
.progress-steps.show { display: block; }
.step {
  display: flex; align-items: center; gap: 8px;
  font-size: 13px; color: #6B7280;
  padding: 3px 0;
  transition: color .2s;
}
.step.active { color: var(--zp-blue); font-weight: 600; }
.step.done   { color: var(--zp-green-d); }
.step-dot {
  width: 7px; height: 7px;
  border-radius: 50%;
  background: #D1D5DB;
  flex-shrink: 0;
  transition: background .2s;
}
.step.active .step-dot { background: var(--zp-blue); }
.step.done   .step-dot { background: var(--zp-green); }

/* ── Sidebar info ── */
.sidebar {}
.info-card {
  background: #fff;
  border: 1px solid var(--zp-border);
  border-radius: var(--radius);
  overflow: hidden;
  box-shadow: 0 1px 4px rgba(0,0,0,.06);
}
.info-card + .info-card { margin-top: 16px; }
.info-header {
  background: var(--zp-navy);
  color: rgba(255,255,255,.85);
  font-size: 11px;
  font-weight: 700;
  letter-spacing: .7px;
  text-transform: uppercase;
  padding: 10px 16px;
  border-left: 3px solid var(--zp-green);
}
.sheet-item {
  display: flex;
  align-items: flex-start;
  gap: 10px;
  padding: 12px 16px;
  border-bottom: 1px solid var(--zp-border);
}
.sheet-item:last-child { border-bottom: none; }
.sheet-num {
  width: 22px; height: 22px;
  background: var(--zp-blue);
  color: #fff;
  font-size: 11px;
  font-weight: 800;
  border-radius: 4px;
  display: flex; align-items: center; justify-content: center;
  flex-shrink: 0;
  margin-top: 1px;
}
.sheet-name {
  font-size: 13px;
  font-weight: 600;
  color: var(--zp-navy);
  margin-bottom: 2px;
}
.sheet-desc {
  font-size: 12px;
  color: var(--zp-gray);
  line-height: 1.5;
}
.tip-item {
  display: flex;
  align-items: flex-start;
  gap: 9px;
  padding: 10px 16px;
  border-bottom: 1px solid var(--zp-border);
  font-size: 13px;
  color: #374151;
  line-height: 1.5;
}
.tip-item:last-child { border-bottom: none; }
.tip-icon { font-size: 15px; flex-shrink: 0; margin-top: 1px; }
.footer-note {
  text-align: center;
  margin-top: 14px;
  font-size: 12px;
  color: #9CA3AF;
}
.footer-note a { color: var(--zp-blue); text-decoration: none; }
</style>
</head>
<body>

<!-- Header -->
<header class="header">
  <div class="header-brand">
    <div class="header-logo">""" + _ZP_LOGO_SVG.replace("#0033C9", "#ffffff").replace("#00CF6A", "#00CF6A") + """</div>
    <div class="header-divider"></div>
    <span class="header-title">Contract Intelligence</span>
    <span class="header-badge">FP&amp;A Tool</span>
  </div>
  <div class="header-right">
    <a class="header-link" href="/docs" target="_blank">API Docs</a>
    <a class="header-link" href="/health">Status</a>
  </div>
</header>

<!-- Main -->
<div class="layout">
  <!-- Left column -->
  <div>
    <div class="hero">
      <h1>Phân tích hợp đồng <span>tự động</span></h1>
      <p>Upload hợp đồng Word / PDF, nhận ngay file Excel chuẩn FP&amp;A với đầy đủ các điều khoản thương mại, lịch thanh toán và đối soát.</p>
    </div>

    <!-- Upload card -->
    <div class="card">
      <div class="card-title">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
        Upload tài liệu
      </div>

      <!-- Contracts -->
      <div class="field-label">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="var(--zp-blue)" stroke-width="2.5"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
        Hợp đồng &amp; Phụ lục
      </div>
      <div class="drop" id="contractDrop"
           onclick="document.getElementById('contractInput').click()"
           ondragover="dnd(event,'over')" ondragleave="dnd(event,'leave')" ondrop="dnd(event,'drop','contractInput')">
        <input type="file" id="contractInput" accept=".docx,.pdf" multiple
               onchange="onFileChange('contractInput')">
        <div class="drop-icon">📂</div>
        <div class="drop-text">Kéo thả vào đây hoặc <b>chọn từ máy tính</b></div>
        <div class="drop-ext">.docx &nbsp;·&nbsp; .pdf &nbsp;·&nbsp; nhiều file</div>
      </div>
      <ul class="file-list" id="contractList"></ul>

      <div class="sep"></div>

      <!-- Emails -->
      <div class="field-label">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#6B7280" stroke-width="2.5"><path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"/><polyline points="22,6 12,13 2,6"/></svg>
        Email alignment
        <span class="badge-opt">tuỳ chọn</span>
      </div>
      <div class="drop" id="emailDrop"
           onclick="document.getElementById('emailInput').click()"
           ondragover="dnd(event,'over')" ondragleave="dnd(event,'leave')" ondrop="dnd(event,'drop','emailInput')">
        <input type="file" id="emailInput" accept=".eml,.msg" multiple
               onchange="onFileChange('emailInput')">
        <div class="drop-icon">📧</div>
        <div class="drop-text">Kéo thả hoặc <b>chọn file email</b></div>
        <div class="drop-ext">.eml &nbsp;·&nbsp; .msg (Outlook)</div>
      </div>
      <ul class="file-list" id="emailList"></ul>

      <div class="sep"></div>

      <button class="btn-analyze" id="btn" onclick="analyze()">
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
        Chạy phân tích
      </button>

      <!-- Status -->
      <div class="status-box loading" id="statusLoading">
        <div class="spinner"></div>
        <div>
          <div style="font-weight:600;margin-bottom:2px">Đang xử lý tài liệu…</div>
          <div style="font-size:12px;opacity:.8">AI đang đọc và trích xuất thông tin, vui lòng chờ</div>
        </div>
      </div>
      <div class="progress-steps" id="progressSteps">
        <div class="step" id="s1"><div class="step-dot"></div>Đọc nội dung tài liệu</div>
        <div class="step" id="s2"><div class="step-dot"></div>Trích xuất điều khoản thương mại</div>
        <div class="step" id="s3"><div class="step-dot"></div>Phân tích payment &amp; đối soát</div>
        <div class="step" id="s4"><div class="step-dot"></div>Xuất file Excel</div>
      </div>
      <div class="status-box error"   id="statusError"></div>
      <div class="status-box success" id="statusSuccess"></div>
    </div>
  </div>

  <!-- Right sidebar -->
  <div class="sidebar">
    <div class="info-card">
      <div class="info-header">Output Excel — 4 Sheets</div>
      <div class="sheet-item">
        <div class="sheet-num">1</div>
        <div>
          <div class="sheet-name">Tóm tắt Hợp đồng</div>
          <div class="sheet-desc">Loại HĐ, các bên, ngày ký, thời hạn, dịch vụ hợp tác, kênh thanh toán</div>
        </div>
      </div>
      <div class="sheet-item">
        <div class="sheet-num">2</div>
        <div>
          <div class="sheet-name">Commercial Terms</div>
          <div class="sheet-desc">Bảng phí theo kênh / nguồn tiền, ngân sách khuyến mãi, lãi &amp; phạt vi phạm</div>
        </div>
      </div>
      <div class="sheet-item">
        <div class="sheet-num">3</div>
        <div>
          <div class="sheet-name">Payment &amp; Reconciliation</div>
          <div class="sheet-desc">Cơ chế thanh toán, công nợ, hồ sơ thanh toán, đối soát, xuất hóa đơn</div>
        </div>
      </div>
      <div class="sheet-item">
        <div class="sheet-num">4</div>
        <div>
          <div class="sheet-name">Email vs Hợp đồng</div>
          <div class="sheet-desc">So sánh điểm khác biệt giữa email alignment và hợp đồng ký kết <em style="color:#9CA3AF">(chỉ khi có file email)</em></div>
        </div>
      </div>
    </div>

    <div class="info-card" style="margin-top:16px">
      <div class="info-header">Lưu ý sử dụng</div>
      <div class="tip-item"><span class="tip-icon">📋</span>Hỗ trợ hợp đồng chính + nhiều phụ lục cùng lúc</div>
      <div class="tip-item"><span class="tip-icon">⏱</span>Thời gian xử lý ~30–50 giây tuỳ độ dài HĐ</div>
      <div class="tip-item"><span class="tip-icon">🔒</span>File không được lưu trên server sau khi phân tích xong</div>
      <div class="tip-item"><span class="tip-icon">⚠️</span>Kiểm tra lại các ô màu đỏ trong Excel — đây là trường AI chưa tìm thấy</div>
    </div>

    <p class="footer-note">""" + _ZP_LOGO_SVG.replace('width="108"', 'width="72"').replace('height="24"', 'height="16"').replace("#0033C9", "#9CA3AF").replace("#00CF6A", "#9CA3AF") + """ &nbsp;FP&amp;A &nbsp;·&nbsp; <a href="/docs" target="_blank">API Docs</a></p>
  </div>
</div>

<script>
const FILE_ICONS = { pdf:'📄', docx:'📝', eml:'📩', msg:'📨' };
const PAIR = { contractInput: 'contractList', emailInput: 'emailList' };

/* Central store: inputId → File[] */
const _store = { contractInput: [], emailInput: [] };

/* Sync store → input.files */
function syncInput(inputId) {
  const dt = new DataTransfer();
  _store[inputId].forEach(f => dt.items.add(f));
  document.getElementById(inputId).files = dt.files;
}

/* Render list with × delete buttons */
function renderList(inputId) {
  const files = _store[inputId];
  const list  = document.getElementById(PAIR[inputId]);
  if (!files.length) { list.innerHTML = ''; return; }
  list.innerHTML = files.map((f, i) => {
    const ext  = f.name.split('.').pop().toLowerCase();
    const icon = FILE_ICONS[ext] || '📄';
    const sz   = f.size < 1024 ? f.size + ' B'
               : f.size < 1048576 ? (f.size/1024).toFixed(1) + ' KB'
               : (f.size/1048576).toFixed(2) + ' MB';
    return `<li>
      <span class="file-icon-${ext}" style="font-size:15px">${icon}</span>
      <span class="fname" title="${f.name}">${f.name}</span>
      <span class="fsize">${sz}</span>
      <button class="del-btn" onclick="removeFile('${inputId}',${i})" title="Xoá">×</button>
    </li>`;
  }).join('');
}

/* Add files (merge, deduplicate by name) */
function addFiles(inputId, newFiles) {
  const names = new Set(_store[inputId].map(f => f.name));
  Array.from(newFiles).forEach(f => {
    if (!names.has(f.name)) { _store[inputId].push(f); names.add(f.name); }
  });
  syncInput(inputId);
  renderList(inputId);
}

/* Remove one file by index */
function removeFile(inputId, idx) {
  _store[inputId].splice(idx, 1);
  syncInput(inputId);
  renderList(inputId);
}

/* onchange handler (called by input element) */
function onFileChange(inputId) {
  const input = document.getElementById(inputId);
  addFiles(inputId, input.files);
}

/* Drag & drop */
function dnd(e, action, inputId) {
  e.preventDefault(); e.stopPropagation();
  const drop = e.currentTarget;
  if (action === 'over')  { drop.classList.add('over'); return; }
  if (action === 'leave') { drop.classList.remove('over'); return; }
  if (action === 'drop')  { drop.classList.remove('over'); addFiles(inputId, e.dataTransfer.files); }
}

let _stepTimer = null;
function startSteps() {
  const steps = ['s1','s2','s3','s4'];
  let i = 0;
  document.querySelectorAll('.step').forEach(s => s.className = 'step');
  document.getElementById('progressSteps').classList.add('show');
  setStep(0);
  _stepTimer = setInterval(() => {
    if (i < steps.length - 1) { i++; setStep(i); }
  }, 9000);
}
function setStep(i) {
  const steps = ['s1','s2','s3','s4'];
  steps.forEach((id,j) => {
    const el = document.getElementById(id);
    el.className = 'step' + (j < i ? ' done' : j === i ? ' active' : '');
  });
}
function stopSteps(success) {
  clearInterval(_stepTimer);
  if (success) {
    document.querySelectorAll('.step').forEach(s => s.className = 'step done');
  }
  document.getElementById('progressSteps').classList.remove('show');
}

function show(id, msg) {
  const el = document.getElementById(id);
  el.style.display = id === 'statusLoading' ? 'flex' : 'block';
  if (msg !== undefined) el.textContent = msg;
}
function hide(id) { document.getElementById(id).style.display = 'none'; }

async function analyze() {
  const contracts = document.getElementById('contractInput').files;
  if (!contracts.length) {
    show('statusError', '⚠️  Vui lòng chọn ít nhất 1 file hợp đồng (.docx hoặc .pdf)');
    return;
  }
  const btn = document.getElementById('btn');
  btn.disabled = true;
  hide('statusError'); hide('statusSuccess');
  show('statusLoading');
  startSteps();

  const fd = new FormData();
  Array.from(contracts).forEach(f => fd.append('contracts', f));
  Array.from(document.getElementById('emailInput').files).forEach(f => fd.append('emails', f));

  const done = (ok, msg) => {
    btn.disabled = false;
    hide('statusLoading');
    stopSteps(ok);
    if (ok) show('statusSuccess', msg);
    else     show('statusError',   msg);
  };

  try {
    /* Bước 1: submit job — nhận job_id ngay lập tức */
    const submitRes = await fetch('/analyze', { method: 'POST', body: fd });
    if (!submitRes.ok) {
      const txt = await submitRes.text();
      let msg = txt;
      try { msg = JSON.parse(txt).detail || txt; } catch {}
      done(false, '❌  ' + msg);
      return;
    }
    const { job_id } = await submitRes.json();

    /* Bước 2: poll /result/{job_id} mỗi 3 giây (tối đa 5 phút) */
    for (let i = 0; i < 100; i++) {
      await new Promise(r => setTimeout(r, 3000));
      const res = await fetch('/result/' + job_id);
      if (res.status === 202) continue;   /* đang xử lý */

      if (!res.ok) {
        const txt = await res.text();
        let msg = txt;
        try { msg = JSON.parse(txt).detail || txt; } catch {}
        done(false, '❌  ' + msg);
        return;
      }

      /* 200 — tải file */
      const blob = await res.blob();
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement('a');
      a.href = url; a.download = 'ContractReview.xlsx'; a.click();
      URL.revokeObjectURL(url);
      done(true, '✅  Phân tích thành công! File ContractReview.xlsx đã được tải về máy.');
      return;
    }
    done(false, '❌  Timeout: phân tích mất quá lâu, vui lòng thử lại.');
  } catch (err) {
    done(false, '❌  Lỗi kết nối: ' + err.message);
  }
}
</script>
</body>
</html>"""


@app.get("/", include_in_schema=False)
def root():
    return HTMLResponse(_UI_HTML)


@app.get("/health", tags=["System"])
def health_check():
    """Kiểm tra trạng thái service."""
    api_key_ok = bool(os.environ.get("LLM_API_KEY", ""))
    return {
        "status": "ok",
        "anthropic_key_configured": api_key_ok,
    }


@app.post(
    "/analyze",
    tags=["Contract Review"],
    summary="Trích xuất thông tin hợp đồng (async)",
    response_description="Job ID để poll kết quả tại GET /result/{job_id}",
    status_code=202,
)
async def analyze(
    contracts: list[UploadFile] = File(
        ...,
        description="Hợp đồng chính và các phụ lục (.docx / .pdf). Có thể upload nhiều file.",
    ),
    emails: list[UploadFile] = File(
        default=[],
        description="(Tuỳ chọn) Email alignment (.eml / .msg) để so sánh với hợp đồng.",
    ),
    credentials: HTTPAuthorizationCredentials | None = Security(security),
    x_user_id: str | None = Header(default=None, alias="X-GreenNode-AgentBase-User-Id"),
    x_session_id: str | None = Header(default=None, alias="X-GreenNode-AgentBase-Session-Id"),
):
    """
    Bắt đầu phân tích hợp đồng trong nền. Trả về `job_id` ngay lập tức (202).
    Poll kết quả tại **GET /result/{job_id}** — trả 202 khi đang xử lý, 200 + file Excel khi xong.
    """
    verify_auth(credentials)

    for f in contracts:
        if not (f.filename or "").lower().endswith((".docx", ".pdf")):
            raise HTTPException(400, f"File '{f.filename}' phải là .docx hoặc .pdf")
    for f in emails:
        if not (f.filename or "").lower().endswith((".eml", ".msg")):
            raise HTTPException(400, f"File '{f.filename}' phải là .eml hoặc .msg")

    api_key = os.environ.get("LLM_API_KEY", "")
    if not api_key:
        raise HTTPException(500, "LLM_API_KEY chưa được cấu hình trên server.")

    # Đọc bytes trước khi UploadFile bị đóng
    contract_data = [(f.filename or f"contract_{uuid.uuid4().hex}.docx", await f.read()) for f in contracts]
    email_data    = [(f.filename or f"email_{uuid.uuid4().hex}.eml",     await f.read()) for f in emails]

    job_id = str(uuid.uuid4())
    _jobs[job_id] = {"status": "processing"}

    asyncio.create_task(_run_analysis(job_id, contract_data, email_data, api_key, x_user_id, x_session_id))

    return JSONResponse({"job_id": job_id, "status": "processing"}, status_code=202)


@app.get(
    "/result/{job_id}",
    tags=["Contract Review"],
    summary="Lấy kết quả phân tích",
)
async def get_result(job_id: str):
    """
    Poll kết quả của job. Trả **202** khi đang xử lý, **200 + file Excel** khi xong, **500** nếu lỗi.
    """
    job = _jobs.get(job_id)
    if not job:
        raise HTTPException(404, "Job không tồn tại hoặc đã hết hạn.")

    if job["status"] == "processing":
        return JSONResponse({"status": "processing"}, status_code=202)

    if job["status"] == "error":
        error = job.get("error", "Unknown error")
        del _jobs[job_id]
        raise HTTPException(500, f"Lỗi phân tích: {error}")

    # Done — trả file và log memory
    excel_bytes  = job["bytes"]
    result_data  = job.get("result_data", {})
    x_user_id    = job.get("x_user_id")
    x_session_id = job.get("x_session_id")
    fnames       = job.get("contract_fnames", [])
    del _jobs[job_id]

    if x_user_id:
        parties = "; ".join(
            b.get("ten_cong_ty", "") for b in (result_data.get("cac_ben") or []) if b.get("ten_cong_ty")
        )
        missing = ", ".join(result_data.get("truong_con_thieu") or []) or "Không có"
        summary = (
            f"Loại HĐ: {result_data.get('loai_hop_dong') or 'N/A'}\n"
            f"Các bên: {parties or 'N/A'}\n"
            f"Ngày ký: {(result_data.get('thoi_han_hop_dong') or {}).get('ngay_ky') or 'N/A'}\n"
            f"Hết hạn: {(result_data.get('thoi_han_hop_dong') or {}).get('ngay_het_han') or 'N/A'}\n"
            f"Trường thiếu: {missing}"
        )
        await _log_to_memory(
            user_id=x_user_id,
            session_id=x_session_id or str(uuid.uuid4()),
            user_msg=f"Phân tích hợp đồng: {', '.join(fnames)}",
            assistant_msg=summary,
        )

    result_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", prefix="ContractReview_")
    result_file.write(excel_bytes)
    result_file.close()

    return FileResponse(
        path=result_file.name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="ContractReview.xlsx",
    )
