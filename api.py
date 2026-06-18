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
import email.header
import json
import os
import sqlite3
import tempfile
import uuid
from datetime import datetime
from pathlib import Path

from fastapi import FastAPI, File, Form, Header, HTTPException, Security, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, RedirectResponse
from fastapi.security import HTTPAuthorizationCredentials, HTTPBearer

# Import các hàm từ agent.py
from agent import collect_files, extract_contract, compare_email, export_excel, answer_custom_queries, chat_with_contract, group_contract_files, identify_email_target

try:
    from greennode_agentbase.memory import MemoryClient
    from greennode_agentbase.memory.models import ChatMessage, EventCreateRequest
    _MEMORY_SDK_AVAILABLE = True
except ImportError:
    _MEMORY_SDK_AVAILABLE = False


# ─── History DB ───────────────────────────────────────────────────────────────

_DB_PATH = os.environ.get("HISTORY_DB", "/tmp/contract_history.db")


def _decode_filename(raw: str | None) -> str:
    """Decode filename that may be RFC 2047 encoded or UTF-8 bytes misread as latin-1."""
    if not raw:
        return ""
    # RFC 2047: =?utf-8?B?...?= or =?utf-8?Q?...?=
    if "=?" in raw:
        try:
            parts = email.header.decode_header(raw)
            decoded = ""
            for part, enc in parts:
                if isinstance(part, bytes):
                    decoded += part.decode(enc or "utf-8", errors="replace")
                else:
                    decoded += part
            cleaned = decoded.strip()
            if cleaned:
                return cleaned
        except Exception:
            pass
    # UTF-8 bytes misinterpreted as latin-1 (common in multipart/form-data from browsers/.NET)
    try:
        return raw.encode("latin-1").decode("utf-8")
    except (UnicodeDecodeError, UnicodeEncodeError):
        return raw


def _db_init() -> None:
    with sqlite3.connect(_DB_PATH) as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS reviews (
                id            TEXT PRIMARY KEY,
                created_at    TEXT NOT NULL,
                filenames     TEXT NOT NULL,
                results_json  TEXT NOT NULL,
                comparison_json TEXT,
                custom_answers_json TEXT,
                excel_bytes   BLOB
            )
        """)


def _db_save(
    review_id: str,
    filenames: list[str],
    results: list[dict],
    comparison: dict | None,
    custom_answers: list[dict] | None,
    excel_bytes: bytes,
) -> None:
    with sqlite3.connect(_DB_PATH) as conn:
        conn.execute(
            "INSERT OR REPLACE INTO reviews VALUES (?,?,?,?,?,?,?)",
            (
                review_id,
                datetime.utcnow().isoformat(),
                json.dumps(filenames, ensure_ascii=False),
                json.dumps(results,   ensure_ascii=False),
                json.dumps(comparison, ensure_ascii=False) if comparison else None,
                json.dumps(custom_answers, ensure_ascii=False) if custom_answers else None,
                excel_bytes,
            ),
        )


def _db_list(limit: int = 30) -> list[dict]:
    try:
        with sqlite3.connect(_DB_PATH) as conn:
            rows = conn.execute(
                "SELECT id, created_at, filenames FROM reviews ORDER BY created_at DESC LIMIT ?",
                (limit,),
            ).fetchall()
        return [{"id": r[0], "created_at": r[1], "filenames": json.loads(r[2])} for r in rows]
    except Exception:
        return []


def _db_get(review_id: str) -> dict | None:
    try:
        with sqlite3.connect(_DB_PATH) as conn:
            row = conn.execute(
                "SELECT id, created_at, filenames, results_json, comparison_json, custom_answers_json "
                "FROM reviews WHERE id=?",
                (review_id,),
            ).fetchone()
        if not row:
            return None
        return {
            "id": row[0],
            "created_at": row[1],
            "filenames": json.loads(row[2]),
            "results": json.loads(row[3]),
            "comparison": json.loads(row[4]) if row[4] else None,
            "custom_answers": json.loads(row[5]) if row[5] else None,
        }
    except Exception:
        return None


def _db_get_excel(review_id: str) -> bytes | None:
    try:
        with sqlite3.connect(_DB_PATH) as conn:
            row = conn.execute(
                "SELECT excel_bytes FROM reviews WHERE id=?", (review_id,)
            ).fetchone()
        return row[0] if row else None
    except Exception:
        return None


def _db_delete(review_id: str) -> bool:
    try:
        with sqlite3.connect(_DB_PATH) as conn:
            conn.execute("DELETE FROM reviews WHERE id=?", (review_id,))
        return True
    except Exception:
        return False


def _db_delete_all() -> int:
    try:
        with sqlite3.connect(_DB_PATH) as conn:
            cur = conn.execute("DELETE FROM reviews")
            return cur.rowcount
    except Exception:
        return 0


def _db_get_learning_example() -> dict | None:
    """Lấy kết quả trích xuất gần nhất làm few-shot example cho lần review tiếp theo."""
    try:
        with sqlite3.connect(_DB_PATH) as conn:
            row = conn.execute(
                "SELECT results_json FROM reviews ORDER BY created_at DESC LIMIT 1"
            ).fetchone()
        if not row:
            return None
        results = json.loads(row[0])
        if results:
            example = results[0]
            # Chỉ giữ phần cốt lõi, bỏ truong_con_thieu và _refs để tránh bias
            example.pop("truong_con_thieu", None)
            example.pop("_refs", None)
            return example
        return None
    except Exception:
        return None


_db_init()


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
    custom_queries: list[str] | None = None,
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

            fnames  = list(contract_texts.keys())
            learning_example = _db_get_learning_example()

            # Nhóm HĐ + phụ lục thành các logical documents
            groups = await loop.run_in_executor(
                None, group_contract_files, fnames, api_key
            )

            results: list[dict] = []
            group_labels: list[str] = []
            for group in groups:
                group_texts = {f: contract_texts[f] for f in group["files"] if f in contract_texts}
                if not group_texts:
                    continue
                r = await loop.run_in_executor(
                    None, extract_contract, group_texts, api_key, learning_example
                )
                results.append(r)
                group_labels.append(group["group_name"])

            # Dùng group labels làm fnames để hiển thị đúng tab
            if group_labels:
                fnames = group_labels

            comparison = None
            if email_texts:
                # Xác định email đang nói về nhóm nào
                matched_names = await loop.run_in_executor(
                    None, identify_email_target, email_texts, groups, api_key
                )
                # Tìm index group phù hợp nhất
                target_idx = 0
                for i, g in enumerate(groups):
                    if g["group_name"] in matched_names:
                        target_idx = i
                        break
                comparison = await loop.run_in_executor(
                    None, compare_email, email_texts, results[target_idx], api_key
                )

            custom_answers: list[dict] = []
            if custom_queries:
                custom_answers = await loop.run_in_executor(
                    None, answer_custom_queries, contract_texts, custom_queries, api_key
                )

            output_path = tmp / "ContractReview.xlsx"
            export_excel(results, comparison, str(output_path), fnames)
            excel_bytes = output_path.read_bytes()

        # Lưu vào lịch sử
        review_id = job_id
        _db_save(review_id, fnames, results, comparison, custom_answers or None, excel_bytes)

        _jobs[job_id] = {
            "status": "done",
            "review_id": review_id,
            "results":   results,
            "comparison": comparison,
            "custom_answers": custom_answers,
            "x_user_id":    x_user_id,
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
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Contract Intelligence — Zalopay FP&amp;A</title>
<style>
:root{
  /* Zalopay brand — green primary, blue accent */
  --zb:#0068FF;--zbd:#0052CC;--zbl:#E8F1FF;--zbm:#99C0FF;
  --zg:#06C755;--zgd:#04A847;--zgl:#E8FAF0;
  --zn:#001A3C;--zgr:#6B7280;--zbo:#E5E7EB;--zbg:#F5F7FA;
  --zok:#06C755;--zer:#DE350B;--rad:10px;
  --zp:#06C755;--zpd:#04A847;--zpl:#E8FAF0;
}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Inter,sans-serif;background:var(--zbg);color:#111827;min-height:100vh}
.hdr{background:var(--zg);padding:0 28px;height:56px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:20;box-shadow:0 2px 10px rgba(6,199,85,.3)}
.hdr-l{display:flex;align-items:center;gap:14px}
.hdr-div{width:1px;height:22px;background:rgba(255,255,255,.35)}
.hdr-t{color:rgba(255,255,255,.85);font-size:13px;font-weight:500}
.hdr-badge{background:rgba(255,255,255,.2);color:#fff;font-size:10px;font-weight:700;padding:3px 8px;border-radius:4px;letter-spacing:.5px;text-transform:uppercase;border:1px solid rgba(255,255,255,.35)}
.hdr-r{display:flex;align-items:center;gap:18px}
.hdr-a{color:rgba(255,255,255,.7);font-size:13px;text-decoration:none;transition:color .15s}
.hdr-a:hover{color:#fff}
.wrap{max-width:1100px;margin:0 auto;padding:28px 20px 60px;display:grid;grid-template-columns:1fr 300px;gap:22px;align-items:start}
@media(max-width:740px){.wrap{grid-template-columns:1fr}}
.card{background:#fff;border:1px solid var(--zbo);border-radius:var(--rad);padding:24px;box-shadow:0 1px 4px rgba(0,0,0,.06)}
.card+.card{margin-top:14px}
.card-t{font-size:12px;font-weight:700;color:var(--zn);text-transform:uppercase;letter-spacing:.6px;margin-bottom:16px;padding-bottom:10px;border-bottom:2px solid var(--zgl);display:flex;align-items:center;gap:7px}
/* upload */
.flbl{font-size:13px;font-weight:600;color:#374151;margin-bottom:6px;display:flex;align-items:center;gap:6px}
.opt{font-size:11px;color:var(--zgr);background:#F3F4F6;border-radius:4px;padding:1px 6px}
.drop{display:block;border:1.5px dashed #8ADCAA;border-radius:var(--rad);padding:18px 14px;text-align:center;cursor:pointer;background:var(--zgl);transition:border-color .15s,background .15s;position:relative}
.drop:hover,.drop.over{border-color:var(--zg);background:#D0F5E2}
.drop input[type=file]{display:none}
.drop-ic{font-size:24px;margin-bottom:4px}
.drop-tx{font-size:13px;color:var(--zgr)}
.drop-tx b{color:var(--zgd);font-weight:600;cursor:pointer}
.drop-ex{margin-top:3px;font-size:11px;color:#9CA3AF;font-family:monospace}
.flist{list-style:none;margin-top:7px}
.flist li{display:flex;align-items:center;gap:6px;padding:5px 9px;background:#F9FAFB;border:1px solid var(--zbo);border-radius:6px;margin-top:3px;font-size:13px}
.flist .fn{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.flist .fsz{color:#9CA3AF;font-size:12px;flex-shrink:0}
.del{flex-shrink:0;background:none;border:none;cursor:pointer;font-size:16px;color:#9CA3AF;padding:0 2px;border-radius:4px;transition:color .15s,background .15s}
.del:hover{color:#EF4444;background:#FEE2E2}
.sep{height:1px;background:var(--zbo);margin:18px 0}
/* custom query */
.qbox{width:100%;border:1.5px solid var(--zbo);border-radius:8px;padding:10px 12px;font-size:13px;font-family:inherit;resize:vertical;min-height:72px;transition:border-color .15s;color:#111827}
.qbox:focus{outline:none;border-color:var(--zg);box-shadow:0 0 0 3px rgba(6,199,85,.1)}
.qbox::placeholder{color:#9CA3AF}
/* btn */
.btn-run{width:100%;padding:13px;background:linear-gradient(135deg,var(--zg) 0%,#04A847 100%);color:#fff;border:none;border-radius:var(--rad);font-size:15px;font-weight:700;cursor:pointer;transition:opacity .15s,box-shadow .15s;display:flex;align-items:center;justify-content:center;gap:8px;letter-spacing:.2px;margin-top:4px;box-shadow:0 2px 10px rgba(6,199,85,.35)}
.btn-run:hover:not(:disabled){opacity:.9;box-shadow:0 4px 16px rgba(0,51,201,.4)}
.btn-run:disabled{background:#9CA3AF;cursor:not-allowed;box-shadow:none}
.btn-dl{display:inline-flex;align-items:center;gap:7px;padding:9px 18px;background:var(--zok);color:#fff;border:none;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer;text-decoration:none;transition:opacity .15s}
.btn-dl:hover{opacity:.85}
.btn-back{display:inline-flex;align-items:center;gap:6px;padding:9px 16px;background:#F3F4F6;color:#374151;border:none;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer;transition:background .15s}
.btn-back:hover{background:var(--zbo)}
/* status */
.sbox{margin-top:12px;border-radius:var(--rad);padding:12px 15px;font-size:14px;display:none}
.sbox.loading{display:flex;align-items:center;gap:11px;background:var(--zgl);color:var(--zgd);border:1px solid #8ADCAA}
.sbox.error{display:block;background:#FFF5F3;color:var(--zer);border:1px solid #FFBDAD}
.spin{width:17px;height:17px;border:2.5px solid #8ADCAA;border-top-color:var(--zg);border-radius:50%;animation:spin .75s linear infinite;flex-shrink:0}
@keyframes spin{to{transform:rotate(360deg)}}
/* results */
.res-hdr{display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:10px;margin-bottom:18px}
.res-title{font-size:20px;font-weight:700;color:var(--zn)}
.res-actions{display:flex;gap:8px;flex-wrap:wrap}
.contract-tabs{display:flex;gap:4px;flex-wrap:wrap;margin-bottom:16px;border-bottom:2px solid var(--zbo);padding-bottom:0}
.ctab{padding:8px 16px;border:1px solid var(--zbo);border-bottom:none;border-radius:8px 8px 0 0;font-size:13px;font-weight:600;cursor:pointer;background:#F9FAFB;color:var(--zgr);transition:background .15s,color .15s;margin-bottom:-2px}
.ctab.active{background:#fff;color:var(--zgd);border-color:var(--zbo);border-bottom-color:#fff;box-shadow:0 -2px 0 0 var(--zg)}
.cpanel{display:none}
.cpanel.active{display:block}
.missing-banner{background:#FFFBEB;border:1px solid #FCD34D;color:#92400E;border-radius:8px;padding:10px 14px;font-size:13px;margin-bottom:14px;font-weight:500}
/* sections (accordion) */
.sec{border:1px solid var(--zbo);border-radius:8px;margin-bottom:10px;overflow:hidden}
.sec-hd{display:flex;align-items:center;justify-content:space-between;padding:12px 16px;background:#F9FAFB;cursor:pointer;font-size:13px;font-weight:700;color:var(--zn);user-select:none;transition:background .15s}
.sec-hd:hover{background:var(--zgl)}
.sec-hd .caret{font-size:11px;color:var(--zgr);transition:transform .2s}
.sec.open .sec-hd{background:var(--zgl)}
.sec.open .caret{transform:rotate(90deg)}
.sec-bd{display:none;padding:0}
.sec.open .sec-bd{display:block}
/* data table */
.dtable{width:100%;border-collapse:collapse;font-size:13px}
.dtable td{padding:11px 14px;border-bottom:1px solid var(--zbo);vertical-align:top;line-height:1.6}
.dtable tr:last-child td{border-bottom:none}
.td-lbl{width:185px;font-weight:600;color:#374151;background:#FAFAFA;white-space:nowrap;font-size:12px}
.td-val{color:#111827;word-break:break-word}
.td-val.null{color:#9CA3AF;font-style:italic}
.td-val.bad{background:#FFF5F3;color:var(--zer)}
.row-ref{display:inline-block;margin-left:7px;vertical-align:middle}
.recon-hdr td{background:var(--zgl);padding:7px 14px;font-size:12px;font-weight:700;color:var(--zgd)}
.dtable tr.alert-row td{background:#FFF5F3;color:var(--zer);font-size:12px;font-weight:500}
/* citation badge */
.ref{display:inline-block;background:var(--zgl);color:var(--zgd);border:1px solid #8ADCAA;border-radius:5px;font-size:11px;font-weight:600;padding:2px 7px;white-space:nowrap}
/* fee table */
.fee-wrap{overflow-x:auto;padding:0}
.fee-tbl{width:100%;border-collapse:collapse;font-size:12.5px}
.fee-tbl th{padding:9px 12px;background:var(--zg);color:#fff;font-weight:600;text-align:left;white-space:nowrap}
.fee-tbl td{padding:9px 12px;border-bottom:1px solid var(--zbo);vertical-align:top;word-break:break-word}
.fee-tbl tr.alt td{background:#F9FAFB}
.fee-tbl .fee-bold{font-weight:700;color:var(--zgd)}
.fee-tbl tr:last-child td{border-bottom:none}
.no-fee{padding:14px 16px;color:var(--zer);font-size:13px;background:#FFF5F3;border-top:1px solid var(--zbo)}
/* comparison */
.cmp-summary{display:flex;gap:8px;flex-wrap:wrap;padding:12px 16px;border-bottom:1px solid var(--zbo);background:#FAFAFA}
.cmp-wrap{overflow-x:auto}
.cmp-tbl{width:100%;border-collapse:collapse;font-size:13px}
.cmp-tbl th{padding:9px 12px;background:var(--zn);color:#fff;font-weight:600;text-align:left;font-size:12px}
.cmp-tbl td{padding:10px 12px;border-bottom:1px solid var(--zbo);vertical-align:top;line-height:1.5}
.cmp-tbl tr:last-child td{border-bottom:none}
.cmp-row-pending td{background:#FFFBEB}
.cmp-row-pending td:first-child{border-left:3px solid #F59E0B}
.cmp-row-diff td{background:#FFF9F9}
.cmp-row-diff td:first-child{border-left:3px solid var(--zer)}
.badge-match{background:var(--zgl);color:var(--zgd);border-radius:5px;font-size:11px;font-weight:700;padding:3px 9px}
.badge-diff{background:#FFF5F3;color:var(--zer);border-radius:5px;font-size:11px;font-weight:700;padding:3px 9px}
.badge-only{background:var(--zgl);color:var(--zgd);border-radius:5px;font-size:11px;font-weight:700;padding:3px 9px}
.badge-gray{background:#F3F4F6;color:#374151;border-radius:5px;font-size:11px;font-weight:700;padding:3px 9px}
.badge-pending{background:#FEF3C7;color:#92400E;border-radius:5px;font-size:11px;font-weight:700;padding:3px 9px;border:1px solid #FCD34D}
/* risk table */
.risk-tbl{width:100%;border-collapse:collapse;font-size:13px}
.risk-tbl th{padding:9px 12px;background:var(--zn);color:rgba(255,255,255,.9);font-weight:600;text-align:left;font-size:12px}
.risk-tbl td{padding:10px 12px;border-bottom:1px solid var(--zbo);vertical-align:top;line-height:1.5}
.risk-tbl tr:last-child td{border-bottom:none}
.risk-tbl tr.risk-cao td{background:#FFF5F3}
.risk-tbl tr.risk-cao td:first-child{border-left:3px solid var(--zer)}
.risk-tbl tr.risk-tb td{background:#FFFBEB}
.risk-tbl tr.risk-tb td:first-child{border-left:3px solid #F59E0B}
.badge-risk-cao{background:#FEE2E2;color:#991B1B;border-radius:5px;font-size:11px;font-weight:700;padding:3px 9px}
.badge-risk-tb{background:#FEF3C7;color:#92400E;border-radius:5px;font-size:11px;font-weight:700;padding:3px 9px}
.badge-risk-thap{background:#F3F4F6;color:#374151;border-radius:5px;font-size:11px;font-weight:700;padding:3px 9px}
.badge-risk-type{background:var(--zbl);color:var(--zb);border-radius:4px;font-size:10px;font-weight:600;padding:2px 6px;white-space:nowrap}
.no-risk{padding:14px 16px;color:var(--zgd);font-size:13px;background:var(--zgl)}
.cmp-note{margin-top:12px;padding:12px 14px;background:#FFFBEB;border:1px solid #FCD34D;border-radius:8px;font-size:13px;color:#92400E}
/* custom answers */
.qa-item{padding:14px 16px;border-bottom:1px solid var(--zbo)}
.qa-item:last-child{border-bottom:none}
.qa-q{font-size:13px;font-weight:700;color:var(--zn);margin-bottom:6px;display:flex;align-items:flex-start;gap:8px}
.qa-q-num{background:var(--zg);color:#fff;border-radius:50%;width:20px;height:20px;display:inline-flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;flex-shrink:0;margin-top:1px}
.qa-a{font-size:13px;color:#374151;margin-left:28px;line-height:1.6}
.qa-ref{margin-left:28px;margin-top:5px}
/* history */
.hist-hdr{display:flex;align-items:center;justify-content:space-between;padding:0 4px 0 0}
.hist-clear{background:none;border:none;cursor:pointer;font-size:11px;color:#9CA3AF;padding:2px 7px;border-radius:4px;transition:color .15s,background .15s}
.hist-clear:hover{color:var(--zer);background:#FEE2E2}
.hist-item{display:flex;align-items:flex-start;gap:10px;padding:10px 14px;border-bottom:1px solid var(--zbo);cursor:pointer;transition:background .15s}
.hist-item:last-child{border-bottom:none}
.hist-item:hover{background:var(--zgl)}
.hist-ic{width:28px;height:28px;background:var(--zgl);border-radius:6px;display:flex;align-items:center;justify-content:center;font-size:14px;flex-shrink:0}
.hist-body{flex:1;min-width:0}
.hist-fnames{font-size:12px;font-weight:600;color:var(--zn);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.hist-date{font-size:11px;color:var(--zgr);margin-top:2px}
.hist-del{flex-shrink:0;background:none;border:none;cursor:pointer;font-size:15px;color:#D1D5DB;padding:0 2px;border-radius:4px;transition:color .15s,background .15s;align-self:center}
.hist-del:hover{color:var(--zer);background:#FEE2E2}
.hist-empty{padding:16px;text-align:center;color:var(--zgr);font-size:13px}
/* live chat */
.chat-fab{position:fixed;bottom:28px;right:28px;z-index:100;width:50px;height:50px;border-radius:50%;background:linear-gradient(135deg,var(--zg) 0%,#04A847 100%);color:#fff;border:none;cursor:pointer;font-size:22px;display:none;align-items:center;justify-content:center;box-shadow:0 4px 16px rgba(6,199,85,.45);transition:transform .2s,box-shadow .2s}
.chat-fab:hover{transform:scale(1.08);box-shadow:0 6px 24px rgba(6,199,85,.55)}
.chat-panel{position:fixed;bottom:90px;right:28px;z-index:100;width:360px;max-height:520px;background:#fff;border-radius:14px;box-shadow:0 8px 40px rgba(0,0,0,.18);display:none;flex-direction:column;overflow:hidden;border:1px solid var(--zbo)}
.chat-panel.open{display:flex}
.chat-ph{background:var(--zg);padding:12px 16px;display:flex;align-items:center;justify-content:space-between}
.chat-ph-t{color:#fff;font-size:13px;font-weight:700;display:flex;align-items:center;gap:7px}
.chat-ph-close{background:none;border:none;color:rgba(255,255,255,.6);font-size:18px;cursor:pointer;padding:0 4px;border-radius:4px;transition:color .15s}
.chat-ph-close:hover{color:#fff}
.chat-msgs{flex:1;overflow-y:auto;padding:12px 14px;display:flex;flex-direction:column;gap:9px;min-height:120px}
.chat-bubble{max-width:88%;padding:9px 13px;border-radius:12px;font-size:13px;line-height:1.55;word-break:break-word}
.chat-bubble.user{background:var(--zb);color:#fff;align-self:flex-end;border-bottom-right-radius:4px}
.chat-bubble.assistant{background:#F3F4F6;color:#111827;align-self:flex-start;border-bottom-left-radius:4px}
.chat-bubble.thinking{background:#F3F4F6;color:#9CA3AF;font-style:italic;align-self:flex-start}
.chat-input-row{display:flex;gap:8px;padding:10px 14px;border-top:1px solid var(--zbo);background:#FAFAFA}
.chat-inp{flex:1;border:1.5px solid var(--zbo);border-radius:8px;padding:8px 10px;font-size:13px;font-family:inherit;resize:none;min-height:38px;max-height:90px;transition:border-color .15s}
.chat-inp:focus{outline:none;border-color:var(--zg)}
.chat-send{flex-shrink:0;background:var(--zg);color:#fff;border:none;border-radius:8px;padding:0 14px;cursor:pointer;font-size:16px;transition:opacity .15s}
.chat-send:hover:not(:disabled){opacity:.85}
.chat-send:disabled{background:#9CA3AF;cursor:not-allowed}
/* sidebar info */
.info-hdr{background:var(--zn);color:rgba(255,255,255,.85);font-size:11px;font-weight:700;letter-spacing:.7px;text-transform:uppercase;padding:9px 14px;border-left:3px solid var(--zg)}
.tip{display:flex;align-items:flex-start;gap:8px;padding:9px 14px;border-bottom:1px solid var(--zbo);font-size:13px;color:#374151;line-height:1.5}
.tip:last-child{border-bottom:none}
.tip-ic{font-size:14px;flex-shrink:0;margin-top:1px}
.fnote{text-align:center;margin-top:12px;font-size:12px;color:#9CA3AF}
.fnote a{color:var(--zb);text-decoration:none}
/* folder upload toggle */
.upload-mode-tabs{display:flex;gap:0;margin-bottom:10px;border:1px solid #8ADCAA;border-radius:8px;overflow:hidden}
.upload-mode-tab{flex:1;padding:7px 10px;border:none;background:#F9FAFB;color:var(--zgr);font-size:12px;font-weight:600;cursor:pointer;transition:background .15s,color .15s}
.upload-mode-tab.active{background:var(--zg);color:#fff}
/* chat markdown */
.chat-md-tbl{width:100%;border-collapse:collapse;font-size:12px;margin:6px 0}
.chat-md-tbl th{padding:5px 8px;background:var(--zg);color:#fff;font-weight:600;text-align:left}
.chat-md-tbl td{padding:5px 8px;border-bottom:1px solid var(--zbo);vertical-align:top}
.chat-md-tbl tr:last-child td{border-bottom:none}
.chat-md-tbl tr:nth-child(even) td{background:#F9FAFB}
.chat-bubble strong{font-weight:700;color:inherit}
.chat-bubble ul{margin:4px 0 4px 16px;padding:0}
.chat-bubble li{margin:2px 0}
/* section highlight flash */
@keyframes sec-flash{0%{outline:2px solid var(--zb);outline-offset:2px}100%{outline:2px solid transparent}}
.sec-highlight{animation:sec-flash 1.5s ease-out forwards}
/* view-in-results btn */
.chat-view-btn{display:inline-flex;align-items:center;gap:5px;margin-top:7px;padding:4px 10px;background:var(--zgl);color:var(--zgd);border:1px solid #8ADCAA;border-radius:6px;font-size:11px;font-weight:600;cursor:pointer;transition:background .15s}
.chat-view-btn:hover{background:#D0F5E2}
.hero{margin-bottom:20px}
.hero h1{font-size:24px;font-weight:700;color:var(--zn);line-height:1.3;margin-bottom:7px}
.hero h1 span{color:var(--zg)}
.hero p{color:var(--zgr);font-size:14px;line-height:1.6}
</style>
</head>
<body>
<header class="hdr">
  <div class="hdr-l">
    <div>""" + _ZP_LOGO_SVG.replace("#0033C9","#ffffff").replace("#00CF6A","#ffffff") + """</div>
    <div class="hdr-div"></div>
    <span class="hdr-t">Contract Intelligence</span>
    <span class="hdr-badge">FP&amp;A Tool</span>
  </div>
  <div class="hdr-r">
    <a class="hdr-a" href="/docs" target="_blank">API Docs</a>
    <a class="hdr-a" href="/health">Status</a>
  </div>
</header>

<div class="wrap">
  <!-- LEFT -->
  <div>
    <!-- UPLOAD PHASE -->
    <div id="phaseUpload">
      <div class="hero">
        <h1>Phân tích hợp đồng <span>tự động</span></h1>
        <p>Upload hợp đồng Word / PDF — kết quả hiển thị trực tiếp trên trang với dẫn chiếu điều khoản cụ thể.</p>
      </div>
      <div class="card">
        <div class="card-t">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
          Upload tài liệu
        </div>
        <div class="flbl">
          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="var(--zb)" stroke-width="2.5"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
          Hợp đồng &amp; Phụ lục
        </div>
        <div class="upload-mode-tabs">
          <button class="upload-mode-tab active" id="tabFiles" onclick="switchUploadMode('files')">📄 Chọn file</button>
          <button class="upload-mode-tab" id="tabFolder" onclick="switchUploadMode('folder')">📁 Chọn thư mục</button>
        </div>
        <div id="dropFiles">
          <div class="drop" id="cDrop"
               onclick="document.getElementById('cInput').click()"
               ondragover="dnd(event,'over')" ondragleave="dnd(event,'leave')"
               ondrop="dndFiles(event,'cInput')">
            <input type="file" id="cInput" accept=".docx,.pdf" multiple onchange="onFC('cInput')">
            <div class="drop-ic">📂</div>
            <div class="drop-tx">Kéo thả hoặc <b>click để chọn file</b></div>
            <div class="drop-ex">.docx &nbsp;·&nbsp; .pdf &nbsp;·&nbsp; nhiều file</div>
          </div>
        </div>
        <div id="dropFolder" style="display:none">
          <div class="drop" id="folderDrop"
               onclick="document.getElementById('folderInput').click()"
               ondragover="dnd(event,'over')" ondragleave="dnd(event,'leave')"
               ondrop="dndFolder(event)">
            <input type="file" id="folderInput" webkitdirectory mozdirectory multiple onchange="onFolderChange()">
            <div id="folderIc" class="drop-ic">📁</div>
            <div id="folderTx" class="drop-tx">Kéo thả thư mục hoặc <b>click để chọn</b></div>
            <div id="folderEx" class="drop-ex">Tự động nhận diện .docx · .pdf (HĐ) và .eml · .msg (email)</div>
          </div>
        </div>
        <ul class="flist" id="cList"></ul>
        <div class="sep"></div>
        <div class="flbl">
          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#6B7280" stroke-width="2.5"><path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"/><polyline points="22,6 12,13 2,6"/></svg>
          Email alignment <span class="opt">tuỳ chọn</span>
        </div>
        <div class="drop" id="eDrop"
             onclick="document.getElementById('eInput').click()"
             ondragover="dnd(event,'over')" ondragleave="dnd(event,'leave')"
             ondrop="dndFiles(event,'eInput')">
          <input type="file" id="eInput" accept=".eml,.msg" multiple onchange="onFC('eInput')">
          <div class="drop-ic">📧</div>
          <div class="drop-tx">Kéo thả hoặc <b>click để chọn file email</b></div>
          <div class="drop-ex">.eml &nbsp;·&nbsp; .msg (Outlook)</div>
        </div>
        <ul class="flist" id="eList"></ul>
        <div class="sep"></div>
        <div class="flbl">
          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#6B7280" stroke-width="2.5"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
          Câu hỏi bổ sung <span class="opt">tuỳ chọn</span>
        </div>
        <textarea id="qInput" class="qbox" placeholder="Mỗi dòng là một câu hỏi. Ví dụ:&#10;Điều khoản chấm dứt hợp đồng là gì?&#10;Hợp đồng có quy định SLA không?"></textarea>
        <div class="sep"></div>
        <button class="btn-run" id="btnRun" onclick="analyze()">
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
          Chạy phân tích
        </button>
        <div class="sbox loading" id="sLoad">
          <div class="spin"></div>
          <div>
            <div style="font-weight:600;margin-bottom:2px">Đang phân tích tài liệu…</div>
            <div style="font-size:12px;opacity:.8">AI đang đọc và trích xuất thông tin, vui lòng chờ</div>
          </div>
        </div>
        <div class="sbox error" id="sErr"></div>
      </div>
    </div>

    <!-- RESULT PHASE (hidden by default) -->
    <div id="phaseResult" style="display:none">
      <div class="res-hdr">
        <div class="res-title" id="resTitle">Kết quả phân tích</div>
        <div class="res-actions">
          <button class="btn-back" onclick="backToUpload()">
            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="15 18 9 12 15 6"/></svg>
            Phân tích mới
          </button>
          <a class="btn-dl" id="dlBtn" href="#" download>
            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
            Tải Excel
          </a>
        </div>
      </div>
      <div class="card" id="resContent"></div>
    </div>
  </div>

  <!-- RIGHT SIDEBAR -->
  <div>
    <div class="card" style="padding:0;overflow:hidden">
      <div class="info-hdr hist-hdr">
        <span>Lịch sử review</span>
        <button class="hist-clear" onclick="clearAllHistory()" title="Xoá toàn bộ lịch sử">Xoá tất cả</button>
      </div>
      <div id="histList"><div class="hist-empty">Chưa có review nào</div></div>
    </div>
    <div class="card" style="padding:0;overflow:hidden;margin-top:14px">
      <div class="info-hdr">Lưu ý sử dụng</div>
      <div class="tip"><span class="tip-ic">📋</span>Kết quả hiển thị ngay trên trang, có dẫn chiếu điều khoản cụ thể</div>
      <div class="tip"><span class="tip-ic">📝</span>Thêm câu hỏi tùy chọn để lấy thêm thông tin theo nhu cầu</div>
      <div class="tip"><span class="tip-ic">⏱</span>Thời gian xử lý ~30–50 giây tuỳ độ dài HĐ</div>
      <div class="tip"><span class="tip-ic">🔒</span>File không được lưu trên server sau khi phân tích xong</div>
      <div class="tip"><span class="tip-ic">⚠️</span>Kiểm tra lại các trường đỏ — đây là thông tin AI chưa tìm thấy</div>
    </div>
    <p class="fnote">""" + _ZP_LOGO_SVG.replace('width="108"','width="68"').replace('height="24"','height="15"').replace("#0033C9","#9CA3AF").replace("#00CF6A","#9CA3AF") + """ &nbsp;FP&amp;A &nbsp;·&nbsp; <a href="/docs" target="_blank">API Docs</a></p>
  </div>
</div>

<!-- Live Chat -->
<button class="chat-fab" id="chatFab" onclick="toggleChat()" title="Hỏi đáp về hợp đồng">💬</button>
<div class="chat-panel" id="chatPanel">
  <div class="chat-ph">
    <span class="chat-ph-t">
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/></svg>
      Hỏi đáp hợp đồng
    </span>
    <button class="chat-ph-close" onclick="toggleChat()">×</button>
  </div>
  <div class="chat-msgs" id="chatMsgs">
    <div class="chat-bubble assistant">Xin chào! Tôi có thể giúp bạn tra cứu và giải đáp thắc mắc về hợp đồng vừa được phân tích. Hãy đặt câu hỏi.</div>
  </div>
  <div class="chat-input-row">
    <textarea class="chat-inp" id="chatInp" placeholder="Nhập câu hỏi…" rows="1"
      onkeydown="if(event.key==='Enter'&&!event.shiftKey){event.preventDefault();sendChat();}"></textarea>
    <button class="chat-send" id="chatSend" onclick="sendChat()">➤</button>
  </div>
</div>

<script>
// ── helpers ──────────────────────────────────────────────────────────
function esc(s){if(s==null)return'';return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}
function val(v){if(v==null||v==='')return null;if(Array.isArray(v))return v.join(', ')||null;return String(v);}
function badge(r){if(!r)return'';return '<span class="ref">'+esc(r)+'</span>';}
function rowHtml(label,v,ref,isBad){
  const d=val(v);
  const nullCls=d==null?' null':'';const badCls=isBad?' bad':'';
  const refHtml=ref?'<span class="row-ref">'+badge(ref)+'</span>':'';
  return '<tr><td class="td-lbl">'+esc(label)+'</td>'
    +'<td class="td-val'+nullCls+badCls+'">'+(d==null?'<em>—</em>':esc(d))+refHtml+'</td></tr>';
}
function sec(title,inner,open,sid){
  return '<div class="sec'+(open?' open':'')+(sid?' id="'+sid+'"':'')+'">'
    +'<div class="sec-hd" onclick="this.parentElement.classList.toggle(\\'open\\')"><span>'+esc(title)+'</span><span class="caret">&#9658;</span></div>'
    +'<div class="sec-bd">'+inner+'</div></div>';
}
function formatInline(s){
  return esc(s)
    .replace(/\*\*([^*\\n]+)\*\*/g,'<strong>$1</strong>')
    .replace(/\*([^*\\n]+)\*/g,'<em>$1</em>')
    .replace(/`([^`\\n]+)`/g,'<code style="background:#F3F4F6;padding:1px 4px;border-radius:3px;font-size:11px">$1</code>');
}
function renderMdTable(rows){
  if(!rows.length)return'';
  const hd=rows[0].map(c=>'<th>'+formatInline(c)+'</th>').join('');
  const bd=rows.slice(1).map((r,i)=>'<tr>'+ r.map(c=>'<td>'+formatInline(c)+'</td>').join('')+'</tr>').join('');
  return'<div style="overflow-x:auto"><table class="chat-md-tbl"><thead><tr>'+hd+'</tr></thead><tbody>'+bd+'</tbody></table></div>';
}
function renderMd(text){
  if(!text)return'';
  const lines=text.split('\\n');
  let html='',inList=false,tableRows=[],inTable=false;
  const flushTable=()=>{if(tableRows.length){html+=renderMdTable(tableRows);tableRows=[];inTable=false;}};
  const flushList=()=>{if(inList){html+='</ul>';inList=false;}};
  for(let i=0;i<lines.length;i++){
    const l=lines[i];
    if(l.trim().startsWith('|')&&l.trim().endsWith('|')){
      if(/^\|[-| :]+\|$/.test(l.trim()))continue;
      inTable=true;
      tableRows.push(l.split('|').slice(1,-1).map(c=>c.trim()));
      continue;
    }else{flushTable();}
    if(/^[-*] /.test(l)){
      flushList();if(!inList){html+='<ul>';inList=true;}
      html+='<li>'+formatInline(l.replace(/^[-*] /,''))+'</li>';continue;
    }else{flushList();}
    if(!l.trim()){html+='<br>';continue;}
    html+='<div>'+formatInline(l)+'</div>';
  }
  flushList();flushTable();
  return html;
}
function scrollToSection(sectionId){
  const el=document.getElementById('sec-'+sectionId);
  if(!el)return;
  document.getElementById('chatPanel').classList.remove('open');
  el.classList.add('open');
  setTimeout(()=>{
    el.scrollIntoView({behavior:'smooth',block:'start'});
    el.classList.add('sec-highlight');
    setTimeout(()=>el.classList.remove('sec-highlight'),1600);
  },120);
}
function switchUploadMode(mode){
  document.getElementById('dropFiles').style.display=mode==='files'?'':'none';
  document.getElementById('dropFolder').style.display=mode==='folder'?'':'none';
  document.getElementById('tabFiles').classList.toggle('active',mode==='files');
  document.getElementById('tabFolder').classList.toggle('active',mode==='folder');
  if(mode==='folder'){
    // reset folder zone display to default state
    document.getElementById('folderIc').textContent='📁';
    document.getElementById('folderTx').innerHTML='Kéo thả thư mục hoặc <b>click để chọn</b>';
    document.getElementById('folderEx').textContent='Tự động nhận diện .docx · .pdf (HĐ) và .eml · .msg (email)';
  }
}
function _showFolderResult(folderName,cCount,eCount){
  document.getElementById('folderIc').textContent='✅';
  document.getElementById('folderTx').innerHTML='<b>'+esc(folderName)+'</b>';
  const parts=[];
  if(cCount)parts.push(cCount+' hợp đồng/phụ lục');
  if(eCount)parts.push(eCount+' email');
  document.getElementById('folderEx').textContent=parts.length?'Đã nhận: '+parts.join(', '):'Không tìm thấy file phù hợp trong thư mục';
}
function _sortFolderFiles(files,folderName){
  const CONTRACT_EXT=['.docx','.pdf'];
  const EMAIL_EXT=['.eml','.msg'];
  const cFiles=[],eFiles=[];
  Array.from(files).forEach(f=>{
    const ext='.'+f.name.split('.').pop().toLowerCase();
    if(CONTRACT_EXT.includes(ext))cFiles.push(f);
    else if(EMAIL_EXT.includes(ext))eFiles.push(f);
  });
  if(cFiles.length)addFiles('cInput',cFiles);
  if(eFiles.length)addFiles('eInput',eFiles);
  if(folderName)_showFolderResult(folderName,cFiles.length,eFiles.length);
  return{c:cFiles.length,e:eFiles.length};
}
function onFolderChange(){
  const inp=document.getElementById('folderInput');
  const files=inp.files;
  if(!files||!files.length)return;
  // derive folder name from first file's webkitRelativePath
  let folderName='Thư mục đã chọn';
  if(files[0].webkitRelativePath){
    folderName=files[0].webkitRelativePath.split('/')[0];
  }
  _sortFolderFiles(files,folderName);
}
// ── file store ────────────────────────────────────────────────────────
const ICONS={pdf:'📄',docx:'📝',eml:'📩',msg:'📨'};
const PAIR={cInput:'cList',eInput:'eList'};
const _store={cInput:[],eInput:[]};
function syncInput(id){const dt=new DataTransfer();_store[id].forEach(f=>dt.items.add(f));document.getElementById(id).files=dt.files;}
function renderList(id){
  const ul=document.getElementById(PAIR[id]);
  if(!_store[id].length){ul.innerHTML='';return;}
  ul.innerHTML=_store[id].map((f,i)=>{
    const ext=f.name.split('.').pop().toLowerCase();
    const sz=f.size<1024?f.size+' B':f.size<1048576?(f.size/1024).toFixed(1)+' KB':(f.size/1048576).toFixed(2)+' MB';
    return '<li><span style="font-size:14px">'+( ICONS[ext]||'📄')+'</span>'
      +'<span class="fn" title="'+esc(f.name)+'">'+esc(f.name)+'</span>'
      +'<span class="fsz">'+sz+'</span>'
      +'<button class="del" onclick="rmFile(\\''+id+'\\','+i+')" title="Xoá">×</button></li>';
  }).join('');
}
function addFiles(id,files){const ns=new Set(_store[id].map(f=>f.name));Array.from(files).forEach(f=>{if(!ns.has(f.name)){_store[id].push(f);ns.add(f.name);}});syncInput(id);renderList(id);}
function rmFile(id,i){_store[id].splice(i,1);syncInput(id);renderList(id);}
function onFC(id){addFiles(id,document.getElementById(id).files);}
// drag helper — hover state only
function dnd(e,action){e.preventDefault();e.stopPropagation();const d=e.currentTarget;if(action==='over')d.classList.add('over');else d.classList.remove('over');}
// drop handler for regular files
function dndFiles(e,id){e.preventDefault();e.stopPropagation();e.currentTarget.classList.remove('over');addFiles(id,e.dataTransfer.files);}
// drop handler for folder — uses FileSystem API when available, else falls back to files list
function dndFolder(e){
  e.preventDefault();e.stopPropagation();e.currentTarget.classList.remove('over');
  const items=e.dataTransfer.items;
  if(items&&items.length){
    const entry=items[0].webkitGetAsEntry?items[0].webkitGetAsEntry():null;
    if(entry&&entry.isDirectory){
      document.getElementById('folderIc').textContent='⏳';
      document.getElementById('folderTx').innerHTML='Đang đọc thư mục…';
      document.getElementById('folderEx').textContent='';
      _readDirEntry(entry,entry.name);return;
    }
  }
  // dropped individual files
  _sortFolderFiles(e.dataTransfer.files,'Files đã kéo thả');
}
function _readDirEntry(dirEntry,folderName){
  const reader=dirEntry.createReader();
  const allFiles=[];
  function readBatch(){
    reader.readEntries(entries=>{
      if(!entries.length){
        Promise.all(allFiles).then(files=>_sortFolderFiles(files,folderName));
        return;
      }
      entries.forEach(en=>{
        if(en.isFile){
          allFiles.push(new Promise(res=>en.file(res)));
        }
        // recurse one level into sub-folders
        else if(en.isDirectory){
          const sub=en.createReader();
          allFiles.push(new Promise(res=>{
            sub.readEntries(subEntries=>{
              Promise.all(subEntries.filter(s=>s.isFile).map(s=>new Promise(r=>s.file(r)))).then(res);
            });
          }));
        }
      });
      readBatch();
    });
  }
  readBatch();
}
// ── status helpers ────────────────────────────────────────────────────
function showLoad(){document.getElementById('sLoad').className='sbox loading';document.getElementById('sErr').className='sbox';}
function showErr(m){document.getElementById('sLoad').className='sbox';const e=document.getElementById('sErr');e.className='sbox error';e.textContent=m;}
function hideStatus(){document.getElementById('sLoad').className='sbox';document.getElementById('sErr').className='sbox';}
// ── analyze ───────────────────────────────────────────────────────────
async function analyze(){
  if(!_store.cInput.length){showErr('⚠️  Vui lòng chọn ít nhất 1 file hợp đồng (.docx hoặc .pdf)');return;}
  document.getElementById('btnRun').disabled=true;
  hideStatus();showLoad();
  const fd=new FormData();
  _store.cInput.forEach(f=>fd.append('contracts',f));
  _store.eInput.forEach(f=>fd.append('emails',f));
  const q=document.getElementById('qInput').value.trim();
  if(q)fd.append('custom_query',q);
  try{
    const r1=await fetch('/analyze',{method:'POST',body:fd});
    if(!r1.ok){let m=await r1.text();try{m=JSON.parse(m).detail||m;}catch{}showErr('❌ '+m);document.getElementById('btnRun').disabled=false;return;}
    const {job_id}=await r1.json();
    for(let i=0;i<120;i++){
      await new Promise(r=>setTimeout(r,3000));
      const r2=await fetch('/result/'+job_id);
      if(r2.status===202)continue;
      if(!r2.ok){let m=await r2.text();try{m=JSON.parse(m).detail||m;}catch{}showErr('❌ '+m);document.getElementById('btnRun').disabled=false;return;}
      const data=await r2.json();
      hideStatus();
      document.getElementById('btnRun').disabled=false;
      renderResults(data);
      loadHistory();
      return;
    }
    showErr('❌ Timeout — phân tích mất quá lâu, vui lòng thử lại.');
    document.getElementById('btnRun').disabled=false;
  }catch(err){showErr('❌ Lỗi kết nối: '+err.message);document.getElementById('btnRun').disabled=false;}
}
function backToUpload(){
  document.getElementById('phaseResult').style.display='none';
  document.getElementById('phaseUpload').style.display='';
  hideChatFab();
}
// ── render results ────────────────────────────────────────────────────
function renderResults(data){
  const rv=data.review_id||'';
  document.getElementById('dlBtn').href='/download/'+rv;
  document.getElementById('resTitle').textContent='Kết quả phân tích'+(data.contracts&&data.contracts.length>1?' ('+data.contracts.length+' hợp đồng)':'');
  let html='';
  const multi=data.contracts&&data.contracts.length>1;
  if(multi){
    html+='<div class="contract-tabs">';
    data.contracts.forEach((c,i)=>html+='<button class="ctab'+(i===0?' active':'')+'" onclick="switchTab('+i+')">'+esc(c.filename||('HĐ '+(i+1)))+'</button>');
    html+='</div>';
  }
  (data.contracts||[]).forEach((c,i)=>{
    html+='<div class="cpanel'+((!multi||i===0)?' active':'')+'" id="cp'+i+'">';
    html+=buildContract(c);
    html+='</div>';
  });
  if(data.comparison&&data.comparison.ket_qua_so_sanh&&data.comparison.ket_qua_so_sanh.length){
    html+=sec('📧 So sánh Email vs Hợp đồng',buildComparison(data.comparison),true);
  }
  if(data.custom_answers&&data.custom_answers.length){
    html+=sec('💬 Câu hỏi bổ sung',buildCustom(data.custom_answers),true);
  }
  document.getElementById('resContent').innerHTML=html;
  document.getElementById('phaseUpload').style.display='none';
  document.getElementById('phaseResult').style.display='';
  showChatFab(rv);
}
function switchTab(idx){
  document.querySelectorAll('.ctab').forEach((b,i)=>b.classList.toggle('active',i===idx));
  document.querySelectorAll('.cpanel').forEach((p,i)=>p.classList.toggle('active',i===idx));
}
function buildContract(c){
  const d=c.data||{};const refs=d._refs||{};
  const miss=new Set(d.truong_con_thieu||[]);
  let html='';
  if(miss.size)html+='<div class="missing-banner">⚠️ Thông tin còn thiếu: '+[...miss].map(esc).join(' · ')+'</div>';
  // S1 — Thông tin chung
  let s1='<table class="dtable">';
  s1+=rowHtml('Loại hợp đồng',d.loai_hop_dong,refs.loai_hop_dong,!d.loai_hop_dong);
  (d.cac_ben||[]).forEach(p=>{
    s1+=rowHtml('Bên: '+esc(p.ten_ben||''),p.ten_cong_ty,refs.cac_ben,!p.ten_cong_ty);
    if(p.nguoi_dai_dien)s1+=rowHtml('  Người đại diện',(p.nguoi_dai_dien||'')+(p.chuc_vu?' — '+p.chuc_vu:''),null,false);
    if(p.ma_so_thue)s1+=rowHtml('  MST',p.ma_so_thue,null,false);
  });
  const t=d.thoi_han_hop_dong||{};
  s1+=rowHtml('Ngày ký',t.ngay_ky,refs.thoi_han_hop_dong,!t.ngay_ky);
  s1+=rowHtml('Ngày hiệu lực',t.ngay_hieu_luc,null,!t.ngay_hieu_luc);
  s1+=rowHtml('Thời gian hiệu lực',t.thoi_gian_hieu_luc,null,false);
  s1+=rowHtml('Ngày hết hạn',t.ngay_het_han,null,false);
  s1+=rowHtml('Điều kiện gia hạn',t.dieu_kien_gia_han,null,false);
  const dv=d.dich_vu_hop_tac||{};
  s1+=rowHtml('Mô tả dịch vụ',dv.mo_ta_chung,refs.dich_vu_hop_tac,false);
  s1+=rowHtml('Kênh thanh toán',(dv.kenh_thanh_toan||[]).join(', ')||null,null,!(dv.kenh_thanh_toan||[]).length);
  s1+=rowHtml('Nguồn tiền',(dv.nguon_tien||[]).join(', ')||null,null,!(dv.nguon_tien||[]).length);
  s1+='</table>';
  html+=sec('1. Thông tin chung & Các bên',s1,true,'sec-summary');
  // S2 — Commercial
  const ct=d.commercial_terms||{};
  let s2='';
  s2+='<table class="dtable">'+rowHtml('Tổng giá trị HĐ',ct.tong_gia_tri_hop_dong,refs['commercial_terms.tong_gia_tri'],false)+'</table>';
  const fees=ct.phi_dich_vu||[];
  if(fees.length){
    s2+='<div class="fee-wrap"><table class="fee-tbl"><thead><tr><th>Loại phí</th><th>Kênh TT</th><th>Nguồn tiền</th><th>Mức phí</th><th>Điều kiện</th><th>Ghi chú</th></tr></thead><tbody>';
    fees.forEach((f,i)=>{s2+='<tr class="'+(i%2?'alt':'')+'"><td>'+esc(f.loai_phi||'')+'</td><td>'+esc(f.kenh_thanh_toan||'')+'</td><td>'+esc(f.nguon_tien||'')+'</td><td class="fee-bold">'+esc(f.muc_phi||'')+'</td><td>'+esc(f.dieu_kien||'')+'</td><td>'+esc(f.ghi_chu||'')+'</td></tr>';});
    s2+='</tbody></table>';
    if(refs['commercial_terms.phi_dich_vu'])s2+='<div style="padding:6px 10px;font-size:12px">Dẫn chiếu: '+badge(refs['commercial_terms.phi_dich_vu'])+'</div>';
    s2+='</div>';
  }else{s2+='<div class="no-fee">⚠️ Không tìm thấy thông tin phí dịch vụ trong hợp đồng</div>';}
  const km=ct.ngan_sach_khuyen_mai||{};
  if(km.tong_ngan_sach||km.the_le_dieu_kien){
    s2+='<table class="dtable">'+rowHtml('Ngân sách KM',km.tong_ngan_sach,refs['commercial_terms.ngan_sach_khuyen_mai'],false)+rowHtml('Thể lệ / Điều kiện KM',km.the_le_dieu_kien,null,false)+'</table>';
  }
  const lp=ct.lai_va_phat||{};
  s2+='<table class="dtable">'+rowHtml('Lãi trả chậm',lp.lai_tra_cham,refs['commercial_terms.lai_va_phat'],false)+rowHtml('Phạt vi phạm',lp.phat_vi_pham,null,false)+rowHtml('Bồi thường thiệt hại',lp.boi_thuong_thiet_hai,null,false)+'</table>';
  html+=sec('2. Commercial Terms & Phí dịch vụ',s2,true,'sec-commercial');
  // S3 — Payment & Recon
  const pt=ct.payment_term||{};const rt=ct.reconciliation_term||{};
  let s3='<table class="dtable">';
  s3+=rowHtml('Cơ chế thanh toán',pt.co_che_thanh_toan,refs['commercial_terms.payment_term'],!pt.co_che_thanh_toan);
  s3+=rowHtml('Tạm ứng / TT trước',pt.tam_ung_thanh_toan_truoc,null,false);
  s3+=rowHtml('Công nợ thanh toán',pt.cong_no_thanh_toan,null,!pt.cong_no_thanh_toan);
  s3+=rowHtml('Hồ sơ thanh toán',(pt.ho_so_thanh_toan||[]).join(', ')||null,null,!(pt.ho_so_thanh_toan||[]).length);
  if(pt.alert_ho_so)s3+='<tr class="alert-row"><td colspan="2">⚠️ '+esc(pt.alert_ho_so)+'</td></tr>';
  s3+='<tr class="recon-hdr"><td colspan="2">Reconciliation Term</td></tr>';
  s3+=rowHtml('Bắt đầu đối soát',rt.thoi_gian_bat_dau_doi_soat,refs['commercial_terms.reconciliation_term'],!rt.thoi_gian_bat_dau_doi_soat);
  s3+=rowHtml('Gửi đối soát',rt.thoi_gian_gui_doi_soat,null,false);
  s3+=rowHtml('Thời hạn phản hồi',rt.thoi_gian_phan_hoi,null,false);
  s3+=rowHtml('Xử lý chênh lệch',rt.xu_ly_chenh_lech,null,false);
  s3+=rowHtml('Zalopay xuất hóa đơn',rt.zalopay_xuat_hoa_don,null,false);
  s3+='</table>';
  html+=sec('3. Payment & Reconciliation Term',s3,true,'sec-payment');
  // S4 — Risk Assessment
  html+=sec('4. ⚠️ Đánh giá rủi ro cho Zalopay',buildRisks(d.danh_gia_rui_ro),true,'sec-risk');
  return html;
}
function buildRisks(risks){
  if(!risks||!risks.length)return '<div class="no-risk">✅ Không phát hiện điều khoản rủi ro đáng kể cho Zalopay.</div>';
  const MUC_CLS={CAO:'badge-risk-cao',TRUNG_BINH:'badge-risk-tb',THAP:'badge-risk-thap'};
  const MUC_LBL={CAO:'CAO',TRUNG_BINH:'TRUNG BÌNH',THAP:'THẤP'};
  const ORDER={CAO:0,TRUNG_BINH:1,THAP:2};
  const sorted=[...risks].sort((a,b)=>((ORDER[a.muc_do]??3)-(ORDER[b.muc_do]??3)));
  let html=\'<div style="overflow-x:auto"><table class="risk-tbl" style="table-layout:fixed;width:100%"><thead><tr>\'
    +\'<th style="width:140px">Mức độ / Loại</th>\'
    +\'<th>Mô tả rủi ro</th>\'
    +\'<th style="width:110px;text-align:center">Điều khoản</th>\'
    +\'</tr></thead><tbody>\';
  sorted.forEach(r=>{
    const muc=r.muc_do||\'\';
    const rowCls=muc===\'CAO\'?\' class="risk-cao"\':(muc===\'TRUNG_BINH\'?\' class="risk-tb"\':\'\');
    html+=\'<tr\'+rowCls+\'>\'
      +\'<td style="word-break:keep-all"><span class="\'+(MUC_CLS[muc]||\'badge-risk-thap\')+\'">\'+( MUC_LBL[muc]||esc(muc))+\'</span><br><span class="badge-risk-type" style="display:inline-block;margin-top:5px">\'+esc(r.loai_rui_ro||\'\')+\'</span></td>\'
      +\'<td style="word-break:break-word">\'+esc(r.tom_tat||\'\')+\'</td>\'
      +\'<td style="text-align:center;word-break:break-word">\'+badge(r.dieu_khoan||null)+\'</td>\'
      +\'</tr>\';
  });
  html+=\'</tbody></table></div>\';
  return html;
}
function buildComparison(cmp){
  const rows=cmp.ket_qua_so_sanh||[];
  const ST_LABEL={KHOP:'KHỚP',KHAC_BIET:'KHÁC BIỆT',CHI_EMAIL:'CHỈ TRONG EMAIL',CHI_HD:'CHỈ TRONG HĐ',CHUA_PHAN_HOI:'CHƯA PHẢN HỒI'};
  const ST_CLS={KHOP:'badge-match',KHAC_BIET:'badge-diff',CHI_EMAIL:'badge-only',CHI_HD:'badge-gray',CHUA_PHAN_HOI:'badge-pending'};
  // normalise legacy Vietnamese status values
  function norm(s){const m={'KHỚP':'KHOP','KHÁC BIỆT':'KHAC_BIET','CHỈ TRONG EMAIL':'CHI_EMAIL','CHỈ TRONG HĐ':'CHI_HD','CHƯA PHẢN HỒI':'CHUA_PHAN_HOI'};return m[s]||s;}
  const cnt={};rows.forEach(r=>{const k=norm(r.trang_thai||'');cnt[k]=(cnt[k]||0)+1;});
  let sumHtml='<div class="cmp-summary">';
  ['CHUA_PHAN_HOI','KHAC_BIET','CHI_EMAIL','CHI_HD','KHOP'].forEach(k=>{
    if(cnt[k])sumHtml+='<span class="'+(ST_CLS[k]||'badge-only')+'">'+cnt[k]+'&nbsp;'+(ST_LABEL[k]||k)+'</span>';
  });
  sumHtml+='</div>';
  const ORDER={CHUA_PHAN_HOI:0,KHAC_BIET:1,CHI_EMAIL:2,CHI_HD:3,KHOP:4};
  const sorted=[...rows].sort((a,b)=>((ORDER[norm(a.trang_thai)]??5)-(ORDER[norm(b.trang_thai)]??5)));
  let html=sumHtml+'<div class="cmp-wrap"><table class="cmp-tbl"><thead><tr>'
    +'<th style="width:148px">Trạng thái</th><th>Điểm so sánh</th>'
    +'<th>Trong Email</th><th>Trong HĐ</th><th>Ghi chú</th>'
    +'</tr></thead><tbody>';
  sorted.forEach(r=>{
    const st=norm(r.trang_thai||'');
    const rowCls=st==='CHUA_PHAN_HOI'?' class="cmp-row-pending"':st==='KHAC_BIET'?' class="cmp-row-diff"':'';
    html+='<tr'+rowCls+'>'
      +'<td><span class="'+(ST_CLS[st]||'badge-only')+'">'+(ST_LABEL[st]||esc(r.trang_thai||''))+'</span></td>'
      +'<td><strong>'+esc(r.diem_so_sanh||'')+'</strong></td>'
      +'<td>'+esc(r.trong_email||'—')+'</td>'
      +'<td>'+esc(r.trong_hop_dong||'—')+'</td>'
      +'<td style="color:#6B7280;font-style:italic">'+esc(r.ghi_chu||'')+'</td>'
      +'</tr>';
  });
  html+='</tbody></table></div>';
  if(cmp.tong_ket)html+='<div class="cmp-note">📋 '+esc(cmp.tong_ket)+'</div>';
  return html;
}
function buildCustom(answers){
  return '<div>'+(answers||[]).map((a,i)=>'<div class="qa-item"><div class="qa-q"><span class="qa-q-num">'+(i+1)+'</span>'+esc(a.question||'')+'</div><div class="qa-a">'+esc(a.answer||'')+'</div>'+(a.ref?'<div class="qa-ref">'+badge(a.ref)+'</div>':'')+'</div>').join('')+'</div>';
}
// ── history ───────────────────────────────────────────────────────────
async function loadHistory(){
  try{
    const r=await fetch('/history');
    const items=await r.json();
    const el=document.getElementById('histList');
    if(!items||!items.length){el.innerHTML='<div class="hist-empty">Chưa có review nào</div>';return;}
    el.innerHTML=items.map(it=>{
      const fnames=(it.filenames||[]).join(', ')||'—';
      const dt=it.created_at?new Date(it.created_at).toLocaleString('vi-VN',{dateStyle:'short',timeStyle:'short'}):'';
      return '<div class="hist-item"><div class="hist-ic" onclick="viewReview(\\''+esc(it.id)+'\\')">📋</div>'
        +'<div class="hist-body" onclick="viewReview(\\''+esc(it.id)+'\\')" style="cursor:pointer"><div class="hist-fnames" title="'+esc(fnames)+'">'+esc(fnames)+'</div><div class="hist-date">'+esc(dt)+'</div></div>'
        +'<button class="hist-del" onclick="deleteReview(\\''+esc(it.id)+'\\',event)" title="Xoá review này">×</button></div>';
    }).join('');
  }catch(e){document.getElementById('histList').innerHTML='<div class="hist-empty">Không tải được lịch sử</div>';}
}
async function deleteReview(id,e){
  e.stopPropagation();
  if(!confirm('Xoá review này khỏi lịch sử?'))return;
  try{
    await fetch('/history/'+id,{method:'DELETE'});
    loadHistory();
  }catch(e){}
}
async function clearAllHistory(){
  if(!confirm('Xoá toàn bộ lịch sử review?'))return;
  try{
    await fetch('/history',{method:'DELETE'});
    loadHistory();
  }catch(e){}
}
async function viewReview(id){
  try{
    const r=await fetch('/history/'+id);
    if(!r.ok)return;
    const item=await r.json();
    const data={
      review_id:item.id,
      contracts:(item.results||[]).map((d,i)=>({filename:(item.filenames||[])[i]||('HĐ '+(i+1)),data:d})),
      comparison:item.comparison,
      custom_answers:item.custom_answers||[],
    };
    renderResults(data);
  }catch(e){}
}
// ── live chat ─────────────────────────────────────────────────────────
let _chatReviewId='';let _chatHistory=[];
function toggleChat(){
  const p=document.getElementById('chatPanel');
  p.classList.toggle('open');
  if(p.classList.contains('open'))document.getElementById('chatInp').focus();
}
function showChatFab(reviewId){
  _chatReviewId=reviewId;_chatHistory=[];
  document.getElementById('chatFab').style.display='flex';
  // Reset chat messages
  document.getElementById('chatMsgs').innerHTML='<div class="chat-bubble assistant">Xin chào! Tôi có thể giúp bạn tra cứu và giải đáp thắc mắc về hợp đồng vừa được phân tích. Hãy đặt câu hỏi.</div>';
}
function hideChatFab(){
  document.getElementById('chatFab').style.display='none';
  document.getElementById('chatPanel').classList.remove('open');
}
async function sendChat(){
  const inp=document.getElementById('chatInp');
  const msg=inp.value.trim();
  if(!msg||!_chatReviewId)return;
  inp.value='';inp.style.height='';
  const msgs=document.getElementById('chatMsgs');
  msgs.innerHTML+='<div class="chat-bubble user">'+esc(msg)+'</div>';
  const thinking=document.createElement('div');
  thinking.className='chat-bubble thinking';thinking.textContent='Đang trả lời…';
  msgs.appendChild(thinking);msgs.scrollTop=msgs.scrollHeight;
  document.getElementById('chatSend').disabled=true;
  try{
    const fd=new FormData();
    fd.append('review_id',_chatReviewId);
    fd.append('message',msg);
    fd.append('history',JSON.stringify(_chatHistory));
    const r=await fetch('/chat',{method:'POST',body:fd});
    const {reply,highlight}=await r.json();
    _chatHistory.push({role:'user',content:msg},{role:'assistant',content:reply});
    thinking.className='chat-bubble assistant';
    thinking.innerHTML=renderMd(reply);
    if(highlight){
      const btn=document.createElement('button');
      btn.className='chat-view-btn';
      btn.innerHTML='📍 Xem trong kết quả';
      btn.onclick=()=>scrollToSection(highlight);
      thinking.appendChild(btn);
    }
  }catch(err){
    thinking.className='chat-bubble assistant';thinking.textContent='❌ Lỗi kết nối, vui lòng thử lại.';
  }
  msgs.scrollTop=msgs.scrollHeight;
  document.getElementById('chatSend').disabled=false;
  inp.focus();
}
// ── init ─────────────────────────────────────────────────────────────
loadHistory();
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


@app.post("/debug-filename", include_in_schema=False)
async def debug_filename(contracts: list[UploadFile] = File(...)):
    """Debug endpoint: echo back raw and decoded filenames."""
    results = []
    for f in contracts:
        raw = f.filename or ""
        decoded = _decode_filename(raw)
        results.append({
            "raw": raw,
            "raw_repr": repr(raw),
            "decoded": decoded,
            "decoded_repr": repr(decoded),
        })
    return JSONResponse({"files": results})


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
    custom_query: str = Form(
        default="",
        description="(Tuỳ chọn) Câu hỏi bổ sung, mỗi dòng là một câu hỏi.",
    ),
    credentials: HTTPAuthorizationCredentials | None = Security(security),
    x_user_id: str | None = Header(default=None, alias="X-GreenNode-AgentBase-User-Id"),
    x_session_id: str | None = Header(default=None, alias="X-GreenNode-AgentBase-Session-Id"),
):
    """
    Bắt đầu phân tích hợp đồng trong nền. Trả về `job_id` ngay lập tức (202).
    Poll kết quả tại **GET /result/{job_id}**.
    """
    verify_auth(credentials)

    for f in contracts:
        fname = _decode_filename(f.filename)
        if not fname.lower().endswith((".docx", ".pdf")):
            raise HTTPException(400, f"File '{fname}' phải là .docx hoặc .pdf")
    for f in emails:
        fname = _decode_filename(f.filename)
        if not fname.lower().endswith((".eml", ".msg")):
            raise HTTPException(400, f"File '{fname}' phải là .eml hoặc .msg")

    api_key = os.environ.get("LLM_API_KEY", "")
    if not api_key:
        raise HTTPException(500, "LLM_API_KEY chưa được cấu hình trên server.")

    contract_data = [(_decode_filename(f.filename) or f"contract_{uuid.uuid4().hex}.docx", await f.read()) for f in contracts]
    email_data    = [(_decode_filename(f.filename) or f"email_{uuid.uuid4().hex}.eml",     await f.read()) for f in emails]
    questions     = [q.strip() for q in custom_query.splitlines() if q.strip()]

    job_id = str(uuid.uuid4())
    _jobs[job_id] = {"status": "processing"}

    asyncio.create_task(
        _run_analysis(job_id, contract_data, email_data, api_key, x_user_id, x_session_id, questions)
    )

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

    # Done — trả JSON và log memory
    results       = job.get("results", [])
    comparison    = job.get("comparison")
    custom_answers= job.get("custom_answers", [])
    x_user_id     = job.get("x_user_id")
    x_session_id  = job.get("x_session_id")
    fnames        = job.get("contract_fnames", [])
    review_id     = job.get("review_id", job_id)
    del _jobs[job_id]

    if x_user_id and results:
        first = results[0]
        parties = "; ".join(
            b.get("ten_cong_ty", "") for b in (first.get("cac_ben") or []) if b.get("ten_cong_ty")
        )
        missing = ", ".join(first.get("truong_con_thieu") or []) or "Không có"
        summary = (
            f"Loại HĐ: {first.get('loai_hop_dong') or 'N/A'}\n"
            f"Các bên: {parties or 'N/A'}\n"
            f"Ngày ký: {(first.get('thoi_han_hop_dong') or {}).get('ngay_ky') or 'N/A'}\n"
            f"Trường thiếu: {missing}"
        )
        await _log_to_memory(
            user_id=x_user_id,
            session_id=x_session_id or str(uuid.uuid4()),
            user_msg=f"Phân tích hợp đồng: {', '.join(fnames)}",
            assistant_msg=summary,
        )

    contracts_out = [
        {"filename": fnames[i] if i < len(fnames) else f"hop_dong_{i+1}", "data": d}
        for i, d in enumerate(results)
    ]
    return JSONResponse({
        "status":         "done",
        "review_id":      review_id,
        "contracts":      contracts_out,
        "comparison":     comparison,
        "custom_answers": custom_answers,
    })


@app.get("/download/{review_id}", tags=["Contract Review"], summary="Tải file Excel")
async def download_excel(review_id: str):
    """Tải file Excel đã được tạo cho một lần phân tích."""
    data = _db_get_excel(review_id)
    if not data:
        raise HTTPException(404, "Không tìm thấy file Excel cho review này.")
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    tmp.write(data); tmp.close()
    return FileResponse(
        path=tmp.name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="ContractReview.xlsx",
    )


@app.get("/history", tags=["History"], summary="Danh sách hợp đồng đã review")
async def list_history():
    """Trả về danh sách 30 review gần nhất."""
    return JSONResponse(_db_list(30))


@app.get("/history/{review_id}", tags=["History"], summary="Chi tiết một review")
async def get_history_item(review_id: str):
    """Trả về đầy đủ kết quả phân tích của một review trong lịch sử."""
    item = _db_get(review_id)
    if not item:
        raise HTTPException(404, "Không tìm thấy review.")
    return JSONResponse(item)


@app.delete("/history/{review_id}", tags=["History"], summary="Xoá một review")
async def delete_history_item(review_id: str):
    """Xoá một review khỏi lịch sử."""
    ok = _db_delete(review_id)
    if not ok:
        raise HTTPException(404, "Không tìm thấy review.")
    return JSONResponse({"status": "deleted", "id": review_id})


@app.delete("/history", tags=["History"], summary="Xoá toàn bộ lịch sử")
async def delete_all_history():
    """Xoá toàn bộ lịch sử review."""
    count = _db_delete_all()
    return JSONResponse({"status": "deleted", "count": count})


@app.post("/chat", tags=["Chat"], summary="Hỏi đáp về hợp đồng đã phân tích")
async def chat(
    review_id: str = Form(..., description="ID của review cần hỏi đáp"),
    message: str = Form(..., description="Câu hỏi của người dùng"),
    history: str = Form(default="[]", description="Lịch sử chat dạng JSON array [{role, content}]"),
):
    """Chat trực tiếp với kết quả phân tích hợp đồng."""
    item = _db_get(review_id)
    if not item:
        raise HTTPException(404, "Không tìm thấy review.")

    api_key = os.environ.get("LLM_API_KEY", "")
    if not api_key:
        raise HTTPException(500, "LLM_API_KEY chưa được cấu hình.")

    results = item.get("results") or []
    if not results:
        raise HTTPException(400, "Review này không có dữ liệu hợp đồng.")

    try:
        chat_history = json.loads(history)
    except Exception:
        chat_history = []

    contract_data = results[0]
    loop = asyncio.get_running_loop()
    result = await loop.run_in_executor(
        None, chat_with_contract, message, contract_data, chat_history, api_key
    )
    return JSONResponse({"reply": result.get("reply", ""), "highlight": result.get("highlight")})
