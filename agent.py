"""
Contract Review Agent — Zalopay FP&A
=====================================
Trích xuất thông tin hợp đồng Word + so sánh email alignment → Excel.

Sử dụng:
    python agent.py --contracts hop_dong.docx phu_luc.docx
    python agent.py --contracts hop_dong.docx --emails aligned.eml
    python agent.py --folder ./ho_so/
    python agent.py --contracts hop_dong.docx --output review.xlsx

Biến môi trường:
    ANTHROPIC_API_KEY   (bắt buộc)
"""

from __future__ import annotations

import argparse
import email as email_lib
import json
import os
import re
import sys
from datetime import datetime
from email import policy as email_policy
from pathlib import Path

# ── Third-party ──────────────────────────────────────────────────────────────
try:
    from docx import Document
    from docx.oxml.ns import qn
    from docx.table import Table as DocxTable
except ImportError:
    sys.exit("ERROR: pip install python-docx")

try:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("ERROR: pip install openpyxl")

try:
    from openai import OpenAI
except ImportError:
    sys.exit("ERROR: pip install openai")

try:
    import pdfplumber
    _PDF_AVAILABLE = True
except ImportError:
    _PDF_AVAILABLE = False

try:
    import extract_msg as _extract_msg
    _MSG_AVAILABLE = True
except ImportError:
    _MSG_AVAILABLE = False


# ═════════════════════════════════════════════════════════════════════════════
# CONSTANTS
# ═════════════════════════════════════════════════════════════════════════════

MODEL          = os.environ.get("LLM_MODEL", "qwen/qwen3-5-27b")
LLM_BASE_URL   = os.environ.get("LLM_BASE_URL", "https://maas-llm-aiplatform-hcm.api.vngcloud.vn/v1")
MAX_CONTRACT   = 50_000    # chars — trimmed to stay under 60s gateway timeout
MAX_EMAIL      = 20_000    # chars
MAX_TOKENS_EXT = 4_000
MAX_TOKENS_CMP = 2_000

# Colour palette
C = {
    "dark_blue":   "1F4E79",
    "mid_blue":    "2E75B6",
    "light_blue":  "D6E4F0",
    "very_light":  "EBF3FB",
    "white":       "FFFFFF",
    "red_head":    "C00000",
    "red_light":   "FDECEA",
    "yellow":      "FFF2CC",
    "green_light": "E2EFDA",
    "gray":        "F2F2F2",
}

# ═════════════════════════════════════════════════════════════════════════════
# PROMPTS
# ═════════════════════════════════════════════════════════════════════════════

SYSTEM_PROMPT = """Bạn là chuyên gia phân tích hợp đồng cho bộ phận FP&A của Zalopay.
Quy tắc bắt buộc:
- Trả về JSON hợp lệ DUY NHẤT, không có giải thích nào ngoài JSON.
- Rephrase nội dung concisely nhưng KHÔNG thay đổi ý nghĩa.
- Nếu không tìm thấy: dùng null. Nếu không rõ ràng: ghi "Cần xác nhận thêm".
- Số tiền/tỷ lệ: ghi đầy đủ đơn vị và ký hiệu (VND, USD, %).
- Điều khoản pháp lý: trích ngắn gọn, đủ ý."""

EXTRACTION_PROMPT = """\
Phân tích hợp đồng/phụ lục dưới đây và trả về JSON theo cấu trúc sau:

{{
  "loai_hop_dong": "string từ tiêu đề",

  "cac_ben": [
    {{
      "ten_ben":       "Bên A / Bên B / tên gọi trong HĐ",
      "ten_cong_ty":   "tên đầy đủ",
      "nguoi_dai_dien":"họ tên",
      "chuc_vu":       "chức vụ",
      "dia_chi":       "địa chỉ hoặc null",
      "ma_so_thue":    "MST hoặc null"
    }}
  ],

  "thoi_han_hop_dong": {{
    "ngay_ky":           "DD/MM/YYYY hoặc null",
    "ngay_hieu_luc":     "DD/MM/YYYY hoặc null",
    "thoi_gian_hieu_luc":"VD: 12 tháng kể từ ngày hiệu lực",
    "ngay_het_han":      "DD/MM/YYYY hoặc null",
    "dieu_kien_gia_han": "mô tả gia hạn tự động/thủ công, thời hạn thông báo v.v."
  }},

  "dich_vu_hop_tac": {{
    "mo_ta_chung":      "mô tả ngành nghề và phạm vi hợp tác",
    "doi_tuong_tich_hop":"đối tượng/sản phẩm tích hợp",
    "kenh_thanh_toan":  ["Zalopay App", "Gateway", "VietQR", "..."],
    "nguon_tien":       ["Số dư ví", "Tài khoản ngân hàng", "Thẻ nội địa", "Thẻ quốc tế", "..."]
  }},

  "commercial_terms": {{
    "tong_gia_tri_hop_dong": "số tiền + đơn vị hoặc null",

    "phi_dich_vu": [
      {{
        "loai_phi":       "tên loại phí (dịch vụ / hoàn trả / chia sẻ / thúc đẩy DS / ...)",
        "kenh_thanh_toan":"kênh áp dụng",
        "nguon_tien":     "nguồn tiền áp dụng",
        "muc_phi":        "con số / tỷ lệ cụ thể",
        "dieu_kien":      "điều kiện áp dụng ngắn gọn",
        "ghi_chu":        "thông tin bổ sung hoặc null"
      }}
    ],

    "ngan_sach_khuyen_mai": {{
      "tong_ngan_sach":    "số tiền hoặc null",
      "the_le_dieu_kien":  "mô tả thể lệ và điều kiện áp dụng hoặc null"
    }},

    "lai_va_phat": {{
      "lai_tra_cham":          "mức lãi và cách tính hoặc null",
      "phat_vi_pham":          "mức phạt hoặc null",
      "boi_thuong_thiet_hai":  "quy định bồi thường hoặc null"
    }},

    "payment_term": {{
      "co_che_thanh_toan":        "cấn trừ D+1 / đối tác TT sau X ngày từ HĐ đối soát / ...",
      "tam_ung_thanh_toan_truoc": "số tiền hoặc null",
      "cong_no_thanh_toan":       "X ngày thường / X ngày làm việc",
      "ho_so_thanh_toan":         ["hóa đơn VAT", "biên bản đối soát", "biên bản nghiệm thu", "..."],
      "alert_ho_so":              "CẢNH BÁO nếu thiếu hóa đơn / biên bản đối soát / nghiệm thu, hoặc null"
    }},

    "reconciliation_term": {{
      "thoi_gian_bat_dau_doi_soat":"mô tả",
      "thoi_gian_gui_doi_soat":    "mô tả",
      "thoi_gian_phan_hoi":        "mô tả",
      "xu_ly_chenh_lech":          "quy trình xử lý chênh lệch",
      "zalopay_xuat_hoa_don":      "X ngày kể từ xác nhận đối soát hoặc null"
    }}
  }},

  "truong_con_thieu": ["trường quan trọng không tìm thấy trong HĐ"]
}}

=== NỘI DUNG TÀI LIỆU ===
{content}
"""

COMPARISON_PROMPT = """\
So sánh nội dung aligned trong EMAIL với nội dung trong HỢP ĐỒNG.
Trả về JSON:

{{
  "ket_qua_so_sanh": [
    {{
      "diem_so_sanh":  "điểm so sánh (VD: Mức phí dịch vụ Gateway)",
      "trong_email":   "nội dung trong email",
      "trong_hop_dong":"nội dung trong HĐ",
      "trang_thai":    "KHỚP | KHÁC BIỆT | CHỈ TRONG EMAIL | CHỈ TRONG HĐ",
      "ghi_chu":       "giải thích ngắn nếu có sai lệch"
    }}
  ],
  "tong_ket": "nhận xét tổng quan mức độ khớp"
}}

=== NỘI DUNG EMAIL ===
{email_content}

=== TÓM TẮT HỢP ĐỒNG ===
{contract_summary}
"""


# ═════════════════════════════════════════════════════════════════════════════
# READERS
# ═════════════════════════════════════════════════════════════════════════════

def read_docx(path: str) -> str:
    """Trả về toàn bộ text từ .docx, bao gồm nội dung bảng."""
    doc = Document(path)
    parts: list[str] = []
    for block in doc.element.body:
        tag = block.tag.split("}")[-1]
        if tag == "p":
            text = "".join(
                node.text or "" for node in block.iter() if node.tag == qn("w:t")
            )
            if text.strip():
                parts.append(text)
        elif tag == "tbl":
            tbl = DocxTable(block, doc)
            for row in tbl.rows:
                row_text = " | ".join(cell.text.strip() for cell in row.cells)
                if row_text.strip(" |"):
                    parts.append(f"[TABLE] {row_text}")
    return "\n".join(parts)


def read_eml(path: str) -> str:
    """Trả về nội dung email (.eml) dạng plain text."""
    with open(path, "rb") as fh:
        msg = email_lib.message_from_bytes(fh.read(), policy=email_policy.default)

    subject = msg.get("Subject", "")
    sender  = msg.get("From", "")
    to      = msg.get("To", "")
    date    = msg.get("Date", "")
    body    = ""

    if msg.is_multipart():
        for part in msg.walk():
            ct = part.get_content_type()
            if ct == "text/plain":
                body += part.get_content() or ""
            elif ct == "text/html" and not body:
                html = part.get_content() or ""
                body += re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", html))
    else:
        body = msg.get_content() or ""

    return (
        f"[EMAIL]\nTừ: {sender}\nGửi đến: {to}\nNgày: {date}\n"
        f"Tiêu đề: {subject}\n\nNội dung:\n{body}"
    )


def read_pdf(path: str) -> str:
    """Trả về toàn bộ text từ .pdf, bao gồm nội dung bảng."""
    if not _PDF_AVAILABLE:
        raise RuntimeError("pdfplumber chưa được cài: pip install pdfplumber")
    parts: list[str] = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=2, y_tolerance=2) or ""
            if text.strip():
                parts.append(text)
            for table in page.extract_tables():
                for row in table:
                    row_text = " | ".join(cell or "" for cell in row)
                    if row_text.strip(" |"):
                        parts.append(f"[TABLE] {row_text}")
    return "\n".join(parts)


def read_msg(path: str) -> str:
    """Trả về nội dung email Outlook (.msg) dạng plain text."""
    if not _MSG_AVAILABLE:
        raise RuntimeError("extract-msg chưa được cài: pip install extract-msg")
    m = _extract_msg.openMsg(path)
    subject = m.subject or ""
    sender  = m.sender or ""
    to      = m.to or ""
    date    = str(m.date or "")
    body    = m.body or ""
    if body.lstrip().startswith("<"):
        body = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", body))
    return (
        f"[EMAIL]\nTừ: {sender}\nGửi đến: {to}\nNgày: {date}\n"
        f"Tiêu đề: {subject}\n\nNội dung:\n{body}"
    )


def collect_files(
    contracts: list[str],
    emails: list[str],
    folder: str | None,
) -> tuple[dict[str, str], dict[str, str]]:
    """Đọc tất cả file, trả về (contract_texts, email_texts)."""
    if folder:
        p = Path(folder)
        contracts = contracts + [str(f) for f in sorted(p.glob("*.docx"))] \
                              + [str(f) for f in sorted(p.glob("*.pdf"))]
        emails    = emails    + [str(f) for f in sorted(p.glob("*.eml"))] \
                              + [str(f) for f in sorted(p.glob("*.msg"))]

    contract_texts: dict[str, str] = {}
    email_texts:    dict[str, str] = {}

    for fp in contracts:
        name = Path(fp).name
        ext  = Path(fp).suffix.lower()
        try:
            contract_texts[name] = read_pdf(fp) if ext == ".pdf" else read_docx(fp)
            print(f"  ✓ HĐ/PL : {name}")
        except Exception as exc:
            print(f"  ✗ Lỗi   : {name} — {exc}")

    for fp in emails:
        name = Path(fp).name
        ext  = Path(fp).suffix.lower()
        try:
            email_texts[name] = read_msg(fp) if ext == ".msg" else read_eml(fp)
            print(f"  ✓ Email : {name}")
        except Exception as exc:
            print(f"  ✗ Lỗi   : {name} — {exc}")

    return contract_texts, email_texts


# ═════════════════════════════════════════════════════════════════════════════
# EXTRACTORS
# ═════════════════════════════════════════════════════════════════════════════

def _call_llm(prompt: str, api_key: str, max_tokens: int) -> dict:
    """Gọi LLM API (OpenAI-compatible), parse JSON từ response."""
    client = OpenAI(api_key=api_key, base_url=LLM_BASE_URL)
    response = client.chat.completions.create(
        model=MODEL,
        max_tokens=max_tokens,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user",   "content": prompt},
        ],
        extra_body={"chat_template_kwargs": {"enable_thinking": False}},
    )
    raw = response.choices[0].message.content.strip()
    # Strip markdown code fences nếu có
    if "```" in raw:
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    return json.loads(raw.strip())


def extract_contract(texts: dict[str, str], api_key: str) -> dict:
    """Trích xuất thông tin có cấu trúc từ hợp đồng."""
    full_content = ""
    for name, text in texts.items():
        full_content += f"\n\n{'='*60}\nTÀI LIỆU: {name}\n{'='*60}\n{text}"
    prompt = EXTRACTION_PROMPT.format(content=full_content[:MAX_CONTRACT])
    print("  → Trích xuất thông tin HĐ...")
    return _call_llm(prompt, api_key, MAX_TOKENS_EXT)


def compare_email(
    email_texts: dict[str, str],
    contract_data: dict,
    api_key: str,
) -> dict:
    """So sánh nội dung email alignment với hợp đồng."""
    email_content    = "\n\n".join(email_texts.values())[:MAX_EMAIL]
    contract_summary = json.dumps(contract_data, ensure_ascii=False, indent=2)[:30_000]
    prompt = COMPARISON_PROMPT.format(
        email_content=email_content,
        contract_summary=contract_summary,
    )
    print("  → So sánh email vs HĐ...")
    return _call_llm(prompt, api_key, MAX_TOKENS_CMP)


# ═════════════════════════════════════════════════════════════════════════════
# EXCEL HELPERS
# ═════════════════════════════════════════════════════════════════════════════

def _fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)

def _thin_border(color: str = "B0C4D8") -> Border:
    s = Side(style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def _font(bold=False, size=10, color="000000", italic=False) -> Font:
    return Font(bold=bold, size=size, color=color, italic=italic)

def _align(h="left", v="top", wrap=True) -> Alignment:
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)


def _header(cell, text: str, bg=C["dark_blue"], size=13):
    cell.value = text
    cell.font  = _font(bold=True, size=size, color=C["white"])
    cell.fill  = _fill(bg)
    cell.alignment = _align(h="center", v="center")

def _section(cell, text: str, bg=C["mid_blue"]):
    cell.value = text
    cell.font  = _font(bold=True, size=10, color=C["white"])
    cell.fill  = _fill(bg)
    cell.alignment = _align(h="left", v="center")
    cell.border = _thin_border()

def _sub_section(cell, text: str):
    cell.value = text
    cell.font  = _font(bold=True, size=10, color=C["dark_blue"])
    cell.fill  = _fill(C["light_blue"])
    cell.alignment = _align(h="left", v="center")
    cell.border = _thin_border()

def _label(cell, text: str):
    cell.value = text
    cell.font  = _font(bold=True, size=10)
    cell.fill  = _fill(C["very_light"])
    cell.alignment = _align()
    cell.border = _thin_border()

def _value(cell, text, alert=False):
    cell.value = text or ""
    cell.font  = _font(size=10, color=C["red_head"] if alert else "000000")
    cell.fill  = _fill(C["red_light"] if alert else C["white"])
    cell.alignment = _align()
    cell.border = _thin_border("C00000" if alert else "B0C4D8")

def _tbl_header(cell, text: str):
    cell.value = text
    cell.font  = _font(bold=True, size=9, color=C["white"])
    cell.fill  = _fill(C["mid_blue"])
    cell.alignment = _align(h="center", v="center")
    cell.border = _thin_border()

def _tbl_cell(cell, text, alt=False, red=False):
    cell.value = text or ""
    cell.font  = _font(size=9, color=C["red_head"] if red else "000000")
    cell.fill  = _fill(C["red_light"] if red else (C["gray"] if alt else C["white"]))
    cell.alignment = _align()
    cell.border = _thin_border()


def _merge_section(ws, row: int, text: str, n_cols=2, bg=C["mid_blue"]) -> int:
    ws.merge_cells(f"A{row}:{get_column_letter(n_cols)}{row}")
    _section(ws.cell(row, 1), text, bg=bg)
    ws.row_dimensions[row].height = 22
    return row + 1


def _write_row(ws, row: int, label: str, value, alert=False, height: int | None = None):
    _label(ws.cell(row, 1), label)
    _value(ws.cell(row, 2), value, alert=alert)
    ws.row_dimensions[row].height = (
        height if height
        else (45 if value and len(str(value)) > 120 else 20)
    )


# ═════════════════════════════════════════════════════════════════════════════
# WRITERS
# ═════════════════════════════════════════════════════════════════════════════

def write_contract_sheet(ws, data: dict, source_files: list[str]) -> None:
    """Ghi toàn bộ thông tin hợp đồng vào 1 sheet duy nhất (6 cột)."""
    NCOLS = 6
    LAST  = "F"
    for col, w in zip("ABCDEF", [30, 28, 20, 20, 30, 28]):
        ws.column_dimensions[col].width = w

    def wr(row, label, value, alert=False, height=None):
        _label(ws.cell(row, 1), label)
        ws.merge_cells(f"B{row}:{LAST}{row}")
        _value(ws.cell(row, 2), value, alert=alert)
        ws.row_dimensions[row].height = (
            height if height else (45 if value and len(str(value)) > 120 else 20)
        )

    def sec(row, text, bg=C["mid_blue"]):
        ws.merge_cells(f"A{row}:{LAST}{row}")
        _section(ws.cell(row, 1), text, bg=bg)
        ws.row_dimensions[row].height = 22
        return row + 1

    row = 1

    # ── Tiêu đề ──
    ws.merge_cells(f"A{row}:{LAST}{row}")
    _header(ws.cell(row, 1), "BẢNG TRÍCH XUẤT THÔNG TIN HỢP ĐỒNG — ZALOPAY FP&A")
    ws.row_dimensions[row].height = 32; row += 1

    ws.merge_cells(f"A{row}:{LAST}{row}")
    c = ws.cell(row, 1, value=(
        f"📅 Ngày xử lý: {datetime.now().strftime('%d/%m/%Y %H:%M')}   |   "
        f"📄 Nguồn: {', '.join(Path(f).name for f in source_files)}"
    ))
    c.font = _font(italic=True, size=9, color="666666")
    c.alignment = _align(h="left", v="center", wrap=False)
    ws.row_dimensions[row].height = 16; row += 1

    missing = data.get("truong_con_thieu") or []
    if missing:
        ws.merge_cells(f"A{row}:{LAST}{row}")
        c = ws.cell(row, 1, value=f"⚠️  THÔNG TIN CÒN THIẾU: {' | '.join(missing)}")
        c.font  = _font(bold=True, size=10, color=C["white"])
        c.fill  = _fill(C["red_head"])
        c.alignment = _align(h="left", v="center")
        ws.row_dimensions[row].height = 22; row += 1
    row += 1

    # ── 1. Loại HĐ ──
    row = sec(row, "1.  LOẠI HỢP ĐỒNG")
    wr(row, "Loại hợp đồng", data.get("loai_hop_dong"), alert=not data.get("loai_hop_dong"))
    row += 2

    # ── 2. Các bên ──
    row = sec(row, "2.  CÁC BÊN THAM GIA")
    for party in data.get("cac_ben") or []:
        _label(ws.cell(row, 1), f"Bên: {party.get('ten_ben', '')}")
        ws.merge_cells(f"B{row}:{LAST}{row}")
        _value(ws.cell(row, 2), party.get("ten_cong_ty"))
        ws.row_dimensions[row].height = 18; row += 1
        wr(row, "  Người đại diện",
           f"{party.get('nguoi_dai_dien','')} — {party.get('chuc_vu','')}"); row += 1
        if party.get("ma_so_thue"):
            wr(row, "  MST", party.get("ma_so_thue")); row += 1
    row += 1

    # ── 3. Thời hạn ──
    row = sec(row, "3.  THỜI HẠN HỢP ĐỒNG")
    t = data.get("thoi_han_hop_dong") or {}
    for lbl, key in [
        ("Ngày ký",            "ngay_ky"),
        ("Ngày hiệu lực",      "ngay_hieu_luc"),
        ("Thời gian hiệu lực", "thoi_gian_hieu_luc"),
        ("Ngày hết hạn",       "ngay_het_han"),
        ("Điều kiện gia hạn",  "dieu_kien_gia_han"),
    ]:
        wr(row, lbl, t.get(key), alert=not t.get(key)); row += 1
    row += 1

    # ── 4. Dịch vụ ──
    row = sec(row, "4.  DỊCH VỤ HỢP TÁC")
    dv = data.get("dich_vu_hop_tac") or {}
    wr(row, "Mô tả chung",        dv.get("mo_ta_chung"),        height=55); row += 1
    wr(row, "Đối tượng tích hợp", dv.get("doi_tuong_tich_hop"));            row += 1
    wr(row, "Kênh thanh toán",
       ", ".join(dv.get("kenh_thanh_toan") or []) or None,
       alert=not dv.get("kenh_thanh_toan")); row += 1
    wr(row, "Nguồn tiền",
       ", ".join(dv.get("nguon_tien") or []) or None,
       alert=not dv.get("nguon_tien")); row += 1
    row += 1

    # ── 5. Tổng giá trị ──
    row = sec(row, "5.  TỔNG GIÁ TRỊ HỢP ĐỒNG")
    ct = data.get("commercial_terms") or {}
    _label(ws.cell(row, 1), "Tổng giá trị")
    ws.merge_cells(f"B{row}:{LAST}{row}")
    c = ws.cell(row, 2, value=ct.get("tong_gia_tri_hop_dong") or "Không xác định trong HĐ")
    c.font = _font(bold=True, size=11, color=C["dark_blue"])
    c.alignment = _align(h="left", v="center"); c.border = _thin_border()
    ws.row_dimensions[row].height = 22; row += 2

    # ── 6. Phí dịch vụ ──
    row = sec(row, "6.  PHÍ DỊCH VỤ (THEO KÊNH TT / NGUỒN TIỀN)")
    for col, hdr in enumerate(
        ["Loại phí","Kênh TT","Nguồn tiền","Mức phí","Điều kiện áp dụng","Ghi chú"], 1
    ):
        _tbl_header(ws.cell(row, col), hdr)
    ws.row_dimensions[row].height = 22; row += 1
    fees = ct.get("phi_dich_vu") or []
    if fees:
        for idx, fee in enumerate(fees):
            alt = idx % 2 == 1
            for col, key in enumerate(
                ["loai_phi","kenh_thanh_toan","nguon_tien","muc_phi","dieu_kien","ghi_chu"], 1
            ):
                _tbl_cell(ws.cell(row, col), fee.get(key), alt=alt)
            ws.row_dimensions[row].height = 35; row += 1
    else:
        ws.merge_cells(f"A{row}:{LAST}{row}")
        c = ws.cell(row, 1, value="⚠️  Không tìm thấy thông tin phí dịch vụ trong hợp đồng")
        c.font = _font(bold=True, color=C["red_head"], size=10)
        c.fill = _fill(C["red_light"]); c.border = _thin_border("C00000")
        ws.row_dimensions[row].height = 22; row += 1
    row += 1

    # ── 7. Ngân sách KM ──
    km = ct.get("ngan_sach_khuyen_mai") or {}
    if km.get("tong_ngan_sach") or km.get("the_le_dieu_kien"):
        row = sec(row, "7.  NGÂN SÁCH & THỂ LỆ KHUYẾN MẠI")
        for lbl, key in [("Tổng ngân sách","tong_ngan_sach"),
                         ("Thể lệ / Điều kiện","the_le_dieu_kien")]:
            _label(ws.cell(row, 1), lbl)
            ws.merge_cells(f"B{row}:{LAST}{row}")
            c = ws.cell(row, 2, value=km.get(key) or "")
            c.alignment = _align(); c.border = _thin_border()
            ws.row_dimensions[row].height = 50 if key == "the_le_dieu_kien" else 20; row += 1
        row += 1

    # ── 8. Lãi & Phạt ──
    lp = ct.get("lai_va_phat") or {}
    row = sec(row, "8.  LÃI TRẢ CHẬM & PHẠT VI PHẠM")
    for lbl, key in [("Lãi trả chậm","lai_tra_cham"),
                     ("Phạt vi phạm","phat_vi_pham"),
                     ("Bồi thường thiệt hại","boi_thuong_thiet_hai")]:
        _label(ws.cell(row, 1), lbl)
        ws.merge_cells(f"B{row}:{LAST}{row}")
        c = ws.cell(row, 2, value=lp.get(key) or "—")
        c.alignment = _align(); c.border = _thin_border()
        ws.row_dimensions[row].height = 35; row += 1
    row += 1

    # ── 9. Payment Term ──
    pt = ct.get("payment_term") or {}
    row = sec(row, "9.  PAYMENT TERM — ĐIỀU KHOẢN THANH TOÁN")
    for lbl, key, req in [
        ("Cơ chế thanh toán",  "co_che_thanh_toan",        True),
        ("Tạm ứng / TT trước", "tam_ung_thanh_toan_truoc", False),
        ("Công nợ thanh toán", "cong_no_thanh_toan",       True),
    ]:
        wr(row, lbl, pt.get(key), alert=(req and not pt.get(key))); row += 1
    ho_so = pt.get("ho_so_thanh_toan") or []
    wr(row, "Hồ sơ thanh toán", ", ".join(ho_so) if ho_so else None, alert=not ho_so); row += 1
    alert_msg = pt.get("alert_ho_so")
    if alert_msg:
        ws.merge_cells(f"A{row}:{LAST}{row}")
        c = ws.cell(row, 1, value=f"⚠️  {alert_msg}")
        c.font = _font(bold=True, color=C["white"], size=10)
        c.fill = _fill(C["red_head"])
        c.alignment = _align(h="left", v="center"); c.border = _thin_border()
        ws.row_dimensions[row].height = 22; row += 1
    row += 1

    # ── 10. Reconciliation ──
    rt = ct.get("reconciliation_term") or {}
    row = sec(row, "10. RECONCILIATION TERM — ĐIỀU KHOẢN ĐỐI SOÁT")
    for lbl, key in [
        ("Bắt đầu đối soát",    "thoi_gian_bat_dau_doi_soat"),
        ("Gửi đối soát",         "thoi_gian_gui_doi_soat"),
        ("Thời hạn phản hồi",    "thoi_gian_phan_hoi"),
        ("Xử lý chênh lệch",     "xu_ly_chenh_lech"),
        ("Zalopay xuất hóa đơn", "zalopay_xuat_hoa_don"),
    ]:
        wr(row, lbl, rt.get(key), alert=not rt.get(key)); row += 1


def write_summary_sheet(ws, data: dict, source_files: list[str]):
    ws.title = "📋 Tóm tắt HĐ"
    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 72
    row = 1

    # ── Tiêu đề ──
    ws.merge_cells(f"A{row}:B{row}")
    _header(ws.cell(row, 1), "BẢNG TRÍCH XUẤT THÔNG TIN HỢP ĐỒNG — ZALOPAY FP&A")
    ws.row_dimensions[row].height = 32
    row += 1

    ws.merge_cells(f"A{row}:B{row}")
    c = ws.cell(row, 1,
        value=(f"📅 Ngày xử lý: {datetime.now().strftime('%d/%m/%Y %H:%M')}   |   "
               f"📄 Nguồn: {', '.join(Path(f).name for f in source_files)}"))
    c.font = _font(italic=True, size=9, color="666666")
    c.alignment = _align(h="left", v="center", wrap=False)
    ws.row_dimensions[row].height = 16
    row += 1

    # ── Alert thiếu thông tin ──
    missing = data.get("truong_con_thieu") or []
    if missing:
        ws.merge_cells(f"A{row}:B{row}")
        c = ws.cell(row, 1, value=f"⚠️  THÔNG TIN CÒN THIẾU: {' | '.join(missing)}")
        c.font  = _font(bold=True, size=10, color=C["white"])
        c.fill  = _fill(C["red_head"])
        c.alignment = _align(h="left", v="center")
        ws.row_dimensions[row].height = 22
        row += 1
    row += 1

    # ── 1. Loại HĐ ──
    row = _merge_section(ws, row, "1.  LOẠI HỢP ĐỒNG")
    _write_row(ws, row, "Loại hợp đồng", data.get("loai_hop_dong"),
               alert=not data.get("loai_hop_dong"))
    row += 2

    # ── 2. Các Bên ──
    row = _merge_section(ws, row, "2.  CÁC BÊN THAM GIA")
    for party in data.get("cac_ben") or []:
        _label(ws.cell(row, 1), f"Bên: {party.get('ten_ben', '')}")
        _value(ws.cell(row, 2), party.get("ten_cong_ty"))
        ws.row_dimensions[row].height = 18
        row += 1
        _write_row(ws, row, "  Người đại diện",
                   f"{party.get('nguoi_dai_dien', '')} — {party.get('chuc_vu', '')}")
        row += 1
        if party.get("ma_so_thue"):
            _write_row(ws, row, "  MST", party.get("ma_so_thue"))
            row += 1
    row += 1

    # ── 3. Thời hạn ──
    row = _merge_section(ws, row, "3.  THỜI HẠN HỢP ĐỒNG")
    t = data.get("thoi_han_hop_dong") or {}
    for label, key in [
        ("Ngày ký",            "ngay_ky"),
        ("Ngày hiệu lực",      "ngay_hieu_luc"),
        ("Thời gian hiệu lực", "thoi_gian_hieu_luc"),
        ("Ngày hết hạn",       "ngay_het_han"),
        ("Điều kiện gia hạn",  "dieu_kien_gia_han"),
    ]:
        _write_row(ws, row, label, t.get(key), alert=not t.get(key))
        row += 1
    row += 1

    # ── 4. Dịch vụ hợp tác ──
    row = _merge_section(ws, row, "4.  DỊCH VỤ HỢP TÁC")
    dv = data.get("dich_vu_hop_tac") or {}
    _write_row(ws, row, "Mô tả chung",       dv.get("mo_ta_chung"),       height=55); row += 1
    _write_row(ws, row, "Đối tượng tích hợp",dv.get("doi_tuong_tich_hop")); row += 1
    _write_row(ws, row, "Kênh thanh toán",
               ", ".join(dv.get("kenh_thanh_toan") or []) or None,
               alert=not dv.get("kenh_thanh_toan")); row += 1
    _write_row(ws, row, "Nguồn tiền",
               ", ".join(dv.get("nguon_tien") or []) or None,
               alert=not dv.get("nguon_tien")); row += 1
    row += 1

    # ── 5. Tổng giá trị ──
    row = _merge_section(ws, row, "5.  TỔNG GIÁ TRỊ HỢP ĐỒNG")
    ct  = data.get("commercial_terms") or {}
    _label(ws.cell(row, 1), "Tổng giá trị")
    c = ws.cell(row, 2, value=ct.get("tong_gia_tri_hop_dong") or "Không xác định trong HĐ")
    c.font = _font(bold=True, size=11, color=C["dark_blue"])
    c.alignment = _align(h="left", v="center")
    c.border = _thin_border()
    ws.row_dimensions[row].height = 22
    row += 1


def write_commercial_sheet(ws, data: dict):
    ws.title = "💰 Commercial Terms"
    ct = data.get("commercial_terms") or {}

    col_letters = "ABCDEF"
    col_widths  = [24, 28, 20, 20, 30, 28]
    for letter, width in zip(col_letters, col_widths):
        ws.column_dimensions[letter].width = width

    row = 1
    ws.merge_cells(f"A{row}:F{row}")
    _header(ws.cell(row, 1), "COMMERCIAL TERMS — PHÍ & ĐIỀU KIỆN THƯƠNG MẠI")
    ws.row_dimensions[row].height = 28
    row += 2

    # ── Bảng phí dịch vụ ──
    ws.merge_cells(f"A{row}:F{row}")
    _section(ws.cell(row, 1), "PHÍ DỊCH VỤ (CHI TIẾT THEO KÊNH TT / NGUỒN TIỀN)")
    ws.row_dimensions[row].height = 22; row += 1

    for col, header in enumerate(
        ["Loại phí", "Kênh TT", "Nguồn tiền", "Mức phí", "Điều kiện áp dụng", "Ghi chú"], 1
    ):
        _tbl_header(ws.cell(row, col), header)
    ws.row_dimensions[row].height = 22; row += 1

    fees = ct.get("phi_dich_vu") or []
    if fees:
        for idx, fee in enumerate(fees):
            alt = idx % 2 == 1
            for col, key in enumerate(
                ["loai_phi","kenh_thanh_toan","nguon_tien","muc_phi","dieu_kien","ghi_chu"], 1
            ):
                _tbl_cell(ws.cell(row, col), fee.get(key), alt=alt)
            ws.row_dimensions[row].height = 35; row += 1
    else:
        ws.merge_cells(f"A{row}:F{row}")
        c = ws.cell(row, 1, value="⚠️  Không tìm thấy thông tin phí dịch vụ trong hợp đồng")
        c.font = _font(bold=True, color=C["red_head"], size=10)
        c.fill = _fill(C["red_light"])
        c.border = _thin_border("C00000")
        ws.row_dimensions[row].height = 22; row += 1
    row += 1

    # ── Ngân sách KM ──
    km = ct.get("ngan_sach_khuyen_mai") or {}
    if km.get("tong_ngan_sach") or km.get("the_le_dieu_kien"):
        ws.merge_cells(f"A{row}:F{row}")
        _section(ws.cell(row, 1), "NGÂN SÁCH & THỂ LỆ KHUYẾN MẠI")
        ws.row_dimensions[row].height = 22; row += 1
        for label, key in [("Tổng ngân sách","tong_ngan_sach"),("Thể lệ / Điều kiện","the_le_dieu_kien")]:
            _label(ws.cell(row, 1), label)
            ws.merge_cells(f"B{row}:F{row}")
            c = ws.cell(row, 2, value=km.get(key) or "")
            c.alignment = _align(); c.border = _thin_border()
            h = 50 if key == "the_le_dieu_kien" else 20
            ws.row_dimensions[row].height = h; row += 1
        row += 1

    # ── Lãi & Phạt ──
    lp = ct.get("lai_va_phat") or {}
    ws.merge_cells(f"A{row}:F{row}")
    _section(ws.cell(row, 1), "LÃI TRẢ CHẬM & PHẠT VI PHẠM")
    ws.row_dimensions[row].height = 22; row += 1
    for label, key in [
        ("Lãi trả chậm",         "lai_tra_cham"),
        ("Phạt vi phạm",         "phat_vi_pham"),
        ("Bồi thường thiệt hại", "boi_thuong_thiet_hai"),
    ]:
        _label(ws.cell(row, 1), label)
        ws.merge_cells(f"B{row}:F{row}")
        c = ws.cell(row, 2, value=lp.get(key) or "—")
        c.alignment = _align(); c.border = _thin_border()
        ws.row_dimensions[row].height = 35; row += 1


def write_payment_sheet(ws, data: dict):
    ws.title = "🏦 Payment & Recon"
    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 72
    ct = data.get("commercial_terms") or {}
    pt = ct.get("payment_term")       or {}
    rt = ct.get("reconciliation_term") or {}
    row = 1

    ws.merge_cells(f"A{row}:B{row}")
    _header(ws.cell(row, 1), "PAYMENT TERM & RECONCILIATION TERM")
    ws.row_dimensions[row].height = 28; row += 2

    # ── Payment Term ──
    row = _merge_section(ws, row, "PAYMENT TERM — ĐIỀU KHOẢN THANH TOÁN")
    for label, key, required in [
        ("Cơ chế thanh toán",        "co_che_thanh_toan",        True),
        ("Tạm ứng / TT trước",       "tam_ung_thanh_toan_truoc", False),
        ("Công nợ thanh toán",        "cong_no_thanh_toan",       True),
    ]:
        val = pt.get(key)
        _write_row(ws, row, label, val, alert=(required and not val))
        row += 1

    ho_so = pt.get("ho_so_thanh_toan") or []
    _write_row(ws, row, "Hồ sơ thanh toán",
               ", ".join(ho_so) if ho_so else None, alert=not ho_so)
    row += 1

    alert_msg = pt.get("alert_ho_so")
    if alert_msg:
        ws.merge_cells(f"A{row}:B{row}")
        c = ws.cell(row, 1, value=f"⚠️  {alert_msg}")
        c.font = _font(bold=True, color=C["white"], size=10)
        c.fill = _fill(C["red_head"])
        c.alignment = _align(h="left", v="center")
        c.border = _thin_border()
        ws.row_dimensions[row].height = 22; row += 1
    row += 1

    # ── Reconciliation Term ──
    row = _merge_section(ws, row, "RECONCILIATION TERM — ĐIỀU KHOẢN ĐỐI SOÁT")
    for label, key in [
        ("Bắt đầu đối soát",       "thoi_gian_bat_dau_doi_soat"),
        ("Gửi đối soát",            "thoi_gian_gui_doi_soat"),
        ("Thời hạn phản hồi",       "thoi_gian_phan_hoi"),
        ("Xử lý chênh lệch",        "xu_ly_chenh_lech"),
        ("Zalopay xuất hóa đơn",    "zalopay_xuat_hoa_don"),
    ]:
        _write_row(ws, row, label, rt.get(key), alert=not rt.get(key))
        row += 1


def write_comparison_sheet(ws, comparison: dict):
    ws.title = "📧 Email vs HĐ"
    for col, width in zip("ABCDE", [28, 40, 40, 16, 32]):
        ws.column_dimensions[col].width = width
    row = 1

    ws.merge_cells(f"A{row}:E{row}")
    _header(ws.cell(row, 1), "SO SÁNH: EMAIL ALIGNMENT vs HỢP ĐỒNG")
    ws.row_dimensions[row].height = 28; row += 1

    summary = comparison.get("tong_ket", "")
    if summary:
        ws.merge_cells(f"A{row}:E{row}")
        c = ws.cell(row, 1, value=f"Nhận xét tổng quan: {summary}")
        c.font = _font(italic=True, size=10)
        c.fill = _fill(C["yellow"])
        c.alignment = _align(h="left", v="center")
        ws.row_dimensions[row].height = 40; row += 1
    row += 1

    for col, header in enumerate(
        ["Điểm so sánh", "Trong email", "Trong hợp đồng", "Trạng thái", "Ghi chú"], 1
    ):
        _tbl_header(ws.cell(row, col), header)
    ws.row_dimensions[row].height = 22; row += 1

    STATUS_BG = {
        "KHỚP":            C["green_light"],
        "KHÁC BIỆT":       C["red_light"],
        "CHỈ TRONG EMAIL": C["yellow"],
        "CHỈ TRONG HĐ":   C["very_light"],
    }

    for idx, item in enumerate(comparison.get("ket_qua_so_sanh") or []):
        status = item.get("trang_thai", "")
        bg = STATUS_BG.get(status, C["white"])
        alt = idx % 2 == 1
        values = [
            item.get("diem_so_sanh"),
            item.get("trong_email"),
            item.get("trong_hop_dong"),
            status,
            item.get("ghi_chu"),
        ]
        for col, val in enumerate(values, 1):
            c = ws.cell(row, col, value=val or "")
            c.font = _font(
                size=9, bold=(col == 4),
                color=C["red_head"] if status == "KHÁC BIỆT" else "000000"
            )
            c.fill = _fill(bg if col != 1 else (C["gray"] if alt else C["white"]))
            c.alignment = _align()
            c.border = _thin_border()
        ws.row_dimensions[row].height = 38; row += 1


def export_excel(
    results: list[dict],
    comparison: dict | None,
    output_path: str,
    source_files: list[str],
):
    """Xuất Excel: 1 sheet mỗi hợp đồng + 1 sheet so sánh email (nếu có)."""
    wb    = openpyxl.Workbook()
    multi = len(results) > 1

    for i, data in enumerate(results):
        fname = source_files[i] if i < len(source_files) else f"hop_dong_{i+1}.docx"
        stem  = Path(fname).stem[:28]

        ws = wb.active if i == 0 else wb.create_sheet()
        ws.title = (f"{i+1}. {stem}"[:31] if multi else "📋 Hợp đồng")
        write_contract_sheet(ws, data, [fname])

    if comparison:
        write_comparison_sheet(wb.create_sheet(), comparison)

    for ws in wb.worksheets:
        ws.freeze_panes = "A4"
    wb.save(output_path)
    print(f"  ✓ Lưu : {output_path}")


# ═════════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ═════════════════════════════════════════════════════════════════════════════

def get_api_key() -> str:
    key = os.environ.get("LLM_API_KEY", "")
    if not key:
        key = input("LLM_API_KEY: ").strip()
    if not key:
        sys.exit("Cần LLM_API_KEY.")
    return key


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Contract Review Agent — Zalopay FP&A",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    p.add_argument("--contracts", "-c", nargs="+", default=[],
                   metavar="FILE.docx",
                   help="Hợp đồng chính và phụ lục (.docx)")
    p.add_argument("--emails", "-e", nargs="+", default=[],
                   metavar="FILE.eml",
                   help="Email alignment để so sánh (.eml)")
    p.add_argument("--folder", "-f", default=None,
                   metavar="DIR",
                   help="Thư mục chứa toàn bộ .docx và .eml")
    p.add_argument("--output", "-o", default=None,
                   metavar="OUTPUT.xlsx",
                   help="Đường dẫn file Excel output")
    return p


def main():
    args = build_parser().parse_args()

    if not args.contracts and not args.folder:
        build_parser().print_help()
        sys.exit(0)

    print(f"\n{'═' * 55}")
    print("  CONTRACT REVIEW AGENT — ZALOPAY FP&A")
    print(f"{'═' * 55}\n")

    print("📂 Đọc file:\n")
    contract_texts, email_texts = collect_files(args.contracts, args.emails, args.folder)
    if not contract_texts:
        sys.exit("Không đọc được nội dung hợp đồng.")

    api_key = get_api_key()

    print("\n🤖 Phân tích:\n")
    fnames  = list(contract_texts.keys())
    results: list[dict] = []
    for fname, text in contract_texts.items():
        print(f"  → Phân tích: {fname}")
        r = extract_contract({fname: text}, api_key)
        results.append(r)
        missing = r.get("truong_con_thieu") or []
        if missing:
            print(f"     ⚠️  Thiếu: {', '.join(missing)}")

    comparison = None
    if email_texts:
        # So sánh email với hợp đồng đầu tiên (hoặc hợp đồng duy nhất)
        comparison = compare_email(email_texts, results[0], api_key)

    # Xác định output path
    if args.output:
        output_path = args.output
    else:
        first = Path(args.contracts[0]) if args.contracts else Path(fnames[0])
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = str(first.parent / f"ContractReview_{ts}.xlsx")

    print("\n📊 Xuất Excel:\n")
    export_excel(results, comparison, output_path, fnames)

    print(f"\n✅ Hoàn thành!\n   📁 {output_path}")
    if len(results) > 1:
        print(f"   📑 {len(results)} hợp đồng được phân tích riêng biệt → {len(results)*3} sheets")
    if comparison:
        n_diff = sum(
            1 for x in (comparison.get("ket_qua_so_sanh") or [])
            if x.get("trang_thai") == "KHÁC BIỆT"
        )
        if n_diff:
            print(f"   ⚠️  {n_diff} điểm khác biệt giữa email và HĐ (xem sheet 📧 Email vs HĐ)")
    print()


if __name__ == "__main__":
    main()
