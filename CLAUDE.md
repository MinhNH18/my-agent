# CLAUDE.md — Contract Review Agent

## Project overview

AI agent trích xuất thông tin hợp đồng cho bộ phận **FP&A của Zalopay**.  
Đọc file Word (.docx) và email (.eml), gọi Claude API để phân tích, xuất kết quả ra Excel (.xlsx).

---

## File structure

```
.
├── agent.py              # Entry point — orchestrates toàn bộ pipeline
├── requirements.txt      # Python dependencies
├── Dockerfile            # Container image
├── README.md             # Hướng dẫn người dùng cuối
└── CLAUDE.md             # File này — context cho Claude
```

---

## Architecture

```
agent.py
├── readers/
│   ├── read_docx(path)       → str   — đọc Word, giữ cấu trúc bảng
│   └── read_eml(path)        → str   — đọc email, strip HTML
│
├── extractors/
│   ├── extract_contract(texts, api_key)   → dict  — gọi Claude để trích xuất HĐ
│   └── compare_email(emails, data, key)   → dict  — gọi Claude để so sánh email vs HĐ
│
└── writers/
    ├── write_summary_sheet(ws, data)      — Sheet 1: Tóm tắt HĐ
    ├── write_commercial_sheet(ws, data)   — Sheet 2: Commercial Terms
    ├── write_payment_sheet(ws, data)      — Sheet 3: Payment & Reconciliation
    ├── write_comparison_sheet(ws, cmp)    — Sheet 4: Email vs HĐ (nếu có email)
    └── export_excel(data, cmp, path, ...)
```

---

## Output fields (JSON schema)

Claude trả về JSON với các trường:

| Trường | Mô tả |
|--------|--------|
| `loai_hop_dong` | Loại HĐ từ tiêu đề |
| `cac_ben[]` | Tên công ty, người đại diện, MST |
| `thoi_han_hop_dong` | Ngày ký, hiệu lực, hết hạn, gia hạn |
| `dich_vu_hop_tac` | Dịch vụ, kênh TT, nguồn tiền |
| `commercial_terms.phi_dich_vu[]` | Bảng phí theo kênh TT / nguồn tiền |
| `commercial_terms.ngan_sach_khuyen_mai` | Ngân sách & thể lệ KM |
| `commercial_terms.lai_va_phat` | Lãi trả chậm, phạt, bồi thường |
| `commercial_terms.payment_term` | Cơ chế TT, công nợ, hồ sơ |
| `commercial_terms.reconciliation_term` | Đối soát, xuất hóa đơn |
| `truong_con_thieu[]` | Danh sách trường thiếu (dùng để alert) |

---

## Claude API usage

- **Model**: `claude-opus-4-8`
- **Max tokens**: 6 000 (extraction), 4 000 (comparison)
- **Input limit**: 130 000 ký tự văn bản HĐ; 40 000 ký tự email
- **System prompt**: Chuyên gia FP&A phân tích hợp đồng — rephrase concise, không đổi ý nghĩa, null nếu thiếu

---

## Conventions

- Hàm reader trả về `str` thuần, không parse.
- Hàm extractor nhận `dict[str, str]` (filename → content) và trả `dict`.
- Hàm writer nhận `openpyxl.Worksheet` và mutate in-place.
- Alert (ô đỏ) được áp dụng khi field là `None` hoặc list rỗng.
- Missing fields được liệt kê trong banner đỏ ở đầu Sheet 1.

---

## Environment variables

| Biến | Bắt buộc | Mô tả |
|------|----------|--------|
| `ANTHROPIC_API_KEY` | ✅ | API key để gọi Claude |

---

## Common tasks

**Thêm trường trích xuất mới:**
1. Thêm field vào `EXTRACTION_PROMPT` trong `agent.py`
2. Thêm `write_section_row` tương ứng trong hàm `write_*_sheet`

**Đổi model:**
Tìm `claude-opus-4-8` trong `agent.py`, thay bằng model khác.

**Tăng giới hạn input:**
Tìm `[:130000]` trong `extract_contract()`, điều chỉnh theo context window của model.
