# Contract Review Agent — Zalopay FP&A

AI agent tự động trích xuất thông tin quan trọng từ hợp đồng Word (.docx) và email alignment (.eml), xuất kết quả ra file Excel có format sẵn cho FP&A review.

---

## Thông tin được trích xuất

| # | Mục | Chi tiết |
|---|-----|----------|
| 1 | **Loại hợp đồng** | Từ tiêu đề hợp đồng |
| 2 | **Các Bên** | Tên công ty, người đại diện, chức vụ, MST |
| 3 | **Thời hạn HĐ** | Ngày ký, ngày hiệu lực, thời gian hiệu lực, ngày hết hạn, điều kiện gia hạn |
| 4 | **Dịch vụ hợp tác** | Mô tả dịch vụ, đối tượng tích hợp, kênh TT (ZLP App / Gateway / VietQR), nguồn tiền (ví / tài khoản NH / thẻ nội địa / thẻ QT...) |
| 5 | **Commercial Terms** | Bảng phí chi tiết theo kênh TT và nguồn tiền; ngân sách & thể lệ KM; lãi trả chậm; phạt vi phạm |
| 6 | **Payment Term** | Cơ chế thanh toán, tạm ứng, công nợ (ngày thường / ngày làm việc), hồ sơ TT |
| 7 | **Reconciliation Term** | Thời gian bắt đầu, gửi, phản hồi đối soát; xử lý chênh lệch; thời điểm xuất hóa đơn |
| 8 | **Tổng giá trị HĐ** | Nếu có |

**Tính năng bổ sung:**
- ⚠️ Banner cảnh báo liệt kê các thông tin còn thiếu
- 🔴 Ô đỏ đánh dấu từng trường bị thiếu
- 📧 Sheet so sánh Email Alignment vs Hợp đồng (KHỚP / KHÁC BIỆT / CHỈ TRONG EMAIL / CHỈ TRONG HĐ)

---

## Cài đặt

### Yêu cầu

- Python ≥ 3.11
- Anthropic API Key ([lấy tại console.anthropic.com](https://console.anthropic.com))

### Cài thư viện

```bash
pip install -r requirements.txt
```

### Cấu hình API Key

```bash
# Windows (PowerShell)
$env:ANTHROPIC_API_KEY = "sk-ant-..."

# Windows (Command Prompt)
set ANTHROPIC_API_KEY=sk-ant-...

# macOS / Linux
export ANTHROPIC_API_KEY=sk-ant-...
```

Nếu không set biến môi trường, agent sẽ hỏi key khi chạy.

---

## Sử dụng

### Chỉ hợp đồng (không có email)

```bash
python agent.py --contracts hop_dong.docx phu_luc_1.docx phu_luc_2.docx
```

### Hợp đồng + email để so sánh

```bash
python agent.py --contracts hop_dong.docx phu_luc.docx --emails aligned.eml
```

### Toàn bộ file trong một thư mục

```bash
python agent.py --folder ./ho_so_hop_dong/
```
Agent sẽ tự nhận tất cả `.docx` làm hợp đồng và `.eml` làm email.

### Chỉ định tên file output

```bash
python agent.py --contracts hop_dong.docx --output review_Q3_2026.xlsx
```

### Xem tất cả options

```bash
python agent.py --help
```

---

## Output Excel

File Excel được tạo cùng thư mục với hợp đồng, tên tự động: `ContractReview_YYYYMMDD_HHMMSS.xlsx`

| Sheet | Nội dung |
|-------|----------|
| 📋 Tóm tắt HĐ | Loại HĐ, Các bên, Thời hạn, Dịch vụ, Tổng giá trị; banner alert thiếu thông tin |
| 💰 Commercial Terms | Bảng phí chi tiết, ngân sách KM, lãi & phạt |
| 🏦 Payment & Recon | Cơ chế TT, tạm ứng, công nợ, hồ sơ TT, toàn bộ reconciliation term |
| 📧 Email vs HĐ | So sánh từng điểm giữa email alignment và hợp đồng *(chỉ có khi truyền file .eml)* |

---

## Chạy bằng Docker

### Build image

```bash
docker build -t contract-agent .
```

### Chạy với thư mục chứa hồ sơ

```bash
docker run --rm \
  -e ANTHROPIC_API_KEY=sk-ant-... \
  -v "$(pwd)/ho_so":/data \
  contract-agent --folder /data --output /data/review.xlsx
```

### Chạy với file cụ thể

```bash
docker run --rm \
  -e ANTHROPIC_API_KEY=sk-ant-... \
  -v "$(pwd)":/data \
  contract-agent \
  --contracts /data/hop_dong.docx /data/phu_luc.docx \
  --emails /data/aligned.eml \
  --output /data/ContractReview.xlsx
```

---

## Chi phí API (ước tính)

| Kích thước | Tokens ước tính | Chi phí (Claude Opus 4.8) |
|------------|-----------------|--------------------------|
| HĐ nhỏ < 10 trang | ~3 000 | ~$0.05 |
| HĐ vừa 10–30 trang | ~8 000 | ~$0.12 |
| HĐ + nhiều PL > 30 trang | ~20 000 | ~$0.30 |
| + so sánh email | thêm ~5 000 | thêm ~$0.08 |

---

## Rule cứng

> Tất cả điều khoản lấy ra được **rephrase concisely** nhưng **không làm thay đổi ý nghĩa** gây confusing cho người đọc.

Điều này được enforce trực tiếp trong system prompt gửi đến Claude.
