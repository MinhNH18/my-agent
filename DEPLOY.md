# Hướng dẫn Deploy Public API

Agent được expose qua FastAPI. Sau khi deploy, bạn có ngay:
- `POST /analyze` — upload docx/eml, nhận file Excel
- `GET  /health`  — health check
- `GET  /docs`    — Swagger UI để test trực tiếp trên browser

---

## Chuẩn bị trước khi deploy

Cần 2 biến môi trường:

| Biến | Bắt buộc | Mô tả |
|------|----------|-------|
| `ANTHROPIC_API_KEY` | ✅ | Key gọi Claude |
| `API_SECRET_KEY` | Khuyến nghị | Bearer token bảo vệ API (nếu bỏ trống: API public hoàn toàn) |

---

## Option 1 — Railway (khuyến nghị, dễ nhất)

**Free tier:** 500 giờ/tháng, tự động HTTPS, custom domain.

### Bước 1: Push code lên GitHub
```bash
git init
git add .
git commit -m "init contract review agent"
git remote add origin https://github.com/<you>/<repo>.git
git push -u origin main
```

### Bước 2: Deploy trên Railway
1. Vào [railway.app](https://railway.app) → **New Project** → **Deploy from GitHub repo**
2. Chọn repo vừa push
3. Railway tự detect Dockerfile và build

### Bước 3: Cấu hình biến môi trường
Vào **Settings → Variables**, thêm:
```
ANTHROPIC_API_KEY = sk-ant-...
API_SECRET_KEY    = chon-mot-password-manh
```

### Bước 4: Lấy URL
Railway tạo URL dạng: `https://your-app.up.railway.app`

Swagger UI: `https://your-app.up.railway.app/docs`

---

## Option 2 — Render

**Free tier:** Auto-sleep sau 15 phút idle (cold start ~30s), tự động HTTPS.

### Bước 1: Push code lên GitHub (như trên)

### Bước 2: Deploy trên Render
1. Vào [render.com](https://render.com) → **New** → **Web Service**
2. Connect GitHub repo
3. Chọn **Environment: Docker**
4. **Start Command** để trống (Dockerfile đã có CMD)

### Bước 3: Biến môi trường
Vào **Environment**, thêm:
```
ANTHROPIC_API_KEY = sk-ant-...
API_SECRET_KEY    = chon-mot-password-manh
```

---

## Option 3 — Fly.io

**Free tier:** 3 shared-cpu VM, tự động HTTPS, deploy nhanh.

```bash
# Cài flyctl
curl -L https://fly.io/install.sh | sh

# Login và deploy
fly auth login
fly launch          # tự detect Dockerfile, tạo fly.toml
fly secrets set ANTHROPIC_API_KEY=sk-ant-...
fly secrets set API_SECRET_KEY=chon-mot-password-manh
fly deploy
```

URL dạng: `https://your-app.fly.dev`

---

## Chạy local để test trước khi deploy

```bash
pip install -r requirements.txt

export ANTHROPIC_API_KEY=sk-ant-...
export API_SECRET_KEY=test-secret   # tuỳ chọn

uvicorn api:app --reload --port 8000
```

Mở `http://localhost:8000/docs` để test Swagger UI.

---

## Cách gọi API sau khi deploy

### Curl
```bash
# Chỉ hợp đồng
curl -X POST https://your-app.up.railway.app/analyze \
  -H "Authorization: Bearer chon-mot-password-manh" \
  -F "contracts=@hop_dong.docx" \
  -F "contracts=@phu_luc.docx" \
  --output ContractReview.xlsx

# Hợp đồng + email
curl -X POST https://your-app.up.railway.app/analyze \
  -H "Authorization: Bearer chon-mot-password-manh" \
  -F "contracts=@hop_dong.docx" \
  -F "emails=@aligned.eml" \
  --output ContractReview.xlsx
```

### Python
```python
import requests

url = "https://your-app.up.railway.app/analyze"
headers = {"Authorization": "Bearer chon-mot-password-manh"}

with open("hop_dong.docx", "rb") as f1, open("phu_luc.docx", "rb") as f2:
    response = requests.post(
        url,
        headers=headers,
        files=[
            ("contracts", ("hop_dong.docx",  f1, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")),
            ("contracts", ("phu_luc.docx",   f2, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")),
        ],
    )

with open("ContractReview.xlsx", "wb") as out:
    out.write(response.content)
```

### Swagger UI (test nhanh không cần code)
1. Vào `https://your-app.up.railway.app/docs`
2. Click **Authorize** → nhập bearer token
3. Dùng endpoint `POST /analyze` để upload file và download Excel

---

## Bảo mật

- Luôn set `API_SECRET_KEY` để API không public hoàn toàn
- `ANTHROPIC_API_KEY` **không bao giờ** commit vào git
- Thêm `.env` vào `.gitignore` nếu dùng file `.env` local

```bash
echo ".env" >> .gitignore
```
