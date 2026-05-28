# Deploy Render hoàn chỉnh

## 1. Đẩy code lên GitHub
Tạo repo mới rồi upload toàn bộ thư mục này.

## 2. Tạo PostgreSQL online
Vào Render → New → PostgreSQL → Create Database.

Copy `External Database URL` hoặc `Internal Database URL` đưa vào biến môi trường `DATABASE_URL`.

## 3. Tạo Web Service
Render → New → Web Service → chọn repo GitHub.

Build command:

```bash
pip install -r requirements.txt
```

Start command:

```bash
gunicorn app:app
```

## 4. Biến môi trường cần thêm

```text
SECRET_KEY
DOMAIN_URL
DATABASE_URL
OPENAI_API_KEY
GOOGLE_CLIENT_ID
GOOGLE_CLIENT_SECRET
BANK_NAME
BANK_ACCOUNT
BANK_OWNER
PAYMENT_SECRET
```

## 5. Google Login
Google Cloud Console → APIs & Services → Credentials → OAuth Client ID.

Authorized redirect URI:

```text
https://your-domain.com/auth/google/callback
```

## 6. Domain riêng
Render Web Service → Settings → Custom Domains → Add Domain.

Sau đó vào nơi mua domain, thêm DNS CNAME theo Render cấp.

## 7. SEO Google
Sau khi có domain:
- Vào Google Search Console.
- Thêm domain.
- Submit sitemap:

```text
https://your-domain.com/sitemap.xml
```

## 8. Thanh toán tự động thật
Kết nối PayOS/VNPAY/Momo Business webhook vào:

```text
POST /webhook/payment
```

Payload demo:

```json
{
  "payment_code": "AR12345678",
  "status": "paid",
  "signature": "HMAC_SHA256(payment_code + status, PAYMENT_SECRET)"
}
```
