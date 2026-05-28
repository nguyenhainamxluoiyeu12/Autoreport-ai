# AutoReport Pro Cloud Edition

Bản Cloud đầy đủ hơn gồm:

- Landing page chuyên nghiệp.
- Dashboard app.
- Dùng thử miễn phí 2 lần/ngày.
- Gói 149k/tuần và 499k/tháng.
- Subscription trả phí.
- Thanh toán demo + webhook mẫu.
- PostgreSQL online qua Render.
- SQLite fallback khi chạy local.
- Google Login OAuth.
- Upload logo công ty.
- Export Word DOCX.
- Export PDF.
- AI thật qua OpenAI API key.
- SEO: meta, canonical, sitemap.xml, robots.txt.
- Giao diện mobile responsive như app.
- File deploy Render: `render.yaml`, `Procfile`.

## Chạy local

```bash
python -m pip install -r requirements.txt
python app.py
```

Mở:

```text
http://127.0.0.1:5000
```

## Deploy

Xem file:

```text
DEPLOY_RENDER.md
```

## Lưu ý quan trọng

Các tính năng cần cấu hình thông tin thật:
- Google Login cần `GOOGLE_CLIENT_ID`, `GOOGLE_CLIENT_SECRET`.
- AI thật cần `OPENAI_API_KEY`.
- PostgreSQL online cần `DATABASE_URL`.
- Domain riêng cần cấu hình DNS ở nơi mua domain.
- Thanh toán tự động thật cần PayOS/VNPAY/Momo webhook.


## Cấu hình đăng nhập Google/Facebook

### Google OAuth
1. Vào Google Cloud Console → APIs & Services → Credentials.
2. Tạo OAuth Client ID loại Web Application.
3. Authorized redirect URI:
   - Local: `http://127.0.0.1:5000/auth/google/callback`
   - Deploy: `https://ten-mien-cua-ban.com/auth/google/callback`
4. Thêm vào biến môi trường: `GOOGLE_CLIENT_ID`, `GOOGLE_CLIENT_SECRET`.

### Facebook Login
1. Vào Meta for Developers → tạo App → thêm Facebook Login.
2. Valid OAuth Redirect URI:
   - Local: `http://127.0.0.1:5000/auth/facebook/callback`
   - Deploy: `https://ten-mien-cua-ban.com/auth/facebook/callback`
3. Thêm vào biến môi trường: `FACEBOOK_CLIENT_ID`, `FACEBOOK_CLIENT_SECRET`.

## Liên hệ hỗ trợ + Admin AI
- Trang `/support` có form liên hệ và chat nhanh.
- Người dùng gửi tin nhắn sẽ nhận phản hồi tự động từ Admin AI.
- Hệ thống tạo phiếu hỗ trợ trạng thái `CHỜ XỬ LÝ`.
- Cấu hình: `SUPPORT_EMAIL`, `SUPPORT_PHONE`, `ADMIN_EMAIL`.

## Bảo mật đã thêm
- CSRF token cho form POST.
- Giới hạn request theo phút.
- Chặn đăng nhập sai quá nhiều lần.
- Security headers: CSP, X-Frame-Options, nosniff, Referrer Policy, Permissions Policy.
- Cookie HttpOnly/SameSite, tùy chọn `COOKIE_SECURE=1` khi chạy HTTPS.
- Log sự kiện bảo mật vào bảng `SecurityLog`.

Lưu ý: Không có hệ thống nào chống hack tuyệt đối. Khi deploy thật cần bật HTTPS, dùng SECRET_KEY mạnh, không public file .env, cập nhật thư viện định kỳ và sao lưu database.

## Nâng cấp bản PRO đã thêm

### 1. Thanh toán thật
App đã có khung tích hợp `PAYMENT_PROVIDER=payos|momo|vnpay|bank`.
- PayOS: tạo checkout URL nếu có `PAYOS_CLIENT_ID`, `PAYOS_API_KEY`, `PAYOS_CHECKSUM_KEY`.
- MoMo: tạo payUrl nếu có `MOMO_PARTNER_CODE`, `MOMO_ACCESS_KEY`, `MOMO_SECRET_KEY`.
- VNPAY: tạo link thanh toán nếu có `VNPAY_TMN_CODE`, `VNPAY_HASH_SECRET`.
- Webhook chung: `/webhook/payment`.

Lưu ý: để chạy thật cần đăng ký merchant và điền API key thật của cổng thanh toán. Không nên để nút kích hoạt demo khi public chính thức.

### 2. Google/Facebook Login thật
Code đã hỗ trợ OAuth thật qua Authlib. Bạn chỉ cần điền:
- `GOOGLE_CLIENT_ID`, `GOOGLE_CLIENT_SECRET`
- `FACEBOOK_CLIENT_ID`, `FACEBOOK_CLIENT_SECRET`

Redirect URI cần đúng domain deploy:
- `/auth/google/callback`
- `/auth/facebook/callback`

### 3. Admin AI thật
Chat hỗ trợ và AI viết báo cáo dùng OpenAI nếu có `OPENAI_API_KEY`. Nếu thiếu key, hệ thống tự fallback trả lời theo mẫu để app vẫn chạy.

### 4. Bảo mật nâng cao
Đã thêm:
- `FORCE_HTTPS=1` để redirect HTTP sang HTTPS.
- HSTS khi bật HTTPS.
- CSRF, rate limit, cookie secure/httpOnly/SameSite.
- Chặn file nguy hiểm và giới hạn dung lượng upload.
- SecurityLog và route admin xem log: `/admin/security-logs?token=ADMIN_TOKEN`.
- Backup qua route: `/admin/backup?token=ADMIN_TOKEN` hoặc chạy `python backup_db.py`.

### 5. UI nâng cấp
Đã làm lại dashboard/pricing/checkout mobile đẹp hơn, thêm trạng thái payment/security/OAuth trên dashboard.
