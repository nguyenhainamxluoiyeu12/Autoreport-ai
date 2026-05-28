from dotenv import load_dotenv
load_dotenv()
import os
import uuid
import hmac
import hashlib
import time
import re
import json
import shutil
import zipfile
from pathlib import Path
from datetime import datetime, date, timedelta
from functools import wraps

import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

from flask import Flask, render_template, request, redirect, url_for, session, send_file, jsonify, Response, abort, flash
from flask_sqlalchemy import SQLAlchemy
from werkzeug.utils import secure_filename
from authlib.integrations.flask_client import OAuth

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

try:
    from openai import OpenAI
except Exception:
    OpenAI = None


BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")
UPLOAD_DIR = BASE_DIR / "static" / "uploads"
GENERATED_DIR = BASE_DIR / "static" / "generated"
OUTPUT_DIR = BASE_DIR / "outputs"
for d in [UPLOAD_DIR, GENERATED_DIR, OUTPUT_DIR]:
    d.mkdir(parents=True, exist_ok=True)

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-change-me")

database_url = os.environ.get("DATABASE_URL", f"sqlite:///{BASE_DIR / 'autoreport_cloud.db'}")
if database_url.startswith("postgres://"):
    database_url = database_url.replace("postgres://", "postgresql://", 1)

app.config["SQLALCHEMY_DATABASE_URI"] = database_url
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["MAX_CONTENT_LENGTH"] = 30 * 1024 * 1024
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SESSION_COOKIE_SECURE"] = os.environ.get("COOKIE_SECURE", "0") == "1"


db = SQLAlchemy(app)

APP_NAME = "AutoReport Pro Cloud"
DOMAIN_URL = os.environ.get("DOMAIN_URL", "http://127.0.0.1:5000")
FREE_DAILY_LIMIT = 2

BANK_NAME = os.environ.get("BANK_NAME", "TECHCOMBANK")
BANK_ACCOUNT = os.environ.get("BANK_ACCOUNT", "67892999999")
BANK_OWNER = os.environ.get("BANK_OWNER", "NGUYEN HAI NAM")
PAYMENT_SECRET = os.environ.get("PAYMENT_SECRET", "demo-payment-secret")
PAYMENT_PROVIDER = os.environ.get("PAYMENT_PROVIDER", "bank").lower()  # bank | payos | momo | vnpay
PAYOS_CLIENT_ID = os.environ.get("PAYOS_CLIENT_ID")
PAYOS_API_KEY = os.environ.get("PAYOS_API_KEY")
PAYOS_CHECKSUM_KEY = os.environ.get("PAYOS_CHECKSUM_KEY")
MOMO_PARTNER_CODE = os.environ.get("MOMO_PARTNER_CODE")
MOMO_ACCESS_KEY = os.environ.get("MOMO_ACCESS_KEY")
MOMO_SECRET_KEY = os.environ.get("MOMO_SECRET_KEY")
MOMO_ENDPOINT = os.environ.get("MOMO_ENDPOINT", "https://test-payment.momo.vn/v2/gateway/api/create")
VNPAY_TMN_CODE = os.environ.get("VNPAY_TMN_CODE")
VNPAY_HASH_SECRET = os.environ.get("VNPAY_HASH_SECRET")
VNPAY_PAYMENT_URL = os.environ.get("VNPAY_PAYMENT_URL", "https://sandbox.vnpayment.vn/paymentv2/vpcpay.html")
ADMIN_TOKEN = os.environ.get("ADMIN_TOKEN", "change-admin-token")
FORCE_HTTPS = os.environ.get("FORCE_HTTPS", "0") == "1"
SUPPORT_EMAIL = os.environ.get("SUPPORT_EMAIL", "support@autoreport.pro")
SUPPORT_PHONE = os.environ.get("SUPPORT_PHONE", "0987654321")
ADMIN_EMAIL = os.environ.get("ADMIN_EMAIL", "admin@autoreport.pro")
RATE_LIMIT_PER_MINUTE = int(os.environ.get("RATE_LIMIT_PER_MINUTE", "60"))


OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY")
AI_MODEL = os.environ.get("AI_MODEL", "gpt-4.1")

PLANS = {
    "weekly": {
        "name": "Gói tuần Pro", "price": 149000, "label": "149.000đ / 7 ngày", "days": 7,
        "level": "pro", "ai": "AI Pro", "exports": ["docx", "pdf"],
        "features": ["Không giới hạn lượt tạo trong 7 ngày", "AI viết báo cáo chi tiết hơn Free", "Xuất Word + PDF", "Upload Excel/CSV/TXT/DOCX/PDF mẫu", "Lưu lịch sử báo cáo"]
    },
    "monthly": {
        "name": "Gói tháng Premium", "price": 499000, "label": "499.000đ / 1 tháng", "days": 30,
        "level": "premium", "ai": "AI Premium mạnh nhất", "exports": ["docx", "pdf"],
        "features": ["Toàn bộ tính năng gói tuần", "AI phân tích sâu, văn phong chuyên nghiệp hơn", "Upload mẫu hầu hết định dạng phổ biến", "Nhận diện bố cục file mẫu để viết theo", "Logo công ty + ưu tiên nâng cấp", "Phù hợp dùng lâu dài / bán dịch vụ"]
    },
}


REPORT_LIBRARY = [
    {
        "title": "Báo cáo thực tập tốt nghiệp",
        "category": "Thực tập",
        "keywords": ["thực tập", "nhật ký", "công ty", "kết quả", "kiến nghị"],
        "description": "Mẫu báo cáo có mở đầu, giới thiệu đơn vị, nội dung công việc, kết quả đạt được và kiến nghị.",
        "outline": "Mở đầu → Giới thiệu đơn vị → Quá trình thực tập → Kết quả → Kết luận"
    },
    {
        "title": "Báo cáo đồ án kỹ thuật",
        "category": "Kỹ thuật",
        "keywords": ["đồ án", "kỹ thuật", "thiết kế", "mô hình", "thực nghiệm"],
        "description": "Phù hợp cho đồ án điện, tự động hóa, IoT, năng lượng mặt trời, có phần cơ sở lý thuyết và triển khai mô hình.",
        "outline": "Tổng quan → Cơ sở lý thuyết → Thiết kế hệ thống → Thực nghiệm → Đánh giá"
    },
    {
        "title": "Báo cáo công việc tuần",
        "category": "Công việc",
        "keywords": ["tuần", "tiến độ", "kpi", "công việc", "kế hoạch"],
        "description": "Mẫu báo cáo nhanh để tổng hợp việc đã làm, việc chưa xong, khó khăn và kế hoạch tuần tiếp theo.",
        "outline": "Tóm tắt → Việc đã hoàn thành → Vấn đề tồn tại → Kế hoạch tuần tới"
    },
    {
        "title": "Báo cáo doanh thu bán hàng",
        "category": "Kinh doanh",
        "keywords": ["doanh thu", "bán hàng", "khách hàng", "lợi nhuận", "biểu đồ"],
        "description": "Mẫu phân tích doanh thu theo thời gian, sản phẩm, khách hàng, kèm nhận xét tăng giảm và đề xuất.",
        "outline": "Tổng quan doanh thu → Bảng số liệu → Biểu đồ → Nhận xét → Giải pháp"
    },
    {
        "title": "Báo cáo sự cố và khắc phục",
        "category": "Vận hành",
        "keywords": ["sự cố", "nguyên nhân", "khắc phục", "rủi ro", "bảo trì"],
        "description": "Mẫu báo cáo incident: mô tả hiện tượng, nguyên nhân gốc, hành động xử lý và phòng ngừa tái diễn.",
        "outline": "Thông tin sự cố → Tác động → Nguyên nhân → Xử lý → Phòng ngừa"
    },
    {
        "title": "Báo cáo nghiên cứu thị trường",
        "category": "Marketing",
        "keywords": ["thị trường", "đối thủ", "khách hàng", "xu hướng", "marketing"],
        "description": "Mẫu tổng hợp insight thị trường, chân dung khách hàng, đối thủ cạnh tranh và cơ hội phát triển.",
        "outline": "Mục tiêu → Dữ liệu khảo sát → Phân tích khách hàng → Đối thủ → Đề xuất"
    },
    {
        "title": "Báo cáo tài chính cơ bản",
        "category": "Tài chính",
        "keywords": ["tài chính", "chi phí", "lợi nhuận", "dòng tiền", "ngân sách"],
        "description": "Mẫu báo cáo thu chi, lợi nhuận, dòng tiền và đánh giá hiệu quả sử dụng ngân sách.",
        "outline": "Tổng quan → Thu nhập → Chi phí → Lợi nhuận → Dự báo"
    },
    {
        "title": "Báo cáo dự án phần mềm",
        "category": "Công nghệ",
        "keywords": ["phần mềm", "web", "app", "database", "deploy"],
        "description": "Mẫu báo cáo tiến độ phát triển web/app, trạng thái module, lỗi tồn tại và kế hoạch release.",
        "outline": "Mục tiêu dự án → Tính năng → Kiến trúc → Tiến độ → Kiểm thử → Triển khai"
    }
]

oauth = OAuth(app)
google = oauth.register(
    name="google",
    client_id=os.environ.get("GOOGLE_CLIENT_ID"),
    client_secret=os.environ.get("GOOGLE_CLIENT_SECRET"),
    server_metadata_url="https://accounts.google.com/.well-known/openid-configuration",
    client_kwargs={"scope": "openid email profile"},
)

facebook = oauth.register(
    name="facebook",
    client_id=os.environ.get("FACEBOOK_CLIENT_ID"),
    client_secret=os.environ.get("FACEBOOK_CLIENT_SECRET"),
    access_token_url="https://graph.facebook.com/v19.0/oauth/access_token",
    authorize_url="https://www.facebook.com/v19.0/dialog/oauth",
    api_base_url="https://graph.facebook.com/v19.0/",
    client_kwargs={"scope": "email,public_profile"},
)


class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_key = db.Column(db.String(80), unique=True, nullable=False)
    email = db.Column(db.String(180), unique=True)
    name = db.Column(db.String(180))
    avatar = db.Column(db.Text)
    password_hash = db.Column(db.String(255))
    company_logo = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


class UsageLog(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_key = db.Column(db.String(80), nullable=False)
    usage_date = db.Column(db.String(20), nullable=False)
    count = db.Column(db.Integer, default=0)


class Report(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    report_id = db.Column(db.String(40), unique=True, nullable=False)
    user_key = db.Column(db.String(80), nullable=False)
    title = db.Column(db.String(220))
    report_type = db.Column(db.String(120))
    author = db.Column(db.String(120))
    department = db.Column(db.String(160))
    docx_filename = db.Column(db.String(260))
    pdf_filename = db.Column(db.String(260))
    has_excel = db.Column(db.Boolean, default=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


class Payment(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_key = db.Column(db.String(80), nullable=False)
    plan = db.Column(db.String(50), nullable=False)
    amount = db.Column(db.Integer, nullable=False)
    payment_code = db.Column(db.String(80), unique=True, nullable=False)
    status = db.Column(db.String(30), default="pending")
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    paid_at = db.Column(db.DateTime)


class Subscription(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_key = db.Column(db.String(80), nullable=False)
    plan = db.Column(db.String(50), nullable=False)
    status = db.Column(db.String(30), default="active")
    start_at = db.Column(db.DateTime, default=datetime.utcnow)
    end_at = db.Column(db.DateTime, nullable=False)
    payment_code = db.Column(db.String(80))


class SupportTicket(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    ticket_code = db.Column(db.String(40), unique=True, nullable=False)
    user_key = db.Column(db.String(80), nullable=False)
    name = db.Column(db.String(180))
    email = db.Column(db.String(180))
    topic = db.Column(db.String(180))
    status = db.Column(db.String(40), default="waiting")
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


class SupportMessage(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    ticket_code = db.Column(db.String(40), nullable=False)
    sender = db.Column(db.String(40), default="user")
    message = db.Column(db.Text, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


class SecurityLog(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    ip = db.Column(db.String(80))
    event = db.Column(db.String(120))
    detail = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)




def client_ip():
    return (request.headers.get("X-Forwarded-For", request.remote_addr or "unknown").split(",")[0].strip())[:80]


def log_security(event, detail=""):
    try:
        db.session.add(SecurityLog(ip=client_ip(), event=event, detail=str(detail)[:500]))
        db.session.commit()
    except Exception:
        db.session.rollback()


def clean_text(value, max_len=1000):
    value = (value or "").strip()
    value = re.sub(r"<[^>]*>", "", value)
    return value[:max_len]


def generate_csrf():
    token = session.get("csrf_token")
    if not token:
        token = hashlib.sha256((str(uuid.uuid4()) + app.secret_key).encode()).hexdigest()
        session["csrf_token"] = token
    return token


def verify_csrf():
    token = request.form.get("csrf_token") or request.headers.get("X-CSRF-Token")
    if not token or token != session.get("csrf_token"):
        log_security("csrf_block", request.path)
        abort(403)


def rate_limit():
    now = int(time.time())
    bucket = session.get("rate_bucket", {"start": now, "count": 0})
    if now - bucket.get("start", now) >= 60:
        bucket = {"start": now, "count": 0}
    bucket["count"] = bucket.get("count", 0) + 1
    session["rate_bucket"] = bucket
    if bucket["count"] > RATE_LIMIT_PER_MINUTE:
        log_security("rate_limit", request.path)
        abort(429)


def require_login(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if "user_key" not in session:
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return wrapper


@app.before_request
def security_guard():
    db.create_all()
    if FORCE_HTTPS and request.headers.get("X-Forwarded-Proto", "https") != "https":
        return redirect(request.url.replace("http://", "https://", 1), code=301)
    rate_limit()
    if request.method == "POST" and request.endpoint not in {"payment_webhook", "support_chat_api"}:
        verify_csrf()


@app.after_request
def add_security_headers(resp):
    resp.headers["X-Frame-Options"] = "DENY"
    resp.headers["X-Content-Type-Options"] = "nosniff"
    resp.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"
    resp.headers["Permissions-Policy"] = "geolocation=(), microphone=(), camera=()"
    if FORCE_HTTPS:
        resp.headers["Strict-Transport-Security"] = "max-age=31536000; includeSubDomains"
    resp.headers["Content-Security-Policy"] = "default-src 'self' https://fonts.googleapis.com https://fonts.gstatic.com; img-src 'self' data: https:; style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; script-src 'self' 'unsafe-inline'; connect-src 'self';"
    return resp


@app.context_processor
def inject_globals():
    return {"csrf_token": generate_csrf(), "support_email": SUPPORT_EMAIL, "support_phone": SUPPORT_PHONE, "admin_email": ADMIN_EMAIL, "payment_provider": PAYMENT_PROVIDER}

@app.before_request
def ensure_db():
    pass




def hash_password(password):
    salt = os.environ.get("PASSWORD_SALT", app.secret_key)
    return hashlib.pbkdf2_hmac("sha256", (password or "").encode(), salt.encode(), 120000).hex()

def today_str():
    return date.today().isoformat()


def get_or_create_user():
    if "user_key" not in session:
        session["user_key"] = "guest_" + str(uuid.uuid4())
    user = User.query.filter_by(user_key=session["user_key"]).first()
    if not user:
        user = User(user_key=session["user_key"], name="Khách dùng thử")
        db.session.add(user)
        db.session.commit()
    return user


def active_subscription(user_key):
    return Subscription.query.filter(
        Subscription.user_key == user_key,
        Subscription.status == "active",
        Subscription.end_at >= datetime.utcnow()
    ).order_by(Subscription.id.desc()).first()


def usage_count(user_key):
    row = UsageLog.query.filter_by(user_key=user_key, usage_date=today_str()).first()
    return row.count if row else 0


def increase_usage(user_key):
    row = UsageLog.query.filter_by(user_key=user_key, usage_date=today_str()).first()
    if not row:
        row = UsageLog(user_key=user_key, usage_date=today_str(), count=0)
        db.session.add(row)
    row.count += 1
    db.session.commit()


def can_generate(user_key):
    sub = active_subscription(user_key)
    if sub:
        return True, "pro", 999999, sub
    used = usage_count(user_key)
    if used < FREE_DAILY_LIMIT:
        return True, "free", FREE_DAILY_LIMIT - used, None
    return False, "locked", 0, None


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def add_page_border(section):
    sectPr = section._sectPr
    pgBorders = OxmlElement("w:pgBorders")
    pgBorders.set(qn("w:offsetFrom"), "page")
    for name in ["top", "left", "bottom", "right"]:
        border = OxmlElement(f"w:{name}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), "8")
        border.set(qn("w:space"), "24")
        border.set(qn("w:color"), "4F46E5")
        pgBorders.append(border)
    sectPr.append(pgBorders)


def add_heading(doc, text, level=1):
    p = doc.add_heading(text, level=level)
    for run in p.runs:
        run.font.name = "Times New Roman"
        run.font.color.rgb = RGBColor(67, 56, 202)
    return p


def add_body(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.25
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run(text)
    r.font.name = "Times New Roman"
    r.font.size = Pt(12)
    return p


def read_excel_and_chart(excel_path, report_id):
    try:
        df = pd.read_excel(excel_path)
        if df.empty:
            return df, None
        numeric_cols = df.select_dtypes(include="number").columns.tolist()
        if not numeric_cols:
            return df, None
        y_col = numeric_cols[0]
        x_col = df.columns[0]
        plt.figure(figsize=(8.4, 4.7))
        plt.plot(df[x_col].astype(str), df[y_col], marker="o")
        plt.title(f"Biểu đồ phân tích: {y_col}")
        plt.xlabel(str(x_col))
        plt.ylabel(str(y_col))
        plt.xticks(rotation=30, ha="right")
        plt.tight_layout()
        chart_path = GENERATED_DIR / f"chart_{report_id}.png"
        plt.savefig(chart_path, dpi=170)
        plt.close()
        return df, chart_path
    except Exception:
        return None, None



ALLOWED_TEMPLATE_EXTS = (".xlsx", ".xls", ".csv", ".txt", ".md", ".json", ".html", ".htm", ".py", ".js", ".css", ".docx", ".pdf", ".pptx")
ALLOWED_DATA_EXTS = (".xlsx", ".xls", ".csv")
ALLOWED_LOGO_EXTS = (".png", ".jpg", ".jpeg", ".webp")
DANGEROUS_EXTS = (".exe", ".bat", ".cmd", ".sh", ".ps1", ".scr", ".php", ".jsp")
MAX_TEMPLATE_SIZE = 15 * 1024 * 1024
MAX_DATA_SIZE = 10 * 1024 * 1024
MAX_LOGO_SIZE = 3 * 1024 * 1024


def is_safe_upload(file_storage, allowed_exts, max_size):
    filename = secure_filename(file_storage.filename or "")
    ext = Path(filename).suffix.lower()
    if not filename or ext in DANGEROUS_EXTS or ext not in allowed_exts:
        return False, "Định dạng file không được phép."
    pos = file_storage.stream.tell()
    file_storage.stream.seek(0, os.SEEK_END)
    size = file_storage.stream.tell()
    file_storage.stream.seek(pos)
    if size > max_size:
        return False, f"File vượt quá giới hạn {max_size // (1024*1024)}MB."
    return True, filename


def get_plan_level(user_key):
    sub = active_subscription(user_key)
    if not sub:
        return "free"
    return PLANS.get(sub.plan, {}).get("level", "pro")


def payment_transfer_content(user):
    raw_name = (user.name or user.email or "NGUOI DUNG").upper()
    safe_name = "".join(ch for ch in raw_name if ch.isalnum() or ch in " _.-").strip() or "NGUOI DUNG"
    return f"{safe_name} 123+CK"


def read_template_file(path):
    """Đọc nhanh nội dung file mẫu để AI bắt chước cấu trúc. Không sửa file gốc."""
    try:
        ext = path.suffix.lower()
        if ext in (".xlsx", ".xls"):
            df = pd.read_excel(path, nrows=20)
            return "Mẫu Excel:\n" + df.head(12).to_csv(index=False)
        if ext == ".csv":
            df = pd.read_csv(path, nrows=20)
            return "Mẫu CSV:\n" + df.head(12).to_csv(index=False)
        if ext == ".docx":
            d = Document(str(path))
            text = "\n".join(p.text for p in d.paragraphs if p.text.strip())
            return "Mẫu Word DOCX:\n" + text[:7000]
        if ext == ".pdf":
            try:
                from pypdf import PdfReader
                reader = PdfReader(str(path))
                text = "\n".join((page.extract_text() or "") for page in reader.pages[:5])
                return "Mẫu PDF:\n" + text[:7000]
            except Exception:
                return "File mẫu PDF đã được upload. Hãy làm báo cáo theo phong cách tài liệu PDF chuyên nghiệp."
        if ext == ".pptx":
            try:
                from pptx import Presentation
                prs = Presentation(str(path))
                chunks = []
                for slide in prs.slides[:12]:
                    for shape in slide.shapes:
                        if hasattr(shape, "text") and shape.text.strip():
                            chunks.append(shape.text.strip())
                return "Mẫu PowerPoint:\n" + "\n".join(chunks)[:7000]
            except Exception:
                return "File mẫu PPTX đã được upload. Hãy nhận diện theo kiểu slide/thuyết trình chuyên nghiệp."
        if ext in (".txt", ".md", ".json", ".html", ".htm", ".py", ".js", ".css"):
            return "File mẫu dạng văn bản/code:\n" + path.read_text(encoding="utf-8", errors="ignore")[:7000]
    except Exception as e:
        return f"Đã nhận file mẫu nhưng chưa đọc được chi tiết: {e}"
    return "Đã nhận file mẫu. Hãy làm theo cấu trúc, văn phong và cách trình bày tương tự."

def ai_insight(data, df):
    fallback = []
    if df is not None and not df.empty:
        rows, cols = df.shape
        fallback.append(f"Dữ liệu đầu vào có {rows} dòng và {cols} cột.")
        numeric_cols = df.select_dtypes(include="number").columns.tolist()
        for col in numeric_cols[:4]:
            fallback.append(f"Chỉ tiêu '{col}' có tổng {df[col].sum():,.2f}, trung bình {df[col].mean():,.2f}, cao nhất {df[col].max():,.2f}, thấp nhất {df[col].min():,.2f}.")
    else:
        fallback.append("Báo cáo được tạo dựa trên thông tin nhập thủ công của người dùng.")
    fallback.append("Nội dung cần được cập nhật định kỳ để nâng cao độ chính xác và giá trị sử dụng.")

    if not OPENAI_API_KEY or OpenAI is None:
        return fallback

    try:
        client = OpenAI(api_key=OPENAI_API_KEY)
        df_summary = ""
        if df is not None and not df.empty:
            df_summary = df.head(10).to_csv(index=False)
        prompt = f"""
Bạn là AI báo cáo cấp Premium, viết tiếng Việt chuyên nghiệp như trợ lý doanh nghiệp.
Hãy phân tích sâu, rõ ý, có nhận xét, rủi ro và đề xuất. Viết 6-8 gạch đầu dòng chất lượng cao.

Loại báo cáo: {data.get('report_type')}
Tiêu đề: {data.get('title')}
Mục tiêu: {data.get('goal')}
Nội dung: {data.get('tasks')}
Dữ liệu Excel mẫu:
{df_summary}

File mẫu người dùng upload để AI làm theo:
{data.get("template_text", "Không có file mẫu.")}
"""
        res = client.chat.completions.create(
            model=AI_MODEL,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.4,
        )
        text = res.choices[0].message.content.strip()
        lines = [x.strip("-• ").strip() for x in text.splitlines() if x.strip()]
        return lines[:8] or fallback
    except Exception:
        return fallback


def ai_full_sections(data, df=None):
    """Tạo nội dung báo cáo đầy đủ bằng OpenAI. Nếu lỗi/thiếu key thì trả None để dùng bản dự phòng."""
    if not OPENAI_API_KEY or OpenAI is None:
        return None
    try:
        client = OpenAI(api_key=OPENAI_API_KEY)
        df_summary = ""
        if df is not None and not df.empty:
            df_summary = df.head(15).to_csv(index=False)

        prompt = f"""
Bạn là trợ lý AI chuyên viết báo cáo tiếng Việt. Hãy tạo một báo cáo HOÀN TOÀN MỚI, không dùng nội dung mặc định.

Thông tin người dùng nhập:
- Tiêu đề: {data.get('title')}
- Loại báo cáo: {data.get('report_type')}
- Người lập: {data.get('author')}
- Đơn vị/lớp/phòng ban: {data.get('department')}
- Mục tiêu: {data.get('goal')}
- Nội dung/công việc/yêu cầu: {data.get('tasks')}
- Kết luận người dùng nhập: {data.get('conclusion')}
- Dữ liệu Excel tóm tắt:
{df_summary}
- File mẫu người dùng upload:
{data.get('template_text', 'Không có file mẫu.')}

Yêu cầu xuất JSON THUẦN, không markdown, không giải thích ngoài JSON.
Dạng JSON:
{{
  "sections": [
    {{"heading": "1. Mở đầu", "paragraphs": ["...", "..."]}},
    {{"heading": "2. Nội dung chính", "paragraphs": ["...", "..."]}},
    {{"heading": "3. Phân tích và đánh giá", "paragraphs": ["...", "..."]}},
    {{"heading": "4. Kết luận và đề xuất", "paragraphs": ["...", "..."]}}
  ]
}}

Viết chuyên nghiệp, dài vừa đủ, bám sát đúng tiêu đề và nội dung người dùng nhập.
"""
        res = client.chat.completions.create(
            model=AI_MODEL,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.55,
        )
        raw = res.choices[0].message.content.strip()
        raw = raw.replace("```json", "").replace("```", "").strip()
        parsed = json.loads(raw)
        sections = parsed.get("sections", [])
        clean = []
        for sec in sections:
            heading = str(sec.get("heading", "")).strip()
            paragraphs = [str(x).strip() for x in sec.get("paragraphs", []) if str(x).strip()]
            if heading and paragraphs:
                clean.append({"heading": heading, "paragraphs": paragraphs})
        return clean or None
    except Exception as e:
        print("[AI_FULL_CONTENT_ERROR]", e)
        return None


def make_docx_and_pdf(user, data, excel_path=None):
    report_id = str(uuid.uuid4())[:8]
    report_type = data.get("report_type", "Báo cáo tổng hợp")
    df, chart_path = (None, None)
    if excel_path:
        df, chart_path = read_excel_and_chart(excel_path, report_id)

    safe_title = secure_filename((data.get("title") or "Bao_cao").replace(" ", "_"))[:36] or "Bao_cao"
    docx_filename = f"{safe_title}_{report_id}.docx"
    pdf_filename = f"{safe_title}_{report_id}.pdf"
    docx_path = OUTPUT_DIR / docx_filename
    pdf_path = OUTPUT_DIR / pdf_filename

    insights = ai_insight(data, df)

    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.68)
    section.bottom_margin = Inches(0.68)
    section.left_margin = Inches(0.85)
    section.right_margin = Inches(0.85)
    add_page_border(section)

    if user.company_logo:
        logo_path = BASE_DIR / user.company_logo.lstrip("/")
        if logo_path.exists():
            doc.add_picture(str(logo_path), width=Inches(1.15))
            doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = title.add_run((data.get("title") or "BÁO CÁO TỔNG HỢP").upper())
    r.bold = True
    r.font.size = Pt(21)
    r.font.name = "Times New Roman"
    r.font.color.rgb = RGBColor(15, 23, 42)

    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    m = meta.add_run(f"Tạo bởi AutoReport Pro Cloud • {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    m.italic = True
    m.font.size = Pt(10.5)
    m.font.name = "Times New Roman"

    table = doc.add_table(rows=5, cols=2)
    table.style = "Table Grid"
    info = [
        ("Người lập báo cáo", data.get("author") or user.name or "Chưa nhập"),
        ("Đơn vị / Lớp / Phòng ban", data.get("department") or "Chưa nhập"),
        ("Loại báo cáo", report_type),
        ("Kỳ báo cáo", f"{data.get('from_date') or '...'} đến {data.get('to_date') or '...'}"),
        ("Mức độ ưu tiên", data.get("priority") or "Bình thường"),
    ]
    for i, (k, v) in enumerate(info):
        table.cell(i, 0).text = k
        table.cell(i, 1).text = v
        set_cell_shading(table.cell(i, 0), "E0E7FF")

    ai_sections = ai_full_sections(data, df)
    if ai_sections:
        for sec in ai_sections:
            add_heading(doc, sec["heading"], 1)
            for para in sec["paragraphs"]:
                add_body(doc, para)
    else:
        add_heading(doc, "1. Mục tiêu báo cáo", 1)
        add_body(doc, data.get("goal") or "Báo cáo nhằm tổng hợp dữ liệu, đánh giá kết quả và đề xuất hướng xử lý tiếp theo.")

        add_heading(doc, "2. Nội dung thực hiện", 1)
        tasks = [x.strip("-• \n") for x in data.get("tasks", "").splitlines() if x.strip()]
        if tasks:
            for task in tasks:
                p = doc.add_paragraph(style="List Bullet")
                rr = p.add_run(task)
                rr.font.name = "Times New Roman"
                rr.font.size = Pt(12)
        else:
            add_body(doc, "Chưa có nội dung công việc cụ thể.")

        add_heading(doc, "3. Phân tích bằng AI", 1)
        for line in insights:
            add_body(doc, line)

    if df is not None and not df.empty:
        add_heading(doc, "4. Bảng dữ liệu", 1)
        preview = df.head(12)
        t = doc.add_table(rows=1, cols=len(preview.columns))
        t.style = "Table Grid"
        for i, col in enumerate(preview.columns):
            t.cell(0, i).text = str(col)
            set_cell_shading(t.cell(0, i), "4F46E5")
        for _, row in preview.iterrows():
            cells = t.add_row().cells
            for i, val in enumerate(row):
                cells[i].text = "" if pd.isna(val) else str(val)

    if chart_path and chart_path.exists():
        add_heading(doc, "5. Biểu đồ minh họa", 1)
        doc.add_picture(str(chart_path), width=Inches(6.45))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    if data.get("template_text"):
        add_heading(doc, "6. Nhận diện file mẫu", 1)
        add_body(doc, "AI đã tham chiếu file mẫu được upload để bám theo bố cục, văn phong và cách trình bày của tài liệu mẫu.")

    if not ai_sections:
        add_heading(doc, "7. Kết luận và đề xuất", 1)
        add_body(doc, data.get("conclusion") or "Báo cáo đã tổng hợp các nội dung trọng tâm. Cần tiếp tục cập nhật dữ liệu, chuẩn hóa quy trình và ứng dụng tự động hóa để nâng cao hiệu quả xử lý.")

    doc.save(docx_path)

    # PDF export
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="VNTitle", fontName="Helvetica-Bold", fontSize=18, leading=24, alignment=1, textColor=colors.HexColor("#1e293b")))
    styles.add(ParagraphStyle(name="VNBody", fontName="Helvetica", fontSize=11, leading=16))
    pdf_doc = SimpleDocTemplate(str(pdf_path), pagesize=A4, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    story = []
    story.append(Paragraph(data.get("title", "Báo cáo tổng hợp"), styles["VNTitle"]))
    story.append(Spacer(1, 14))
    story.append(Paragraph(f"Người lập: {data.get('author') or user.name or '...'}", styles["VNBody"]))
    story.append(Paragraph(f"Loại báo cáo: {report_type}", styles["VNBody"]))
    story.append(Spacer(1, 12))
    if ai_sections:
        for sec in ai_sections:
            story.append(Paragraph(f"<b>{sec['heading']}</b>", styles["VNBody"]))
            for para in sec["paragraphs"]:
                story.append(Paragraph(para, styles["VNBody"]))
            story.append(Spacer(1, 8))
    else:
        story.append(Paragraph("<b>1. Mục tiêu báo cáo</b>", styles["VNBody"]))
        story.append(Paragraph(data.get("goal") or "Báo cáo nhằm tổng hợp dữ liệu, đánh giá kết quả và đề xuất hướng xử lý tiếp theo.", styles["VNBody"]))
        story.append(Spacer(1, 8))
        story.append(Paragraph("<b>2. Phân tích bằng AI</b>", styles["VNBody"]))
        for line in insights:
            story.append(Paragraph("• " + line, styles["VNBody"]))
    if chart_path and chart_path.exists():
        story.append(Spacer(1, 12))
        story.append(RLImage(str(chart_path), width=420, height=235))
    if not ai_sections:
        story.append(Spacer(1, 12))
        story.append(Paragraph("<b>3. Kết luận</b>", styles["VNBody"]))
        story.append(Paragraph(data.get("conclusion") or "Báo cáo đã tổng hợp các nội dung trọng tâm.", styles["VNBody"]))
    pdf_doc.build(story)

    report = Report(
        report_id=report_id,
        user_key=user.user_key,
        title=data.get("title", "Báo cáo"),
        report_type=report_type,
        author=data.get("author") or user.name,
        department=data.get("department"),
        docx_filename=docx_filename,
        pdf_filename=pdf_filename,
        has_excel=excel_path is not None,
    )
    db.session.add(report)
    db.session.commit()
    return docx_path, pdf_path


@app.route("/")
def landing():
    return render_template("landing.html", domain=DOMAIN_URL)


@app.route("/app")
def dashboard():
    user = get_or_create_user()
    allowed, mode, remaining, sub = can_generate(user.user_key)
    recent = Report.query.filter_by(user_key=user.user_key).order_by(Report.id.desc()).limit(5).all()
    return render_template("dashboard.html", user=user, allowed=allowed, mode=mode, remaining=remaining, used=usage_count(user.user_key), free_limit=FREE_DAILY_LIMIT, subscription=sub, recent=recent, plans=PLANS, report_library=REPORT_LIBRARY)


@app.route("/login/google")
def google_login():
    if not os.environ.get("GOOGLE_CLIENT_ID"):
        return "Chưa cấu hình GOOGLE_CLIENT_ID và GOOGLE_CLIENT_SECRET.", 400
    redirect_uri = url_for("google_callback", _external=True)
    return google.authorize_redirect(redirect_uri)


@app.route("/auth/google/callback")
def google_callback():
    token = google.authorize_access_token()
    info = token.get("userinfo") or google.parse_id_token(token)
    email = info.get("email")
    user = User.query.filter_by(email=email).first()
    if not user:
        user = User(
            user_key="google_" + str(uuid.uuid4()),
            email=email,
            name=info.get("name"),
            avatar=info.get("picture"),
        )
        db.session.add(user)
        db.session.commit()
    session["user_key"] = user.user_key
    return redirect(url_for("dashboard"))




@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        email = clean_text(request.form.get("email"), 180).lower()
        password = request.form.get("password")
        attempts = session.get("login_attempts", 0)
        if attempts >= 8:
            log_security("login_blocked", email)
            return render_template("auth.html", mode="login", error="Tài khoản/IP đăng nhập sai quá nhiều lần. Vui lòng thử lại sau."), 429
        user = User.query.filter_by(email=email).first()
        if user and user.password_hash == hash_password(password):
            session["user_key"] = user.user_key
            session["login_attempts"] = 0
            return redirect(url_for("dashboard"))
        session["login_attempts"] = attempts + 1
        log_security("login_failed", email)
        return render_template("auth.html", mode="login", error="Sai tài khoản hoặc mật khẩu")
    return render_template("auth.html", mode="login")


@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        name = clean_text(request.form.get("name"), 180)
        email = clean_text(request.form.get("email"), 180).lower()
        password = request.form.get("password")
        if len(password or "") < 6:
            return render_template("auth.html", mode="register", error="Mật khẩu tối thiểu 6 ký tự")
        if User.query.filter_by(email=email).first():
            return render_template("auth.html", mode="register", error="Email đã tồn tại")
        user = User(user_key="user_" + str(uuid.uuid4()), name=name, email=email, password_hash=hash_password(password))
        db.session.add(user)
        db.session.commit()
        session["user_key"] = user.user_key
        return redirect(url_for("dashboard"))
    return render_template("auth.html", mode="register")


@app.route("/login/facebook")
def login_facebook():
    if not os.environ.get("FACEBOOK_CLIENT_ID"):
        return "Chưa cấu hình FACEBOOK_CLIENT_ID và FACEBOOK_CLIENT_SECRET.", 400
    redirect_uri = url_for("facebook_callback", _external=True)
    return facebook.authorize_redirect(redirect_uri)


@app.route("/auth/facebook/callback")
def facebook_callback():
    token = facebook.authorize_access_token()
    resp = facebook.get("me?fields=id,name,email,picture", token=token)
    info = resp.json()
    email = info.get("email") or f"facebook_{info.get('id')}@facebook.local"
    user = User.query.filter_by(email=email).first()
    if not user:
        pic = (((info.get("picture") or {}).get("data") or {}).get("url"))
        user = User(user_key="facebook_" + str(uuid.uuid4()), email=email, name=info.get("name"), avatar=pic)
        db.session.add(user)
        db.session.commit()
    session["user_key"] = user.user_key
    return redirect(url_for("dashboard"))


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("landing"))


@app.route("/upload-logo", methods=["POST"])
def upload_logo():
    user = get_or_create_user()
    f = request.files.get("logo")
    if not f or not f.filename:
        return redirect(url_for("dashboard"))
    ok, result = is_safe_upload(f, ALLOWED_LOGO_EXTS, MAX_LOGO_SIZE)
    if not ok:
        log_security("blocked_logo_upload", result)
        return result, 400
    filename = result
    save_name = f"logo_{user.id}_{uuid.uuid4().hex[:8]}_{filename}"
    path = UPLOAD_DIR / save_name
    f.save(path)
    user.company_logo = f"/static/uploads/{save_name}"
    db.session.commit()
    return redirect(url_for("dashboard"))


@app.route("/generate", methods=["POST"])
def generate():
    user = get_or_create_user()
    allowed, mode, remaining, sub = can_generate(user.user_key)
    if not allowed:
        return redirect(url_for("pricing"))

    data = request.form.to_dict()
    plan_level = get_plan_level(user.user_key)

    file = request.files.get("excel_file")
    excel_path = None
    if file and file.filename:
        ok, result = is_safe_upload(file, ALLOWED_DATA_EXTS, MAX_DATA_SIZE)
        if not ok:
            log_security("blocked_data_upload", result)
            return result, 400
        filename = result
        excel_path = UPLOAD_DIR / f"data_{uuid.uuid4()}_{filename}"
        file.save(excel_path)

    template_file = request.files.get("template_file")
    if template_file and template_file.filename:
        if plan_level == "free":
            return "Tính năng AI đọc file mẫu chỉ có trong gói 7 ngày và 1 tháng.", 403
        ok, result = is_safe_upload(template_file, ALLOWED_TEMPLATE_EXTS, MAX_TEMPLATE_SIZE)
        if not ok:
            log_security("blocked_template_upload", result)
            return result, 400
        tname = result
        tpath = UPLOAD_DIR / f"template_{uuid.uuid4()}_{tname}"
        template_file.save(tpath)
        data["template_text"] = read_template_file(tpath)
        data["template_filename"] = tname

    if plan_level == "free" and data.get("export_format") == "pdf":
        return "Gói Free chỉ xuất Word DOCX. Nâng cấp gói 7 ngày hoặc 1 tháng để xuất PDF.", 403

    docx_path, pdf_path = make_docx_and_pdf(user, data, excel_path)

    if mode == "free":
        increase_usage(user.user_key)

    fmt = data.get("export_format", "docx")
    return send_file(pdf_path if fmt == "pdf" else docx_path, as_attachment=True)


@app.route("/api/report-library")
def api_report_library():
    q = clean_text(request.args.get("q", ""), 120).lower()
    items = REPORT_LIBRARY
    if q:
        items = [r for r in REPORT_LIBRARY if q in r["title"].lower() or q in r["category"].lower() or any(q in k.lower() for k in r["keywords"])]
    return jsonify({"keywords": sorted({k for r in REPORT_LIBRARY for k in r["keywords"]}), "reports": items})


@app.route("/history")
def history():
    user = get_or_create_user()
    reports = Report.query.filter_by(user_key=user.user_key).order_by(Report.id.desc()).limit(100).all()
    return render_template("history.html", reports=reports)


@app.route("/download/<fmt>/<filename>")
def download(fmt, filename):
    filename = secure_filename(filename)
    path = OUTPUT_DIR / filename
    if not path.exists():
        return "File không tồn tại", 404
    return send_file(path, as_attachment=True)


@app.route("/pricing")
def pricing():
    user = get_or_create_user()
    return render_template("pricing.html", plans=PLANS, used=usage_count(user.user_key), free_limit=FREE_DAILY_LIMIT, subscription=active_subscription(user.user_key))




def create_gateway_payment(provider, payment, plan_info, user):
    """Tạo link thanh toán thật nếu đã cấu hình env. Thiếu env thì fallback chuyển khoản ngân hàng."""
    return_url = f"{DOMAIN_URL.rstrip('/')}/payment/success/{payment.payment_code}"
    webhook_url = f"{DOMAIN_URL.rstrip('/')}/webhook/payment"
    description = payment_transfer_content(user)
    amount = int(plan_info["price"])
    provider = (provider or "bank").lower()
    try:
        import requests
        if provider == "payos" and PAYOS_CLIENT_ID and PAYOS_API_KEY and PAYOS_CHECKSUM_KEY:
            order_code = int(''.join(filter(str.isdigit, payment.payment_code))[:12] or str(int(time.time())))
            raw = f"amount={amount}&cancelUrl={return_url}&description={description[:25]}&orderCode={order_code}&returnUrl={return_url}"
            signature = hmac.new(PAYOS_CHECKSUM_KEY.encode(), raw.encode(), hashlib.sha256).hexdigest()
            payload = {"orderCode": order_code, "amount": amount, "description": description[:25], "returnUrl": return_url, "cancelUrl": return_url, "signature": signature, "items": [{"name": plan_info['name'], "quantity": 1, "price": amount}]}
            headers = {"x-client-id": PAYOS_CLIENT_ID, "x-api-key": PAYOS_API_KEY, "Content-Type": "application/json"}
            res = requests.post("https://api-merchant.payos.vn/v2/payment-requests", json=payload, headers=headers, timeout=12)
            data = res.json()
            return (data.get("data") or {}).get("checkoutUrl"), data
        if provider == "momo" and MOMO_PARTNER_CODE and MOMO_ACCESS_KEY and MOMO_SECRET_KEY:
            order_id = payment.payment_code
            request_id = payment.payment_code + str(int(time.time()))
            extra_data = ""
            order_info = f"Thanh toan {APP_NAME} {payment.payment_code}"
            raw = f"accessKey={MOMO_ACCESS_KEY}&amount={amount}&extraData={extra_data}&ipnUrl={webhook_url}&orderId={order_id}&orderInfo={order_info}&partnerCode={MOMO_PARTNER_CODE}&redirectUrl={return_url}&requestId={request_id}&requestType=captureWallet"
            sig = hmac.new(MOMO_SECRET_KEY.encode(), raw.encode(), hashlib.sha256).hexdigest()
            payload = {"partnerCode": MOMO_PARTNER_CODE, "accessKey": MOMO_ACCESS_KEY, "requestId": request_id, "amount": str(amount), "orderId": order_id, "orderInfo": order_info, "redirectUrl": return_url, "ipnUrl": webhook_url, "extraData": extra_data, "requestType": "captureWallet", "signature": sig, "lang": "vi"}
            res = requests.post(MOMO_ENDPOINT, json=payload, timeout=12)
            data = res.json()
            return data.get("payUrl") or data.get("deeplink"), data
        if provider == "vnpay" and VNPAY_TMN_CODE and VNPAY_HASH_SECRET:
            from urllib.parse import urlencode, quote_plus
            params = {
                "vnp_Version": "2.1.0", "vnp_Command": "pay", "vnp_TmnCode": VNPAY_TMN_CODE,
                "vnp_Amount": amount * 100, "vnp_CurrCode": "VND", "vnp_TxnRef": payment.payment_code,
                "vnp_OrderInfo": description, "vnp_OrderType": "other", "vnp_Locale": "vn",
                "vnp_ReturnUrl": return_url, "vnp_IpAddr": client_ip(), "vnp_CreateDate": datetime.utcnow().strftime("%Y%m%d%H%M%S")
            }
            qs = urlencode(sorted(params.items()), quote_via=quote_plus)
            sig = hmac.new(VNPAY_HASH_SECRET.encode(), qs.encode(), hashlib.sha512).hexdigest()
            return VNPAY_PAYMENT_URL + "?" + qs + "&vnp_SecureHash=" + sig, params
    except Exception as e:
        log_security("payment_gateway_error", f"{provider}: {e}")
    return None, {"provider": "bank", "message": "fallback bank transfer"}


def verify_payment_webhook(payload):
    """Xác thực webhook phổ biến: PayOS/MoMo/VNPAY hoặc HMAC nội bộ."""
    provider = (payload.get("provider") or PAYMENT_PROVIDER or "bank").lower()
    if provider == "payos" and PAYOS_CHECKSUM_KEY and payload.get("data"):
        data = payload.get("data") or {}
        payment_code = str(data.get("orderCode") or payload.get("payment_code") or "")
        status = "paid" if str(data.get("code", "")).upper() in {"00", "PAID", "SUCCESS"} else payload.get("status", "")
        return payment_code, status == "paid"
    if provider == "momo":
        return payload.get("orderId") or payload.get("payment_code"), str(payload.get("resultCode")) == "0" or payload.get("status") == "paid"
    if provider == "vnpay":
        return payload.get("vnp_TxnRef") or payload.get("payment_code"), payload.get("vnp_ResponseCode") == "00" or payload.get("status") == "paid"
    payment_code = payload.get("payment_code", "")
    status = payload.get("status", "")
    signature = payload.get("signature", "")
    expected = hmac.new(PAYMENT_SECRET.encode(), (payment_code + status).encode(), hashlib.sha256).hexdigest()
    if not hmac.compare_digest(signature, expected):
        return "", False
    return payment_code, status == "paid"

@app.route("/checkout/<plan>")
def checkout(plan):
    user = get_or_create_user()
    if plan not in PLANS:
        return "Gói không tồn tại", 404
    payment_code = (str(int(time.time() * 1000)) if PAYMENT_PROVIDER == "payos" else "AR" + str(uuid.uuid4())[:8].upper())
    p = PLANS[plan]
    payment = Payment(user_key=user.user_key, plan=plan, amount=p["price"], payment_code=payment_code)
    db.session.add(payment)
    db.session.commit()
    gateway_url, gateway_raw = create_gateway_payment(PAYMENT_PROVIDER, payment, p, user)
    return render_template("checkout.html", info=p, plan=plan, payment_code=payment_code, payment_content=payment_transfer_content(user), bank_name=BANK_NAME, bank_account=BANK_ACCOUNT, bank_owner=BANK_OWNER, gateway_url=gateway_url, provider=PAYMENT_PROVIDER, gateway_raw=gateway_raw)


def activate_payment(payment_code):
    payment = Payment.query.filter_by(payment_code=payment_code).first()
    if not payment:
        return False
    if payment.status == "paid":
        return True
    p = PLANS[payment.plan]
    payment.status = "paid"
    payment.paid_at = datetime.utcnow()
    sub = Subscription(
        user_key=payment.user_key,
        plan=payment.plan,
        status="active",
        start_at=datetime.utcnow(),
        end_at=datetime.utcnow() + timedelta(days=p["days"]),
        payment_code=payment.payment_code,
    )
    db.session.add(sub)
    db.session.commit()
    return True


@app.route("/payment/success/<payment_code>")
def payment_success_demo(payment_code):
    activate_payment(payment_code)
    return redirect(url_for("dashboard"))


@app.route("/webhook/payment", methods=["POST"])
def payment_webhook():
    """
    Webhook mẫu cho PayOS/VNPAY/Momo.
    Gửi JSON: {"payment_code":"ARxxxx", "status":"paid", "signature":"..."}
    signature demo = HMAC_SHA256(payment_code + status, PAYMENT_SECRET)
    """
    payload = request.get_json(silent=True) or request.form.to_dict() or {}
    payment_code, is_paid = verify_payment_webhook(payload)
    if not payment_code:
        log_security("payment_webhook_invalid", payload)
        return jsonify({"ok": False, "error": "invalid signature or missing payment code"}), 403
    if is_paid:
        activate_payment(payment_code)
    return jsonify({"ok": True, "payment_code": payment_code, "paid": is_paid})




def support_ai_reply(message):
    if OPENAI_API_KEY and OpenAI is not None:
        try:
            client = OpenAI(api_key=OPENAI_API_KEY)
            res = client.chat.completions.create(
                model=AI_MODEL,
                messages=[
                    {"role": "system", "content": f"Bạn là Admin AI của {APP_NAME}. Trả lời tiếng Việt ngắn gọn, lịch sự. Luôn nói đã tạo phiếu CHỜ XỬ LÝ nếu cần admin kiểm tra. Không hứa quá mức. Thông tin thanh toán: STK {BANK_ACCOUNT} {BANK_NAME}, chủ TK {BANK_OWNER}, nội dung CK: TÊN NGƯỜI DÙNG 123+CK."},
                    {"role": "user", "content": message[:2000]},
                ],
                temperature=0.35,
                max_tokens=220,
            )
            text = res.choices[0].message.content.strip()
            return "AI Admin: " + text.replace("AI Admin:", "").strip()
        except Exception as e:
            log_security("openai_support_error", e)
    msg = message.lower()
    if any(k in msg for k in ["thanh toán", "chuyển khoản", "stk", "qr"]):
        return f"AI Admin: Bạn vui lòng chuyển khoản đúng STK {BANK_ACCOUNT}, nội dung: TÊN NGƯỜI DÙNG 123+CK. Nếu đã chuyển, hệ thống sẽ báo admin kiểm tra và xử lý."
    if any(k in msg for k in ["đăng nhập", "google", "facebook", "tài khoản"]):
        return "AI Admin: Bạn kiểm tra lại email/mật khẩu hoặc dùng nút Google/Facebook. Nếu OAuth chưa hoạt động, admin cần cấu hình Client ID/Secret trong biến môi trường."
    if any(k in msg for k in ["lỗi", "hack", "report", "bảo mật", "không chạy"]):
        return "AI Admin: Mình đã ghi nhận vấn đề bảo mật/lỗi. Yêu cầu của bạn đang ở trạng thái CHỜ XỬ LÝ, admin sẽ kiểm tra log và phản hồi sớm."
    return "AI Admin: Mình đã tiếp nhận yêu cầu. Trạng thái hiện tại: CHỜ XỬ LÝ. Admin sẽ xem và phản hồi khi có kết quả."


@app.route("/support", methods=["GET", "POST"])
def support():
    user = get_or_create_user()
    if request.method == "POST":
        name = clean_text(request.form.get("name"), 180) or user.name
        email = clean_text(request.form.get("email"), 180) or user.email
        topic = clean_text(request.form.get("topic"), 180)
        message = clean_text(request.form.get("message"), 2500)
        if not message:
            return render_template("support.html", user=user, tickets=[], error="Bạn chưa nhập nội dung hỗ trợ.")
        ticket_code = "SP" + str(uuid.uuid4())[:8].upper()
        ticket = SupportTicket(ticket_code=ticket_code, user_key=user.user_key, name=name, email=email, topic=topic, status="waiting")
        db.session.add(ticket)
        db.session.add(SupportMessage(ticket_code=ticket_code, sender="user", message=message))
        db.session.add(SupportMessage(ticket_code=ticket_code, sender="ai_admin", message=support_ai_reply(message)))
        db.session.commit()
        return redirect(url_for("support", created=ticket_code))
    tickets = SupportTicket.query.filter_by(user_key=user.user_key).order_by(SupportTicket.id.desc()).limit(10).all()
    return render_template("support.html", user=user, tickets=tickets, created=request.args.get("created"))


@app.route("/api/support-chat", methods=["POST"])
def support_chat_api():
    # Chat nhanh dùng X-CSRF-Token để tránh bot spam từ web ngoài.
    verify_csrf()
    user = get_or_create_user()
    payload = request.get_json(silent=True) or {}
    message = clean_text(payload.get("message"), 2000)
    ticket_code = clean_text(payload.get("ticket_code"), 40) or "SP" + str(uuid.uuid4())[:8].upper()
    if not message:
        return jsonify({"ok": False, "error": "empty message"}), 400
    ticket = SupportTicket.query.filter_by(ticket_code=ticket_code).first()
    if not ticket:
        ticket = SupportTicket(ticket_code=ticket_code, user_key=user.user_key, name=user.name, email=user.email, topic="Chat hỗ trợ", status="waiting")
        db.session.add(ticket)
    reply = support_ai_reply(message)
    db.session.add(SupportMessage(ticket_code=ticket_code, sender="user", message=message))
    db.session.add(SupportMessage(ticket_code=ticket_code, sender="ai_admin", message=reply))
    db.session.commit()
    return jsonify({"ok": True, "ticket_code": ticket_code, "reply": reply, "status": "waiting"})


@app.route("/security")
def security_page():
    return render_template("security.html", rate=RATE_LIMIT_PER_MINUTE)

@app.route("/robots.txt")
def robots():
    txt = "User-agent: *\nAllow: /\nSitemap: " + DOMAIN_URL.rstrip("/") + "/sitemap.xml\n"
    return Response(txt, mimetype="text/plain")


@app.route("/sitemap.xml")
def sitemap():
    pages = ["", "/pricing", "/app"]
    xml = ['<?xml version="1.0" encoding="UTF-8"?>', '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">']
    for p in pages:
        xml.append(f"<url><loc>{DOMAIN_URL.rstrip('/')}{p}</loc><changefreq>weekly</changefreq><priority>0.8</priority></url>")
    xml.append("</urlset>")
    return Response("\n".join(xml), mimetype="application/xml")




def require_admin_token():
    token = request.headers.get("X-Admin-Token") or request.args.get("token")
    if not token or not hmac.compare_digest(token, ADMIN_TOKEN):
        log_security("admin_token_invalid", request.path)
        abort(403)

@app.route("/admin/security-logs")
def admin_security_logs():
    require_admin_token()
    logs = SecurityLog.query.order_by(SecurityLog.id.desc()).limit(100).all()
    return jsonify([{"time": x.created_at.isoformat(), "ip": x.ip, "event": x.event, "detail": x.detail} for x in logs])

@app.route("/admin/backup")
def admin_backup():
    require_admin_token()
    backup_dir = BASE_DIR / "backups"
    backup_dir.mkdir(exist_ok=True)
    stamp = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
    zip_path = backup_dir / f"autoreport_backup_{stamp}.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        db_file = BASE_DIR / "autoreport_cloud.db"
        if db_file.exists():
            z.write(db_file, "autoreport_cloud.db")
        for folder in [OUTPUT_DIR, UPLOAD_DIR]:
            if folder.exists():
                for fp in folder.rglob("*"):
                    if fp.is_file():
                        z.write(fp, str(fp.relative_to(BASE_DIR)))
    return send_file(zip_path, as_attachment=True)

@app.route("/api/status")
def api_status():
    user = get_or_create_user()
    return jsonify({
        "app": APP_NAME,
        "user": user.email or user.user_key,
        "used_today": usage_count(user.user_key),
        "free_limit": FREE_DAILY_LIMIT,
        "subscription": bool(active_subscription(user.user_key)),
        "ai_enabled": bool(OPENAI_API_KEY),
        "database": "postgresql" if "postgresql" in database_url else "sqlite"
    })


if __name__ == "__main__":
    with app.app_context():
        db.create_all()
    app.run(debug=True, port=5000)


def generate_ai_report(topic):
    try:
        response = client.responses.create(
            model="gpt-4.1-mini",
            input=f"Viết báo cáo chuyên nghiệp bằng tiếng Việt về chủ đề: {topic}"
        )
        return response.output_text
    except Exception as e:
        return f"Lỗi AI: {str(e)}"

