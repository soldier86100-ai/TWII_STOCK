"""
run_and_send.py
================================================================
由 GitHub Actions 每天 08:00（台灣時間）自動執行：
  1. 判斷今天是否為交易日（週末 / 台灣國定假日 → 跳過）
  2. 執行 daily_strategy_report.py 產生 PPTX
  3. 用 Gmail 寄信附檔到指定信箱
================================================================
"""
import os
import sys
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import date
from pathlib import Path
# ── 1. 判斷是否為台灣交易日 ────────────────────────────────────
try:
    import holidays
except ImportError:
    print("⚠️  holidays 套件未安裝，跳過假日判斷")
    holidays = None
today = date.today()
# 週末（週六=5, 週日=6）→ 跳過
if today.weekday() >= 5:
    print(f"📅 今天是週末（{today}），不產生日報。")
    sys.exit(0)
# 台灣國定假日 → 跳過
if holidays is not None:
    tw_holidays = holidays.country_holidays("TW", years=today.year)
    if today in tw_holidays:
        holiday_name = tw_holidays[today]
        print(f"🎌 今天是台灣國定假日（{today}：{holiday_name}），不產生日報。")
        sys.exit(0)
print(f"✅ 今天是交易日（{today}），開始產生台指策略日報...")
print("=" * 60)
# ── 2. 執行日報產生 ─────────────────────────────────────────────
try:
    from daily_strategy_report import generate_daily_report
    output_path = generate_daily_report()
except Exception as e:
    print(f"❌ 日報產生失敗：{e}")
    sys.exit(1)
if not Path(str(output_path)).exists():
    print(f"❌ 找不到輸出檔案：{output_path}")
    sys.exit(1)
print(f"\n✅ 日報檔案已產生：{output_path}")
# ── 3. 寄送 Email ───────────────────────────────────────────────
gmail_user = os.environ.get("GMAIL_USER", "")
gmail_pass = os.environ.get("GMAIL_PASS", "")
to_email   = os.environ.get("TO_EMAIL", "")
if not all([gmail_user, gmail_pass, to_email]):
    print("❌ 環境變數 GMAIL_USER / GMAIL_PASS / TO_EMAIL 未設定")
    sys.exit(1)
today_str = today.strftime("%Y.%m.%d")
subject   = f"【台股策略日報】{today_str}"
body      = f"""\
您好，
附件為今日台股策略日報（{today_str}），請查收。
此信由 GitHub Actions 系統自動寄出，請勿直接回覆。
"""
# 建立郵件
msg = MIMEMultipart()
msg["From"]    = gmail_user
msg["To"]      = to_email
msg["Subject"] = subject
msg.attach(MIMEText(body, "plain", "utf-8"))
# 附上 PPTX 附件（改用純英文檔名，避免 noname 亂碼問題）
en_filename = f"TW_Strategy_Report_{today_str}.pptx"
with open(str(output_path), "rb") as f:
    part = MIMEBase("application", "octet-stream")
    part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f'attachment; filename="{en_filename}"')
    msg.attach(part)
# 寄出
print(f"\n📧 寄送中：{gmail_user} → {to_email}")
try:
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(gmail_user, gmail_pass)
        server.sendmail(gmail_user, to_email, msg.as_string())
    print(f"✅ 日報已成功寄出至 {to_email}")
except smtplib.SMTPAuthenticationError:
    print("❌ Gmail 驗證失敗，請確認應用程式密碼是否正確")
    sys.exit(1)
except Exception as e:
    print(f"❌ 寄信失敗：{e}")
    sys.exit(1)
