import base64
import html
import jpholiday
import secrets
import msal
import shutil
import time
from pathlib import Path

from calendar import monthrange
from datetime import date, datetime, timedelta, timezone
from email.utils import parsedate_to_datetime
import json
import os
import re
import requests
import sqlite3
import unicodedata

from dotenv import load_dotenv

from flask import (
    Flask,
    abort,
    has_request_context,
    make_response,
    g,
    redirect,
    render_template,
    request,
    session,
    url_for,
)

from werkzeug.security import (
    check_password_hash,
    generate_password_hash,
)

from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request as GoogleRequest
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from cryptography.fernet import Fernet

# .env を最初に読み込む
load_dotenv()


app = Flask(__name__)

# 採用担当者向けデモ版。既存のローカルアプリの機能を保持して分離する。
DEMO_MODE = os.environ.get("INTERNTRACK_DEMO_MODE") == "1"
DEMO_HOME = Path(os.environ.get("INTERNTRACK_DEMO_DIR", "/tmp/interntrack-demo"))
if DEMO_MODE:
    DEMO_HOME.mkdir(parents=True, exist_ok=True)
    if not os.environ.get("INTERNTRACK_SECRET_KEY"):
        raise RuntimeError("デモ公開にはINTERNTRACK_SECRET_KEYの設定が必要です")

app.config["SECRET_KEY"] = os.environ.get(
    "INTERNTRACK_SECRET_KEY",
    "interntrack-local-development-key",
)
if DEMO_MODE:
    app.config.update(
        SESSION_COOKIE_HTTPONLY=True,
        SESSION_COOKIE_SAMESITE="Lax",
        SESSION_COOKIE_SECURE=os.environ.get("INTERNTRACK_DEMO_LOCAL") != "1",
        MAX_CONTENT_LENGTH=65536,
        PREFERRED_URL_SCHEME="https",
    )

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DATABASE = (
    str(DEMO_HOME / "seed-template.db") if DEMO_MODE
    else os.path.join(BASE_DIR, "interntrack.db")
)

GOOGLE_CLIENT_SECRETS_FILE = os.path.join(
    BASE_DIR,
    "google_oauth_client.json",
)

GOOGLE_OAUTH_SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
]

GOOGLE_OAUTH_REDIRECT_URI = "http://127.0.0.1:5000/oauth2callback"

INTERNTRACK_TOKEN_KEY = os.environ.get("INTERNTRACK_TOKEN_KEY")

# Microsoft / Outlook OAuth
MICROSOFT_CLIENT_ID = os.environ.get("MICROSOFT_CLIENT_ID")
MICROSOFT_CLIENT_SECRET = os.environ.get("MICROSOFT_CLIENT_SECRET")

MICROSOFT_AUTHORITY = os.environ.get(
    "MICROSOFT_AUTHORITY",
    "https://login.microsoftonline.com/common",
)

MICROSOFT_REDIRECT_URI = os.environ.get(
    "MICROSOFT_REDIRECT_URI",
    "http://localhost:5000/outlook/callback",
)

MICROSOFT_SCOPES = [
    "User.Read",
    "Mail.Read",
]


def get_token_cipher():
    """OAuthトークンを暗号化・復号するFernetを返す。"""

    if not INTERNTRACK_TOKEN_KEY:
        raise RuntimeError("INTERNTRACK_TOKEN_KEYが設定されていません。")

    try:
        return Fernet(INTERNTRACK_TOKEN_KEY.encode("utf-8"))

    except (TypeError, ValueError) as error:
        raise RuntimeError("INTERNTRACK_TOKEN_KEYの形式が正しくありません。") from error


def encrypt_token(token):
    """文字列トークンを暗号化してDB保存用文字列にする。"""

    if not token:
        return None

    cipher = get_token_cipher()

    return cipher.encrypt(token.encode("utf-8")).decode("utf-8")


def decrypt_token(encrypted_token):
    """DBに保存した暗号化トークンを復号する。"""

    if not encrypted_token:
        return None

    cipher = get_token_cipher()

    return cipher.decrypt(encrypted_token.encode("utf-8")).decode("utf-8")


def load_google_client_config():
    """Google OAuthクライアントJSONからWebアプリ設定を読み込む。"""

    with open(
        GOOGLE_CLIENT_SECRETS_FILE,
        "r",
        encoding="utf-8",
    ) as file:
        configuration = json.load(file)

    web_configuration = configuration.get("web")

    if not web_configuration:
        raise RuntimeError("google_oauth_client.jsonにweb設定がありません。")

    return web_configuration


def parse_google_expiry(expires_at):
    """
    DBに保存したGoogle OAuthの有効期限を、
    google-authが比較できるUTCのdatetimeへ変換する。
    """

    if not expires_at:
        return None

    try:
        expiry = datetime.fromisoformat(expires_at.replace("Z", "+00:00"))

    except (TypeError, ValueError):
        return None

    # google-authはタイムゾーンなしのUTC日時と比較するため、
    # タイムゾーン付き日時はUTCへ変換してからtzinfoを外す。
    if expiry.tzinfo is not None:
        expiry = expiry.astimezone(timezone.utc).replace(tzinfo=None)

    return expiry


def build_google_credentials(token_row):
    """暗号化保存されたOAuth情報からGoogle Credentialsを復元する。"""

    client_configuration = load_google_client_config()
    scopes = token_row["scopes"].split() if token_row["scopes"] else GOOGLE_OAUTH_SCOPES

    return Credentials(
        token=decrypt_token(token_row["access_token"]),
        refresh_token=decrypt_token(token_row["refresh_token"]),
        token_uri=(
            token_row["token_uri"]
            or client_configuration.get("token_uri")
            or "https://oauth2.googleapis.com/token"
        ),
        client_id=(token_row["client_id"] or client_configuration.get("client_id")),
        client_secret=client_configuration.get("client_secret"),
        scopes=scopes,
        expiry=parse_google_expiry(token_row["expires_at"]),
    )


def get_gmail_header(message, header_name):
    """Gmailメッセージの指定ヘッダーを取得する。"""

    headers = message.get("payload", {}).get("headers", [])

    for header in headers:
        if header.get("name", "").casefold() == header_name.casefold():
            return header.get("value", "").strip()

    return ""


def normalize_gmail_received_at(message):
    """Gmailメッセージの受信日時をISO文字列へ整形する。"""

    internal_date = message.get("internalDate")

    if internal_date:
        try:
            received = datetime.fromtimestamp(
                int(internal_date) / 1000,
                tz=timezone.utc,
            )
            return received.isoformat()
        except (TypeError, ValueError, OSError):
            pass

    date_header = get_gmail_header(message, "Date")

    if date_header:
        try:
            received = parsedate_to_datetime(date_header)

            if received.tzinfo is None:
                received = received.replace(tzinfo=timezone.utc)

            return received.isoformat()
        except (TypeError, ValueError, OverflowError):
            pass

    return None


def clean_email_text(value):
    """メール解析用にHTML・空白を軽く正規化する。"""

    text = html.unescape(value or "")
    text = re.sub(
        r"<(script|style)\b[^>]*>.*?</\1>",
        " ",
        text,
        flags=re.IGNORECASE | re.DOTALL,
    )
    text = re.sub(r"<br\s*/?>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"</p\s*>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"<[^>]+>", " ", text)
    text = unicodedata.normalize("NFKC", text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def decode_gmail_body_data(data):
    """Gmail APIのbase64url本文をUTF-8文字列へ変換する。"""

    if not data:
        return ""

    try:
        padded = data + ("=" * (-len(data) % 4))
        return base64.urlsafe_b64decode(
            padded.encode("ascii")
        ).decode("utf-8", errors="replace")
    except (ValueError, TypeError, UnicodeError):
        return ""


def extract_gmail_body_text(message):
    """Gmailメッセージから解析用本文を取り出す。plain textを優先する。"""

    plain_parts = []
    html_parts = []

    def walk(part):
        mime_type = (part.get("mimeType") or "").casefold()
        body = part.get("body") or {}
        decoded = decode_gmail_body_data(body.get("data"))

        if decoded:
            if mime_type == "text/plain":
                plain_parts.append(decoded)
            elif mime_type == "text/html":
                html_parts.append(decoded)

        for child in part.get("parts") or []:
            walk(child)

    walk(message.get("payload") or {})

    if plain_parts:
        return clean_email_text("\n".join(plain_parts))

    if html_parts:
        return clean_email_text("\n".join(html_parts))

    return ""


# =========================================================
# データベース接続
# =========================================================


def get_db_connection():
    path = DATABASE
    if DEMO_MODE and has_request_context():
        # セッションは署名済みcookieに保持し、異なる訪問者には別DBを割り当てる。
        identifier = session.get("demo_id", "")
        if re.fullmatch(r"[0-9a-f]{32}", identifier):
            path = str(DEMO_HOME / f"visit-{identifier}.db")
    connection = sqlite3.connect(
        path,
        timeout=30,
    )

    connection.row_factory = sqlite3.Row

    connection.execute("PRAGMA foreign_keys = ON")

    connection.execute("PRAGMA busy_timeout = 30000")

    return connection


# =========================================================
# 選考履歴からステータスを判定
# =========================================================


def determine_status(event_type, title):
    """
    選考履歴の種類とタイトルから、
    企業の現在ステータスを判定する。
    """

    event_type = (event_type or "").strip()
    title = (title or "").strip()

    combined_text = f"{event_type} {title}".lower()

    # 結果関係
    if "不合格" in combined_text or "お祈り" in combined_text:
        return "不合格"

    if "内定" in combined_text or "合格" in combined_text:
        return "合格"

    if "結果待ち" in combined_text or event_type == "結果通知":
        return "結果待ち"

    # 面接関係
    if "最終面接" in combined_text:
        return "最終面接"

    if "三次面接" in combined_text:
        return "三次面接"

    if "二次面接" in combined_text:
        return "二次面接"

    if "一次面接" in combined_text:
        return "一次面接"

    if "面接" in combined_text:
        return "面接"

    # 適性検査関係
    if (
        "適性検査" in combined_text
        or "spi" in combined_text
        or "webテスト" in combined_text
        or "web test" in combined_text
        or "玉手箱" in combined_text
        or "テストセンター" in combined_text
    ):
        return "適性検査"

    # ES関係
    if "es提出" in combined_text or "エントリーシート提出" in combined_text:
        return "ES提出済"

    if "es作成" in combined_text or "エントリーシート作成" in combined_text:
        return "ES作成中"

    if event_type == "ES":
        return "ES提出済"

    # その他
    if event_type == "インターン":
        return "インターン"

    if event_type == "説明会":
        return "説明会"

    # 自動判定できない場合
    return None


def update_company_status_from_events(
    connection,
    company_id,
):
    """
    残っている最新の選考履歴から
    企業ステータスを再計算する。
    """

    events = connection.execute(
        """
        SELECT event_type, title
        FROM company_events
        WHERE company_id = ?
        ORDER BY id DESC
        """,
        (company_id,),
    ).fetchall()

    for event in events:
        status = determine_status(
            event["event_type"],
            event["title"],
        )

        if status:
            connection.execute(
                """
                UPDATE companies
                SET status = ?
                WHERE id = ?
                """,
                (
                    status,
                    company_id,
                ),
            )

            return

    # 判定可能な履歴が1件もない場合は未応募に戻す
    connection.execute(
        """
        UPDATE companies
        SET status = ?
        WHERE id = ?
        """,
        (
            "未応募",
            company_id,
        ),
    )


# =========================================================
# Gmailメールと企業の自動紐付け
# =========================================================


def normalize_company_match_text(value):
    """
    会社名やメール本文を比較しやすい文字列へ変換する。
    大文字・小文字、全角・半角、空白、記号の違いを吸収する。
    """

    normalized = unicodedata.normalize(
        "NFKC",
        value or "",
    ).casefold()

    return re.sub(
        r"[^0-9a-zぁ-んァ-ヶ一-龠]+",
        "",
        normalized,
    )


def build_company_match_aliases(company_name):
    """
    企業名から照合用の候補名を作成する。
    例:
        株式会社NTTデータ
        → 株式会社nttデータ / nttデータ
    """

    raw_name = unicodedata.normalize(
        "NFKC",
        company_name or "",
    ).strip()

    aliases = {
        normalize_company_match_text(raw_name),
    }

    removable_words = (
        "株式会社",
        "有限会社",
        "合同会社",
        "一般社団法人",
        "一般財団法人",
        "Inc.",
        "Inc",
        "Ltd.",
        "Ltd",
        "Co., Ltd.",
        "Co Ltd",
        "Corporation",
        "Corp.",
        "Corp",
        "Holdings",
        "ホールディングス",
        "グループ",
    )

    simplified_name = raw_name

    for removable_word in removable_words:
        simplified_name = simplified_name.replace(
            removable_word,
            "",
        )

    simplified_alias = normalize_company_match_text(simplified_name)

    if simplified_alias:
        aliases.add(simplified_alias)

    # 短すぎる語は誤判定しやすいため除外する。
    return {alias for alias in aliases if len(alias) >= 3}


def extract_sender_domain(sender):
    """
    Fromヘッダーからメールドメインを取り出す。
    """

    match = re.search(
        r"@([a-z0-9.-]+\.[a-z]{2,})",
        (sender or "").casefold(),
    )

    if match is None:
        return ""

    return match.group(1)


def find_matching_company(
    connection,
    user_id,
    subject,
    sender,
    snippet,
):
    """
    件名・送信者・スニペットから登録企業を推定する。

    戻り値:
        (company_id, match_method, confidence)
        一意に判定できない場合は (None, None, None)
    """

    companies = connection.execute(
        """
        SELECT id, name
        FROM companies
        WHERE user_id = ?
        ORDER BY length(name) DESC, id ASC
        """,
        (user_id,),
    ).fetchall()

    if not companies:
        return None, None, None

    subject_text = subject or ""
    sender_text = sender or ""
    snippet_text = snippet or ""

    combined_text = " ".join(
        (
            subject_text,
            sender_text,
            snippet_text,
        )
    )

    normalized_combined = normalize_company_match_text(combined_text)
    sender_domain = extract_sender_domain(sender_text)
    normalized_domain = normalize_company_match_text(sender_domain)

    candidates = []

    for company in companies:
        company_name = company["name"] or ""
        aliases = build_company_match_aliases(company_name)

        best_score = 0
        best_method = None

        # 会社名が件名・本文などにそのまま含まれる場合を最優先。
        if company_name and company_name.casefold() in combined_text.casefold():
            best_score = 100
            best_method = "exact_name"

        # 全角・空白・記号を除いた状態で一致する場合。
        for alias in aliases:
            if alias in normalized_combined:
                score = 92

                if alias in normalize_company_match_text(subject_text):
                    score = 97

                if score > best_score:
                    best_score = score
                    best_method = "normalized_name"

            # 英数字4文字以上の企業名が送信ドメインに含まれる場合。
            if (
                len(alias) >= 4
                and re.fullmatch(
                    r"[0-9a-z]+",
                    alias,
                )
                and alias in normalized_domain
                and 82 > best_score
            ):
                best_score = 82
                best_method = "sender_domain"

        if best_score > 0:
            candidates.append(
                {
                    "company_id": company["id"],
                    "score": best_score,
                    "method": best_method,
                }
            )

    if not candidates:
        return None, None, None

    candidates.sort(
        key=lambda candidate: (
            candidate["score"],
            -candidate["company_id"],
        ),
        reverse=True,
    )

    best_candidate = candidates[0]

    # 同点で複数企業が候補になる場合は、誤紐付けを避ける。
    if len(candidates) >= 2 and candidates[1]["score"] == best_candidate["score"]:
        return None, None, None

    return (
        best_candidate["company_id"],
        best_candidate["method"],
        best_candidate["score"],
    )


def update_email_message_company_link(
    connection,
    user_id,
    email_account_id,
    provider,
    message_id,
    subject,
    sender,
    snippet,
):
    """
    取得したGmailメール1件を企業へ自動紐付けする。
    """

    (
        company_id,
        match_method,
        match_confidence,
    ) = find_matching_company(
        connection,
        user_id,
        subject,
        sender,
        snippet,
    )

    connection.execute(
        """
        UPDATE email_messages
        SET
            company_id = ?,
            match_method = ?,
            match_confidence = ?,
            matched_at = CASE
                WHEN ? IS NULL THEN NULL
                ELSE CURRENT_TIMESTAMP
            END
          WHERE user_id = ?
          AND email_account_id = ?
          AND provider = ?
          AND message_id = ?
        """,
        (
            (
                company_id,
                match_method,
                match_confidence,
                company_id,
                user_id,
                email_account_id,
                provider,
                message_id,
            )
        ),
    )

    return company_id


EMAIL_COMPANY_TARGET_COUNT = 20
EMAIL_BACKFILL_SCAN_LIMIT = 500


def count_company_linked_emails(
    connection,
    user_id,
    provider=None,
):
    """InternTrack上で登録企業に紐付いているメール件数を返す。"""

    conditions = [
        "email_messages.user_id = ?",
        "email_messages.company_id IS NOT NULL",
    ]
    parameters = [user_id]

    if provider:
        conditions.append("email_messages.provider = ?")
        parameters.append(provider)

    row = connection.execute(
        f"""
        SELECT COUNT(*) AS count
        FROM email_messages
        WHERE {' AND '.join(conditions)}
        """,
        parameters,
    ).fetchone()

    return int(row["count"] or 0)


def has_registered_companies(connection, user_id):
    """ログイン中ユーザーに登録企業が1社以上あるか確認する。"""

    row = connection.execute(
        """
        SELECT 1
        FROM companies
        WHERE user_id = ?
        LIMIT 1
        """,
        (user_id,),
    ).fetchone()

    return row is not None


# =========================================================
# メールから企業イベントを自動生成
# =========================================================


def normalize_email_analysis_text(*parts):
    """件名・概要・本文をイベント解析しやすい文字列へまとめる。"""

    text = "\n".join(
        str(part)
        for part in parts
        if part
    )
    text = unicodedata.normalize("NFKC", text)
    text = text.replace("\u3000", " ")
    text = re.sub(r"[ \t]+", " ", text)
    return text.strip()


def detect_email_events(subject, snippet, body=""):
    """
    メールから就活イベント種別を判定する。

    1通に「説明会日時」と「予約締切」などが併記される場合に備え、
    複数イベントを返せる形式にする。
    """

    normalized_text = normalize_email_analysis_text(
        subject,
        snippet,
        body,
    ).casefold()

    events = []
    used_keys = set()

    def add_event(event_key, event_type, title, date_role="event"):
        if event_key in used_keys:
            return
        used_keys.add(event_key)
        events.append(
            {
                "event_key": event_key,
                "event_type": event_type,
                "title": title,
                "date_role": date_role,
            }
        )

    # 結果通知は予定ではないため、日付が本文になければカレンダーには出さない。
    negative_keywords = (
        "不合格",
        "不採用",
        "選考見送り",
        "選考を見送",
        "ご期待に添えない",
        "ご希望に添えない",
        "今回は見送",
        "お祈り申し上げ",
    )
    if any(keyword.casefold() in normalized_text for keyword in negative_keywords):
        add_event("result", "結果通知", "不合格", "result")
        return events

    offer_keywords = (
        "内定",
        "内々定",
    )
    if any(keyword.casefold() in normalized_text for keyword in offer_keywords):
        add_event("result", "結果通知", "内定", "result")
        return events

    pass_keywords = (
        "選考通過",
        "合格のお知らせ",
        "合格のご連絡",
        "合格通知",
    )
    if any(keyword.casefold() in normalized_text for keyword in pass_keywords):
        add_event("result", "結果通知", "合格", "result")
        return events

    interview_rules = (
        (
            "最終面接",
            ("最終面接", "最終選考"),
        ),
        (
            "三次面接",
            ("三次面接", "3次面接", "第三次面接"),
        ),
        (
            "二次面接",
            ("二次面接", "2次面接", "第二次面接"),
        ),
        (
            "一次面接",
            ("一次面接", "1次面接", "第一次面接"),
        ),
    )

    interview_found = False
    for title, keywords in interview_rules:
        if any(keyword.casefold() in normalized_text for keyword in keywords):
            add_event("interview", "面接", title, "event")
            interview_found = True
            break

    interview_keywords = (
        "面接のご案内",
        "面接日時",
        "面接日程",
        "面接予約",
        "面接について",
        "面接のお知らせ",
        "面談のご案内",
        "面談日時",
    )
    if (
        not interview_found
        and any(keyword.casefold() in normalized_text for keyword in interview_keywords)
    ):
        add_event("interview", "面接", "面接", "event")
        interview_found = True

    test_keywords = (
        "適性検査",
        "webテスト",
        "web test",
        "spi受検",
        "spi受験",
        "spi試験",
        "玉手箱",
        "テストセンター",
    )
    test_found = any(
        keyword.casefold() in normalized_text
        for keyword in test_keywords
    )
    if test_found:
        add_event("test", "適性検査", "適性検査", "deadline")

    es_keywords = (
        "es提出",
        "エントリーシート提出",
        "エントリーシートの提出",
        "応募書類提出",
        "応募書類の提出",
        "履歴書提出",
        "履歴書の提出",
    )
    es_found = any(
        keyword.casefold() in normalized_text
        for keyword in es_keywords
    )
    if es_found:
        add_event("es", "ES", "ES提出締切", "deadline")

    briefing_keywords = (
        "会社説明会",
        "企業説明会",
        "説明会のご案内",
        "説明会予約",
        "説明会のお知らせ",
    )
    briefing_found = any(
        keyword.casefold() in normalized_text
        for keyword in briefing_keywords
    )
    if briefing_found:
        add_event("briefing", "説明会", "説明会", "event")

    intern_keywords = (
        "インターンシップ",
        "インターン参加",
        "インターンのご案内",
        "インターン選考",
        "仕事体験",
        "オープンカンパニー",
    )
    intern_found = any(
        keyword.casefold() in normalized_text
        for keyword in intern_keywords
    )
    if intern_found:
        add_event("intern", "インターン", "インターン", "event")

    deadline_keywords = (
        "応募締切",
        "応募期限",
        "エントリー締切",
        "エントリー期限",
        "申込締切",
        "申込期限",
        "申し込み締切",
        "申し込み期限",
        "予約締切",
        "予約期限",
        "提出締切",
        "提出期限",
        "受検期限",
        "受験期限",
    )
    has_explicit_deadline = any(
        keyword.casefold() in normalized_text
        for keyword in deadline_keywords
    )

    # 1通に開催日時と予約締切の両方がある場合は、締切も別予定として保持する。
    if has_explicit_deadline:
        if interview_found and (
            "面接予約締切" in normalized_text
            or "面接予約期限" in normalized_text
        ):
            add_event(
                "interview_deadline",
                "締切",
                "面接予約締切",
                "deadline",
            )
        elif briefing_found:
            add_event(
                "briefing_deadline",
                "締切",
                "説明会予約締切",
                "deadline",
            )
        elif intern_found:
            add_event(
                "intern_deadline",
                "締切",
                "インターン応募締切",
                "deadline",
            )
        elif not es_found and not test_found:
            add_event(
                "deadline",
                "締切",
                "応募・予約締切",
                "deadline",
            )

    return events


def detect_email_event(subject, snippet, body=""):
    """既存コードとの互換用。最初に検出されたイベントを返す。"""

    events = detect_email_events(subject, snippet, body)
    if not events:
        return None
    return events[0]


def parse_email_received_date(received_at):
    """メール受信日時から日付だけ安全に取り出す。"""

    if not received_at:
        return date.today()

    try:
        return date.fromisoformat(str(received_at)[:10])
    except (TypeError, ValueError):
        return date.today()


def extract_email_event_schedule(
    subject,
    snippet,
    body,
    received_at,
    event_type,
    title,
    date_role="event",
):
    """
    メール本文中の日付・時刻を抽出する。

    対応例:
        2026年10月20日
        2026/10/20
        10月20日(火)
        10/20 14:00
        10月20日 23時59分まで
    """

    text = normalize_email_analysis_text(
        subject,
        snippet,
        body,
    )
    if not text:
        return {
            "event_date": None,
            "event_time": None,
            "matched_text": None,
        }

    folded_text = text.casefold()
    base_date = parse_email_received_date(received_at)
    candidates = []
    full_date_spans = []

    full_date_pattern = re.compile(
        r"(?<!\d)"
        r"(20\d{2})"
        r"\s*(?:年|[./-])\s*"
        r"(\d{1,2})"
        r"\s*(?:月|[./-])\s*"
        r"(\d{1,2})"
        r"\s*日?"
    )

    for match in full_date_pattern.finditer(text):
        try:
            candidate_date = date(
                int(match.group(1)),
                int(match.group(2)),
                int(match.group(3)),
            )
        except ValueError:
            continue

        full_date_spans.append(match.span())
        candidates.append(
            {
                "date": candidate_date,
                "start": match.start(),
                "end": match.end(),
                "explicit_year": True,
            }
        )

    short_date_pattern = re.compile(
        r"(?<!\d)"
        r"(\d{1,2})"
        r"\s*(?:月|[/-])\s*"
        r"(\d{1,2})"
        r"\s*日?"
    )

    for match in short_date_pattern.finditer(text):
        if any(
            span_start <= match.start() < span_end
            for span_start, span_end in full_date_spans
        ):
            continue

        month = int(match.group(1))
        day = int(match.group(2))

        try:
            candidate_date = date(
                base_date.year,
                month,
                day,
            )
        except ValueError:
            continue

        # 年末に翌年の日程が案内されるケースを補正する。
        if candidate_date < base_date - timedelta(days=90):
            try:
                candidate_date = date(
                    base_date.year + 1,
                    month,
                    day,
                )
            except ValueError:
                continue

        candidates.append(
            {
                "date": candidate_date,
                "start": match.start(),
                "end": match.end(),
                "explicit_year": False,
            }
        )

    if not candidates:
        return {
            "event_date": None,
            "event_time": None,
            "matched_text": None,
        }

    event_keyword_map = {
        "面接": (
            ("面接日時", 11),
            ("面接日程", 11),
            ("面接", 8),
            ("面談", 7),
            ("選考", 4),
        ),
        "ES": (
            ("エントリーシート", 9),
            ("es", 6),
            ("提出", 7),
        ),
        "適性検査": (
            ("適性検査", 9),
            ("webテスト", 9),
            ("spi", 9),
            ("受検", 7),
            ("受験", 7),
        ),
        "説明会": (
            ("説明会日時", 11),
            ("説明会", 9),
        ),
        "インターン": (
            ("インターンシップ", 9),
            ("インターン", 9),
            ("仕事体験", 9),
            ("オープンカンパニー", 9),
        ),
        "締切": (
            ("応募", 6),
            ("申込", 6),
            ("申し込み", 6),
            ("予約", 6),
            ("提出", 6),
        ),
    }

    def keyword_distance(candidate, keyword):
        keyword_text = keyword.casefold()
        best_distance = None

        for match in re.finditer(
            re.escape(keyword_text),
            folded_text,
        ):
            if match.end() <= candidate["start"]:
                distance = candidate["start"] - match.end()
            elif match.start() >= candidate["end"]:
                distance = match.start() - candidate["end"]
            else:
                distance = 0

            if best_distance is None or distance < best_distance:
                best_distance = distance

        return best_distance

    def proximity_score(candidate, keyword, weight):
        distance = keyword_distance(candidate, keyword)

        if distance is None or distance > 40:
            return 0
        if distance <= 6:
            return weight * 2
        if distance <= 16:
            return weight
        return max(1, weight // 2)

    scored_candidates = []

    for candidate in candidates:
        score = 2 if candidate["explicit_year"] else 0

        if candidate["date"] >= base_date:
            score += 2

        day_distance = abs((candidate["date"] - base_date).days)
        if day_distance <= 370:
            score += 1

        for keyword, weight in event_keyword_map.get(event_type, ()):
            score += proximity_score(
                candidate,
                keyword,
                weight,
            )

        if date_role == "deadline":
            score += proximity_score(candidate, "締切", 14)
            score += proximity_score(candidate, "期限", 14)
            score += proximity_score(candidate, "まで", 7)
            score += proximity_score(candidate, "提出", 5)
            score += proximity_score(candidate, "応募", 5)
            score += proximity_score(candidate, "予約", 5)
            score += proximity_score(candidate, "受検", 5)
            score += proximity_score(candidate, "受験", 5)

            # 開催日時より締切日を優先する。
            if keyword_distance(candidate, "開催") is not None:
                if keyword_distance(candidate, "開催") <= 8:
                    score -= 8
            if keyword_distance(candidate, "日時") is not None:
                if keyword_distance(candidate, "日時") <= 8:
                    score -= 6

        elif date_role == "event":
            score += proximity_score(candidate, "日時", 11)
            score += proximity_score(candidate, "日程", 10)
            score += proximity_score(candidate, "開催", 10)
            score += proximity_score(candidate, "実施", 9)
            score += proximity_score(candidate, "参加", 4)

            # 開催日と締切日が同じメールにある場合、締切日を避ける。
            for deadline_word in ("締切", "期限"):
                distance = keyword_distance(candidate, deadline_word)
                if distance is not None:
                    if distance <= 8:
                        score -= 18
                    elif distance <= 16:
                        score -= 9

        scored_candidates.append(
            (
                score,
                candidate["date"],
                candidate["start"],
                candidate["end"],
            )
        )

    scored_candidates.sort(
        key=lambda item: (
            item[0],
            1 if item[1] >= base_date else 0,
            -abs((item[1] - base_date).days),
        ),
        reverse=True,
    )

    best_score, best_date, best_start, best_end = scored_candidates[0]

    # 文脈が弱すぎる数字は予定として採用しない。
    if best_score < 5:
        return {
            "event_date": None,
            "event_time": None,
            "matched_text": None,
        }

    time_text = text[best_end:min(len(text), best_end + 45)]
    if not time_text.strip():
        time_text = text[max(0, best_start - 20):min(len(text), best_end + 45)]

    event_time = None

    if "正午" in time_text:
        event_time = "12:00"
    else:
        time_match = re.search(
            r"(午前|午後)?\s*"
            r"(\d{1,2})\s*"
            r"(?::|時)\s*"
            r"(\d{1,2})?\s*分?",
            time_text,
        )

        if time_match:
            meridiem = time_match.group(1)
            hour = int(time_match.group(2))
            minute = int(time_match.group(3) or 0)

            if meridiem == "午後" and hour < 12:
                hour += 12
            elif meridiem == "午前" and hour == 12:
                hour = 0

            if 0 <= hour <= 23 and 0 <= minute <= 59:
                event_time = f"{hour:02d}:{minute:02d}"

    matched_context = text[
        max(0, best_start - 25):min(len(text), best_end + 35)
    ].strip()

    return {
        "event_date": best_date.isoformat(),
        "event_time": event_time,
        "matched_text": matched_context,
    }


def build_email_event_memo(
    provider,
    sender,
    subject,
    schedule,
):
    """メール由来イベントのメモを統一形式で作る。"""

    lines = [
        f"{provider}から自動登録",
        f"送信者: {sender or '不明'}",
        f"件名: {subject or '（件名なし）'}",
    ]

    if schedule.get("event_date"):
        schedule_text = schedule["event_date"]
        if schedule.get("event_time"):
            schedule_text += f" {schedule['event_time']}"
        lines.append(f"抽出日時: {schedule_text}")

    if schedule.get("matched_text"):
        lines.append(f"判定箇所: {schedule['matched_text'][:160]}")

    lines.append("※メール内容から自動判定した予定です。必要に応じて確認・編集してください。")
    return "\n".join(lines)


def create_event_from_email_message(
    connection,
    user_id,
    email_account_id,
    provider,
    message_id,
):
    """
    企業と紐付いたメールからcompany_eventsを自動生成する。
    1通に複数の予定が含まれる場合にも対応する。
    """

    email_message = connection.execute(
        """
        SELECT
            email_messages.id,
            email_messages.company_id,
            email_messages.subject,
            email_messages.sender,
            email_messages.snippet,
            email_messages.body,
            email_messages.received_at
        FROM email_messages
        JOIN companies
          ON companies.id = email_messages.company_id
        WHERE email_messages.user_id = ?
          AND email_messages.email_account_id = ?
          AND email_messages.provider = ?
          AND email_messages.message_id = ?
          AND email_messages.company_id IS NOT NULL
          AND companies.user_id = ?
        """,
        (
            user_id,
            email_account_id,
            provider,
            message_id,
            user_id,
        ),
    ).fetchone()

    if email_message is None:
        return 0

    detected_events = detect_email_events(
        email_message["subject"],
        email_message["snippet"],
        email_message["body"],
    )

    if not detected_events:
        return 0

    changed_count = 0

    for detected_event in detected_events:
        schedule = extract_email_event_schedule(
            email_message["subject"],
            email_message["snippet"],
            email_message["body"],
            email_message["received_at"],
            detected_event["event_type"],
            detected_event["title"],
            detected_event["date_role"],
        )

        event_title = detected_event["title"]
        if schedule.get("event_time"):
            event_title = f"{event_title} {schedule['event_time']}"

        auto_memo = build_email_event_memo(
            provider,
            email_message["sender"],
            email_message["subject"],
            schedule,
        )

        existing_event = connection.execute(
            """
            SELECT
                id,
                event_date
            FROM company_events
            WHERE source_email_provider = ?
              AND source_email_account_id = ?
              AND source_email_message_id = ?
              AND source_email_event_key = ?
            LIMIT 1
            """,
            (
                provider,
                email_account_id,
                message_id,
                detected_event["event_key"],
            ),
        ).fetchone()

        if existing_event is not None:
            # 以前の版で日付なし登録された予定だけを安全に補完する。
            if (
                (existing_event["event_date"] is None or existing_event["event_date"] == "")
                and schedule.get("event_date")
            ):
                connection.execute(
                    """
                    UPDATE company_events
                    SET
                        event_date = ?,
                        event_type = ?,
                        title = ?,
                        memo = ?
                    WHERE id = ?
                    """,
                    (
                        schedule["event_date"],
                        detected_event["event_type"],
                        event_title,
                        auto_memo,
                        existing_event["id"],
                    ),
                )
                changed_count += 1
            continue

        connection.execute(
            """
            INSERT INTO company_events (
                company_id,
                event_date,
                event_type,
                title,
                memo,
                source_email_provider,
                source_email_account_id,
                source_email_message_id,
                source_email_event_key
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                email_message["company_id"],
                schedule.get("event_date"),
                detected_event["event_type"],
                event_title,
                auto_memo,
                provider,
                email_account_id,
                message_id,
                detected_event["event_key"],
            ),
        )
        changed_count += 1

    if changed_count:
        update_company_status_from_events(
            connection,
            email_message["company_id"],
        )

    return changed_count


def reprocess_email_messages_for_user(
    connection,
    user_id,
    limit=300,
):
    """
    企業追加・名称変更後に既存メールを再照合する。
    新しく登録した企業に過去メールが該当すれば自動で予定化される。
    """

    rows = connection.execute(
        """
        SELECT
            email_account_id,
            provider,
            message_id,
            subject,
            sender,
            snippet,
            body
        FROM email_messages
        WHERE user_id = ?
        ORDER BY
            CASE
                WHEN received_at IS NULL THEN 1
                ELSE 0
            END,
            received_at DESC,
            id DESC
        LIMIT ?
        """,
        (
            user_id,
            limit,
        ),
    ).fetchall()

    event_count = 0

    for message in rows:
        analysis_text = normalize_email_analysis_text(
            message["snippet"],
            message["body"],
        )

        update_email_message_company_link(
            connection,
            user_id,
            message["email_account_id"],
            message["provider"],
            message["message_id"],
            message["subject"],
            message["sender"],
            analysis_text,
        )

        event_count += create_event_from_email_message(
            connection,
            user_id,
            message["email_account_id"],
            message["provider"],
            message["message_id"],
        )

    return event_count


# =========================================================
# テーブル作成
# =========================================================


def create_tables():
    connection = get_db_connection()

    # ユーザー情報
    connection.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            email TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
        """)

    # 企業情報
    connection.execute("""
        CREATE TABLE IF NOT EXISTS companies (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            name TEXT NOT NULL,
            status TEXT NOT NULL,
            next_schedule TEXT,
            deadline TEXT,
            mypage_url TEXT,
            memo TEXT,

            FOREIGN KEY (user_id)
                REFERENCES users (id)
                ON DELETE CASCADE
        )
        """)

    # 選考履歴
    connection.execute("""
        CREATE TABLE IF NOT EXISTS company_events (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER NOT NULL,
            event_date TEXT,
            event_type TEXT NOT NULL,
            title TEXT NOT NULL,
            memo TEXT,
            source_email_provider TEXT,
            source_email_account_id INTEGER,
            source_email_message_id TEXT,
            source_email_event_key TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,

            FOREIGN KEY (company_id)
                REFERENCES companies (id)
                ON DELETE CASCADE,

            FOREIGN KEY (source_email_account_id)
                REFERENCES email_accounts (id)
                ON DELETE SET NULL
        )
        """)

    # 企業メモ
    connection.execute("""
        CREATE TABLE IF NOT EXISTS company_notes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER NOT NULL,
            title TEXT NOT NULL,
            content TEXT NOT NULL,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,

            FOREIGN KEY (company_id)
                REFERENCES companies (id)
                ON DELETE CASCADE
        )
        """)

    # メールアカウント設定
    #
    # この段階ではパスワードやアクセストークンは保存しません。
    # Gmail・Outlookは後の工程でOAuth認証を追加し、
    # Yahoo!・iCloud・大学メールなどはIMAP設定に対応させます。

    connection.execute("""
        CREATE TABLE IF NOT EXISTS email_accounts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            provider TEXT NOT NULL,
            email_address TEXT NOT NULL,
            display_name TEXT,
            imap_server TEXT,
            imap_port INTEGER,
            use_ssl INTEGER NOT NULL DEFAULT 1,
            auth_type TEXT NOT NULL DEFAULT 'oauth',
            is_active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,

            FOREIGN KEY (user_id)
                REFERENCES users (id)
                ON DELETE CASCADE,

            UNIQUE (user_id, email_address)
        )
        """)

    # Google OAuthトークン
    # access_token / refresh_token はFernetで暗号化して保存します。
    connection.execute("""
        CREATE TABLE IF NOT EXISTS google_oauth_tokens (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            email_account_id INTEGER NOT NULL,
            access_token TEXT NOT NULL,
            refresh_token TEXT,
            token_uri TEXT NOT NULL,
            client_id TEXT NOT NULL,
            scopes TEXT,
            expires_at TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP,

            FOREIGN KEY (user_id)
                REFERENCES users (id)
                ON DELETE CASCADE,

            FOREIGN KEY (email_account_id)
                REFERENCES email_accounts (id)
                ON DELETE CASCADE,

            UNIQUE (user_id, email_account_id)
        )
        """)

    # Microsoft / Outlook OAuthトークンキャッシュ
    connection.execute("""
        CREATE TABLE IF NOT EXISTS microsoft_oauth_tokens (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            email_account_id INTEGER NOT NULL,
            token_cache TEXT NOT NULL,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP,

            FOREIGN KEY (user_id)
                REFERENCES users (id)
                ON DELETE CASCADE,

            FOREIGN KEY (email_account_id)
                REFERENCES email_accounts (id)
                ON DELETE CASCADE,

            UNIQUE (user_id, email_account_id)
        )
        """)

    # Ver2.09-1: メール由来イベントを識別する列
    company_event_columns = connection.execute("""
        PRAGMA table_info(company_events)
        """).fetchall()

    company_event_column_names = {column["name"] for column in company_event_columns}

    company_event_new_columns = {
        "source_email_provider": "TEXT",
        "source_email_account_id": "INTEGER",
        "source_email_message_id": "TEXT",
        "source_email_event_key": "TEXT",
    }

    for column_name, column_type in company_event_new_columns.items():
        if column_name not in company_event_column_names:
            connection.execute(f"""
                ALTER TABLE company_events
                ADD COLUMN {column_name} {column_type}
                """)

    # 旧版は「1メール=1イベント」の一意制約だったため、
    # 複数イベント対応用のevent_keyへ移行する。
    connection.execute(
        """
        DROP INDEX IF EXISTS idx_company_events_email_source
        """
    )

    connection.execute(
        """
        UPDATE company_events
        SET source_email_event_key = CASE
            WHEN event_type = '面接' THEN 'interview'
            WHEN event_type = '適性検査' THEN 'test'
            WHEN event_type = 'ES' THEN 'es'
            WHEN event_type = '説明会' THEN 'briefing'
            WHEN event_type = 'インターン' THEN 'intern'
            WHEN event_type = '結果通知' THEN 'result'
            WHEN event_type = '締切' THEN 'deadline'
            ELSE 'primary'
        END
        WHERE source_email_message_id IS NOT NULL
          AND (
                source_email_event_key IS NULL
                OR source_email_event_key = ''
              )
        """
    )

    connection.execute(
        """
        CREATE UNIQUE INDEX IF NOT EXISTS
            idx_company_events_email_source
        ON company_events (
            source_email_provider,
            source_email_account_id,
            source_email_message_id,
            source_email_event_key
        )
        WHERE source_email_message_id IS NOT NULL
        """
    )

    # Gmail・Outlook共通のメール保存テーブル
    connection.execute("""
        CREATE TABLE IF NOT EXISTS email_messages (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            email_account_id INTEGER NOT NULL,
            provider TEXT NOT NULL,
            message_id TEXT NOT NULL,
            thread_id TEXT,
            company_id INTEGER,
            subject TEXT,
            sender TEXT,
            received_at TEXT,
            snippet TEXT,
            body TEXT,
            is_unread INTEGER NOT NULL DEFAULT 0,
            match_method TEXT,
            match_confidence INTEGER,
            matched_at TEXT,
            fetched_at TEXT DEFAULT CURRENT_TIMESTAMP,

            FOREIGN KEY (user_id)
                REFERENCES users (id)
                ON DELETE CASCADE,

            FOREIGN KEY (email_account_id)
                REFERENCES email_accounts (id)
                ON DELETE CASCADE,

            FOREIGN KEY (company_id)
                REFERENCES companies (id)
                ON DELETE SET NULL,

            UNIQUE (
                email_account_id,
                provider,
                message_id
            )
        )
        """)

    # 古いデータベースにmemo列がない場合だけ追加
    columns = connection.execute("""
        PRAGMA table_info(companies)
        """).fetchall()

    column_names = [column["name"] for column in columns]

    if "memo" not in column_names:
        connection.execute("""
            ALTER TABLE companies
            ADD COLUMN memo TEXT
            """)

    email_columns = connection.execute("""
        PRAGMA table_info(email_accounts)
        """).fetchall()

    email_column_names = [column["name"] for column in email_columns]

    if "user_id" not in email_column_names:
        connection.execute("""
            ALTER TABLE email_accounts
            ADD COLUMN user_id INTEGER
            """)

    # Ver2-01: 企業をユーザーごとに分離する
    if "user_id" not in column_names:
        connection.execute("""
            ALTER TABLE companies
            ADD COLUMN user_id INTEGER
            """)

    # 既存企業は最初に登録されたユーザーへ割り当てる
    first_user = connection.execute("""
        SELECT id
        FROM users
        ORDER BY id ASC
        LIMIT 1
        """).fetchone()

    if first_user is not None:
        connection.execute(
            """
            UPDATE companies
            SET user_id = ?
            WHERE user_id IS NULL
            """,
            (first_user["id"],),
        )

        connection.execute(
            """
            UPDATE email_accounts
            SET user_id = ?
            WHERE user_id IS NULL
            """,
            (first_user["id"],),
        )

    connection.commit()
    connection.close()


create_tables()


# =========================================================
# デモ専用：メールに記載された企業名での自動照合（外部通信なし）
# =========================================================
def relink_demo_mail_data(connection, user_id):
    """本来の企業名判定ロジックを、架空メールのみへ安全に適用する。"""
    rows = connection.execute(
        """SELECT id, subject, sender, snippet FROM email_messages
           WHERE user_id = ? ORDER BY id""",
        (user_id,),
    ).fetchall()
    matched = 0
    for message in rows:
        company_id, method, confidence = find_matching_company(
            connection, user_id, message["subject"], message["sender"], message["snippet"]
        )
        connection.execute(
            """UPDATE email_messages
               SET company_id = ?, match_method = ?, match_confidence = ?,
                   matched_at = CASE WHEN ? IS NULL THEN NULL ELSE CURRENT_TIMESTAMP END
               WHERE id = ? AND user_id = ?""",
            (company_id, method, confidence, company_id, message["id"], user_id),
        )
        matched += int(company_id is not None)
    return matched


# =========================================================
# 公開ポートフォリオ専用：外部連携をせず架空メールを再現する
# =========================================================
def seed_demo_mail_data(connection):
    """デモ用の架空アカウント/メールを登録し、実際の照合処理を適用する。

    実ユーザーのメール、OAuth資格情報、外部APIには一切アクセスしない。
    既存レコードはUNIQUEキーにより二重登録しない。
    """
    dummy_accounts = (
        ("gmail", "sample-gmail@example.invalid", "Gmail（サンプル）"),
        ("outlook", "sample-outlook@example.invalid", "Outlook（サンプル）"),
    )
    account_ids = {}
    for provider, address, label in dummy_accounts:
        connection.execute(
            """INSERT OR IGNORE INTO email_accounts
               (user_id,provider,email_address,display_name,auth_type,is_active)
               VALUES(1,?,?,?,'oauth',1)""",
            (provider,address,label),
        )
        account_ids[provider] = connection.execute(
            "SELECT id FROM email_accounts WHERE user_id=1 AND email_address=?",
            (address,),
        ).fetchone()["id"]

    # 架空の企業名はデモの企業マスタに対応。一部は意図的に未紐付けにする。
    # 後から新しい企業を登録して「企業を再照合」するとマッチングを再体験できる。
    samples = [
        ("gmail", "sample-g-1", "【青空テクノロジーズ】一次面接のご案内", "採用担当 <recruit@aozora.example.invalid>",
         "青空テクノロジーズの一次面接は2026年11月12日 14:00です。", "2026-10-18T13:10:00"),
        ("gmail", "sample-g-2", "未来データ研究所 エントリーシート提出完了", "採用窓口 <hr@mirai.example.invalid>",
         "未来データ研究所へのES提出を受け付けました。", "2026-10-18T11:40:00"),
        ("gmail", "sample-g-3", "【星空システムズ】インターン参加のご案内", "インターン担当 <intern@hoshizora.example.invalid>",
         "星空システムズのインターンは2026年12月3日に開催されます。", "2026-10-17T19:30:00"),
        ("gmail", "sample-g-4", "【北斗ソリューションズ】会社説明会のご案内", "採用窓口 <recruit@hokuto.example.invalid>",
         "北斗ソリューションズの説明会は2026年11月18日に開催予定です。未登録企業のサンプルです。", "2026-10-16T15:20:00"),
        ("outlook", "sample-o-1", "さくらクラウド 適性検査のご案内", "採用チーム <career@sakura.example.invalid>",
         "さくらクラウドの適性検査を2026年11月10日までに受検してください。", "2026-10-18T09:20:00"),
        ("outlook", "sample-o-2", "【未来データ研究所】選考通過のお知らせ", "採用窓口 <hr@mirai.example.invalid>",
         "未来データ研究所の選考を通過しました。次回の日程は別途ご案内します。", "2026-10-17T17:10:00"),
        ("outlook", "sample-o-3", "大学キャリアセンターからのお知らせ", "キャリアセンター <career@university.example.invalid>",
         "合同説明会の情報をお送りします。企業名は未登録のため紐付けされません。", "2026-10-15T12:00:00"),
    ]
    for provider, message_id, subject, sender, snippet, received_at in samples:
        connection.execute(
            """INSERT OR IGNORE INTO email_messages
               (user_id,email_account_id,provider,message_id,subject,sender,
                received_at,snippet,body,is_unread)
               VALUES(1,?,?,?,?,?,?,?,?,?)""",
            (account_ids[provider],provider,message_id,subject,sender,
             received_at,snippet,snippet,1 if message_id.endswith('1') else 0),
        )

    # 本体と同じ企業名判定ロジックを使う。再実行しても重複しない。
    return relink_demo_mail_data(connection, 1)


# =========================================================
# ポートフォリオデモ：訪問者ごとの仮データを用意
# =========================================================
if DEMO_MODE:
    # 既存DBからのコピーではなく、公開用に空から生成したDBだけを利用する。
    with sqlite3.connect(DATABASE) as demo_connection:
        demo_connection.execute(
            "INSERT OR IGNORE INTO users (id,name,email,password_hash) VALUES (1,?,?,?)",
            ("デモユーザー", "demo@example.invalid", generate_password_hash(secrets.token_urlsafe(32))),
        )
        if not demo_connection.execute("SELECT 1 FROM companies LIMIT 1").fetchone():
            sample = [
                ("青空テクノロジーズ", "一次面接", "2026-11-12", "2026-11-05", "Webサービス開発職。架空の企業です。"),
                ("未来データ研究所", "ES提出済", "2026-11-08", "2026-10-31", "データ分析職。架空の企業です。"),
                ("さくらクラウド", "適性検査", "2026-11-15", "2026-11-10", "クラウドエンジニア職。架空の企業です。"),
                ("星空システムズ", "インターン", "2026-12-03", "2026-11-20", "ソフトウェアエンジニア職。架空の企業です。"),
            ]
            for name,status,next_date,deadline,memo in sample:
                cursor = demo_connection.execute(
                    "INSERT INTO companies(user_id,name,status,next_schedule,deadline,mypage_url,memo) VALUES(1,?,?,?,?,?,?)",
                    (name,status,next_date,deadline,"",memo),
                )
                demo_connection.execute(
                    "INSERT INTO company_events(company_id,event_date,event_type,title,memo) VALUES(?,?,?,?,?)",
                    (cursor.lastrowid, next_date, "面接" if status=="一次面接" else "その他", status, "デモ用の架空イベント"),
                )
            demo_connection.commit()

    # 起動時にデモ用アカウントとサンプル受信メールを生成。
    with sqlite3.connect(DATABASE) as demo_mail_connection:
        demo_mail_connection.row_factory = sqlite3.Row
        demo_mail_connection.execute("PRAGMA foreign_keys = ON")
        seed_demo_mail_data(demo_mail_connection)
        demo_mail_connection.commit()


@app.before_request
def prepare_portfolio_demo():
    if not DEMO_MODE:
        return None
    identifier = session.get("demo_id", "")
    if not isinstance(identifier, str) or not re.fullmatch(r"[0-9a-f]{32}", identifier):
        identifier = secrets.token_hex(16)
        session.clear()
        session["demo_id"] = identifier
    demo_path = DEMO_HOME / f"visit-{identifier}.db"
    if not demo_path.exists():
        shutil.copyfile(DATABASE, demo_path)
    session["user_id"] = 1

    # 旧いデモDBを削除。永続保存は行わない（Render無料版でも利用可能）。
    if secrets.randbelow(100) == 0:
        cutoff = time.time() - 86400
        for old in DEMO_HOME.glob("visit-*.db"):
            try:
                if old.stat().st_mtime < cutoff:
                    old.unlink()
            except OSError:
                pass

    # デモでは外部メール接続・会員登録・パスワード変更を公開しない。
    blocked = {"register", "login", "logout", "account",
               "google_connect", "google_oauth_callback", "outlook_connect",
               "outlook_callback", "sync_outlook_messages", "sync_gmail_messages",
               "toggle_email_account", "delete_email_account"}
    if request.endpoint in blocked:
        if request.method != "GET":
            abort(403)
        return redirect(url_for("index"))


@app.context_processor
def expose_portfolio_demo_flag():
    return {"demo_mode": DEMO_MODE}


@app.after_request
def portfolio_demo_security_headers(response):
    if DEMO_MODE:
        response.headers["X-Content-Type-Options"] = "nosniff"
        response.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"
        response.headers["X-Frame-Options"] = "DENY"
        response.headers["Cache-Control"] = "private, no-store"
    return response


# =========================================================
# ログイン状態の読み込み
# =========================================================


@app.before_request
def load_logged_in_user():
    user_id = session.get("user_id")

    if user_id is None:
        g.user = None
        return

    connection = get_db_connection()

    g.user = connection.execute(
        """
        SELECT id, name, email, created_at
        FROM users
        WHERE id = ?
        """,
        (user_id,),
    ).fetchone()

    connection.close()

    if g.user is None:
        session.clear()


# =========================================================
# 未ログイン時のアクセス制限
# =========================================================


@app.before_request
def require_login():
    allowed_endpoints = {
        "login",
        "register",
        "static",
    }

    if request.endpoint in allowed_endpoints:
        return

    if g.user is None:
        return redirect(
            url_for(
                "login",
                next=request.full_path,
            )
        )


# =========================================================
# 新規登録
# =========================================================


@app.route(
    "/register",
    methods=["GET", "POST"],
)
def register():
    if g.user is not None:
        return redirect(url_for("index"))

    error = None

    if request.method == "POST":
        name = request.form.get(
            "name",
            "",
        ).strip()

        email = (
            request.form.get(
                "email",
                "",
            )
            .strip()
            .lower()
        )

        password = request.form.get(
            "password",
            "",
        )

        password_confirm = request.form.get(
            "password_confirm",
            "",
        )

        if not name:
            error = "名前を入力してください。"

        elif not email:
            error = "メールアドレスを入力してください。"

        elif "@" not in email:
            error = "正しいメールアドレスを入力してください。"

        elif len(password) < 8:
            error = "パスワードは8文字以上にしてください。"

        elif password != password_confirm:
            error = "確認用パスワードが一致しません。"

        else:
            connection = get_db_connection()

            existing_user = connection.execute(
                """
                SELECT id
                FROM users
                WHERE lower(email) = lower(?)
                """,
                (email,),
            ).fetchone()

            if existing_user:
                error = "このメールアドレスはすでに" "登録されています。"
                connection.close()

            else:
                cursor = connection.execute(
                    """
                    INSERT INTO users (
                        name,
                        email,
                        password_hash
                    )
                    VALUES (?, ?, ?)
                    """,
                    (
                        name,
                        email,
                        generate_password_hash(password),
                    ),
                )

                connection.commit()

                user_id = cursor.lastrowid
                connection.close()

                session.clear()
                session["user_id"] = user_id

                return redirect(url_for("index"))

    return render_template(
        "register.html",
        error=error,
    )


# =========================================================
# ログイン
# =========================================================


@app.route(
    "/login",
    methods=["GET", "POST"],
)
def login():
    if g.user is not None:
        return redirect(url_for("index"))

    error = None

    if request.method == "POST":
        email = (
            request.form.get(
                "email",
                "",
            )
            .strip()
            .lower()
        )

        password = request.form.get(
            "password",
            "",
        )

        connection = get_db_connection()

        user = connection.execute(
            """
            SELECT *
            FROM users
            WHERE lower(email) = lower(?)
            """,
            (email,),
        ).fetchone()

        connection.close()

        if user is None:
            error = "メールアドレスまたはパスワードが" "正しくありません。"

        elif not check_password_hash(
            user["password_hash"],
            password,
        ):
            error = "メールアドレスまたはパスワードが" "正しくありません。"

        else:
            session.clear()
            session["user_id"] = user["id"]

            next_url = request.args.get("next")

            if next_url and next_url.startswith("/") and not next_url.startswith("//"):
                return redirect(next_url)

            return redirect(url_for("index"))

    return render_template(
        "login.html",
        error=error,
    )


# =========================================================
# アカウント
# =========================================================


@app.route(
    "/account",
    methods=["GET", "POST"],
)
def account():
    error = None
    success = None

    if request.method == "POST":
        action = request.form.get(
            "action",
            "",
        )

        connection = get_db_connection()

        current_user = connection.execute(
            """
            SELECT *
            FROM users
            WHERE id = ?
            """,
            (g.user["id"],),
        ).fetchone()

        if current_user is None:
            connection.close()
            session.clear()

            return redirect(url_for("login"))

        # ---------------------------------
        # 名前の変更
        # ---------------------------------

        if action == "update_name":
            new_name = request.form.get(
                "name",
                "",
            ).strip()

            if not new_name:
                error = "名前を入力してください。"

            elif len(new_name) > 100:
                error = "名前は100文字以内で入力してください。"

            else:
                connection.execute(
                    """
                    UPDATE users
                    SET name = ?
                    WHERE id = ?
                    """,
                    (
                        new_name,
                        g.user["id"],
                    ),
                )

                connection.commit()

                success = "名前を変更しました。"

        # ---------------------------------
        # メールアドレスの変更
        # ---------------------------------

        elif action == "update_email":
            new_email = (
                request.form.get(
                    "email",
                    "",
                )
                .strip()
                .lower()
            )

            current_password = request.form.get(
                "current_password",
                "",
            )

            if not new_email:
                error = "新しいメールアドレスを" "入力してください。"

            elif "@" not in new_email:
                error = "正しいメールアドレスを" "入力してください。"

            elif not current_password:
                error = "現在のパスワードを" "入力してください。"

            elif not check_password_hash(
                current_user["password_hash"],
                current_password,
            ):
                error = "現在のパスワードが" "正しくありません。"

            else:
                existing_user = connection.execute(
                    """
                    SELECT id
                    FROM users
                    WHERE lower(email) = lower(?)
                      AND id != ?
                    """,
                    (
                        new_email,
                        g.user["id"],
                    ),
                ).fetchone()

                if existing_user:
                    error = "このメールアドレスは" "すでに使用されています。"

                else:
                    connection.execute(
                        """
                        UPDATE users
                        SET email = ?
                        WHERE id = ?
                        """,
                        (
                            new_email,
                            g.user["id"],
                        ),
                    )

                    connection.commit()

                    success = "メールアドレスを変更しました。"

        # ---------------------------------
        # パスワードの変更
        # ---------------------------------

        elif action == "update_password":
            current_password = request.form.get(
                "current_password",
                "",
            )

            new_password = request.form.get(
                "new_password",
                "",
            )

            new_password_confirm = request.form.get(
                "new_password_confirm",
                "",
            )

            if not current_password:
                error = "現在のパスワードを" "入力してください。"

            elif not check_password_hash(
                current_user["password_hash"],
                current_password,
            ):
                error = "現在のパスワードが" "正しくありません。"

            elif len(new_password) < 8:
                error = "新しいパスワードは" "8文字以上にしてください。"

            elif new_password != new_password_confirm:
                error = "新しいパスワードと" "確認用パスワードが一致しません。"

            elif check_password_hash(
                current_user["password_hash"],
                new_password,
            ):
                error = "現在と異なるパスワードを" "設定してください。"

            else:
                new_password_hash = generate_password_hash(new_password)

                connection.execute(
                    """
                    UPDATE users
                    SET password_hash = ?
                    WHERE id = ?
                    """,
                    (
                        new_password_hash,
                        g.user["id"],
                    ),
                )

                connection.commit()

                success = "パスワードを変更しました。"

        else:
            error = "正しい操作を選択してください。"

        connection.close()

    connection = get_db_connection()

    user = connection.execute(
        """
        SELECT
            id,
            name,
            email,
            created_at
        FROM users
        WHERE id = ?
        """,
        (g.user["id"],),
    ).fetchone()

    connection.close()

    if user is None:
        session.clear()

        return redirect(url_for("login"))

    return render_template(
        "account.html",
        user=user,
        error=error,
        success=success,
    )


# =========================================================
# ログアウト
# =========================================================


@app.route(
    "/logout",
    methods=["POST"],
)
def logout():
    session.clear()

    return redirect(url_for("login"))


# =========================================================
# ホーム画面
# =========================================================


@app.route("/")
def index():
    connection = get_db_connection()

    today = date.today()
    seven_days_later = today + timedelta(days=7)
    user_id = g.user["id"]

    # -------------------------
    # 登録企業数
    # -------------------------

    company_count = connection.execute(
        """
        SELECT COUNT(*) AS count
        FROM companies
        WHERE user_id = ?
        """,
        (user_id,),
    ).fetchone()["count"]

    # -------------------------
    # 選考中の企業数
    # -------------------------

    selection_statuses = (
        "ES作成中",
        "ES提出済",
        "適性検査",
        "面接",
        "一次面接",
        "二次面接",
        "三次面接",
        "最終面接",
    )

    placeholders = ",".join("?" for _ in selection_statuses)

    selection_count = connection.execute(
        f"""
        SELECT COUNT(*) AS count
        FROM companies
        WHERE user_id = ?
          AND status IN ({placeholders})
        """,
        (
            user_id,
            *selection_statuses,
        ),
    ).fetchone()["count"]

    # -------------------------
    # 7日以内の締切数
    # -------------------------

    deadline_count = connection.execute(
        """
        SELECT COUNT(*) AS count
        FROM companies
        WHERE user_id = ?
          AND deadline IS NOT NULL
          AND deadline != ''
          AND date(deadline)
              BETWEEN date(?) AND date(?)
        """,
        (
            user_id,
            today.isoformat(),
            seven_days_later.isoformat(),
        ),
    ).fetchone()["count"]

    # -------------------------
    # 結果待ちの企業数
    # -------------------------

    waiting_count = connection.execute(
        """
        SELECT COUNT(*) AS count
        FROM companies
        WHERE user_id = ?
          AND status = ?
        """,
        (
            user_id,
            "結果待ち",
        ),
    ).fetchone()["count"]

    # -------------------------
    # 今日の予定
    # -------------------------

    today_events = connection.execute(
        """
        SELECT
            company_events.*,
            companies.name AS company_name
        FROM company_events
        JOIN companies
          ON companies.id = company_events.company_id
        WHERE companies.user_id = ?
          AND date(company_events.event_date) = date(?)
        ORDER BY company_events.event_date ASC
        """,
        (
            user_id,
            today.isoformat(),
        ),
    ).fetchall()

    # -------------------------
    # 明日から7日以内の予定
    # -------------------------

    upcoming_events = connection.execute(
        """
        SELECT
            company_events.*,
            companies.name AS company_name
        FROM company_events
        JOIN companies
          ON companies.id = company_events.company_id
        WHERE companies.user_id = ?
          AND date(company_events.event_date)
              BETWEEN date(?, '+1 day')
              AND date(?)
        ORDER BY company_events.event_date ASC
        """,
        (
            user_id,
            today.isoformat(),
            seven_days_later.isoformat(),
        ),
    ).fetchall()

    # -------------------------
    # 7日以内の締切
    # -------------------------

    upcoming_deadlines = connection.execute(
        """
        SELECT *
        FROM companies
        WHERE user_id = ?
          AND deadline IS NOT NULL
          AND deadline != ''
          AND date(deadline)
              BETWEEN date(?) AND date(?)
        ORDER BY deadline ASC
        """,
        (
            user_id,
            today.isoformat(),
            seven_days_later.isoformat(),
        ),
    ).fetchall()

    connection.close()

    return render_template(
        "index.html",
        company_count=company_count,
        selection_count=selection_count,
        deadline_count=deadline_count,
        waiting_count=waiting_count,
        today_events=today_events,
        upcoming_events=upcoming_events,
        upcoming_deadlines=upcoming_deadlines,
    )


# =========================================================
# カレンダー
# =========================================================


@app.route("/calendar")
def calendar_view():
    today = date.today()

    year_text = request.args.get(
        "year",
        str(today.year),
    )

    month_text = request.args.get(
        "month",
        str(today.month),
    )

    view_mode = request.args.get(
        "view",
        "month",
    ).strip()

    search_query = request.args.get(
        "q",
        "",
    ).strip()

    selected_type = request.args.get(
        "type",
        "all",
    ).strip()

    if view_mode not in {
        "month",
        "list",
    }:
        view_mode = "month"

    try:
        selected_year = int(year_text)
        selected_month = int(month_text)

    except (TypeError, ValueError):
        selected_year = today.year
        selected_month = today.month

    if selected_year < 2000 or selected_year > 2100:
        selected_year = today.year

    if selected_month < 1 or selected_month > 12:
        selected_month = today.month

    first_day = date(
        selected_year,
        selected_month,
        1,
    )

    days_in_month = monthrange(
        selected_year,
        selected_month,
    )[1]

    last_day = date(
        selected_year,
        selected_month,
        days_in_month,
    )

    calendar_start = first_day - timedelta(days=first_day.weekday())

    calendar_end = calendar_start + timedelta(days=41)

    connection = get_db_connection()
    user_id = g.user["id"]

    event_rows = connection.execute(
        """
        SELECT
            company_events.id,
            company_events.company_id,
            company_events.event_date,
            company_events.event_type,
            company_events.title,
            company_events.memo,
            companies.name AS company_name
        FROM company_events
        JOIN companies
          ON companies.id = company_events.company_id
        WHERE companies.user_id = ?
          AND company_events.event_date IS NOT NULL
          AND company_events.event_date != ''
          AND date(company_events.event_date)
              BETWEEN date(?) AND date(?)
        ORDER BY
            company_events.event_date ASC,
            company_events.id ASC
        """,
        (
            user_id,
            calendar_start.isoformat(),
            calendar_end.isoformat(),
        ),
    ).fetchall()

    deadline_rows = connection.execute(
        """
        SELECT
            id,
            name,
            deadline,
            status
        FROM companies
        WHERE user_id = ?
          AND deadline IS NOT NULL
          AND deadline != ''
          AND date(deadline)
              BETWEEN date(?) AND date(?)
        ORDER BY
            deadline ASC,
            id ASC
        """,
        (
            user_id,
            calendar_start.isoformat(),
            calendar_end.isoformat(),
        ),
    ).fetchall()

    company_rows = connection.execute(
        """
        SELECT
            id,
            name
        FROM companies
        WHERE user_id = ?
        ORDER BY name COLLATE NOCASE ASC
        """,
        (user_id,),
    ).fetchall()

    connection.close()

    items_by_date = {}

    for event in event_rows:
        event_date = event["event_date"][:10]

        items_by_date.setdefault(
            event_date,
            [],
        )

        items_by_date[event_date].append(
            {
                "kind": "event",
                "event_id": event["id"],
                "company_id": event["company_id"],
                "company_name": event["company_name"],
                "title": (event["title"] or event["event_type"] or "予定"),
                "event_type": (event["event_type"] or "その他"),
                "memo": event["memo"] or "",
                "date": event_date,
            }
        )

    for company in deadline_rows:
        deadline_date = company["deadline"][:10]

        items_by_date.setdefault(
            deadline_date,
            [],
        )

        items_by_date[deadline_date].append(
            {
                "kind": "deadline",
                "event_id": None,
                "company_id": company["id"],
                "company_name": company["name"],
                "title": "応募・提出締切",
                "event_type": "締切",
                "memo": "",
                "date": deadline_date,
            }
        )

    calendar_days = []
    current_day = calendar_start

    while current_day <= calendar_end:
        current_date_text = current_day.isoformat()

        holiday_name = jpholiday.is_holiday_name(current_day)

        calendar_days.append(
            {
                "date": current_day,
                "date_text": current_date_text,
                "day": current_day.day,
                "weekday": current_day.weekday(),
                "is_saturday": (current_day.weekday() == 5),
                "is_sunday": (current_day.weekday() == 6),
                "holiday_name": holiday_name,
                "is_holiday": bool(holiday_name),
                "is_current_month": (
                    current_day.year == selected_year
                    and current_day.month == selected_month
                ),
                "is_today": current_day == today,
                "items": items_by_date.get(
                    current_date_text,
                    [],
                ),
            }
        )

        current_day += timedelta(days=1)

    if selected_month == 1:
        previous_year = selected_year - 1
        previous_month = 12

    else:
        previous_year = selected_year
        previous_month = selected_month - 1

    if selected_month == 12:
        next_year = selected_year + 1
        next_month = 1

    else:
        next_year = selected_year
        next_month = selected_month + 1

    monthly_items = []

    for item_date, items in items_by_date.items():
        if first_day.isoformat() <= item_date <= last_day.isoformat():
            monthly_items.extend(items)

    monthly_items.sort(
        key=lambda item: (
            item["date"],
            0 if item["kind"] == "deadline" else 1,
            item["company_name"],
        )
    )

    today_count = len(
        items_by_date.get(
            today.isoformat(),
            [],
        )
    )

    deadline_count = sum(1 for item in monthly_items if item["kind"] == "deadline")

    interview_count = sum(
        1
        for item in monthly_items
        if (item["kind"] == "event" and "面接" in item["event_type"])
    )

    type_options = (
        "締切",
        "説明会",
        "ES",
        "適性検査",
        "面接",
        "面談",
        "インターン",
        "結果発表",
        "その他",
    )

    if selected_type != "all" and selected_type not in type_options:
        selected_type = "all"

    normalized_query = search_query.casefold()
    filtered_items = []

    for item in monthly_items:
        if selected_type != "all":
            if selected_type == "締切" and item["kind"] != "deadline":
                continue

            if selected_type != "締切" and item["event_type"] != selected_type:
                continue

        searchable_text = " ".join(
            (
                item["company_name"],
                item["title"],
                item["event_type"],
                item["memo"],
            )
        ).casefold()

        if normalized_query and normalized_query not in searchable_text:
            continue

        filtered_items.append(item)

    tomorrow = today + timedelta(days=1)
    week_end = today + timedelta(days=(6 - today.weekday()))

    grouped_items = {
        "今日": [],
        "明日": [],
        "今週": [],
        "今月": [],
    }

    for item in filtered_items:
        try:
            item_date = date.fromisoformat(item["date"])

        except ValueError:
            grouped_items["今月"].append(item)
            continue

        if item_date == today:
            group_title = "今日"

        elif item_date == tomorrow:
            group_title = "明日"

        elif today < item_date <= week_end and item_date != tomorrow:
            group_title = "今週"

        else:
            group_title = "今月"

        grouped_items[group_title].append(item)

    list_groups = [
        {
            "title": group_title,
            "items": grouped_items[group_title],
        }
        for group_title in (
            "今日",
            "明日",
            "今週",
            "今月",
        )
        if grouped_items[group_title]
    ]

    return render_template(
        "calendar.html",
        selected_year=selected_year,
        selected_month=selected_month,
        view_mode=view_mode,
        search_query=search_query,
        selected_type=selected_type,
        type_options=type_options,
        calendar_days=calendar_days,
        monthly_items=monthly_items,
        filtered_items=filtered_items,
        list_groups=list_groups,
        companies=company_rows,
        previous_year=previous_year,
        previous_month=previous_month,
        next_year=next_year,
        next_month=next_month,
        today=today,
        today_count=today_count,
        deadline_count=deadline_count,
        interview_count=interview_count,
    )


# =========================================================
# カレンダーから予定を追加・編集
# =========================================================


@app.route(
    "/calendar/event/save",
    methods=["POST"],
)
def save_calendar_event():
    event_id = request.form.get(
        "event_id",
        "",
    ).strip()

    company_id = request.form.get(
        "company_id",
        "",
    ).strip()

    event_date = request.form.get(
        "event_date",
        "",
    ).strip()

    event_type = request.form.get(
        "event_type",
        "",
    ).strip()

    title = request.form.get(
        "title",
        "",
    ).strip()

    memo = request.form.get(
        "memo",
        "",
    ).strip()

    selected_year = request.form.get(
        "selected_year",
        str(date.today().year),
    ).strip()

    selected_month = request.form.get(
        "selected_month",
        str(date.today().month),
    ).strip()

    view_mode = request.form.get(
        "view_mode",
        "month",
    ).strip()

    if view_mode not in {
        "month",
        "list",
    }:
        view_mode = "month"

    redirect_arguments = {
        "year": selected_year,
        "month": selected_month,
        "view": view_mode,
    }

    if not company_id or not event_date:
        return redirect(
            url_for(
                "calendar_view",
                **redirect_arguments,
            )
        )

    if not event_type:
        event_type = "その他"

    if not title:
        title = event_type

    try:
        saved_date = date.fromisoformat(event_date)
        company_id_number = int(company_id)

    except (TypeError, ValueError):
        return redirect(
            url_for(
                "calendar_view",
                **redirect_arguments,
            )
        )

    connection = get_db_connection()

    # 選択された企業がログイン中ユーザーのものか確認
    company = connection.execute(
        """
        SELECT id
        FROM companies
        WHERE id = ?
          AND user_id = ?
        """,
        (
            company_id_number,
            g.user["id"],
        ),
    ).fetchone()

    if company is None:
        connection.close()

        return redirect(
            url_for(
                "calendar_view",
                **redirect_arguments,
            )
        )

    previous_company_id = None

    if event_id:
        try:
            event_id_number = int(event_id)

        except ValueError:
            connection.close()

            return redirect(
                url_for(
                    "calendar_view",
                    **redirect_arguments,
                )
            )

        # 編集対象がログイン中ユーザーの予定か確認
        existing_event = connection.execute(
            """
            SELECT
                company_events.company_id
            FROM company_events
            JOIN companies
              ON companies.id = company_events.company_id
            WHERE company_events.id = ?
              AND companies.user_id = ?
            """,
            (
                event_id_number,
                g.user["id"],
            ),
        ).fetchone()

        if existing_event is None:
            connection.close()

            return redirect(
                url_for(
                    "calendar_view",
                    **redirect_arguments,
                )
            )

        previous_company_id = existing_event["company_id"]

        connection.execute(
            """
            UPDATE company_events
            SET
                company_id = ?,
                event_date = ?,
                event_type = ?,
                title = ?,
                memo = ?
            WHERE id = ?
            """,
            (
                company_id_number,
                event_date,
                event_type,
                title,
                memo,
                event_id_number,
            ),
        )

    else:
        connection.execute(
            """
            INSERT INTO company_events (
                company_id,
                event_date,
                event_type,
                title,
                memo
            )
            VALUES (?, ?, ?, ?, ?)
            """,
            (
                company_id_number,
                event_date,
                event_type,
                title,
                memo,
            ),
        )

    update_company_status_from_events(
        connection,
        company_id_number,
    )

    if previous_company_id is not None and previous_company_id != company_id_number:
        update_company_status_from_events(
            connection,
            previous_company_id,
        )

    connection.commit()
    connection.close()

    return redirect(
        url_for(
            "calendar_view",
            year=saved_date.year,
            month=saved_date.month,
            view=view_mode,
        )
    )


# =========================================================
# カレンダーから予定を削除
# =========================================================


@app.route(
    "/calendar/event/<int:event_id>/delete",
    methods=["POST"],
)
def delete_calendar_event(event_id):
    today = date.today()

    selected_year = request.form.get(
        "selected_year",
        str(today.year),
    ).strip()

    selected_month = request.form.get(
        "selected_month",
        str(today.month),
    ).strip()

    view_mode = request.form.get(
        "view_mode",
        "month",
    ).strip()

    if view_mode not in {
        "month",
        "list",
    }:
        view_mode = "month"

    connection = get_db_connection()

    # ログイン中のユーザーが所有する予定か確認
    event = connection.execute(
        """
        SELECT
            company_events.company_id
        FROM company_events
        JOIN companies
          ON companies.id = company_events.company_id
        WHERE company_events.id = ?
          AND companies.user_id = ?
        """,
        (
            event_id,
            g.user["id"],
        ),
    ).fetchone()

    if event is not None:
        connection.execute(
            """
            DELETE FROM company_events
            WHERE id = ?
            """,
            (event_id,),
        )

        update_company_status_from_events(
            connection,
            event["company_id"],
        )

        connection.commit()

    connection.close()

    return redirect(
        url_for(
            "calendar_view",
            year=selected_year,
            month=selected_month,
            view=view_mode,
        )
    )


# =========================================================
# 企業一覧
# =========================================================


@app.route("/companies")
def companies():
    selected_filter = request.args.get(
        "filter",
        "all",
    ).strip()

    search_query = request.args.get(
        "q",
        "",
    ).strip()

    selected_sort = request.args.get(
        "sort",
        "newest",
    ).strip()

    connection = get_db_connection()

    filter_title = "すべての企業"
    filter_description = "登録されているすべての企業を表示しています。"
    conditions = ["user_id = ?"]
    parameters = [g.user["id"]]

    if selected_filter == "selection":
        selection_statuses = (
            "ES作成中",
            "ES提出済",
            "適性検査",
            "面接",
            "一次面接",
            "二次面接",
            "三次面接",
            "最終面接",
        )

        placeholders = ",".join("?" for _ in selection_statuses)

        conditions.append(f"status IN ({placeholders})")
        parameters.extend(selection_statuses)

        filter_title = "選考中の企業"
        filter_description = "現在、選考が進んでいる企業を表示しています。"

    elif selected_filter == "deadline":
        today = date.today()
        seven_days_later = today + timedelta(days=7)

        conditions.extend(
            (
                "deadline IS NOT NULL",
                "deadline != ''",
                "date(deadline) BETWEEN date(?) AND date(?)",
            )
        )
        parameters.extend(
            (
                today.isoformat(),
                seven_days_later.isoformat(),
            )
        )

        filter_title = "7日以内に締切がある企業"
        filter_description = "今日から7日以内に締切を迎える企業を" "表示しています。"

    elif selected_filter == "waiting":
        conditions.append("status = ?")
        parameters.append("結果待ち")

        filter_title = "結果待ちの企業"
        filter_description = "現在、選考結果を待っている企業を" "表示しています。"

    else:
        selected_filter = "all"

    if search_query:
        escaped_query = (
            search_query.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")
        )
        conditions.append("name LIKE ? ESCAPE '\\'")
        parameters.append(f"%{escaped_query}%")

    sort_options = {
        "newest": "id DESC",
        "deadline": (
            "CASE "
            "WHEN deadline IS NULL OR deadline = '' "
            "THEN 1 ELSE 0 END, "
            "deadline ASC, id DESC"
        ),
        "name": "name COLLATE NOCASE ASC, id DESC",
        "status": "status COLLATE NOCASE ASC, id DESC",
    }

    if selected_sort not in sort_options:
        selected_sort = "newest"

    where_clause = ""

    if conditions:
        where_clause = "WHERE " + " AND ".join(conditions)

    company_list = connection.execute(
        f"""
        SELECT *
        FROM companies
        {where_clause}
        ORDER BY {sort_options[selected_sort]}
        """,
        parameters,
    ).fetchall()

    connection.close()

    return render_template(
        "companies.html",
        companies=company_list,
        selected_filter=selected_filter,
        filter_title=filter_title,
        filter_description=filter_description,
        search_query=search_query,
        selected_sort=selected_sort,
    )


# =========================================================
# 企業追加
# =========================================================


@app.route(
    "/companies/add",
    methods=["GET", "POST"],
)
def add_company():
    error = None

    if request.method == "POST":
        name = request.form.get(
            "name",
            "",
        ).strip()

        status = request.form.get(
            "status",
            "",
        ).strip()

        next_schedule = request.form.get(
            "next_schedule",
            "",
        ).strip()

        deadline = request.form.get(
            "deadline",
            "",
        ).strip()

        mypage_url = request.form.get(
            "mypage_url",
            "",
        ).strip()

        memo = request.form.get(
            "memo",
            "",
        ).strip()

        if not name:
            error = "企業名を入力してください。"

        elif not status:
            error = "選考状況を選択してください。"

        else:
            connection = get_db_connection()

            connection.execute(
                """
                INSERT INTO companies (
                    user_id,
                    name,
                    status,
                    next_schedule,
                    deadline,
                    mypage_url,
                    memo
                )
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    g.user["id"],
                    name,
                    status,
                    next_schedule,
                    deadline,
                    mypage_url,
                    memo,
                ),
            )

            # 新しく企業を登録した直後に、既存メールを再照合する。
            reprocess_email_messages_for_user(
                connection,
                g.user["id"],
            )

            connection.commit()
            connection.close()

            return redirect(url_for("companies"))

    return render_template(
        "add_company.html",
        error=error,
    )


# =========================================================
# 企業詳細
# =========================================================


@app.route("/companies/<int:company_id>")
def company_detail(company_id):
    connection = get_db_connection()

    company = connection.execute(
        """
        SELECT *
        FROM companies
        WHERE id = ?
          AND user_id = ?
        """,
        (
            company_id,
            g.user["id"],
        ),
    ).fetchone()

    if company is None:
        connection.close()
        abort(404)

    events = connection.execute(
        """
        SELECT *
        FROM company_events
        WHERE company_id = ?
        ORDER BY
            CASE
                WHEN event_date IS NULL
                  OR event_date = ''
                THEN 1
                ELSE 0
            END,
            event_date ASC,
            id ASC
        """,
        (company_id,),
    ).fetchall()

    notes = connection.execute(
        """
        SELECT *
        FROM company_notes
        WHERE company_id = ?
        ORDER BY id DESC
        """,
        (company_id,),
    ).fetchall()

    connection.close()

    return render_template(
        "company_detail.html",
        company=company,
        events=events,
        notes=notes,
    )


# =========================================================
# 企業編集
# =========================================================


@app.route(
    "/companies/<int:company_id>/edit",
    methods=["GET", "POST"],
)
def edit_company(company_id):
    connection = get_db_connection()

    company = connection.execute(
        """
        SELECT *
        FROM companies
        WHERE id = ?
          AND user_id = ?
        """,
        (
            company_id,
            g.user["id"],
        ),
    ).fetchone()

    if company is None:
        connection.close()
        abort(404)

    error = None

    if request.method == "POST":
        name = request.form.get(
            "name",
            "",
        ).strip()

        status = request.form.get(
            "status",
            "",
        ).strip()

        next_schedule = request.form.get(
            "next_schedule",
            "",
        ).strip()

        deadline = request.form.get(
            "deadline",
            "",
        ).strip()

        mypage_url = request.form.get(
            "mypage_url",
            "",
        ).strip()

        memo = request.form.get(
            "memo",
            "",
        ).strip()

        if not name:
            error = "企業名を入力してください。"

        elif not status:
            error = "選考状況を選択してください。"

        else:
            connection.execute(
                """
                UPDATE companies
                SET
                    name = ?,
                    status = ?,
                    next_schedule = ?,
                    deadline = ?,
                    mypage_url = ?,
                    memo = ?
                WHERE id = ?
                  AND user_id = ?
                """,
                (
                    name,
                    status,
                    next_schedule,
                    deadline,
                    mypage_url,
                    memo,
                    company_id,
                    g.user["id"],
                ),
            )

            # 企業名変更後も既存メールとの紐付けを再判定する。
            reprocess_email_messages_for_user(
                connection,
                g.user["id"],
            )

            connection.commit()
            connection.close()

            return redirect(
                url_for(
                    "company_detail",
                    company_id=company_id,
                )
            )

    connection.close()

    return render_template(
        "edit_company.html",
        company=company,
        error=error,
    )


# =========================================================
# 企業削除
# =========================================================


@app.route(
    "/companies/<int:company_id>/delete",
    methods=["POST"],
)
def delete_company(company_id):
    connection = get_db_connection()

    company = connection.execute(
        """
        SELECT *
        FROM companies
        WHERE id = ?
          AND user_id = ?
        """,
        (
            company_id,
            g.user["id"],
        ),
    ).fetchone()

    if company is None:
        connection.close()
        abort(404)

    connection.execute(
        """
        DELETE FROM companies
        WHERE id = ?
          AND user_id = ?
        """,
        (
            company_id,
            g.user["id"],
        ),
    )

    connection.commit()
    connection.close()

    return redirect(url_for("companies"))


# =========================================================
# 選考履歴追加
# =========================================================


@app.route(
    "/companies/<int:company_id>/events/add",
    methods=["POST"],
)
def add_company_event(company_id):
    event_date = request.form.get(
        "event_date",
        "",
    ).strip()

    event_type = request.form.get(
        "event_type",
        "",
    ).strip()

    title = request.form.get(
        "title",
        "",
    ).strip()

    memo = request.form.get(
        "event_memo",
        "",
    ).strip()

    connection = get_db_connection()

    company = connection.execute(
        """
        SELECT *
        FROM companies
        WHERE id = ?
          AND user_id = ?
        """,
        (
            company_id,
            g.user["id"],
        ),
    ).fetchone()

    if company is None:
        connection.close()
        abort(404)

    if not event_type or not title:
        connection.close()

        return redirect(
            url_for(
                "company_detail",
                company_id=company_id,
            )
        )

    connection.execute(
        """
        INSERT INTO company_events (
            company_id,
            event_date,
            event_type,
            title,
            memo
        )
        VALUES (?, ?, ?, ?, ?)
        """,
        (
            company_id,
            event_date,
            event_type,
            title,
            memo,
        ),
    )

    # 選考履歴から企業ステータスを自動更新
    new_status = determine_status(
        event_type,
        title,
    )

    if new_status:
        connection.execute(
            """
            UPDATE companies
            SET status = ?
            WHERE id = ?
            """,
            (
                new_status,
                company_id,
            ),
        )

    connection.commit()
    connection.close()

    return redirect(
        url_for(
            "company_detail",
            company_id=company_id,
        )
    )


# =========================================================
# 選考履歴削除
# =========================================================


@app.route(
    "/companies/<int:company_id>/events/<int:event_id>/delete",
    methods=["POST"],
)
def delete_company_event(
    company_id,
    event_id,
):
    connection = get_db_connection()

    # ログイン中のユーザーが所有する企業の履歴か確認する
    event = connection.execute(
        """
        SELECT
            company_events.*
        FROM company_events
        JOIN companies
          ON companies.id = company_events.company_id
        WHERE company_events.id = ?
          AND company_events.company_id = ?
          AND companies.user_id = ?
        """,
        (
            event_id,
            company_id,
            g.user["id"],
        ),
    ).fetchone()

    if event is None:
        connection.close()
        abort(404)

    connection.execute(
        """
        DELETE FROM company_events
        WHERE id = ?
          AND company_id = ?
        """,
        (
            event_id,
            company_id,
        ),
    )

    # 残った履歴から企業ステータスを再計算
    update_company_status_from_events(
        connection,
        company_id,
    )

    connection.commit()
    connection.close()

    return redirect(
        url_for(
            "company_detail",
            company_id=company_id,
        )
    )


# =========================================================
# 選考履歴編集
# =========================================================


@app.route(
    "/companies/<int:company_id>/events/<int:event_id>/edit",
    methods=["GET", "POST"],
)
def edit_company_event(
    company_id,
    event_id,
):
    connection = get_db_connection()

    company = connection.execute(
        """
        SELECT *
        FROM companies
        WHERE id = ?
          AND user_id = ?
        """,
        (
            company_id,
            g.user["id"],
        ),
    ).fetchone()

    if company is None:
        connection.close()
        abort(404)

    event = connection.execute(
        """
        SELECT *
        FROM company_events
        WHERE id = ?
          AND company_id = ?
        """,
        (
            event_id,
            company_id,
        ),
    ).fetchone()

    if event is None:
        connection.close()
        abort(404)

    error = None

    if request.method == "POST":
        event_date = request.form.get(
            "event_date",
            "",
        ).strip()

        event_type = request.form.get(
            "event_type",
            "",
        ).strip()

        title = request.form.get(
            "title",
            "",
        ).strip()

        memo = request.form.get(
            "event_memo",
            "",
        ).strip()

        if not event_type:
            error = "種類を選択してください。"

        elif not title:
            error = "タイトルを入力してください。"

        else:
            connection.execute(
                """
                UPDATE company_events
                SET
                    event_date = ?,
                    event_type = ?,
                    title = ?,
                    memo = ?
                WHERE id = ?
                  AND company_id = ?
                """,
                (
                    event_date,
                    event_type,
                    title,
                    memo,
                    event_id,
                    company_id,
                ),
            )

            # 編集後の履歴を基に企業ステータスを再計算
            update_company_status_from_events(
                connection,
                company_id,
            )

            connection.commit()
            connection.close()

            return redirect(
                url_for(
                    "company_detail",
                    company_id=company_id,
                )
            )

    connection.close()

    return render_template(
        "edit_event.html",
        company=company,
        event=event,
        error=error,
    )


# =========================================================
# Microsoft / Outlook OAuth連携開始
# =========================================================


def get_microsoft_oauth_app():
    return msal.ConfidentialClientApplication(
        MICROSOFT_CLIENT_ID,
        authority=MICROSOFT_AUTHORITY,
        client_credential=MICROSOFT_CLIENT_SECRET,
    )


@app.route("/outlook/connect")
def outlook_connect():
    microsoft_app = get_microsoft_oauth_app()

    flow = microsoft_app.initiate_auth_code_flow(
        scopes=MICROSOFT_SCOPES,
        redirect_uri=MICROSOFT_REDIRECT_URI,
    )

    if "error" in flow:
        abort(500)

    session["microsoft_auth_flow"] = flow

    return redirect(flow["auth_uri"])


# =========================================================
# Microsoft / Outlook OAuthコールバック
# =========================================================


@app.route("/outlook/callback")
def outlook_callback():
    flow = session.get("microsoft_auth_flow")

    if not flow:
        abort(400)

    cache = msal.SerializableTokenCache()

    microsoft_app = msal.ConfidentialClientApplication(
        MICROSOFT_CLIENT_ID,
        authority=MICROSOFT_AUTHORITY,
        client_credential=MICROSOFT_CLIENT_SECRET,
        token_cache=cache,
    )

    try:
        result = microsoft_app.acquire_token_by_auth_code_flow(
            flow,
            request.args,
        )
    except ValueError:
        abort(400)

    session.pop("microsoft_auth_flow", None)

    if "access_token" not in result:
        print("Microsoft OAuth error:", result)
        abort(400)

    access_token = result["access_token"]

    profile_response = requests.get(
        "https://graph.microsoft.com/v1.0/me",
        headers={"Authorization": f"Bearer {access_token}"},
        params={"$select": "displayName,mail,userPrincipalName"},
        timeout=20,
    )

    if not profile_response.ok:
        print(
            "Microsoft Graph error:",
            profile_response.text,
        )
        abort(400)

    profile = profile_response.json()

    outlook_address = (
        (profile.get("mail") or profile.get("userPrincipalName") or "").strip().lower()
    )

    display_name = (profile.get("displayName") or "Outlook / Microsoft 365").strip()

    if not outlook_address:
        abort(400)

    encrypted_token_cache = encrypt_token(cache.serialize())

    connection = get_db_connection()

    try:
        existing_account = connection.execute(
            """
            SELECT id
            FROM email_accounts
            WHERE user_id = ?
              AND lower(email_address) = lower(?)
            """,
            (
                g.user["id"],
                outlook_address,
            ),
        ).fetchone()

        if existing_account is None:
            cursor = connection.execute(
                """
                INSERT INTO email_accounts (
                    user_id,
                    provider,
                    email_address,
                    display_name,
                    imap_server,
                    imap_port,
                    use_ssl,
                    auth_type,
                    is_active
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, 1)
                """,
                (
                    g.user["id"],
                    "outlook",
                    outlook_address,
                    display_name,
                    "",
                    None,
                    1,
                    "oauth",
                ),
            )

            email_account_id = cursor.lastrowid

        else:
            email_account_id = existing_account["id"]

            connection.execute(
                """
                UPDATE email_accounts
                SET
                    provider = 'outlook',
                    display_name = ?,
                    auth_type = 'oauth',
                    is_active = 1
                WHERE id = ?
                  AND user_id = ?
                """,
                (
                    display_name,
                    email_account_id,
                    g.user["id"],
                ),
            )

        connection.execute(
            """
            INSERT INTO microsoft_oauth_tokens (
                user_id,
                email_account_id,
                token_cache,
                updated_at
            )
            VALUES (?, ?, ?, CURRENT_TIMESTAMP)
            ON CONFLICT(user_id, email_account_id)
            DO UPDATE SET
                token_cache = excluded.token_cache,
                updated_at = CURRENT_TIMESTAMP
            """,
            (
                g.user["id"],
                email_account_id,
                encrypted_token_cache,
            ),
        )

        connection.commit()

    except Exception:
        connection.rollback()
        raise

    finally:
        connection.close()

    return redirect(url_for("email_settings"))


# =========================================================
# Outlook / Microsoft 365 最新メール取得
# =========================================================


@app.route(
    "/outlook/sync",
    methods=["POST"],
)
def sync_outlook_messages():
    connection = get_db_connection()

    token_rows = connection.execute(
        """
        SELECT
            microsoft_oauth_tokens.*,
            email_accounts.email_address
        FROM microsoft_oauth_tokens
        JOIN email_accounts
          ON email_accounts.id = microsoft_oauth_tokens.email_account_id
        WHERE microsoft_oauth_tokens.user_id = ?
          AND email_accounts.user_id = ?
          AND email_accounts.provider = 'outlook'
          AND email_accounts.is_active = 1
        ORDER BY email_accounts.id ASC
        """,
        (
            g.user["id"],
            g.user["id"],
        ),
    ).fetchall()

    if not token_rows:
        connection.close()

        return redirect(
            url_for(
                "email_settings",
                sync="error",
                message="先にOutlookアカウントを連携してください。",
            )
        )

    saved_count = 0
    event_count = 0
    can_backfill = has_registered_companies(
        connection,
        g.user["id"],
    )

    try:
        for token_row in token_rows:
            cache = msal.SerializableTokenCache()
            cache.deserialize(decrypt_token(token_row["token_cache"]))

            microsoft_app = msal.ConfidentialClientApplication(
                MICROSOFT_CLIENT_ID,
                authority=MICROSOFT_AUTHORITY,
                client_credential=MICROSOFT_CLIENT_SECRET,
                token_cache=cache,
            )

            accounts = microsoft_app.get_accounts()

            if not accounts:
                raise RuntimeError(
                    f"{token_row['email_address']}を再連携してください。"
                )

            result = microsoft_app.acquire_token_silent(
                MICROSOFT_SCOPES,
                account=accounts[0],
            )

            if not result or "access_token" not in result:
                raise RuntimeError(
                    f"{token_row['email_address']}を再連携してください。"
                )

            if cache.has_state_changed:
                connection.execute(
                    """
                    UPDATE microsoft_oauth_tokens
                    SET
                        token_cache = ?,
                        updated_at = CURRENT_TIMESTAMP
                    WHERE id = ?
                      AND user_id = ?
                    """,
                    (
                        encrypt_token(cache.serialize()),
                        token_row["id"],
                        g.user["id"],
                    ),
                )

            headers = {
                "Authorization": f"Bearer {result['access_token']}",
                "Prefer": 'outlook.body-content-type="text"',
            }

            next_url = (
                "https://graph.microsoft.com/v1.0/"
                "me/mailFolders/inbox/messages"
            )
            first_page = True
            account_scanned_count = 0

            while (
                next_url
                and account_scanned_count < EMAIL_BACKFILL_SCAN_LIMIT
            ):
                params = None

                if first_page:
                    params = {
                        "$top": "20",
                        "$select": (
                            "id,"
                            "conversationId,"
                            "subject,"
                            "from,"
                            "receivedDateTime,"
                            "bodyPreview,"
                            "body,"
                            "isRead"
                        ),
                        "$orderby": "receivedDateTime desc",
                    }

                response = requests.get(
                    next_url,
                    headers=headers,
                    params=params,
                    timeout=20,
                )

                if not response.ok:
                    raise RuntimeError(
                        "Outlookメールの取得に失敗しました。"
                    )

                payload = response.json()
                messages = payload.get("value", [])

                if not messages:
                    break

                for message in messages:
                    if account_scanned_count >= EMAIL_BACKFILL_SCAN_LIMIT:
                        break

                    account_scanned_count += 1

                    sender_info = (message.get("from") or {}).get(
                        "emailAddress",
                        {},
                    )

                    sender_name = (sender_info.get("name") or "").strip()
                    sender_address = (sender_info.get("address") or "").strip()

                    if sender_name and sender_address:
                        sender = f"{sender_name} <{sender_address}>"
                    else:
                        sender = sender_address or sender_name or "（送信者不明）"

                    subject = message.get("subject") or "（件名なし）"
                    snippet = (message.get("bodyPreview") or "").strip()
                    body_text = clean_email_text(
                        (message.get("body") or {}).get("content") or ""
                    )[:8000]
                    analysis_text = normalize_email_analysis_text(
                        snippet,
                        body_text,
                    )
                    message_id = message.get("id")

                    if not message_id:
                        continue

                    connection.execute(
                        """
                        INSERT INTO email_messages (
                            user_id,
                            email_account_id,
                            provider,
                            message_id,
                            thread_id,
                            subject,
                            sender,
                            received_at,
                            snippet,
                            body,
                            is_unread,
                            fetched_at
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
                        ON CONFLICT(
                            email_account_id,
                            provider,
                            message_id
                        )
                        DO UPDATE SET
                            thread_id = excluded.thread_id,
                            subject = excluded.subject,
                            sender = excluded.sender,
                            received_at = excluded.received_at,
                            snippet = excluded.snippet,
                            body = excluded.body,
                            is_unread = excluded.is_unread,
                            fetched_at = CURRENT_TIMESTAMP
                        """,
                        (
                            g.user["id"],
                            token_row["email_account_id"],
                            "outlook",
                            message_id,
                            message.get("conversationId"),
                            subject,
                            sender,
                            message.get("receivedDateTime"),
                            snippet,
                            body_text,
                            (0 if message.get("isRead") else 1),
                        ),
                    )

                    update_email_message_company_link(
                        connection,
                        g.user["id"],
                        token_row["email_account_id"],
                        "outlook",
                        message_id,
                        subject,
                        sender,
                        analysis_text,
                    )

                    event_count += create_event_from_email_message(
                        connection,
                        g.user["id"],
                        token_row["email_account_id"],
                        "outlook",
                        message_id,
                    )

                    saved_count += 1

                connection.commit()

                company_mail_count = count_company_linked_emails(
                    connection,
                    g.user["id"],
                    "outlook",
                )

                # 最新20件は必ず取得する。
                # 登録企業メールが20件未満なら、古いメールへ遡る。
                if (
                    not can_backfill
                    or company_mail_count >= EMAIL_COMPANY_TARGET_COUNT
                ):
                    break

                next_url = payload.get("@odata.nextLink")
                first_page = False

        connection.commit()

    except Exception as error:
        connection.rollback()

        print(
            "Outlook sync error:",
            error,
        )

        return redirect(
            url_for(
                "email_settings",
                sync="error",
                message=str(error),
                company_only=request.form.get("company_only", "0"),
            )
        )

    finally:
        connection.close()

    return redirect(
        url_for(
            "email_settings",
            provider="outlook",
            sync="success",
            count=saved_count,
            events=event_count,
            company_only=request.form.get("company_only", "0"),
        )
    )


# =========================================================
# Google OAuth連携開始
# =========================================================


@app.route("/google/connect")
def google_connect():
    # PKCEで使用する一時的な検証文字列
    code_verifier = secrets.token_urlsafe(64)

    flow = Flow.from_client_secrets_file(
        GOOGLE_CLIENT_SECRETS_FILE,
        scopes=GOOGLE_OAUTH_SCOPES,
        redirect_uri=GOOGLE_OAUTH_REDIRECT_URI,
        code_verifier=code_verifier,
    )

    authorization_url, state = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
    )

    session["google_oauth_state"] = state
    session["google_code_verifier"] = code_verifier

    return redirect(authorization_url)


# =========================================================
# Google OAuthコールバック
# =========================================================


@app.route("/oauth2callback")
def google_oauth_callback():
    expected_state = session.get("google_oauth_state")

    code_verifier = session.get("google_code_verifier")

    received_state = request.args.get("state")

    if (
        not expected_state
        or not received_state
        or expected_state != received_state
        or not code_verifier
    ):
        abort(400)

    flow = Flow.from_client_secrets_file(
        GOOGLE_CLIENT_SECRETS_FILE,
        scopes=GOOGLE_OAUTH_SCOPES,
        state=expected_state,
        redirect_uri=GOOGLE_OAUTH_REDIRECT_URI,
        code_verifier=code_verifier,
    )

    flow.fetch_token(authorization_response=request.url)

    credentials = flow.credentials

    gmail_service = build(
        "gmail",
        "v1",
        credentials=credentials,
        cache_discovery=False,
    )

    profile = gmail_service.users().getProfile(userId="me").execute()

    gmail_address = (
        profile.get(
            "emailAddress",
            "",
        )
        .strip()
        .lower()
    )

    # 一時的なOAuth情報は使用後にセッションから削除する。
    session.pop(
        "google_oauth_state",
        None,
    )
    session.pop(
        "google_code_verifier",
        None,
    )

    if not gmail_address:
        abort(400)

    if not credentials.token:
        abort(400)

    encrypted_access_token = encrypt_token(credentials.token)
    encrypted_refresh_token = encrypt_token(credentials.refresh_token)

    expires_at = None

    if credentials.expiry is not None:
        expires_at = credentials.expiry.isoformat()

    granted_scopes = credentials.scopes or GOOGLE_OAUTH_SCOPES
    scopes_text = " ".join(granted_scopes)

    connection = get_db_connection()

    try:
        existing_account = connection.execute(
            """
            SELECT id
            FROM email_accounts
            WHERE user_id = ?
              AND lower(email_address) = lower(?)
            """,
            (
                g.user["id"],
                gmail_address,
            ),
        ).fetchone()

        if existing_account is None:
            cursor = connection.execute(
                """
                INSERT INTO email_accounts (
                    user_id,
                    provider,
                    email_address,
                    display_name,
                    imap_server,
                    imap_port,
                    use_ssl,
                    auth_type,
                    is_active
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, 1)
                """,
                (
                    g.user["id"],
                    "gmail",
                    gmail_address,
                    "Gmail",
                    "",
                    None,
                    1,
                    "oauth",
                ),
            )

            email_account_id = cursor.lastrowid

        else:
            email_account_id = existing_account["id"]

            connection.execute(
                """
                UPDATE email_accounts
                SET
                    provider = 'gmail',
                    display_name = 'Gmail',
                    auth_type = 'oauth',
                    is_active = 1
                WHERE id = ?
                  AND user_id = ?
                """,
                (
                    email_account_id,
                    g.user["id"],
                ),
            )

        # 同じGoogleアカウントを再連携した場合はトークンを更新する。
        # refresh_tokenが再発行されなかった場合は既存値を維持する。
        connection.execute(
            """
            INSERT INTO google_oauth_tokens (
                user_id,
                email_account_id,
                access_token,
                refresh_token,
                token_uri,
                client_id,
                scopes,
                expires_at,
                updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
            ON CONFLICT(user_id, email_account_id)
            DO UPDATE SET
                access_token = excluded.access_token,
                refresh_token = COALESCE(
                    excluded.refresh_token,
                    google_oauth_tokens.refresh_token
                ),
                token_uri = excluded.token_uri,
                client_id = excluded.client_id,
                scopes = excluded.scopes,
                expires_at = excluded.expires_at,
                updated_at = CURRENT_TIMESTAMP
            """,
            (
                g.user["id"],
                email_account_id,
                encrypted_access_token,
                encrypted_refresh_token,
                credentials.token_uri,
                credentials.client_id,
                scopes_text,
                expires_at,
            ),
        )

        connection.commit()

    except Exception:
        connection.rollback()
        raise

    finally:
        connection.close()

    return redirect(url_for("email_settings"))


# =========================================================
# デモ専用：架空メールを登録企業と再照合（外部通信なし）
# =========================================================
@app.route("/demo/mail/relink", methods=["POST"])
def demo_mail_relink():
    if not DEMO_MODE:
        abort(404)
    connection = get_db_connection()
    try:
        # DBは訪問者ごとに分離済み。メールはexample.invalidのサンプルだけ。
        matched = relink_demo_mail_data(connection, g.user["id"])
        connection.commit()
    finally:
        connection.close()
    return redirect(url_for("email_settings", sync="demo", count=matched))


# =========================================================
# メールアカウント設定
# =========================================================


@app.route(
    "/email-settings",
    methods=["GET", "POST"],
)
def email_settings():
    # 公開デモは画面閲覧と架空メールの再照合のみ。メールアカウント登録は禁止。
    if DEMO_MODE and request.method != "GET":
        abort(403)
    error = None

    provider_defaults = {
        "gmail": {
            "display_name": "Gmail",
            "auth_type": "oauth",
            "imap_server": "",
            "imap_port": "",
        },
        "outlook": {
            "display_name": "Outlook / Microsoft 365",
            "auth_type": "oauth",
            "imap_server": "",
            "imap_port": "",
        },
        "yahoo": {
            "display_name": "Yahoo!メール",
            "auth_type": "imap",
            "imap_server": "imap.mail.yahoo.co.jp",
            "imap_port": "993",
        },
        "icloud": {
            "display_name": "iCloudメール",
            "auth_type": "imap",
            "imap_server": "imap.mail.me.com",
            "imap_port": "993",
        },
        "imap": {
            "display_name": "その他のIMAPメール",
            "auth_type": "imap",
            "imap_server": "",
            "imap_port": "993",
        },
    }

    if request.method == "POST":
        provider = (
            request.form.get(
                "provider",
                "",
            )
            .strip()
            .lower()
        )

        email_address = (
            request.form.get(
                "email_address",
                "",
            )
            .strip()
            .lower()
        )

        display_name = request.form.get(
            "display_name",
            "",
        ).strip()

        imap_server = request.form.get(
            "imap_server",
            "",
        ).strip()

        imap_port_text = request.form.get(
            "imap_port",
            "",
        ).strip()

        use_ssl = 1 if request.form.get("use_ssl") == "on" else 0

        if provider not in provider_defaults:
            error = "メールサービスを選択してください。"

        elif not email_address:
            error = "メールアドレスを入力してください。"

        elif "@" not in email_address:
            error = "正しいメールアドレスを入力してください。"

        else:
            defaults = provider_defaults[provider]

            if not display_name:
                display_name = defaults["display_name"]

            auth_type = defaults["auth_type"]

            if auth_type == "oauth":
                imap_server = ""
                imap_port = None

            else:
                if not imap_server:
                    imap_server = defaults["imap_server"]

                if not imap_port_text:
                    imap_port_text = defaults["imap_port"]

                try:
                    imap_port = int(imap_port_text)
                except ValueError:
                    imap_port = None
                    error = "IMAPポートには数字を入力してください。"

                if not error and not imap_server:
                    error = "IMAPサーバーを入力してください。"

                if not error and (
                    imap_port is None or imap_port <= 0 or imap_port > 65535
                ):
                    error = "IMAPポートは1～65535の範囲で" "入力してください。"

            if not error:
                connection = get_db_connection()

                existing_account = connection.execute(
                    """
                    SELECT id
                    FROM email_accounts
                    WHERE user_id = ?
                      AND lower(email_address) = lower(?)
                    """,
                    (
                        g.user["id"],
                        email_address,
                    ),
                ).fetchone()

                if existing_account:
                    error = "このメールアドレスはすでに" "登録されています。"
                    connection.close()

                else:
                    connection.execute(
                        """
                        INSERT INTO email_accounts (
                            user_id,
                            provider,
                            email_address,
                            display_name,
                            imap_server,
                            imap_port,
                            use_ssl,
                            auth_type,
                            is_active
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, 1)
                        """,
                        (
                            g.user["id"],
                            provider,
                            email_address,
                            display_name,
                            imap_server,
                            imap_port,
                            use_ssl,
                            auth_type,
                        ),
                    )

                    connection.commit()
                    connection.close()

                    return redirect(url_for("email_settings"))

    connection = get_db_connection()

    accounts = connection.execute(
        """
        SELECT
            email_accounts.*,
            CASE
                WHEN email_accounts.provider = 'gmail'
                    AND google_oauth_tokens.id IS NOT NULL
                    THEN 1

                WHEN email_accounts.provider = 'outlook'
                    AND microsoft_oauth_tokens.id IS NOT NULL
                    THEN 1

                ELSE 0
            END AS oauth_connected
        FROM email_accounts
        LEFT JOIN google_oauth_tokens
          ON google_oauth_tokens.email_account_id = email_accounts.id
         AND google_oauth_tokens.user_id = email_accounts.user_id
         LEFT JOIN microsoft_oauth_tokens
            ON microsoft_oauth_tokens.email_account_id = email_accounts.id
            AND microsoft_oauth_tokens.user_id = email_accounts.user_id
        WHERE email_accounts.user_id = ?
        ORDER BY
            email_accounts.is_active DESC,
            email_accounts.id DESC
        """,
        (g.user["id"],),
    ).fetchall()

    selected_provider = (
        request.args.get(
            "provider",
            "all",
        )
        .strip()
        .lower()
    )

    if selected_provider not in {"all", "gmail", "outlook"}:
        selected_provider = "all"

    company_only = request.args.get(
        "company_only",
        "0",
    ).strip() == "1"

    message_conditions = [
        "email_messages.user_id = ?",
        "email_accounts.user_id = ?",
    ]
    message_parameters = [
        g.user["id"],
        g.user["id"],
    ]

    if selected_provider != "all":
        message_conditions.append(
            "email_messages.provider = ?"
        )
        message_parameters.append(selected_provider)

    if company_only:
        message_conditions.append(
            "email_messages.company_id IS NOT NULL"
        )

    message_where_clause = " AND ".join(message_conditions)

    email_messages = connection.execute(
        f"""
        SELECT
            email_messages.*,
            email_accounts.email_address,
            email_accounts.provider AS account_provider,
            companies.name AS company_name
        FROM email_messages
        JOIN email_accounts
          ON email_accounts.id = email_messages.email_account_id
         AND email_accounts.user_id = email_messages.user_id
        LEFT JOIN companies
          ON companies.id = email_messages.company_id
         AND companies.user_id = email_messages.user_id
        WHERE {message_where_clause}
        ORDER BY
            CASE
                WHEN email_messages.received_at IS NULL THEN 1
                ELSE 0
            END,
            email_messages.received_at DESC,
            email_messages.id DESC
        LIMIT 20
        """,
        message_parameters,
    ).fetchall()

    email_connected_count = connection.execute(
        """
        SELECT COUNT(*) AS count
        FROM email_accounts
        WHERE email_accounts.user_id = ?
        AND email_accounts.is_active = 1
        AND (
                (
                    email_accounts.provider = 'gmail'
                    AND EXISTS (
                        SELECT 1
                        FROM google_oauth_tokens
                        WHERE google_oauth_tokens.email_account_id = email_accounts.id
                        AND google_oauth_tokens.user_id = email_accounts.user_id
                    )
                )
                OR
                (
                    email_accounts.provider = 'outlook'
                    AND EXISTS (
                        SELECT 1
                        FROM microsoft_oauth_tokens
                        WHERE microsoft_oauth_tokens.email_account_id = email_accounts.id
                        AND microsoft_oauth_tokens.user_id = email_accounts.user_id
                    )
                )
            )
        """,
        (g.user["id"],),
    ).fetchone()["count"]

    connection.close()

    sync_status = request.args.get("sync", "").strip()
    sync_count = request.args.get("count", "0").strip()
    sync_event_count = request.args.get("events", "0").strip()
    sync_error = request.args.get("message", "").strip()

    return render_template(
        "email_settings.html",
        accounts=accounts,
        email_messages=email_messages,
        email_connected_count=email_connected_count,
        selected_provider=selected_provider,
        company_only=company_only,
        sync_status=sync_status,
        sync_count=sync_count,
        sync_event_count=sync_event_count,
        sync_error=sync_error,
        error=error,
        demo_mail_count=(len(email_messages) if DEMO_MODE else None),
    )


# =========================================================
# Gmail最新メール取得
# =========================================================


@app.route(
    "/gmail/sync",
    methods=["POST"],
)
def sync_gmail_messages():
    connection = get_db_connection()

    token_rows = connection.execute(
        """
        SELECT
            google_oauth_tokens.*,
            email_accounts.email_address
        FROM google_oauth_tokens
        JOIN email_accounts
          ON email_accounts.id = google_oauth_tokens.email_account_id
        WHERE google_oauth_tokens.user_id = ?
          AND email_accounts.user_id = ?
          AND email_accounts.provider = 'gmail'
          AND email_accounts.is_active = 1
        ORDER BY email_accounts.id ASC
        """,
        (
            g.user["id"],
            g.user["id"],
        ),
    ).fetchall()

    if not token_rows:
        connection.close()
        return redirect(
            url_for(
                "email_settings",
                sync="error",
                message="先にGoogleアカウントを連携してください。",
            )
        )

    saved_count = 0
    event_count = 0
    can_backfill = has_registered_companies(
        connection,
        g.user["id"],
    )

    try:
        for token_row in token_rows:
            credentials = build_google_credentials(token_row)

            if credentials.expired:
                if not credentials.refresh_token:
                    raise RuntimeError(
                        f"{token_row['email_address']}を再連携してください。"
                    )

                credentials.refresh(GoogleRequest())

                connection.execute(
                    """
                    UPDATE google_oauth_tokens
                    SET
                        access_token = ?,
                        refresh_token = COALESCE(?, refresh_token),
                        expires_at = ?,
                        updated_at = CURRENT_TIMESTAMP
                    WHERE id = ?
                      AND user_id = ?
                    """,
                    (
                        encrypt_token(credentials.token),
                        encrypt_token(credentials.refresh_token),
                        (
                            credentials.expiry.isoformat()
                            if credentials.expiry
                            else None
                        ),
                        token_row["id"],
                        g.user["id"],
                    ),
                )

            gmail_service = build(
                "gmail",
                "v1",
                credentials=credentials,
                cache_discovery=False,
            )

            page_token = None
            first_page = True
            account_scanned_count = 0

            while account_scanned_count < EMAIL_BACKFILL_SCAN_LIMIT:
                list_arguments = {
                    "userId": "me",
                    "maxResults": 20 if first_page else 50,
                    "labelIds": ["INBOX"],
                }

                if page_token:
                    list_arguments["pageToken"] = page_token

                message_list = (
                    gmail_service.users()
                    .messages()
                    .list(**list_arguments)
                    .execute()
                )

                message_references = message_list.get("messages", [])

                if not message_references:
                    break

                for message_reference in message_references:
                    if account_scanned_count >= EMAIL_BACKFILL_SCAN_LIMIT:
                        break

                    account_scanned_count += 1

                    message = (
                        gmail_service.users()
                        .messages()
                        .get(
                            userId="me",
                            id=message_reference["id"],
                            format="full",
                        )
                        .execute()
                    )

                    label_ids = message.get("labelIds", [])
                    body_text = extract_gmail_body_text(message)[:8000]
                    snippet = message.get("snippet", "").strip()
                    subject = get_gmail_header(message, "Subject") or "（件名なし）"
                    sender = get_gmail_header(message, "From") or "（送信者不明）"
                    analysis_text = normalize_email_analysis_text(
                        snippet,
                        body_text,
                    )

                    connection.execute(
                        """
                        INSERT INTO email_messages (
                            user_id,
                            email_account_id,
                            provider,
                            message_id,
                            thread_id,
                            subject,
                            sender,
                            received_at,
                            snippet,
                            body,
                            is_unread,
                            fetched_at
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
                        ON CONFLICT(email_account_id, provider, message_id)
                        DO UPDATE SET
                            thread_id = excluded.thread_id,
                            subject = excluded.subject,
                            sender = excluded.sender,
                            received_at = excluded.received_at,
                            snippet = excluded.snippet,
                            body = excluded.body,
                            is_unread = excluded.is_unread,
                            fetched_at = CURRENT_TIMESTAMP
                        """,
                        (
                            g.user["id"],
                            token_row["email_account_id"],
                            "gmail",
                            message.get("id"),
                            message.get("threadId"),
                            subject,
                            sender,
                            normalize_gmail_received_at(message),
                            snippet,
                            body_text,
                            1 if "UNREAD" in label_ids else 0,
                        ),
                    )

                    update_email_message_company_link(
                        connection,
                        g.user["id"],
                        token_row["email_account_id"],
                        "gmail",
                        message.get("id"),
                        subject,
                        sender,
                        analysis_text,
                    )

                    event_count += create_event_from_email_message(
                        connection,
                        g.user["id"],
                        token_row["email_account_id"],
                        "gmail",
                        message.get("id"),
                    )
                    saved_count += 1

                connection.commit()

                company_mail_count = count_company_linked_emails(
                    connection,
                    g.user["id"],
                    "gmail",
                )

                # 最新20件は必ず取得する。
                # 登録企業メールが20件未満なら、古いメールへ遡る。
                if (
                    not can_backfill
                    or company_mail_count >= EMAIL_COMPANY_TARGET_COUNT
                ):
                    break

                page_token = message_list.get("nextPageToken")
                first_page = False

                if not page_token:
                    break

        connection.commit()

    except Exception as error:

        print("GMAIL SYNC ERROR:", repr(error))

        connection.rollback()
        connection.close()

        return redirect(
            url_for(
                "email_settings",
                sync="error",
                message=str(error)[:180],
                company_only=request.form.get("company_only", "0"),
            )
        )

    connection.close()

    return redirect(
        url_for(
            "email_settings",
            provider="gmail",
            sync="success",
            count=saved_count,
            events=event_count,
            company_only=request.form.get("company_only", "0"),
        )
    )


# =========================================================
# メールアカウントの有効・無効切り替え
# =========================================================


@app.route(
    "/email-settings/<int:account_id>/toggle",
    methods=["POST"],
)
def toggle_email_account(account_id):
    connection = get_db_connection()

    # ログイン中ユーザーのメール設定だけ取得
    account = connection.execute(
        """
        SELECT *
        FROM email_accounts
        WHERE id = ?
          AND user_id = ?
        """,
        (
            account_id,
            g.user["id"],
        ),
    ).fetchone()

    if account is None:
        connection.close()
        abort(404)

    new_value = 0 if account["is_active"] else 1

    connection.execute(
        """
        UPDATE email_accounts
        SET is_active = ?
        WHERE id = ?
          AND user_id = ?
        """,
        (
            new_value,
            account_id,
            g.user["id"],
        ),
    )

    connection.commit()
    connection.close()

    return redirect(url_for("email_settings"))


# =========================================================
# メールアカウント削除
# =========================================================


@app.route(
    "/email-settings/<int:account_id>/delete",
    methods=["POST"],
)
def delete_email_account(account_id):
    connection = get_db_connection()

    # ログイン中ユーザーのメール設定だけ取得
    account = connection.execute(
        """
        SELECT *
        FROM email_accounts
        WHERE id = ?
          AND user_id = ?
        """,
        (
            account_id,
            g.user["id"],
        ),
    ).fetchone()

    if account is None:
        connection.close()
        abort(404)

    connection.execute(
        """
        DELETE FROM email_accounts
        WHERE id = ?
          AND user_id = ?
        """,
        (
            account_id,
            g.user["id"],
        ),
    )

    connection.commit()
    connection.close()

    return redirect(url_for("email_settings"))


# =========================================================
# アプリ起動
# =========================================================

if __name__ == "__main__":
    app.run(debug=not DEMO_MODE)
