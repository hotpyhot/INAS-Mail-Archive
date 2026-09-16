from __future__ import annotations

import csv
import hashlib
import html
import json
import logging
import os
import re
import shutil
import sys
import threading
import time
import tempfile
import tkinter as tk
import winreg
from dataclasses import dataclass
from datetime import datetime, timedelta
from email import policy
from email.header import decode_header, make_header
from email.parser import BytesParser
from email.utils import parseaddr, parsedate_to_datetime
from pathlib import Path

from PIL import Image, ImageDraw
import pystray
from xml.sax.saxutils import escape
from tkinter import filedialog, messagebox, ttk

from reportlab.lib.pagesizes import A4
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from tkinterdnd2 import DND_FILES, TkinterDnD
from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer
import webbrowser
from urllib.parse import urlparse

from app_info import APP_NAME, APP_VERSION, DISCLAIMER, about_text

CONFIG_DIR = Path(os.getenv("APPDATA", Path.home())) / "INAS" / "MailArchiveCommunity"
SETTINGS_PATH = CONFIG_DIR / "settings.json"
HISTORY_PATH = CONFIG_DIR / "processed_history.json"
SENDER_DICT_PATH = CONFIG_DIR / "sender_dictionary.json"
LOG_DIR = CONFIG_DIR / "logs"
LOG_PATH = LOG_DIR / "inas_mail_archive.log"
EXCLUDED_LOG_PATH = LOG_DIR / "excluded_mail.log"
DUPLICATE_LOG_PATH = LOG_DIR / "duplicate_mail.log"
ERROR_LOG_PATH = LOG_DIR / "error_mail.log"
MESSAGE_ID_REGISTRY_PATH = CONFIG_DIR / "message_id_registry.json"
DEFAULT_IMPORT = Path.home() / "Documents" / "INAS" / "MailImport"
DEFAULT_OUTPUT = Path.home() / "Documents" / "INAS" / "MailArchiveCommunity"
STARTUP_VALUE_NAME = "INAS Mail Archive Community Edition"
DEFAULT_PDF_TEMPLATE = "メール_{sender_label}{sender_honorific}_{datetime}.pdf"
DEFAULT_FOLDER_TEMPLATE = "{date}_追加資料_メール{sender_label}{sender_honorific}"

SUPPORTED_LANGUAGES = {"ja": "日本語", "en": "English", "vi": "Tiếng Việt"}
# INAS Mail Archive の標準表示言語。新規利用時は日本語で起動する。
DEFAULT_LANGUAGE = "ja"
_LOCALE_CACHE: dict[str, dict[str, str]] = {}


def format_bytes(size: int) -> str:
    """ファイルサイズを読みやすい単位に整形する。"""
    try:
        value = float(int(size or 0))
    except Exception:
        value = 0.0
    units = ("B", "KB", "MB", "GB", "TB")
    for unit in units:
        if value < 1024.0 or unit == units[-1]:
            if unit == "B":
                return f"{int(value)} {unit}"
            return f"{value:.1f} {unit}"
        value /= 1024.0
    return f"{int(size or 0)} B"


def load_locale(language: str) -> dict[str, str]:
    """外部localesフォルダから言語ファイルを読み込む。EXE配布時も言語データはEXE外に保持する。"""
    lang = language if language in SUPPORTED_LANGUAGES else DEFAULT_LANGUAGE
    if lang in _LOCALE_CACHE:
        return _LOCALE_CACHE[lang]
    candidates = []
    if getattr(sys, "frozen", False):
        candidates.append(Path(sys.executable).resolve().parent / "locales" / f"{lang}.json")
    candidates.append(Path(__file__).resolve().parent.parent / "locales" / f"{lang}.json")
    if "resource_path" in globals():
        candidates.append(resource_path(f"locales/{lang}.json"))
    for path in candidates:
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                _LOCALE_CACHE[lang] = {str(k): str(v) for k, v in data.items()}
                return _LOCALE_CACHE[lang]
        except Exception:
            continue
    _LOCALE_CACHE[lang] = {}
    return _LOCALE_CACHE[lang]


def tr(language: str, key: str, **kwargs) -> str:
    lang = language if language in SUPPORTED_LANGUAGES else DEFAULT_LANGUAGE
    data = load_locale(lang)
    fallback = load_locale("ja") if lang != "ja" else data
    text = data.get(key, fallback.get(key, key))
    try:
        return text.format(**kwargs)
    except Exception:
        return text

PDF_DISPLAY_DEFAULTS = {
    # メール内容（基本情報）
    "pdf_show_subject": True,
    "pdf_show_from": True,  # EMLに保存されている元のFromを表示
    "pdf_show_to": True,
    "pdf_show_cc": True,
    "pdf_show_bcc": False,
    "pdf_show_sent_datetime": True,
    "pdf_show_attachments": True,
    "pdf_show_body": True,
    # 詳細情報
    "pdf_show_sender_display_name": False,  # 命名に使う編集後の表示名
    "pdf_show_reply_to": False,
    "pdf_show_message_id": False,
    "pdf_show_importance": False,
    "pdf_show_attachment_count": False,
    "pdf_show_attachment_size": False,
    # ページ表示
    "pdf_show_page_number": True,
    "pdf_show_header": True,
    "pdf_show_footer": True,
    "pdf_show_app_name": False,
    "pdf_show_saved_datetime": False,
}


def resource_path(relative_path: str) -> Path:
    """Return resource paths for source and PyInstaller execution."""
    if getattr(sys, "frozen", False):
        base_path = Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
    else:
        base_path = Path(__file__).resolve().parent.parent
    return base_path / relative_path


APP_ICON_PATH = resource_path("assets/inas_mail_archive.ico")
APP_ICON_PNG_PATH = resource_path("assets/inas_mail_archive.png")

WINDOWS_RESERVED = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}


def setup_logging() -> None:
    """アプリの処理状況と例外をローテーションなしのUTF-8ログへ記録する。"""
    try:
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s [%(levelname)s] %(message)s",
            handlers=[
                logging.FileHandler(LOG_PATH, encoding="utf-8"),
            ],
            force=True,
        )
    except Exception:
        # ログ初期化に失敗してもアプリ本体は起動させる。
        pass


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as file:
        for chunk in iter(lambda: file.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def load_processed_history() -> dict:
    try:
        if HISTORY_PATH.exists():
            data = json.loads(HISTORY_PATH.read_text(encoding="utf-8"))
            return data if isinstance(data, dict) else {}
    except Exception:
        logging.exception("処理履歴を読み込めませんでした")
    return {}


def save_processed_history(history: dict) -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    HISTORY_PATH.write_text(
        json.dumps(history, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def append_event_log(path: Path, event_type: str, source_name: str, sender: str = "", detail: str = "") -> None:
    """Append a compact audit line without storing mail body or attachments."""
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        line = "\t".join([
            time.strftime("%Y/%m/%d %H:%M:%S"),
            event_type,
            source_name,
            sender,
            detail.replace("\t", " ").replace("\r", " ").replace("\n", " "),
        ])
        with path.open("a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        logging.exception("監査ログを書き込めませんでした: %s", path)


def load_message_id_registry() -> dict:
    try:
        if MESSAGE_ID_REGISTRY_PATH.exists():
            data = json.loads(MESSAGE_ID_REGISTRY_PATH.read_text(encoding="utf-8"))
            return data if isinstance(data, dict) else {}
    except Exception:
        logging.exception("Message-ID台帳を読み込めませんでした")
    return {}


def save_message_id_registry(registry: dict) -> None:
    try:
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        MESSAGE_ID_REGISTRY_PATH.write_text(
            json.dumps(registry, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except Exception:
        logging.exception("Message-ID台帳を保存できませんでした")


def load_sender_dictionary() -> dict[str, str]:
    try:
        if SENDER_DICT_PATH.exists():
            data = json.loads(SENDER_DICT_PATH.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                return {str(k): str(v) for k, v in data.items() if str(v).strip()}
    except Exception:
        logging.exception("差出人辞書を読み込めませんでした")
    return {}


def save_sender_dictionary(data: dict[str, str]) -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    SENDER_DICT_PATH.write_text(
        json.dumps(data, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def sender_key(from_text: str) -> str:
    display_name, address = parseaddr(from_text or "")
    if address:
        return address.strip().lower()
    return (from_text or "").strip().lower()


def clean_subject(subject: str) -> str:
    text = (subject or "").strip()
    # Re:/Fw:/Fwd: が連続している場合も先頭からまとめて除去する。
    while True:
        cleaned = re.sub(r"^(?:(?:re|fw|fwd)\s*:\s*)+", "", text, flags=re.IGNORECASE).strip()
        if cleaned == text:
            break
        text = cleaned
    return re.sub(r"\s+", " ", text).strip()


TEMPLATE_ITEM_SPECS = (
    ("From（元の差出人情報）", "{sender}", "田中 太郎 <tanaka@example.com>"),
    ("命名用表示名", "{sender_label}", "田中"),
    ("敬称", "{sender_honorific}", "氏"),
    ("件名", "{subject}", "追加資料送付"),
    ("To（宛先）", "{to}", "suzuki@example.com"),
    ("CC", "{cc}", "sato@example.com"),
    ("BCC", "{bcc}", "secret@example.com"),
    ("送信日", "{date}", "20260815"),
    ("送信時刻", "{time}", "151500"),
    ("送信日時", "{datetime}", "20260815_151500"),
    ("Reply-To（返信先）", "{reply_to}", "tanaka@example.com"),
    ("Message-ID", "{message_id}", "<sample@example.com>"),
    ("重要度", "{importance}", "High"),
    ("添付ファイル数", "{attachment_count}", "2"),
)


def render_name_template(template: str, *, context: dict[str, str]) -> str:
    """共通メール項目を命名テンプレートへ展開する。未知の変数はエラーとする。"""
    try:
        return (template or "").format_map(context).strip()
    except KeyError as exc:
        raise ValueError(f"使用できない置換項目があります：{{{exc.args[0]}}}") from exc
    except ValueError as exc:
        raise ValueError(f"命名テンプレートの書式が正しくありません：{exc}") from exc


def build_template_context(
    mail_data: "MailData | None",
    *,
    sender_label: str,
    subject: str,
    sender_honorific: str = "",
) -> dict[str, str]:
    """PDF出力と同じメール項目を基準に、命名用の共通コンテキストを作る。"""
    if mail_data is None:
        return {
            "sender": "田中 太郎 <tanaka@example.com>",
            "sender_label": "田中",
            "sender_honorific": sender_honorific or "氏",
            "subject": subject or "追加資料送付",
            "to": "suzuki@example.com",
            "cc": "sato@example.com",
            "bcc": "",
            "date": "20260815",
            "time": "151500",
            "datetime": "20260815_151500",
            "reply_to": "tanaka@example.com",
            "message_id": "<sample@example.com>",
            "importance": "High",
            "attachment_count": "2",
        }

    header_values = {label: value for label, value in parse_header_fields(mail_data.header_block)}
    date_part, _, time_part = (mail_data.formatted_date or "").partition("_")
    return {
        "sender": (mail_data.from_name or "").strip(),
        "sender_label": (sender_label or "").strip(),
        "sender_honorific": sender_honorific or "",
        "subject": (subject or "").strip(),
        "to": header_values.get("宛先", ""),
        "cc": header_values.get("CC", ""),
        "bcc": header_values.get("BCC", ""),
        "date": date_part,
        "time": time_part,
        "datetime": mail_data.formatted_date or "",
        "reply_to": header_values.get("Reply-To", ""),
        "message_id": header_values.get("Message-ID", ""),
        "importance": header_values.get("重要度", ""),
        "attachment_count": str(len(mail_data.attachments)),
    }


def unique_folder_path(path: Path) -> Path:
    if not path.exists():
        return path
    index = 2
    while True:
        candidate = path.with_name(f"{path.name}_{index}")
        if not candidate.exists():
            return candidate
        index += 1


@dataclass
class MailData:
    source_path: Path
    from_name: str
    formatted_date: str
    header_block: str
    body_text: str
    attachments: list[tuple[str, bytes]]
    source_hash: str


def dheader(value: str | None) -> str:
    if value is None:
        return ""
    try:
        return str(make_header(decode_header(value)))
    except Exception:
        return str(value)

def sanitize_filename(name: str, fallback: str = "noname") -> str:
    name = re.sub(r'[\\/:*?"<>|\x00-\x1f]', "_", name or "")
    name = name.strip().rstrip(". ")
    if not name:
        name = fallback
    stem = name.split(".")[0].upper()
    if stem in WINDOWS_RESERVED:
        name = f"_{name}"
    return name[:180]


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    index = 2
    while True:
        candidate = path.with_name(f"{path.stem}_{index}{path.suffix}")
        if not candidate.exists():
            return candidate
        index += 1


def html_to_text(html_str: str) -> str:
    text = html_str or ""
    text = re.sub(r"(?is)<head.*?>.*?</head>", "", text)
    text = re.sub(r"(?is)<script.*?>.*?</script>", "", text)
    text = re.sub(r"(?is)<style.*?>.*?</style>", "", text)
    text = re.sub(r"(?i)<br\s*/?>", "\n", text)
    text = re.sub(r"(?i)</p\s*>", "\n\n", text)
    text = re.sub(r"(?i)<li\s*>", "\n・", text)
    text = re.sub(r"(?s)<[^>]+>", "", text)
    text = html.unescape(text).replace("\r\n", "\n").replace("\r", "\n")
    text = "\n".join(line.rstrip() for line in text.splitlines())
    return re.sub(r"\n{3,}", "\n\n", text).strip()


def normalize_text(text: str | None) -> str:
    return (text or "").replace("\t", "    ").replace("\r\n", "\n").replace("\r", "\n")


def build_header_block(msg) -> str:
    lines: list[str] = []
    labels = (
        ("From", "From"), ("To", "To"), ("Cc", "Cc"), ("Bcc", "Bcc"),
        ("Reply-To", "Reply-To"), ("Message-ID", "Message-ID"),
        ("Subject", "Subject"), ("Date", "Date"),
    )
    for source, label in labels:
        for value in msg.get_all(source, []):
            lines.append(f"{label}: {dheader(value)}")

    # 重要度はメールソフトによってヘッダー名が異なるため、代表的な順で取得する。
    importance = msg.get("Importance") or msg.get("Priority") or msg.get("X-Priority")
    if importance:
        lines.append(f"Importance: {dheader(importance)}")
    if msg.get("Date"):
        try:
            dt = parsedate_to_datetime(msg["Date"])
            lines.append(f"送信日時: {dt.strftime('%Y/%m/%d %H:%M:%S %z')}")
        except Exception:
            pass
    return "\n".join(lines)



def read_eml_summary(eml_path: Path) -> dict[str, object]:
    """EMLの一覧・画面表示用ヘッダーを一元取得する。"""
    with eml_path.open("rb") as fh:
        msg = BytesParser(policy=policy.default).parse(fh)

    sender = dheader(msg.get("From")).strip()
    subject = dheader(msg.get("Subject")).strip()
    to_value = dheader(msg.get("To")).strip()
    cc_value = dheader(msg.get("Cc")).strip()

    raw_date = msg.get("Date")
    display_date = ""
    if raw_date:
        try:
            dt = parsedate_to_datetime(raw_date)
            display_date = dt.strftime("%Y/%m/%d %H:%M:%S %z")
        except Exception:
            display_date = dheader(raw_date).strip()

    attachment_count = 0
    for part in msg.iter_attachments():
        filename = part.get_filename()
        payload = part.get_payload(decode=True)
        if not filename or payload is None:
            continue
        decoded_name = sanitize_filename(dheader(filename), "attachment")
        disposition = (part.get_content_disposition() or "").lower()
        content_id = (part.get("Content-ID") or "").strip()
        looks_like_signature_image = bool(
            re.match(r"(?i)^image\d{3,}\.(png|jpe?g|gif|bmp)$", decoded_name)
        )
        if disposition == "inline" and (content_id or looks_like_signature_image):
            continue
        attachment_count += 1

    return {
        "sender": sender,
        "subject": subject,
        "to": to_value,
        "cc": cc_value,
        "date": display_date,
        "attachments": attachment_count,
    }


def detect_external_download_links(text_value: str) -> list[dict[str, str]]:
    """メール本文から外部ファイル転送・ダウンロード候補URLを抽出する。"""
    if not text_value:
        return []

    # URL抽出。末尾の句読点や括弧は除外。
    urls = re.findall(r'https?://[^\s<>"\']+', text_value)
    cleaned = []
    seen = set()

    service_patterns = [
        ("ギガファイル便", ("gigafile.nu",)),
        ("firestorage", ("firestorage.jp",)),
        ("データ便", ("datadeliver.net", "datadeliver.jp")),
        ("tenpu", ("tenpu.me",)),
        ("FilePost", ("file-post.net",)),
        ("Bizストレージ ファイルシェア", ("fs2.biz-storage.jp", "biz-storage.jp")),
        ("SECURE DELIVER", ("secure-deliver.jp",)),
        ("Smooth File", ("smoothfile.jp",)),
        ("Box", ("box.com", "app.box.com")),
        ("Dropbox", ("dropbox.com",)),
        ("Google Drive", ("drive.google.com",)),
        ("OneDrive / SharePoint", ("1drv.ms", "sharepoint.com")),
    ]

    generic_keywords = (
        "download", "dl", "file", "files", "transfer", "share",
        "storage", "deliver", "gigafile", "firestorage"
    )

    for raw in urls:
        url = raw.rstrip('.,;:!?)]}）】」』')
        key = url.lower()
        if key in seen:
            continue
        seen.add(key)

        try:
            host = (urlparse(url).hostname or "").lower()
        except Exception:
            host = ""

        service = ""
        for label, domains in service_patterns:
            if any(host == d or host.endswith("." + d) for d in domains):
                service = label
                break

        # 既知サービス以外も、URL文字列にダウンロード系語が含まれる場合は候補にする。
        looks_download = bool(service) or any(k in key for k in generic_keywords)
        if not looks_download:
            continue

        cleaned.append({
            "service": service or host or "外部リンク",
            "url": url,
        })

    return cleaned

def parse_eml(eml_path: Path, ignore_inline_images: bool = True) -> MailData:
    with eml_path.open("rb") as file:
        msg = BytesParser(policy=policy.default).parse(file)

    header_block = build_header_block(msg)
    from_name = dheader(msg.get("From"))
    try:
        dt = parsedate_to_datetime(msg.get("Date")) if msg.get("Date") else None
    except Exception:
        dt = None
    formatted_date = dt.strftime("%Y%m%d_%H%M%S") if dt else time.strftime("%Y%m%d_%H%M%S")

    body = msg.get_body(preferencelist=("plain", "html"))
    if body is None:
        body_text = ""
    elif body.get_content_type() == "text/html":
        body_text = html_to_text(body.get_content())
    else:
        body_text = normalize_text(body.get_content())

    attachments: list[tuple[str, bytes]] = []
    for part in msg.iter_attachments():
        filename = part.get_filename()
        payload = part.get_payload(decode=True)
        if not filename or payload is None:
            continue
        decoded_name = sanitize_filename(dheader(filename), "attachment")
        if ignore_inline_images:
            disposition = (part.get_content_disposition() or "").lower()
            content_id = (part.get("Content-ID") or "").strip()
            looks_like_signature_image = bool(re.match(r"(?i)^image\d{3,}\.(png|jpe?g|gif|bmp)$", decoded_name))
            if disposition == "inline" and (content_id or looks_like_signature_image):
                continue
        attachments.append((decoded_name, payload))

    return MailData(
        eml_path,
        from_name,
        formatted_date,
        header_block,
        body_text,
        attachments,
        file_sha256(eml_path),
    )


def register_pdf_font() -> str:
    candidates = [
        ("MSGothic", Path(r"C:\Windows\Fonts\msgothic.ttc"), 0),
        ("YuGothic", Path(r"C:\Windows\Fonts\YuGothR.ttc"), 0),
        ("Meiryo", Path(r"C:\Windows\Fonts\meiryo.ttc"), 0),
    ]
    for name, path, index in candidates:
        if not path.exists():
            continue
        try:
            pdfmetrics.registerFont(TTFont(name, str(path), subfontIndex=index))
            return name
        except Exception:
            continue
    raise RuntimeError("日本語PDFフォントを登録できません。Windowsの日本語フォントを確認してください。")


def parse_header_fields(header_block: str) -> list[tuple[str, str]]:
    """EMLヘッダー文字列をPDF表示用の項目へ整理する。"""
    collected: dict[str, list[str]] = {}
    sent_datetime = ""
    for line in normalize_text(header_block).splitlines():
        if line.startswith("送信日時:"):
            sent_datetime = line.partition(":")[2].strip()
            continue
        key, sep, value = line.partition(":")
        if not sep:
            continue
        key = key.strip()
        value = value.strip()
        if value:
            collected.setdefault(key, []).append(value)

    rows: list[tuple[str, str]] = []
    label_map = (
        ("From", "差出人"),
        ("To", "宛先"),
        ("Cc", "CC"),
        ("Bcc", "BCC"),
        ("Reply-To", "Reply-To"),
        ("Message-ID", "Message-ID"),
        ("Importance", "重要度"),
        ("Subject", "件名"),
    )
    for key, label in label_map:
        values = collected.get(key, [])
        if values:
            rows.append((label, "\n".join(values)))

    if sent_datetime:
        rows.append(("送信日時", sent_datetime))
    elif collected.get("Date"):
        rows.append(("送信日時", "\n".join(collected["Date"])))
    return rows


def text_to_pdf_markup(text: str, make_links: bool = True) -> str:
    """プレーンテキストをReportLab Paragraph用に安全化し、URLをリンク化する。"""
    source = normalize_text(text or "")
    url_pattern = re.compile(r"https?://[^\s<>\"']+")
    pieces: list[str] = []
    position = 0
    for match in url_pattern.finditer(source):
        pieces.append(escape(source[position:match.start()]))
        raw_url = match.group(0)
        # 文末の句読点や閉じ括弧はリンクから外す。
        trimmed = raw_url.rstrip(".,;:!?、。)]}）］」』")
        trailing = raw_url[len(trimmed):]
        if make_links and trimmed:
            href = escape(trimmed, {'"': '&quot;', "'": '&apos;'})
            label = escape(trimmed)
            pieces.append(f'<link href="{href}" color="#245B9E"><u>{label}</u></link>')
        else:
            pieces.append(escape(trimmed))
        pieces.append(escape(trailing))
        position = match.end()
    pieces.append(escape(source[position:]))
    return "".join(pieces).replace("\n", "<br/>")


def create_pdf(
    pdf_path: Path,
    header_block: str,
    body_text: str,
    attachments: list[tuple[str, bytes]],
    display_subject: str = "",
    display_sender_name: str = "",
    display_options: dict | None = None,
    pdf_language: str = "ja",
) -> None:
    """設定された表示項目に従って、メールを見やすいA4 PDFとして出力する。"""
    options = dict(PDF_DISPLAY_DEFAULTS)
    if isinstance(display_options, dict):
        for key in options:
            if key in display_options:
                options[key] = bool(display_options[key])

    font_name = register_pdf_font()
    subject = (display_subject or tr(pdf_language, "mail_default")).strip() or tr(pdf_language, "mail_default")
    doc = SimpleDocTemplate(
        str(pdf_path),
        pagesize=A4,
        leftMargin=40,
        rightMargin=40,
        topMargin=64,
        bottomMargin=58,
        title=subject,
        author=APP_NAME,
        subject=subject,
    )

    title_style = ParagraphStyle(
        name="MailTitle",
        fontName=font_name,
        fontSize=15,
        leading=19,
        alignment=TA_LEFT,
        wordWrap="CJK",
        spaceAfter=10,
    )
    section_style = ParagraphStyle(
        name="SectionTitle",
        fontName=font_name,
        fontSize=10,
        leading=13,
        alignment=TA_LEFT,
        spaceBefore=5,
        spaceAfter=5,
    )
    label_style = ParagraphStyle(
        name="HeaderLabel",
        fontName=font_name,
        fontSize=9,
        leading=11,
        alignment=TA_LEFT,
    )
    value_style = ParagraphStyle(
        name="HeaderValue",
        fontName=font_name,
        fontSize=9,
        leading=12,
        alignment=TA_LEFT,
        wordWrap="CJK",
        splitLongWords=True,
    )
    body_style = ParagraphStyle(
        name="MailBody",
        fontName=font_name,
        fontSize=10,
        leading=14,
        alignment=TA_LEFT,
        wordWrap="CJK",
        splitLongWords=True,
        allowWidows=1,
        allowOrphans=1,
    )
    small_style = ParagraphStyle(
        name="SmallText",
        fontName=font_name,
        fontSize=8.5,
        leading=11,
        alignment=TA_LEFT,
        wordWrap="CJK",
    )

    story = []
    if options["pdf_show_subject"]:
        story.append(Paragraph(text_to_pdf_markup(subject, make_links=False), title_style))

    label_option = {
        "差出人": "pdf_show_from",
        "宛先": "pdf_show_to",
        "CC": "pdf_show_cc",
        "BCC": "pdf_show_bcc",
        "Reply-To": "pdf_show_reply_to",
        "Message-ID": "pdf_show_message_id",
        "重要度": "pdf_show_importance",
        "件名": "pdf_show_subject",
        "送信日時": "pdf_show_sent_datetime",
    }
    pdf_label_keys = {"差出人":"from_label", "宛先":"to_label", "CC":"cc", "BCC":"bcc", "Reply-To":"reply_to", "Message-ID":"message_id", "重要度":"importance", "件名":"subject", "送信日時":"sent_datetime"}
    header_rows = [
        (tr(pdf_language, pdf_label_keys.get(label, label)), value)
        for label, value in parse_header_fields(header_block)
        if options.get(label_option.get(label, ""), True)
        and not (label == "件名" and options["pdf_show_subject"])
    ]

    # Fromの元情報と命名用表示名は別データとして扱う。
    # 例：元のFrom「田中 太郎 <tanaka@example.com>」／命名用表示名「田中」。
    naming_sender = (display_sender_name or "").strip()
    if options["pdf_show_sender_display_name"] and naming_sender:
        header_rows.append((tr(pdf_language, "sender_display_name"), naming_sender))
    if options["pdf_show_attachment_count"]:
        header_rows.append((tr(pdf_language, "attachment_count_label"), tr(pdf_language, "count_items", count=len(attachments))))
    if header_rows:
        table_data = []
        for label, value in header_rows:
            table_data.append([
                Paragraph(escape(label), label_style),
                Paragraph(text_to_pdf_markup(value), value_style),
            ])
        info_table = Table(table_data, colWidths=[70, doc.width - 70], hAlign="LEFT")
        info_table.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#F2F4F7")),
            ("BOX", (0, 0), (-1, -1), 0.4, colors.HexColor("#C9CED6")),
            ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#D9DDE3")),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))
        story.extend([info_table, Spacer(1, 10)])

    if attachments and options["pdf_show_attachments"]:
        story.append(Paragraph(tr(pdf_language, "attachments"), section_style))

        def format_bytes(size: int) -> str:
            if size < 1024:
                return f"{size} B"
            if size < 1024 * 1024:
                return f"{size / 1024:.1f} KB"
            return f"{size / (1024 * 1024):.1f} MB"

        if options["pdf_show_attachment_size"]:
            attachment_markup = "<br/>".join(
                f"・{escape(filename)}（{format_bytes(len(data))}）" for filename, data in attachments
            )
        else:
            attachment_markup = "<br/>".join(
                f"・{escape(filename)}" for filename, _ in attachments
            )
        story.extend([
            Paragraph(attachment_markup, small_style),
            Spacer(1, 10),
        ])

    if options["pdf_show_body"]:
        story.append(Paragraph(tr(pdf_language, "body"), section_style))
        source_body = normalize_text(body_text).strip() or tr(pdf_language, "no_body")
        story.append(Paragraph(text_to_pdf_markup(source_body), body_style))

    if not story:
        story.append(Paragraph(tr(pdf_language, "no_pdf_items"), body_style))

    saved_at = datetime.now().strftime("%Y/%m/%d %H:%M:%S")

    def draw_page(canvas, doc_obj):
        canvas.saveState()
        canvas.setTitle(subject)
        canvas.setAuthor(APP_NAME)
        canvas.setSubject(subject)
        page_width, page_height = A4
        canvas.setFont(font_name, 7.5)
        canvas.setFillColor(colors.HexColor("#68717D"))

        if options["pdf_show_header"] and doc_obj.page > 1:
            header_text = subject if options["pdf_show_subject"] else APP_NAME
            if len(header_text) > 55:
                header_text = header_text[:52] + "..."
            canvas.drawString(doc.leftMargin, page_height - 30, header_text)
            canvas.setStrokeColor(colors.HexColor("#D9DDE3"))
            canvas.setLineWidth(0.4)
            canvas.line(doc.leftMargin, page_height - 35, page_width - doc.rightMargin, page_height - 35)

        footer_texts = []
        if options["pdf_show_app_name"]:
            footer_texts.append(APP_NAME)
        if options["pdf_show_saved_datetime"]:
            footer_texts.append(f"{tr(pdf_language, 'saved_datetime')} {saved_at}")
        if options["pdf_show_footer"] or footer_texts or options["pdf_show_page_number"]:
            if options["pdf_show_footer"]:
                canvas.setStrokeColor(colors.HexColor("#D9DDE3"))
                canvas.setLineWidth(0.4)
                canvas.line(doc.leftMargin, 35, page_width - doc.rightMargin, 35)
            if footer_texts:
                canvas.drawString(doc.leftMargin, 23, "  /  ".join(footer_texts))
            if options["pdf_show_page_number"]:
                page_label = tr(pdf_language, "page_label", page=doc_obj.page)
                canvas.drawRightString(page_width - doc.rightMargin, 23, page_label)
        canvas.restoreState()

    doc.build(story, onFirstPage=draw_page, onLaterPages=draw_page)


def load_settings() -> dict:
    defaults = {
        "import_folder": str(DEFAULT_IMPORT),
        "watch_enabled": True,
        "post_action": "move_to_output",
        "open_after_save": True,
        "auto_start": False,
        "resident_enabled": True,
        "ignore_inline_images": True,
        "pdf_name_template": DEFAULT_PDF_TEMPLATE,
        "folder_name_template": DEFAULT_FOLDER_TEMPLATE,
        "language": DEFAULT_LANGUAGE,
        "pdf_language": "same",
        "default_output_folder": str(DEFAULT_OUTPUT),
        "excluded_addresses": [],
        "excluded_domains": [],
        "confirm_auto_detected_mail": True,
        "delete_excluded_temp_eml": True,
        "registered_save_locations": [],
        "self_email_addresses": [],
        "sender_honorific": "氏",
        "sender_honorific_rules": {},
        "duplicate_check_enabled": True,
        "recent_destination_suggestions": True,
        "attachment_selective_save": True,
        "external_link_detection_enabled": True,
        "auto_show_pending_on_new_eml": True,
        **PDF_DISPLAY_DEFAULTS,
    }
    try:
        if SETTINGS_PATH.exists():
            defaults.update(json.loads(SETTINGS_PATH.read_text(encoding="utf-8")))
            # Ver.1.4.10以前の標準テンプレートは、{sender_label} 自体に「氏」を含めていた。
            # Ver.1.4.11では {sender_label}=命名用表示名 に統一するため、標準値だけ安全に移行する。
            if defaults.get("pdf_name_template") == "メール_{sender_label}_{datetime}.pdf":
                defaults["pdf_name_template"] = DEFAULT_PDF_TEMPLATE
            if defaults.get("folder_name_template") == "{date}_追加資料_メール{sender_label}":
                defaults["folder_name_template"] = DEFAULT_FOLDER_TEMPLATE
            # Ver.1.6.22: 標準テンプレートの固定敬称を動的トークンへ移行する。
            if defaults.get("pdf_name_template") in {
                "メール_{sender_label}氏_{datetime}.pdf",
                "メール_{sender_label}様_{datetime}.pdf",
            }:
                defaults["pdf_name_template"] = DEFAULT_PDF_TEMPLATE
            if defaults.get("folder_name_template") in {
                "{date}_追加資料_メール{sender_label}氏",
                "{date}_追加資料_メール{sender_label}様",
            }:
                defaults["folder_name_template"] = DEFAULT_FOLDER_TEMPLATE
    except Exception:
        pass
    return defaults


def save_settings(settings: dict) -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    SETTINGS_PATH.write_text(json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8")


_SINGLE_INSTANCE_MUTEX = None


def acquire_single_instance_mutex() -> bool:
    """Prevent multiple resident instances from watching the same import folder.

    Returns True for the first instance.  A second instance returns False and
    exits before creating a watchdog observer.  The mutex handle is kept in a
    module global for the lifetime of the process.
    """
    global _SINGLE_INSTANCE_MUTEX
    if os.name != "nt":
        return True
    try:
        import ctypes
        from ctypes import wintypes

        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel32.CreateMutexW.argtypes = (ctypes.c_void_p, wintypes.BOOL, wintypes.LPCWSTR)
        kernel32.CreateMutexW.restype = wintypes.HANDLE
        handle = kernel32.CreateMutexW(None, False, "Local\\INAS_Mail_Archive_SingleInstance")
        if not handle:
            return True
        ERROR_ALREADY_EXISTS = 183
        if ctypes.get_last_error() == ERROR_ALREADY_EXISTS:
            kernel32.CloseHandle(handle)
            return False
        _SINGLE_INSTANCE_MUTEX = handle
        return True
    except Exception:
        # Single-instance protection must never make the application unusable.
        logging.exception("単一起動確認に失敗しました")
        return True


def release_single_instance_mutex() -> None:
    global _SINGLE_INSTANCE_MUTEX
    if os.name != "nt" or not _SINGLE_INSTANCE_MUTEX:
        return
    try:
        import ctypes
        ctypes.windll.kernel32.CloseHandle(_SINGLE_INSTANCE_MUTEX)
    except Exception:
        pass
    _SINGLE_INSTANCE_MUTEX = None


class EmlCreatedHandler(FileSystemEventHandler):
    def __init__(self, callback):
        self.callback = callback


    def on_created(self, event):
        if not event.is_directory and event.src_path.lower().endswith(".eml"):
            self.callback(Path(event.src_path))

    def on_moved(self, event):
        if not event.is_directory and event.dest_path.lower().endswith(".eml"):
            self.callback(Path(event.dest_path))


class MailArchiveApp:
    def __init__(self):
        setup_logging()
        logging.info("%s Ver.%s を起動しました", APP_NAME, APP_VERSION)
        self.root = TkinterDnD.Tk()
        self.root.title(f"{APP_NAME} Community Edition Ver.{APP_VERSION}")
        try:
            self.root.iconbitmap(default=str(APP_ICON_PATH))
        except Exception:
            logging.exception("ウィンドウアイコンを設定できませんでした")
        self.root.geometry("760x700")
        # 小さい画面でも下部の固定操作ボタンを表示できるよう、最小高さを抑える。
        self.root.minsize(700, 560)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        self.settings = load_settings()
        self.processed_history = load_processed_history()
        self.sender_dictionary = load_sender_dictionary()
        Path(self.settings["import_folder"]).mkdir(parents=True, exist_ok=True)

        self.mail_data: MailData | None = None
        self.observer: Observer | None = None
        self.pending_paths: set[Path] = set()
        self.eml_queue: list[Path] = []
        self.queued_paths: set[Path] = set()
        self.known_eml_signatures: dict[Path, tuple[int, int]] = {}
        # OneDrive/Power Automate can generate multiple filesystem events for one EML.
        # Keep a short in-process guard so one file produces only one confirmation.
        self.recently_detected_eml: dict[Path, tuple[tuple[int, int], float]] = {}
        self.poll_job: str | None = None
        self.tray_icon: pystray.Icon | None = None
        self.tray_thread: threading.Thread | None = None
        self.is_exiting = False
        self.tray_notice_shown = False
        self.last_detected_at: datetime | None = None

        self.var_from = tk.StringVar()
        self.var_subject = tk.StringVar()
        self.var_pdf = tk.StringVar()
        self.var_folder = tk.StringVar()
        self.var_import = tk.StringVar(value=self.settings["import_folder"])
        self.var_output = tk.StringVar(value="")
        self.var_default_output = tk.StringVar(value=str(self.settings.get("default_output_folder", DEFAULT_OUTPUT)))
        self.var_watch = tk.BooleanVar(value=bool(self.settings["watch_enabled"]))
        self.var_post_action = tk.StringVar(value=self.settings["post_action"])
        self.var_open_after = tk.BooleanVar(value=bool(self.settings["open_after_save"]))
        self.var_auto_start = tk.BooleanVar(value=bool(self.settings.get("auto_start", False)))
        self.var_resident = tk.BooleanVar(value=bool(self.settings.get("resident_enabled", True)))
        self.var_ignore_inline = tk.BooleanVar(value=bool(self.settings.get("ignore_inline_images", True)))
        self.var_language = tk.StringVar(value=self.settings.get("language", DEFAULT_LANGUAGE) if self.settings.get("language", DEFAULT_LANGUAGE) in SUPPORTED_LANGUAGES else DEFAULT_LANGUAGE)
        self.var_status = tk.StringVar(value=self.t("status_ready"))
        self.var_monitor_state = tk.StringVar()
        self.var_last_detected = tk.StringVar()
        self.var_pending_count = tk.StringVar()
        self.var_from.trace_add("write", self.update_names)
        self.var_subject.trace_add("write", self.update_names)

        self.build_menu()
        self.build_ui()
        self.sync_startup_setting(show_error=False)
        if self.var_resident.get():
            self.start_tray_icon()
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind("<<Drop>>", self.drop_eml)

        if self.var_watch.get():
            self.start_watcher()
        self.root.after(500, self.poll_import_folder)
        # 起動時にPower Automate等で蓄積された未処理EMLを一覧確認する。
        self.root.after(700, self.show_pending_mail_list)
        startup_mode = "--startup" in sys.argv[1:]
        eml_args = [Path(arg) for arg in sys.argv[1:] if arg != "--startup"]
        valid_eml_args = [arg for arg in eml_args if arg.is_file() and arg.suffix.lower() == ".eml"]
        if valid_eml_args:
            self.root.after(200, lambda paths=valid_eml_args: self.queue_emls(paths))
        elif startup_mode and self.var_resident.get():
            self.root.after(100, self.root.withdraw)

    def t(self, key: str, **kwargs) -> str:
        return tr(self.var_language.get() if hasattr(self, "var_language") else self.settings.get("language", DEFAULT_LANGUAGE), key, **kwargs)

    def change_language(self, language: str):
        if language not in SUPPORTED_LANGUAGES:
            return
        self.var_language.set(language)
        self.settings["language"] = language
        self.persist_settings()
        self.root.title(f"{APP_NAME} Community Edition Ver.{APP_VERSION}")
        for child in self.root.winfo_children():
            child.destroy()
        self.build_menu()
        self.build_ui()
        if self.tray_icon is not None:
            self.stop_tray_icon()
            if self.var_resident.get():
                self.start_tray_icon()
        self.var_status.set(self.t("status_language_changed"))
        if self.mail_data:
            self.attachment_label.configure(text=self.t("attachment_count", count=len(self.mail_data.attachments)))

    def build_menu(self):
        menu_bar = tk.Menu(self.root)

        file_menu = tk.Menu(menu_bar, tearoff=False)
        file_menu.add_command(label=self.t("menu_select_eml_multi"), command=self.select_eml)
        file_menu.add_command(label=self.t("history"), command=self.show_history)
        file_menu.add_command(label=self.t("menu_open_import_folder"), command=lambda: self.open_folder(Path(self.var_import.get())))
        file_menu.add_separator()
        file_menu.add_command(label=self.t("menu_minimize_tray"), command=self.minimize_to_tray, state="normal" if self.var_resident.get() else "disabled")
        self.file_menu = file_menu
        file_menu.add_separator()
        file_menu.add_command(label=self.t("menu_exit"), command=self.confirm_exit)
        menu_bar.add_cascade(label=self.t("menu_file"), menu=file_menu)

        settings_menu = tk.Menu(menu_bar, tearoff=False)
        settings_menu.add_command(label=self.t("settings_ellipsis"), command=self.show_settings_dialog)
        settings_menu.add_separator()
        settings_menu.add_checkbutton(label=self.t("auto_detect_new_eml"), variable=self.var_watch, command=self.toggle_watcher)
        settings_menu.add_checkbutton(label=self.t("resident_tray"), variable=self.var_resident, command=self.on_resident_changed)
        menu_bar.add_cascade(label=self.t("menu_settings"), menu=settings_menu)

        language_menu = tk.Menu(menu_bar, tearoff=False)
        for code, native_name in SUPPORTED_LANGUAGES.items():
            language_menu.add_radiobutton(label=native_name, variable=self.var_language, value=code, command=lambda c=code: self.change_language(c))
        menu_bar.add_cascade(label=self.t("menu_language"), menu=language_menu)

        help_menu = tk.Menu(menu_bar, tearoff=False)
        help_menu.add_command(label=self.t("usage"), command=self.show_usage)
        help_menu.add_command(label=self.t("disclaimer"), command=self.show_disclaimer)
        help_menu.add_separator()
        help_menu.add_command(label=self.t("about"), command=self.show_about)
        menu_bar.add_cascade(label=self.t("menu_help"), menu=help_menu)
        self.root.config(menu=menu_bar)

    def show_usage(self):
        messagebox.showinfo(self.t("usage"), self.t("usage_text"))

    def show_disclaimer(self):
        messagebox.showinfo(self.t("disclaimer"), DISCLAIMER)

    def show_about(self):
        messagebox.showinfo(self.t("about"), about_text())

    def minimize_to_tray(self):
        if not self.var_resident.get():
            messagebox.showinfo(
                "常駐設定",
                "タスクトレイ常駐が無効です。\n"
                "［設定］－［タスクトレイに常駐する］を有効にしてください。",
            )
            return
        if self.tray_icon is None:
            self.start_tray_icon()
        self.persist_settings()
        self.root.withdraw()

    def confirm_exit(self):
        if messagebox.askyesno(
            self.t("exit_confirmation"),
            "INAS Mail Archiveを完全に終了しますか？\n"
            "タスクトレイ常駐とEMLの自動監視も終了します。",
        ):
            self.exit_application()

    def on_resident_changed(self):
        enabled = self.var_resident.get()
        if enabled:
            self.start_tray_icon()
            try:
                self.file_menu.entryconfig("タスクトレイに格納", state="normal")
            except tk.TclError:
                pass
            self.var_status.set("タスクトレイ常駐を有効にしました。")
        else:
            self.stop_tray_icon()
            try:
                self.file_menu.entryconfig("タスクトレイに格納", state="disabled")
            except tk.TclError:
                pass
            self.var_status.set(
                "タスクトレイ常駐を無効にしました。"
            )
        self.persist_settings()

    def build_ui(self):
        style = ttk.Style()
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass

        # 画面上部～設定項目は縦スクロール可能、状態表示と操作ボタンは常に下部へ固定する。
        root_frame = ttk.Frame(self.root)
        root_frame.pack(fill="both", expand=True)

        scroll_host = ttk.Frame(root_frame)
        scroll_host.pack(fill="both", expand=True)

        canvas = tk.Canvas(scroll_host, highlightthickness=0, borderwidth=0)
        scrollbar = ttk.Scrollbar(scroll_host, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        outer = ttk.Frame(canvas, padding=(14, 14, 10, 10))
        self._scroll_window = canvas.create_window((0, 0), window=outer, anchor="nw")
        self._scroll_canvas = canvas

        def update_scroll_region(_event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def fit_scroll_width(event):
            # スクロール領域の内容をCanvas幅に合わせ、横スクロールを発生させない。
            canvas.itemconfigure(self._scroll_window, width=event.width)

        outer.bind("<Configure>", update_scroll_region)
        canvas.bind("<Configure>", fit_scroll_width)

        def on_mousewheel(event):
            if canvas.bbox("all") is None:
                return
            first, last = canvas.yview()
            if first <= 0.0 and last >= 1.0:
                return
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        # メイン画面内の子ウィジェット上でもホイールを受け取れるよう、
        # CanvasのEnter/Leaveによるbind_allではなくメインToplevelへ直接バインドする。
        # これにより、Entry・Button・LabelFrame上でも上下スクロールでき、
        # 設定ダイアログなど別Toplevelのホイール処理とは干渉しない。
        self.root.bind("<MouseWheel>", on_mousewheel, add="+")

        ttk.Label(outer, text=APP_NAME, font=("Yu Gothic UI", 16, "bold")).pack(anchor="w")
        ttk.Label(outer, text=self.t("intro"), foreground="#444").pack(anchor="w", pady=(2, 12))

        watch = ttk.LabelFrame(outer, text=self.t("auto_import"), padding=10)
        watch.pack(fill="x")
        ttk.Entry(watch, textvariable=self.var_import).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(watch, text=self.t("change"), command=self.choose_import_folder).grid(row=0, column=1)
        ttk.Button(watch, text=self.t("open"), command=lambda: self.open_folder(Path(self.var_import.get()))).grid(row=0, column=2, padx=(6, 0))
        ttk.Checkbutton(watch, text=self.t("auto_detect_new_eml"), variable=self.var_watch, command=self.toggle_watcher).grid(row=1, column=0, columnspan=3, sticky="w", pady=(8, 0))
        monitor_row = ttk.Frame(watch)
        monitor_row.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        self.monitor_state_label = ttk.Label(monitor_row, textvariable=self.var_monitor_state)
        self.monitor_state_label.pack(side="left")
        ttk.Label(monitor_row, textvariable=self.var_last_detected, foreground="#555").pack(side="left", padx=(14, 0))
        ttk.Label(monitor_row, textvariable=self.var_pending_count, foreground="#555").pack(side="left", padx=(14, 0))
        ttk.Button(monitor_row, text=self.t("rescan"), command=self.rescan_import_folder).pack(side="right")
        watch.columnconfigure(0, weight=1)
        self.update_monitor_panel()

        info = ttk.LabelFrame(outer, text=self.t("loaded_mail"), padding=10)
        info.pack(fill="x", pady=10)
        ttk.Label(info, text=self.t("from_editable")).grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        self.from_entry = ttk.Entry(info, textvariable=self.var_from)
        self.from_entry.grid(row=0, column=1, sticky="ew", pady=4)
        ttk.Button(info, text=self.t("candidate"), command=self.apply_from_candidate).grid(
            row=0, column=2, padx=(6, 0), pady=4
        )
        ttk.Button(info, text=self.t("register"), command=self.register_current_sender).grid(
            row=0, column=3, padx=(6, 0), pady=4
        )
        ttk.Label(info, text=self.t("subject_editable")).grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(info, textvariable=self.var_subject).grid(row=1, column=1, sticky="ew", pady=4)
        ttk.Button(info, text=self.t("cleanup"), command=self.apply_subject_cleanup).grid(
            row=1, column=2, padx=(6, 0), pady=4
        )
        self.attachment_label = ttk.Label(info, text=self.t("attachment_zero"))
        self.attachment_label.grid(row=2, column=1, sticky="w", pady=4)
        info.columnconfigure(1, weight=1)

        output = ttk.LabelFrame(outer, text=self.t("output"), padding=10)
        output.pack(fill="x")
        ttk.Label(output, text=self.t("save_location")).grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(output, textvariable=self.var_output).grid(row=0, column=1, sticky="ew", pady=4)
        ttk.Button(output, text=self.t("browse"), command=self.choose_output_folder).grid(row=0, column=2, padx=(6, 0))
        self.add_labeled_entry(output, 1, self.t("pdf_filename"), self.var_pdf)
        self.add_labeled_entry(output, 2, self.t("folder_name"), self.var_folder)
        output.columnconfigure(1, weight=1)

        options = ttk.LabelFrame(outer, text=self.t("processing_settings"), padding=10)
        options.pack(fill="x", pady=(10, 0))
        post_action_row = ttk.Frame(options)
        post_action_row.grid(row=0, column=0, columnspan=2, sticky="w")
        values = [
            (self.t("move_to_folder"), "move_to_output"),
            (self.t("keep_in_import"), "keep"),
            (self.t("delete"), "delete"),
        ]
        for label, value in values:
            ttk.Radiobutton(
                post_action_row,
                text=label,
                value=value,
                variable=self.var_post_action,
            ).pack(side="left", padx=(0, 18))

        ttk.Checkbutton(options, text=self.t("open_after_save"), variable=self.var_open_after).grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(8, 0)
        )
        ttk.Button(
            options,
            text=self.t("detailed_settings"),
            command=self.show_settings_dialog,
        ).grid(row=2, column=0, sticky="w", pady=(8, 0))
        ttk.Label(
            options,
            text=self.t("details_hint"),
            foreground="#555",
        ).grid(row=2, column=1, sticky="w", padx=(8, 0), pady=(8, 0))
        options.columnconfigure(1, weight=1)

        # 下部はスクロールさせず常時表示する。
        footer = ttk.Frame(root_frame, padding=(14, 6, 14, 12))
        footer.pack(fill="x", side="bottom")
        ttk.Separator(footer, orient="horizontal").pack(fill="x", pady=(0, 7))
        ttk.Label(footer, textvariable=self.var_status, foreground="#2457a6", wraplength=660).pack(anchor="w", pady=(0, 7))

        # 通常操作と終了操作を分離する。
        # 1段目：日常的に使用する操作
        buttons = ttk.Frame(footer)
        buttons.pack(fill="x")
        ttk.Button(buttons, text=self.t("select_eml"), command=self.select_eml).pack(side="left")
        ttk.Button(buttons, text="未処理メールを開く", command=self.show_pending_mail_list).pack(side="left", padx=(6, 0))
        ttk.Button(buttons, text=self.t("history"), command=self.show_history).pack(side="left", padx=6)
        ttk.Button(buttons, text=self.t("settings"), command=self.show_settings_dialog).pack(side="left")
        ttk.Button(buttons, text=self.t("clear_input"), command=self.clear_form).pack(side="left", padx=(6, 0))
        ttk.Button(buttons, text=self.t("execute"), command=self.execute).pack(side="right")


    def add_labeled_entry(self, parent, row, label, variable, readonly=False):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=(0, 8), pady=4)
        entry = ttk.Entry(parent, textvariable=variable, state="readonly" if readonly else "normal")
        entry.grid(row=row, column=1, columnspan=2, sticky="ew", pady=4)
        return entry

    def choose_import_folder(self):
        folder = filedialog.askdirectory(initialdir=self.var_import.get())
        if folder:
            self.var_import.set(folder)
            Path(folder).mkdir(parents=True, exist_ok=True)
            self.restart_watcher()
            self.persist_settings()

    def choose_output_folder(self):
        current = self.var_output.get().strip()
        initial_dir = current if current and Path(current).is_dir() else str(Path.home())
        folder = filedialog.askdirectory(initialdir=initial_dir)
        if folder:
            self.var_output.set(folder)
            self.persist_settings()

    def select_eml(self):
        paths = filedialog.askopenfilenames(
            filetypes=[(self.t("eml_files"), "*.eml")],
            initialdir=self.var_import.get(),
        )
        if paths:
            self.queue_emls([Path(path) for path in paths])

    def drop_eml(self, event):
        paths = [Path(p) for p in self.root.tk.splitlist(event.data)]
        emls = [p for p in paths if p.is_file() and p.suffix.lower() == ".eml"]
        if emls:
            self.queue_emls(emls)
        else:
            messagebox.showerror("エラー", "有効なEMLファイルをドロップしてください。")

    @staticmethod
    def eml_signature(path: Path) -> tuple[int, int] | None:
        try:
            stat = path.stat()
            return (stat.st_size, stat.st_mtime_ns)
        except OSError:
            return None

    def mark_eml_seen(self, path: Path) -> None:
        """Record the current file state before showing a modal confirmation."""
        signature = self.eml_signature(path)
        if signature is not None:
            self.known_eml_signatures[path] = signature
            self.recently_detected_eml[path] = (signature, time.monotonic())

    def is_recent_duplicate_detection(self, path: Path, cooldown: float = 10.0) -> bool:
        signature = self.eml_signature(path)
        if signature is None:
            return False
        previous = self.recently_detected_eml.get(path)
        if not previous:
            return False
        old_signature, detected_at = previous
        if old_signature != signature:
            return False
        return (time.monotonic() - detected_at) < cooldown

    def queue_emls(self, paths: list[Path], auto_detected: bool = False):
        added = 0
        for path in paths:
            path = Path(path)
            if not path.exists() or path.suffix.lower() != ".eml":
                continue
            if self.mail_data and path == self.mail_data.source_path:
                continue
            if path in self.queued_paths:
                continue
            # A path that is currently in pending_paths has already been marked by
            # wait_until_stable; allow that first hand-off, but reject later repeats.
            if path not in self.pending_paths and self.is_recent_duplicate_detection(path):
                continue
            self.mark_eml_seen(path)
            self.eml_queue.append(path)
            self.queued_paths.add(path)
            added += 1
        # Power Automate / OneDrive監視からの自動検出時は、
        # メイン画面へ直接読み込まず「未処理メール一覧」を作業入口にする。
        if auto_detected and added and self.settings.get("auto_show_pending_on_new_eml", True):
            self.update_monitor_panel()
            self.show_window()
            self.root.after(100, self.show_pending_mail_list)
            return

        if not self.mail_data:
            self.load_next_queued_eml()
        elif added:
            self.var_status.set(f"現在のメールを確認中です。待機EML：{len(self.eml_queue)}件")
        self.update_monitor_panel()



    def validate_eml_file(self, path: Path) -> tuple[bool, str, dict]:
        """Lightweight EML health check before normal parsing."""
        try:
            path = Path(path)
            if not path.exists():
                return False, "ファイルが存在しません", {}
            size = path.stat().st_size
            if size <= 0:
                return False, "0バイトのEMLです", {}

            raw = path.read_bytes()
            # Real EML should have recognizable headers. Avoid rejecting unusual but valid mail too aggressively.
            if b":" not in raw[:65536]:
                return False, "メールヘッダーを確認できません", {}

            msg = BytesParser(policy=policy.default).parsebytes(raw)
            raw_from = dheader(msg.get("From")).strip()
            raw_subject = dheader(msg.get("Subject")).strip()
            raw_date = dheader(msg.get("Date")).strip()
            message_id = dheader(msg.get("Message-ID")).strip()
            _name, sender = parseaddr(raw_from)
            sender = (sender or "").strip().lower()

            # Require at least one normal mail header so arbitrary text renamed .eml is isolated.
            if not any((raw_from, raw_subject, raw_date, message_id, msg.get("To"))):
                return False, "主要メールヘッダーがありません", {}

            return True, "", {
                "sender": sender,
                "from": raw_from,
                "subject": raw_subject,
                "date": raw_date,
                "message_id": message_id.strip().lower(),
                "sha256": hashlib.sha256(raw).hexdigest(),
                "size": size,
            }
        except Exception as exc:
            return False, f"EML解析エラー: {exc}", {}

    def _quarantine_eml(self, path: Path, folder_name: str, reason: str, log_path: Path) -> Path | None:
        """Move an unusable/duplicate EML out of the monitored root so it cannot loop forever."""
        try:
            source = Path(path)
            import_dir = Path(self.var_import.get().strip())
            target_dir = import_dir / folder_name
            target_dir.mkdir(parents=True, exist_ok=True)
            target = unique_path(target_dir / source.name)
            shutil.move(str(source), str(target))
            sender = self.get_eml_sender_address(target)
            append_event_log(log_path, folder_name, target.name, sender, reason)
            logging.info("EMLを%sへ隔離: %s / %s", folder_name, target, reason)
            return target
        except Exception:
            logging.exception("EML隔離に失敗しました: %s", path)
            append_event_log(log_path, folder_name + "_MOVE_FAILED", Path(path).name, "", reason)
            return None

    def quarantine_error_eml(self, path: Path, reason: str) -> Path | None:
        return self._quarantine_eml(path, "Error", reason, ERROR_LOG_PATH)

    def quarantine_duplicate_eml(self, path: Path, reason: str) -> Path | None:
        return self._quarantine_eml(path, "Duplicate", reason, DUPLICATE_LOG_PATH)

    def is_processed_duplicate_path(self, path: Path, health: dict | None = None) -> tuple[bool, str]:
        """Message-ID first, SHA-256 second. Only compares against successfully processed mail."""
        if not self.settings.get("duplicate_check_enabled", True):
            return False, ""
        info = health or {}
        try:
            if not info:
                ok, _reason, info = self.validate_eml_file(path)
                if not ok:
                    return False, ""
            message_id = str(info.get("message_id", "") or "").strip().lower()
            if message_id:
                registry = load_message_id_registry()
                record = registry.get(message_id)
                if isinstance(record, dict):
                    return True, f"Message-ID重複 / 前回保存先: {record.get('output_folder', '')}"
            source_hash = str(info.get("sha256", "") or "")
            if source_hash and source_hash in self.processed_history:
                record = self.processed_history.get(source_hash, {})
                return True, f"SHA-256重複 / 前回処理: {record.get('processed_at', '')}"
        except Exception:
            logging.exception("事前重複判定に失敗しました: %s", path)
        return False, ""

    def preflight_eml(self, path: Path, quarantine: bool = True) -> tuple[bool, str]:
        """Health -> exclusion -> duplicate. Returns True only for normal pending mail."""
        ok, reason, health = self.validate_eml_file(path)
        if not ok:
            if quarantine and Path(path).exists():
                self.quarantine_error_eml(path, reason)
            return False, f"エラーEML: {reason}"

        excluded, excluded_reason = self.is_excluded_eml(path)
        if excluded:
            sender = str(health.get("sender", "") or "")
            append_event_log(EXCLUDED_LOG_PATH, "EXCLUDED", Path(path).name, sender, excluded_reason)
            if self.settings.get("delete_excluded_temp_eml", True):
                try:
                    Path(path).unlink()
                except Exception:
                    logging.exception("保存対象外EMLの削除に失敗: %s", path)
            return False, excluded_reason

        duplicate, duplicate_reason = self.is_processed_duplicate_path(path, health)
        if duplicate:
            if quarantine and Path(path).exists():
                self.quarantine_duplicate_eml(path, duplicate_reason)
            return False, duplicate_reason

        return True, ""

    def _normalized_excluded_addresses(self) -> set[str]:
        values = self.settings.get("excluded_addresses", [])
        if isinstance(values, str):
            values = values.splitlines()
        return {str(v).strip().lower() for v in (values or []) if str(v).strip()}

    def _normalized_excluded_domains(self) -> set[str]:
        values = self.settings.get("excluded_domains", [])
        if isinstance(values, str):
            values = values.splitlines()
        result = set()
        for v in values or []:
            s = str(v).strip().lower()
            if not s:
                continue
            if s.startswith("@"):
                s = s[1:]
            result.add(s)
        return result

    def get_eml_sender_address(self, path: Path) -> str:
        """Read only the From header and return a normalized sender address."""
        try:
            with Path(path).open("rb") as fh:
                msg = BytesParser(policy=policy.default).parse(fh, headersonly=True)
            raw_from = dheader(msg.get("From")).strip()
            _name, address = parseaddr(raw_from)
            return (address or "").strip().lower()
        except Exception:
            logging.exception("除外判定用From取得に失敗しました: %s", path)
            return ""

    def is_excluded_eml(self, path: Path) -> tuple[bool, str]:
        address = self.get_eml_sender_address(path)
        if not address:
            return False, ""

        if address in self._normalized_excluded_addresses():
            return True, f"除外メールアドレス: {address}"

        if "@" in address:
            domain = address.rsplit("@", 1)[1].lower()
            if domain in self._normalized_excluded_domains():
                return True, f"除外ドメイン: {domain}"

        return False, ""

    def apply_exclusion_to_pending_emls(self, update_ui: bool = True) -> dict:
        """Apply current exclusion settings to EMLs already in the import folder."""
        result = {"checked": 0, "excluded": 0, "deleted": 0, "failed": 0}
        try:
            import_dir = Path(self.var_import.get().strip())
        except Exception:
            return result
        if not import_dir.exists():
            return result

        delete_excluded = bool(self.settings.get("delete_excluded_temp_eml", True))

        for p in list(import_dir.glob("*.eml")):
            result["checked"] += 1
            try:
                excluded, reason = self.is_excluded_eml(p)
                if not excluded:
                    continue
                result["excluded"] += 1
                logging.info("保存対象外EMLを検出: %s (%s)", p.name, reason)
                append_event_log(EXCLUDED_LOG_PATH, "EXCLUDED", p.name, self.get_eml_sender_address(p), reason)
                if delete_excluded:
                    try:
                        p.unlink()
                        result["deleted"] += 1
                        logging.info("保存対象外の一時EMLを削除: %s", p)
                    except Exception:
                        result["failed"] += 1
                        logging.exception("保存対象外EMLの削除に失敗: %s", p)
            except Exception:
                result["failed"] += 1
                logging.exception("保存対象外EMLの再判定に失敗: %s", p)

        if update_ui:
            try:
                self.update_monitor_panel()
            except Exception:
                pass
            try:
                dlg = getattr(self, "_pending_dialog", None)
                if dlg and dlg.winfo_exists():
                    dlg.destroy()
                    self._pending_dialog = None
                    self.root.after(50, self.show_pending_mail_list)
            except Exception:
                pass
        return result

    def show_pending_mail_list(self):
        """Show all currently unprocessed EML files in one list.

        This is used especially at startup when Power Automate has accumulated
        multiple EML files overnight.  No EML is deleted unless the user
        explicitly processes it or marks it as unnecessary.
        """
        try:
            import_dir = Path(self.var_import.get().strip())
        except Exception:
            return
        if not import_dir.exists():
            return

        paths = []
        for p in sorted(import_dir.glob("*.eml"), key=lambda x: x.stat().st_mtime):
            try:
                accepted, preflight_reason = self.preflight_eml(p, quarantine=True)
                if not accepted:
                    logging.info("未処理一覧から除外: %s (%s)", p.name, preflight_reason)
                    continue
                paths.append(p)
            except Exception:
                logging.exception("未処理一覧の事前確認に失敗: %s", p)
                if p.exists():
                    self.quarantine_error_eml(p, "未処理一覧の事前確認例外")

        if not paths:
            self.var_status.set("未処理EMLはありません")
            return

        if getattr(self, "_pending_dialog", None):
            try:
                if self._pending_dialog.winfo_exists():
                    refresh_cb = getattr(self, "_pending_refresh_callback", None)
                    if callable(refresh_cb):
                        refresh_cb()
                    self._pending_dialog.deiconify()
                    self._pending_dialog.lift()
                    try:
                        self._pending_dialog.focus_force()
                    except Exception:
                        pass
                    return
            except Exception:
                pass

        dialog = tk.Toplevel(self.root)
        self._pending_dialog = dialog
        dialog.title(f"未処理メール一覧（{len(paths)}件）")
        dialog.geometry("980x560")
        dialog.minsize(820, 420)
        dialog.transient(self.root)

        outer = ttk.Frame(dialog, padding=10)
        outer.pack(fill="both", expand=True)

        ttk.Label(
            outer,
            text="未処理メールを選択して内容を確認します。保存または保存不要の処理後、この一覧へ戻ります。"
        ).pack(anchor="w", pady=(0, 8))

        tree_frame = ttk.Frame(outer)
        tree_frame.pack(fill="both", expand=True)

        cols = ("datetime", "sender", "subject", "attachments", "file")
        tree = ttk.Treeview(tree_frame, columns=cols, show="headings", selectmode="extended")
        for c, title, width in [
            ("datetime", "受信日時", 145),
            ("sender", "送信者", 220),
            ("subject", "件名", 360),
            ("attachments", "添付", 60),
            ("file", "EMLファイル", 170),
        ]:
            tree.heading(c, text=title)
            tree.column(c, width=width, anchor="w" if c != "attachments" else "center")

        vs = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        hs = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        row_paths = {}
        summary_cache = {}

        def read_summary(p: Path):
            """未処理一覧表示。取得失敗時は直前の正常値を保持する。"""
            cache_key = str(p.resolve()).lower()
            try:
                summary = read_eml_summary(p)
                row = (
                    str(summary.get("date") or "（不明）"),
                    str(summary.get("sender") or "（差出人不明）"),
                    str(summary.get("subject") or "（件名なし）"),
                    int(summary.get("attachments") or 0),
                )
                summary_cache[cache_key] = row
                return row
            except Exception:
                logging.exception("未処理一覧のEML概要取得に失敗しました: %s", p)
                cached = summary_cache.get(cache_key)
                if cached:
                    return cached
                # parse_emlが読める場合は最低限From/SubjectをMailDataから補完
                try:
                    data = parse_eml(p)
                    sender = (data.from_name or "").strip() or "（差出人不明）"
                    subject = (self.extract_subject(data.header_block) or "").strip() or "（件名なし）"
                    row = ("（不明）", sender, subject, len(data.attachments))
                    summary_cache[cache_key] = row
                    return row
                except Exception:
                    return "（不明）", "（取得失敗）", p.stem, 0

        def refresh():
            for item in tree.get_children():
                tree.delete(item)
            row_paths.clear()
            current = []
            try:
                current_all = sorted(import_dir.glob("*.eml"), key=lambda x: x.stat().st_mtime)
            except Exception:
                current_all = []

            current = []
            for p in current_all:
                try:
                    accepted, preflight_reason = self.preflight_eml(p, quarantine=True)
                    if not accepted:
                        logging.info("未処理一覧更新時に除外: %s (%s)", p.name, preflight_reason)
                        continue
                    current.append(p)
                except Exception:
                    logging.exception("未処理一覧更新時の事前確認に失敗: %s", p)
                    if p.exists():
                        self.quarantine_error_eml(p, "未処理一覧更新の事前確認例外")

            for idx, p in enumerate(current):
                dt, sender, subject, att = read_summary(p)
                iid = str(idx)
                row_paths[iid] = p
                tree.insert("", "end", iid=iid, values=(dt, sender, subject, att, p.name))
            dialog.title(f"未処理メール一覧（{len(current)}件）")
            self.var_status.set(f"未処理EML：{len(current)}件")

        # 保存後に同じ未処理一覧へ戻るため、外部から再読込できるよう保持する。
        self._pending_refresh_callback = refresh

        def open_selected(_event=None):
            sel = tree.selection()
            if len(sel) != 1:
                messagebox.showinfo("未処理メール", "確認するメールを1件選択してください。", parent=dialog)
                return
            p = row_paths.get(sel[0])
            if not p or not p.exists():
                refresh()
                return

            handoff_to_main = False
            try:
                data = parse_eml(p)
                self.mail_data = data
                self.current_eml_path = p

                summary = read_eml_summary(p)
                raw_from = str(summary.get("sender") or data.from_name or "").strip()
                raw_subject = str(
                    summary.get("subject")
                    or self.extract_subject(data.header_block)
                    or ""
                ).strip()

                key = sender_key(data.from_name)
                mapped_name = self.sender_dictionary.get(key)
                self.var_from.set(mapped_name if mapped_name else raw_from)
                self.var_subject.set(raw_subject)
                self.attachment_label.config(text=f"添付：{len(data.attachments)}件")
                self.update_names()

                # 詳細確認中は未処理一覧を隠す。
                try:
                    dialog.withdraw()
                except Exception:
                    pass

                action = self.show_mail_preview_dialog(data, p)

                if action == "save":
                    # ここでは保存実行しない。保存先・添付選択・メール情報を
                    # メイン画面へ引き渡し、ユーザーが［実行］で最終保存する。
                    self.update_names()
                    self._return_to_pending_after_execute = True
                    self._pending_source_path = p
                    handoff_to_main = True
                    self.var_status.set("未処理メールをメイン画面へ反映しました。内容を確認して［実行］してください。")
                    self.root.deiconify()
                    self.root.lift()
                    try:
                        self.root.attributes("-topmost", True)
                        self.root.after(400, lambda: self.root.attributes("-topmost", False))
                    except Exception:
                        pass

                elif action == "discard":
                    # 保存不要：一時EMLのみ削除。Outlook側の元メールは残す。
                    try:
                        if p.exists():
                            p.unlink()
                    except Exception:
                        logging.exception("保存不要EMLの削除に失敗しました: %s", p)
                    self.clear_form(load_next=False)

                else:
                    # ×で閉じた場合は未処理のまま残す。
                    self.mail_data = None
                    self.current_eml_path = None
                    self.var_from.set("")
                    self.var_subject.set("")
                    self.attachment_label.config(text="添付：0件")
                    self.update_monitor_panel()

            except Exception as exc:
                logging.exception("未処理メールを開けませんでした: %s", p)
                messagebox.showerror(
                    "未処理メール",
                    f"メールを開けませんでした。\n\n{p.name}\n\n詳細：{exc}",
                    parent=self.root,
                )
            finally:
                # メイン画面へ引き渡した場合は未処理一覧を隠したままにする。
                # ［実行］完了後に execute() から再表示する。
                if not handoff_to_main:
                    try:
                        refresh()
                        remaining = len(list(import_dir.glob("*.eml"))) if import_dir.exists() else 0
                        if remaining > 0 and dialog.winfo_exists():
                            dialog.deiconify()
                            dialog.lift()
                            try:
                                dialog.focus_force()
                            except Exception:
                                pass
                        elif remaining == 0:
                            self._pending_dialog = None
                            self._pending_refresh_callback = None
                            if dialog.winfo_exists():
                                dialog.destroy()
                            self.root.lift()
                            self.update_monitor_panel()
                    except Exception:
                        pass

        def discard_selected():
            sel = list(tree.selection())
            if not sel:
                messagebox.showinfo("未処理メール", "保存不要にするメールを選択してください。", parent=dialog)
                return
            if not messagebox.askyesno(
                "まとめて保存不要",
                f"選択した {len(sel)} 件を保存不要として、一時EMLを削除しますか？\n\nOutlook側の元メールは削除されません。",
                parent=dialog,
            ):
                return
            failed = []
            for iid in sel:
                p = row_paths.get(iid)
                if not p:
                    continue
                try:
                    if p.exists() and p.resolve().parent == import_dir.resolve():
                        p.unlink()
                except Exception:
                    failed.append(p.name)
            refresh()
            try:
                remaining = len(list(import_dir.glob("*.eml"))) if import_dir.exists() else 0
                if remaining == 0:
                    self._pending_dialog = None
                    dialog.destroy()
                    self.root.lift()
                    self.update_monitor_panel()
                    return
            except Exception:
                pass
            if failed:
                messagebox.showwarning(
                    "まとめて保存不要",
                    "一部のEMLを削除できませんでした。\n\n" + "\n".join(failed),
                    parent=dialog,
                )

        buttons = ttk.Frame(outer)
        buttons.pack(fill="x", pady=(8, 0))
        ttk.Button(buttons, text="内容を確認", command=open_selected).pack(side="left")
        ttk.Button(buttons, text="選択を保存不要", command=discard_selected).pack(side="left", padx=6)
        ttk.Button(buttons, text="再スキャン", command=refresh).pack(side="left")
        ttk.Button(buttons, text="閉じる", command=dialog.destroy).pack(side="right")

        tree.bind("<Double-1>", open_selected)
        refresh()

        def on_close():
            try:
                self._pending_dialog = None
                self._pending_refresh_callback = None
            except Exception:
                pass
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", on_close)

    def load_next_queued_eml(self):
        # 複数EMLは一覧画面から処理し、確認画面を連続表示しない。
        try:
            import_dir = Path(self.var_import.get().strip())
            pending = list(import_dir.glob("*.eml")) if import_dir.exists() else []
            if len(pending) > 1 and not getattr(self, "_processing_one_from_pending", False):
                self.root.after(50, self.show_pending_mail_list)
                return
        except Exception:
            pass
        while self.eml_queue and not self.mail_data:
            path = self.eml_queue.pop(0)
            self.queued_paths.discard(path)
            if path.exists():
                self.load_eml(path)
                break
        self.update_monitor_panel()

    def apply_from_candidate(self):
        """From欄の最初の半角・全角スペース以降を削除する。"""
        current = self.var_from.get()
        match = re.search(r"[ \u3000]", current)
        if not match:
            return
        candidate = current[:match.start()].strip()
        self.var_from.set(candidate)
        self.from_entry.icursor(tk.END)
        self.from_entry.focus_set()

    def apply_subject_cleanup(self):
        cleaned = clean_subject(self.var_subject.get())
        self.var_subject.set(cleaned)
        self.var_status.set("件名の Re:/Fw:/Fwd: を整理しました。")

    def register_current_sender(self):
        if not self.mail_data:
            messagebox.showinfo("差出人登録", "先にEMLを読み込んでください。")
            return
        value = self.var_from.get().strip()
        key = sender_key(self.mail_data.from_name)
        if not key or not value:
            messagebox.showerror("差出人登録", "登録する差出人名が空欄です。")
            return
        self.sender_dictionary[key] = value
        save_sender_dictionary(self.sender_dictionary)
        rules = dict(self.settings.get("sender_honorific_rules", {}) or {})
        rules.setdefault(key, {"mode": "standard", "custom": ""})
        self.settings["sender_honorific_rules"] = rules
        save_settings(self.settings)
        self.var_status.set(f"差出人設定へ登録しました：{key} → {value}")


    def get_sender_email_address(self) -> str:
        """Return the current sender's normalized email address."""
        # mail_dataがある場合は元Fromを優先する。
        try:
            if self.mail_data:
                key = sender_key(self.mail_data.from_name)
                if key:
                    return key
        except Exception:
            pass
        try:
            _name, address = parseaddr(self.var_from.get().strip())
            return (address or "").strip().lower()
        except Exception:
            return ""

    def get_sender_honorific_rule(self, sender_address: str | None = None) -> dict:
        key = (sender_address or self.get_sender_email_address() or "").strip().lower()
        rules = self.settings.get("sender_honorific_rules", {})
        if not isinstance(rules, dict):
            return {}
        row = rules.get(key, {})
        return row if isinstance(row, dict) else {}

    def current_sender_honorific(self, address: str | None = None) -> str:
        """Resolve sender-specific honorific; otherwise use the standard honorific."""
        key = (address or self.get_sender_email_address() or "").strip().lower()

        # Ver.1.6.22～1.6.33の自分メール設定は互換移行として「敬称なし」扱い。
        old_self = self.settings.get("self_email_addresses", [])
        if isinstance(old_self, str):
            old_self = [old_self]
        if key and key in {str(x).strip().lower() for x in old_self if str(x).strip()}:
            return ""

        row = self.get_sender_honorific_rule(key)
        mode = str(row.get("mode", "standard") or "standard")
        custom = str(row.get("custom", "") or "")

        if mode == "none":
            return ""
        if mode == "shi":
            return "氏"
        if mode == "sama":
            return "様"
        if mode == "custom":
            return custom
        return str(self.settings.get("sender_honorific", "氏") or "")

    def show_settings_dialog(self, initial_tab: str = "基本設定"):
        """主要設定を1つのタブ画面へ集約する。"""
        dialog = tk.Toplevel(self.root)
        dialog.title(self.t("settings"))
        dialog.geometry("760x560")
        dialog.minsize(900, 700)
        dialog.minsize(680, 500)
        dialog.transient(self.root)
        dialog.grab_set()

        # 設定画面は、入力欄・一覧だけを白背景とし、余白部分は淡いグレーで統一する。
        SETTINGS_BG = "#F3F4F6"
        SETTINGS_BORDER = "#D8DEE8"

        dialog.configure(bg=SETTINGS_BG)
        style = ttk.Style(dialog)
        style.configure("Settings.TFrame", background=SETTINGS_BG)
        style.configure("Settings.TEntry", padding=(6, 4))
        style.configure("Settings.TButton", padding=(12, 4))
        style.configure("Settings.TCheckbutton", background=SETTINGS_BG)
        style.configure("Settings.TLabel", background=SETTINGS_BG)

        frame = tk.Frame(dialog, bg=SETTINGS_BG, padx=12, pady=12, bd=0, highlightthickness=0)
        frame.pack(fill="both", expand=True)

        # 設定タブは ttk.Notebook の標準表示では選択状態が分かりにくいため、
        # INAS UIとして選択中／未選択を明確に区別するカスタムタブバーを使用する。
        tab_bar = tk.Frame(frame, bg="#E5E9EF", bd=0, highlightthickness=0)
        tab_bar.pack(fill="x", pady=(0, 0))
        tab_body = tk.Frame(frame, bg=SETTINGS_BG, bd=0, relief="flat", highlightbackground=SETTINGS_BORDER, highlightthickness=1)
        tab_body.pack(fill="both", expand=True)

        tab_frames = {}
        tab_buttons = {}
        tab_canvases = {}
        current_tab = {"name": None}

        def show_settings_tab(name):
            for tab_name, tab_frame in tab_frames.items():
                if tab_name == name:
                    tab_frame.pack(fill="both", expand=True)
                else:
                    tab_frame.pack_forget()
            for tab_name, button in tab_buttons.items():
                selected = tab_name == name
                button.configure(
                    bg="#FFFFFF" if selected else "#F2F4F7",
                    fg="#1F4E8C" if selected else "#5F6368",
                    activebackground="#FFFFFF" if selected else "#E7EBF0",
                    activeforeground="#1F4E8C" if selected else "#333333",
                    font=("Yu Gothic UI", 9, "bold" if selected else "normal"),
                    relief="flat",
                    bd=0,
                )
                accent = getattr(button, "_inas_accent", None)
                if accent is not None:
                    accent.configure(bg="#2F6FB3" if selected else "#F2F4F7")
            current_tab["name"] = name

        def add_settings_tab(name):
            holder = tk.Frame(tab_bar, bg="#F2F4F7", bd=0, highlightthickness=0)
            holder.pack(side="left", padx=(0, 1))
            button = tk.Button(
                holder, text=name, command=lambda n=name: show_settings_tab(n),
                padx=14, pady=8, cursor="hand2", takefocus=True,
                bg="#F2F4F7", fg="#5F6368", activebackground="#E7EBF0",
                activeforeground="#333333", relief="flat", bd=0,
                font=("Yu Gothic UI", 9),
            )
            button.pack(fill="x")
            accent = tk.Frame(holder, height=3, bg="#F2F4F7", bd=0, highlightthickness=0)
            accent.pack(fill="x")
            button._inas_accent = accent
            tab_buttons[name] = button

            # 各設定タブの内容は独立したスクロール領域にする。
            # タブバーと下部の保存／キャンセルは固定し、内容だけを縦スクロールする。
            page = tk.Frame(tab_body, bg=SETTINGS_BG, bd=0, highlightthickness=0)
            page.rowconfigure(0, weight=1)
            page.columnconfigure(0, weight=1)

            canvas = tk.Canvas(page, bg=SETTINGS_BG, highlightthickness=0, bd=0)
            scrollbar = ttk.Scrollbar(page, orient="vertical", command=canvas.yview)
            canvas.configure(yscrollcommand=scrollbar.set)
            canvas.grid(row=0, column=0, sticky="nsew")
            scrollbar.grid(row=0, column=1, sticky="ns")

            body = tk.Frame(canvas, bg=SETTINGS_BG, padx=14, pady=14, bd=0, highlightthickness=0)
            body_window = canvas.create_window((0, 0), window=body, anchor="nw")

            def update_scrollregion(_event=None, c=canvas):
                c.configure(scrollregion=c.bbox("all"))

            def fit_body_size(event, c=canvas, window_id=body_window, b=body):
                # 内容が少ないタブでも設定領域の下まで背景色を埋める。
                # 内容が多い場合は要求高さを維持し、縦スクロールできるようにする。
                c.itemconfigure(
                    window_id,
                    width=max(1, event.width),
                    height=max(1, event.height, b.winfo_reqheight()),
                )
                c.after_idle(update_scrollregion)

            body.bind("<Configure>", update_scrollregion)
            canvas.bind("<Configure>", fit_body_size)

            tab_frames[name] = page
            tab_canvases[name] = canvas
            return body

        def scroll_settings(event):
            """現在表示中の設定タブをマウスホイールで縦スクロールする。"""
            # Treeview は一覧自体のホイール操作を優先する。
            widget = event.widget
            if isinstance(widget, ttk.Treeview):
                return
            name = current_tab.get("name")
            canvas = tab_canvases.get(name)
            if canvas is None:
                return
            try:
                if canvas.bbox("all") is None:
                    return
                content_height = canvas.bbox("all")[3]
                if content_height <= canvas.winfo_height():
                    return
                delta = int(-1 * (event.delta / 120)) if event.delta else 0
                if delta:
                    canvas.yview_scroll(delta, "units")
                    return "break"
            except tk.TclError:
                return

        # Windowsで設定ダイアログ上のどこにマウスがあってもスクロールできる。
        dialog.bind("<MouseWheel>", scroll_settings, add="+")

        # --- 基本設定 ---
        basic = add_settings_tab(self.t("tab_basic"))
        basic.columnconfigure(1, weight=1)

        import_var = tk.StringVar(value=self.var_import.get())
        watch_var = tk.BooleanVar(value=self.var_watch.get())
        resident_var = tk.BooleanVar(value=self.var_resident.get())
        auto_start_var = tk.BooleanVar(value=self.var_auto_start.get())
        ignore_inline_var = tk.BooleanVar(value=self.var_ignore_inline.get())

        ttk.Label(basic, text=self.t("import_folder"), style="Settings.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=5)
        ttk.Entry(basic, textvariable=import_var, style="Settings.TEntry").grid(row=0, column=1, sticky="ew", pady=5)

        def choose_import():
            folder = filedialog.askdirectory(initialdir=import_var.get() or str(Path.home()), parent=dialog)
            if folder:
                import_var.set(folder)

        ttk.Button(basic, text=self.t("browse"), command=choose_import, style="Settings.TButton").grid(row=0, column=2, padx=(6, 0), pady=5)
        ttk.Checkbutton(basic, text=self.t("auto_detect_new_eml"), variable=watch_var, style="Settings.TCheckbutton").grid(row=1, column=0, columnspan=3, sticky="w", pady=(10, 2))
        ttk.Checkbutton(basic, text=self.t("resident_tray"), variable=resident_var, style="Settings.TCheckbutton").grid(row=2, column=0, columnspan=3, sticky="w", pady=2)
        ttk.Checkbutton(basic, text=self.t("auto_start_windows"), variable=auto_start_var, style="Settings.TCheckbutton").grid(row=3, column=0, columnspan=3, sticky="w", pady=2)
        ttk.Checkbutton(basic, text=self.t("ignore_inline_images"), variable=ignore_inline_var, style="Settings.TCheckbutton").grid(row=4, column=0, columnspan=3, sticky="w", pady=2)
        ttk.Label(
            basic,
            text=self.t("auto_start_note"),
            foreground="#555",
            style="Settings.TLabel",
        ).grid(row=5, column=0, columnspan=3, sticky="w", pady=(10, 0))

        # --- 自動取込・保存 ---
        auto_save = add_settings_tab("自動取込・保存")
        auto_save.columnconfigure(1, weight=1)
        default_output_var = tk.StringVar(value=str(self.settings.get("default_output_folder", DEFAULT_OUTPUT)))
        confirm_auto_var = tk.BooleanVar(value=bool(self.settings.get("confirm_auto_detected_mail", True)))
        delete_excluded_var = tk.BooleanVar(value=bool(self.settings.get("delete_excluded_temp_eml", True)))

        ttk.Label(auto_save, text="標準保存先", style="Settings.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=5)
        ttk.Entry(auto_save, textvariable=default_output_var, style="Settings.TEntry").grid(row=0, column=1, sticky="ew", pady=5)
        def choose_default_output():
            folder = filedialog.askdirectory(initialdir=default_output_var.get() or str(Path.home()), parent=dialog)
            if folder:
                default_output_var.set(folder)
        ttk.Button(auto_save, text=self.t("browse"), command=choose_default_output, style="Settings.TButton").grid(row=0, column=2, padx=(6, 0), pady=5)
        ttk.Label(auto_save, text="個別の保存先が特定できない場合、この保存先を候補として表示します。", wraplength=620, style="Settings.TLabel").grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 10))
        ttk.Checkbutton(auto_save, text="自動検出したメールは、最初に保存対象か確認する", variable=confirm_auto_var, style="Settings.TCheckbutton").grid(row=2, column=0, columnspan=3, sticky="w", pady=2)
        ttk.Checkbutton(auto_save, text="保存対象外メールの一時EMLを取込フォルダから削除する", variable=delete_excluded_var, style="Settings.TCheckbutton").grid(row=3, column=0, columnspan=3, sticky="w", pady=2)

        ttk.Label(auto_save, text="保存対象外メールアドレス（1行1件）", style="Settings.TLabel").grid(row=4, column=0, columnspan=3, sticky="w", pady=(14, 4))
        excluded_addresses_text = tk.Text(auto_save, height=6, wrap="none")
        excluded_addresses_text.grid(row=5, column=0, columnspan=3, sticky="ew")
        excluded_addresses_text.insert("1.0", "\n".join(self.settings.get("excluded_addresses", [])))

        ttk.Label(auto_save, text="保存対象外ドメイン（例：example.com / @example.com、1行1件）", style="Settings.TLabel").grid(row=6, column=0, columnspan=3, sticky="w", pady=(12, 4))
        excluded_domains_text = tk.Text(auto_save, height=6, wrap="none")
        excluded_domains_text.grid(row=7, column=0, columnspan=3, sticky="ew")
        excluded_domains_text.insert("1.0", "\n".join(self.settings.get("excluded_domains", [])))

        ttk.Separator(auto_save, orient="horizontal").grid(row=7, column=0, columnspan=3, sticky="ew", pady=(12, 8))
        ttk.Label(auto_save, text="よく使う保存先", style="Settings.TLabel").grid(row=8, column=0, sticky="w", pady=5)
        ttk.Label(
            auto_save,
            text="登録済みの保存先は、メール内容確認画面からすぐに選択できます。",
            wraplength=520,
            style="Settings.TLabel",
        ).grid(row=8, column=1, sticky="w", pady=5)
        ttk.Button(
            auto_save,
            text="登録・管理...",
            command=lambda: self.show_registered_save_locations(parent=dialog),
            style="Settings.TButton",
        ).grid(row=8, column=2, padx=(6, 0), pady=5)

        duplicate_check_var = tk.BooleanVar(value=bool(self.settings.get("duplicate_check_enabled", True)))
        recent_dest_var = tk.BooleanVar(value=bool(self.settings.get("recent_destination_suggestions", True)))
        selective_att_var = tk.BooleanVar(value=bool(self.settings.get("attachment_selective_save", True)))
        external_link_var = tk.BooleanVar(value=bool(self.settings.get("external_link_detection_enabled", True)))
        auto_show_pending_var = tk.BooleanVar(value=bool(self.settings.get("auto_show_pending_on_new_eml", True)))

        ttk.Separator(auto_save, orient="horizontal").grid(row=9, column=0, columnspan=3, sticky="ew", pady=(12, 8))
        ttk.Checkbutton(
            auto_save,
            text="重複保存チェックを有効にする",
            variable=duplicate_check_var,
        ).grid(row=10, column=0, columnspan=2, sticky="w", pady=4)
        ttk.Checkbutton(
            auto_save,
            text="履歴から最近の保存先候補を表示する",
            variable=recent_dest_var,
        ).grid(row=11, column=0, columnspan=2, sticky="w", pady=4)
        ttk.Checkbutton(
            auto_save,
            text="添付ファイルを個別に保存選択できるようにする",
            variable=selective_att_var,
        ).grid(row=12, column=0, columnspan=2, sticky="w", pady=4)
        ttk.Checkbutton(
            auto_save,
            text="外部ファイル便・ダウンロードURLを自動検出する",
            variable=external_link_var,
        ).grid(row=13, column=0, columnspan=2, sticky="w", pady=4)
        ttk.Checkbutton(
            auto_save,
            text="新着EML検出時に未処理メール一覧を自動表示する",
            variable=auto_show_pending_var,
        ).grid(row=14, column=0, columnspan=3, sticky="w", pady=4)


        # --- PDF出力 ---
        pdf_output = add_settings_tab(self.t("tab_pdf"))
        pdf_lang_frame = ttk.Frame(pdf_output)
        pdf_lang_frame.pack(fill="x", pady=(0, 10))
        ttk.Label(pdf_lang_frame, text=self.t("pdf_language")).pack(side="left")
        pdf_language_var = tk.StringVar(value=self.settings.get("pdf_language", "same"))
        pdf_lang_combo = ttk.Combobox(pdf_lang_frame, state="readonly", width=24)
        pdf_lang_values = [self.t("same_as_app"), "日本語", "English", "Tiếng Việt"]
        pdf_lang_codes = ["same", "ja", "en", "vi"]
        pdf_lang_combo["values"] = pdf_lang_values
        try:
            pdf_lang_combo.current(pdf_lang_codes.index(pdf_language_var.get()))
        except ValueError:
            pdf_lang_combo.current(0)
        def _sync_pdf_lang(_event=None):
            idx = pdf_lang_combo.current()
            pdf_language_var.set(pdf_lang_codes[idx] if 0 <= idx < len(pdf_lang_codes) else "same")
        pdf_lang_combo.bind("<<ComboboxSelected>>", _sync_pdf_lang)
        pdf_lang_combo.pack(side="left", padx=(8, 0))

        ttk.Label(
            pdf_output,
            text=self.t("pdf_select_note"),
            wraplength=650,
        ).pack(anchor="w", pady=(0, 12))

        pdf_option_vars = {
            key: tk.BooleanVar(value=bool(self.settings.get(key, default)))
            for key, default in PDF_DISPLAY_DEFAULTS.items()
        }

        content_box = ttk.LabelFrame(pdf_output, text=self.t("mail_content_basic"), padding=10)
        content_box.pack(fill="x", pady=(0, 10))
        for text, key in (
            (self.t("subject"), "pdf_show_subject"),
            (self.t("from_original"), "pdf_show_from"),
            (self.t("to_recipient"), "pdf_show_to"),
            ("CC", "pdf_show_cc"),
            ("BCC", "pdf_show_bcc"),
            (self.t("sent_datetime"), "pdf_show_sent_datetime"),
            (self.t("attachment_list"), "pdf_show_attachments"),
            (self.t("body"), "pdf_show_body"),
        ):
            ttk.Checkbutton(content_box, text=text, variable=pdf_option_vars[key]).pack(anchor="w", pady=2)

        detail_box = ttk.LabelFrame(pdf_output, text=self.t("detail_info"), padding=10)
        detail_box.pack(fill="x", pady=(0, 10))
        for text, key in (
            (self.t("sender_display_name"), "pdf_show_sender_display_name"),
            (self.t("reply_to"), "pdf_show_reply_to"),
            ("Message-ID", "pdf_show_message_id"),
            (self.t("importance"), "pdf_show_importance"),
            (self.t("attachment_count_label"), "pdf_show_attachment_count"),
            (self.t("attachment_size"), "pdf_show_attachment_size"),
        ):
            ttk.Checkbutton(detail_box, text=text, variable=pdf_option_vars[key]).pack(anchor="w", pady=2)
        ttk.Label(
            detail_box,
            text=self.t("attachment_size_note"),
            foreground="#666",
            wraplength=620,
        ).pack(anchor="w", pady=(6, 0))

        page_box = ttk.LabelFrame(pdf_output, text=self.t("page_display"), padding=10)
        page_box.pack(fill="x")
        for text, key in (
            (self.t("header"), "pdf_show_header"),
            (self.t("footer_separator"), "pdf_show_footer"),
            (self.t("page_number"), "pdf_show_page_number"),
            (self.t("app_name_display"), "pdf_show_app_name"),
            (self.t("saved_datetime"), "pdf_show_saved_datetime"),
        ):
            ttk.Checkbutton(page_box, text=text, variable=pdf_option_vars[key]).pack(anchor="w", pady=2)

        def reset_pdf_output_options():
            for key, default in PDF_DISPLAY_DEFAULTS.items():
                pdf_option_vars[key].set(default)

        ttk.Button(pdf_output, text=self.t("reset_defaults"), command=reset_pdf_output_options).pack(anchor="w", pady=(12, 0))

        # --- 命名ルール ---
        naming = add_settings_tab(self.t("tab_naming"))
        naming.columnconfigure(1, weight=1)
        pdf_var = tk.StringVar(value=str(self.settings.get("pdf_name_template", DEFAULT_PDF_TEMPLATE)))
        folder_var = tk.StringVar(value=str(self.settings.get("folder_name_template", DEFAULT_FOLDER_TEMPLATE)))

        ttk.Label(naming, text=self.t("pdf_filename")).grid(row=0, column=0, sticky="w", padx=(0, 8), pady=5)
        pdf_entry = ttk.Entry(naming, textvariable=pdf_var, style="Settings.TEntry")
        pdf_entry.grid(row=0, column=1, sticky="ew", pady=5)
        ttk.Label(naming, text=self.t("folder_name")).grid(row=1, column=0, sticky="w", padx=(0, 8), pady=5)
        folder_entry = ttk.Entry(naming, textvariable=folder_var, style="Settings.TEntry")
        folder_entry.grid(row=1, column=1, sticky="ew", pady=5)

        honorific_var = tk.StringVar(value=str(self.settings.get("sender_honorific", "氏")))

        ttk.Label(naming, text="標準の敬称").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=5)
        honorific_combo = ttk.Combobox(
            naming,
            textvariable=honorific_var,
            values=("氏", "様", ""),
            state="normal",
            width=18,
        )
        honorific_combo.grid(row=2, column=1, sticky="w", pady=5)
        ttk.Label(
            naming,
            text="差出人設定で「標準」を選んだ場合に使用します。任意文字も入力できます。",
            wraplength=650,
            foreground="#666",
        ).grid(row=3, column=0, columnspan=2, sticky="w", pady=(0, 5))

        ttk.Label(
            naming,
            text="PDF出力と同じメール項目を基準にしています。使いたい項目の置換文字を命名欄へ挿入できます。",
            wraplength=650,
        ).grid(row=4, column=0, columnspan=2, sticky="w", pady=(8, 6))

        item_frame = ttk.LabelFrame(naming, text=self.t("available_tokens"), padding=8)
        item_frame.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        item_frame.columnconfigure(0, weight=1)
        item_tree = ttk.Treeview(
            item_frame,
            columns=("item", "token", "example"),
            show="headings",
            selectmode="browse",
            height=10,
        )
        item_tree.heading("item", text=self.t("item"))
        item_tree.heading("token", text=self.t("token"))
        item_tree.heading("example", text=self.t("example"))
        item_tree.column("item", width=190, anchor="w")
        item_tree.column("token", width=150, anchor="w")
        item_tree.column("example", width=280, anchor="w")
        item_tree.grid(row=0, column=0, sticky="ew")
        item_scroll = ttk.Scrollbar(item_frame, orient="vertical", command=item_tree.yview)
        item_scroll.grid(row=0, column=1, sticky="ns")
        item_tree.configure(yscrollcommand=item_scroll.set)
        token_item_keys = ["from_original", "sender_display_name", "sender_honorific", "subject", "to_recipient", "cc", "bcc", "date", "time", "datetime", "reply_to", "message_id", "importance", "attachment_count_label"]
        for (item_name, token, example), item_key in zip(TEMPLATE_ITEM_SPECS, token_item_keys):
            item_tree.insert("", "end", values=(self.t(item_key), token, example))

        last_target = {"entry": pdf_entry}
        pdf_entry.bind("<FocusIn>", lambda _e: last_target.update(entry=pdf_entry))
        folder_entry.bind("<FocusIn>", lambda _e: last_target.update(entry=folder_entry))

        def selected_token() -> str:
            selected = item_tree.selection()
            if not selected:
                return ""
            values = item_tree.item(selected[0], "values")
            return str(values[1]) if len(values) > 1 else ""

        def insert_token(entry):
            token = selected_token()
            if not token:
                messagebox.showinfo("命名ルール", "挿入する項目を一覧から選択してください。", parent=dialog)
                return
            entry.insert(tk.INSERT, token)
            entry.focus_set()

        insert_buttons = ttk.Frame(naming)
        insert_buttons.grid(row=8, column=0, columnspan=2, sticky="w", pady=(0, 6))
        ttk.Button(insert_buttons, text=self.t("insert_pdf"), command=lambda: insert_token(pdf_entry)).pack(side="left")
        ttk.Button(insert_buttons, text=self.t("insert_folder"), command=lambda: insert_token(folder_entry)).pack(side="left", padx=(6, 0))
        ttk.Label(insert_buttons, text=self.t("doubleclick_hint"), foreground="#666").pack(side="left", padx=(12, 0))
        item_tree.bind("<Double-1>", lambda _e: insert_token(last_target["entry"]))

        preview_var = tk.StringVar()
        ttk.Label(naming, textvariable=preview_var, wraplength=650, justify="left").grid(row=9, column=0, columnspan=2, sticky="w", pady=(8, 10))

        def naming_context():
            if self.mail_data:
                edited_from = self.var_from.get().strip()
                display_name, address = parseaddr(edited_from)
                sender_label = display_name.strip() if display_name else edited_from
                if not sender_label and address:
                    sender_label = address.split("@", 1)[0]
                return build_template_context(
                    self.mail_data, sender_label=sender_label, subject=self.var_subject.get().strip(),
                    sender_honorific=self.current_sender_honorific(address)
                )
            return build_template_context(None, sender_label="田中", subject="追加資料送付", sender_honorific=honorific_var.get().strip())

        def refresh_preview(*_):
            context = naming_context()
            try:
                pdf = render_name_template(pdf_var.get(), context=context)
                folder = render_name_template(folder_var.get(), context=context)
                preview_var.set(f"プレビュー\nPDF：{sanitize_filename(pdf, 'mail.pdf')}\nフォルダ：{sanitize_filename(folder, 'mail')}")
            except ValueError as exc:
                preview_var.set(f"プレビューエラー：{exc}")

        pdf_var.trace_add("write", refresh_preview)
        folder_var.trace_add("write", refresh_preview)
        refresh_preview()
        ttk.Button(
            naming, text=self.t("reset_defaults"), style="Settings.TButton",
            command=lambda: (pdf_var.set(DEFAULT_PDF_TEMPLATE), folder_var.set(DEFAULT_FOLDER_TEMPLATE)),
        ).grid(row=6, column=0, sticky="w", pady=(8, 0))

        # --- 差出人辞書 ---
        sender_tab = add_settings_tab("差出人設定")
        sender_list_frame = tk.Frame(
            sender_tab, bg="#FFFFFF", bd=0,
            highlightbackground=SETTINGS_BORDER, highlightthickness=1,
        )
        sender_list_frame.pack(fill="both", expand=True)

        tree = ttk.Treeview(
            sender_list_frame,
            columns=("key", "name", "honorific"),
            show="headings",
            selectmode="browse",
        )
        tree.heading("key", text="メールアドレス / From")
        tree.heading("name", text="表示名")
        tree.heading("honorific", text="敬称")
        tree.column("key", width=360)
        tree.column("name", width=180)
        tree.column("honorific", width=100, anchor="center")
        tree.pack(fill="both", expand=True, padx=0, pady=0)

        sender_rules_local = dict(self.settings.get("sender_honorific_rules", {}) or {})

        # 旧「自分のメールアドレス」は敬称なしとして自動移行表示する。
        old_self = self.settings.get("self_email_addresses", [])
        if isinstance(old_self, str):
            old_self = [old_self]
        for email in old_self:
            key = str(email).strip().lower()
            if key and key not in sender_rules_local:
                sender_rules_local[key] = {"mode": "none", "custom": ""}

        def honorific_label(key):
            row = sender_rules_local.get(str(key).strip().lower(), {})
            mode = str(row.get("mode", "standard") or "standard")
            if mode == "none":
                return "なし"
            if mode == "shi":
                return "氏"
            if mode == "sama":
                return "様"
            if mode == "custom":
                return str(row.get("custom", "") or "任意")
            return "標準"

        def refresh_sender_tree():
            for item in tree.get_children():
                tree.delete(item)
            keys = set(self.sender_dictionary.keys()) | set(sender_rules_local.keys())
            for key in sorted(keys):
                tree.insert(
                    "", "end",
                    iid=str(key),
                    values=(key, self.sender_dictionary.get(key, ""), honorific_label(key)),
                )

        def edit_sender():
            selected = tree.selection()
            if not selected:
                return
            key = str(selected[0])
            current_name = self.sender_dictionary.get(key, "")
            current_rule = dict(sender_rules_local.get(key, {}) or {})

            win = tk.Toplevel(dialog)
            win.title("差出人設定の編集")
            win.transient(dialog)
            win.grab_set()

            frm = ttk.Frame(win, padding=12)
            frm.pack(fill="both", expand=True)

            ttk.Label(frm, text="メールアドレス / From").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=5)
            key_var = tk.StringVar(value=key)
            ttk.Entry(frm, textvariable=key_var, width=46).grid(row=0, column=1, sticky="ew", pady=5)

            ttk.Label(frm, text="表示名").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=5)
            name_var = tk.StringVar(value=current_name)
            ttk.Entry(frm, textvariable=name_var, width=46).grid(row=1, column=1, sticky="ew", pady=5)

            mode_to_label = {
                "standard": "標準",
                "none": "なし",
                "shi": "氏",
                "sama": "様",
                "custom": "任意",
            }
            label_to_mode = {v: k for k, v in mode_to_label.items()}

            ttk.Label(frm, text="敬称").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=5)
            mode_var = tk.StringVar(value=mode_to_label.get(str(current_rule.get("mode", "standard")), "標準"))
            ttk.Combobox(
                frm,
                textvariable=mode_var,
                values=("標準", "なし", "氏", "様", "任意"),
                state="readonly",
                width=18,
            ).grid(row=2, column=1, sticky="w", pady=5)

            ttk.Label(frm, text="任意敬称").grid(row=3, column=0, sticky="w", padx=(0, 8), pady=5)
            custom_var = tk.StringVar(value=str(current_rule.get("custom", "") or ""))
            ttk.Entry(frm, textvariable=custom_var, width=20).grid(row=3, column=1, sticky="w", pady=5)

            ttk.Label(
                frm,
                text="「標準」は命名ルールの標準敬称を使用します。「なし」は氏・様などを付けません。",
                wraplength=500,
                foreground="#666",
            ).grid(row=4, column=0, columnspan=2, sticky="w", pady=(4, 10))

            frm.columnconfigure(1, weight=1)

            def apply_sender():
                new_key = key_var.get().strip().lower()
                if not new_key:
                    messagebox.showerror("差出人設定", "メールアドレス / From を入力してください。", parent=win)
                    return

                if new_key != key:
                    self.sender_dictionary.pop(key, None)
                    sender_rules_local.pop(key, None)

                display_name = name_var.get().strip()
                if display_name:
                    self.sender_dictionary[new_key] = display_name
                else:
                    self.sender_dictionary.pop(new_key, None)

                sender_rules_local[new_key] = {
                    "mode": label_to_mode.get(mode_var.get(), "standard"),
                    "custom": custom_var.get().strip(),
                }

                save_sender_dictionary(self.sender_dictionary)
                refresh_sender_tree()
                win.destroy()

            btns = ttk.Frame(frm)
            btns.grid(row=5, column=0, columnspan=2, sticky="e")
            ttk.Button(btns, text="キャンセル", command=win.destroy).pack(side="right")
            ttk.Button(btns, text="OK", command=apply_sender).pack(side="right", padx=(0, 6))

        def delete_sender():
            selected = tree.selection()
            if not selected:
                return
            key = str(selected[0])
            if messagebox.askyesno("削除確認", f"差出人設定から削除しますか？\n\n{key}", parent=dialog):
                self.sender_dictionary.pop(key, None)
                sender_rules_local.pop(key, None)
                save_sender_dictionary(self.sender_dictionary)
                refresh_sender_tree()

        sender_buttons = tk.Frame(sender_tab, bg=SETTINGS_BG, bd=0, highlightthickness=0)
        sender_buttons.pack(fill="x", pady=(10, 0))
        ttk.Button(sender_buttons, text=self.t("edit"), command=edit_sender, style="Settings.TButton").pack(side="left")
        ttk.Button(sender_buttons, text=self.t("delete_short"), command=delete_sender, style="Settings.TButton").pack(side="left", padx=(8, 0))
        ttk.Label(
            sender_buttons,
            text="差出人ごとに表示名と敬称を設定します。自分のメールは敬称を「なし」に設定してください。",
            foreground="#555", style="Settings.TLabel",
        ).pack(side="left", padx=(14, 0))
        tree.bind("<Double-1>", lambda _e: edit_sender())
        refresh_sender_tree()

        # --- バックアップ ---
        backup = add_settings_tab(self.t("tab_backup"))
        ttk.Label(backup, text=self.t("backup_note"), wraplength=620).pack(anchor="w", pady=(0, 12))
        ttk.Button(backup, text=self.t("export_settings"), command=self.export_settings).pack(anchor="w", pady=4)
        ttk.Button(backup, text=self.t("import_settings"), command=self.import_settings).pack(anchor="w", pady=4)
        ttk.Separator(backup, orient="horizontal").pack(fill="x", pady=14)
        ttk.Label(backup, text=f"設定保存先：\n{SETTINGS_PATH}\n\n差出人辞書：\n{SENDER_DICT_PATH}", foreground="#555", justify="left").pack(anchor="w")

        show_settings_tab(self.t({"基本設定":"tab_basic","PDF出力":"tab_pdf","命名ルール":"tab_naming","差出人辞書":"tab_sender_dict","差出人設定":"tab_sender_dict","バックアップ":"tab_backup"}.get(initial_tab, "tab_basic")))

        footer = tk.Frame(frame, bg=SETTINGS_BG, bd=0, highlightthickness=0)
        footer.pack(fill="x", pady=(10, 0))

        def save_all():
            context = naming_context()
            try:
                pdf_test = render_name_template(pdf_var.get(), context=context)
                folder_test = render_name_template(folder_var.get(), context=context)
                if not pdf_test.strip() or not folder_test.strip():
                    raise ValueError("PDFファイル名と作成フォルダ名は空欄にできません。")
            except ValueError as exc:
                messagebox.showerror("設定", str(exc), parent=dialog)
                show_settings_tab(self.t("tab_naming"))
                return

            old_import = self.var_import.get()
            old_watch = self.var_watch.get()
            old_resident = self.var_resident.get()
            old_auto = self.var_auto_start.get()

            self.var_import.set(import_var.get().strip() or str(DEFAULT_IMPORT))
            self.var_watch.set(watch_var.get())
            self.var_resident.set(resident_var.get())
            self.var_auto_start.set(auto_start_var.get())
            self.var_ignore_inline.set(ignore_inline_var.get())
            pdf_template_value = pdf_var.get().strip() or DEFAULT_PDF_TEMPLATE
            folder_template_value = folder_var.get().strip() or DEFAULT_FOLDER_TEMPLATE
            if pdf_template_value in {"メール_{sender_label}氏_{datetime}.pdf", "メール_{sender_label}様_{datetime}.pdf"}:
                pdf_template_value = DEFAULT_PDF_TEMPLATE
            if folder_template_value in {"{date}_追加資料_メール{sender_label}氏", "{date}_追加資料_メール{sender_label}様"}:
                folder_template_value = DEFAULT_FOLDER_TEMPLATE
            self.settings["pdf_name_template"] = pdf_template_value
            self.settings["folder_name_template"] = folder_template_value
            self.settings["sender_honorific"] = honorific_var.get().strip()
            self.settings["sender_honorific_rules"] = sender_rules_local
            # 旧設定は新ルールへ移行済みなので空にする。
            self.settings["self_email_addresses"] = []
            self.settings["pdf_language"] = pdf_language_var.get()
            self.settings["default_output_folder"] = default_output_var.get().strip() or str(DEFAULT_OUTPUT)
            self.var_default_output.set(self.settings["default_output_folder"])
            self.settings["confirm_auto_detected_mail"] = bool(confirm_auto_var.get())
            self.settings["delete_excluded_temp_eml"] = bool(delete_excluded_var.get())
            self.settings["registered_save_locations"] = self.get_registered_save_locations()
            self.settings["duplicate_check_enabled"] = bool(duplicate_check_var.get())
            self.settings["recent_destination_suggestions"] = bool(recent_dest_var.get())
            self.settings["attachment_selective_save"] = bool(selective_att_var.get())
            self.settings["external_link_detection_enabled"] = bool(external_link_var.get())
            self.settings["auto_show_pending_on_new_eml"] = bool(auto_show_pending_var.get())
            self.settings["excluded_addresses"] = [line.strip().lower() for line in excluded_addresses_text.get("1.0", "end").splitlines() if line.strip()]
            self.settings["excluded_domains"] = [line.strip().lower().lstrip("@") for line in excluded_domains_text.get("1.0", "end").splitlines() if line.strip()]
            for key in PDF_DISPLAY_DEFAULTS:
                self.settings[key] = bool(pdf_option_vars[key].get())
            Path(self.var_import.get()).mkdir(parents=True, exist_ok=True)
            self.persist_settings()
            self.update_names()

            # 除外設定を保存した時点で、既存の未処理EMLも即時再判定する。
            try:
                self.apply_exclusion_to_pending_emls(update_ui=True)
            except Exception:
                logging.exception("保存対象外設定反映後の未処理EML再判定に失敗しました")

            if self.var_import.get() != old_import or self.var_watch.get() != old_watch:
                self.restart_watcher()
            if self.var_auto_start.get() != old_auto:
                self.sync_startup_setting(show_error=True)
            if self.var_resident.get() != old_resident:
                if self.var_resident.get():
                    self.start_tray_icon()
                    try:
                        self.file_menu.entryconfig("タスクトレイに格納", state="normal")
                    except tk.TclError:
                        pass
                else:
                    self.stop_tray_icon()
                    try:
                        self.file_menu.entryconfig("タスクトレイに格納", state="disabled")
                    except tk.TclError:
                        pass
            self.var_status.set("設定を保存しました。")
            dialog.destroy()

        ttk.Button(footer, text=self.t("cancel"), command=dialog.destroy, style="Settings.TButton").pack(side="right")
        ttk.Button(footer, text=self.t("save"), command=save_all, style="Settings.TButton").pack(side="right", padx=(0, 6))
        dialog.bind("<Escape>", lambda _e: dialog.destroy())

    def show_sender_dictionary(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("差出人辞書")
        dialog.geometry("640x420")
        dialog.transient(self.root)

        frame = ttk.Frame(dialog, padding=10)
        frame.pack(fill="both", expand=True)
        tree = ttk.Treeview(frame, columns=("key", "name"), show="headings", selectmode="browse")
        tree.heading("key", text=self.t("email_or_from"))
        tree.heading("name", text=self.t("display_name"))
        tree.column("key", width=360)
        tree.column("name", width=180)
        tree.pack(fill="both", expand=True)

        def refresh():
            for item in tree.get_children():
                tree.delete(item)
            for key, value in sorted(self.sender_dictionary.items()):
                tree.insert("", "end", values=(key, value))

        def edit_selected():
            selected = tree.selection()
            if not selected:
                return
            key, value = tree.item(selected[0], "values")
            new_value = self.simple_text_prompt(dialog, "差出人名の編集", f"{key} の表示名", value)
            if new_value is None:
                return
            new_value = new_value.strip()
            if new_value:
                self.sender_dictionary[key] = new_value
                save_sender_dictionary(self.sender_dictionary)
                refresh()

        def delete_selected():
            selected = tree.selection()
            if not selected:
                return
            key, _value = tree.item(selected[0], "values")
            if messagebox.askyesno("削除確認", f"差出人辞書から削除しますか？\n\n{key}", parent=dialog):
                self.sender_dictionary.pop(key, None)
                save_sender_dictionary(self.sender_dictionary)
                refresh()

        buttons = ttk.Frame(frame)
        buttons.pack(fill="x", pady=(8, 0))
        ttk.Button(buttons, text=self.t("edit"), command=edit_selected).pack(side="left")
        ttk.Button(buttons, text=self.t("delete_short"), command=delete_selected).pack(side="left", padx=6)
        ttk.Button(buttons, text=self.t("close"), command=dialog.destroy).pack(side="right")
        tree.bind("<Double-1>", lambda _e: edit_selected())
        refresh()

    def show_naming_rules(self):
        """旧メニュー互換：統合設定画面の「命名ルール」タブを開く。"""
        self.show_settings_dialog(initial_tab="命名ルール")

    @staticmethod
    def simple_text_prompt(parent, title: str, label: str, initial: str = "") -> str | None:
        result = {"value": None}
        dialog = tk.Toplevel(parent)
        dialog.title(title)
        dialog.transient(parent)
        dialog.grab_set()
        dialog.resizable(False, False)
        frame = ttk.Frame(dialog, padding=12)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text=label).pack(anchor="w")
        var = tk.StringVar(value=initial)
        entry = ttk.Entry(frame, textvariable=var, width=48)
        entry.pack(fill="x", pady=(6, 10))
        entry.select_range(0, tk.END)
        entry.focus_set()
        btns = ttk.Frame(frame)
        btns.pack(fill="x")
        def ok():
            result["value"] = var.get()
            dialog.destroy()
        ttk.Button(btns, text="OK", command=ok).pack(side="right")
        ttk.Button(btns, text=self.t("cancel"), command=dialog.destroy).pack(side="right", padx=(0, 6))
        dialog.bind("<Return>", lambda _e: ok())
        dialog.bind("<Escape>", lambda _e: dialog.destroy())
        parent.wait_window(dialog)
        return result["value"]

    def export_settings(self):
        path = filedialog.asksaveasfilename(
            title="設定をエクスポート",
            defaultextension=".json",
            filetypes=[("JSON", "*.json")],
            initialfile="INAS_Mail_Archive_Settings.json",
        )
        if not path:
            return
        bundle = {
            "app": APP_NAME,
            "version": APP_VERSION,
            "settings": {
                "import_folder": self.var_import.get(),
                "watch_enabled": self.var_watch.get(),
                "post_action": self.var_post_action.get(),
                "open_after_save": self.var_open_after.get(),
                "auto_start": self.var_auto_start.get(),
                "resident_enabled": self.var_resident.get(),
                "ignore_inline_images": self.var_ignore_inline.get(),
                "pdf_name_template": self.settings.get("pdf_name_template", DEFAULT_PDF_TEMPLATE),
                "folder_name_template": self.settings.get("folder_name_template", DEFAULT_FOLDER_TEMPLATE),
                "self_email_addresses": self.settings.get("self_email_addresses", []),
                "sender_honorific": self.settings.get("sender_honorific", "氏"),
            "sender_honorific_rules": self.settings.get("sender_honorific_rules", {}),
                **{key: bool(self.settings.get(key, default)) for key, default in PDF_DISPLAY_DEFAULTS.items()},
            },
            "sender_dictionary": self.sender_dictionary,
        }
        Path(path).write_text(json.dumps(bundle, ensure_ascii=False, indent=2), encoding="utf-8")
        messagebox.showinfo("設定エクスポート", "設定と差出人辞書を保存しました。")

    def import_settings(self):
        path = filedialog.askopenfilename(
            title="設定をインポート",
            filetypes=[("JSON", "*.json"), ("すべてのファイル", "*.*")],
        )
        if not path:
            return
        try:
            data = json.loads(Path(path).read_text(encoding="utf-8"))
            settings = data.get("settings", {}) if isinstance(data, dict) else {}
            senders = data.get("sender_dictionary", {}) if isinstance(data, dict) else {}
            if not isinstance(settings, dict) or not isinstance(senders, dict):
                raise ValueError("設定ファイルの形式が正しくありません。")
            self.var_import.set(str(settings.get("import_folder", self.var_import.get())))
            self.var_watch.set(bool(settings.get("watch_enabled", self.var_watch.get())))
            self.var_post_action.set(str(settings.get("post_action", self.var_post_action.get())))
            self.var_open_after.set(bool(settings.get("open_after_save", self.var_open_after.get())))
            self.var_auto_start.set(bool(settings.get("auto_start", self.var_auto_start.get())))
            self.var_resident.set(bool(settings.get("resident_enabled", self.var_resident.get())))
            self.var_ignore_inline.set(bool(settings.get("ignore_inline_images", self.var_ignore_inline.get())))
            self.settings["default_output_folder"] = str(settings.get("default_output_folder", self.settings.get("default_output_folder", str(DEFAULT_OUTPUT))))
            self.var_default_output.set(self.settings["default_output_folder"])
            self.settings["excluded_addresses"] = [str(v).strip().lower() for v in settings.get("excluded_addresses", self.settings.get("excluded_addresses", [])) if str(v).strip()]
            self.settings["excluded_domains"] = [str(v).strip().lower().lstrip("@") for v in settings.get("excluded_domains", self.settings.get("excluded_domains", [])) if str(v).strip()]
            self.settings["confirm_auto_detected_mail"] = bool(settings.get("confirm_auto_detected_mail", self.settings.get("confirm_auto_detected_mail", True)))
            self.settings["delete_excluded_temp_eml"] = bool(settings.get("delete_excluded_temp_eml", self.settings.get("delete_excluded_temp_eml", True)))
            self.settings["registered_save_locations"] = settings.get("registered_save_locations", self.settings.get("registered_save_locations", []))
            self.settings["duplicate_check_enabled"] = bool(settings.get("duplicate_check_enabled", self.settings.get("duplicate_check_enabled", True)))
            self.settings["recent_destination_suggestions"] = bool(settings.get("recent_destination_suggestions", self.settings.get("recent_destination_suggestions", True)))
            self.settings["attachment_selective_save"] = bool(settings.get("attachment_selective_save", self.settings.get("attachment_selective_save", True)))
            self.settings["external_link_detection_enabled"] = bool(settings.get("external_link_detection_enabled", self.settings.get("external_link_detection_enabled", True)))
            self.settings["auto_show_pending_on_new_eml"] = bool(settings.get("auto_show_pending_on_new_eml", self.settings.get("auto_show_pending_on_new_eml", True)))
            self.settings["pdf_name_template"] = str(settings.get("pdf_name_template", self.settings.get("pdf_name_template", DEFAULT_PDF_TEMPLATE)))
            self.settings["folder_name_template"] = str(settings.get("folder_name_template", self.settings.get("folder_name_template", DEFAULT_FOLDER_TEMPLATE)))
            for key, default in PDF_DISPLAY_DEFAULTS.items():
                self.settings[key] = bool(settings.get(key, self.settings.get(key, default)))
            self.sender_dictionary = {str(k): str(v) for k, v in senders.items() if str(v).strip()}
            save_sender_dictionary(self.sender_dictionary)
            self.persist_settings()
            self.restart_watcher()
            self.sync_startup_setting(show_error=False)
            if self.var_resident.get():
                self.start_tray_icon()
            else:
                self.stop_tray_icon()
            messagebox.showinfo("設定インポート", "設定と差出人辞書を読み込みました。")
        except Exception as exc:
            logging.exception("設定インポートに失敗しました")
            messagebox.showerror("設定インポート", f"設定を読み込めませんでした。\n\n{exc}")

    def update_names(self, *_):
        if not self.mail_data:
            return

        edited_from = self.var_from.get().strip()
        display_name, address = parseaddr(edited_from)
        sender_label = display_name.strip() if display_name else edited_from
        if not sender_label and address:
            sender_label = address.split("@", 1)[0]

        subject = self.var_subject.get().strip()
        # 敬称は編集後のFrom文字列ではなく、元EMLのFromメールアドレスから
        # 差出人設定を再照合して決定する。
        sender_address = self.sender_address_from_mail(self.mail_data)
        sender_honorific = self.current_sender_honorific(sender_address)
        context = build_template_context(
            self.mail_data,
            sender_label=sender_label,
            subject=subject,
            sender_honorific=sender_honorific,
        )
        pdf_template = str(self.settings.get("pdf_name_template", DEFAULT_PDF_TEMPLATE))
        folder_template = str(self.settings.get("folder_name_template", DEFAULT_FOLDER_TEMPLATE))
        try:
            pdf_name = render_name_template(pdf_template, context=context)
            folder_name = render_name_template(folder_template, context=context)
        except ValueError as exc:
            logging.warning("命名テンプレートの展開に失敗しました: %s", exc)
            pdf_name = render_name_template(DEFAULT_PDF_TEMPLATE, context=context)
            folder_name = render_name_template(DEFAULT_FOLDER_TEMPLATE, context=context)

        pdf_name = sanitize_filename(pdf_name, "mail.pdf")
        if not pdf_name.lower().endswith(".pdf"):
            pdf_name += ".pdf"
        self.var_pdf.set(pdf_name)
        self.var_folder.set(sanitize_filename(folder_name, "mail"))

    def sender_address_from_mail(self, data: MailData) -> str:
        """EMLの元Fromからメールアドレスを小文字で返す。"""
        _name, address = parseaddr(data.from_name or "")
        return address.strip().lower()

    def is_excluded_mail(self, data: MailData) -> tuple[bool, str]:
        address = self.sender_address_from_mail(data)
        if not address:
            return False, ""
        addresses = {str(v).strip().lower() for v in self.settings.get("excluded_addresses", []) if str(v).strip()}
        domains = {str(v).strip().lower().lstrip("@") for v in self.settings.get("excluded_domains", []) if str(v).strip()}
        if address in addresses:
            return True, address
        domain = address.rsplit("@", 1)[-1] if "@" in address else ""
        if domain and domain in domains:
            return True, "@" + domain
        return False, ""

    def discard_excluded_temp_eml(self, path: Path) -> None:
        """Power Automate等の取込フォルダ内EMLだけを、除外時に安全に削除する。"""
        if not self.settings.get("delete_excluded_temp_eml", True):
            return
        try:
            import_dir = Path(self.var_import.get()).resolve()
            if path.exists() and path.resolve().parent == import_dir:
                path.unlink()
        except Exception:
            logging.exception("除外EMLの一時ファイル削除に失敗しました: %s", path)

    @staticmethod
    def preview_header_value(header_block: str, key: str) -> str:
        """プレビュー画面用にヘッダー項目を1件取り出す。"""
        prefix = key.lower() + ":"
        for line in normalize_text(header_block).splitlines():
            if line.lower().startswith(prefix):
                return line.split(":", 1)[1].strip()
        return ""


    def get_registered_save_locations(self) -> list[dict]:
        """Return normalized registered save locations."""
        result = []
        raw = self.settings.get("registered_save_locations", [])
        if not isinstance(raw, list):
            return result
        seen = set()
        for item in raw:
            if isinstance(item, str):
                name = Path(item).name or item
                path = item
            elif isinstance(item, dict):
                name = str(item.get("name", "")).strip()
                path = str(item.get("path", "")).strip()
            else:
                continue
            if not path:
                continue
            key = path.lower()
            if key in seen:
                continue
            seen.add(key)
            result.append({"name": name or Path(path).name or path, "path": path})
        return result

    def save_registered_save_locations(self, locations: list[dict]) -> None:
        self.settings["registered_save_locations"] = locations
        save_settings(self.settings)

    def show_registered_save_locations(self, parent=None, on_select=None):
        """Manage frequently used save locations.

        on_select(path) is called when the user chooses a registered location
        from the manager using the Select button.
        """
        dialog = tk.Toplevel(parent or self.root)
        dialog.title("保存先の登録")
        dialog.transient(parent or self.root)
        dialog.geometry("760x430")
        try:
            dialog.grab_set()
        except Exception:
            pass

        outer = ttk.Frame(dialog, padding=12)
        outer.pack(fill="both", expand=True)

        ttk.Label(
            outer,
            text="よく使う保存先を登録します。メール確認画面からすぐに選択できます。",
            wraplength=700,
            justify="left",
        ).pack(anchor="w", pady=(0, 8))

        tree_frame = ttk.Frame(outer)
        tree_frame.pack(fill="both", expand=True)
        tree = ttk.Treeview(tree_frame, columns=("name", "path"), show="headings", selectmode="extended")
        tree.heading("name", text="登録名")
        tree.heading("path", text="保存先")
        tree.column("name", width=180, anchor="w")
        tree.column("path", width=520, anchor="w")
        yscroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=yscroll.set)
        tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        status_var = tk.StringVar(value="")

        def refresh():
            for iid in tree.get_children():
                tree.delete(iid)
            for idx, item in enumerate(self.get_registered_save_locations()):
                tree.insert("", "end", iid=str(idx), values=(item["name"], item["path"]))
            status_var.set(f"登録 {len(self.get_registered_save_locations())}件")

        def add_location():
            folder = filedialog.askdirectory(parent=dialog, title="登録する保存先を選択")
            if not folder:
                return
            name = Path(folder).name or folder
            locations = self.get_registered_save_locations()
            if any(x["path"].lower() == folder.lower() for x in locations):
                messagebox.showinfo("保存先の登録", "この保存先はすでに登録されています。", parent=dialog)
                return
            locations.append({"name": name, "path": folder})
            self.save_registered_save_locations(locations)
            refresh()

        def rename_location():
            selected = tree.selection()
            if len(selected) != 1:
                messagebox.showinfo("保存先の登録", "名前を変更する保存先を1件選択してください。", parent=dialog)
                return
            idx = int(selected[0])
            locations = self.get_registered_save_locations()
            if idx >= len(locations):
                return
            item = locations[idx]

            edit = tk.Toplevel(dialog)
            edit.title("登録名の変更")
            edit.transient(dialog)
            edit.grab_set()
            frame = ttk.Frame(edit, padding=12)
            frame.pack(fill="both", expand=True)
            ttk.Label(frame, text="登録名").grid(row=0, column=0, sticky="w", padx=(0, 8))
            var_name = tk.StringVar(value=item["name"])
            entry = ttk.Entry(frame, textvariable=var_name, width=44)
            entry.grid(row=0, column=1, sticky="ew")
            ttk.Label(frame, text=item["path"], wraplength=480, foreground="#555").grid(
                row=1, column=0, columnspan=2, sticky="w", pady=(8, 12)
            )
            frame.columnconfigure(1, weight=1)

            def apply():
                name = var_name.get().strip()
                if not name:
                    messagebox.showerror("登録名", "登録名を入力してください。", parent=edit)
                    return
                locations[idx]["name"] = name
                self.save_registered_save_locations(locations)
                refresh()
                edit.destroy()

            buttons = ttk.Frame(frame)
            buttons.grid(row=2, column=0, columnspan=2, sticky="e")
            ttk.Button(buttons, text="キャンセル", command=edit.destroy).pack(side="right")
            ttk.Button(buttons, text="OK", command=apply).pack(side="right", padx=(0, 6))
            entry.focus_set()
            entry.select_range(0, "end")

        def delete_selected():
            selected = list(tree.selection())
            if not selected:
                messagebox.showinfo("保存先の登録", "削除する保存先を選択してください。", parent=dialog)
                return
            if not messagebox.askyesno(
                "保存先登録を削除",
                f"選択した {len(selected)} 件の登録を削除しますか？\n\n実際のフォルダや保存済みデータは削除されません。",
                parent=dialog,
            ):
                return
            indexes = {int(i) for i in selected}
            locations = [x for i, x in enumerate(self.get_registered_save_locations()) if i not in indexes]
            self.save_registered_save_locations(locations)
            refresh()

        def delete_all():
            locations = self.get_registered_save_locations()
            if not locations:
                return
            if not messagebox.askyesno(
                "保存先登録をすべて削除",
                "登録した保存先をすべて削除しますか？\n\n実際のフォルダや保存済みデータは削除されません。",
                parent=dialog,
            ):
                return
            self.save_registered_save_locations([])
            refresh()

        def select_location():
            if on_select is None:
                return
            selected = tree.selection()
            if len(selected) != 1:
                messagebox.showinfo("保存先の登録", "使用する保存先を1件選択してください。", parent=dialog)
                return
            idx = int(selected[0])
            locations = self.get_registered_save_locations()
            if idx < len(locations):
                on_select(locations[idx]["path"])
                dialog.destroy()

        buttons = ttk.Frame(outer)
        buttons.pack(fill="x", pady=(8, 0))
        ttk.Button(buttons, text="追加...", command=add_location).pack(side="left")
        ttk.Button(buttons, text="登録名変更", command=rename_location).pack(side="left", padx=6)
        ttk.Button(buttons, text="選択削除", command=delete_selected).pack(side="left")
        ttk.Button(buttons, text="すべて削除", command=delete_all).pack(side="left", padx=6)
        if on_select is not None:
            ttk.Button(buttons, text="この保存先を使用", command=select_location).pack(side="right", padx=(6, 0))
        ttk.Button(buttons, text="閉じる", command=dialog.destroy).pack(side="right")
        ttk.Label(outer, textvariable=status_var).pack(anchor="w", pady=(6, 0))

        tree.bind("<Double-1>", lambda _e: select_location() if on_select is not None else None)
        refresh()
        return dialog


    def _history_entries(self) -> list[dict]:
        """Return normalized history rows from existing history storage if available."""
        candidates = []
        for attr in ("history", "process_history", "history_entries"):
            value = getattr(self, attr, None)
            if isinstance(value, list):
                candidates.extend([x for x in value if isinstance(x, dict)])
        # fallback: load a common json history file if used by older versions
        for p in (
            getattr(self, "history_file", None),
            getattr(self, "history_path", None),
        ):
            try:
                if p and Path(p).exists():
                    data = json.loads(Path(p).read_text(encoding="utf-8"))
                    if isinstance(data, list):
                        candidates.extend([x for x in data if isinstance(x, dict)])
            except Exception:
                pass
        return candidates

    def _message_identity(self, data: MailData, path: Path | None = None) -> str:
        """Stable identity for duplicate checks: Message-ID preferred, file hash fallback."""
        for name in ("message_id", "messageid", "messageId"):
            value = getattr(data, name, None)
            if value:
                return "msgid:" + str(value).strip().lower()
        try:
            raw = getattr(data, "raw_bytes", None)
            if raw:
                return "sha256:" + hashlib.sha256(raw).hexdigest()
        except Exception:
            pass
        try:
            if path and Path(path).exists():
                return "sha256:" + hashlib.sha256(Path(path).read_bytes()).hexdigest()
        except Exception:
            pass
        return ""

    def _duplicate_registry_path(self) -> Path:
        base = Path.home() / "Documents" / "INAS_Mail_Archive"
        base.mkdir(parents=True, exist_ok=True)
        return base / "duplicate_registry.json"

    def _load_duplicate_registry(self) -> dict:
        p = self._duplicate_registry_path()
        try:
            if p.exists():
                data = json.loads(p.read_text(encoding="utf-8"))
                if isinstance(data, dict):
                    return data
        except Exception:
            logging.exception("重複保存台帳の読込に失敗しました")
        return {}

    def _save_duplicate_registry(self, data: dict) -> None:
        try:
            self._duplicate_registry_path().write_text(
                json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
            )
        except Exception:
            logging.exception("重複保存台帳の保存に失敗しました")

    def check_duplicate_mail(self, data: MailData, path: Path | None = None) -> tuple[bool, dict]:
        identity = self._message_identity(data, path)
        if not identity:
            return False, {}
        registry = self._load_duplicate_registry()
        row = registry.get(identity)
        return bool(row), (row if isinstance(row, dict) else {})

    def register_saved_mail_identity(self, data: MailData, saved_path: str, path: Path | None = None) -> None:
        identity = self._message_identity(data, path)
        if not identity:
            return
        registry = self._load_duplicate_registry()
        registry[identity] = {
            "saved_path": saved_path,
            "subject": getattr(data, "subject", "") or "",
            "sender": getattr(data, "sender", "") or getattr(data, "from_value", "") or "",
        }
        # keep registry bounded
        if len(registry) > 5000:
            for key in list(registry.keys())[:-4000]:
                registry.pop(key, None)
        self._save_duplicate_registry(registry)

    def recent_destination_candidates(self, data: MailData | None = None, limit: int = 10) -> list[str]:
        if not self.settings.get("recent_destination_suggestions", True):
            return []
        rows = self._history_entries()
        sender = ""
        subject = ""
        try:
            sender = (getattr(data, "sender", "") or getattr(data, "from_value", "") or "").lower()
            subject = (getattr(data, "subject", "") or "").lower()
        except Exception:
            pass

        scored = []
        seen = set()
        for idx, row in enumerate(reversed(rows)):
            path = str(
                row.get("save_path")
                or row.get("destination")
                or row.get("output_dir")
                or row.get("folder")
                or ""
            ).strip()
            if not path or path.lower() in seen:
                continue
            seen.add(path.lower())
            score = 0
            rsender = str(row.get("sender") or row.get("from") or "").lower()
            rsubject = str(row.get("subject") or "").lower()
            if sender and rsender and sender == rsender:
                score += 3
            if subject and rsubject:
                # simple shared-word similarity
                sw = {w for w in re.split(r"\s+", subject) if len(w) >= 3}
                rw = {w for w in re.split(r"\s+", rsubject) if len(w) >= 3}
                score += min(3, len(sw & rw))
            score += max(0, 2 - idx // 10)
            scored.append((score, idx, path))
        scored.sort(key=lambda x: (-x[0], x[1]))
        return [x[2] for x in scored[:limit]]


    def external_download_links_for_mail(self, data: MailData) -> list[dict[str, str]]:
        """MailData本文から外部ダウンロード候補を取得する。"""
        if not self.settings.get("external_link_detection_enabled", True):
            return []
        body = ""
        for attr in ("body_text", "text_body", "body", "plain_text"):
            value = getattr(data, attr, None)
            if value:
                body = str(value)
                break
        # MailDataに本文属性がない場合、body_parts相当を補完
        if not body:
            try:
                body = str(getattr(data, "body_content", "") or "")
            except Exception:
                body = ""
        return detect_external_download_links(body)

    def show_external_download_links(self, parent, links: list[dict[str, str]]):
        if not links:
            messagebox.showinfo("外部ファイル便", "外部ダウンロードURLは検出されませんでした。", parent=parent)
            return

        dialog = tk.Toplevel(parent)
        dialog.title(f"外部ダウンロードリンク（{len(links)}件）")
        dialog.transient(parent)
        dialog.geometry("820x360")
        dialog.minsize(680, 300)

        outer = ttk.Frame(dialog, padding=10)
        outer.pack(fill="both", expand=True)

        ttk.Label(
            outer,
            text="メール本文から検出した外部ダウンロード候補です。開くリンクを選択してください。",
            wraplength=760,
        ).pack(anchor="w", pady=(0, 8))

        tree_frame = ttk.Frame(outer)
        tree_frame.pack(fill="both", expand=True)
        tree = ttk.Treeview(tree_frame, columns=("service", "url"), show="headings", selectmode="browse")
        tree.heading("service", text="サービス")
        tree.heading("url", text="URL")
        tree.column("service", width=180, anchor="w")
        tree.column("url", width=600, anchor="w")
        ys = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        xs = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        tree.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        for idx, item in enumerate(links):
            tree.insert("", "end", iid=str(idx), values=(item["service"], item["url"]))

        def open_selected(_event=None):
            sel = tree.selection()
            if len(sel) != 1:
                messagebox.showinfo("外部ファイル便", "開くリンクを1件選択してください。", parent=dialog)
                return
            idx = int(sel[0])
            url = links[idx]["url"]
            try:
                webbrowser.open(url, new=2)
            except Exception as exc:
                logging.exception("外部リンクを開けませんでした: %s", url)
                messagebox.showerror(
                    "外部ファイル便",
                    f"リンクを開けませんでした。\n\n{url}\n\n{exc}",
                    parent=dialog,
                )

        buttons = ttk.Frame(outer)
        buttons.pack(fill="x", pady=(8, 0))
        ttk.Button(buttons, text="ブラウザで開く", command=open_selected).pack(side="left")
        ttk.Button(buttons, text="閉じる", command=dialog.destroy).pack(side="right")
        tree.bind("<Double-1>", open_selected)

    def show_mail_preview_dialog(self, data: MailData, path: Path) -> str:
        """自動取得EMLの内容を確認し、save / discard / cancel を返す。"""
        result = {"action": "cancel"}

        dialog = tk.Toplevel(self.root)
        dialog.title("メール内容確認")
        external_links = self.external_download_links_for_mail(data)

        if self.settings.get("duplicate_check_enabled", True):
            is_dup, dup_info = self.check_duplicate_mail(data, path)
            if is_dup:
                saved_to = str(dup_info.get("saved_path", "")).strip()
                msg = "このメールはすでに保存済みの可能性があります。"
                if saved_to:
                    msg += f"\n\n保存先：{saved_to}"
                msg += "\n\n内容を確認して、必要な場合のみ再保存してください。"
                messagebox.showwarning("重複保存チェック", msg, parent=dialog)
        dialog.geometry("1000x760")
        dialog.minsize(820, 620)
        dialog.transient(self.root)
        dialog.grab_set()

        outer = ttk.Frame(dialog, padding=12)
        outer.pack(side="top", fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        # 本文・添付欄だけを可変領域とし、保存先と下部ボタンは常に表示する。
        outer.rowconfigure(2, weight=4, minsize=110)
        outer.rowconfigure(4, weight=1, minsize=72)

        subject = self.extract_subject(data.header_block) or "（件名なし）"
        sender = data.from_name or "（差出人不明）"
        to_value = self.preview_header_value(data.header_block, "To") or "（なし）"
        cc_value = self.preview_header_value(data.header_block, "Cc") or "（なし）"
        sent_value = self.preview_header_value(data.header_block, "送信日時") or self.preview_header_value(data.header_block, "Date") or "（不明）"

        info = ttk.LabelFrame(outer, text="メール情報", padding=10)
        info.grid(row=0, column=0, sticky="ew")
        info.columnconfigure(1, weight=1)
        rows = [
            ("送信者", sender),
            ("宛先", to_value),
            ("Cc", cc_value),
            ("送信日時", sent_value),
            ("件名", subject),
        ]
        for r, (label, value) in enumerate(rows):
            ttk.Label(info, text=label + "：", width=10).grid(row=r, column=0, sticky="nw", padx=(0, 6), pady=2)
            ttk.Label(info, text=value, wraplength=800).grid(row=r, column=1, sticky="nw", pady=2)

        # 本文見出し＋文字サイズ操作
        body_header = ttk.Frame(outer)
        body_header.grid(row=1, column=0, sticky="ew", pady=(10, 4))
        ttk.Label(body_header, text="メール本文", font=("Yu Gothic UI", 10, "bold")).pack(side="left")

        base_body_font_size = 10
        body_font_size = tk.IntVar(value=base_body_font_size)
        body_zoom_label = tk.StringVar(value="100%")

        body_frame = ttk.Frame(outer)
        body_frame.grid(row=2, column=0, sticky="nsew")
        body_frame.columnconfigure(0, weight=1)
        body_frame.rowconfigure(0, weight=1)
        body_text = tk.Text(body_frame, wrap="word", font=("Yu Gothic UI", base_body_font_size), relief="solid", bd=1)
        body_scroll = ttk.Scrollbar(body_frame, orient="vertical", command=body_text.yview)
        body_text.configure(yscrollcommand=body_scroll.set)
        body_text.grid(row=0, column=0, sticky="nsew")
        body_scroll.grid(row=0, column=1, sticky="ns")
        body_text.insert("1.0", data.body_text or "（本文なし）")
        body_text.configure(state="disabled")

        def apply_body_zoom(size: int):
            size = max(7, min(24, int(size)))
            body_font_size.set(size)
            body_text.configure(font=("Yu Gothic UI", size))
            body_zoom_label.set(f"{round(size / base_body_font_size * 100)}%")

        def zoom_out():
            apply_body_zoom(body_font_size.get() - 1)

        def zoom_in():
            apply_body_zoom(body_font_size.get() + 1)

        def zoom_reset():
            apply_body_zoom(base_body_font_size)

        ttk.Label(body_header, text="文字").pack(side="right", padx=(8, 4))
        ttk.Button(body_header, text="＋", width=3, command=zoom_in).pack(side="right")
        ttk.Button(body_header, textvariable=body_zoom_label, width=7, command=zoom_reset).pack(side="right", padx=3)
        ttk.Button(body_header, text="－", width=3, command=zoom_out).pack(side="right")

        body_expanded = tk.BooleanVar(value=False)

        def toggle_body_area():
            expanded = not body_expanded.get()
            body_expanded.set(expanded)
            if expanded:
                # 本文を優先。添付欄は最小表示に戻す。
                outer.rowconfigure(2, weight=8, minsize=280)
                outer.rowconfigure(4, weight=1, minsize=72)
                try:
                    tree.configure(height=3)
                except Exception:
                    pass
                try:
                    attachment_expanded.set(False)
                    attachment_toggle_btn.configure(text="▼ 添付を拡大")
                except Exception:
                    pass
                body_toggle_btn.configure(text="▲ 本文を縮小")
            else:
                outer.rowconfigure(2, weight=4, minsize=110)
                outer.rowconfigure(4, weight=1, minsize=72)
                body_toggle_btn.configure(text="▼ 本文を拡大")
            dialog.update_idletasks()

        body_toggle_btn = ttk.Button(
            body_header, text="▼ 本文を拡大", command=toggle_body_area
        )
        body_toggle_btn.pack(side="right", padx=(10, 0))

        # 添付ファイル欄
        attachment_header = ttk.Frame(outer)
        attachment_header.grid(row=3, column=0, sticky="ew", pady=(10, 4))
        ttk.Label(
            attachment_header,
            text=f"添付ファイル（{len(data.attachments)}件）",
            font=("Yu Gothic UI", 10, "bold"),
        ).pack(side="left")

        attachment_frame = ttk.Frame(outer)
        attachment_frame.grid(row=4, column=0, sticky="nsew")
        attachment_frame.columnconfigure(0, weight=1)
        attachment_frame.rowconfigure(0, weight=1)
        tree = ttk.Treeview(
            attachment_frame,
            columns=("save", "name", "size"),
            show="headings",
            height=3,
        )
        tree.heading("save", text="保存")
        tree.heading("name", text="ファイル名")
        tree.heading("size", text="サイズ")
        tree.column("save", width=55, anchor="center", stretch=False)
        tree.column("name", width=595, anchor="w")
        tree.column("size", width=120, anchor="e")
        scroll = ttk.Scrollbar(attachment_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        tree.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")

        # data.attachments は (実ファイル名, bytes) のリスト。
        attachment_save_flags = {}
        for idx, (filename, payload) in enumerate(data.attachments):
            attachment_save_flags[idx] = True
            tree.insert(
                "", "end",
                iid=str(idx),
                values=("☑", filename, format_bytes(len(payload))),
            )

        def toggle_attachment_save(event):
            """「保存」列をクリックして個別保存対象を切り替える。"""
            try:
                if tree.identify("region", event.x, event.y) != "cell":
                    return
                if tree.identify_column(event.x) != "#1":
                    return
                row = tree.identify_row(event.y)
                if not row:
                    return
                idx = int(row)
                attachment_save_flags[idx] = not attachment_save_flags.get(idx, True)
                values = list(tree.item(row, "values"))
                values[0] = "☑" if attachment_save_flags[idx] else "☐"
                tree.item(row, values=values)
                return "break"
            except Exception:
                return

        tree.bind("<Button-1>", toggle_attachment_save, add="+")

        def open_attachment(event=None):
            selection = tree.selection()
            if not selection:
                messagebox.showinfo("添付ファイル", "開く添付ファイルを選択してください。", parent=dialog)
                return
            idx = int(selection[0])
            filename, payload = data.attachments[idx]

            risky_exts = {
                ".exe", ".com", ".bat", ".cmd", ".msi", ".msp",
                ".ps1", ".psm1", ".vbs", ".vbe", ".js", ".jse",
                ".wsf", ".wsh", ".scr", ".hta", ".reg", ".lnk",
            }
            ext = Path(filename).suffix.lower()
            if ext in risky_exts:
                warning = (
                    "実行可能ファイルまたはスクリプト形式の添付です。\n"
                    "安全性を確認した場合のみ開いてください。\n\n"
                    + filename + "\n\n開きますか？"
                )
                if not messagebox.askyesno("添付ファイルの警告", warning, parent=dialog):
                    return

            try:
                preview_dir = Path(tempfile.gettempdir()) / "INAS_Mail_Archive_Preview"
                preview_dir.mkdir(parents=True, exist_ok=True)
                preview_path = unique_path(preview_dir / sanitize_filename(filename, "attachment"))
                preview_path.write_bytes(payload)
                os.startfile(str(preview_path))
            except OSError as exc:
                logging.exception("添付ファイルの関連付け起動に失敗しました: %s", filename)
                messagebox.showerror(
                    "添付ファイル",
                    "添付ファイルを開けませんでした。\n\n" + filename +
                    "\n\nWindowsでこの拡張子を開くアプリが関連付けされているか確認してください。\n\n" + str(exc),
                    parent=dialog,
                )
            except Exception as exc:
                logging.exception("添付ファイルのプレビューに失敗しました: %s", filename)
                messagebox.showerror("添付ファイル", "添付ファイルを開けませんでした。\n\n" + str(exc), parent=dialog)

        tree.bind("<Double-1>", open_attachment)
        ttk.Button(attachment_header, text="添付を開く", command=open_attachment).pack(side="right", padx=(6, 0))

        attachment_expanded = tk.BooleanVar(value=False)
        body_expanded = tk.BooleanVar(value=False)
        attachment_toggle_btn = ttk.Button(attachment_header)

        def toggle_attachment_area():
            expanded = not attachment_expanded.get()
            attachment_expanded.set(expanded)
            if expanded:
                # 添付を優先する場合は本文拡大を解除。
                body_expanded.set(False)
                body_toggle_btn.configure(text="▼ 本文を拡大")
                outer.rowconfigure(2, weight=1, minsize=70)
                outer.rowconfigure(4, weight=5, minsize=180)
                tree.configure(height=12)
                attachment_toggle_btn.configure(text="▲ 添付を縮小")
            else:
                outer.rowconfigure(2, weight=4, minsize=110)
                outer.rowconfigure(4, weight=1, minsize=72)
                tree.configure(height=3)
                attachment_toggle_btn.configure(text="▼ 添付を拡大")
            dialog.update_idletasks()

        attachment_toggle_btn.configure(text="▼ 添付を拡大", command=toggle_attachment_area)
        attachment_toggle_btn.pack(side="right")

        # 保存先（固定領域）


        if external_links:
            external_frame = ttk.LabelFrame(
                outer,
                text=f"外部ファイル便・ダウンロードリンク（{len(external_links)}件）",
                padding=6,
            )
            external_frame.grid(row=6, column=0, sticky="ew", pady=(4, 8))
            ttk.Label(
                external_frame,
                text="メール本文に外部ダウンロード候補を検出しました。",
            ).pack(side="left")
            ttk.Button(
                external_frame,
                text="リンクを確認",
                command=lambda: self.show_external_download_links(dialog, external_links),
            ).pack(side="right")

        save_frame = ttk.LabelFrame(outer, text="保存先", padding=8)
        save_frame.grid(row=5, column=0, sticky="ew", pady=(10, 0))
        save_frame.columnconfigure(0, weight=1)
        default_output = str(self.settings.get("default_output_folder", "")).strip()
        initial_output = self.var_output.get().strip() or default_output
        preview_output = tk.StringVar(value=initial_output)
        output_entry = ttk.Entry(save_frame, textvariable=preview_output)
        output_entry.grid(row=0, column=0, sticky="ew", padx=(0, 6))

        def browse_output():
            initial = preview_output.get().strip()
            folder = filedialog.askdirectory(parent=dialog, initialdir=initial if initial and Path(initial).is_dir() else None)
            if folder:
                preview_output.set(folder)

        ttk.Button(save_frame, text="変更...", command=browse_output).grid(row=0, column=1)

        registered = self.get_registered_save_locations()
        registered_display = [""] + [f'{item["name"]}  |  {item["path"]}' for item in registered]
        registered_var = tk.StringVar(value="")
        ttk.Label(save_frame, text="登録済み保存先").grid(row=1, column=0, sticky="w", pady=(8, 2))
        registered_combo = ttk.Combobox(save_frame, textvariable=registered_var, values=registered_display, state="readonly")
        registered_combo.grid(row=2, column=0, sticky="ew", padx=(0, 6))

        def apply_registered(_event=None):
            value = registered_var.get()
            if not value:
                return
            for item in self.get_registered_save_locations():
                display = f'{item["name"]}  |  {item["path"]}'
                if display == value:
                    preview_output.set(item["path"])
                    return

        registered_combo.bind("<<ComboboxSelected>>", apply_registered)


        recent_candidates = self.recent_destination_candidates(data)
        recent_var = tk.StringVar(value="")
        ttk.Label(save_frame, text="最近の保存先候補").grid(row=3, column=0, sticky="w", pady=(8, 2))
        recent_combo = ttk.Combobox(
            save_frame,
            textvariable=recent_var,
            values=[""] + recent_candidates,
            state="readonly",
        )
        recent_combo.grid(row=4, column=0, sticky="ew", padx=(0, 6))

        def apply_recent(_event=None):
            if recent_var.get():
                preview_output.set(recent_var.get())

        recent_combo.bind("<<ComboboxSelected>>", apply_recent)

        def manage_registered():
            def selected(path_value):
                preview_output.set(path_value)
            dlg = self.show_registered_save_locations(parent=dialog, on_select=selected)

            def refresh_after_close(_event=None):
                values = [""] + [f'{item["name"]}  |  {item["path"]}' for item in self.get_registered_save_locations()]
                try:
                    registered_combo.configure(values=values)
                except Exception:
                    pass

            dlg.bind("<Destroy>", refresh_after_close)

        ttk.Button(save_frame, text="登録・管理...", command=manage_registered).grid(row=2, column=1)

        # 下部操作ボタン（常時固定表示）
        buttons = ttk.Frame(outer)
        buttons.grid(row=6, column=0, sticky="ew", pady=(12, 0))

        def open_eml():
            try:
                os.startfile(str(path))
            except Exception as exc:
                messagebox.showerror("EMLを開く", f"EMLを開けませんでした。\n\n{exc}", parent=dialog)

        def do_discard():
            result["action"] = "discard"
            dialog.destroy()

        def do_save():
            output = preview_output.get().strip()
            if not output:
                messagebox.showerror("保存先", "保存先を指定してください。", parent=dialog)
                return
            self.var_output.set(output)
            self._preview_selected_attachment_indexes = [
                idx for idx in range(len(data.attachments))
                if attachment_save_flags.get(idx, True)
            ]
            result["action"] = "save"
            dialog.destroy()

        ttk.Button(buttons, text="EMLを開く", command=open_eml).pack(side="left")
        ttk.Button(buttons, text="添付を開く", command=open_attachment).pack(side="left", padx=(6, 0))
        ttk.Button(buttons, text="保存しない", command=do_discard).pack(side="right")
        ttk.Button(buttons, text="メイン画面へ反映", command=do_save).pack(side="right", padx=(0, 8))

        def on_close():
            result["action"] = "cancel"
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", on_close)
        dialog.wait_window()
        return result["action"]

    def load_eml(self, path: Path):
        try:
            data = parse_eml(path, ignore_inline_images=self.var_ignore_inline.get())
        except Exception as exc:
            logging.exception("EMLを読み込めませんでした: %s", path)
            try:
                if path.exists() and path.resolve().parent == Path(self.var_import.get()).resolve():
                    self.quarantine_error_eml(path, f"通常読込エラー: {exc}")
            except Exception:
                pass
            messagebox.showerror(
                "読込エラー",
                "EMLを読み込めませんでした。Errorフォルダへ隔離しました。\n\n"
                f"{exc}\n\n"
                f"ログ：{LOG_PATH}",
            )
            self.update_monitor_panel()
            return
        excluded, excluded_rule = self.is_excluded_mail(data)
        if excluded:
            logging.info("保存対象外メールを除外しました: %s / rule=%s", path, excluded_rule)
            self.discard_excluded_temp_eml(path)
            self.var_status.set(f"保存対象外：{excluded_rule} / {path.name}")
            self.root.after(100, self.load_next_queued_eml)
            return

        if self.settings.get("confirm_auto_detected_mail", True):
            try:
                is_import_mail = path.resolve().parent == Path(self.var_import.get()).resolve()
            except Exception:
                is_import_mail = False
            if is_import_mail:
                default_output = str(self.settings.get("default_output_folder", "")).strip()
                if default_output and not self.var_output.get().strip():
                    self.var_output.set(default_output)
                preview_action = self.show_mail_preview_dialog(data, path)
                if preview_action == "discard":
                    # 保存不要を選んだ場合は、一時EMLを必ず削除する。
                    # Outlook側の元メールには影響しない。
                    self.mark_eml_seen(path)
                    try:
                        if path.exists() and path.resolve().parent == Path(self.var_import.get()).resolve():
                            path.unlink()
                            logging.info("保存不要の一時EMLを削除しました: %s", path)
                            self.var_status.set(f"保存不要・一時EML削除：{path.name}")
                        else:
                            self.var_status.set(f"今回は保存しません：{path.name}")
                    except Exception:
                        logging.exception("保存不要EMLの一時ファイル削除に失敗しました: %s", path)
                        self.var_status.set(f"保存不要（EML削除失敗）：{path.name}")
                    self.root.after(100, self.load_next_queued_eml)
                    return
                if preview_action != "save":
                    # ×で閉じた場合はEMLを削除せず、今回は処理しない。
                    self.mark_eml_seen(path)
                    self.var_status.set(f"メール確認を中断しました：{path.name}")
                    self.root.after(100, self.load_next_queued_eml)
                    return

        self.mail_data = data
        default_output = str(self.settings.get("default_output_folder", "")).strip()
        if default_output and not self.var_output.get().strip():
            self.var_output.set(default_output)
        try:
            summary = read_eml_summary(path)
        except Exception:
            summary = {}
        raw_subject = str(
            summary.get("subject")
            or self.extract_subject(data.header_block)
            or ""
        ).strip()
        raw_from = str(summary.get("sender") or data.from_name or "").strip()

        # EML読込時は元の件名をそのまま表示する。
        # Re:/Fw:/Fwd: は返信・転送を示す情報なので、自動では除去しない。
        # ユーザーが［整理］を押した場合だけ clean_subject() を適用する。
        self.var_subject.set(raw_subject)
        key = sender_key(data.from_name)
        mapped_name = self.sender_dictionary.get(key)
        self.var_from.set(mapped_name if mapped_name else raw_from)
        self.attachment_label.config(text=f"添付：{len(data.attachments)}件")
        duplicate = self.processed_history.get(data.source_hash)
        if duplicate:
            processed_at = duplicate.get("processed_at", "日時不明")
            self.var_status.set(
                f"読込済み：{path.name}（同じEMLを処理済み：{processed_at}）"
            )
            logging.warning(
                "処理済みEMLを再読込しました: %s / hash=%s",
                path,
                data.source_hash,
            )
        else:
            self.var_status.set(f"読込済み：{path.name}")
            logging.info("EMLを読み込みました: %s", path)
        try:
            stat = path.stat()
            self.known_eml_signatures[path] = (stat.st_size, stat.st_mtime_ns)
        except OSError:
            pass
        self.root.deiconify()
        self.root.lift()
        self.root.attributes("-topmost", True)
        self.root.after(500, lambda: self.root.attributes("-topmost", False))

    @staticmethod
    def extract_subject(header_block: str) -> str:
        for line in header_block.splitlines():
            if line.startswith("Subject:"):
                return line.partition(":")[2].strip()
        return ""

    def ask_duplicate_action(self, duplicate: dict) -> str:
        result = {"action": "cancel"}
        dialog = tk.Toplevel(self.root)
        dialog.title("重複メールの確認")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)
        frame = ttk.Frame(dialog, padding=14)
        frame.pack(fill="both", expand=True)
        ttk.Label(
            frame,
            text=(
                "同一内容のEMLはすでに処理されています。\n\n"
                f"前回処理日時：{duplicate.get('processed_at', '不明')}\n"
                f"前回保存先：{duplicate.get('output_folder', '不明')}"
            ),
            justify="left",
        ).pack(anchor="w")
        ttk.Label(frame, text="処理方法を選択してください。").pack(anchor="w", pady=(12, 8))

        def choose(action):
            result["action"] = action
            dialog.destroy()

        buttons = ttk.Frame(frame)
        buttons.pack(fill="x")
        ttk.Button(buttons, text="処理しない", command=lambda: choose("cancel")).pack(side="left")
        ttk.Button(buttons, text="別名フォルダで保存", command=lambda: choose("separate")).pack(side="left", padx=6)
        ttk.Button(buttons, text="再処理", command=lambda: choose("force")).pack(side="left")
        ttk.Button(buttons, text="保存済みを開く", command=lambda: choose("open")).pack(side="left", padx=(6, 0))
        dialog.protocol("WM_DELETE_WINDOW", lambda: choose("cancel"))
        self.root.wait_window(dialog)
        return result["action"]

    def show_history(self):
        dialog = tk.Toplevel(self.root)
        dialog.title(self.t("history"))
        dialog.geometry("1080x620")
        dialog.minsize(820, 440)
        dialog.transient(self.root)

        frame = ttk.Frame(dialog, padding=10)
        frame.pack(fill="both", expand=True)

        # 検索・絞り込み
        filter_frame = ttk.LabelFrame(frame, text=self.t("search_filter"), padding=8)
        filter_frame.pack(fill="x", pady=(0, 8))

        ttk.Label(filter_frame, text=self.t("keyword")).grid(row=0, column=0, sticky="w")
        var_keyword = tk.StringVar()
        keyword_entry = ttk.Entry(filter_frame, textvariable=var_keyword)
        keyword_entry.grid(row=0, column=1, sticky="ew", padx=(6, 12))

        ttk.Label(filter_frame, text=self.t("period")).grid(row=0, column=2, sticky="w")
        var_period = tk.StringVar(value="すべて")
        period_combo = ttk.Combobox(
            filter_frame,
            textvariable=var_period,
            values=("すべて", "今日", "7日以内", "30日以内", "90日以内"),
            state="readonly",
            width=12,
        )
        period_combo.grid(row=0, column=3, sticky="w", padx=(6, 12))

        var_missing = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            filter_frame,
            text=self.t("missing_only"),
            variable=var_missing,
        ).grid(row=0, column=4, sticky="w")
        filter_frame.columnconfigure(1, weight=1)

        count_var = tk.StringVar(value="0件")
        ttk.Label(filter_frame, textvariable=count_var).grid(row=1, column=0, columnspan=5, sticky="w", pady=(6, 0))

        # 履歴一覧
        tree_frame = ttk.Frame(frame)
        tree_frame.pack(fill="both", expand=True)
        columns = ("time", "sender", "subject", "source", "folder")
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="extended")
        tree.heading("time", text=self.t("processed_datetime"))
        tree.heading("sender", text=self.t("from_label"))
        tree.heading("subject", text=self.t("subject"))
        tree.heading("source", text=self.t("source_eml"))
        tree.heading("folder", text=self.t("save_location"))
        tree.column("time", width=145, anchor="w", stretch=False)
        tree.column("sender", width=130, anchor="w")
        tree.column("subject", width=270, anchor="w")
        tree.column("source", width=170, anchor="w")
        tree.column("folder", width=360, anchor="w")

        yscroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        xscroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        def record_matches(record: dict) -> bool:
            keyword = var_keyword.get().strip().lower()
            if keyword:
                haystack = " ".join(
                    str(record.get(key, ""))
                    for key in ("processed_at", "sender", "subject", "source_name", "output_folder", "pdf_path")
                ).lower()
                if keyword not in haystack:
                    return False

            period = var_period.get()
            if period != "すべて":
                try:
                    processed = datetime.strptime(record.get("processed_at", ""), "%Y/%m/%d %H:%M:%S")
                    now = datetime.now()
                    if period == "今日" and processed.date() != now.date():
                        return False
                    days_map = {"7日以内": 7, "30日以内": 30, "90日以内": 90}
                    if period in days_map and processed < now - timedelta(days=days_map[period]):
                        return False
                except (TypeError, ValueError):
                    return False

            if var_missing.get():
                folder_ok = Path(record.get("output_folder", "")).is_dir()
                pdf_ok = Path(record.get("pdf_path", "")).is_file()
                if folder_ok and pdf_ok:
                    return False
            return True

        visible_hashes: list[str] = []

        def refresh_history(*_args):
            nonlocal visible_hashes
            for item in tree.get_children():
                tree.delete(item)

            records = sorted(
                self.processed_history.items(),
                key=lambda item: item[1].get("processed_at", ""),
                reverse=True,
            )
            visible_hashes = []
            for source_hash, record in records:
                if not record_matches(record):
                    continue
                visible_hashes.append(source_hash)
                tree.insert(
                    "",
                    "end",
                    iid=source_hash,
                    values=(
                        record.get("processed_at", ""),
                        record.get("sender", ""),
                        record.get("subject", ""),
                        record.get("source_name", ""),
                        record.get("output_folder", ""),
                    ),
                )
            count_var.set(f"表示 {len(visible_hashes)}件 / 全 {len(self.processed_history)}件")

        def selected_records():
            result = []
            for source_hash in tree.selection():
                record = self.processed_history.get(source_hash)
                if record:
                    result.append((source_hash, record))
            return result

        def first_selected_record():
            records = selected_records()
            return records[0][1] if records else None

        def open_selected_folder():
            record = first_selected_record()
            if not record:
                return
            path = Path(record.get("output_folder", ""))
            if path.is_dir():
                os.startfile(str(path))
            else:
                messagebox.showerror("処理履歴", "保存フォルダが見つかりません。", parent=dialog)

        def open_selected_pdf():
            record = first_selected_record()
            if not record:
                return
            path = Path(record.get("pdf_path", ""))
            if path.is_file():
                os.startfile(str(path))
            else:
                messagebox.showerror("処理履歴", "PDFファイルが見つかりません。", parent=dialog)

        def copy_selected_path():
            record = first_selected_record()
            if not record:
                return
            path = record.get("output_folder", "")
            self.root.clipboard_clear()
            self.root.clipboard_append(path)
            count_var.set("保存先パスをコピーしました。")

        def delete_selected_history():
            selected = selected_records()
            if not selected:
                messagebox.showinfo("処理履歴", "削除する履歴を選択してください。", parent=dialog)
                return
            if not messagebox.askyesno(
                "処理履歴を削除",
                f"選択した {len(selected)} 件の履歴を削除しますか？\n\n"
                "保存済みのPDF・添付ファイル・EMLは削除されません。\n"
                "履歴を削除すると、そのメールは重複判定の対象外になります。",
                parent=dialog,
            ):
                return
            for source_hash, _record in selected:
                self.processed_history.pop(source_hash, None)
            save_processed_history(self.processed_history)
            refresh_history()

        def clear_all_history():
            if not self.processed_history:
                return
            if not messagebox.askyesno(
                "全履歴を削除",
                "すべての処理履歴を削除しますか？\n\n"
                "保存済みファイルは削除されませんが、過去メールの重複判定ができなくなります。",
                parent=dialog,
            ):
                return
            self.processed_history.clear()
            save_processed_history(self.processed_history)
            refresh_history()

        def export_visible_csv():
            if not visible_hashes:
                messagebox.showinfo("CSV出力", "出力対象の履歴がありません。", parent=dialog)
                return
            default_name = f"INAS_Mail_Archive_History_{time.strftime('%Y%m%d_%H%M%S')}.csv"
            path = filedialog.asksaveasfilename(
                parent=dialog,
                title="処理履歴をCSV出力",
                defaultextension=".csv",
                initialfile=default_name,
                filetypes=[("CSVファイル", "*.csv"), ("すべてのファイル", "*.*")],
            )
            if not path:
                return
            try:
                with open(path, "w", newline="", encoding="utf-8-sig") as f:
                    writer = csv.writer(f)
                    writer.writerow(["処理日時", "差出人", "件名", "元EML", "保存先", "PDFファイル", "SHA-256"])
                    for source_hash in visible_hashes:
                        record = self.processed_history.get(source_hash, {})
                        writer.writerow([
                            record.get("processed_at", ""),
                            record.get("sender", ""),
                            record.get("subject", ""),
                            record.get("source_name", ""),
                            record.get("output_folder", ""),
                            record.get("pdf_path", ""),
                            source_hash,
                        ])
                messagebox.showinfo("CSV出力", f"処理履歴を出力しました。\n{path}", parent=dialog)
            except Exception as exc:
                logging.exception("処理履歴CSV出力に失敗しました")
                messagebox.showerror("CSV出力", f"CSV出力に失敗しました。\n{exc}", parent=dialog)

        buttons = ttk.Frame(frame)
        buttons.pack(fill="x", pady=(8, 0))
        ttk.Button(buttons, text=self.t("open_save_folder"), command=open_selected_folder).pack(side="left")
        ttk.Button(buttons, text=self.t("open_pdf"), command=open_selected_pdf).pack(side="left", padx=6)
        ttk.Button(buttons, text=self.t("copy_path"), command=copy_selected_path).pack(side="left")
        ttk.Separator(buttons, orient="vertical").pack(side="left", fill="y", padx=8)
        ttk.Button(buttons, text=self.t("export_visible_csv"), command=export_visible_csv).pack(side="left")
        ttk.Button(buttons, text=self.t("delete_selected_history"), command=delete_selected_history).pack(side="left", padx=6)
        ttk.Button(buttons, text=self.t("delete_all_history"), command=clear_all_history).pack(side="left")
        ttk.Button(buttons, text=self.t("close"), command=dialog.destroy).pack(side="right")

        var_keyword.trace_add("write", refresh_history)
        period_combo.bind("<<ComboboxSelected>>", refresh_history)
        var_missing.trace_add("write", refresh_history)
        tree.bind("<Double-1>", lambda _e: open_selected_folder())
        dialog.bind("<Control-f>", lambda _e: keyword_entry.focus_set())
        keyword_entry.focus_set()
        refresh_history()

    def show_completion(self, final_folder: Path, pdf_path: Path):
        dialog = tk.Toplevel(self.root)
        dialog.title(self.t("save_complete"))
        dialog.transient(self.root)
        dialog.grab_set()
        frame = ttk.Frame(dialog, padding=14)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text=self.t("saved_pdf_attachments"), font=("Yu Gothic UI", 10, "bold")).pack(anchor="w")
        ttk.Label(frame, text=str(final_folder), wraplength=620).pack(anchor="w", pady=(8, 12))

        def copy_path():
            self.root.clipboard_clear()
            self.root.clipboard_append(str(final_folder))
            self.var_status.set("保存先パスをクリップボードへコピーしました。")

        buttons = ttk.Frame(frame)
        buttons.pack(fill="x")
        ttk.Button(buttons, text="フォルダを開く", command=lambda: os.startfile(str(final_folder))).pack(side="left")
        ttk.Button(buttons, text=self.t("open_pdf"), command=lambda: os.startfile(str(pdf_path))).pack(side="left", padx=6)
        ttk.Button(buttons, text=self.t("copy_path"), command=copy_path).pack(side="left")
        ttk.Button(buttons, text=self.t("close"), command=dialog.destroy).pack(side="right")
        self.root.wait_window(dialog)

    def execute(self, skip_confirmation: bool = False, suppress_completion: bool = False):
        if not self.mail_data:
            messagebox.showerror("エラー", "先にEMLを読み込んでください。")
            return

        duplicate_mode = "force"
        duplicate = self.processed_history.get(self.mail_data.source_hash)
        if duplicate:
            duplicate_mode = self.ask_duplicate_action(duplicate)
            if duplicate_mode == "cancel":
                logging.info("重複EMLの再処理をキャンセルしました: %s", self.mail_data.source_path)
                return
            if duplicate_mode == "open":
                previous_output = Path(duplicate.get("output_folder", ""))
                if previous_output.is_dir():
                    os.startfile(str(previous_output))
                else:
                    messagebox.showerror("重複確認", "前回の保存先が見つかりません。")
                return

        output_text = self.var_output.get().strip()
        if not output_text:
            messagebox.showerror("エラー", "保存先を指定してください。")
            return
        output_root = Path(output_text)
        if not output_root.is_dir():
            messagebox.showerror("エラー", "保存先フォルダが無効です。")
            return

        folder_name = sanitize_filename(self.var_folder.get(), "mail")
        pdf_name = sanitize_filename(self.var_pdf.get(), "mail.pdf")
        if not pdf_name.lower().endswith(".pdf"):
            pdf_name += ".pdf"
        final_folder = output_root / folder_name
        if duplicate_mode == "separate":
            final_folder = unique_folder_path(final_folder)

        preview = (
            f"保存先：{final_folder}\n"
            f"PDF：{pdf_name}\n"
            f"添付（個別保存）：{len(getattr(self, '_preview_selected_attachment_indexes', list(range(len(self.mail_data.attachments)))))} / {len(self.mail_data.attachments)}件\n"
            f"{self.t('original_eml_label')}：{dict(move_to_output=self.t('move_to_folder'), keep=self.t('keep_in_import'), delete=self.t('delete')).get(self.var_post_action.get(), self.var_post_action.get())}\n\n"
            + self.t("save_with_content")
        )
        if not skip_confirmation:
            if not messagebox.askyesno(self.t("save_confirmation"), preview):
                return

        try:
            final_folder.mkdir(parents=True, exist_ok=True)
            pdf_path = unique_path(final_folder / pdf_name)
            create_pdf(
                pdf_path,
                self.mail_data.header_block,
                self.mail_data.body_text,
                self.mail_data.attachments,
                display_subject=self.var_subject.get().strip(),
                display_sender_name=self.var_from.get().strip(),
                display_options=self.settings,
                pdf_language=(self.var_language.get() if self.settings.get("pdf_language", "same") == "same" else self.settings.get("pdf_language", "ja")),
            )

            selected_attachment_indexes = set(
                getattr(
                    self,
                    "_preview_selected_attachment_indexes",
                    range(len(self.mail_data.attachments)),
                )
            )
            for idx, (filename, data) in enumerate(self.mail_data.attachments):
                if idx not in selected_attachment_indexes:
                    continue
                attachment_path = unique_path(final_folder / sanitize_filename(filename, "attachment"))
                attachment_path.write_bytes(data)

            source_name = self.mail_data.source_path.name
            source_hash = self.mail_data.source_hash
            subject = self.var_subject.get().strip()
            sender = self.var_from.get().strip()
            try:
                _ok, _reason, health = self.validate_eml_file(self.mail_data.source_path)
            except Exception:
                health = {}
            self.handle_original_eml(final_folder)

            message_id = str(health.get("message_id", "") or "").strip().lower()

            self.processed_history[source_hash] = {
                "processed_at": time.strftime("%Y/%m/%d %H:%M:%S"),
                "source_name": source_name,
                "sender": sender,
                "subject": subject,
                "output_folder": str(final_folder),
                "pdf_path": str(pdf_path),
                "message_id": message_id,
            }
            save_processed_history(self.processed_history)

            if message_id:
                registry = load_message_id_registry()
                registry[message_id] = {
                    "processed_at": self.processed_history[source_hash]["processed_at"],
                    "source_hash": source_hash,
                    "source_name": source_name,
                    "sender": sender,
                    "subject": subject,
                    "output_folder": str(final_folder),
                    "pdf_path": str(pdf_path),
                }
                if len(registry) > 10000:
                    # preserve newest insertion order entries
                    registry = dict(list(registry.items())[-8000:])
                save_message_id_registry(registry)

            self.persist_settings()
            self.var_status.set(f"保存完了：{final_folder}")
            logging.info(
                "保存完了: source=%s / pdf=%s / attachments=%d",
                source_name,
                pdf_path,
                len(self.mail_data.attachments),
            )
            if self.var_open_after.get():
                os.startfile(str(final_folder))
            if not suppress_completion:
                self.show_completion(final_folder, pdf_path)

            return_to_pending = bool(getattr(self, "_return_to_pending_after_execute", False))
            self._return_to_pending_after_execute = False
            self._pending_source_path = None
            self.clear_form(load_next=(not suppress_completion and not return_to_pending))

            if return_to_pending:
                # 保存完了後は未処理一覧を再読込して次のメール処理へ戻る。
                refresh_cb = getattr(self, "_pending_refresh_callback", None)
                if callable(refresh_cb):
                    try:
                        refresh_cb()
                    except Exception:
                        logging.exception("未処理一覧の再読込に失敗しました")
                self.root.after(100, self.show_pending_mail_list)
        except Exception as exc:
            logging.exception(
                "保存処理に失敗しました: %s",
                self.mail_data.source_path if self.mail_data else "",
            )
            messagebox.showerror(
                "保存エラー",
                "保存処理に失敗しました。\n\n"
                f"{exc}\n\n"
                f"ログ：{LOG_PATH}",
            )

    def handle_original_eml(self, final_folder: Path):
        source = self.mail_data.source_path
        action = self.var_post_action.get()
        if action == "keep" or not source.exists():
            return
        if action == "delete":
            source.unlink()
        elif action == "move_to_output":
            destination = unique_path(final_folder / sanitize_filename(source.name, "original.eml"))
            try:
                shutil.move(str(source), str(destination))
            except PermissionError:
                time.sleep(0.5)
                shutil.move(str(source), str(destination))

    def clear_form(self, load_next: bool = False):
        self._preview_selected_attachment_indexes = []
        self.mail_data = None
        self.var_from.set("")
        self.var_subject.set("")
        self.var_pdf.set("")
        self.var_folder.set("")
        self.var_output.set("")
        self.attachment_label.config(text=self.t("attachment_zero"))
        if load_next and self.eml_queue:
            self.var_status.set(f"次のEMLを読み込みます。待機：{len(self.eml_queue)}件")
            self.root.after(100, self.load_next_queued_eml)
        else:
            waiting = f" 待機EML：{len(self.eml_queue)}件" if self.eml_queue else ""
            self.var_status.set("EMLを取込フォルダへ保存するか、画面へドロップしてください。" + waiting)
        self.update_monitor_panel()

    def pending_eml_count(self) -> int:
        count = len(self.eml_queue) + len(self.pending_paths)
        if self.mail_data is not None:
            count += 1
        return count

    def update_monitor_panel(self):
        if not hasattr(self, "var_monitor_state"):
            return
        active = bool(self.observer and self.observer.is_alive() and self.var_watch.get())
        self.var_monitor_state.set(self.t("monitor_active") if active else self.t("monitor_stopped"))
        if hasattr(self, "monitor_state_label"):
            self.monitor_state_label.configure(foreground="#237A3B" if active else "#8A4B4B")
        if self.last_detected_at:
            self.var_last_detected.set(self.t("last_detected", time=self.last_detected_at.strftime("%Y/%m/%d %H:%M:%S")))
        else:
            self.var_last_detected.set(self.t("last_detected_none"))
        self.var_pending_count.set(self.t("pending_eml_count", count=self.pending_eml_count()))

    def rescan_import_folder(self):
        try:
            folder = Path(self.var_import.get())
            folder.mkdir(parents=True, exist_ok=True)
            paths = sorted(folder.glob("*.eml"), key=lambda p: p.stat().st_mtime_ns if p.exists() else 0)
            added = []
            for path in paths:
                if self.mail_data and path == self.mail_data.source_path:
                    continue
                if path in self.queued_paths or path in self.pending_paths:
                    continue
                added.append(path)
            if added:
                self.last_detected_at = datetime.now()
                self.queue_emls(added)
            self.var_status.set(self.t("rescan_complete", count=len(added)))
            self.update_monitor_panel()
            logging.info("取込フォルダを再スキャンしました: %s / %d件", folder, len(added))
        except Exception as exc:
            logging.exception("取込フォルダの再スキャンに失敗しました")
            self.var_status.set(self.t("rescan_error", error=str(exc)))
            self.update_monitor_panel()

    def toggle_watcher(self):
        if self.var_watch.get():
            self.start_watcher()
        else:
            self.stop_watcher()
        self.persist_settings()

    def start_watcher(self):
        if self.observer and self.observer.is_alive():
            return
        folder = Path(self.var_import.get())
        folder.mkdir(parents=True, exist_ok=True)
        self.observer = Observer()
        self.observer.schedule(EmlCreatedHandler(self.on_eml_detected), str(folder), recursive=False)
        self.observer.daemon = True
        self.observer.start()
        self.var_status.set(f"監視中：{folder}")
        self.update_monitor_panel()

    def stop_watcher(self):
        if self.observer:
            self.observer.stop()
            self.observer.join(timeout=2)
            self.observer = None
        self.var_status.set("自動監視を停止しています。")
        self.update_monitor_panel()

    def restart_watcher(self):
        self.stop_watcher()
        if self.var_watch.get():
            self.start_watcher()

    def on_eml_detected(self, path: Path):
        path = Path(path)
        if path in self.pending_paths or path in self.queued_paths:
            return
        if self.mail_data and path == self.mail_data.source_path:
            return
        if self.is_recent_duplicate_detection(path):
            return
        self.last_detected_at = datetime.now()
        self.pending_paths.add(path)
        self.update_monitor_panel()
        threading.Thread(target=self.wait_until_stable, args=(path,), daemon=True).start()

    def wait_until_stable(self, path: Path):
        try:
            last_size = -1
            stable_count = 0
            for _ in range(30):
                if not path.exists():
                    time.sleep(0.3)
                    continue
                size = path.stat().st_size
                if size > 0 and size == last_size:
                    stable_count += 1
                    if stable_count >= 3:
                        # Mark before crossing to the Tk thread.  This closes the
                        # watchdog/polling race window while OneDrive finishes sync.
                        self.mark_eml_seen(path)
                        def _handle_stable_eml(p=path):
                            try:
                                accepted, preflight_reason = self.preflight_eml(p, quarantine=True)
                                if not accepted:
                                    logging.info("新着EMLを通常処理から除外: %s (%s)", p.name, preflight_reason)
                                    self.update_monitor_panel()
                                    return
                            except Exception:
                                logging.exception("新着EMLの事前確認に失敗: %s", p)
                                if p.exists():
                                    self.quarantine_error_eml(p, "新着EML事前確認例外")
                                self.update_monitor_panel()
                                return
                            self.queue_emls([p], auto_detected=True)

                        self.root.after(0, _handle_stable_eml)
                        return
                else:
                    stable_count = 0
                last_size = size
                time.sleep(0.5)
        finally:
            self.pending_paths.discard(path)
            try:
                self.root.after(0, self.update_monitor_panel)
            except Exception:
                pass

    def poll_import_folder(self):
        """watchdogが取りこぼしたEMLも定期確認で検出する。"""
        try:
            if self.var_watch.get():
                folder = Path(self.var_import.get())
                folder.mkdir(parents=True, exist_ok=True)
                current: dict[Path, tuple[int, int]] = {}
                candidates: list[tuple[int, Path]] = []

                for path in folder.glob("*.eml"):
                    try:
                        stat = path.stat()
                    except OSError:
                        continue
                    signature = (stat.st_size, stat.st_mtime_ns)
                    current[path] = signature
                    if self.known_eml_signatures.get(path) != signature:
                        candidates.append((stat.st_mtime_ns, path))

                # 起動時・監視中とも、未取得のEMLを古い順にキューへ追加する。
                if candidates:
                    candidates.sort()
                    for _mtime, candidate in candidates:
                        self.on_eml_detected(candidate)

                self.known_eml_signatures = current
        except Exception as exc:
            logging.exception("監視確認エラー")
            self.var_status.set(f"監視確認エラー：{exc}")
        finally:
            if not self.is_exiting:
                self.poll_job = self.root.after(1000, self.poll_import_folder)

    def persist_settings(self):
        self.settings.update({
            "import_folder": self.var_import.get(),
            "watch_enabled": self.var_watch.get(),
            "post_action": self.var_post_action.get(),
            "open_after_save": self.var_open_after.get(),
            "auto_start": self.var_auto_start.get(),
            "resident_enabled": self.var_resident.get(),
            "ignore_inline_images": self.var_ignore_inline.get(),
            "default_output_folder": self.settings.get("default_output_folder", str(DEFAULT_OUTPUT)),
            "excluded_addresses": self.settings.get("excluded_addresses", []),
            "excluded_domains": self.settings.get("excluded_domains", []),
            "confirm_auto_detected_mail": bool(self.settings.get("confirm_auto_detected_mail", True)),
            "delete_excluded_temp_eml": bool(self.settings.get("delete_excluded_temp_eml", True)),
            "registered_save_locations": self.get_registered_save_locations(),
            "duplicate_check_enabled": bool(self.settings.get("duplicate_check_enabled", True)),
            "recent_destination_suggestions": bool(self.settings.get("recent_destination_suggestions", True)),
            "attachment_selective_save": bool(self.settings.get("attachment_selective_save", True)),
            "external_link_detection_enabled": bool(self.settings.get("external_link_detection_enabled", True)),
            "auto_show_pending_on_new_eml": bool(self.settings.get("auto_show_pending_on_new_eml", True)),
            "pdf_name_template": self.settings.get("pdf_name_template", DEFAULT_PDF_TEMPLATE),
            "folder_name_template": self.settings.get("folder_name_template", DEFAULT_FOLDER_TEMPLATE),
            **{key: bool(self.settings.get(key, default)) for key, default in PDF_DISPLAY_DEFAULTS.items()},
        })
        save_settings(self.settings)

    @staticmethod
    def open_folder(path: Path):
        path.mkdir(parents=True, exist_ok=True)
        os.startfile(str(path))

    @staticmethod
    def startup_command() -> str:
        """現在の実行形態に応じたWindows自動起動コマンドを返す。"""
        if getattr(sys, "frozen", False):
            return f'"{sys.executable}" --startup'

        script_path = Path(__file__).resolve()
        python_exe = Path(sys.executable)
        pythonw_exe = python_exe.with_name("pythonw.exe")
        launcher = pythonw_exe if pythonw_exe.exists() else python_exe
        return f'"{launcher}" "{script_path}" --startup'

    def set_windows_startup(self, enabled: bool) -> None:
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            key_path,
            0,
            winreg.KEY_SET_VALUE,
        ) as key:
            if enabled:
                winreg.SetValueEx(
                    key,
                    STARTUP_VALUE_NAME,
                    0,
                    winreg.REG_SZ,
                    self.startup_command(),
                )
            else:
                try:
                    winreg.DeleteValue(key, STARTUP_VALUE_NAME)
                except FileNotFoundError:
                    pass

    def sync_startup_setting(self, show_error: bool = True) -> None:
        try:
            self.set_windows_startup(self.var_auto_start.get())
        except OSError as exc:
            self.var_auto_start.set(False)
            if show_error:
                messagebox.showerror(
                    "自動起動設定エラー",
                    f"Windows自動起動の設定を変更できませんでした。\n\n{exc}",
                )

    def on_auto_start_changed(self):
        self.sync_startup_setting(show_error=True)
        self.persist_settings()

    @staticmethod
    def create_tray_image() -> Image.Image:
        try:
            return Image.open(APP_ICON_PNG_PATH).convert("RGBA")
        except Exception:
            logging.exception("タスクトレイ用アイコンを読み込めませんでした")
            image = Image.new("RGB", (64, 64), "white")
            draw = ImageDraw.Draw(image)
            draw.rounded_rectangle((5, 9, 59, 55), radius=7, outline="#2457A6", width=4)
            draw.line((8, 13, 32, 34, 56, 13), fill="#2457A6", width=4)
            draw.text((20, 36), "IN", fill="#2457A6")
            return image

    def start_tray_icon(self):
        if self.tray_icon is not None:
            return

        menu = pystray.Menu(
            pystray.MenuItem(self.t("tray_show"), self.tray_show_window, default=True),
            pystray.MenuItem(
                self.t("toggle_monitoring"),
                self.tray_toggle_watcher,
                checked=lambda _item: bool(self.var_watch.get()),
            ),
            pystray.MenuItem(self.t("menu_open_import_folder"), self.tray_open_import_folder),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(self.t("exit_short"), self.tray_exit),
        )
        self.tray_icon = pystray.Icon(
            "inas_mail_archive",
            self.create_tray_image(),
            f"{APP_NAME} Community Edition Ver.{APP_VERSION}",
            menu,
        )
        self.tray_thread = threading.Thread(target=self.tray_icon.run, daemon=True)
        self.tray_thread.start()

    def stop_tray_icon(self):
        if self.tray_icon:
            try:
                self.tray_icon.stop()
            except Exception:
                logging.exception("タスクトレイアイコンの停止に失敗しました")
            self.tray_icon = None
            self.tray_thread = None

    def tray_show_window(self, _icon=None, _item=None):
        self.root.after(0, self.show_window)

    def show_window(self):
        self.root.deiconify()
        self.root.state("normal")
        self.root.lift()
        self.root.attributes("-topmost", True)
        self.root.after(400, lambda: self.root.attributes("-topmost", False))

    def tray_toggle_watcher(self, _icon=None, _item=None):
        self.root.after(0, self.toggle_watcher_from_tray)

    def toggle_watcher_from_tray(self):
        self.var_watch.set(not self.var_watch.get())
        self.toggle_watcher()
        if self.tray_icon:
            self.tray_icon.update_menu()

    def tray_open_import_folder(self, _icon=None, _item=None):
        self.root.after(0, lambda: self.open_folder(Path(self.var_import.get())))

    def tray_exit(self, _icon=None, _item=None):
        self.root.after(0, self.exit_application)

    def on_close(self):
        """右上の×では画面だけ閉じ、常駐設定が有効なら監視を継続する。"""
        self.persist_settings()
        if self.var_resident.get():
            if self.tray_icon is None:
                self.start_tray_icon()
            self.root.withdraw()
            return
        self.exit_application()

    def exit_application(self):
        if self.is_exiting:
            return
        self.is_exiting = True
        logging.info("アプリを終了します")
        self.persist_settings()
        self.stop_watcher()
        self.stop_tray_icon()
        release_single_instance_mutex()
        self.root.destroy()

    def run(self):
        self.root.mainloop()


def hide_console_window_for_frozen_app():
    """Keep the executable as CONSOLE subsystem, but hide its console window after startup.

    INAS Mail Archive の正式ビルド方式として使用する。
    PyInstaller は --console のままにし、--windowed は使用しない。
    起動後にコンソールウィンドウだけを非表示にする。
    """
    if os.name != "nt" or not getattr(sys, "frozen", False):
        return
    try:
        import ctypes

        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd:
            ctypes.windll.user32.ShowWindow(hwnd, 0)  # SW_HIDE
    except Exception:
        # Console hiding must never prevent the application from starting.
        pass


if __name__ == "__main__":
    hide_console_window_for_frozen_app()
    if not acquire_single_instance_mutex():
        # A resident copy is already monitoring the import folder.
        try:
            temp_root = tk.Tk()
            temp_root.withdraw()
            messagebox.showinfo(
                "INAS Mail Archive",
                "INAS Mail Archive はすでに起動しています。\n\n"
                "タスクトレイのアイコンから既存の画面を開いてください。",
                parent=temp_root,
            )
            temp_root.destroy()
        except Exception:
            pass
        raise SystemExit(0)
    MailArchiveApp().run()
