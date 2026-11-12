import os
import re
import time
import base64
import json
import html as html_lib
import tempfile
import threading
from datetime import datetime
from urllib.parse import parse_qs, unquote, urlencode, urljoin, urlparse, urlunparse

import streamlit as st
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from curl_cffi import requests

st.set_page_config(
    page_title="TikTok Comment Extractor",
    page_icon="⚡",
    layout="centered",
)

EXCEL_FILE = "tiktok_comment_info_API_Profile.xlsx"
ERROR_EXCEL_FILE = "tiktok_comment_errors.xlsx"
EXCEL_LOCK = threading.RLock()
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
)
REQUEST_TIMEOUT = 15
MAX_COMMENT_PAGES = 100
MAX_RETRIES = 3
TIKWM_MIN_INTERVAL = 1.15
_TIKWM_RATE_LOCK = threading.Lock()
_TIKWM_LAST_REQUEST = 0.0
_CACHE_LOCK = threading.RLock()
_CID_CACHE = None
_APP_DIR = os.path.dirname(os.path.abspath(__file__))
_CACHE_DIR = os.path.join(_APP_DIR, "work")
_CID_CACHE_FILE = os.path.join(_CACHE_DIR, "tiktok_cid_position_cache.json")


def _cache_key(video_id, target_id):
    normalized = _decode_numeric_base64(target_id) or str(target_id or "").strip()
    return f"{video_id}:{normalized}"


def _read_cid_cache():
    global _CID_CACHE
    if _CID_CACHE is not None:
        return _CID_CACHE
    try:
        with open(_CID_CACHE_FILE, "r", encoding="utf-8") as cache_file:
            value = json.load(cache_file)
            _CID_CACHE = value if isinstance(value, dict) else {}
    except (OSError, ValueError):
        _CID_CACHE = {}
    return _CID_CACHE


def _get_cid_hint(video_id, target_id, provider):
    if not video_id or not target_id:
        return {}
    with _CACHE_LOCK:
        entry = _read_cid_cache().get(_cache_key(video_id, target_id), {})
        hint = entry.get(provider, {}) if isinstance(entry, dict) else {}
        if not isinstance(hint, dict) or time.time() - hint.get("updated", 0) > 30 * 86400:
            return {}
        return dict(hint)


def _save_cid_hint(video_id, target_id, provider, **values):
    if not video_id or not target_id:
        return
    with _CACHE_LOCK:
        cache = _read_cid_cache()
        key = _cache_key(video_id, target_id)
        entry = cache.setdefault(key, {})
        hint = entry.setdefault(provider, {})
        hint.update(values)
        hint["updated"] = time.time()
        try:
            os.makedirs(_CACHE_DIR, exist_ok=True)
            temporary_file = _CID_CACHE_FILE + ".tmp"
            with open(temporary_file, "w", encoding="utf-8") as cache_file:
                json.dump(cache, cache_file, ensure_ascii=False)
            os.replace(temporary_file, _CID_CACHE_FILE)
        except OSError:
            pass


def _wait_for_tikwm_slot():
    """Respect TikWM's observed one-request-per-second free-tier limit."""
    global _TIKWM_LAST_REQUEST
    with _TIKWM_RATE_LOCK:
        wait = TIKWM_MIN_INTERVAL - (time.monotonic() - _TIKWM_LAST_REQUEST)
        if wait > 0:
            time.sleep(wait)
        _TIKWM_LAST_REQUEST = time.monotonic()


def response_json(response):
    """Decode JSON as UTF-8 explicitly; avoid requests' ASCII guessing."""
    content = response.content or b""
    if isinstance(content, str):
        content = content.encode("utf-8")
    try:
        body = content.decode("utf-8-sig")
    except UnicodeDecodeError:
        body = content.decode("utf-8", errors="replace")
    return json.loads(body)


def create_tiktok_session(logger):
    """Warm one browser-like HTTP session so TikTok returns its JSON feed."""
    session = requests.Session()
    session.headers.update({
        "User-Agent": USER_AGENT,
        "Accept-Language": "vi-VN,vi;q=0.9,en-US;q=0.8",
        "Accept": "text/html,application/xhtml+xml,application/json;q=0.9,*/*;q=0.8",
    })
    try:
        response = session.get(
            "https://www.tiktok.com/",
            timeout=REQUEST_TIMEOUT,
            impersonate="chrome120",
        )
        logger(f"   + Đã khởi tạo session TikTok (HTTP {response.status_code}).")
    except Exception as exc:
        logger(f"   ! Không khởi tạo được session TikTok: {exc}")
    return session


def _decode_numeric_base64(value):
    """Return the numeric ID represented by a URL-safe Base64 value, if any."""
    if not value:
        return None
    try:
        encoded = str(value).replace("-", "+").replace("_", "/")
        encoded += "=" * ((4 - len(encoded) % 4) % 4)
        decoded = base64.b64decode(encoded, validate=True).decode("utf-8")
        return decoded if decoded.isdigit() else None
    except (ValueError, UnicodeDecodeError):
        return None


def comment_id_variants(value):
    """Create comparable forms without ever matching on comment text or author."""
    if value is None:
        return set()
    raw = str(value).strip()
    if not raw:
        return set()
    variants = {raw}
    decoded = _decode_numeric_base64(raw)
    if decoded:
        variants.add(decoded)
    if raw.isdigit():
        variants.add(base64.b64encode(raw.encode("utf-8")).decode("ascii").rstrip("="))
    return variants


def expand_short_url(session, url, logger):
    """Resolve TikTok short links; keep the original if resolution is blocked."""
    url = html_lib.unescape(str(url).strip())
    host = (urlparse(url).hostname or "").lower()
    short_host = (
        host in {"vt.tiktok.com", "vm.tiktok.com"}
        or host.endswith(".vt.tiktok.com")
        or host.endswith(".vm.tiktok.com")
        or (host.endswith("tiktok.com") and "/t/" in urlparse(url).path)
    )
    if not short_host:
        return url
    try:
        response = session.get(
            url,
            headers={"User-Agent": USER_AGENT},
            allow_redirects=True,
            timeout=REQUEST_TIMEOUT,
            impersonate="chrome120",
        )
        final_url = html_lib.unescape(response.url or url)
        if final_url != url:
            logger(f"   + URL sau khi mở rộng: {final_url[:240]}")
            return final_url
        page = response.text or ""
        patterns = (
            r'<meta[^>]+content=["\'][^"\']*?url=([^"\']+)',
            r'window\.location\.replace\(["\']([^"\']+)["\']\)',
            r'href=["\'](https://www\.tiktok\.com/[^"\']+)["\']',
            r'<link[^>]+rel=["\']canonical["\'][^>]+href=["\']([^"\']+)',
            r'<meta[^>]+property=["\']og:url["\'][^>]+content=["\']([^"\']+)',
        )
        for pattern in patterns:
            match = re.search(pattern, page, re.IGNORECASE)
            if match:
                target = html_lib.unescape(match.group(1))
                target = target.replace("\\u0026", "&").replace("\\/", "/")
                expanded = urljoin(final_url, target)
                logger(f"   + URL sau khi mở rộng: {expanded[:240]}")
                return expanded
        logger(f"   ! TikTok chưa chuyển hướng link (HTTP {response.status_code}).")
        return final_url
    except Exception as exc:
        logger(f"   ! Không mở rộng được link ngắn ({exc}); sẽ thử với link gốc.")
        return url


def parse_tiktok_link_info(url):
    """Read the target comment CID and video ID from a TikTok URL."""
    url = html_lib.unescape(unquote(str(url or "").strip()))
    parsed = urlparse(url)
    query = parse_qs(parsed.query, keep_blank_values=True)
    query_lower = {key.lower(): values for key, values in query.items()}

    target_keys = ("share_comment_id", "cid", "commentid", "comment_id")
    target_id = None
    for key in target_keys:
        values = query_lower.get(key)
        if values and values[0].strip():
            target_id = values[0].strip()
            break
    if not target_id:
        match = re.search(
            r"(?:share_comment_id|cid|commentId|comment_id)=([^&?#\s]+)",
            url,
            re.IGNORECASE,
        )
        if match:
            target_id = unquote(match.group(1)).strip()

    video_match = re.search(r"/(?:@[^/]+/)?(?:video|photo)/(\d+)", parsed.path)
    video_id = video_match.group(1) if video_match else None
    if not video_id:
        for key in ("item_id", "aweme_id", "share_item_id"):
            values = query_lower.get(key)
            if values and values[0].isdigit():
                video_id = values[0]
                break
    return target_id, video_id


def build_cid_page_url(video_url, target_id):
    """Open the video page with TikTok's Base64 CID query parameter."""
    cid = str(target_id or "").strip()
    if cid.isdigit():
        cid = base64.b64encode(cid.encode("utf-8")).decode("ascii").rstrip("=")
    parsed = urlparse(video_url)
    query = parse_qs(parsed.query, keep_blank_values=True)
    query["cid"] = [cid]
    return urlunparse((
        parsed.scheme, parsed.netloc, parsed.path, parsed.params,
        urlencode(query, doseq=True), parsed.fragment,
    ))


def _comment_ids(comment):
    for key in ("cid", "comment_id", "commentId", "id", "id_str"):
        value = comment.get(key)
        if value is not None and str(value).strip():
            yield str(value).strip()


def _comment_text(comment):
    for key in ("text", "content", "comment_text"):
        value = comment.get(key)
        if isinstance(value, str):
            return value.strip()
    return ""


def _comment_author(comment):
    user = comment.get("user") or comment.get("author") or comment.get("user_info") or {}
    if not isinstance(user, dict):
        user = {}
    return (
        user.get("nickname") or user.get("display_name") or user.get("unique_id")
        or user.get("uniqueId") or comment.get("nickname") or "Tài_Khoản_Ẩn"
    )


def _format_comment_match(comment, targets, target_id):
    """Build a result only when this comment object's own CID matches."""
    ids = set()
    for raw_id in _comment_ids(comment):
        ids.update(comment_id_variants(raw_id))
    if not ids.intersection(targets):
        return None
    return {
        "id": next((i for i in _comment_ids(comment)
                    if comment_id_variants(i).intersection(targets)), str(target_id)),
        "author": str(_comment_author(comment)),
        "text": _comment_text(comment),
    }


def find_exact_comment(payload, target_id):
    """Match exact IDs only in the API's top-level comment list."""
    targets = comment_id_variants(target_id)
    if not targets:
        return None
    comments = _comment_list(payload)
    for comment in comments:
        if isinstance(comment, dict):
            match = _format_comment_match(comment, targets, target_id)
            if match:
                return match
    return None


def _top_level_comment_lists(node):
    """Find comment arrays in page hydration without descending into replies."""
    if isinstance(node, list):
        for item in node:
            yield from _top_level_comment_lists(item)
        return
    if not isinstance(node, dict):
        return
    for key, value in node.items():
        key_lower = str(key).lower()
        if "reply" in key_lower or key_lower in {"child_comments", "children"}:
            continue
        if key_lower in {"comments", "comment_list", "commentlist"} and isinstance(value, list):
            yield value
            continue
        if isinstance(value, (dict, list)):
            yield from _top_level_comment_lists(value)


def find_exact_hydrated_comment(payload, target_id):
    """Exact CID match against top-level comment arrays in page hydration."""
    targets = comment_id_variants(target_id)
    if not targets:
        return None
    for comments in _top_level_comment_lists(payload):
        for comment in comments:
            if not isinstance(comment, dict):
                continue
            match = _format_comment_match(comment, targets, target_id)
            if match:
                return match
    return None


def _payload_data(payload):
    data = payload.get("data", {}) if isinstance(payload, dict) else {}
    return data if isinstance(data, dict) else {}


def _comment_list(payload):
    data = _payload_data(payload)
    comments = data.get("comments")
    if comments is None and isinstance(payload, dict):
        comments = payload.get("comments")
    return comments if isinstance(comments, list) else []


def _has_more(payload):
    data = _payload_data(payload)
    value = data.get("has_more", data.get("hasMore", payload.get("has_more", False)))
    return str(value).lower() not in ("", "0", "false", "none")


def _next_cursor(payload, current):
    data = _payload_data(payload)
    value = data.get("cursor", payload.get("cursor", current))
    return str(value if value is not None else current)


def request_json(session, endpoint, params, logger, label):
    """Bounded retries: retry transient failures, never reuse a previous response."""
    last_error = "không có phản hồi"
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            _wait_for_tikwm_slot()
            response = session.get(
                endpoint,
                params=params,
                headers={"User-Agent": USER_AGENT,
                         "Accept": "application/json, text/plain, */*"},
                timeout=REQUEST_TIMEOUT,
                impersonate="chrome120",
            )
            try:
                payload = response_json(response)
            except Exception:
                payload = None
            if response.status_code == 200 and isinstance(payload, dict):
                code = payload.get("code")
                msg = str(payload.get("msg", payload.get("message", "")))
                if code in (None, 0, "0"):
                    return payload, None
                last_error = f"API code={code}: {msg}".strip()
                retryable = any(word in msg.lower() for word in ("limit", "rate", "busy", "try again"))
            else:
                last_error = f"HTTP {response.status_code}" if response.status_code != 200 else "phản hồi không phải JSON"
                retryable = response.status_code == 429 or response.status_code >= 500
        except Exception as exc:
            last_error = str(exc)
            retryable = True

        if attempt < MAX_RETRIES and retryable:
            pause = min(2 ** (attempt - 1), 6)
            logger(f"   ! {label}: {last_error}; thử lại sau {pause} giây ({attempt}/{MAX_RETRIES - 1}).")
            time.sleep(pause)
        else:
            break
    return None, last_error


def fetch_tikwm_comment(session, video_url, video_id, target_id, logger):
    """Search only top-level comments; retry with bare video ID if needed."""
    endpoint = "https://www.tikwm.com/api/comment/list"
    sources = [("video link", video_url)]
    if video_id and str(video_id) != str(video_url):
        sources.append(("Video ID", str(video_id)))

    cache = _get_cid_hint(video_id, target_id, "tikwm")
    preferred_source = cache.get("top_source")
    if preferred_source:
        sources.sort(key=lambda item: item[0] != preferred_source)

    total_pages = 0
    last_error = None
    for source_name, source_value in sources:
        cached_cursor = (
            str(cache.get("top_cursor", "0"))
            if source_name == preferred_source else "0"
        )
        if cached_cursor not in ("", "0"):
            logger(f"   + TikWM: thử trang đã lưu ({source_name}, cursor {cached_cursor})...")
            cached_page, cached_error = request_json(
                session, endpoint,
                {"url": source_value, "count": 50, "cursor": cached_cursor},
                logger, "TikWM comment/list cache",
            )
            total_pages += 1
            if cached_error:
                last_error = cached_error
                if any(marker in last_error.lower() for marker in ("limit", "429", "rate")):
                    return None, f"TikWM đang giới hạn request: {last_error}"
            elif cached_page:
                found = find_exact_comment(cached_page, target_id)
                if found:
                    _save_cid_hint(
                        video_id, target_id, "tikwm",
                        top_source=source_name, top_cursor=cached_cursor,
                    )
                    return found, None

        cursor = "0"
        seen_cursors = set()
        for page in range(MAX_COMMENT_PAGES):
            logger(
                f"   + TikWM ({source_name}): trang {page + 1}/"
                f"{MAX_COMMENT_PAGES}, cursor {cursor}..."
            )
            payload, error = request_json(
                session, endpoint,
                {"url": source_value, "count": 50, "cursor": cursor},
                logger, "TikWM comment/list",
            )
            total_pages += 1
            if error:
                last_error = error
                if any(marker in last_error.lower() for marker in ("limit", "429", "rate")):
                    return None, f"TikWM đang giới hạn request: {last_error}"
                break

            found = find_exact_comment(payload, target_id)
            if found:
                _save_cid_hint(
                    video_id, target_id, "tikwm",
                    top_source=source_name, top_cursor=cursor,
                )
                return found, None

            comments = _comment_list(payload)
            if not comments:
                break
            next_cursor = _next_cursor(payload, cursor)
            if (
                not _has_more(payload)
                or next_cursor in ("", "0", cursor)
                or next_cursor in seen_cursors
            ):
                break
            seen_cursors.add(cursor)
            cursor = next_cursor

        if source_name != sources[-1][0]:
            logger("   + Chưa thấy CID theo link; thử lại danh sách bằng Video ID...")

    detail = f"; API báo: {last_error}" if last_error else ""
    return None, (
        f"Không thấy CID {target_id} trong danh sách comment chính "
        f"sau {total_pages} lượt trang{detail}"
    )


def fetch_tiktok_web_comment(session, video_id, target_id, logger, video_url=None):
    """Search TikTok's top-level comment feed only."""
    if not video_id:
        return None, "Link thiếu Video ID để gọi TikTok web API"

    endpoint = "https://www.tiktok.com/api/comment/list/"
    headers = {
        "User-Agent": USER_AGENT,
        "Referer": video_url or f"https://www.tiktok.com/video/{video_id}",
        "Origin": "https://www.tiktok.com",
        "Accept-Language": "vi-VN,vi;q=0.9,en-US;q=0.8",
        "Accept": "application/json, text/plain, */*",
    }
    cursor = "0"
    seen_cursors = set()
    last_error = None

    for page in range(MAX_COMMENT_PAGES):
        logger(
            f"   + TikTok: trang comment chính {page + 1}/"
            f"{MAX_COMMENT_PAGES}, cursor {cursor}..."
        )
        params = {
            "aid": "1988", "aweme_id": str(video_id),
            "count": "50", "cursor": cursor,
            "device_platform": "web_pc", "app_name": "tiktok_web",
            "channel": "tiktok_web", "browser_language": "vi-VN",
            "browser_name": "Chrome", "browser_platform": "Win32",
            "browser_version": "120.0.0.0", "cookie_enabled": "true",
        }
        try:
            response = session.get(
                endpoint, params=params, headers=headers,
                timeout=REQUEST_TIMEOUT, impersonate="chrome120",
            )
            if response.status_code != 200:
                last_error = f"HTTP {response.status_code}"
                break
            payload = response_json(response)
            if not isinstance(payload, dict):
                last_error = "JSON không hợp lệ"
                break
            status = payload.get("status_code", 0)
            if status not in (0, "0", None):
                last_error = f"status_code={status}: {payload.get('status_msg', '')}"
                break
        except Exception as exc:
            last_error = str(exc)
            break

        found = find_exact_comment(payload, target_id)
        if found:
            _save_cid_hint(video_id, target_id, "tiktok", top_cursor=cursor)
            return found, None

        if not _comment_list(payload):
            break
        next_cursor = _next_cursor(payload, cursor)
        if (
            not _has_more(payload)
            or next_cursor in ("", "0", cursor)
            or next_cursor in seen_cursors
        ):
            break
        seen_cursors.add(cursor)
        cursor = next_cursor

    detail = f"; API báo: {last_error}" if last_error else ""
    return None, f"TikTok web API không thấy CID trong comment chính{detail}"


def fetch_cid_page_comment(session, video_url, target_id, logger):
    """Read top-level comment data from the page opened with the target CID."""
    try:
        response = session.get(
            video_url,
            headers={"User-Agent": USER_AGENT,
                     "Accept-Language": "vi-VN,vi;q=0.9,en-US;q=0.8"},
            timeout=REQUEST_TIMEOUT,
            allow_redirects=True,
            impersonate="chrome120",
        )
        page = response.text or ""
        if "captcha_verify_container" in page or "verify-ele" in page:
            return None, "TikTok yêu cầu xác minh Captcha"
        if response.status_code != 200:
            return None, f"Trang video trả HTTP {response.status_code}"
        script = re.search(
            r'<script[^>]+id=["\']__UNIVERSAL_DATA_FOR_REHYDRATION__["\'][^>]*>(.*?)</script>',
            page,
            re.DOTALL | re.IGNORECASE,
        )
        if not script:
            return None, "Trang video không có dữ liệu bình luận nhúng"
        payload = json.loads(html_lib.unescape(script.group(1)))
        return find_exact_hydrated_comment(payload, target_id), None
    except Exception as exc:
        return None, str(exc)


def fetch_comment_data(session, input_url, logger):
    expanded = expand_short_url(session, input_url, logger)
    target_id, video_id = parse_tiktok_link_info(expanded)
    original_target, original_video = parse_tiktok_link_info(input_url)
    target_id = target_id or original_target
    video_id = video_id or original_video
    if not target_id:
        return None, None, None, (
            "Link sau khi mở rộng không có CID "
            f"(share_comment_id/comment_id/cid). URL nhận được: {expanded[:240]}"
        )

    numeric_cid = _decode_numeric_base64(target_id)
    if numeric_cid:
        logger(f"   + CID gốc từ link (Base64): {target_id}")
        logger(f"   + CID số giải mã để đối chiếu API: {numeric_cid}")
    else:
        logger(f"   + CID gốc từ link: {target_id}")
    if video_id:
        logger(f"   + Video ID: {video_id}")
    else:
        logger("   ! Không trích xuất được Video ID; TikTok web API sẽ bị bỏ qua.")

    cid_url = build_cid_page_url(expanded, target_id)
    logger("   + Mở trang video kèm CID để lấy bình luận được trỏ tới...")
    page_match, page_error = fetch_cid_page_comment(session, cid_url, target_id, logger)
    if page_match:
        return page_match["author"], page_match["text"], page_match["id"], "Success (CID page)"

    direct_match, direct_error = fetch_tiktok_web_comment(
        session, video_id, target_id, logger, expanded
    )
    if direct_match:
        return direct_match["author"], direct_match["text"], direct_match["id"], "Success (TikTok)"

    match, tikwm_error = fetch_tikwm_comment(
        session, expanded, video_id, target_id, logger
    )
    if match:
        return match["author"], match["text"], match["id"], "Success (TikWM)"

    return None, None, None, (
        f"Không xác minh được bình luận {target_id}. {tikwm_error}; "
        f"{direct_error}; trang CID: {page_error}"
    )


def setup_excel_file(filename, second_header="Tài Khoản + Nội Dung"):
    headers = ["Link Comment Gốc", second_header]
    if not os.path.exists(filename):
        workbook = openpyxl.Workbook()
        try:
            sheet = workbook.active
            sheet.title = "TikTok Comments"
            for col, title in enumerate(headers, 1):
                sheet.cell(row=1, column=col, value=title)
            _format_output_sheet(sheet)
            atomic_save_workbook(workbook, filename)
        finally:
            workbook.close()
        return
    with EXCEL_LOCK:
        workbook = openpyxl.load_workbook(filename)
        try:
            sheet = workbook.active
            sheet.title = "TikTok Comments"
            result_col = next(
                (
                    col for col in range(2, sheet.max_column + 1)
                    if any(token in str(sheet.cell(1, col).value or "").casefold()
                           for token in ("tài khoản", "nội dung", "lỗi"))
                ),
                None,
            )
            if result_col and result_col != 2:
                for row_number in range(2, sheet.max_row + 1):
                    sheet.cell(row_number, 2).value = sheet.cell(row_number, result_col).value
            if sheet.max_column > 2:
                sheet.delete_cols(3, sheet.max_column - 2)
            for col, title in enumerate(headers, 1):
                sheet.cell(row=1, column=col, value=title)
            _format_output_sheet(sheet)
            atomic_save_workbook(workbook, filename)
        finally:
            workbook.close()


def _format_output_sheet(sheet):
    header_fill = PatternFill(fill_type="solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    header_alignment = Alignment(horizontal="center", vertical="center")
    for cell in sheet[1][:2]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment
    sheet.row_dimensions[1].height = 22
    sheet.column_dimensions["A"].width = 52
    sheet.column_dimensions["B"].width = 88


def atomic_save_workbook(workbook, filename):
    filename = os.path.abspath(filename)
    directory = os.path.dirname(filename)
    os.makedirs(directory, exist_ok=True)
    handle, temporary = tempfile.mkstemp(
        prefix=".tiktok_result_", suffix=".tmp.xlsx", dir=directory
    )
    os.close(handle)
    try:
        workbook.save(temporary)
        os.replace(temporary, filename)
    except Exception:
        try:
            os.remove(temporary)
        except OSError:
            pass
        raise


def save_result(filename, url, result_text, second_header="Tài Khoản + Nội Dung"):
    with EXCEL_LOCK:
        last_error = None
        for attempt in range(3):
            workbook = None
            try:
                if not os.path.exists(filename):
                    setup_excel_file(filename, second_header=second_header)
                workbook = openpyxl.load_workbook(filename)
                sheet = workbook.active
                sheet.append([url, result_text])
                atomic_save_workbook(workbook, filename)
                return True, None
            except Exception as exc:
                last_error = str(exc)
                time.sleep(0.5 * (attempt + 1))
            finally:
                if workbook:
                    workbook.close()
        return False, last_error


def normalize_link_key(url):
    raw = html_lib.unescape(str(url or "").strip())
    parsed = urlparse(raw)
    host = (parsed.hostname or "").lower()
    if host.startswith("www."):
        host = host[4:]
    if host in {"vt.tiktok.com", "vm.tiktok.com"}:
        host = "short.tiktok.com"
    path = parsed.path.rstrip("/")
    query = {key.lower(): values for key, values in
             parse_qs(parsed.query, keep_blank_values=True).items()}
    cid = None
    for key in ("cid", "share_comment_id", "comment_id", "commentid"):
        values = query.get(key)
        if values and values[0].strip():
            value = values[0].strip()
            cid = _decode_numeric_base64(value) or value
            break
    if cid:
        return f"cid:{cid}"
    if host:
        return f"{host}{path}"
    return re.sub(r"\s+", "", raw).rstrip("/").lower()


def load_processed_link_keys(filenames):
    processed = set()
    for filename in filenames:
        if not os.path.exists(filename):
            continue
        workbook = None
        try:
            workbook = openpyxl.load_workbook(filename, read_only=True, data_only=True)
            sheet = workbook.active
            for (link,) in sheet.iter_rows(min_row=2, min_col=1, max_col=1, values_only=True):
                if link:
                    processed.add(normalize_link_key(link))
        finally:
            if workbook:
                workbook.close()
    return processed


def filter_unprocessed_links(links, processed_keys):
    unique_links = []
    processed = set(processed_keys)
    seen = set()
    duplicate_count = 0
    processed_count = 0
    for link in links:
        key = normalize_link_key(link)
        if key in processed:
            processed_count += 1
            continue
        if key in seen:
            duplicate_count += 1
            continue
        seen.add(key)
        unique_links.append(link)
    return unique_links, duplicate_count, processed_count


def extract_tiktok_links(text):
    pattern = re.compile(
        r"(?<![\w./-])(?:https?://)?(?:[\w-]+\.)*tiktok\.com/[^\s<>\"']+",
        re.IGNORECASE,
    )
    links = []
    for match in pattern.finditer(str(text or "")):
        candidate = match.group(0).rstrip(".,;!?)]}»")
        if not candidate.lower().startswith(("http://", "https://")):
            candidate = "https://" + candidate
        host = (urlparse(candidate).hostname or "").lower()
        if host == "tiktok.com" or host.endswith(".tiktok.com"):
            links.append(candidate)
    return links


def move_legacy_errors_to_file(main_filename, error_filename):
    if not os.path.exists(main_filename):
        return 0
    with EXCEL_LOCK:
        main_workbook = None
        error_workbook = None
        try:
            main_workbook = openpyxl.load_workbook(main_filename)
            error_workbook = openpyxl.load_workbook(error_filename)
            main_sheet = main_workbook.active
            error_sheet = error_workbook.active
            existing_errors = {
                normalize_link_key(row[0])
                for row in error_sheet.iter_rows(min_row=2, min_col=1, max_col=1, values_only=True)
                if row[0]
            }
            remove_rows = []
            moved = 0
            for row_number in range(2, main_sheet.max_row + 1):
                link = main_sheet.cell(row_number, 1).value
                result = main_sheet.cell(row_number, 2).value
                if isinstance(result, str) and result.strip().upper().startswith(("LỖI:", "LỖI HỆ THỐNG:")):
                    if not link:
                        continue
                    key = normalize_link_key(link)
                    if key not in existing_errors:
                        error_sheet.append([link, result])
                        existing_errors.add(key)
                    remove_rows.append(row_number)
                    moved += 1
            for row_number in reversed(remove_rows):
                main_sheet.delete_rows(row_number, 1)
            if moved:
                atomic_save_workbook(error_workbook, error_filename)
                atomic_save_workbook(main_workbook, main_filename)
            return moved
        finally:
            if main_workbook:
                main_workbook.close()
            if error_workbook:
                error_workbook.close()


def read_excel_bytes(filename):
    if not os.path.exists(filename):
        return None
    with EXCEL_LOCK:
        with open(filename, "rb") as f:
            return f.read()


# ================= GIAO DIỆN MOBILE WEB (STREAMLIT) =================
st.title("⚡ TikTok Comment Extractor")
st.caption("Tìm chính xác bình luận theo CID — Tối ưu cho trình duyệt điện thoại")

main_path = os.path.abspath(EXCEL_FILE)
error_path = os.path.abspath(ERROR_EXCEL_FILE)
setup_excel_file(main_path)
setup_excel_file(error_path, second_header="Lỗi")

if "logs" not in st.session_state:
    st.session_state.logs = []

with st.expander("📤 Nạp file Excel cũ (tùy chọn để nối tiếp & tránh trùng)"):
    up_main = st.file_uploader("File Excel chính cũ (.xlsx)", type=["xlsx"], key="up_main")
    if up_main is not None and st.button("Khôi phục Excel chính"):
        with EXCEL_LOCK:
            with open(main_path, "wb") as f:
                f.write(up_main.getvalue())
        setup_excel_file(main_path)
        st.success("Đã nạp lại file Excel chính!")

    up_err = st.file_uploader("File Excel lỗi cũ (.xlsx)", type=["xlsx"], key="up_err")
    if up_err is not None and st.button("Khôi phục Excel lỗi"):
        with EXCEL_LOCK:
            with open(error_path, "wb") as f:
                f.write(up_err.getvalue())
        setup_excel_file(error_path, second_header="Lỗi")
        st.success("Đã nạp lại file Excel lỗi!")

raw_input = st.text_area(
    "Danh sách link comment TikTok:",
    height=150,
    placeholder="Dán các link comment TikTok vào đây (mỗi dòng 1 link hoặc dán cả đoạn văn bản)...",
)

col1, col2, col3 = st.columns([2, 1, 1])
start_clicked = col1.button("⚡ BẮT ĐẦU TÌM", type="primary", use_container_width=True)
stop_clicked = col2.button("🛑 Dừng", use_container_width=True)
clear_clicked = col3.button("🗑 Xóa Excel", use_container_width=True)

if stop_clicked:
    st.warning("Đã dừng tiến trình. Các link đã chạy xong trước đó đều đã được lưu vào Excel.")

if clear_clicked:
    for fpath in (main_path, error_path):
        if os.path.exists(fpath):
            os.remove(fpath)
    setup_excel_file(main_path)
    setup_excel_file(error_path, second_header="Lỗi")
    st.session_state.logs = []
    st.success("Đã xóa dữ liệu Excel cũ và tạo file mới trắng!")

log_placeholder = st.empty()


def append_log(message):
    line = f"[{datetime.now().strftime('%H:%M:%S')}] {message}"
    st.session_state.logs.append(line)
    log_placeholder.code("\n".join(st.session_state.logs[-45:]), language="text")


if start_clicked:
    links = extract_tiktok_links(raw_input)
    if not links:
        st.error("Không tìm thấy link TikTok hợp lệ!")
    else:
        st.session_state.logs = []
        moved_errors = move_legacy_errors_to_file(main_path, error_path)
        if moved_errors:
            append_log(f"   + Đã chuyển {moved_errors} dòng lỗi cũ sang Excel lỗi.")

        known_links = load_processed_link_keys((main_path, error_path))
        links, duplicate_count, processed_count = filter_unprocessed_links(links, known_links)
        append_log(
            f"   + Bỏ qua {duplicate_count} link trùng và {processed_count} link đã có; "
            f"còn {len(links)} link mới."
        )

        if not links:
            append_log("🏁 Không có link mới cần xử lý.")
        else:
            session = create_tiktok_session(append_log)
            outcomes = {"success": 0, "error": 0}
            progress_bar = st.progress(0)
            try:
                append_log(f"🚀 Bắt đầu xử lý {len(links)} link...")
                for idx, url in enumerate(links, 1):
                    append_log("--------------------------------------------------")
                    append_log(f"🔍 [{idx}/{len(links)}] Đang xử lý: {url}")
                    try:
                        author, text, matched_id, status = fetch_comment_data(
                            session, url, append_log
                        )
                        if author is not None and text is not None:
                            clean_text = text.replace("\r", " ").replace("\n", " ").strip()
                            result = f"{author}: {clean_text}"
                            saved, err = save_result(main_path, url, result)
                            if saved:
                                append_log(f"✅ [KHỚP CID {matched_id}] {result}")
                                append_log(f"   + Nguồn xác minh: {status}")
                                outcomes["success"] += 1
                            else:
                                append_log(f"❌ Không lưu được Excel: {err}")
                                save_result(error_path, url, f"LỖI LƯU KẾT QUẢ: {err}", second_header="Lỗi")
                                outcomes["error"] += 1
                        else:
                            result = f"LỖI: {status}"
                            save_result(error_path, url, result, second_header="Lỗi")
                            append_log(f"❌ {result}")
                            outcomes["error"] += 1
                    except Exception as exc:
                        result = f"LỖI HỆ THỐNG: {exc}"
                        save_result(error_path, url, result, second_header="Lỗi")
                        append_log(f"❌ {result}")
                        outcomes["error"] += 1

                    progress_bar.progress(idx / len(links))

                append_log(
                    f"🏁 Hoàn tất: {outcomes['success']} thành công, {outcomes['error']} lỗi."
                )
            finally:
                session.close()

elif st.session_state.logs:
    log_placeholder.code("\n".join(st.session_state.logs[-45:]), language="text")

st.divider()
st.subheader("📥 Tải kết quả về điện thoại")
dl_col1, dl_col2 = st.columns(2)

main_bytes = read_excel_bytes(main_path)
if main_bytes:
    dl_col1.download_button(
        label="📂 Tải Excel chính (.xlsx)",
        data=main_bytes,
        file_name=EXCEL_FILE,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

error_bytes = read_excel_bytes(error_path)
if error_bytes:
    dl_col2.download_button(
        label="⚠️ Tải Excel lỗi (.xlsx)",
        data=error_bytes,
        file_name=ERROR_EXCEL_FILE,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )