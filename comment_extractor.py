from __future__ import annotations

"""
comment_extractor.py

Local Campaign Automation용 통합 댓글 URL 추출 모듈.

지원 Channel
------------
YT : YouTube      -> yt-dlp
IG : Instagram    -> yt-dlp
X  : X / Twitter  -> Playwright persistent Microsoft Edge (visible)
FB : Facebook     -> Playwright persistent Chromium
TT : TikTok       -> Playwright + Edge CDP/network response

Public API
----------
extract_comment_url(channel, post_url) -> str | None

build_url_cell_value(channel, post_url) -> str | None

정확성 정책
-----------
- 게시물 작성자 본인의 댓글/self-reply는 제외한다.
- 댓글 작성자를 확인할 수 없으면 잘못 선택하지 않고 skip한다.
- 댓글 추출 실패는 None을 반환한다.
- raw_to_processed.py에서는 실패 시 원 게시물 URL만 유지한다.
"""

import atexit
import base64
import html
import logging
import re
import subprocess
import threading
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable
from urllib.parse import (
    parse_qs,
    quote,
    urlencode,
    urljoin,
    urlparse,
    urlsplit,
    urlunsplit,
)
from urllib.request import urlopen


# =============================================================================
# 공통 설정
# =============================================================================

PROJECT_ROOT = Path(__file__).resolve().parent

LOGGER = logging.getLogger(__name__)

SUPPORTED_CHANNELS = {
    "YT",
    "IG",
    "X",
    "FB",
    "TT",
}

CHANNEL_ALIASES = {
    "YT": "YT",
    "YOUTUBE": "YT",
    "IG": "IG",
    "INSTAGRAM": "IG",
    "X": "X",
    "TWITTER": "X",
    "FB": "FB",
    "FACEBOOK": "FB",
    "TT": "TT",
    "TIKTOK": "TT",
    "TIK TOK": "TT",
}

PAGE_TIMEOUT_MS = 45_000

# 첫 로그인 때만 사람이 인증할 수 있게 유지한다.
INTERACTIVE_LOGIN = True

# Facebook 등 기존 browser automation의 기본 headless 설정.
BROWSER_HEADLESS = True

# X / Twitter는 headless Edge에서 X가 HTTP response failure를 반환하는
# 환경이 확인되어 Microsoft Edge visible mode로 고정한다.
X_BROWSER_CHANNEL = "msedge"
X_HEADLESS = False

# 댓글별 상세 로그(self-comment 제외, 개별 실패 URL 등)는 출력하지 않는다.
# 전체 작업 종료 시 CommentExtractorSession.close()에서 요약만 출력한다.
DEBUG = False
COMMENT_SUMMARY_ENABLED = True


# =============================================================================
# 공통 helper
# =============================================================================

def normalize_channel(
    value: object,
) -> str | None:
    if value is None:
        return None

    text = str(value).strip().upper()

    if not text:
        return None

    return CHANNEL_ALIASES.get(text)


def normalize_url(
    value: object,
) -> str | None:
    if value is None:
        return None

    text = str(value).strip()

    if not text:
        return None

    try:
        parsed = urlparse(text)
    except Exception:
        return None

    if parsed.scheme.lower() not in {
        "http",
        "https",
    }:
        return None

    if not parsed.netloc:
        return None

    return text


def normalize_identity(
    value: object,
) -> str | None:
    if value is None:
        return None

    text = (
        str(value)
        .strip()
        .lstrip("@")
        .casefold()
    )

    return text or None


def identity_aliases(
    *values: object,
) -> set[str]:
    result: set[str] = set()

    for value in values:
        normalized = normalize_identity(
            value
        )

        if normalized:
            result.add(normalized)

    return result


def is_self_comment(
    post_author_aliases: set[str],
    comment_author_aliases: set[str],
) -> bool:
    if (
        not post_author_aliases
        or not comment_author_aliases
    ):
        return False

    return bool(
        post_author_aliases
        & comment_author_aliases
    )


# =============================================================================
# yt-dlp 공통: YouTube / Instagram
# =============================================================================

def _collect_comments_from_info(
    info: dict[str, Any],
) -> list[dict[str, Any]]:
    """
    일반 단일 게시물과 Instagram carousel 형태 모두 처리한다.
    """

    comments: list[dict[str, Any]] = []

    raw_comments = info.get("comments")

    if isinstance(raw_comments, list):
        comments.extend(
            comment
            for comment in raw_comments
            if isinstance(comment, dict)
        )

    entries = info.get("entries")

    if isinstance(entries, list):
        for entry in entries:
            if not isinstance(entry, dict):
                continue

            entry_comments = entry.get(
                "comments"
            )

            if isinstance(
                entry_comments,
                list,
            ):
                comments.extend(
                    comment
                    for comment
                    in entry_comments
                    if isinstance(
                        comment,
                        dict,
                    )
                )

    # 동일 comment ID 중복 제거
    deduped: list[dict[str, Any]] = []
    seen_ids: set[str] = set()

    for comment in comments:
        comment_id = comment.get("id")

        if comment_id is None:
            continue

        comment_id_text = str(
            comment_id
        ).strip()

        if not comment_id_text:
            continue

        if comment_id_text in seen_ids:
            continue

        seen_ids.add(
            comment_id_text
        )

        deduped.append(
            comment
        )

    return deduped


def _collect_post_author_aliases_from_info(
    info: dict[str, Any],
) -> set[str]:
    aliases = identity_aliases(
        info.get("uploader_id"),
        info.get("uploader"),
        info.get("channel_id"),
        info.get("channel"),
        info.get("creator"),
    )

    entries = info.get("entries")

    if isinstance(entries, list):
        for entry in entries:
            if not isinstance(entry, dict):
                continue

            aliases.update(
                identity_aliases(
                    entry.get("uploader_id"),
                    entry.get("uploader"),
                    entry.get("channel_id"),
                    entry.get("channel"),
                    entry.get("creator"),
                )
            )

    return aliases


def _extract_with_ytdlp(
    post_url: str,
    platform: str,
) -> tuple[
    dict[str, Any],
    list[dict[str, Any]],
    set[str],
]:
    """
    yt-dlp Python API로 게시물 metadata + comments를 가져온다.
    """

    try:
        import yt_dlp
    except ImportError as exc:
        raise RuntimeError(
            "yt-dlp가 설치되어 있지 않습니다. "
            "현재 프로젝트 가상환경에 yt-dlp를 설치하세요."
        ) from exc

    options: dict[str, Any] = {
        "skip_download": True,
        "getcomments": True,
        "quiet": True,
        "no_warnings": True,
        "ignoreerrors": False,
    }

    # Instagram은 사진-only 게시물이나 carousel 내부 이미지처럼
    # video format이 없는 media도 정상적인 게시물로 취급해야 한다.
    # 댓글 URL 추출 목적에서는 video format 자체가 필요하지 않으므로,
    # "There is no video in this post" / "No video formats found!"
    # 오류 때문에 metadata/comment extraction 전체가 중단되지 않도록 한다.
    #
    # 이 옵션은 IG에만 적용하며 YouTube의 기존 동작은 변경하지 않는다.
    if platform == "IG":
        options["ignore_no_formats_error"] = True

    # YouTube는 댓글이 매우 많은 경우가 있으므로
    # 첫 50개 top-level 댓글만 조회한다.
    if platform == "YT":
        options["extractor_args"] = {
            "youtube": {
                "comment_sort": [
                    "top"
                ],
                "max_comments": [
                    "50,50,0,0,1"
                ],
            }
        }

    with yt_dlp.YoutubeDL(
        options
    ) as ydl:
        info = ydl.extract_info(
            post_url,
            download=False,
        )

    if not isinstance(info, dict):
        raise RuntimeError(
            f"{platform} yt-dlp metadata가 "
            "dict 형태가 아닙니다."
        )

    comments = (
        _collect_comments_from_info(
            info
        )
    )

    post_author_aliases = (
        _collect_post_author_aliases_from_info(
            info
        )
    )

    return (
        info,
        comments,
        post_author_aliases,
    )


def _first_non_self_ytdlp_comment(
    comments: list[dict[str, Any]],
    post_author_aliases: set[str],
) -> dict[str, Any] | None:
    """
    작성자 ID/username을 확인할 수 있는 댓글만 후보로 사용한다.
    """

    for comment in comments:
        comment_id = comment.get("id")

        if comment_id is None:
            continue

        comment_author_aliases = (
            identity_aliases(
                comment.get("author_id"),
                comment.get("author"),
            )
        )

        # 작성자를 확인할 수 없으면 self-comment 검증도 불가능.
        if not comment_author_aliases:
            continue

        if is_self_comment(
            post_author_aliases,
            comment_author_aliases,
        ):
            if DEBUG:
                print(
                    "[SELF COMMENT 제외] "
                    f"id={comment_id}, "
                    f"author={comment.get('author')!r}"
                )

            continue

        return comment

    return None


# =============================================================================
# YouTube
# =============================================================================

def _youtube_video_id_from_url(
    post_url: str,
) -> str | None:
    parsed = urlparse(
        post_url
    )

    host = (
        parsed.netloc
        .lower()
        .split(":")[0]
    )

    if host in {
        "youtu.be",
        "www.youtu.be",
    }:
        value = (
            parsed.path
            .strip("/")
            .split("/")[0]
        )

        return value or None

    if host.endswith(
        "youtube.com"
    ):
        if parsed.path.rstrip("/") == "/watch":
            values = parse_qs(
                parsed.query
            ).get("v")

            if values and values[0]:
                return values[0]

        match = re.search(
            r"/(?:shorts|embed|live)/([^/?#]+)",
            parsed.path,
            flags=re.IGNORECASE,
        )

        if match:
            return match.group(1)

    return None


def extract_youtube_comment_url(
    post_url: str,
) -> str | None:
    info, comments, post_author_aliases = (
        _extract_with_ytdlp(
            post_url=post_url,
            platform="YT",
        )
    )

    if not post_author_aliases:
        LOGGER.debug(
            "[YT] 게시물 작성자를 식별하지 못해 "
            "댓글 선택을 중단합니다: %s",
            post_url,
        )
        return None

    comment = (
        _first_non_self_ytdlp_comment(
            comments=comments,
            post_author_aliases=(
                post_author_aliases
            ),
        )
    )

    if comment is None:
        return None

    comment_id = str(
        comment["id"]
    ).strip()

    video_id = (
        str(info.get("id")).strip()
        if info.get("id")
        else _youtube_video_id_from_url(
            post_url
        )
    )

    if not video_id:
        return None

    return (
        "https://www.youtube.com/watch?"
        + urlencode(
            {
                "v": video_id,
                "lc": comment_id,
            }
        )
    )


# =============================================================================
# Instagram
# =============================================================================

INSTAGRAM_POST_PATTERN = re.compile(
    r"/(?:p|reel|reels|tv)/([^/?#]+)",
    flags=re.IGNORECASE,
)


def _instagram_shortcode(
    post_url: str,
    info: dict[str, Any],
) -> str | None:
    info_id = info.get("id")

    if isinstance(info_id, str):
        info_id = info_id.strip()

        # 일반 post/reel의 yt-dlp id는 shortcode다.
        if info_id and not info_id.isdigit():
            return info_id

    match = INSTAGRAM_POST_PATTERN.search(
        urlparse(post_url).path
    )

    if not match:
        return None

    return match.group(1)


def extract_instagram_comment_url(
    post_url: str,
) -> str | None:
    info, comments, post_author_aliases = (
        _extract_with_ytdlp(
            post_url=post_url,
            platform="IG",
        )
    )

    if not post_author_aliases:
        LOGGER.debug(
            "[IG] 게시물 작성자를 식별하지 못해 "
            "댓글 선택을 중단합니다: %s",
            post_url,
        )
        return None

    comment = (
        _first_non_self_ytdlp_comment(
            comments=comments,
            post_author_aliases=(
                post_author_aliases
            ),
        )
    )

    if comment is None:
        return None

    comment_id = str(
        comment["id"]
    ).strip()

    shortcode = _instagram_shortcode(
        post_url=post_url,
        info=info,
    )

    if not shortcode:
        return None

    # Instagram comment permalink는 post shortcode 중심으로 사용한다.
    return (
        "https://www.instagram.com/"
        f"p/{quote(shortcode, safe='')}/"
        f"c/{quote(comment_id, safe='')}/"
    )



# =============================================================================
# Browser Session Manager
# =============================================================================

class CommentExtractorSession:
    """
    X / Facebook / TikTok의 Playwright browser context를
    여러 게시물 행에서 재사용한다.

    - X: .x_browser_profile persistent context 1개
    - FB: .fb_browser_profile persistent context 1개
    - TT: Edge CDP browser/context 연결 1개
    - 각 게시물에서는 page(tab)만 만들고 닫는다.
    - YT / IG는 yt-dlp 기반이므로 이 session을 사용하지 않는다.
    """

    def __init__(
        self,
        *,
        headless: bool = BROWSER_HEADLESS,
    ) -> None:
        self.headless = headless

        self._playwright = None
        self._x_context = None
        self._fb_context = None
        self._tt_browser = None
        self._tt_context = None

        # 댓글 추출 summary용 통계.
        self._stats_lock = threading.Lock()
        self._comment_attempts = 0
        self._comment_extracted = 0
        self._comment_not_extracted = 0
        self._comment_errors = 0
        self._summary_printed = False

        self._closed = False

    def record_comment_result(
        self,
        *,
        extracted: bool,
        error: bool = False,
    ) -> None:
        """댓글 추출 1건의 결과를 summary 통계에 기록한다."""

        with self._stats_lock:
            self._comment_attempts += 1

            if extracted:
                self._comment_extracted += 1
            else:
                self._comment_not_extracted += 1

            if error:
                self._comment_errors += 1

    def print_comment_summary(
        self,
    ) -> None:
        """세션 전체 댓글 추출 결과를 한 줄로 출력한다."""

        if (
            not COMMENT_SUMMARY_ENABLED
            or self._summary_printed
        ):
            return

        with self._stats_lock:
            attempts = self._comment_attempts
            extracted = self._comment_extracted
            not_extracted = self._comment_not_extracted
            errors = self._comment_errors
            self._summary_printed = True

        if attempts <= 0:
            return

        print(
            "[댓글 추출 완료] "
            f"총 {attempts}건 | "
            f"추출 {extracted}건 | "
            f"미추출 {not_extracted}건"
            + (
                f" (오류 {errors}건)"
                if errors
                else ""
            )
        )

    def __enter__(
        self,
    ) -> "CommentExtractorSession":
        self.start()
        return self

    def __exit__(
        self,
        exc_type,
        exc_value,
        traceback,
    ) -> None:
        self.close()

    def start(
        self,
    ) -> "CommentExtractorSession":
        if self._closed:
            raise RuntimeError(
                "이미 종료된 CommentExtractorSession은 "
                "다시 시작할 수 없습니다."
            )

        if self._playwright is not None:
            return self

        try:
            from playwright.sync_api import (
                sync_playwright,
            )
        except ImportError as exc:
            raise RuntimeError(
                "Playwright가 설치되어 있지 않습니다."
            ) from exc

        self._playwright = (
            sync_playwright().start()
        )

        return self

    def _require_playwright(
        self,
    ):
        self.start()

        assert self._playwright is not None
        return self._playwright

    def get_x_context(
        self,
    ):
        if self._x_context is not None:
            return self._x_context

        playwright = (
            self._require_playwright()
        )

        X_PROFILE_DIR.mkdir(
            parents=True,
            exist_ok=True,
        )

        self._x_context = (
            playwright.chromium
            .launch_persistent_context(
                user_data_dir=str(
                    X_PROFILE_DIR
                ),
                channel=X_BROWSER_CHANNEL,
                headless=X_HEADLESS,
                viewport={
                    "width": 1440,
                    "height": 1000,
                },
            )
        )

        return self._x_context

    def get_fb_context(
        self,
    ):
        if self._fb_context is not None:
            return self._fb_context

        playwright = (
            self._require_playwright()
        )

        FB_PROFILE_DIR.mkdir(
            parents=True,
            exist_ok=True,
        )

        self._fb_context = (
            playwright.chromium
            .launch_persistent_context(
                user_data_dir=str(
                    FB_PROFILE_DIR
                ),
                headless=self.headless,
                viewport={
                    "width": 1440,
                    "height": 1000,
                },
            )
        )

        return self._fb_context

    def get_tiktok_context(
        self,
    ):
        if self._tt_context is not None:
            return self._tt_context

        playwright = (
            self._require_playwright()
        )

        _tt_start_edge_if_needed()

        self._tt_browser = (
            playwright.chromium
            .connect_over_cdp(
                TIKTOK_CDP_ENDPOINT,
                timeout=10_000,
            )
        )

        if not self._tt_browser.contexts:
            self._tt_browser = None
            raise RuntimeError(
                "TikTok CDP Browser에는 연결했지만 "
                "BrowserContext가 없습니다."
            )

        self._tt_context = (
            self._tt_browser.contexts[0]
        )

        return self._tt_context

    def close(
        self,
    ) -> None:
        if self._closed:
            return

        # X / Facebook은 이 session이 직접 생성한
        # persistent context이므로 종료한다.
        for context in (
            self._x_context,
            self._fb_context,
        ):
            if context is None:
                continue

            try:
                context.close()
            except Exception:
                pass

        self._x_context = None
        self._fb_context = None

        # TikTok은 외부 Edge/CDP에 연결한 것이므로
        # browser.close()를 호출하지 않는다.
        self._tt_context = None
        self._tt_browser = None

        if self._playwright is not None:
            try:
                self._playwright.stop()
            except Exception:
                pass

        self._playwright = None

        # 행별 상세 로그 대신 세션 종료 시 전체 결과만 출력한다.
        self.print_comment_summary()

        self._closed = True


_DEFAULT_COMMENT_SESSION: CommentExtractorSession | None = None
_DEFAULT_COMMENT_SESSION_LOCK = threading.Lock()


def get_default_comment_extractor_session(
) -> CommentExtractorSession:
    """
    기존 호출부가 session을 명시하지 않아도
    프로세스 전체에서 browser context를 재사용한다.
    """

    global _DEFAULT_COMMENT_SESSION

    with _DEFAULT_COMMENT_SESSION_LOCK:
        if (
            _DEFAULT_COMMENT_SESSION is None
            or _DEFAULT_COMMENT_SESSION._closed
        ):
            _DEFAULT_COMMENT_SESSION = (
                CommentExtractorSession(
                    headless=BROWSER_HEADLESS,
                )
            )

        return _DEFAULT_COMMENT_SESSION


def close_default_comment_extractor_session(
) -> None:
    """
    module-level shared session을 명시적으로 종료한다.
    프로세스 종료 시에도 atexit으로 자동 호출된다.
    """

    global _DEFAULT_COMMENT_SESSION

    with _DEFAULT_COMMENT_SESSION_LOCK:
        session = _DEFAULT_COMMENT_SESSION
        _DEFAULT_COMMENT_SESSION = None

    if session is not None:
        session.close()


atexit.register(
    close_default_comment_extractor_session
)


# =============================================================================
# X / Twitter
# =============================================================================

X_PROFILE_DIR = (
    PROJECT_ROOT
    / ".x_browser_profile"
)

X_STATUS_PATH_PATTERN = re.compile(
    r"^/([^/?#]+)/status/(\d+)(?:[/\?#]|$)",
    flags=re.IGNORECASE,
)


def _x_original_identity(
    post_url: str,
) -> tuple[str, str]:
    parsed = urlparse(
        post_url
    )

    match = re.search(
        r"/([^/?#]+)/status/(\d+)",
        parsed.path,
        flags=re.IGNORECASE,
    )

    if not match:
        raise ValueError(
            "X/Twitter URL에서 username/status ID를 "
            f"찾을 수 없습니다: {post_url}"
        )

    username = normalize_identity(
        match.group(1)
    )

    if not username:
        raise ValueError(
            "X 원 게시물 작성자 username이 비어 있습니다."
        )

    return (
        username,
        match.group(2),
    )


def _x_parse_status_href(
    href: str | None,
) -> tuple[str, str, str] | None:
    if not href:
        return None

    if href.startswith(
        ("http://", "https://")
    ):
        path = urlparse(
            href
        ).path
    else:
        path = href

    match = X_STATUS_PATH_PATTERN.match(
        path
    )

    if not match:
        return None

    username = (
        normalize_identity(
            match.group(1)
        )
    )

    if not username:
        return None

    tweet_id = match.group(2)

    return (
        f"https://x.com/"
        f"{match.group(1)}/status/{tweet_id}",
        username,
        tweet_id,
    )


def _x_find_article_status(
    article,
) -> tuple[str, str, str] | None:
    selectors = (
        "a[href*='/status/']:has(time)",
        "a[href*='/status/']",
    )

    for selector in selectors:
        links = article.locator(
            selector
        )

        for index in range(
            links.count()
        ):
            try:
                href = (
                    links
                    .nth(index)
                    .get_attribute("href")
                )
            except Exception:
                continue

            result = (
                _x_parse_status_href(
                    href
                )
            )

            if result:
                return result

    return None


def _x_requires_login(
    page,
) -> bool:
    current_url = (
        page.url.lower()
    )

    return (
        "/i/flow/login" in current_url
        or "/login" in current_url
    )


def _x_handle_login(
    page,
    post_url: str,
    *,
    headless: bool = False,
) -> bool:
    if not _x_requires_login(
        page
    ):
        return True

    if headless:
        print()
        print(
            "[X] 로그인 세션이 필요하지만 현재 headless 모드입니다."
        )
        print(
            "X는 visible Edge로 실행하도록 설정되어 있습니다. "
            "X_HEADLESS 설정을 확인하세요."
        )
        return False

    if not INTERACTIVE_LOGIN:
        return False

    print()
    print(
        "[X] 최초 로그인이 필요합니다."
    )
    print(
        "열린 Microsoft Edge에서 로그인한 뒤 "
        "터미널로 돌아와 Enter를 누르세요."
    )

    input(
        "로그인 완료 후 Enter: "
    )

    page.goto(
        post_url,
        wait_until="domcontentloaded",
        timeout=PAGE_TIMEOUT_MS,
    )

    page.wait_for_timeout(
        3_000
    )

    return not _x_requires_login(
        page
    )


def _x_find_consumer_reply_url(
    page,
    original_username: str,
    original_tweet_id: str,
    seen_ids: set[str],
) -> str | None:
    articles = page.locator(
        "article"
    )

    for index in range(
        articles.count()
    ):
        article = articles.nth(
            index
        )

        result = _x_find_article_status(
            article
        )

        if not result:
            continue

        (
            tweet_url,
            username,
            tweet_id,
        ) = result

        if tweet_id in seen_ids:
            continue

        seen_ids.add(
            tweet_id
        )

        if tweet_id == original_tweet_id:
            continue

        # 기존 테스트본에서 추가한 핵심 보완:
        # 원 게시물 작성자의 self-reply도 제외.
        if username == original_username:
            if DEBUG:
                print(
                    "[X SELF REPLY 제외] "
                    f"{tweet_url}"
                )
            continue

        return tweet_url

    return None


def extract_twitter_comment_url(
    post_url: str,
    *,
    session: CommentExtractorSession | None = None,
) -> str | None:
    try:
        from playwright.sync_api import (
            Error as PlaywrightError,
            TimeoutError as PlaywrightTimeoutError,
        )
    except ImportError as exc:
        raise RuntimeError(
            "Playwright가 설치되어 있지 않습니다."
        ) from exc

    (
        original_username,
        original_tweet_id,
    ) = _x_original_identity(
        post_url
    )

    active_session = (
        session
        or get_default_comment_extractor_session()
    )

    context = (
        active_session.get_x_context()
    )

    page = context.new_page()

    try:
        page.set_default_timeout(
            10_000
        )

        page.goto(
            post_url,
            wait_until="domcontentloaded",
            timeout=PAGE_TIMEOUT_MS,
        )

        page.wait_for_timeout(
            3_000
        )

        if not _x_handle_login(
            page=page,
            post_url=post_url,
            headless=X_HEADLESS,
        ):
            return None

        seen_ids: set[str] = set()

        for _ in range(6):
            page.wait_for_timeout(
                1_500
            )

            result = (
                _x_find_consumer_reply_url(
                    page=page,
                    original_username=(
                        original_username
                    ),
                    original_tweet_id=(
                        original_tweet_id
                    ),
                    seen_ids=seen_ids,
                )
            )

            if result:
                return result

            page.mouse.wheel(
                0,
                1400,
            )

        return None

    except (
        PlaywrightTimeoutError,
        PlaywrightError,
    ):
        # 특정 X 게시물의 navigation/runtime 실패가
        # 전체 pipeline을 중단시키지 않도록 미추출로 처리한다.
        return None

    finally:
        # BrowserContext는 전체 작업 동안 재사용하고,
        # 이 게시물용 tab만 닫는다.
        try:
            page.close()
        except Exception:
            pass


# =============================================================================
# Facebook
# =============================================================================

FB_PROFILE_DIR = (
    PROJECT_ROOT
    / ".fb_browser_profile"
)


def _fb_normalize_url(
    href: str | None,
) -> str | None:
    if not href:
        return None

    href = html.unescape(
        href.strip()
    )

    if not href:
        return None

    if href.startswith("//"):
        return f"https:{href}"

    if href.startswith("/"):
        return urljoin(
            "https://www.facebook.com",
            href,
        )

    if href.startswith(
        ("http://", "https://")
    ):
        return href

    return None


def _fb_is_domain(
    host: str,
) -> bool:
    host = (
        host.lower()
        .split(":")[0]
    )

    return (
        host == "facebook.com"
        or host.endswith(
            ".facebook.com"
        )
    )


def _fb_profile_aliases_from_url(
    url: str | None,
) -> set[str]:
    aliases: set[str] = set()

    if not url:
        return aliases

    try:
        parsed = urlparse(
            url
        )
    except Exception:
        return aliases

    if not _fb_is_domain(
        parsed.netloc
    ):
        return aliases

    path = (
        parsed.path
        .strip("/")
    )

    query = parse_qs(
        parsed.query
    )

    if path.lower() == "profile.php":
        ids = query.get("id")

        if (
            ids
            and ids[0]
            and ids[0].isdigit()
        ):
            aliases.add(
                f"id:{ids[0]}"
            )

        return aliases

    parts = [
        part
        for part in path.split("/")
        if part
    ]

    if not parts:
        return aliases

    if (
        len(parts) >= 3
        and parts[0].lower()
        in {
            "people",
            "pages",
        }
    ):
        possible_id = parts[-1]

        if possible_id.isdigit():
            aliases.add(
                f"id:{possible_id}"
            )

        return aliases

    reserved = {
        "groups",
        "watch",
        "reel",
        "reels",
        "events",
        "marketplace",
        "gaming",
        "login",
        "help",
        "settings",
        "messages",
        "notifications",
        "story.php",
        "permalink.php",
        "photo.php",
        "video.php",
        "share",
    }

    first = parts[0].lower()

    if first in reserved:
        return aliases

    if len(parts) == 1:
        if first.isdigit():
            aliases.add(
                f"id:{first}"
            )
        else:
            aliases.add(
                f"username:{first}"
            )

    return aliases


def _fb_post_author_aliases_from_url(
    post_url: str,
) -> set[str]:
    aliases: set[str] = set()

    try:
        parsed = urlparse(
            post_url
        )
    except Exception:
        return aliases

    query = parse_qs(
        parsed.query
    )

    owner_ids = query.get("id")

    if (
        owner_ids
        and owner_ids[0]
        and owner_ids[0].isdigit()
    ):
        aliases.add(
            f"id:{owner_ids[0]}"
        )

    parts = [
        part
        for part in (
            parsed.path
            .strip("/")
            .split("/")
        )
        if part
    ]

    post_markers = {
        "posts",
        "videos",
        "photos",
        "reel",
        "reels",
    }

    for index, part in enumerate(
        parts
    ):
        if (
            part.lower()
            not in post_markers
            or index < 1
        ):
            continue

        owner = (
            parts[index - 1]
            .lower()
        )

        if owner.isdigit():
            aliases.add(
                f"id:{owner}"
            )
        else:
            aliases.add(
                f"username:{owner}"
            )

        break

    return aliases


def _fb_decode_comment_reference(
    value: str | None,
) -> str | None:
    if not value:
        return None

    value = value.strip()

    if value.isdigit():
        return None

    try:
        padded = (
            value
            + "=" * (
                (-len(value)) % 4
            )
        )

        decoded = (
            base64
            .urlsafe_b64decode(
                padded
            )
            .decode("utf-8")
        )

    except Exception:
        return None

    match = re.fullmatch(
        r"comment:(\d+)_(\d+)",
        decoded,
    )

    return (
        match.group(2)
        if match
        else None
    )


def _fb_build_comment_author_map(
    page,
) -> dict[
    str,
    tuple[set[str], str],
]:
    result: dict[
        str,
        tuple[set[str], str],
    ] = {}

    links = page.locator(
        "a[href*='comment_id=']"
    )

    for index in range(
        links.count()
    ):
        try:
            href = (
                links
                .nth(index)
                .get_attribute("href")
            )
        except Exception:
            continue

        full_url = (
            _fb_normalize_url(
                href
            )
        )

        if not full_url:
            continue

        try:
            query = parse_qs(
                urlparse(
                    full_url
                ).query
            )
        except Exception:
            continue

        values = query.get(
            "comment_id"
        )

        if not values:
            continue

        real_comment_id = (
            _fb_decode_comment_reference(
                values[0]
            )
        )

        if not real_comment_id:
            continue

        aliases = (
            _fb_profile_aliases_from_url(
                full_url
            )
        )

        if not aliases:
            continue

        result.setdefault(
            real_comment_id,
            (
                aliases,
                full_url,
            ),
        )

    return result


def _fb_is_comment_permalink(
    url: str,
) -> bool:
    try:
        parsed = urlparse(
            url
        )
    except Exception:
        return False

    if not _fb_is_domain(
        parsed.netloc
    ):
        return False

    query = parse_qs(
        parsed.query
    )

    values = query.get(
        "comment_id"
    )

    if (
        not values
        or not values[0].isdigit()
    ):
        return False

    path = parsed.path.lower()

    patterns = (
        "/posts/",
        "/story.php",
        "/permalink.php",
        "/photo.php",
        "/videos/",
        "/video.php",
        "/reel/",
    )

    if any(
        pattern in path
        for pattern in patterns
    ):
        return True

    return any(
        key in query
        for key in (
            "story_fbid",
            "fbid",
            "v",
        )
    )


def _fb_clean_comment_url(
    url: str,
) -> str:
    parsed = urlparse(
        url
    )

    values = parse_qs(
        parsed.query
    ).get("comment_id")

    if not values:
        return url

    return (
        f"{parsed.scheme}://"
        f"{parsed.netloc}"
        f"{parsed.path}"
        f"?comment_id={values[0]}"
    )


def _fb_enrich_post_author_aliases(
    aliases: set[str],
    comment_permalink_url: str,
) -> None:
    aliases.update(
        _fb_post_author_aliases_from_url(
            comment_permalink_url
        )
    )


def _fb_find_consumer_comment_url(
    page,
    post_author_aliases: set[str],
) -> str | None:
    author_map = (
        _fb_build_comment_author_map(
            page
        )
    )

    if not author_map:
        return None

    links = page.locator(
        "a[href*='comment_id=']"
    )

    seen_ids: set[str] = set()

    for index in range(
        links.count()
    ):
        try:
            href = (
                links
                .nth(index)
                .get_attribute("href")
            )
        except Exception:
            continue

        full_url = (
            _fb_normalize_url(
                href
            )
        )

        if not full_url:
            continue

        try:
            values = parse_qs(
                urlparse(
                    full_url
                ).query
            ).get("comment_id")
        except Exception:
            continue

        if not values:
            continue

        comment_id = (
            values[0]
            .strip()
        )

        if (
            not comment_id.isdigit()
            or comment_id in seen_ids
            or not _fb_is_comment_permalink(
                full_url
            )
        ):
            continue

        seen_ids.add(
            comment_id
        )

        # 숫자 ID 원본 URL과 username 기반 렌더링 URL이
        # 같은 페이지를 가리킬 수 있으므로 owner alias 보강.
        _fb_enrich_post_author_aliases(
            aliases=post_author_aliases,
            comment_permalink_url=(
                full_url
            ),
        )

        author_info = (
            author_map.get(
                comment_id
            )
        )

        # 작성자를 확인할 수 없으면 skip.
        if not author_info:
            continue

        comment_author_aliases = (
            author_info[0]
        )

        if (
            post_author_aliases
            & comment_author_aliases
        ):
            if DEBUG:
                print(
                    "[FB SELF COMMENT 제외] "
                    f"id={comment_id}"
                )

            continue

        return _fb_clean_comment_url(
            full_url
        )

    return None


def _fb_requires_login(
    page,
) -> bool:
    current_url = (
        page.url.lower()
    )

    if (
        "/login" in current_url
        or "/checkpoint"
        in current_url
    ):
        return True

    try:
        return (
            page.locator(
                "input[name='email']"
            ).count() > 0
            and page.locator(
                "input[name='pass']"
            ).count() > 0
        )
    except Exception:
        return False


def _fb_handle_login(
    page,
    post_url: str,
    *,
    headless: bool = False,
) -> bool:
    if not _fb_requires_login(
        page
    ):
        return True

    if headless:
        print()
        print(
            "[FB] 로그인 세션이 필요하지만 현재 headless 모드입니다."
        )
        print(
            "최초 인증이 필요한 경우 BROWSER_HEADLESS=False로 "
            "1회 로그인한 뒤 다시 headless 모드로 실행하세요."
        )
        return False

    if not INTERACTIVE_LOGIN:
        return False

    print()
    print(
        "[FB] 최초 로그인이 필요합니다."
    )
    print(
        "열린 Chromium에서 로그인한 뒤 "
        "터미널로 돌아와 Enter를 누르세요."
    )

    input(
        "로그인 완료 후 Enter: "
    )

    page.goto(
        post_url,
        wait_until="domcontentloaded",
        timeout=PAGE_TIMEOUT_MS,
    )

    page.wait_for_timeout(
        4_000
    )

    return not _fb_requires_login(
        page
    )


def _fb_try_expand_comments(
    page,
) -> None:
    keywords = (
        "more comments",
        "view more comments",
        "see more comments",
        "previous comments",
        "view previous comments",
        "댓글 더 보기",
        "댓글 더보기",
        "이전 댓글",
        "댓글 보기",
    )

    try:
        buttons = page.locator(
            "[role='button']"
        )

        count = min(
            buttons.count(),
            300,
        )
    except Exception:
        return

    for index in range(count):
        button = buttons.nth(
            index
        )

        try:
            text = (
                button.inner_text(
                    timeout=300
                )
                or ""
            ).strip()
        except Exception:
            continue

        if (
            not text
            or not any(
                keyword in text.lower()
                for keyword in keywords
            )
        ):
            continue

        try:
            if not button.is_visible():
                continue

            button.click(
                timeout=1_000
            )

            page.wait_for_timeout(
                1_000
            )

            return
        except Exception:
            continue


def extract_facebook_comment_url(
    post_url: str,
    *,
    session: CommentExtractorSession | None = None,
) -> str | None:
    try:
        from playwright.sync_api import (
            TimeoutError as PlaywrightTimeoutError,
        )
    except ImportError as exc:
        raise RuntimeError(
            "Playwright가 설치되어 있지 않습니다."
        ) from exc

    post_author_aliases = (
        _fb_post_author_aliases_from_url(
            post_url
        )
    )

    # 정확성 우선:
    # 게시물 작성자를 식별할 수 없으면 self-comment 제외를 보장할 수 없다.
    if not post_author_aliases:
        return None

    active_session = (
        session
        or get_default_comment_extractor_session()
    )

    context = (
        active_session.get_fb_context()
    )

    page = context.new_page()

    try:
        page.set_default_timeout(
            10_000
        )

        page.goto(
            post_url,
            wait_until="domcontentloaded",
            timeout=PAGE_TIMEOUT_MS,
        )

        page.wait_for_timeout(
            4_000
        )

        if not _fb_handle_login(
            page=page,
            post_url=post_url,
            headless=active_session.headless,
        ):
            return None

        for _ in range(10):
            _fb_try_expand_comments(
                page
            )

            page.wait_for_timeout(
                1_000
            )

            result = (
                _fb_find_consumer_comment_url(
                    page=page,
                    post_author_aliases=(
                        post_author_aliases
                    ),
                )
            )

            if result:
                return result

            page.mouse.wheel(
                0,
                1200,
            )

            page.wait_for_timeout(
                1_500
            )

        return None

    except PlaywrightTimeoutError:
        return None

    finally:
        # BrowserContext는 전체 작업 동안 재사용하고,
        # 이 게시물용 tab만 닫는다.
        try:
            page.close()
        except Exception:
            pass


# =============================================================================
# TikTok
# =============================================================================

TIKTOK_CDP_ENDPOINT = (
    "http://127.0.0.1:9222"
)

TIKTOK_PROFILE_DIR = (
    PROJECT_ROOT
    / ".tiktok_edge_profile"
)

TIKTOK_COMMENT_WAIT_SECONDS = 20.0


@dataclass(frozen=True)
class TikTokCommentCandidate:
    comment_id: str
    author_username: str
    text: str | None
    direct_url: str | None = None


def _tt_clean_post_url(
    post_url: str,
) -> str:
    parts = urlsplit(
        post_url.strip()
    )

    return urlunsplit(
        (
            parts.scheme or "https",
            parts.netloc,
            parts.path.rstrip("/"),
            "",
            "",
        )
    )


def _tt_parse_post_identity(
    post_url: str,
) -> tuple[str, str]:
    match = re.search(
        r"(?:https?://)?(?:www\.)?"
        r"tiktok\.com/@([^/?#]+)/video/(\d+)",
        post_url,
        flags=re.IGNORECASE,
    )

    if not match:
        raise ValueError(
            "지원하는 TikTok 게시물 URL 형식이 아닙니다: "
            f"{post_url}"
        )

    username = (
        normalize_identity(
            match.group(1)
        )
    )

    if not username:
        raise ValueError(
            "TikTok username을 식별하지 못했습니다."
        )

    return (
        username,
        match.group(2),
    )


def _tt_first_scalar(
    mapping: dict[str, Any],
    keys: Iterable[str],
) -> Any:
    for key in keys:
        value = mapping.get(
            key
        )

        if (
            value is not None
            and not isinstance(
                value,
                (dict, list),
            )
        ):
            return value

    return None


def _tt_extract_username(
    obj: dict[str, Any],
) -> str | None:
    user = obj.get("user")

    if isinstance(user, dict):
        value = _tt_first_scalar(
            user,
            (
                "unique_id",
                "uniqueId",
                "username",
                "user_name",
                "userName",
            ),
        )

        normalized = (
            normalize_identity(
                value
            )
        )

        if normalized:
            return normalized

    return normalize_identity(
        _tt_first_scalar(
            obj,
            (
                "unique_id",
                "uniqueId",
                "username",
                "user_name",
                "userName",
            ),
        )
    )


def _tt_direct_url(
    obj: dict[str, Any],
) -> str | None:
    for key in (
        "share_url",
        "shareUrl",
        "comment_url",
        "commentUrl",
        "permalink",
        "url",
    ):
        value = obj.get(
            key
        )

        if not isinstance(
            value,
            str,
        ):
            continue

        value = value.strip()

        if (
            value.startswith(
                ("http://", "https://")
            )
            and "tiktok.com"
            in value.casefold()
        ):
            return value

    return None


def _tt_parse_comment_object(
    obj: dict[str, Any],
) -> TikTokCommentCandidate | None:
    raw_id = _tt_first_scalar(
        obj,
        (
            "cid",
            "comment_id",
            "commentId",
        ),
    )

    if raw_id is None:
        has_text = any(
            key in obj
            for key in (
                "text",
                "comment_text",
                "commentText",
            )
        )

        has_user = (
            isinstance(
                obj.get("user"),
                dict,
            )
            or any(
                key in obj
                for key in (
                    "unique_id",
                    "uniqueId",
                    "username",
                    "user_name",
                    "userName",
                )
            )
        )

        if (
            has_text
            and has_user
        ):
            raw_id = obj.get("id")

    if raw_id is None:
        return None

    comment_id = str(
        raw_id
    ).strip()

    if not comment_id.isdigit():
        return None

    author = (
        _tt_extract_username(
            obj
        )
    )

    if not author:
        return None

    text_value = _tt_first_scalar(
        obj,
        (
            "text",
            "comment_text",
            "commentText",
        ),
    )

    return TikTokCommentCandidate(
        comment_id=comment_id,
        author_username=author,
        text=(
            str(text_value)
            if text_value is not None
            else None
        ),
        direct_url=_tt_direct_url(
            obj
        ),
    )


def _tt_iter_dicts(
    value: Any,
):
    if isinstance(value, dict):
        yield value

        for child in value.values():
            yield from _tt_iter_dicts(
                child
            )

    elif isinstance(value, list):
        for child in value:
            yield from _tt_iter_dicts(
                child
            )


class _TikTokCollector:
    def __init__(
        self,
        post_author: str,
    ) -> None:
        self.post_author = post_author
        self._candidates: dict[
            str,
            TikTokCommentCandidate,
        ] = {}
        self._lock = threading.Lock()

    def on_response(
        self,
        response,
    ) -> None:
        try:
            resource_type = (
                response.request
                .resource_type
                .casefold()
            )

            response_url = (
                response.url
            )
        except Exception:
            return

        if resource_type not in {
            "xhr",
            "fetch",
        }:
            return

        if (
            "comment"
            not in response_url.casefold()
        ):
            return

        try:
            payload = response.json()
        except Exception:
            return

        for mapping in _tt_iter_dicts(
            payload
        ):
            candidate = (
                _tt_parse_comment_object(
                    mapping
                )
            )

            if not candidate:
                continue

            with self._lock:
                self._candidates.setdefault(
                    candidate.comment_id,
                    candidate,
                )

    def first_consumer(
        self,
    ) -> TikTokCommentCandidate | None:
        with self._lock:
            candidates = list(
                self._candidates.values()
            )

        for candidate in candidates:
            if (
                normalize_identity(
                    candidate.author_username
                )
                == self.post_author
            ):
                if DEBUG:
                    print(
                        "[TT SELF COMMENT 제외] "
                        f"id={candidate.comment_id}"
                    )
                continue

            return candidate

        return None


def _tt_build_comment_url(
    post_url: str,
    comment_id: str,
) -> str:
    cid = (
        base64
        .urlsafe_b64encode(
            comment_id.encode(
                "utf-8"
            )
        )
        .decode("ascii")
        .rstrip("=")
    )

    return (
        f"{_tt_clean_post_url(post_url)}"
        f"?cid={cid}"
    )


def _tt_cdp_available() -> bool:
    try:
        with urlopen(
            (
                f"{TIKTOK_CDP_ENDPOINT}"
                "/json/version"
            ),
            timeout=1.5,
        ) as response:
            return (
                200
                <= response.status
                < 300
            )
    except Exception:
        return False


def _find_edge_executable() -> Path | None:
    import os

    candidates: list[Path] = []

    for env_name in (
        "PROGRAMFILES",
        "PROGRAMFILES(X86)",
        "LOCALAPPDATA",
    ):
        root = os.getenv(
            env_name
        )

        if not root:
            continue

        candidates.append(
            Path(root)
            / "Microsoft"
            / "Edge"
            / "Application"
            / "msedge.exe"
        )

    for candidate in candidates:
        if candidate.is_file():
            return candidate

    return None


def _tt_start_edge_if_needed() -> None:
    """
    운영 시 start_tiktok_edge.ps1을 매번 수동 실행하지 않도록 한다.
    """

    if _tt_cdp_available():
        return

    edge_path = (
        _find_edge_executable()
    )

    if edge_path is None:
        raise RuntimeError(
            "Microsoft Edge 실행 파일을 찾지 못했습니다."
        )

    TIKTOK_PROFILE_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    subprocess.Popen(
        [
            str(edge_path),
            "--headless=new",
            "--remote-debugging-port=9222",
            (
                "--user-data-dir="
                f"{TIKTOK_PROFILE_DIR}"
            ),
            "--no-first-run",
            "--no-default-browser-check",
            "https://www.tiktok.com/",
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )

    deadline = (
        time.monotonic()
        + 15.0
    )

    while (
        time.monotonic()
        < deadline
    ):
        if _tt_cdp_available():
            return

        time.sleep(
            0.5
        )

    raise RuntimeError(
        "TikTok 전용 Edge는 실행했지만 "
        "CDP 9222 포트가 준비되지 않았습니다."
    )


def _tt_trigger_comment_loading(
    page,
) -> None:
    selectors = (
        "[data-e2e*='comment']",
        "button[aria-label*='comment' i]",
        (
            "[role='button']"
            "[aria-label*='comment' i]"
        ),
    )

    for selector in selectors:
        try:
            locator = page.locator(
                selector
            )

            count = min(
                locator.count(),
                20,
            )
        except Exception:
            continue

        for index in range(
            count
        ):
            candidate = locator.nth(
                index
            )

            try:
                if not candidate.is_visible():
                    continue

                candidate.click(
                    timeout=1_000
                )

                page.wait_for_timeout(
                    1_000
                )

                return
            except Exception:
                continue


def extract_tiktok_comment_url(
    post_url: str,
    *,
    session: CommentExtractorSession | None = None,
) -> str | None:
    try:
        from playwright.sync_api import (
            Error as PlaywrightError,
        )
    except ImportError as exc:
        raise RuntimeError(
            "Playwright가 설치되어 있지 않습니다."
        ) from exc

    (
        post_author,
        video_id,
    ) = _tt_parse_post_identity(
        post_url
    )

    active_session = (
        session
        or get_default_comment_extractor_session()
    )

    try:
        context = (
            active_session.get_tiktok_context()
        )
    except PlaywrightError:
        return None

    page = context.new_page()

    try:
        collector = (
            _TikTokCollector(
                post_author
            )
        )

        page.on(
            "response",
            collector.on_response,
        )

        page.set_default_timeout(
            10_000
        )

        page.goto(
            post_url,
            wait_until="domcontentloaded",
            timeout=PAGE_TIMEOUT_MS,
        )

        page.wait_for_timeout(
            3_000
        )

        canonical_url = (
            _tt_clean_post_url(
                post_url
            )
        )

        try:
            (
                resolved_author,
                resolved_video_id,
            ) = _tt_parse_post_identity(
                page.url
            )

            if (
                resolved_video_id
                == video_id
            ):
                collector.post_author = (
                    resolved_author
                )

                canonical_url = (
                    _tt_clean_post_url(
                        page.url
                    )
                )
        except ValueError:
            pass

        _tt_trigger_comment_loading(
            page
        )

        deadline = (
            time.monotonic()
            + TIKTOK_COMMENT_WAIT_SECONDS
        )

        while (
            time.monotonic()
            < deadline
        ):
            candidate = (
                collector.first_consumer()
            )

            if candidate:
                return (
                    candidate.direct_url
                    or _tt_build_comment_url(
                        post_url=canonical_url,
                        comment_id=(
                            candidate.comment_id
                        ),
                    )
                )

            try:
                page.mouse.wheel(
                    0,
                    700,
                )
            except Exception:
                pass

            page.wait_for_timeout(
                1_000
            )

        return None

    finally:
        # Edge/CDP browser와 context는 전체 작업 동안 재사용하고,
        # 이 게시물용 tab만 닫는다.
        try:
            page.close()
        except Exception:
            pass


# =============================================================================
# Dispatcher
# =============================================================================

def extract_comment_url(
    channel: object,
    post_url: object,
    *,
    session: CommentExtractorSession | None = None,
    raise_on_error: bool = False,
) -> str | None:
    """
    raw_to_processed.py가 호출할 단일 public function.

    성공:
        댓글 URL 반환

    댓글 없음 / 추출 실패:
        None 반환

    raise_on_error=False가 기본값이므로
    특정 행의 댓글 추출 실패가 전체 pipeline을 중단하지 않는다.

    session을 전달하면 X / FB / TT browser context를
    해당 session 범위에서 재사용한다.

    댓글별 상세 로그는 출력하지 않고, session 종료 시
    전체 시도/성공/미추출 건수만 요약 출력한다.
    """

    normalized_channel = (
        normalize_channel(
            channel
        )
    )

    normalized_post_url = (
        normalize_url(
            post_url
        )
    )

    if (
        normalized_channel is None
        or normalized_post_url is None
    ):
        return None

    # raw_to_processed.py는 하나의 session을 재사용하므로
    # 모든 플랫폼의 댓글 추출 결과를 동일한 summary에 집계한다.
    stats_session = (
        session
        or get_default_comment_extractor_session()
    )

    try:
        if normalized_channel == "YT":
            comment_url = (
                extract_youtube_comment_url(
                    normalized_post_url
                )
            )

        elif normalized_channel == "IG":
            comment_url = (
                extract_instagram_comment_url(
                    normalized_post_url
                )
            )

        elif normalized_channel == "X":
            comment_url = (
                extract_twitter_comment_url(
                    normalized_post_url,
                    session=stats_session,
                )
            )

        elif normalized_channel == "FB":
            comment_url = (
                extract_facebook_comment_url(
                    normalized_post_url,
                    session=stats_session,
                )
            )

        elif normalized_channel == "TT":
            comment_url = (
                extract_tiktok_comment_url(
                    normalized_post_url,
                    session=stats_session,
                )
            )

        else:
            return None

        normalized_comment_url = (
            normalize_url(
                comment_url
            )
        )

        stats_session.record_comment_result(
            extracted=(
                normalized_comment_url
                is not None
            ),
            error=False,
        )

        return normalized_comment_url

    except Exception as exc:
        stats_session.record_comment_result(
            extracted=False,
            error=True,
        )

        message = (
            "[COMMENT EXTRACT FAILED] "
            f"channel={normalized_channel}, "
            f"url={normalized_post_url}, "
            f"{type(exc).__name__}: {exc}"
        )

        if raise_on_error:
            raise RuntimeError(
                message
            ) from exc

        # 운영 로그에는 행별 상세 오류를 출력하지 않는다.
        # 필요 시 logging DEBUG level에서만 확인 가능하다.
        LOGGER.debug(
            message,
            exc_info=True,
        )

        return None

def build_url_cell_value(
    channel: object,
    post_url: object,
    *,
    session: CommentExtractorSession | None = None,
    separator: str = "\n",
    raise_on_error: bool = False,
) -> str | None:
    """
    Excel URL 셀에 적재할 최종 문자열.

    댓글 성공:
        POST_URL
        COMMENT_URL

    댓글 없음/실패:
        POST_URL
        N/A
    """

    normalized_post_url = (
        normalize_url(
            post_url
        )
    )

    if normalized_post_url is None:
        return None

    comment_url = (
        extract_comment_url(
            channel=channel,
            post_url=normalized_post_url,
            session=session,
            raise_on_error=raise_on_error,
        )
    )

    if not comment_url:
        return separator.join(
            (
                normalized_post_url,
                "N/A",
            )
        )

    return separator.join(
        (
            normalized_post_url,
            comment_url,
        )
    )


# =============================================================================
# 단독 테스트
# =============================================================================

def main() -> None:
    print(
        "지원 Channel: YT, IG, X, FB, TT"
    )

    channel = input(
        "Channel: "
    ).strip()

    post_url = input(
        "게시물 URL: "
    ).strip()

    normalized_post_url = (
        normalize_url(
            post_url
        )
    )

    with CommentExtractorSession(
        headless=BROWSER_HEADLESS,
    ) as session:
        comment_url = (
            extract_comment_url(
                channel=channel,
                post_url=post_url,
                session=session,
                raise_on_error=False,
            )
        )

    if normalized_post_url is None:
        url_cell_value = None
    elif comment_url:
        url_cell_value = "\n".join(
            (
                normalized_post_url,
                comment_url,
            )
        )
    else:
        url_cell_value = "\n".join(
            (
                normalized_post_url,
                "N/A",
            )
        )

    print()
    print(
        "========================================"
    )
    print(
        "COMMENT URL"
    )
    print(
        "========================================"
    )
    print(
        comment_url
        if comment_url
        else "None"
    )

    print()
    print(
        "========================================"
    )
    print(
        "EXCEL URL CELL VALUE"
    )
    print(
        "========================================"
    )
    print(
        url_cell_value
        if url_cell_value
        else "None"
    )


if __name__ == "__main__":
    main()
