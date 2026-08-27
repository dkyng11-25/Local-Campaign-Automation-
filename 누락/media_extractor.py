from __future__ import annotations

import argparse
import ast
import os
import re
import shutil
import subprocess
import sys
from dataclasses import dataclass
from datetime import datetime
from enum import StrEnum
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, urlparse

import pandas as pd
import requests


# =========================================================
# 1. Configuration
# =========================================================

BASE_DIR = Path(__file__).resolve().parent

# 이 파일이 Local_Campaign_Automation_version4/누락/ 안에 있을 때:
#
# Excel:
#   누락/output_누락/{YYMMDD}_누락/
#
# Media:
#   누락/media_누락/{YYMMDD}_누락/
OUTPUT_DIR = BASE_DIR / "output_누락"
LOCAL_MEDIA_ROOT = BASE_DIR / "media_누락"

ENV_INPUT_DATE = "LOCAL_CAMPAIGN_INPUT_DATE"
ENV_RUN_NUMBER = "LOCAL_CAMPAIGN_RUN_NUMBER"
ENV_OUTPUT_DIR = "LOCAL_CAMPAIGN_OUTPUT_DIR"
ENV_MEDIA_DIR = "LOCAL_CAMPAIGN_MEDIA_DIR"

RAW_SHEET_NAMES = ["Raw Data_원문", "Raw Data_전략법인"]
LLM_INPUT_SHEET_NAME = "llm_input"

# Sprinklr Raw Data 시트의 헤더는 현재 1행이지만,
# 과거 파일의 2행 헤더도 처리할 수 있도록 상단 행에서 자동 탐색한다.
HEADER_SEARCH_MAX_ROWS = 10
HEADER_DETECTION_COLUMNS = {
    "Conversation Stream",
    "Campaign ID",
    "Permalink",
    "snType column",
}

REQUEST_TIMEOUT = (10, 60)
PROBE_BYTES = 64 * 1024
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/142.0 Safari/537.36"
)

# X 로그인 쿠키를 사용할 수 없는 환경을 기본값으로 둔다.
# True로 바꾸면 TWITTER_BROWSER에 로그인된 브라우저 쿠키를 사용한다.
TWITTER_USE_BROWSER_COOKIES = True
TWITTER_BROWSER = "edge"
TWITTER_INCLUDE_CARD_IMAGES = False

# TikTok도 기본적으로 쿠키 없이 먼저 시도한다.
# gallery-dl에서 로그인/쿠키 관련 오류가 발생할 때만 True로 변경한다.
# True인 경우 TIKTOK_BROWSER에서 TikTok에 로그인되어 있어야 한다.
TIKTOK_USE_BROWSER_COOKIES = False
TIKTOK_BROWSER = "edge"

# TikTok 게시물은 gallery-dl 대신 yt-dlp로 직접 다운로드한다.
# 로컬 테스트에서 성공한 Chrome impersonation을 우선 사용하며,
# 현재 환경에서 impersonation을 사용할 수 없으면 자동으로 일반 요청을 재시도한다.
TIKTOK_YT_DLP_IMPERSONATE = "chrome"
TIKTOK_YT_DLP_TIMEOUT_SECONDS = 300

SHOW_PROGRESS_LOGS = True
SHOW_SUCCESS_LOGS = False
SHOW_FAILURE_LOGS = False
GALLERY_DL_VERBOSE = False


# =========================================================
# 2. Enum / Data Models
# =========================================================

class Platform(StrEnum):
    YOUTUBE = "youtube"
    INSTAGRAM = "instagram"
    FACEBOOK = "facebook"
    TWITTER = "twitter"
    TIKTOK = "tiktok"
    UNKNOWN = "unknown"


class MediaType(StrEnum):
    VIDEO = "video"
    IMAGE = "image"
    LINK = "link"
    CAROUSEL = "carousel"
    UNKNOWN = "unknown"
    NONE = "none"


@dataclass(frozen=True, slots=True)
class AccountInfo:
    """게시자 분류 및 국가 판별에 사용할 Sender Profile 정보."""

    account_name: str | None = None
    account_handle: str | None = None
    follower_count: int | None = None
    profile_url: str | None = None
    location: str | None = None
    detailed_location: str | None = None
    bio: str | None = None
    website: str | None = None
    verified: bool | None = None
    verified_type: str | None = None
    profile_tags: str | None = None


@dataclass(frozen=True, slots=True)
class LoadedRawSheet:
    """Raw Data 시트와 실제 Excel 헤더 행 번호를 함께 보관한다."""

    dataframe: pd.DataFrame
    header_row: int


@dataclass(frozen=True, slots=True)
class SourceMedia:
    """Raw Data의 같은 순번에 있는 URL과 media type을 하나로 묶은 객체."""

    source_url: str
    media_type: MediaType


@dataclass(frozen=True, slots=True)
class StructuredMediaInput:
    """Excel의 게시글 한 행을 표준화한 게시글 단위 입력 객체."""

    campaign_id: str
    platform: Platform
    media_type: MediaType
    original_post_url: str | None
    conversation_stream: str | None
    account: AccountInfo
    source_medias: tuple[SourceMedia, ...] = ()
    source_sheet_name: str | None = None
    raw_row_number: int | None = None


@dataclass(slots=True)
class MediaAsset:
    """실제로 검사하고 다운로드할 개별 이미지 또는 영상."""

    campaign_id: str
    asset_id: str
    asset_index: int
    platform: Platform
    media_type: MediaType
    original_post_url: str | None
    source_url: str | None

    # 최초 입력 타입은 내부 라우팅 판단에만 사용하며 Excel에는 출력하지 않는다.
    input_media_type: MediaType = MediaType.UNKNOWN

    source_sheet_name: str | None = None
    raw_row_number: int | None = None

    http_status: int | None = None
    content_type: str | None = None
    file_size_bytes: int | None = None

    local_path: Path | None = None
    extraction_method: str | None = None
    status: str = "pending"
    error_message: str | None = None


@dataclass(frozen=True, slots=True)
class UrlFeasibilityResult:
    feasible: bool
    http_status: int | None
    content_type: str | None
    detected_media_type: MediaType
    extension: str | None
    error_message: str | None = None


class GalleryDlNoNativeMediaError(RuntimeError):
    """
    gallery-dl 실행 자체는 정상 완료되었지만 게시물에서
    실제 이미지/영상 URL을 하나도 찾지 못한 경우에만 사용한다.

    Twitter/X에서는 이 상태를 추출 실패가 아니라
    정상적인 text-only 게시물 후보로 처리한다.
    """


# 표준 컬럼명 -> 원본 Excel 컬럼명
COLUMN_MAPPINGS: dict[str, str] = {
    "conversation_stream": "Conversation Stream",
    "campaign_id": "Campaign ID",
    "profile_url": "Profile URL",
    "user_name": "User Name",
    "permalink": "Permalink",
    "platform": "snType column",
    "media_type": "Media Type",
    "source_url": "Media URL",

    # Sender Profile 컬럼
    "sender_profile_available": "Sender Profile Available",
    "sender_screen_name": "Sender Screen Name",
    "sender_follower_count": "Sender Follower Count",
    "sender_location": "Sender Location",
    "sender_detailed_location": "Sender Detailed Location",
    "sender_bio": "Sender Bio",
    "sender_website": "Sender Website",
    "sender_verified": "Sender Verified",
    "sender_verified_type": "Sender Verified Type",
    "sender_profile_tags": "Sender Profile Tags",
}

# 이 컬럼들은 원본 시트에 반드시 존재해야 함
REQUIRED_STANDARD_COLUMNS = {
    "campaign_id",
    "platform",
}

MANIFEST_COLUMNS = [
    "campaign_id",
    "asset_id",
    "asset_index",
    "platform",
    "media_type",
    "source_sheet_name",
    "raw_row_number",
    "original_post_url",
    "source_url",
    "http_status",
    "content_type",
    "file_size_bytes",
    "local_path",
    "extraction_method",
    "status",
    "error_message",
    "llm_input_mode",
    "llm_input_value",
    "llm_ready",
    "user_action_required",
    "user_action",
]

LLM_INPUT_COLUMNS = [
    "campaign_id",
    "source_sheet",
    "raw_row_number",
    "platform",
    "permalink",
    "conversation_stream",
    "user_name",
    "profile_url",

    # 게시자 판별 및 국가 판별용 Sender Profile 정보
    "sender_profile_available",
    "sender_screen_name",
    "sender_follower_count",
    "sender_location",
    "sender_detailed_location",
    "sender_bio",
    "sender_website",
    "sender_verified",
    "sender_verified_type",
    "sender_profile_tags",

    # 미디어 처리 결과
    "post_media_type",
    "media_count",
    "media_types",
    "media_source_urls",
    "local_media_paths",
    "successful_media_count",
    "failed_media_count",
    "extraction_methods",
    "status",
    "error_message",
    "llm_input_mode",
    "llm_input_value",
    "llm_ready",
    "user_action_required",
    "user_action",
]


# =========================================================
# 3. 실행 인자 및 경로 결정
# =========================================================

def parse_arguments() -> argparse.Namespace:
    """
    전체 파이프라인 또는 모듈 단독 실행에 필요한 선택 인자를 읽는다.

    전체 파이프라인:
        run_pipeline.py가 output/media 경로를 환경변수로 전달한다.

    단독 실행:
        --output-dir과 --media-dir로 기존 실행 차수 폴더를 지정한다.
    """

    parser = argparse.ArgumentParser(
        description=(
            "누락 전용 output_누락/media_누락 폴더를 사용하여 "
            "소셜 미디어 파일을 추출합니다."
        )
    )

    parser.add_argument(
        "--output-dir",
        type=Path,
        help=(
            "누락 output 폴더 override용. "
            "현재 main은 날짜 기준 output_누락/{YYMMDD}_누락을 사용합니다."
        ),
    )

    parser.add_argument(
        "--media-dir",
        type=Path,
        help=(
            "누락 media 폴더 override용. "
            "현재 main은 날짜 기준 media_누락/{YYMMDD}_누락을 사용합니다."
        ),
    )

    parser.add_argument(
        "--overwrite",
        action="store_true",
        help=(
            "기존 campaign_media_result.xlsx와 미디어 결과가 있을 때 "
            "새 결과가 모두 완성된 후 기존 결과를 교체합니다."
        ),
    )

    return parser.parse_args()


def resolve_path_from_project(
    path: Path,
) -> Path:
    """
    상대경로는 현재 터미널 위치가 아니라 프로젝트 루트 기준으로 해석한다.
    """

    expanded_path = path.expanduser()

    if not expanded_path.is_absolute():
        expanded_path = BASE_DIR / expanded_path

    return expanded_path.resolve()


def extract_run_number_from_directory_name(
    directory_name: str,
    input_date: str,
) -> int:
    """
    실행 폴더명에서 차수를 추출한다.

    예:
        260805      -> 1
        260805_1차  -> 1
        260805_2차  -> 2
    """

    if directory_name == input_date:
        return 1

    match = re.fullmatch(
        rf"{re.escape(input_date)}_(\d+)차",
        directory_name,
    )

    if match is None:
        raise ValueError(
            "실행 폴더명이 입력 날짜와 일치하지 않습니다.\n"
            f"입력 날짜: {input_date}\n"
            f"폴더명: {directory_name}\n"
            "허용 형식 예: "
            f"{input_date}, {input_date}_2차"
        )

    run_number = int(match.group(1))

    if run_number < 1:
        raise ValueError(
            "실행 차수는 1 이상이어야 합니다: "
            f"{directory_name}"
        )

    return run_number


def resolve_execution_directories(
    input_date: str,
    cli_output_dir: Path | None,
    cli_media_dir: Path | None,
) -> tuple[Path, Path, int]:
    """
    이번 모듈이 사용할 기존 output/media 실행 폴더를 확정한다.

    각 경로의 우선순위:
        1. 명령행 인자
        2. run_pipeline.py가 전달한 환경변수
        3. 둘 다 없으면 오류

    이 함수는 새로운 차수 폴더를 생성하지 않는다.
    """

    pipeline_input_date = os.getenv(
        ENV_INPUT_DATE
    )

    if (
        pipeline_input_date
        and pipeline_input_date != input_date
    ):
        raise ValueError(
            "run_pipeline.py에서 전달된 작업 날짜와 "
            "입력 날짜가 일치하지 않습니다.\n"
            f"전달된 작업 날짜: {pipeline_input_date}\n"
            f"입력한 날짜: {input_date}"
        )

    environment_output_dir_text = os.getenv(
        ENV_OUTPUT_DIR
    )
    environment_media_dir_text = os.getenv(
        ENV_MEDIA_DIR
    )

    environment_output_dir = (
        resolve_path_from_project(
            Path(environment_output_dir_text)
        )
        if environment_output_dir_text
        else None
    )
    environment_media_dir = (
        resolve_path_from_project(
            Path(environment_media_dir_text)
        )
        if environment_media_dir_text
        else None
    )

    resolved_cli_output_dir = (
        resolve_path_from_project(
            cli_output_dir
        )
        if cli_output_dir is not None
        else None
    )
    resolved_cli_media_dir = (
        resolve_path_from_project(
            cli_media_dir
        )
        if cli_media_dir is not None
        else None
    )

    if (
        resolved_cli_output_dir is not None
        and environment_output_dir is not None
        and resolved_cli_output_dir != environment_output_dir
    ):
        raise ValueError(
            "--output-dir과 run_pipeline.py가 전달한 "
            "output 경로가 서로 다릅니다.\n"
            f"--output-dir: {resolved_cli_output_dir}\n"
            f"환경변수 경로: {environment_output_dir}"
        )

    if (
        resolved_cli_media_dir is not None
        and environment_media_dir is not None
        and resolved_cli_media_dir != environment_media_dir
    ):
        raise ValueError(
            "--media-dir과 run_pipeline.py가 전달한 "
            "media 경로가 서로 다릅니다.\n"
            f"--media-dir: {resolved_cli_media_dir}\n"
            f"환경변수 경로: {environment_media_dir}"
        )

    output_dir = (
        resolved_cli_output_dir
        or environment_output_dir
    )
    media_dir = (
        resolved_cli_media_dir
        or environment_media_dir
    )

    if output_dir is None or media_dir is None:
        missing_values: list[str] = []

        if output_dir is None:
            missing_values.append("output 경로")
        if media_dir is None:
            missing_values.append("media 경로")

        raise RuntimeError(
            "실행 경로가 모두 지정되지 않았습니다.\n"
            f"누락: {', '.join(missing_values)}\n"
            "전체 실행은 run_pipeline.py를 통해 시작하세요.\n"
            "기존 차수에서 이 모듈만 단독 실행하는 경우:\n"
            "python media_extractor.py "
            '--output-dir "output\\260805_2차" '
            '--media-dir "media\\260805_2차"'
        )

    for path_name, directory in (
        ("output", output_dir),
        ("media", media_dir),
    ):
        if not directory.exists():
            raise FileNotFoundError(
                f"지정된 {path_name} 실행 폴더를 찾을 수 없습니다.\n"
                f"경로: {directory}\n"
                "이 모듈은 차수 폴더를 생성하지 않습니다."
            )

        if not directory.is_dir():
            raise NotADirectoryError(
                f"지정된 {path_name} 경로가 폴더가 아닙니다.\n"
                f"경로: {directory}"
            )

    output_run_number = extract_run_number_from_directory_name(
        directory_name=output_dir.name,
        input_date=input_date,
    )
    media_run_number = extract_run_number_from_directory_name(
        directory_name=media_dir.name,
        input_date=input_date,
    )

    if output_dir.name != media_dir.name:
        raise ValueError(
            "output 폴더와 media 폴더의 실행 차수가 일치하지 않습니다.\n"
            f"output 폴더: {output_dir}\n"
            f"media 폴더: {media_dir}"
        )

    if output_run_number != media_run_number:
        raise ValueError(
            "output/media 실행 차수 계산 결과가 일치하지 않습니다.\n"
            f"output 차수: {output_run_number}\n"
            f"media 차수: {media_run_number}"
        )

    environment_run_number_text = os.getenv(
        ENV_RUN_NUMBER
    )

    if environment_run_number_text:
        try:
            environment_run_number = int(
                environment_run_number_text
            )
        except ValueError as exc:
            raise ValueError(
                "LOCAL_CAMPAIGN_RUN_NUMBER가 정수가 아닙니다: "
                f"{environment_run_number_text}"
            ) from exc

        if environment_run_number != output_run_number:
            raise ValueError(
                "run_pipeline.py가 전달한 실행 차수와 "
                "폴더명에서 계산한 차수가 일치하지 않습니다.\n"
                f"전달 차수: {environment_run_number}\n"
                f"폴더 차수: {output_run_number}"
            )

    print(
        f"[INFO] 실행 차수: {output_run_number}차"
    )
    print(
        f"[INFO] 실행 output 폴더: {output_dir}"
    )
    print(
        f"[INFO] 실행 media 폴더: {media_dir}"
    )

    return (
        output_dir,
        media_dir,
        output_run_number,
    )


def resolve_missing_execution_directories(
    input_date: str,
) -> tuple[Path, Path]:
    """
    누락 서브 파이프라인의 날짜별 output/media 폴더를 확정한다.

    이 파일이 다음 위치에 있다고 가정한다.
        Local_Campaign_Automation_version4/누락/media_extractor.py

    Excel 폴더:
        Local_Campaign_Automation_version4/
        누락/output_누락/{YYMMDD}_누락/

    Media 폴더:
        Local_Campaign_Automation_version4/
        누락/media_누락/{YYMMDD}_누락/
    """

    output_dir = (
        OUTPUT_DIR
        / f"{input_date}_누락"
    )

    media_dir = (
        LOCAL_MEDIA_ROOT
        / f"{input_date}_누락"
    )

    # 입력 Excel은 이미 output 날짜 폴더에 있어야 한다.
    if not output_dir.exists():
        raise FileNotFoundError(
            "누락 데이터 output 폴더를 찾을 수 없습니다.\n"
            f"경로: {output_dir}"
        )

    if not output_dir.is_dir():
        raise NotADirectoryError(
            "누락 데이터 output 경로가 폴더가 아닙니다.\n"
            f"경로: {output_dir}"
        )

    # 미디어 기준 폴더와 날짜 폴더는 이 모듈에서 안전하게 생성한다.
    media_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    print(
        f"[INFO] 누락 데이터 output 폴더: {output_dir}"
    )
    print(
        f"[INFO] 누락 데이터 media 폴더: {media_dir}"
    )

    return output_dir, media_dir


def build_excel_paths(
    input_date: str,
    output_dir: Path,
) -> tuple[Path, Path]:
    """
    output_누락/{YYMMDD}_누락 폴더에서 입력 Excel 경로를 생성한다.

    입력 및 최종 출력:
        {YYMMDD}_SLCC_SOV_Local Campaign Tracking_{월}월_v01.xlsx

    최종 결과는 원본 파일명을 그대로 유지한다.
    실제 원본 교체는 모든 처리가 성공한 뒤 수행한다.
    """

    try:
        input_date_obj = datetime.strptime(
            input_date,
            "%y%m%d",
        )
    except ValueError as exc:
        raise ValueError(
            "날짜는 YYMMDD 형식으로 입력해야 합니다. 예시) 260714"
        ) from exc

    input_month = input_date_obj.month

    input_excel_path = (
        output_dir
        / (
            f"{input_date}_SLCC_SOV_Local Campaign Tracking_"
            f"{input_month}월_v01.xlsx"
        )
    )

    # 입력 파일명을 그대로 유지한다.
    output_excel_path = input_excel_path

    if not input_excel_path.is_file():
        raise FileNotFoundError(
            f"입력 Excel 파일을 찾을 수 없습니다: {input_excel_path}"
        )

    return input_excel_path, output_excel_path


def directory_has_contents(
    directory: Path,
) -> bool:
    return any(
        directory.iterdir()
    )


def prepare_temporary_artifacts(
    input_excel_path: Path,
    output_excel_path: Path,
    media_dir: Path,
    overwrite: bool,
) -> tuple[Path, Path, Path]:
    """
    기존 결과를 직접 수정하지 않고 임시 Excel과 임시 media 폴더를 준비한다.

    반환:
        temporary_excel_path
        temporary_media_dir
        backup_media_dir
    """

    media_has_existing_results = directory_has_contents(
        media_dir
    )

    # 누락 처리에서는 입력 Excel과 최종 출력 Excel의 경로가 같다.
    # 따라서 입력 원본 파일의 존재 자체는 기존 처리 결과로 보지 않는다.
    output_is_input_file = (
        output_excel_path.resolve()
        == input_excel_path.resolve()
    )
    output_has_existing_result = (
        output_excel_path.exists()
        and not output_is_input_file
    )

    if (
        output_has_existing_result
        or media_has_existing_results
    ) and not overwrite:
        existing_items: list[str] = []

        if output_has_existing_result:
            existing_items.append(
                str(output_excel_path)
            )

        if media_has_existing_results:
            existing_items.append(
                f"{media_dir} 내부 미디어 파일"
            )

        raise FileExistsError(
            "동일한 실행 차수의 미디어 처리 결과가 이미 존재합니다.\n"
            + "\n".join(
                f"- {item}"
                for item in existing_items
            )
            + "\n기존 결과를 새 결과로 교체하려면 "
            "--overwrite 옵션을 명시하세요."
        )

    temporary_excel_path = output_excel_path.with_name(
        f".{output_excel_path.stem}.partial.xlsx"
    )

    temporary_media_dir = media_dir.with_name(
        f".{media_dir.name}.partial"
    )

    backup_media_dir = media_dir.with_name(
        f".{media_dir.name}.backup"
    )

    if backup_media_dir.exists():
        raise RuntimeError(
            "이전 실행의 media 백업 폴더가 남아 있습니다.\n"
            f"백업 폴더: {backup_media_dir}\n"
            "자동으로 삭제하지 않습니다. 기존 media 폴더와 내용을 "
            "확인한 뒤 수동으로 복구 또는 정리하세요."
        )

    if temporary_excel_path.exists():
        temporary_excel_path.unlink()

    if temporary_media_dir.exists():
        shutil.rmtree(
            temporary_media_dir
        )

    temporary_media_dir.mkdir(
        parents=False,
        exist_ok=False,
    )

    return (
        temporary_excel_path,
        temporary_media_dir,
        backup_media_dir,
    )


def rebase_media_asset_local_paths(
    media_assets: list[MediaAsset],
    temporary_media_dir: Path,
    final_media_dir: Path,
) -> list[MediaAsset]:
    """
    임시 media 폴더 경로를 최종 실행 media 폴더 경로로 변경한다.

    Excel에는 임시 폴더가 아니라 최종 폴더의 경로가 저장되어야 한다.
    """

    resolved_temporary_dir = temporary_media_dir.resolve()
    resolved_final_dir = final_media_dir.resolve()

    for asset in media_assets:
        if asset.local_path is None:
            continue

        resolved_local_path = asset.local_path.resolve()

        try:
            relative_path = resolved_local_path.relative_to(
                resolved_temporary_dir
            )
        except ValueError as exc:
            raise ValueError(
                "다운로드된 미디어가 임시 media 폴더 밖에 저장되었습니다.\n"
                f"미디어 경로: {resolved_local_path}\n"
                f"임시 폴더: {resolved_temporary_dir}"
            ) from exc

        asset.local_path = (
            resolved_final_dir
            / relative_path
        )

    return media_assets


def commit_output_artifacts(
    temporary_excel_path: Path,
    output_excel_path: Path,
    temporary_media_dir: Path,
    media_dir: Path,
    backup_media_dir: Path,
) -> None:
    """
    모든 미디어 처리와 임시 Excel 저장이 성공한 경우에만
    기존 media 폴더와 최종 Excel을 새 결과로 교체한다.
    """

    if not temporary_excel_path.is_file():
        raise FileNotFoundError(
            "최종 반영할 임시 Excel 파일이 없습니다: "
            f"{temporary_excel_path}"
        )

    if not temporary_media_dir.is_dir():
        raise FileNotFoundError(
            "최종 반영할 임시 media 폴더가 없습니다: "
            f"{temporary_media_dir}"
        )

    # 기존 실행 media 폴더를 백업한 뒤 임시 폴더를 최종 경로로 교체한다.
    media_dir.rename(
        backup_media_dir
    )

    try:
        temporary_media_dir.rename(
            media_dir
        )

        # 같은 output 폴더 안에서 완성된 임시 Excel만 최종 파일로 교체한다.
        os.replace(
            temporary_excel_path,
            output_excel_path,
        )

    except Exception:
        # media 교체 이후 Excel 반영이 실패하면 기존 media 폴더를 복구한다.
        if media_dir.exists():
            shutil.rmtree(
                media_dir,
                ignore_errors=True,
            )

        if backup_media_dir.exists():
            backup_media_dir.rename(
                media_dir
            )

        raise

    shutil.rmtree(
        backup_media_dir,
        ignore_errors=True,
    )


def cleanup_temporary_artifacts(
    temporary_excel_path: Path,
    temporary_media_dir: Path,
) -> None:
    """
    처리 실패 시 임시 산출물만 정리한다.
    기존 완성 결과는 변경하지 않는다.
    """

    if temporary_excel_path.exists():
        temporary_excel_path.unlink()

    if temporary_media_dir.exists():
        shutil.rmtree(
            temporary_media_dir,
            ignore_errors=True,
        )


# =========================================================
# 4. Excel Input / Sheet Functions
# =========================================================

def _normalize_header_value(value: Any) -> str | None:
    if value is None:
        return None

    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass

    text = str(value).strip()
    return text or None


def detect_header_row(
    excel_file: pd.ExcelFile,
    sheet_name: str,
    required_headers: set[str] = HEADER_DETECTION_COLUMNS,
    max_search_rows: int = HEADER_SEARCH_MAX_ROWS,
) -> int | None:
    """시트 상단에서 필수 헤더가 모두 존재하는 Excel 행을 찾는다."""

    preview_df = pd.read_excel(
        excel_file,
        sheet_name=sheet_name,
        header=None,
        nrows=max_search_rows,
    )

    if preview_df.dropna(how="all").empty:
        return None

    for dataframe_index, row in preview_df.iterrows():
        row_headers = {
            header
            for header in (
                _normalize_header_value(value)
                for value in row.tolist()
            )
            if header
        }

        if required_headers.issubset(row_headers):
            return int(dataframe_index) + 1

    return None


def load_raw_sheets(
    excel_path: Path,
    sheet_names: list[str],
) -> dict[str, LoadedRawSheet]:
    """Raw Data 시트를 읽고 각 시트의 실제 헤더 행을 자동 탐색한다.

    - 현재 파일의 1행 헤더와 과거 파일의 2행 헤더를 모두 지원한다.
    - 시트가 없거나 완전히 비어 있어도 빈 DataFrame으로 유지한다.
    """

    loaded_sheets: dict[str, LoadedRawSheet] = {}

    with pd.ExcelFile(excel_path) as excel_file:
        available_sheet_names = set(excel_file.sheet_names)

        for sheet_name in sheet_names:
            if sheet_name not in available_sheet_names:
                print(
                    f"[WARNING] 시트가 없어 빈 시트로 처리합니다: {sheet_name}"
                )
                loaded_sheets[sheet_name] = LoadedRawSheet(
                    dataframe=pd.DataFrame(
                        columns=list(COLUMN_MAPPINGS.values())
                    ),
                    header_row=1,
                )
                continue

            header_row = detect_header_row(
                excel_file=excel_file,
                sheet_name=sheet_name,
            )

            if header_row is None:
                preview_df = pd.read_excel(
                    excel_file,
                    sheet_name=sheet_name,
                    header=None,
                    nrows=HEADER_SEARCH_MAX_ROWS,
                )

                if preview_df.dropna(how="all").empty:
                    print(
                        f"[WARNING] 데이터와 헤더가 없는 빈 시트입니다: "
                        f"{sheet_name}"
                    )
                    loaded_sheets[sheet_name] = LoadedRawSheet(
                        dataframe=pd.DataFrame(
                            columns=list(COLUMN_MAPPINGS.values())
                        ),
                        header_row=1,
                    )
                    continue

                raise ValueError(
                    f"[{sheet_name}] 헤더 행을 찾지 못했습니다. "
                    f"필수 헤더: {sorted(HEADER_DETECTION_COLUMNS)}"
                )

            dataframe = pd.read_excel(
                excel_file,
                sheet_name=sheet_name,
                header=header_row - 1,
            )
            dataframe.columns = [
                str(column).strip()
                for column in dataframe.columns
            ]

            loaded_sheets[sheet_name] = LoadedRawSheet(
                dataframe=dataframe.copy(),
                header_row=header_row,
            )

            print(
                f"[INFO] {sheet_name}: header row={header_row}, "
                f"data rows={len(dataframe)}"
            )

    return loaded_sheets


def prepare_input_df(
    raw_df: pd.DataFrame,
    sheet_name: str,
    column_mapping: dict[str, str],
    header_row: int,
) -> pd.DataFrame:
    """필요 컬럼을 표준명으로 변경한다. 선택 컬럼이 없으면 pd.NA를 넣는다."""

    missing_required_columns = [
        column_mapping[standard_name]
        for standard_name in REQUIRED_STANDARD_COLUMNS
        if (
            standard_name not in column_mapping
            or column_mapping[standard_name] not in raw_df.columns
        )
    ]

    if missing_required_columns:
        raise ValueError(
            f"[{sheet_name}] 필수 컬럼이 없습니다: {missing_required_columns}"
        )

    available_mapping = {
        standard_name: original_name
        for standard_name, original_name in column_mapping.items()
        if original_name in raw_df.columns
    }

    selected_original_columns = list(
        dict.fromkeys(available_mapping.values())
    )
    prepared_df = raw_df[selected_original_columns].copy()

    rename_mapping = {
        original_name: standard_name
        for standard_name, original_name in available_mapping.items()
    }
    prepared_df = prepared_df.rename(columns=rename_mapping)

    for standard_name in column_mapping:
        if standard_name not in prepared_df.columns:
            prepared_df[standard_name] = pd.NA

    prepared_df = prepared_df[list(column_mapping.keys())]
    prepared_df["source_sheet"] = sheet_name

    # DataFrame index 0은 실제 헤더 행의 다음 Excel 행이다.
    prepared_df["raw_row_number"] = (
        raw_df.index.to_series().astype(int)
        + header_row
        + 1
    ).to_numpy()

    return prepared_df


def build_media_input_dataframe(
    raw_sheets: dict[str, LoadedRawSheet],
    column_mapping: dict[str, str],
) -> pd.DataFrame:
    prepared_dataframes: list[pd.DataFrame] = []

    for sheet_name, loaded_sheet in raw_sheets.items():
        prepared_df = prepare_input_df(
            raw_df=loaded_sheet.dataframe,
            sheet_name=sheet_name,
            column_mapping=column_mapping,
            header_row=loaded_sheet.header_row,
        )
        prepared_dataframes.append(prepared_df)

    if not prepared_dataframes:
        return pd.DataFrame(
            columns=[
                *column_mapping.keys(),
                "source_sheet",
                "raw_row_number",
            ]
        )

    return pd.concat(
        prepared_dataframes,
        ignore_index=True,
        sort=False,
    )


def clean_input_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    cleaned_df = df.copy()

    # media_type/source_url은 list 또는 줄바꿈 문자열일 수 있으므로
    # 여기서 무조건 astype("string") 하지 않는다.
    scalar_string_columns = [
        "conversation_stream",
        "campaign_id",
        "profile_url",
        "user_name",
        "permalink",
        "platform",
        "source_sheet",
        "sender_screen_name",
        "sender_location",
        "sender_detailed_location",
        "sender_bio",
        "sender_website",
        "sender_verified_type",
        "sender_profile_tags",
    ]

    for column in scalar_string_columns:
        if column not in cleaned_df.columns:
            continue

        cleaned_df[column] = (
            cleaned_df[column]
            .astype("string")
            .str.strip()
        )

    cleaned_df = cleaned_df.replace(
        {
            "": pd.NA,
            "nan": pd.NA,
            "None": pd.NA,
        }
    )

    if "sender_follower_count" in cleaned_df.columns:
        cleaned_df["sender_follower_count"] = (
            cleaned_df["sender_follower_count"]
            .apply(optional_follower_count)
            .astype("Int64")
        )

    cleaned_df = cleaned_df.dropna(
        subset=["campaign_id"]
    ).copy()

    return cleaned_df


# =========================================================
# 5. Input Mapper Helpers
# =========================================================

def optional_text(value: Any) -> str | None:
    if value is None:
        return None

    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass

    text = str(value).strip()
    return text or None


def optional_follower_count(value: Any) -> int | None:
    """Excel의 팔로워 수를 0 이상의 정수로 정규화한다.

    지원 예:
    - 1234567
    - 1,234,567
    - 1234567.0
    - 1.234567E+06
    """

    if value is None:
        return None

    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass

    if isinstance(value, bool):
        return None

    normalized_value = str(value).strip().replace(",", "")

    if not normalized_value:
        return None

    try:
        numeric_value = float(normalized_value)
    except (TypeError, ValueError):
        return None

    if numeric_value < 0:
        return None

    return int(round(numeric_value))


def optional_bool(value: Any) -> bool | None:
    if value is None:
        return None

    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass

    if isinstance(value, bool):
        return value

    if isinstance(value, (int, float)):
        if value == 1:
            return True
        if value == 0:
            return False

    text = str(value).strip().lower()

    if text in {"true", "t", "yes", "y", "1"}:
        return True

    if text in {"false", "f", "no", "n", "0"}:
        return False

    return None


def identify_platform(platform_value: Any) -> Platform:
    text = optional_text(platform_value)

    if text is None:
        return Platform.UNKNOWN

    platform_mapping = {
        "youtube": Platform.YOUTUBE,
        "instagram": Platform.INSTAGRAM,
        "facebook": Platform.FACEBOOK,
        "twitter": Platform.TWITTER,
        "x": Platform.TWITTER,
        "tiktok": Platform.TIKTOK,
    }

    return platform_mapping.get(
        text.lower(),
        Platform.UNKNOWN,
    )


def identify_media_type(media_type_value: Any) -> MediaType:
    text = optional_text(media_type_value)

    if text is None:
        return MediaType.UNKNOWN

    media_type_mapping = {
        "video": MediaType.VIDEO,
        "photo": MediaType.IMAGE,
        "image": MediaType.IMAGE,
        "picture": MediaType.IMAGE,
        "link": MediaType.LINK,
        "carousel": MediaType.CAROUSEL,
    }

    return media_type_mapping.get(
        text.lower(),
        MediaType.UNKNOWN,
    )


def normalize_cell_values(value: Any) -> tuple[str, ...]:
    """단일값, 리스트, 리스트 문자열, 줄바꿈 문자열을 문자열 tuple로 통일."""

    if value is None:
        return ()

    if isinstance(value, (list, tuple, set)):
        normalized_values: list[str] = []

        for item in value:
            text = optional_text(item)
            if not text:
                continue

            normalized_values.extend(
                line.strip()
                for line in text.splitlines()
                if line.strip()
            )

        return tuple(normalized_values)

    try:
        if pd.isna(value):
            return ()
    except (TypeError, ValueError):
        pass

    text = str(value).strip()
    if not text:
        return ()

    # Excel에 "['video', 'photo']" 같은 문자열로 저장된 경우도 지원
    if text.startswith("[") and text.endswith("]"):
        try:
            parsed_value = ast.literal_eval(text)
        except (ValueError, SyntaxError):
            parsed_value = None

        if isinstance(parsed_value, (list, tuple, set)):
            return normalize_cell_values(parsed_value)

    return tuple(
        line.strip()
        for line in text.splitlines()
        if line.strip()
    )


def prepare_campaign_input_for_processing(
    df: pd.DataFrame,
) -> pd.DataFrame:
    """campaign_input을 미디어 처리 및 향후 LLM 입력 생성에 맞게 정리한다.

    - YouTube는 Permalink를 LLM에 직접 전달한다.
    - 다른 플랫폼은 처리 결과를 나중에 행 단위로 반영한다.
    - final_url 계열 컬럼은 만들지 않는다.
    """

    prepared_df = df.copy()

    llm_columns: dict[str, Any] = {
        "resolved_media_type": pd.NA,
        "resolved_source_urls": pd.NA,
        "local_media_paths": pd.NA,
        "asset_statuses": pd.NA,
        "llm_input_mode": "pending_media_resolution",
        "llm_primary_input": pd.NA,
        "llm_ready": False,
        "user_action_required": False,
        "user_action": pd.NA,
        "llm_error_message": pd.NA,
    }

    for column_name, default_value in llm_columns.items():
        prepared_df[column_name] = default_value

    for row_index, row in prepared_df.iterrows():
        platform = identify_platform(row.get("platform"))
        permalink = optional_text(row.get("permalink"))
        source_urls = normalize_cell_values(row.get("source_url"))
        raw_media_types = normalize_cell_values(row.get("media_type"))

        normalized_media_types = tuple(
            identify_media_type(value).value
            for value in raw_media_types
        )

        if platform == Platform.YOUTUBE:
            youtube_url = permalink or (source_urls[0] if source_urls else None)
            prepared_df.at[row_index, "media_type"] = MediaType.VIDEO.value
            prepared_df.at[row_index, "resolved_media_type"] = MediaType.VIDEO.value
            prepared_df.at[row_index, "llm_input_mode"] = "youtube_url"

            if youtube_url:
                prepared_df.at[row_index, "source_url"] = youtube_url
                prepared_df.at[row_index, "resolved_source_urls"] = youtube_url
                prepared_df.at[row_index, "llm_primary_input"] = youtube_url
                prepared_df.at[row_index, "llm_ready"] = True
            else:
                prepared_df.at[row_index, "source_url"] = pd.NA
                prepared_df.at[row_index, "llm_input_mode"] = "manual_media_upload"
                prepared_df.at[row_index, "user_action_required"] = True
                prepared_df.at[row_index, "user_action"] = (
                    "게시물 URL 확인 후 미디어를 수동 확보하고 LLM에 직접 입력"
                )
                prepared_df.at[row_index, "llm_error_message"] = (
                    "YouTube Permalink와 Media URL이 모두 비어 있습니다."
                )
            continue

        if normalized_media_types:
            prepared_df.at[row_index, "media_type"] = "\n".join(normalized_media_types)
        else:
            prepared_df.at[row_index, "media_type"] = MediaType.UNKNOWN.value

        prepared_df.at[row_index, "source_url"] = (
            "\n".join(source_urls) if source_urls else pd.NA
        )

    return prepared_df


def normalize_source_medias(
    source_url_value: Any,
    media_type_value: Any,
) -> tuple[SourceMedia, ...]:
    """같은 순번의 URL과 media type을 SourceMedia로 묶는다."""

    source_urls = normalize_cell_values(source_url_value)
    raw_media_types = normalize_cell_values(media_type_value)

    if not source_urls:
        return ()

    media_types = tuple(
        identify_media_type(value)
        for value in raw_media_types
    )

    if not media_types:
        media_types = (
            MediaType.UNKNOWN,
        ) * len(source_urls)

    elif len(media_types) == 1 and len(source_urls) > 1:
        single_type = media_types[0]

        # carousel은 게시글 전체 타입이므로 개별 asset 타입으로 쓰지 않음
        if single_type == MediaType.CAROUSEL:
            media_types = (
                MediaType.UNKNOWN,
            ) * len(source_urls)
        else:
            media_types = media_types * len(source_urls)

    elif len(media_types) != len(source_urls):
        raise ValueError(
            "Media URL 개수와 Media Type 개수가 일치하지 않습니다. "
            f"URL={len(source_urls)}, Type={len(media_types)}"
        )

    return tuple(
        SourceMedia(
            source_url=source_url,
            media_type=media_type,
        )
        for source_url, media_type in zip(
            source_urls,
            media_types,
            strict=True,
        )
    )


def identify_post_media_type(
    source_medias: tuple[SourceMedia, ...],
    raw_media_type_value: Any,
) -> MediaType:
    if len(source_medias) > 1:
        return MediaType.CAROUSEL

    if len(source_medias) == 1:
        return source_medias[0].media_type

    raw_types = normalize_cell_values(raw_media_type_value)
    if len(raw_types) == 1:
        return identify_media_type(raw_types[0])

    return MediaType.UNKNOWN


def extract_rows_to_inputs(
    df: pd.DataFrame,
) -> list[StructuredMediaInput]:
    media_inputs: list[StructuredMediaInput] = []

    for row in df.to_dict(orient="records"):
        campaign_id = optional_text(row.get("campaign_id"))

        if campaign_id is None:
            raise ValueError("campaign_id가 없는 행이 존재합니다.")

        raw_row_number_value = row.get("raw_row_number")
        if (
            raw_row_number_value is not None
            and not pd.isna(raw_row_number_value)
        ):
            raw_row_number = int(raw_row_number_value)
        else:
            raw_row_number = None

        try:
            source_medias = normalize_source_medias(
                source_url_value=row.get("source_url"),
                media_type_value=row.get("media_type"),
            )
        except ValueError as exc:
            raise ValueError(
                "Source media 매핑 실패: "
                f"sheet={row.get('source_sheet')}, "
                f"row={raw_row_number}, "
                f"campaign_id={campaign_id}. {exc}"
            ) from exc

        media_input = StructuredMediaInput(
            campaign_id=campaign_id,
            platform=identify_platform(row.get("platform")),
            media_type=identify_post_media_type(
                source_medias=source_medias,
                raw_media_type_value=row.get("media_type"),
            ),
            original_post_url=optional_text(row.get("permalink")),
            conversation_stream=optional_text(
                row.get("conversation_stream")
            ),
            account=AccountInfo(
                account_name=optional_text(row.get("user_name")),
                account_handle=optional_text(
                    row.get("sender_screen_name")
                ),
                follower_count=optional_follower_count(
                    row.get("sender_follower_count")
                ),
                profile_url=optional_text(row.get("profile_url")),
                location=optional_text(row.get("sender_location")),
                detailed_location=optional_text(
                    row.get("sender_detailed_location")
                ),
                bio=optional_text(row.get("sender_bio")),
                website=optional_text(row.get("sender_website")),
                verified=optional_bool(row.get("sender_verified")),
                verified_type=optional_text(
                    row.get("sender_verified_type")
                ),
                profile_tags=optional_text(
                    row.get("sender_profile_tags")
                ),
            ),
            source_medias=source_medias,
            source_sheet_name=optional_text(row.get("source_sheet")),
            raw_row_number=raw_row_number,
        )
        media_inputs.append(media_input)

    return media_inputs


# =========================================================
# 6. Structured Input -> MediaAsset
# =========================================================

def build_media_assets(
    media_inputs: list[StructuredMediaInput],
) -> list[MediaAsset]:
    media_assets: list[MediaAsset] = []

    for media_input in media_inputs:
        if not media_input.source_medias:
            media_assets.append(
                MediaAsset(
                    campaign_id=media_input.campaign_id,
                    asset_id=f"{media_input.campaign_id}_01",
                    asset_index=1,
                    platform=media_input.platform,
                    media_type=media_input.media_type,
                    original_post_url=media_input.original_post_url,
                    source_url=None,
                    input_media_type=media_input.media_type,
                    source_sheet_name=media_input.source_sheet_name,
                    raw_row_number=media_input.raw_row_number,
                    status="source_url_missing",
                    error_message="source URL이 존재하지 않습니다.",
                )
            )
            continue

        for asset_index, source_media in enumerate(
            media_input.source_medias,
            start=1,
        ):
            media_assets.append(
                MediaAsset(
                    campaign_id=media_input.campaign_id,
                    asset_id=(
                        f"{media_input.campaign_id}_"
                        f"{asset_index:02d}"
                    ),
                    asset_index=asset_index,
                    platform=media_input.platform,
                    media_type=source_media.media_type,
                    original_post_url=media_input.original_post_url,
                    source_url=source_media.source_url,
                    input_media_type=media_input.media_type,
                    source_sheet_name=media_input.source_sheet_name,
                    raw_row_number=media_input.raw_row_number,
                )
            )

    return media_assets


# =========================================================
# 6-1. gallery-dl Media URL Resolution
# =========================================================

def is_http_url(value: str | None) -> bool:
    if not value:
        return False

    try:
        parsed_url = urlparse(value.strip())
    except (TypeError, ValueError):
        return False

    return (
        parsed_url.scheme.lower() in {"http", "https"}
        and bool(parsed_url.netloc)
    )


def is_twitter_status_url(
    source_url: str | None,
) -> bool:
    """Twitter/X의 개별 status URL인지 확인한다."""

    if not is_http_url(source_url):
        return False

    parsed_url = urlparse(source_url.strip())
    hostname = (parsed_url.hostname or "").lower()

    if hostname not in {
        "x.com",
        "www.x.com",
        "twitter.com",
        "www.twitter.com",
        "mobile.twitter.com",
    }:
        return False

    return bool(
        re.search(
            r"/status/\d+",
            parsed_url.path,
        )
    )


def is_tiktok_post_url(
    source_url: str | None,
) -> bool:
    """TikTok 개별 게시물 또는 단축 URL인지 확인한다.

    지원 예:
        https://www.tiktok.com/@username/video/1234567890
        https://tiktok.com/@username/video/1234567890
        https://m.tiktok.com/@username/video/1234567890
        https://vm.tiktok.com/...
        https://vt.tiktok.com/...
    """

    if not is_http_url(source_url):
        return False

    parsed_url = urlparse(source_url.strip())
    hostname = (parsed_url.hostname or "").lower()

    if hostname in {
        "vm.tiktok.com",
        "vt.tiktok.com",
    }:
        return True

    if hostname not in {
        "tiktok.com",
        "www.tiktok.com",
        "m.tiktok.com",
    }:
        return False

    return bool(
        re.search(
            r"/@[^/]+/video/\d+",
            parsed_url.path,
            flags=re.IGNORECASE,
        )
    )


def twitter_media_identity_key(media_url: str) -> str:
    """X가 같은 미디어를 여러 해상도 URL로 반환할 때 동일 자산을 식별한다.

    - pbs.twimg.com/media/... 이미지 URL은 name=small/large/orig 차이를 무시한다.
    - video.twimg.com URL은 해상도 경로만 다른 동일 파일을 하나로 묶는다.
    - 식별 규칙을 확신할 수 없는 URL은 전체 URL을 key로 사용한다.
    """

    parsed = urlparse(media_url)
    hostname = (parsed.hostname or "").lower()
    path = parsed.path

    if hostname == "pbs.twimg.com" and path.startswith("/media/"):
        query = parse_qs(parsed.query)
        path_object = Path(path)
        path_suffix = path_object.suffix.lower()
        normalized_path = path

        if path_suffix in {".jpg", ".jpeg", ".png", ".webp", ".gif"}:
            normalized_path = str(path_object.with_suffix(""))

        format_value = (
            query.get("format", [None])[0]
            or path_suffix.lstrip(".")
            or "unknown"
        )
        return (
            f"twitter_image:{hostname}{normalized_path.lower()}:"
            f"{format_value.lower()}"
        )

    if hostname == "video.twimg.com":
        # 같은 영상의 320x180 / 640x360 / 1280x720 경로 차이를 제거한다.
        normalized_path = re.sub(
            r"/(?:avc1/)?\d+x\d+/",
            "/{resolution}/",
            path,
            flags=re.IGNORECASE,
        )
        return f"twitter_video:{hostname}{normalized_path.lower()}"

    return f"exact:{media_url}"


def deduplicate_gallery_dl_urls(
    page_url: str,
    extracted_urls: tuple[str, ...],
) -> tuple[str, ...]:
    """gallery-dl 결과를 순서 보존하며 중복 제거한다.

    Twitter/X는 동일 이미지의 여러 사이즈 URL만 하나로 줄이고,
    서로 다른 이미지 또는 영상은 모두 유지한다.
    TikTok과 다른 플랫폼은 완전히 같은 URL만 제거한다.
    """

    deduplicated_urls: list[str] = []
    seen_keys: set[str] = set()

    for media_url in extracted_urls:
        if is_twitter_status_url(page_url):
            identity_key = twitter_media_identity_key(media_url)
        else:
            identity_key = f"exact:{media_url}"

        if identity_key in seen_keys:
            continue

        seen_keys.add(identity_key)
        deduplicated_urls.append(media_url)

    return tuple(deduplicated_urls)


def extract_media_urls_with_gallery_dl(
    page_url: str,
    timeout_seconds: int = 180,
) -> tuple[str, ...]:
    """게시물 URL에서 실제 미디어 URL들을 추출한다.

    Twitter/X에서는 리포스트와 인용 게시물의 원본 미디어를 조회한다.
    TikTok 게시물 URL은 gallery-dl로 실제 CDN 미디어 URL을 조회한다.
    같은 이미지의 해상도별 중복 URL은 첫 번째 URL만 남기지만,
    게시물에 포함된 서로 다른 미디어는 모두 반환한다.
    """

    if not is_http_url(page_url):
        raise ValueError(
            "gallery-dl에는 유효한 HTTP(S) URL만 전달할 수 있습니다. "
            f"page_url={page_url}"
        )

    command = [
        sys.executable,
        "-m",
        "gallery_dl",
        "-G",
        "--config-ignore",
        "--no-input",
        "--no-colors",
    ]

    if is_twitter_status_url(page_url):
        command.extend(
            [
                "-o",
                "extractor.twitter.retweets=original",
                "-o",
                "extractor.twitter.quoted=true",
            ]
        )

        if TWITTER_INCLUDE_CARD_IMAGES:
            command.extend(
                [
                    "-o",
                    "extractor.twitter.cards=true",
                ]
            )

        if TWITTER_USE_BROWSER_COOKIES:
            command.extend(
                [
                    "--cookies-from-browser",
                    TWITTER_BROWSER,
                ]
            )

    elif is_tiktok_post_url(page_url):
        if TIKTOK_USE_BROWSER_COOKIES:
            command.extend(
                [
                    "--cookies-from-browser",
                    TIKTOK_BROWSER,
                ]
            )

    command.append(page_url)

    try:
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout_seconds,
            check=False,
        )
    except subprocess.TimeoutExpired as exc:
        raise RuntimeError(
            "gallery-dl 실행 시간이 초과되었습니다. "
            f"page_url={page_url}"
        ) from exc
    except OSError as exc:
        raise RuntimeError(
            "gallery-dl을 실행하지 못했습니다. 현재 Python 환경에서 "
            "python -m gallery_dl --version이 실행되는지 확인하세요."
        ) from exc

    raw_urls = tuple(
        dict.fromkeys(
            line.strip()
            for line in result.stdout.splitlines()
            if is_http_url(line.strip())
        )
    )
    extracted_urls = deduplicate_gallery_dl_urls(
        page_url=page_url,
        extracted_urls=raw_urls,
    )

    if extracted_urls:
        if len(raw_urls) != len(extracted_urls):
            print(
                "  gallery-dl 동일 미디어 해상도 중복 제거: "
                f"{len(raw_urls)}개 -> {len(extracted_urls)}개"
            )
        return extracted_urls

    stderr_text = result.stderr.strip()

    if result.returncode != 0:
        raise RuntimeError(
            "gallery-dl 미디어 URL 추출에 실패했습니다. "
            f"return_code={result.returncode}, page_url={page_url}"
        )

    # returncode=0인데 실제 미디어 URL이 하나도 없으면
    # gallery-dl 자체 오류와 구분할 수 있도록 전용 예외를 사용한다.
    raise GalleryDlNoNativeMediaError(
        "gallery-dl 실행은 정상 완료되었지만 실제 미디어 URL을 "
        "찾지 못했습니다. "
        f"page_url={page_url}"
    )


def extract_twitter_media_urls_with_gallery_dl(
    source_url: str,
    timeout_seconds: int = 180,
) -> tuple[str, ...]:
    """기존 호출부 호환용 Twitter 전용 wrapper."""

    if not is_twitter_status_url(source_url):
        raise ValueError(
            "Twitter/X status URL이 아닙니다. "
            f"source_url={source_url}"
        )

    return extract_media_urls_with_gallery_dl(
        page_url=source_url,
        timeout_seconds=timeout_seconds,
    )


def reindex_media_assets(
    media_assets: list[MediaAsset],
) -> list[MediaAsset]:
    """gallery-dl 확장 이후 같은 Excel 행 안에서 asset 순번을 다시 부여한다."""

    next_index_by_row: dict[tuple[str | None, int | None, str], int] = {}

    for asset in media_assets:
        row_key = (
            asset.source_sheet_name,
            asset.raw_row_number,
            asset.campaign_id,
        )
        next_index = next_index_by_row.get(row_key, 1)
        asset.asset_index = next_index
        asset.asset_id = f"{asset.campaign_id}_{next_index:02d}"
        next_index_by_row[row_key] = next_index + 1

    return media_assets


def get_gallery_dl_page_url(asset: MediaAsset) -> str | None:
    """gallery-dl에는 원본 게시물 URL을 우선 전달한다."""

    if is_http_url(asset.original_post_url):
        return asset.original_post_url

    if is_http_url(asset.source_url):
        return asset.source_url

    return None


def resolve_page_media_assets_with_gallery_dl(
    media_assets: list[MediaAsset],
    *,
    only_failed_unknown: bool = False,
) -> list[MediaAsset]:
    """플랫폼 게시물 URL을 원본 게시물 URL로 gallery-dl 해석한다.

    첫 호출에서는 Twitter LINK/UNKNOWN 게시물을 처리한다.
    TikTok은 별도의 yt-dlp 직접 다운로드 함수에서 처리한다.

    두 번째 호출에서는 TikTok을 제외하고 직접 URL 처리에 실패한
    UNKNOWN 행을 처리한다. final_url, parent_asset_id, resolver_url은
    사용하지 않는다.
    """

    resolved_assets: list[MediaAsset] = []

    for asset in media_assets:
        if asset.platform == Platform.YOUTUBE:
            resolved_assets.append(asset)
            continue

        page_url_candidate = get_gallery_dl_page_url(asset)

        if only_failed_unknown:
            should_resolve = (
                asset.platform != Platform.TIKTOK
                and asset.input_media_type == MediaType.UNKNOWN
                and asset.status in {
                    "platform_extractor_required",
                    "feasibility_failed",
                    "source_url_missing",
                }
            )
        else:
            should_resolve = (
                asset.platform == Platform.TWITTER
                and asset.input_media_type in {
                    MediaType.LINK,
                    MediaType.UNKNOWN,
                }
                and asset.status in {"pending", "source_url_missing"}
            )

        if not should_resolve:
            resolved_assets.append(asset)
            continue

        page_url = page_url_candidate
        if not page_url:
            asset.status = "manual_action_required"
            asset.extraction_method = "gallery_dl_original_post_url"
            asset.error_message = (
                "gallery-dl에 전달할 원본 게시물 URL이 없습니다."
            )
            resolved_assets.append(asset)
            continue

        try:
            extracted_urls = extract_media_urls_with_gallery_dl(page_url)

            print(
                "  original_post_url gallery-dl 추출 완료: "
                f"{asset.asset_id} -> {len(extracted_urls)}개"
            )

            for extracted_url in extracted_urls:
                resolved_assets.append(
                    MediaAsset(
                        campaign_id=asset.campaign_id,
                        asset_id=asset.asset_id,
                        asset_index=asset.asset_index,
                        platform=asset.platform,
                        media_type=MediaType.UNKNOWN,
                        original_post_url=asset.original_post_url or page_url,
                        source_url=extracted_url,
                        input_media_type=asset.input_media_type,
                        source_sheet_name=asset.source_sheet_name,
                        raw_row_number=asset.raw_row_number,
                        extraction_method="gallery_dl_original_post_url",
                    )
                )

        except GalleryDlNoNativeMediaError as exc:
            if asset.platform == Platform.TWITTER:
                # Twitter/X 게시물은 정상적으로 확인했지만 실제 이미지/영상이
                # 하나도 없는 경우 text-only 게시물로 확정한다.
                asset.media_type = MediaType.NONE
                asset.source_url = None
                asset.local_path = None
                asset.extraction_method = "gallery_dl_no_native_media"
                asset.status = "no_native_media"
                asset.error_message = None
                resolved_assets.append(asset)

                print(
                    "  Twitter/X 네이티브 미디어 없음 확인: "
                    f"{asset.asset_id}"
                )
            else:
                # Twitter 외 플랫폼은 기존 정책을 그대로 유지한다.
                asset.extraction_method = "gallery_dl_original_post_url"
                asset.status = "manual_action_required"
                asset.error_message = f"{type(exc).__name__}: {exc}"
                resolved_assets.append(asset)

        except Exception as exc:
            asset.extraction_method = "gallery_dl_original_post_url"
            asset.status = "manual_action_required"
            asset.error_message = f"{type(exc).__name__}: {exc}"
            resolved_assets.append(asset)

    return reindex_media_assets(resolved_assets)


def mark_remaining_failures_for_manual_action(
    media_assets: list[MediaAsset],
) -> list[MediaAsset]:
    """자동 처리가 끝난 뒤 미해결 행을 사용자 수동 처리 대상으로 표시한다."""

    failure_statuses = {
        "gallery_dl_failed",
        "platform_extractor_required",
        "feasibility_failed",
        "download_failed",
        "source_url_missing",
    }

    for asset in media_assets:
        if asset.status in failure_statuses:
            asset.status = "manual_action_required"

    return media_assets


# =========================================================
# 7. URL Feasibility
# =========================================================

def detect_media_from_response(
    content_type: str | None,
    sample: bytes,
) -> tuple[MediaType, str | None]:
    normalized_content_type = (
        content_type or ""
    ).split(";", 1)[0].strip().lower()

    content_type_mapping = {
        "video/mp4": (MediaType.VIDEO, ".mp4"),
        "video/webm": (MediaType.VIDEO, ".webm"),
        "video/quicktime": (MediaType.VIDEO, ".mov"),
        "image/jpeg": (MediaType.IMAGE, ".jpg"),
        "image/png": (MediaType.IMAGE, ".png"),
        "image/webp": (MediaType.IMAGE, ".webp"),
        "image/gif": (MediaType.IMAGE, ".gif"),
    }

    if normalized_content_type in content_type_mapping:
        return content_type_mapping[normalized_content_type]

    # 서버 Content-Type이 부정확할 수 있으므로 magic bytes도 확인
    if sample.startswith(b"\xff\xd8\xff"):
        return MediaType.IMAGE, ".jpg"

    if sample.startswith(b"\x89PNG\r\n\x1a\n"):
        return MediaType.IMAGE, ".png"

    if sample.startswith((b"GIF87a", b"GIF89a")):
        return MediaType.IMAGE, ".gif"

    if (
        len(sample) >= 12
        and sample[:4] == b"RIFF"
        and sample[8:12] == b"WEBP"
    ):
        return MediaType.IMAGE, ".webp"

    # MP4/M4V/MOV 계열의 대표 signature
    if len(sample) >= 12 and sample[4:8] == b"ftyp":
        return MediaType.VIDEO, ".mp4"

    if normalized_content_type.startswith("video/"):
        return MediaType.VIDEO, ".mp4"

    if normalized_content_type.startswith("image/"):
        return MediaType.IMAGE, ".jpg"

    if normalized_content_type in {
        "text/html",
        "application/xhtml+xml",
    }:
        return MediaType.LINK, None

    return MediaType.UNKNOWN, None


def read_probe_sample(
    response: requests.Response,
    max_bytes: int = PROBE_BYTES,
) -> bytes:
    chunks: list[bytes] = []
    total_size = 0

    for chunk in response.iter_content(chunk_size=8192):
        if not chunk:
            continue

        remaining = max_bytes - total_size
        chunks.append(chunk[:remaining])
        total_size += min(len(chunk), remaining)

        if total_size >= max_bytes:
            break

    return b"".join(chunks)


def build_media_request_headers(
    asset: MediaAsset,
    *,
    include_range: bool,
) -> dict[str, str]:
    """플랫폼별 직접 미디어 요청 헤더를 생성한다.

    기존 플랫폼은 이전과 동일하게 Range probe를 사용한다.
    TikTok CDN은 Range 요청 또는 Referer 누락 시 403을 반환할 수 있으므로
    TikTok 자산에만 Referer/Origin을 추가하고 probe 단계에서도 Range를 쓰지 않는다.
    """

    headers: dict[str, str] = {}

    if include_range and asset.platform != Platform.TIKTOK:
        headers["Range"] = f"bytes=0-{PROBE_BYTES - 1}"

    if asset.platform == Platform.TIKTOK:
        referer_url = (
            asset.original_post_url
            if is_http_url(asset.original_post_url)
            else "https://www.tiktok.com/"
        )
        headers.update(
            {
                "Referer": referer_url,
                "Origin": "https://www.tiktok.com",
                "Accept-Language": "en-US,en;q=0.9",
            }
        )

    return headers


def check_media_url(
    asset: MediaAsset,
    session: requests.Session,
) -> UrlFeasibilityResult:
    if not asset.source_url:
        return UrlFeasibilityResult(
            feasible=False,
            http_status=None,
            content_type=None,
            detected_media_type=MediaType.UNKNOWN,
            extension=None,
            error_message="검사할 source URL이 없습니다.",
        )

    source_url = asset.source_url
    headers = build_media_request_headers(
        asset,
        include_range=True,
    )

    try:
        with session.get(
            source_url,
            headers=headers,
            stream=True,
            allow_redirects=True,
            timeout=REQUEST_TIMEOUT,
        ) as response:
            sample = read_probe_sample(response)
            content_type = response.headers.get("Content-Type")
            detected_media_type, extension = detect_media_from_response(
                content_type=content_type,
                sample=sample,
            )

            if response.status_code not in {200, 206}:
                return UrlFeasibilityResult(
                    feasible=False,
                    http_status=response.status_code,
                    content_type=content_type,
                    detected_media_type=detected_media_type,
                    extension=None,
                    error_message=(
                        "정상적인 미디어 응답이 아닙니다: "
                        f"HTTP {response.status_code}"
                    ),
                )

            if detected_media_type == MediaType.LINK:
                return UrlFeasibilityResult(
                    feasible=False,
                    http_status=response.status_code,
                    content_type=content_type,
                    detected_media_type=MediaType.LINK,
                    extension=None,
                    error_message=(
                        "HTML 페이지 URL입니다. 직접 파일 다운로드가 아니라 "
                        "플랫폼 extractor가 필요합니다."
                    ),
                )

            if detected_media_type not in {
                MediaType.VIDEO,
                MediaType.IMAGE,
            }:
                return UrlFeasibilityResult(
                    feasible=False,
                    http_status=response.status_code,
                    content_type=content_type,
                    detected_media_type=MediaType.UNKNOWN,
                    extension=None,
                    error_message=(
                        "URL에는 접근했지만 이미지 또는 영상 파일로 "
                        "판별하지 못했습니다."
                    ),
                )

            return UrlFeasibilityResult(
                feasible=True,
                http_status=response.status_code,
                content_type=content_type,
                detected_media_type=detected_media_type,
                extension=extension,
            )

    except requests.exceptions.Timeout:
        return UrlFeasibilityResult(
            feasible=False,
            http_status=None,
            content_type=None,
            detected_media_type=MediaType.UNKNOWN,
            extension=None,
            error_message="서버 응답 시간이 초과되었습니다.",
        )

    except requests.exceptions.RequestException as exc:
        return UrlFeasibilityResult(
            feasible=False,
            http_status=None,
            content_type=None,
            detected_media_type=MediaType.UNKNOWN,
            extension=None,
            error_message=f"URL 요청 실패: {type(exc).__name__}: {exc}",
        )


# =========================================================
# 8. Direct Media Download
# =========================================================

def sanitize_filename(value: str) -> str:
    sanitized = re.sub(
        r'[<>:"/\\|?*\x00-\x1f]',
        "_",
        value,
    )
    sanitized = sanitized.strip(" .")
    return sanitized or "unknown"


def build_media_filename_stem(
    asset: MediaAsset,
) -> str:
    """
    저장 파일명 규칙:
        첫 번째 asset: {sheet_name}_{excel_row}
        두 번째 이후: {sheet_name}_{excel_row}_{asset_index:02d}

    한 Excel 행에 미디어가 여러 개 있을 때 파일명 충돌을 방지한다.
    """

    safe_sheet_name = sanitize_filename(
        asset.source_sheet_name or "unknown_sheet"
    )

    if asset.raw_row_number is None:
        row_text = "unknown_row"
    else:
        row_text = str(asset.raw_row_number)

    if asset.asset_index == 1:
        return f"{safe_sheet_name}_{row_text}"

    return (
        f"{safe_sheet_name}_{row_text}_"
        f"{asset.asset_index:02d}"
    )



TIKTOK_LOCAL_VIDEO_EXTENSIONS = {
    ".mp4",
    ".webm",
    ".mkv",
    ".mov",
    ".m4v",
}

TIKTOK_LOCAL_IMAGE_EXTENSIONS = {
    ".jpg",
    ".jpeg",
    ".png",
    ".webp",
    ".gif",
}

TIKTOK_YT_DLP_AUXILIARY_EXTENSIONS = {
    ".part",
    ".ytdl",
    ".json",
    ".description",
    ".vtt",
    ".srt",
    ".ass",
    ".lrc",
    ".txt",
}


def get_tiktok_page_url(asset: MediaAsset) -> str | None:
    """yt-dlp에 전달할 TikTok 게시물 URL을 반환한다."""

    for candidate_url in (
        asset.original_post_url,
        asset.source_url,
    ):
        if is_tiktok_post_url(candidate_url):
            return candidate_url.strip()

    return None


def detect_local_media_file(
    media_path: Path,
) -> tuple[MediaType, str | None]:
    """yt-dlp가 저장한 로컬 파일의 미디어 타입을 판별한다."""

    suffix = media_path.suffix.lower()

    if suffix in TIKTOK_LOCAL_VIDEO_EXTENSIONS:
        content_type_mapping = {
            ".mp4": "video/mp4",
            ".webm": "video/webm",
            ".mov": "video/quicktime",
            ".m4v": "video/x-m4v",
            ".mkv": "video/x-matroska",
        }
        return MediaType.VIDEO, content_type_mapping.get(
            suffix,
            "video/*",
        )

    if suffix in TIKTOK_LOCAL_IMAGE_EXTENSIONS:
        content_type_mapping = {
            ".jpg": "image/jpeg",
            ".jpeg": "image/jpeg",
            ".png": "image/png",
            ".webp": "image/webp",
            ".gif": "image/gif",
        }
        return MediaType.IMAGE, content_type_mapping.get(
            suffix,
            "image/*",
        )

    with media_path.open("rb") as file:
        sample = file.read(PROBE_BYTES)

    detected_media_type, _ = detect_media_from_response(
        content_type=None,
        sample=sample,
    )

    if detected_media_type == MediaType.VIDEO:
        return detected_media_type, "video/*"

    if detected_media_type == MediaType.IMAGE:
        return detected_media_type, "image/*"

    return MediaType.UNKNOWN, None


def collect_yt_dlp_downloaded_files(
    temporary_directory: Path,
    stdout_text: str,
) -> list[Path]:
    """yt-dlp 출력과 임시 폴더를 이용해 실제 저장 파일을 찾는다."""

    downloaded_files: list[Path] = []
    seen_paths: set[Path] = set()

    for raw_line in stdout_text.splitlines():
        candidate_text = raw_line.strip().strip('"')
        if not candidate_text:
            continue

        candidate_path = Path(candidate_text)
        if not candidate_path.is_absolute():
            candidate_path = temporary_directory / candidate_path

        try:
            resolved_path = candidate_path.resolve()
        except OSError:
            continue

        if (
            resolved_path.is_file()
            and resolved_path.suffix.lower()
            not in TIKTOK_YT_DLP_AUXILIARY_EXTENSIONS
            and resolved_path not in seen_paths
        ):
            seen_paths.add(resolved_path)
            downloaded_files.append(resolved_path)

    for candidate_path in sorted(
        temporary_directory.rglob("*"),
        key=lambda path: str(path).casefold(),
    ):
        if not candidate_path.is_file():
            continue

        if (
            candidate_path.suffix.lower()
            in TIKTOK_YT_DLP_AUXILIARY_EXTENSIONS
        ):
            continue

        resolved_path = candidate_path.resolve()
        if resolved_path in seen_paths:
            continue

        seen_paths.add(resolved_path)
        downloaded_files.append(resolved_path)

    return downloaded_files


def build_tiktok_yt_dlp_command(
    page_url: str,
    temporary_directory: Path,
    *,
    use_impersonation: bool,
) -> list[str]:
    """현재 Python 환경의 yt-dlp를 호출하는 명령어를 생성한다."""

    output_template = (
        temporary_directory
        / "%(id)s_%(autonumber)02d.%(ext)s"
    )

    command = [
        sys.executable,
        "-m",
        "yt_dlp",
        "--no-playlist",
        "--no-progress",
        "--newline",
        "--no-colors",
        "--force-overwrites",
        "--format",
        "best[ext=mp4]/best",
        "--output",
        str(output_template),
        "--print",
        "after_move:filepath",
    ]

    if use_impersonation and TIKTOK_YT_DLP_IMPERSONATE:
        command.extend(
            [
                "--impersonate",
                TIKTOK_YT_DLP_IMPERSONATE,
            ]
        )

    if TIKTOK_USE_BROWSER_COOKIES:
        command.extend(
            [
                "--cookies-from-browser",
                TIKTOK_BROWSER,
            ]
        )

    command.append(page_url)
    return command


def run_tiktok_yt_dlp(
    page_url: str,
    temporary_directory: Path,
) -> subprocess.CompletedProcess[str]:
    """Chrome impersonation을 우선 사용하고 필요하면 일반 요청으로 재시도한다."""

    command = build_tiktok_yt_dlp_command(
        page_url=page_url,
        temporary_directory=temporary_directory,
        use_impersonation=True,
    )

    try:
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=TIKTOK_YT_DLP_TIMEOUT_SECONDS,
            check=False,
        )
    except subprocess.TimeoutExpired as exc:
        raise RuntimeError(
            "yt-dlp TikTok 다운로드 시간이 초과되었습니다. "
            f"page_url={page_url}"
        ) from exc
    except OSError as exc:
        raise RuntimeError(
            "yt-dlp를 실행하지 못했습니다. 현재 Python 환경에서 "
            "python -m yt_dlp --version이 실행되는지 확인하세요."
        ) from exc

    if result.returncode == 0:
        return result

    combined_error = (
        f"{result.stderr}\n{result.stdout}"
    ).casefold()
    impersonation_unavailable_markers = (
        "impersonate target",
        "impersonation target",
        "curl_cffi",
        "curl-cffi",
        "unsupported impersonate",
        "not available for impersonation",
    )

    should_retry_without_impersonation = (
        bool(TIKTOK_YT_DLP_IMPERSONATE)
        and any(
            marker in combined_error
            for marker in impersonation_unavailable_markers
        )
    )

    if not should_retry_without_impersonation:
        return result

    print(
        "  TikTok yt-dlp Chrome impersonation을 사용할 수 없어 "
        "일반 요청으로 한 번 더 시도합니다."
    )

    fallback_command = build_tiktok_yt_dlp_command(
        page_url=page_url,
        temporary_directory=temporary_directory,
        use_impersonation=False,
    )

    try:
        return subprocess.run(
            fallback_command,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=TIKTOK_YT_DLP_TIMEOUT_SECONDS,
            check=False,
        )
    except subprocess.TimeoutExpired as exc:
        raise RuntimeError(
            "yt-dlp TikTok 일반 요청 재시도 시간이 초과되었습니다. "
            f"page_url={page_url}"
        ) from exc


def download_tiktok_post_with_yt_dlp(
    asset: MediaAsset,
    local_media_root: Path,
) -> list[MediaAsset]:
    """TikTok 게시물 URL을 yt-dlp로 내려받아 MediaAsset으로 변환한다."""

    page_url = get_tiktok_page_url(asset)

    if page_url is None:
        # TikTok CDN 직접 URL이면 기존 URL feasibility/HTTP 다운로드 흐름을 유지한다.
        return [asset]

    safe_sheet_name = sanitize_filename(
        asset.source_sheet_name or "unknown_sheet"
    )
    sheet_directory = local_media_root / safe_sheet_name
    sheet_directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    row_token = sanitize_filename(
        f"{asset.campaign_id}_{asset.raw_row_number or 'unknown_row'}"
    )
    temporary_directory = (
        sheet_directory
        / f".tiktok_yt_dlp_{row_token}"
    )

    if temporary_directory.exists():
        shutil.rmtree(temporary_directory)

    temporary_directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    try:
        result = run_tiktok_yt_dlp(
            page_url=page_url,
            temporary_directory=temporary_directory,
        )

        downloaded_files = collect_yt_dlp_downloaded_files(
            temporary_directory=temporary_directory,
            stdout_text=result.stdout,
        )

        if not downloaded_files:
            error_text = (
                result.stderr.strip()
                or result.stdout.strip()
                or "상세 오류 메시지가 없습니다."
            )
            asset.status = "manual_action_required"
            asset.extraction_method = "yt_dlp_tiktok"
            asset.error_message = (
                "yt-dlp TikTok 미디어 다운로드에 실패했습니다. "
                f"return_code={result.returncode}, "
                f"page_url={page_url}, "
                f"message={error_text[:2000]}"
            )
            return [asset]

        downloaded_assets: list[MediaAsset] = []

        for downloaded_index, downloaded_path in enumerate(
            downloaded_files,
            start=1,
        ):
            if downloaded_path.stat().st_size <= 0:
                continue

            media_type, content_type = detect_local_media_file(
                downloaded_path
            )

            if media_type not in {
                MediaType.VIDEO,
                MediaType.IMAGE,
            }:
                continue

            downloaded_asset = MediaAsset(
                campaign_id=asset.campaign_id,
                asset_id=asset.asset_id,
                asset_index=downloaded_index,
                platform=asset.platform,
                media_type=media_type,
                original_post_url=asset.original_post_url or page_url,
                source_url=page_url,
                input_media_type=asset.input_media_type,
                source_sheet_name=asset.source_sheet_name,
                raw_row_number=asset.raw_row_number,
                content_type=content_type,
                extraction_method="yt_dlp_tiktok",
                status="downloaded",
            )

            file_stem = build_media_filename_stem(
                downloaded_asset
            )
            final_path = (
                sheet_directory
                / f"{file_stem}{downloaded_path.suffix.lower()}"
            )

            if final_path.exists():
                final_path.unlink()

            downloaded_path.replace(final_path)

            downloaded_asset.file_size_bytes = final_path.stat().st_size
            downloaded_asset.local_path = final_path
            downloaded_asset.error_message = None
            downloaded_assets.append(downloaded_asset)

        if not downloaded_assets:
            asset.status = "manual_action_required"
            asset.extraction_method = "yt_dlp_tiktok"
            asset.error_message = (
                "yt-dlp가 파일을 생성했지만 지원 가능한 이미지 또는 "
                f"영상으로 판별하지 못했습니다. page_url={page_url}"
            )
            return [asset]

        print(
            "  TikTok yt-dlp 저장 완료: "
            f"{asset.asset_id} -> {len(downloaded_assets)}개"
        )
        return downloaded_assets

    except Exception as exc:
        asset.status = "manual_action_required"
        asset.extraction_method = "yt_dlp_tiktok"
        asset.error_message = f"{type(exc).__name__}: {exc}"
        return [asset]

    finally:
        if temporary_directory.exists():
            shutil.rmtree(
                temporary_directory,
                ignore_errors=True,
            )


def process_tiktok_media_assets_with_yt_dlp(
    media_assets: list[MediaAsset],
    local_media_root: Path,
) -> list[MediaAsset]:
    """TikTok 게시물만 yt-dlp로 처리하고 다른 플랫폼은 그대로 유지한다."""

    processed_assets: list[MediaAsset] = []
    processed_tiktok_rows: set[
        tuple[str | None, int | None, str]
    ] = set()

    for asset in media_assets:
        if asset.platform != Platform.TIKTOK:
            processed_assets.append(asset)
            continue

        page_url = get_tiktok_page_url(asset)
        if page_url is None:
            # 이미 직접 다운로드 가능한 TikTok CDN URL인 경우 기존 로직을 사용한다.
            processed_assets.append(asset)
            continue

        row_key = (
            asset.source_sheet_name,
            asset.raw_row_number,
            asset.campaign_id,
        )

        # Sprinklr가 같은 게시물 URL을 여러 MediaAsset으로 만들더라도
        # 게시물 한 행당 yt-dlp는 한 번만 호출한다.
        if row_key in processed_tiktok_rows:
            continue

        processed_tiktok_rows.add(row_key)
        processed_assets.extend(
            download_tiktok_post_with_yt_dlp(
                asset=asset,
                local_media_root=local_media_root,
            )
        )

    return reindex_media_assets(processed_assets)


def download_media_asset(
    asset: MediaAsset,
    feasibility: UrlFeasibilityResult,
    local_media_root: Path,
    session: requests.Session,
) -> MediaAsset:
    if not feasibility.feasible:
        asset.status = "feasibility_failed"
        asset.error_message = feasibility.error_message
        return asset

    if asset.source_url is None:
        asset.status = "source_url_missing"
        asset.error_message = "다운로드할 source URL이 없습니다."
        return asset

    safe_sheet_name = sanitize_filename(
        asset.source_sheet_name or "unknown_sheet"
    )

    sheet_directory = (
        local_media_root
        / safe_sheet_name
    )
    sheet_directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    try:
        download_headers = build_media_request_headers(
            asset,
            include_range=False,
        )

        with session.get(
            asset.source_url,
            headers=download_headers,
            stream=True,
            allow_redirects=True,
            timeout=REQUEST_TIMEOUT,
        ) as response:
            response.raise_for_status()

            chunk_iterator = response.iter_content(
                chunk_size=1024 * 1024
            )
            first_chunk = next(
                (chunk for chunk in chunk_iterator if chunk),
                b"",
            )

            actual_media_type, actual_extension = detect_media_from_response(
                content_type=response.headers.get("Content-Type"),
                sample=first_chunk,
            )

            if actual_media_type not in {
                MediaType.VIDEO,
                MediaType.IMAGE,
            }:
                raise ValueError(
                    "다운로드 응답이 이미지 또는 영상 파일이 아닙니다. "
                    f"Content-Type={response.headers.get('Content-Type')}"
                )

            extension = actual_extension or feasibility.extension
            if extension is None:
                raise ValueError("파일 확장자를 판별하지 못했습니다.")

            file_stem = build_media_filename_stem(
                asset
            )

            output_path = (
                sheet_directory
                / f"{file_stem}{extension}"
            )
            temporary_path = output_path.with_suffix(
                output_path.suffix + ".part"
            )

            if temporary_path.exists():
                temporary_path.unlink()

            with temporary_path.open("wb") as file:
                if first_chunk:
                    file.write(first_chunk)

                for chunk in chunk_iterator:
                    if chunk:
                        file.write(chunk)

            temporary_path.replace(output_path)

            asset.http_status = response.status_code
            asset.content_type = response.headers.get("Content-Type")
            asset.media_type = actual_media_type
            asset.file_size_bytes = output_path.stat().st_size
            asset.local_path = output_path
            if (
                asset.extraction_method
                and asset.extraction_method.startswith("gallery_dl")
            ):
                asset.extraction_method = (
                    f"{asset.extraction_method} + direct_http_download"
                )
            else:
                asset.extraction_method = "direct_http_download"

            asset.status = "downloaded"
            asset.error_message = None

            return asset

    except Exception as exc:
        asset.status = "download_failed"
        asset.error_message = f"{type(exc).__name__}: {exc}"
        return asset


def process_media_assets(
    media_assets: list[MediaAsset],
    local_media_root: Path,
) -> list[MediaAsset]:
    """
    status가 pending인 asset만 실제 검사 및 다운로드한다.

    이미 downloaded, llm_url_ready, manual_action_required 등으로
    처리된 asset은 출력과 네트워크 요청 없이 그대로 유지한다.

    YouTube는 파일을 다운로드하지 않고 게시물 URL을
    llm_url_ready 상태로 준비한다.
    """

    processed_assets: list[MediaAsset] = []

    # 실제 처리가 필요한 pending asset 개수만 계산
    pending_asset_count = sum(
        1
        for asset in media_assets
        if asset.status == "pending"
    )

    pending_index = 0

    with requests.Session() as session:
        session.headers.update(
            {
                "User-Agent": USER_AGENT,
                "Accept": "*/*",
            }
        )

        for asset in media_assets:
            # =====================================================
            # 이미 처리된 asset은 출력도 하지 않고 그대로 유지
            # =====================================================
            if asset.status != "pending":
                processed_assets.append(asset)
                continue

            pending_index += 1

            print(
                f"[{pending_index}/{pending_asset_count}] 검사: "
                f"{asset.asset_id}"
            )

            # =====================================================
            # YouTube
            # URL 자체를 LLM 입력으로 사용하며 파일 다운로드하지 않음
            # =====================================================
            if asset.platform == Platform.YOUTUBE:
                youtube_url = (
                    asset.original_post_url
                    or asset.source_url
                )

                if youtube_url:
                    asset.source_url = (
                        asset.source_url
                        or youtube_url
                    )
                    asset.media_type = MediaType.VIDEO
                    asset.extraction_method = (
                        "youtube_url_direct_to_llm"
                    )
                    asset.status = "llm_url_ready"
                    asset.error_message = None

                    print(
                        "  YouTube LLM URL 준비 완료: "
                        f"{youtube_url}"
                    )

                else:
                    asset.status = "source_url_missing"
                    asset.error_message = (
                        "YouTube Permalink와 source URL이 "
                        "모두 없습니다."
                    )

                processed_assets.append(asset)
                continue

            # =====================================================
            # source URL 누락
            # =====================================================
            if not asset.source_url:
                asset.status = "source_url_missing"
                asset.error_message = (
                    "source URL이 존재하지 않습니다."
                )
                processed_assets.append(asset)
                continue

            # =====================================================
            # URL feasibility 검사
            # =====================================================
            feasibility = check_media_url(
                asset=asset,
                session=session,
            )

            asset.http_status = feasibility.http_status
            asset.content_type = feasibility.content_type

            if not feasibility.feasible:
                if (
                    feasibility.detected_media_type
                    == MediaType.LINK
                ):
                    asset.status = (
                        "platform_extractor_required"
                    )
                else:
                    asset.status = "feasibility_failed"

                asset.error_message = (
                    feasibility.error_message
                )
                processed_assets.append(asset)

                print(
                    f"  실패: {asset.status} - "
                    f"{asset.error_message}"
                )
                continue

            # Raw Data의 media type보다 실제 응답 타입을 사용
            asset.media_type = (
                feasibility.detected_media_type
            )

            # =====================================================
            # 실제 다운로드
            # =====================================================
            asset = download_media_asset(
                asset=asset,
                feasibility=feasibility,
                local_media_root=local_media_root,
                session=session,
            )
            processed_assets.append(asset)

            if asset.status == "downloaded":
                print(
                    f"  저장 완료: {asset.local_path}"
                )
            else:
                print(
                    "  다운로드 실패: "
                    f"{asset.error_message}"
                )

    return processed_assets


# =========================================================
# 9. Manifest / Excel Output
# =========================================================

def build_llm_manifest_fields(
    asset: MediaAsset,
) -> dict[str, Any]:
    """각 manifest asset의 LLM 입력 상태와 사용자 액션 필요 여부를 표시한다."""

    if (
        asset.platform == Platform.YOUTUBE
        and asset.status == "llm_url_ready"
        and asset.source_url
    ):
        return {
            "llm_input_mode": "youtube_url",
            "llm_input_value": asset.source_url,
            "llm_ready": True,
            "user_action_required": False,
            "user_action": None,
        }

    if asset.status == "downloaded" and asset.local_path is not None:
        return {
            "llm_input_mode": "local_media_file",
            "llm_input_value": str(asset.local_path),
            "llm_ready": True,
            "user_action_required": False,
            "user_action": None,
        }

    if (
        asset.platform == Platform.TWITTER
        and asset.status == "no_native_media"
        and asset.media_type == MediaType.NONE
    ):
        return {
            "llm_input_mode": "post_text_only",
            "llm_input_value": (
                asset.original_post_url
                or asset.source_url
            ),
            "llm_ready": True,
            "user_action_required": False,
            "user_action": None,
        }

    if asset.status == "manual_action_required":
        return {
            "llm_input_mode": "manual_media_upload",
            "llm_input_value": asset.original_post_url or asset.source_url,
            "llm_ready": False,
            "user_action_required": True,
            "user_action": "미디어를 수동 추출한 뒤 LLM에 직접 입력",
        }

    return {
        "llm_input_mode": "unresolved",
        "llm_input_value": None,
        "llm_ready": False,
        "user_action_required": False,
        "user_action": None,
    }


def media_assets_to_dataframe(
    media_assets: list[MediaAsset],
) -> pd.DataFrame:
    records: list[dict[str, Any]] = []

    for asset in media_assets:
        llm_fields = build_llm_manifest_fields(asset)

        records.append(
            {
                "campaign_id": asset.campaign_id,
                "asset_id": asset.asset_id,
                "asset_index": asset.asset_index,
                "platform": asset.platform.value,
                "media_type": asset.media_type.value,
                "source_sheet_name": asset.source_sheet_name,
                "raw_row_number": asset.raw_row_number,
                "original_post_url": asset.original_post_url,
                "source_url": asset.source_url,
                "http_status": asset.http_status,
                "content_type": asset.content_type,
                "file_size_bytes": asset.file_size_bytes,
                "local_path": (
                    str(asset.local_path)
                    if asset.local_path is not None
                    else None
                ),
                "extraction_method": asset.extraction_method,
                "status": asset.status,
                "error_message": asset.error_message,
                **llm_fields,
            }
        )

    return pd.DataFrame(
        records,
        columns=MANIFEST_COLUMNS,
    )


def normalize_row_number(value: Any) -> int | None:
    if value is None:
        return None

    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass

    return int(value)


def make_row_key(
    source_sheet_name: Any,
    raw_row_number: Any,
    campaign_id: Any,
) -> tuple[str | None, int | None, str | None]:
    return (
        optional_text(source_sheet_name),
        normalize_row_number(raw_row_number),
        optional_text(campaign_id),
    )


def join_nonempty(values: list[Any], *, unique: bool = False) -> str | None:
    normalized_values: list[str] = []
    seen: set[str] = set()

    for value in values:
        text = optional_text(value)
        if not text:
            continue

        if unique and text in seen:
            continue

        seen.add(text)
        normalized_values.append(text)

    if not normalized_values:
        return None

    return "\n".join(normalized_values)


def derive_post_media_type(
    successful_rows: list[dict[str, Any]],
    fallback_value: Any,
) -> str:
    successful_types = [
        optional_text(row.get("media_type"))
        for row in successful_rows
    ]
    successful_types = [
        value
        for value in successful_types
        if value in {MediaType.VIDEO.value, MediaType.IMAGE.value}
    ]

    if len(successful_types) > 1:
        return MediaType.CAROUSEL.value

    if len(successful_types) == 1:
        return successful_types[0]

    fallback_types = normalize_cell_values(fallback_value)
    if len(fallback_types) == 1:
        return identify_media_type(fallback_types[0]).value

    return MediaType.UNKNOWN.value


def build_sender_profile_output_fields(
    campaign_row: dict[str, Any],
) -> dict[str, Any]:
    """LLM 게시자 분류에 필요한 Sender Profile 필드를 한 번에 반환한다."""

    return {
        "sender_profile_available": optional_bool(
            campaign_row.get("sender_profile_available")
        ),
        "sender_screen_name": optional_text(
            campaign_row.get("sender_screen_name")
        ),
        "sender_follower_count": optional_follower_count(
            campaign_row.get("sender_follower_count")
        ),
        "sender_location": optional_text(
            campaign_row.get("sender_location")
        ),
        "sender_detailed_location": optional_text(
            campaign_row.get("sender_detailed_location")
        ),
        "sender_bio": optional_text(
            campaign_row.get("sender_bio")
        ),
        "sender_website": optional_text(
            campaign_row.get("sender_website")
        ),
        "sender_verified": optional_bool(
            campaign_row.get("sender_verified")
        ),
        "sender_verified_type": optional_text(
            campaign_row.get("sender_verified_type")
        ),
        "sender_profile_tags": optional_text(
            campaign_row.get("sender_profile_tags")
        ),
    }


def build_llm_input_dataframe(
    campaign_input_df: pd.DataFrame,
    manifest_df: pd.DataFrame,
) -> pd.DataFrame:
    """asset 처리 결과를 post 단위 1행의 LLM 입력 테이블로 집계한다.

    한 게시물에 여러 미디어가 있으면 media_types, media_source_urls,
    local_media_paths에 같은 순서로 줄바꿈 저장한다. LLM 호출 시
    local_media_paths를 splitlines()하여 하나의 요청에 모두 전달한다.
    """

    manifest_rows_by_key: dict[
        tuple[str | None, int | None, str | None],
        list[dict[str, Any]],
    ] = {}

    for manifest_row in manifest_df.to_dict(orient="records"):
        row_key = make_row_key(
            manifest_row.get("source_sheet_name"),
            manifest_row.get("raw_row_number"),
            manifest_row.get("campaign_id"),
        )
        manifest_rows_by_key.setdefault(row_key, []).append(manifest_row)

    records: list[dict[str, Any]] = []

    for campaign_row in campaign_input_df.to_dict(orient="records"):
        row_key = make_row_key(
            campaign_row.get("source_sheet"),
            campaign_row.get("raw_row_number"),
            campaign_row.get("campaign_id"),
        )
        manifest_rows = sorted(
            manifest_rows_by_key.get(row_key, []),
            key=lambda row: int(row.get("asset_index") or 0),
        )
        platform = identify_platform(campaign_row.get("platform"))
        permalink = optional_text(campaign_row.get("permalink"))

        if platform == Platform.YOUTUBE:
            youtube_url = (
                permalink
                or optional_text(campaign_row.get("source_url"))
            )
            ready = bool(youtube_url)

            records.append(
                {
                    "campaign_id": optional_text(campaign_row.get("campaign_id")),
                    "source_sheet": optional_text(campaign_row.get("source_sheet")),
                    "raw_row_number": normalize_row_number(
                        campaign_row.get("raw_row_number")
                    ),
                    "platform": Platform.YOUTUBE.value,
                    "permalink": permalink,
                    "conversation_stream": optional_text(
                        campaign_row.get("conversation_stream")
                    ),
                    "user_name": optional_text(campaign_row.get("user_name")),
                    "profile_url": optional_text(campaign_row.get("profile_url")),
                    **build_sender_profile_output_fields(campaign_row),
                    "post_media_type": MediaType.VIDEO.value,
                    "media_count": 1 if ready else 0,
                    "media_types": MediaType.VIDEO.value if ready else None,
                    "media_source_urls": youtube_url,
                    "local_media_paths": None,
                    "successful_media_count": 1 if ready else 0,
                    "failed_media_count": 0 if ready else 1,
                    "extraction_methods": (
                        "youtube_url_direct_to_llm" if ready else None
                    ),
                    "status": (
                        "llm_url_ready" if ready else "manual_action_required"
                    ),
                    "error_message": (
                        None
                        if ready
                        else "YouTube Permalink와 Media URL이 모두 없습니다."
                    ),
                    "llm_input_mode": (
                        "youtube_url" if ready else "manual_media_upload"
                    ),
                    "llm_input_value": youtube_url,
                    "llm_ready": ready,
                    "user_action_required": not ready,
                    "user_action": (
                        None
                        if ready
                        else "게시물 URL을 확인하고 미디어를 수동 확보한 뒤 LLM에 입력"
                    ),
                }
            )
            continue

        successful_rows = [
            row
            for row in manifest_rows
            if row.get("status") == "downloaded"
            and optional_text(row.get("local_path"))
        ]
        no_native_media_rows = [
            row
            for row in manifest_rows
            if (
                row.get("status") == "no_native_media"
                and row.get("media_type") == MediaType.NONE.value
            )
        ]
        failed_rows = [
            row
            for row in manifest_rows
            if row.get("status")
            not in {
                "downloaded",
                "no_native_media",
            }
        ]

        successful_count = len(successful_rows)
        failed_count = len(failed_rows)
        total_count = len(manifest_rows)

        media_types = join_nonempty(
            [row.get("media_type") for row in successful_rows]
        )
        media_source_urls = join_nonempty(
            [row.get("source_url") for row in successful_rows]
        )
        local_media_paths = join_nonempty(
            [row.get("local_path") for row in successful_rows]
        )
        extraction_methods = join_nonempty(
            [row.get("extraction_method") for row in manifest_rows],
            unique=True,
        )
        error_message = join_nonempty(
            [row.get("error_message") for row in failed_rows],
            unique=True,
        )

        post_media_type = derive_post_media_type(
            successful_rows=successful_rows,
            fallback_value=campaign_row.get("media_type"),
        )
        output_media_count = total_count

        if successful_count > 0 and failed_count == 0:
            status = "ready"
            llm_input_mode = "local_media_files"
            llm_input_value = local_media_paths
            llm_ready = True
            user_action_required = False
            user_action = None
        elif successful_count > 0 and failed_count > 0:
            status = "partial_media_manual_action_required"
            llm_input_mode = "local_media_files_plus_manual_upload"
            llm_input_value = local_media_paths
            llm_ready = False
            user_action_required = True
            user_action = (
                "일부 미디어 자동 추출에 실패했습니다. 실패한 미디어를 수동 저장한 뒤 "
                "local_media_paths의 파일들과 함께 하나의 LLM 입력으로 전달하세요."
            )
        elif (
            platform == Platform.TWITTER
            and no_native_media_rows
            and successful_count == 0
            and failed_count == 0
        ):
            # gallery-dl이 게시물을 정상 확인했지만 네이티브 이미지/영상이
            # 하나도 없는 Twitter/X 게시물은 정상적인 text-only 입력이다.
            post_media_type = MediaType.NONE.value
            output_media_count = 0
            status = "ready_no_native_media"
            llm_input_mode = "post_text_only"
            llm_input_value = permalink
            llm_ready = True
            user_action_required = False
            user_action = None
        else:
            status = "manual_action_required"
            llm_input_mode = "manual_media_upload"
            llm_input_value = permalink
            llm_ready = False
            user_action_required = True
            user_action = (
                "게시물의 미디어를 수동 저장한 뒤 게시물의 모든 미디어를 "
                "하나의 LLM 입력으로 전달하세요."
            )

        records.append(
            {
                "campaign_id": optional_text(campaign_row.get("campaign_id")),
                "source_sheet": optional_text(campaign_row.get("source_sheet")),
                "raw_row_number": normalize_row_number(
                    campaign_row.get("raw_row_number")
                ),
                "platform": platform.value,
                "permalink": permalink,
                "conversation_stream": optional_text(
                    campaign_row.get("conversation_stream")
                ),
                "user_name": optional_text(campaign_row.get("user_name")),
                "profile_url": optional_text(campaign_row.get("profile_url")),
                **build_sender_profile_output_fields(campaign_row),
                "post_media_type": post_media_type,
                "media_count": output_media_count,
                "media_types": media_types,
                "media_source_urls": media_source_urls,
                "local_media_paths": local_media_paths,
                "successful_media_count": successful_count,
                "failed_media_count": failed_count,
                "extraction_methods": extraction_methods,
                "status": status,
                "error_message": error_message,
                "llm_input_mode": llm_input_mode,
                "llm_input_value": llm_input_value,
                "llm_ready": llm_ready,
                "user_action_required": user_action_required,
                "user_action": user_action,
            }
        )

    return pd.DataFrame(
        records,
        columns=LLM_INPUT_COLUMNS,
    )


def save_result_excel(
    output_excel_path: Path,
    llm_input_df: pd.DataFrame,
) -> None:
    """최종 결과를 post 단위 llm_input 시트 하나로 저장한다."""

    with pd.ExcelWriter(
        output_excel_path,
        engine="openpyxl",
    ) as writer:
        llm_input_df.to_excel(
            writer,
            sheet_name=LLM_INPUT_SHEET_NAME,
            index=False,
        )

        worksheet = writer.book[LLM_INPUT_SHEET_NAME]
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions

        wrap_columns = {
            "permalink",
            "conversation_stream",
            "sender_location",
            "sender_detailed_location",
            "sender_bio",
            "sender_website",
            "sender_profile_tags",
            "media_types",
            "media_source_urls",
            "local_media_paths",
            "extraction_methods",
            "error_message",
            "llm_input_value",
            "user_action",
        }

        for cell in worksheet[1]:
            cell.font = cell.font.copy(bold=True)

        for column_cells in worksheet.columns:
            header = str(column_cells[0].value or "")
            width = 18

            if header in {
                "conversation_stream",
                "sender_bio",
                "sender_profile_tags",
                "media_source_urls",
                "local_media_paths",
                "error_message",
                "llm_input_value",
                "user_action",
            }:
                width = 45
            elif header in {
                "permalink",
                "profile_url",
                "sender_website",
                "sender_location",
                "sender_detailed_location",
            }:
                width = 35
            elif header in {
                "campaign_id",
                "extraction_methods",
            }:
                width = 28
            elif header == "sender_follower_count":
                width = 22

            worksheet.column_dimensions[column_cells[0].column_letter].width = width

            if header in wrap_columns:
                for cell in column_cells[1:]:
                    cell.alignment = cell.alignment.copy(
                        wrap_text=True,
                        vertical="top",
                    )


# =========================================================
# 10. Main
# =========================================================

def main() -> None:
    args = parse_arguments()

    input_date = input(
        "조회 날짜를 입력하세요 (YYMMDD): "
    ).strip()

    try:
        datetime.strptime(
            input_date,
            "%y%m%d",
        )
    except ValueError as exc:
        raise ValueError(
            "날짜는 YYMMDD 형식으로 입력해야 합니다. 예시) 260714"
        ) from exc

    output_dir, media_dir = (
        resolve_missing_execution_directories(
            input_date=input_date,
        )
    )

    input_excel_path, output_excel_path = build_excel_paths(
        input_date=input_date,
        output_dir=output_dir,
    )

    (
        temporary_excel_path,
        temporary_media_dir,
        backup_media_dir,
    ) = prepare_temporary_artifacts(
        input_excel_path=input_excel_path,
        output_excel_path=output_excel_path,
        media_dir=media_dir,
        overwrite=args.overwrite,
    )

    print(f"입력 파일: {input_excel_path}")
    print(f"출력 파일: {output_excel_path}")
    print(f"미디어 저장 폴더: {media_dir}")

    try:
        # 1. Raw sheet 읽기
        raw_sheets = load_raw_sheets(
            excel_path=input_excel_path,
            sheet_names=RAW_SHEET_NAMES,
        )
        print(
            f"Raw sheet 로드 완료: "
            f"{list(raw_sheets.keys())}"
        )

        # 2. 필요한 컬럼 추출 및 표준화
        media_input_df = build_media_input_dataframe(
            raw_sheets=raw_sheets,
            column_mapping=COLUMN_MAPPINGS,
        )
        print(
            f"컬럼 표준화 완료: "
            f"{media_input_df.shape}"
        )

        # 3. 값 정리
        media_input_df = clean_input_dataframe(
            media_input_df
        )
        media_input_df = prepare_campaign_input_for_processing(
            media_input_df
        )
        print(
            f"입력 데이터 정리 완료: "
            f"{media_input_df.shape}"
        )

        # 4. DataFrame 행 -> StructuredMediaInput
        structured_inputs = extract_rows_to_inputs(
            media_input_df
        )
        print(
            "StructuredMediaInput 생성 완료: "
            f"{len(structured_inputs)}개"
        )

        # 5. 게시글 단위 -> 개별 MediaAsset
        media_assets = build_media_assets(
            structured_inputs
        )
        print(
            f"MediaAsset 생성 완료: "
            f"{len(media_assets)}개"
        )

        # 5-1. TikTok 게시물 URL은 yt-dlp로 임시 media 폴더에 저장한다.
        media_assets = process_tiktok_media_assets_with_yt_dlp(
            media_assets=media_assets,
            local_media_root=temporary_media_dir,
        )

        # 5-2. Twitter LINK/UNKNOWN 게시물을 gallery-dl로 처리한다.
        media_assets = resolve_page_media_assets_with_gallery_dl(
            media_assets=media_assets,
            only_failed_unknown=False,
        )

        # 6. 1차 URL feasibility 검사 및 직접 다운로드
        processed_assets = process_media_assets(
            media_assets=media_assets,
            local_media_root=temporary_media_dir,
        )

        # 6-1. TikTok을 제외하고 직접 URL 처리에 실패한 UNKNOWN 행은
        #      원본 게시물 URL을 gallery-dl에 다시 전달한다.
        processed_assets = resolve_page_media_assets_with_gallery_dl(
            media_assets=processed_assets,
            only_failed_unknown=True,
        )

        # 6-2. gallery-dl fallback으로 새 pending asset이 생성된 경우
        #      2차 다운로드를 실행한다.
        has_pending_assets = any(
            asset.status == "pending"
            for asset in processed_assets
        )

        if has_pending_assets:
            processed_assets = process_media_assets(
                media_assets=processed_assets,
                local_media_root=temporary_media_dir,
            )

        # 6-3. 자동 처리 실패 건은 사용자 수동 처리 대상으로 표시한다.
        processed_assets = mark_remaining_failures_for_manual_action(
            media_assets=processed_assets,
        )

        # 임시 폴더에 저장된 실제 파일 경로를
        # 최종 media/{date}_{차수} 경로로 변환한다.
        processed_assets = rebase_media_asset_local_paths(
            media_assets=processed_assets,
            temporary_media_dir=temporary_media_dir,
            final_media_dir=media_dir,
        )

        # 7. asset 결과를 post 단위 LLM input 1행으로 집계한다.
        manifest_df = media_assets_to_dataframe(
            processed_assets
        )
        llm_input_df = build_llm_input_dataframe(
            campaign_input_df=media_input_df,
            manifest_df=manifest_df,
        )

        # 최종 Excel에 직접 쓰지 않고 임시 Excel을 먼저 완성한다.
        save_result_excel(
            output_excel_path=temporary_excel_path,
            llm_input_df=llm_input_df,
        )

        # 미디어와 Excel이 모두 완성된 뒤에만 기존 결과를 교체한다.
        commit_output_artifacts(
            temporary_excel_path=temporary_excel_path,
            output_excel_path=output_excel_path,
            temporary_media_dir=temporary_media_dir,
            media_dir=media_dir,
            backup_media_dir=backup_media_dir,
        )

    except Exception:
        cleanup_temporary_artifacts(
            temporary_excel_path=temporary_excel_path,
            temporary_media_dir=temporary_media_dir,
        )
        raise

    if llm_input_df.empty:
        status_counts: dict[Any, int] = {}
        print(
            "처리할 게시물이 없어 헤더만 있는 "
            "llm_input 시트를 생성했습니다."
        )
    else:
        status_counts = llm_input_df[
            "status"
        ].value_counts(
            dropna=False
        ).to_dict()

    follower_count_rows = (
        int(
            llm_input_df[
                "sender_follower_count"
            ].notna().sum()
        )
        if (
            not llm_input_df.empty
            and "sender_follower_count"
            in llm_input_df.columns
        )
        else 0
    )

    print(
        f"결과 Excel 저장 완료: "
        f"{output_excel_path}"
    )
    print(
        f"게시물 단위 처리 상태 요약: "
        f"{status_counts}"
    )
    print(
        "Sender Follower Count 전달 완료 행 수: "
        f"{follower_count_rows}/{len(llm_input_df)}"
    )


if __name__ == "__main__":
    main()
