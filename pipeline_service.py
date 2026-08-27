from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
from typing import Callable, Literal

from pipeline_run_paths import (
    PipelineRunPaths,
    find_numbered_run_directories,
    prepare_pipeline_run_paths,
    validate_input_date,
)


PROJECT_ROOT = Path(__file__).resolve().parent
OUTPUT_ROOT = PROJECT_ROOT / "output"
MEDIA_ROOT = PROJECT_ROOT / "media"

# Buzz Volume은 날짜별/차수별 실행 폴더와 분리된 공용 작업 폴더다.
BUZZ_VOLUME_ROOT = OUTPUT_ROOT / "Buzz_Volume"
BUZZ_VOLUME_COMPLETED_DIR = (
    BUZZ_VOLUME_ROOT / "completed"
)

# 누락건은 기존 운영 방식 그대로 프로젝트 루트의 `누락/` 서브 파이프라인을 사용한다.
# Streamlit은 이 폴더로 직접 이동하는 대신 subprocess의 cwd를 MISSING_ROOT로 지정한다.
MISSING_ROOT = PROJECT_ROOT / "누락"
MISSING_RUN_PIPELINE = MISSING_ROOT / "run_pipeline.py"
MISSING_PAYLOAD = PROJECT_ROOT / "payload" / "payload_6_1_누락건.json"
MISSING_OUTPUT_ROOT = MISSING_ROOT / "output_누락"
MISSING_MEDIA_ROOT = MISSING_ROOT / "media_누락"

MISSING_PIPELINE_MODULES = (
    "sprinklr_export_excel.py",
    "raw_to_processed.py",
    "media_extractor.py",
    "llm_analysis_pipeline.py",
)

MISSING_REQUIRED_PATHS = (
    MISSING_RUN_PIPELINE,
    *(MISSING_ROOT / module_name for module_name in MISSING_PIPELINE_MODULES),
    MISSING_PAYLOAD,
)

SPRINKLR_MODULE = "sprinklr_export_excel.py"
BUZZ_VOLUME_MODULE = "buzz_volume_adaptor.py"
BUZZ_VOLUME_PAYLOAD = PROJECT_ROOT / "payload" / "buzz_volume_base_payload.json"
BUZZ_VOLUME_DAILY_START_DATE = date(2026, 7, 1)

FOLLOW_UP_MODULES = (
    "raw_to_processed.py",
    "media_extractor.py",
    "llm_analysis_pipeline.py",
)

PIPELINE_MODULES = (
    SPRINKLR_MODULE,
    *FOLLOW_UP_MODULES,
)

MODULE_OVERWRITE_ARGUMENTS = {
    "sprinklr_export_excel.py": "--overwrite",
    "raw_to_processed.py": "--overwrite",
    "media_extractor.py": "--overwrite",
    "llm_analysis_pipeline.py": "--overwrite-results",
}

# 하위 모듈에 공통 실행 경로를 전달하기 위한 환경변수명
ENV_INPUT_DATE = "LOCAL_CAMPAIGN_INPUT_DATE"
ENV_RUN_NUMBER = "LOCAL_CAMPAIGN_RUN_NUMBER"
ENV_OUTPUT_DIR = "LOCAL_CAMPAIGN_OUTPUT_DIR"
ENV_MEDIA_DIR = "LOCAL_CAMPAIGN_MEDIA_DIR"


PipelineStepStatus = Literal[
    "started",
    "completed",
]


@dataclass(frozen=True)
class PipelineProgress:
    """UI/CLI에 전달할 파이프라인 진행 상태."""

    current_step: int
    total_steps: int
    module_name: str
    status: PipelineStepStatus

    @property
    def fraction(self) -> float:
        if self.total_steps <= 0:
            return 0.0

        completed_steps = (
            self.current_step
            if self.status == "completed"
            else self.current_step - 1
        )

        return max(
            0.0,
            min(
                completed_steps / self.total_steps,
                1.0,
            ),
        )


@dataclass(frozen=True)
class PipelineRunResult:
    """한 번의 Local Campaign 실행 결과 메타데이터."""

    input_date: str
    normalized_start_datetime: str
    normalized_end_datetime: str
    run_paths: PipelineRunPaths
    buzz_volume_root: Path
    buzz_volume_completed_dir: Path


@dataclass(frozen=True)
class ModuleDependencyCheck:
    """개별 모듈 실행 전 선행조건 확인 결과."""

    name: str
    passed: bool
    detail: str


@dataclass(frozen=True)
class SingleModuleRunResult:
    """개별 모듈 1회 실행 결과 메타데이터."""

    module_name: str
    input_date: str
    run_paths: PipelineRunPaths
    overwrite_existing: bool
    normalized_start_datetime: str | None = None
    normalized_end_datetime: str | None = None


@dataclass(frozen=True)
class BuzzVolumeRunResult:
    """Buzz Volume 1회 실행 결과 메타데이터."""

    data_cut_type: str
    reference_date: str
    input_excel_path: Path
    output_excel_path: Path
    buzz_volume_root: Path
    buzz_volume_completed_dir: Path
    overwrite_existing_result: bool


@dataclass(frozen=True)
class MissingCasesRunResult:
    """누락건 서브 파이프라인 1회 실행 결과 메타데이터."""

    input_date: str
    normalized_start_datetime: str
    normalized_end_datetime: str
    query_text: str
    payload_path: Path
    missing_root: Path
    output_dir: Path
    media_dir: Path


LogCallback = Callable[[str], None]
ProgressCallback = Callable[[PipelineProgress], None]


def _emit_log(
    message: str = "",
    *,
    log_callback: LogCallback | None = None,
) -> None:
    """
    CLI에서는 stdout으로 출력하고,
    UI에서는 callback으로 같은 로그를 전달한다.
    """

    if log_callback is None:
        print(
            message,
            flush=True,
        )
        return

    log_callback(
        message
    )


def _emit_progress(
    *,
    current_step: int,
    total_steps: int,
    module_name: str,
    status: PipelineStepStatus,
    progress_callback: ProgressCallback | None,
) -> None:
    if progress_callback is None:
        return

    progress_callback(
        PipelineProgress(
            current_step=current_step,
            total_steps=total_steps,
            module_name=module_name,
            status=status,
        )
    )


def ensure_buzz_volume_directories() -> tuple[Path, Path]:
    """
    Buzz Volume 공용 입력/완료 폴더가 존재하도록 보장한다.

    구조:
        output/
        └─ Buzz_Volume/
           ├─ 사용자가 최종 통합·정제 Excel을 넣는 위치
           └─ completed/
              └─ Buzz Volume 적재 완료 결과 저장 위치

    이 폴더들은 날짜별 실행 차수와 독립적이므로,
    이미 존재하면 그대로 유지한다.
    """

    BUZZ_VOLUME_ROOT.mkdir(
        parents=True,
        exist_ok=True,
    )

    BUZZ_VOLUME_COMPLETED_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    return (
        BUZZ_VOLUME_ROOT,
        BUZZ_VOLUME_COMPLETED_DIR,
    )


def _build_gcloud_command(
    gcloud_path: str,
    *arguments: str,
) -> list[str]:
    """
    현재 OS에서 gcloud 명령을 안정적으로 실행할 command list를 만든다.

    Windows에서는 gcloud가 gcloud.cmd / gcloud.bat 형태일 수 있으므로
    cmd.exe를 통해 실행한다.
    """

    suffix = Path(gcloud_path).suffix.casefold()

    if os.name == "nt" and suffix in {
        ".cmd",
        ".bat",
    }:
        return [
            "cmd.exe",
            "/d",
            "/c",
            gcloud_path,
            *arguments,
        ]

    return [
        gcloud_path,
        *arguments,
    ]


def ensure_google_adc(
    *,
    log_callback: LogCallback | None = None,
) -> None:
    """
    Google Application Default Credentials(ADC)를 확인한다.

    처리 순서:
    1. 현재 ADC로 access token 발급 가능 여부 확인
    2. 정상이라면 그대로 pipeline 진행
    3. ADC가 없거나 만료되었다면
       `gcloud auth application-default login` 자동 실행
    4. 로그인 완료 후 ADC를 다시 검증
    """

    gcloud_path = shutil.which(
        "gcloud"
    )

    if not gcloud_path:
        raise RuntimeError(
            "gcloud CLI를 찾을 수 없습니다.\n"
            "Google Cloud SDK가 설치되어 있고 "
            "PATH에 등록되어 있는지 확인하세요."
        )

    _emit_log(
        log_callback=log_callback,
    )
    _emit_log(
        "=" * 70,
        log_callback=log_callback,
    )
    _emit_log(
        "Google Cloud 인증 확인",
        log_callback=log_callback,
    )
    _emit_log(
        "=" * 70,
        log_callback=log_callback,
    )

    check_result = subprocess.run(
        _build_gcloud_command(
            gcloud_path,
            "auth",
            "application-default",
            "print-access-token",
        ),
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        check=False,
        cwd=PROJECT_ROOT,
    )

    if check_result.returncode == 0:
        _emit_log(
            "✅ Google Cloud Application Default Credentials 정상",
            log_callback=log_callback,
        )
        return

    _emit_log(
        "Google Cloud 인증이 없거나 만료되었습니다.",
        log_callback=log_callback,
    )
    _emit_log(
        "gcloud auth application-default login을 실행합니다.",
        log_callback=log_callback,
    )
    _emit_log(
        log_callback=log_callback,
    )

    login_result = subprocess.run(
        _build_gcloud_command(
            gcloud_path,
            "auth",
            "application-default",
            "login",
        ),
        check=False,
        cwd=PROJECT_ROOT,
    )

    if login_result.returncode != 0:
        raise RuntimeError(
            "Google Cloud Application Default Credentials "
            "로그인에 실패했습니다.\n"
            f"gcloud return code: "
            f"{login_result.returncode}"
        )

    verify_result = subprocess.run(
        _build_gcloud_command(
            gcloud_path,
            "auth",
            "application-default",
            "print-access-token",
        ),
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        check=False,
        cwd=PROJECT_ROOT,
    )

    if verify_result.returncode != 0:
        raise RuntimeError(
            "Google Cloud 로그인은 완료되었지만 "
            "Application Default Credentials 검증에 "
            "실패했습니다."
        )

    _emit_log(
        log_callback=log_callback,
    )
    _emit_log(
        "✅ Google Cloud Application Default Credentials 인증 완료",
        log_callback=log_callback,
    )


def parse_datetime(
    datetime_text: str,
    input_name: str,
) -> datetime:
    """
    YYYY-MM-DD HH:MM:SS 형식의 문자열을 datetime으로 변환한다.
    """

    datetime_text = datetime_text.strip()

    try:
        return datetime.strptime(
            datetime_text,
            "%Y-%m-%d %H:%M:%S",
        )

    except ValueError as exc:
        raise ValueError(
            f"{input_name}은 YYYY-MM-DD HH:MM:SS "
            "형식이어야 합니다.\n"
            f"입력값: {datetime_text}\n"
            "예시: 2026-07-27 19:00:00"
        ) from exc


def normalize_pipeline_datetimes(
    start_datetime: datetime | str,
    end_datetime: datetime | str,
) -> tuple[datetime, datetime, str, str]:
    """CLI와 Web UI가 공통으로 사용할 날짜/시간 정규화."""

    if isinstance(
        start_datetime,
        str,
    ):
        start_datetime_obj = parse_datetime(
            datetime_text=start_datetime,
            input_name="시작 날짜와 시간",
        )
    elif isinstance(
        start_datetime,
        datetime,
    ):
        start_datetime_obj = start_datetime
    else:
        raise TypeError(
            "start_datetime은 datetime 또는 문자열이어야 합니다."
        )

    if isinstance(
        end_datetime,
        str,
    ):
        end_datetime_obj = parse_datetime(
            datetime_text=end_datetime,
            input_name="종료 날짜와 시간",
        )
    elif isinstance(
        end_datetime,
        datetime,
    ):
        end_datetime_obj = end_datetime
    else:
        raise TypeError(
            "end_datetime은 datetime 또는 문자열이어야 합니다."
        )

    if end_datetime_obj <= start_datetime_obj:
        raise ValueError(
            "종료 날짜와 시간은 시작 날짜와 시간보다 "
            "늦어야 합니다.\n"
            f"시작: {start_datetime_obj}\n"
            f"종료: {end_datetime_obj}"
        )

    normalized_start_datetime = (
        start_datetime_obj.strftime(
            "%Y-%m-%d %H:%M:%S"
        )
    )

    normalized_end_datetime = (
        end_datetime_obj.strftime(
            "%Y-%m-%d %H:%M:%S"
        )
    )

    return (
        start_datetime_obj,
        end_datetime_obj,
        normalized_start_datetime,
        normalized_end_datetime,
    )


def build_module_environment(
    run_paths: PipelineRunPaths,
) -> dict[str, str]:
    """
    하위 모듈에 전달할 환경변수를 생성한다.

    기존 모듈의 input() 입력 순서는 그대로 유지하고,
    실행 차수별 output/media 경로는 환경변수로 별도 전달한다.
    """

    module_environment = os.environ.copy()

    module_environment.update(
        {
            ENV_INPUT_DATE: run_paths.input_date,
            ENV_RUN_NUMBER: str(
                run_paths.run_number
            ),
            ENV_OUTPUT_DIR: str(
                run_paths.output_dir
            ),
            ENV_MEDIA_DIR: str(
                run_paths.media_dir
            ),
            "PYTHONUNBUFFERED": "1",
        }
    )

    return module_environment


def run_module(
    module_name: str,
    module_inputs: list[str],
    run_paths: PipelineRunPaths,
    *,
    module_arguments: list[str] | None = None,
    log_callback: LogCallback | None = None,
) -> None:
    """
    기존 하위 모듈의 input() 계약을 그대로 유지하면서 실행한다.

    stdout/stderr는 한 줄씩 전달하므로,
    CLI에서는 기존처럼 터미널에서 볼 수 있고 향후 Streamlit에서는
    log_callback을 이용해 실시간 로그 영역에 표시할 수 있다.
    """

    module_path = (
        PROJECT_ROOT / module_name
    )

    if not module_path.exists():
        raise FileNotFoundError(
            "모듈 파일을 찾을 수 없습니다: "
            f"{module_path}"
        )

    if not module_path.is_file():
        raise FileNotFoundError(
            "모듈 경로가 파일이 아닙니다: "
            f"{module_path}"
        )

    stdin_text = (
        "\n".join(module_inputs)
        + "\n"
    )

    module_environment = (
        build_module_environment(
            run_paths=run_paths
        )
    )

    _emit_log(
        log_callback=log_callback,
    )
    _emit_log(
        "=" * 70,
        log_callback=log_callback,
    )
    _emit_log(
        f"실행 시작: {module_name}",
        log_callback=log_callback,
    )

    for index, input_value in enumerate(
        module_inputs,
        start=1,
    ):
        _emit_log(
            f"전달 입력값 {index}: {input_value}",
            log_callback=log_callback,
        )

    _emit_log(
        f"실행 차수: {run_paths.run_label}",
        log_callback=log_callback,
    )
    _emit_log(
        f"Output 폴더: {run_paths.output_dir}",
        log_callback=log_callback,
    )
    _emit_log(
        f"Media 폴더: {run_paths.media_dir}",
        log_callback=log_callback,
    )
    _emit_log(
        "=" * 70,
        log_callback=log_callback,
    )

    command = [
        sys.executable,
        str(module_path),
        *(module_arguments or []),
    ]

    if module_arguments:
        _emit_log(
            "실행 옵션: " + " ".join(module_arguments),
            log_callback=log_callback,
        )

    process = subprocess.Popen(
        command,
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1,
        cwd=PROJECT_ROOT,
        env=module_environment,
    )

    if process.stdin is None:
        raise RuntimeError(
            f"{module_name} stdin pipe를 생성하지 못했습니다."
        )

    process.stdin.write(
        stdin_text
    )
    process.stdin.close()

    if process.stdout is not None:
        for line in process.stdout:
            _emit_log(
                line.rstrip("\r\n"),
                log_callback=log_callback,
            )

    return_code = process.wait()

    if return_code != 0:
        raise RuntimeError(
            f"{module_name} 실행 실패 "
            f"(return code: {return_code})\n"
            f"실행 차수: {run_paths.run_label}\n"
            f"Output 폴더: {run_paths.output_dir}\n"
            f"Media 폴더: {run_paths.media_dir}"
        )

    _emit_log(
        f"✅ 실행 완료: {module_name}",
        log_callback=log_callback,
    )


def run_local_campaign_pipeline(
    start_datetime: datetime | str,
    end_datetime: datetime | str,
    *,
    log_callback: LogCallback | None = None,
    progress_callback: ProgressCallback | None = None,
    check_google_adc: bool = True,
) -> PipelineRunResult:
    """
    Local Campaign 1~4단계를 실행하는 공통 Backend 서비스.

    CLI(run_pipeline.py)와 향후 Web UI(streamlit_app.py)가
    이 함수를 동일하게 호출한다.
    """

    if check_google_adc:
        ensure_google_adc(
            log_callback=log_callback,
        )

    (
        buzz_volume_root,
        buzz_volume_completed_dir,
    ) = ensure_buzz_volume_directories()

    _emit_log(
        log_callback=log_callback,
    )
    _emit_log(
        "Buzz Volume 공용 폴더 확인",
        log_callback=log_callback,
    )
    _emit_log(
        f"- 입력 파일 위치: {buzz_volume_root}",
        log_callback=log_callback,
    )
    _emit_log(
        f"- 완료 결과 위치: {buzz_volume_completed_dir}",
        log_callback=log_callback,
    )

    (
        _start_datetime_obj,
        end_datetime_obj,
        normalized_start_datetime,
        normalized_end_datetime,
    ) = normalize_pipeline_datetimes(
        start_datetime=start_datetime,
        end_datetime=end_datetime,
    )

    input_date = (
        end_datetime_obj.strftime(
            "%y%m%d"
        )
    )

    _emit_log(
        log_callback=log_callback,
    )
    _emit_log(
        "입력값 확인",
        log_callback=log_callback,
    )
    _emit_log(
        f"- 조회 시작 시각: {normalized_start_datetime}",
        log_callback=log_callback,
    )
    _emit_log(
        f"- 조회 종료 시각: {normalized_end_datetime}",
        log_callback=log_callback,
    )
    _emit_log(
        f"- 후속 모듈 작업 날짜: {input_date}",
        log_callback=log_callback,
    )

    run_paths = prepare_pipeline_run_paths(
        input_date=input_date,
        output_root=OUTPUT_ROOT,
        media_root=MEDIA_ROOT,
    )

    _emit_log(
        log_callback=log_callback,
    )
    _emit_log(
        "실행 폴더 확인",
        log_callback=log_callback,
    )
    _emit_log(
        f"- 실행 차수: {run_paths.run_label}",
        log_callback=log_callback,
    )
    _emit_log(
        f"- Output 폴더: {run_paths.output_dir}",
        log_callback=log_callback,
    )
    _emit_log(
        f"- Media 폴더: {run_paths.media_dir}",
        log_callback=log_callback,
    )

    module_jobs = (
        (
            SPRINKLR_MODULE,
            [
                normalized_start_datetime,
                normalized_end_datetime,
            ],
        ),
        *(
            (
                module_name,
                [input_date],
            )
            for module_name in FOLLOW_UP_MODULES
        ),
    )

    total_steps = len(
        module_jobs
    )

    for current_step, (
        module_name,
        module_inputs,
    ) in enumerate(
        module_jobs,
        start=1,
    ):
        _emit_progress(
            current_step=current_step,
            total_steps=total_steps,
            module_name=module_name,
            status="started",
            progress_callback=progress_callback,
        )

        run_module(
            module_name=module_name,
            module_inputs=module_inputs,
            run_paths=run_paths,
            log_callback=log_callback,
        )

        _emit_progress(
            current_step=current_step,
            total_steps=total_steps,
            module_name=module_name,
            status="completed",
            progress_callback=progress_callback,
        )

    _emit_log(
        log_callback=log_callback,
    )
    _emit_log(
        "=" * 70,
        log_callback=log_callback,
    )
    _emit_log(
        "✅ 전체 Local Campaign 자동화 파이프라인 실행 완료",
        log_callback=log_callback,
    )
    _emit_log(
        f"실행 차수: {run_paths.run_label}",
        log_callback=log_callback,
    )
    _emit_log(
        f"Output 폴더: {run_paths.output_dir}",
        log_callback=log_callback,
    )
    _emit_log(
        f"Media 폴더: {run_paths.media_dir}",
        log_callback=log_callback,
    )
    _emit_log(
        f"Buzz Volume 입력 폴더: {buzz_volume_root}",
        log_callback=log_callback,
    )
    _emit_log(
        f"Buzz Volume 완료 폴더: {buzz_volume_completed_dir}",
        log_callback=log_callback,
    )
    _emit_log(
        log_callback=log_callback,
    )
    _emit_log(
        "1차·2차 결과를 통합하고 최종 정제한 Excel을 "
        "Buzz Volume 입력 폴더에 넣은 뒤 "
        "buzz_volume_adaptor.py를 별도로 실행하세요.",
        log_callback=log_callback,
    )
    _emit_log(
        "=" * 70,
        log_callback=log_callback,
    )

    return PipelineRunResult(
        input_date=input_date,
        normalized_start_datetime=(
            normalized_start_datetime
        ),
        normalized_end_datetime=(
            normalized_end_datetime
        ),
        run_paths=run_paths,
        buzz_volume_root=buzz_volume_root,
        buzz_volume_completed_dir=(
            buzz_volume_completed_dir
        ),
    )


# =============================================================================
# Individual Module Service
# =============================================================================

def list_existing_run_paths(
    input_date: str,
) -> tuple[PipelineRunPaths, ...]:
    """
    특정 날짜의 기존 실행 차수를 안전하게 조회한다.

    이 함수는 새 폴더를 만들거나 이름을 변경하지 않는다.

    1차는 두 형태 중 하나일 수 있다.
    - 첫 실행만 존재: output/{YYMMDD}, media/{YYMMDD}
    - 2차 이상 존재: output/{YYMMDD}_1차, media/{YYMMDD}_1차
    """

    normalized_date = validate_input_date(
        input_date
    )

    output_root = OUTPUT_ROOT.resolve()
    media_root = MEDIA_ROOT.resolve()

    base_output_dir = (
        output_root / normalized_date
    )
    base_media_dir = (
        media_root / normalized_date
    )

    base_output_exists = base_output_dir.is_dir()
    base_media_exists = base_media_dir.is_dir()

    if base_output_exists != base_media_exists:
        raise RuntimeError(
            "기본 날짜 output/media 폴더 상태가 서로 다릅니다.\n"
            f"output: {base_output_dir} -> {base_output_exists}\n"
            f"media: {base_media_dir} -> {base_media_exists}"
        )

    output_numbered = find_numbered_run_directories(
        output_root,
        normalized_date,
    )
    media_numbered = find_numbered_run_directories(
        media_root,
        normalized_date,
    )

    if set(output_numbered) != set(media_numbered):
        raise RuntimeError(
            "output과 media의 실행 차수가 서로 다릅니다.\n"
            f"output 차수={sorted(output_numbered)}\n"
            f"media 차수={sorted(media_numbered)}"
        )

    run_numbers = set(output_numbered)
    if run_numbers:
        expected_run_numbers = set(
            range(1, max(run_numbers) + 1)
        )
        if run_numbers != expected_run_numbers:
            raise RuntimeError(
                "기존 실행 차수 폴더가 1차부터 연속적이지 않습니다.\n"
                f"현재 차수={sorted(run_numbers)}\n"
                f"누락 차수={sorted(expected_run_numbers - run_numbers)}"
            )

    if base_output_exists and output_numbered:
        raise RuntimeError(
            "기본 날짜 폴더와 차수 폴더가 동시에 존재합니다. "
            "실행 이력이 모호하므로 개별 실행을 중단합니다.\n"
            f"기본 output: {base_output_dir}\n"
            f"번호 차수: {sorted(output_numbered)}"
        )

    runs: list[PipelineRunPaths] = []

    if base_output_exists:
        runs.append(
            PipelineRunPaths(
                input_date=normalized_date,
                run_number=1,
                output_dir=base_output_dir,
                media_dir=base_media_dir,
            )
        )
        return tuple(runs)

    for run_number in sorted(output_numbered):
        runs.append(
            PipelineRunPaths(
                input_date=normalized_date,
                run_number=run_number,
                output_dir=output_numbered[run_number],
                media_dir=media_numbered[run_number],
            )
        )

    return tuple(runs)


def resolve_existing_run_paths(
    input_date: str,
    run_number: int,
) -> PipelineRunPaths:
    """선택한 날짜/차수의 기존 output/media 폴더를 확정한다."""

    for run_paths in list_existing_run_paths(
        input_date
    ):
        if run_paths.run_number == run_number:
            return run_paths

    normalized_date = validate_input_date(
        input_date
    )

    raise FileNotFoundError(
        "선택한 실행 차수의 기존 output/media 폴더를 찾지 못했습니다.\n"
        f"작업 날짜: {normalized_date}\n"
        f"실행 차수: {run_number}차\n"
        "먼저 Full Pipeline을 실행하거나 실제 존재하는 차수를 선택하세요."
    )


def build_expected_run_artifacts(
    run_paths: PipelineRunPaths,
) -> dict[str, Path]:
    """현재 파일명 contract 기준으로 모듈별 주요 산출물 경로를 계산한다."""

    input_date_obj = datetime.strptime(
        run_paths.input_date,
        "%y%m%d",
    )
    input_month = input_date_obj.month

    raw_excel = run_paths.output_dir / (
        f"{run_paths.input_date}_SLCC_SOV_Local Campaign Tracking_"
        f"{input_month}월_v01.xlsx"
    )

    formatted_excel = raw_excel.with_name(
        raw_excel.stem + "_formatted.xlsx"
    )

    media_result_excel = run_paths.output_dir / (
        f"{run_paths.input_date}_campaign_media_result.xlsx"
    )

    llm_log_excel = run_paths.output_dir / (
        f"{run_paths.input_date}_campaign_media_result_llm_result.xlsx"
    )

    llm_completed_excel = formatted_excel.with_name(
        formatted_excel.stem
        + "_llm_completed"
        + formatted_excel.suffix
    )

    return {
        "raw_excel": raw_excel,
        "formatted_excel": formatted_excel,
        "media_result_excel": media_result_excel,
        "llm_log_excel": llm_log_excel,
        "llm_completed_excel": llm_completed_excel,
    }


def validate_module_dependencies(
    module_name: str,
    input_date: str,
    run_number: int,
) -> tuple[ModuleDependencyCheck, ...]:
    """
    개별 모듈 실행 전 최소 선행조건을 점검한다.

    파일 존재 여부를 확인할 뿐 파일을 변경하지 않는다.
    """

    if module_name not in PIPELINE_MODULES:
        raise ValueError(
            "지원하지 않는 모듈입니다: "
            f"{module_name}"
        )

    try:
        run_paths = resolve_existing_run_paths(
            input_date=input_date,
            run_number=run_number,
        )
    except Exception as exc:
        return (
            ModuleDependencyCheck(
                name="기존 실행 폴더",
                passed=False,
                detail=f"{type(exc).__name__}: {exc}",
            ),
        )

    artifacts = build_expected_run_artifacts(
        run_paths
    )

    checks: list[ModuleDependencyCheck] = [
        ModuleDependencyCheck(
            name="Output directory",
            passed=run_paths.output_dir.is_dir(),
            detail=str(run_paths.output_dir),
        ),
        ModuleDependencyCheck(
            name="Media directory",
            passed=run_paths.media_dir.is_dir(),
            detail=str(run_paths.media_dir),
        ),
    ]

    if module_name == SPRINKLR_MODULE:
        # 1단계 재실행은 기존 차수 폴더만 있으면 실행할 수 있다.
        return tuple(checks)

    if module_name == "raw_to_processed.py":
        checks.append(
            ModuleDependencyCheck(
                name="Sprinklr Raw Excel",
                passed=artifacts["raw_excel"].is_file(),
                detail=str(artifacts["raw_excel"]),
            )
        )

    elif module_name == "media_extractor.py":
        # media_extractor.py는 Processed Excel이 아니라
        # 같은 실행 차수의 Sprinklr Raw Excel(Raw Data 시트)을 읽는다.
        checks.append(
            ModuleDependencyCheck(
                name="Sprinklr Raw Excel",
                passed=artifacts["raw_excel"].is_file(),
                detail=str(artifacts["raw_excel"]),
            )
        )

    elif module_name == "llm_analysis_pipeline.py":
        checks.extend(
            (
                ModuleDependencyCheck(
                    name="Formatted Excel",
                    passed=artifacts["formatted_excel"].is_file(),
                    detail=str(artifacts["formatted_excel"]),
                ),
                ModuleDependencyCheck(
                    name="Media Result Excel",
                    passed=artifacts["media_result_excel"].is_file(),
                    detail=str(artifacts["media_result_excel"]),
                ),
            )
        )

    return tuple(checks)


def run_single_module(
    module_name: str,
    input_date: str,
    run_number: int,
    *,
    start_datetime: datetime | str | None = None,
    end_datetime: datetime | str | None = None,
    overwrite_existing: bool = False,
    log_callback: LogCallback | None = None,
    progress_callback: ProgressCallback | None = None,
    check_google_adc: bool = True,
) -> SingleModuleRunResult:
    """
    기존 실행 차수에서 선택한 모듈 하나만 재실행한다.

    주의:
    - 새 차수 폴더를 만들지 않는다.
    - 선택한 기존 output/media 폴더만 사용한다.
    - Sprinklr Export만 start/end datetime이 필요하다.
    - overwrite_existing=True일 때 모듈별 기존 overwrite 옵션을 전달한다.
    """

    if module_name not in PIPELINE_MODULES:
        raise ValueError(
            "지원하지 않는 모듈입니다: "
            f"{module_name}"
        )

    normalized_date = validate_input_date(
        input_date
    )

    run_paths = resolve_existing_run_paths(
        input_date=normalized_date,
        run_number=run_number,
    )

    dependency_checks = validate_module_dependencies(
        module_name=module_name,
        input_date=normalized_date,
        run_number=run_number,
    )

    failed_checks = [
        check
        for check in dependency_checks
        if not check.passed
    ]

    if failed_checks:
        details = "\n".join(
            f"- {check.name}: {check.detail}"
            for check in failed_checks
        )
        raise RuntimeError(
            "개별 모듈 실행 선행조건을 충족하지 못했습니다.\n"
            + details
        )

    normalized_start_datetime: str | None = None
    normalized_end_datetime: str | None = None

    if module_name == SPRINKLR_MODULE:
        if start_datetime is None or end_datetime is None:
            raise ValueError(
                "Sprinklr Export 개별 실행에는 시작/종료 날짜와 시간이 모두 필요합니다."
            )

        (
            _start_obj,
            end_obj,
            normalized_start_datetime,
            normalized_end_datetime,
        ) = normalize_pipeline_datetimes(
            start_datetime=start_datetime,
            end_datetime=end_datetime,
        )

        end_input_date = end_obj.strftime(
            "%y%m%d"
        )

        if end_input_date != normalized_date:
            raise ValueError(
                "Sprinklr 종료 시각의 날짜와 선택한 작업 날짜가 일치하지 않습니다.\n"
                f"선택한 작업 날짜: {normalized_date}\n"
                f"종료 시각 기준 날짜: {end_input_date}"
            )

        module_inputs = [
            normalized_start_datetime,
            normalized_end_datetime,
        ]
    else:
        module_inputs = [
            normalized_date
        ]

    if (
        check_google_adc
        and module_name == "llm_analysis_pipeline.py"
    ):
        ensure_google_adc(
            log_callback=log_callback,
        )

    module_arguments: list[str] = []

    if overwrite_existing:
        overwrite_argument = (
            MODULE_OVERWRITE_ARGUMENTS.get(
                module_name
            )
        )
        if overwrite_argument:
            module_arguments.append(
                overwrite_argument
            )

    _emit_log(
        log_callback=log_callback,
    )
    _emit_log(
        "=" * 70,
        log_callback=log_callback,
    )
    _emit_log(
        "Individual Module 실행",
        log_callback=log_callback,
    )
    _emit_log(
        f"- 모듈: {module_name}",
        log_callback=log_callback,
    )
    _emit_log(
        f"- 작업 날짜: {normalized_date}",
        log_callback=log_callback,
    )
    _emit_log(
        f"- 실행 차수: {run_paths.run_label}",
        log_callback=log_callback,
    )
    _emit_log(
        f"- overwrite: {overwrite_existing}",
        log_callback=log_callback,
    )
    _emit_log(
        "=" * 70,
        log_callback=log_callback,
    )

    _emit_progress(
        current_step=1,
        total_steps=1,
        module_name=module_name,
        status="started",
        progress_callback=progress_callback,
    )

    run_module(
        module_name=module_name,
        module_inputs=module_inputs,
        run_paths=run_paths,
        module_arguments=module_arguments,
        log_callback=log_callback,
    )

    _emit_progress(
        current_step=1,
        total_steps=1,
        module_name=module_name,
        status="completed",
        progress_callback=progress_callback,
    )

    _emit_log(
        f"✅ Individual Module 실행 완료: {module_name}",
        log_callback=log_callback,
    )

    return SingleModuleRunResult(
        module_name=module_name,
        input_date=normalized_date,
        run_paths=run_paths,
        overwrite_existing=overwrite_existing,
        normalized_start_datetime=(
            normalized_start_datetime
        ),
        normalized_end_datetime=(
            normalized_end_datetime
        ),
    )

# =============================================================================
# Buzz Volume Service
# =============================================================================

def normalize_buzz_volume_data_cut_type(
    value: str,
) -> str:
    """Buzz Volume data cut을 daily/weekly 중 하나로 정규화한다."""

    normalized = str(value).strip().lower()

    if normalized not in {
        "daily",
        "weekly",
    }:
        raise ValueError(
            "Buzz Volume data_cut_type은 'daily' 또는 'weekly'여야 합니다. "
            f"입력값: {value!r}"
        )

    return normalized


def normalize_buzz_volume_reference_date(
    value: date | datetime | str,
) -> date:
    """Buzz Volume 기준 날짜를 date로 정규화한다."""

    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    if isinstance(value, str):
        try:
            return datetime.strptime(
                value.strip(),
                "%Y-%m-%d",
            ).date()
        except ValueError as exc:
            raise ValueError(
                "Buzz Volume 기준 날짜는 YYYY-MM-DD 형식이어야 합니다. "
                f"입력값: {value!r}"
            ) from exc

    raise TypeError(
        "Buzz Volume 기준 날짜는 date, datetime 또는 문자열이어야 합니다."
    )


def validate_buzz_volume_date_cut(
    data_cut_type: str,
    reference_date: date | datetime | str,
) -> tuple[str, date]:
    """현재 buzz_volume_adaptor.py의 Daily/Weekly 날짜 규칙을 사전 검증한다."""

    normalized_type = normalize_buzz_volume_data_cut_type(
        data_cut_type
    )
    normalized_date = normalize_buzz_volume_reference_date(
        reference_date
    )

    if normalized_type == "daily":
        if normalized_date < BUZZ_VOLUME_DAILY_START_DATE:
            raise ValueError(
                "Daily 기준 날짜는 고정 시작일보다 이전일 수 없습니다.\n"
                f"고정 시작일: {BUZZ_VOLUME_DAILY_START_DATE}\n"
                f"입력 날짜: {normalized_date}"
            )
    else:
        # Python weekday(): 월요일=0
        if normalized_date.weekday() != 0:
            raise ValueError(
                "Weekly 기준 날짜는 반드시 월요일이어야 합니다.\n"
                f"입력 날짜: {normalized_date}"
            )

    return normalized_type, normalized_date


def build_buzz_volume_paths(
    reference_date: date | datetime | str,
) -> tuple[Path, Path]:
    """
    Buzz Volume 공용 작업 폴더의 canonical input/output Excel 경로를 계산한다.

    input:
        output/Buzz_Volume/{YYMMDD}_SLCC_SOV_Local Campaign Tracking_{월}월_v01.xlsx

    output:
        output/Buzz_Volume/completed/{...}_mentions_updated.xlsx
    """

    normalized_date = normalize_buzz_volume_reference_date(
        reference_date
    )
    date_text = normalized_date.strftime("%y%m%d")
    month = normalized_date.month

    ensure_buzz_volume_directories()

    input_excel_path = BUZZ_VOLUME_ROOT / (
        f"{date_text}_SLCC_SOV_Local Campaign Tracking_"
        f"{month}월_v01.xlsx"
    )

    output_excel_path = BUZZ_VOLUME_COMPLETED_DIR / (
        f"{input_excel_path.stem}_mentions_updated.xlsx"
    )

    return input_excel_path, output_excel_path


def validate_buzz_volume_input(
    data_cut_type: str,
    reference_date: date | datetime | str,
    *,
    require_input_file: bool = True,
) -> tuple[ModuleDependencyCheck, ...]:
    """Buzz Volume UI/서비스가 공통으로 사용할 실행 전 검증."""

    checks: list[ModuleDependencyCheck] = []

    try:
        normalized_type, normalized_date = validate_buzz_volume_date_cut(
            data_cut_type=data_cut_type,
            reference_date=reference_date,
        )
    except Exception as exc:
        checks.append(
            ModuleDependencyCheck(
                name="Data Cut / Reference Date",
                passed=False,
                detail=f"{type(exc).__name__}: {exc}",
            )
        )
        return tuple(checks)

    checks.append(
        ModuleDependencyCheck(
            name="Data Cut / Reference Date",
            passed=True,
            detail=f"{normalized_type} / {normalized_date.isoformat()}",
        )
    )

    module_path = PROJECT_ROOT / BUZZ_VOLUME_MODULE
    checks.append(
        ModuleDependencyCheck(
            name="buzz_volume_adaptor.py",
            passed=module_path.is_file(),
            detail=str(module_path),
        )
    )

    checks.append(
        ModuleDependencyCheck(
            name="Buzz Volume Base Payload",
            passed=BUZZ_VOLUME_PAYLOAD.is_file(),
            detail=str(BUZZ_VOLUME_PAYLOAD),
        )
    )

    input_excel_path, _ = build_buzz_volume_paths(
        normalized_date
    )

    input_exists = input_excel_path.is_file()
    checks.append(
        ModuleDependencyCheck(
            name="Buzz Volume Input Excel",
            passed=(input_exists or not require_input_file),
            detail=(
                str(input_excel_path)
                if input_exists
                else (
                    f"{input_excel_path}"
                    + (
                        " (업로드 예정)"
                        if not require_input_file
                        else " (파일 없음)"
                    )
                )
            ),
        )
    )

    return tuple(checks)


def save_buzz_volume_input_file(
    file_bytes: bytes,
    reference_date: date | datetime | str,
    *,
    overwrite_existing: bool = False,
) -> Path:
    """
    Streamlit 등 UI에서 업로드한 최종 통합·정제 Excel을
    Buzz Volume canonical 파일명으로 안전하게 저장한다.
    """

    if not isinstance(file_bytes, (bytes, bytearray)):
        raise TypeError(
            "Buzz Volume 업로드 파일은 bytes여야 합니다."
        )

    if not file_bytes:
        raise ValueError(
            "업로드된 Buzz Volume Excel 파일이 비어 있습니다."
        )

    input_excel_path, _ = build_buzz_volume_paths(
        reference_date
    )

    if input_excel_path.exists() and not overwrite_existing:
        raise FileExistsError(
            "동일 기준 날짜의 Buzz Volume 입력 파일이 이미 존재합니다.\n"
            f"파일: {input_excel_path}\n"
            "기존 입력 파일을 교체하려면 overwrite_existing=True를 사용하세요."
        )

    input_excel_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    temporary_path = input_excel_path.with_name(
        f".{input_excel_path.stem}_uploading{input_excel_path.suffix}"
    )

    try:
        temporary_path.write_bytes(bytes(file_bytes))

        # xlsx는 ZIP container다. 명백히 잘못된 파일을 저장하지 않는다.
        import zipfile

        if not zipfile.is_zipfile(temporary_path):
            raise ValueError(
                "업로드 파일이 올바른 .xlsx 형식이 아닙니다."
            )

        temporary_path.replace(input_excel_path)
    finally:
        if temporary_path.exists():
            temporary_path.unlink()

    return input_excel_path


def run_buzz_volume(
    data_cut_type: str,
    reference_date: date | datetime | str,
    *,
    overwrite_existing_result: bool = False,
    log_callback: LogCallback | None = None,
    progress_callback: ProgressCallback | None = None,
) -> BuzzVolumeRunResult:
    """
    기존 buzz_volume_adaptor.py를 수정하지 않고 별도 subprocess로 실행한다.

    기존 입력 계약:
        input #1 = daily / weekly
        input #2 = YYYY-MM-DD

    Buzz Volume은 날짜별 실행 차수와 독립적인 output/Buzz_Volume을 사용한다.
    """

    normalized_type, normalized_date = validate_buzz_volume_date_cut(
        data_cut_type=data_cut_type,
        reference_date=reference_date,
    )
    reference_date_text = normalized_date.isoformat()

    dependency_checks = validate_buzz_volume_input(
        data_cut_type=normalized_type,
        reference_date=normalized_date,
        require_input_file=True,
    )

    failed_checks = [
        check
        for check in dependency_checks
        if not check.passed
    ]

    if failed_checks:
        details = "\n".join(
            f"- {check.name}: {check.detail}"
            for check in failed_checks
        )
        raise RuntimeError(
            "Buzz Volume 실행 선행조건을 충족하지 못했습니다.\n"
            + details
        )

    input_excel_path, output_excel_path = build_buzz_volume_paths(
        normalized_date
    )

    if output_excel_path.exists() and not overwrite_existing_result:
        raise FileExistsError(
            "동일 기준 날짜의 Buzz Volume 완료 결과가 이미 존재합니다.\n"
            f"파일: {output_excel_path}\n"
            "의도적인 재실행이면 overwrite_existing_result=True를 사용하세요."
        )

    module_path = PROJECT_ROOT / BUZZ_VOLUME_MODULE

    _emit_log(
        log_callback=log_callback,
    )
    _emit_log(
        "=" * 70,
        log_callback=log_callback,
    )
    _emit_log(
        "Buzz Volume 실행",
        log_callback=log_callback,
    )
    _emit_log(
        f"- Data Cut: {normalized_type}",
        log_callback=log_callback,
    )
    _emit_log(
        f"- 기준 날짜: {reference_date_text}",
        log_callback=log_callback,
    )
    _emit_log(
        f"- Input Excel: {input_excel_path}",
        log_callback=log_callback,
    )
    _emit_log(
        f"- Completed Excel: {output_excel_path}",
        log_callback=log_callback,
    )
    _emit_log(
        f"- overwrite completed result: {overwrite_existing_result}",
        log_callback=log_callback,
    )
    _emit_log(
        "=" * 70,
        log_callback=log_callback,
    )

    _emit_progress(
        current_step=1,
        total_steps=1,
        module_name=BUZZ_VOLUME_MODULE,
        status="started",
        progress_callback=progress_callback,
    )

    process_environment = os.environ.copy()
    process_environment["PYTHONUNBUFFERED"] = "1"

    process = subprocess.Popen(
        [
            sys.executable,
            str(module_path),
        ],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1,
        cwd=PROJECT_ROOT,
        env=process_environment,
    )

    if process.stdin is None:
        raise RuntimeError(
            "buzz_volume_adaptor.py stdin pipe를 생성하지 못했습니다."
        )

    process.stdin.write(
        f"{normalized_type}\n{reference_date_text}\n"
    )
    process.stdin.close()

    if process.stdout is not None:
        for line in process.stdout:
            _emit_log(
                line.rstrip("\r\n"),
                log_callback=log_callback,
            )

    return_code = process.wait()

    if return_code != 0:
        raise RuntimeError(
            "buzz_volume_adaptor.py 실행 실패 "
            f"(return code: {return_code})\n"
            f"Data Cut: {normalized_type}\n"
            f"기준 날짜: {reference_date_text}\n"
            f"Input Excel: {input_excel_path}\n"
            f"Expected Output: {output_excel_path}"
        )

    if not output_excel_path.is_file():
        raise RuntimeError(
            "Buzz Volume 모듈은 정상 종료되었지만 예상 완료 Excel을 찾지 못했습니다.\n"
            f"확인 경로: {output_excel_path}"
        )

    _emit_progress(
        current_step=1,
        total_steps=1,
        module_name=BUZZ_VOLUME_MODULE,
        status="completed",
        progress_callback=progress_callback,
    )

    _emit_log(
        "✅ Buzz Volume 실행 완료",
        log_callback=log_callback,
    )

    return BuzzVolumeRunResult(
        data_cut_type=normalized_type,
        reference_date=reference_date_text,
        input_excel_path=input_excel_path,
        output_excel_path=output_excel_path,
        buzz_volume_root=BUZZ_VOLUME_ROOT,
        buzz_volume_completed_dir=BUZZ_VOLUME_COMPLETED_DIR,
        overwrite_existing_result=overwrite_existing_result,
    )

# =============================================================================
# Missing Cases Service
# =============================================================================

def _load_missing_payload() -> dict:
    """누락건 payload JSON을 읽고 최상위 object 여부를 검증한다."""

    if not MISSING_PAYLOAD.is_file():
        raise FileNotFoundError(
            "누락건 payload 파일을 찾을 수 없습니다.\n"
            f"경로: {MISSING_PAYLOAD}"
        )

    try:
        with MISSING_PAYLOAD.open(
            "r",
            encoding="utf-8",
        ) as file:
            payload = json.load(file)
    except json.JSONDecodeError as exc:
        raise ValueError(
            "누락건 payload가 올바른 JSON 형식이 아닙니다.\n"
            f"경로: {MISSING_PAYLOAD}\n"
            f"오류: {exc}"
        ) from exc

    if not isinstance(payload, dict):
        raise ValueError(
            "누락건 payload의 최상위 값은 JSON object여야 합니다.\n"
            f"실제 타입: {type(payload).__name__}"
        )

    return payload


def _find_missing_query_filter(
    payload: dict,
) -> dict:
    """
    payload.filters에서 dimensionName == QUERY인 필터를 정확히 하나 찾는다.

    기존 Buzz Volume payload 처리와 동일하게 QUERY 필터는
    filterType=IN, values=[query] 계약을 사용한다.
    """

    filters = payload.get("filters")

    if not isinstance(filters, list):
        raise ValueError(
            "누락건 payload의 filters가 list가 아닙니다."
        )

    query_filters = [
        filter_item
        for filter_item in filters
        if isinstance(filter_item, dict)
        and filter_item.get("dimensionName") == "QUERY"
    ]

    if len(query_filters) != 1:
        raise ValueError(
            "누락건 payload의 QUERY 필터가 정확히 1개여야 합니다. "
            f"현재 발견된 개수: {len(query_filters)}"
        )

    query_filter = query_filters[0]

    if query_filter.get("filterType") != "IN":
        raise ValueError(
            "누락건 payload QUERY 필터의 filterType이 'IN'이 아닙니다."
        )

    return query_filter


def normalize_missing_query(
    query_text: str,
) -> str:
    """UI에서 받은 누락 Query를 원문 의미를 바꾸지 않고 앞뒤 공백만 제거한다."""

    if not isinstance(query_text, str):
        raise TypeError(
            "누락 Query는 문자열이어야 합니다."
        )

    normalized = query_text.strip()

    if not normalized:
        raise ValueError(
            "누락 Query가 비어 있습니다. "
            "처리할 Query를 입력하세요."
        )

    return normalized


def read_missing_payload_query() -> str:
    """현재 payload_6_1_누락건.json에 저장된 QUERY 값을 읽는다."""

    payload = _load_missing_payload()
    query_filter = _find_missing_query_filter(
        payload
    )
    values = query_filter.get("values")

    if not isinstance(values, list):
        raise ValueError(
            "누락건 payload QUERY 필터의 values가 list가 아닙니다."
        )

    if len(values) != 1:
        raise ValueError(
            "누락건 payload QUERY 필터의 values는 정확히 1개여야 합니다. "
            f"현재 개수: {len(values)}"
        )

    value = values[0]

    if value is None:
        return ""

    if not isinstance(value, str):
        raise ValueError(
            "누락건 payload QUERY 값이 문자열이 아닙니다. "
            f"실제 타입: {type(value).__name__}"
        )

    return value


def update_missing_payload_query(
    query_text: str,
    *,
    log_callback: LogCallback | None = None,
) -> str:
    """
    기존 수기 운영에서 하던 작업과 동일하게 QUERY 필터의 values만 교체한다.

    JSON은 임시 파일에 먼저 쓴 뒤 os.replace로 교체하여
    저장 도중 중단되어 원본 JSON이 반쯤 쓰이는 상황을 방지한다.
    """

    normalized_query = normalize_missing_query(
        query_text
    )

    payload = _load_missing_payload()
    query_filter = _find_missing_query_filter(
        payload
    )
    query_filter["values"] = [
        normalized_query
    ]

    temporary_path = MISSING_PAYLOAD.with_name(
        MISSING_PAYLOAD.name + ".tmp"
    )

    try:
        with temporary_path.open(
            "w",
            encoding="utf-8",
        ) as file:
            json.dump(
                payload,
                file,
                ensure_ascii=False,
                indent=2,
            )
            file.write("\n")

        os.replace(
            temporary_path,
            MISSING_PAYLOAD,
        )

    finally:
        if temporary_path.exists():
            try:
                temporary_path.unlink()
            except OSError:
                pass

    # 저장 직후 다시 읽어서 실제 반영 여부를 검증한다.
    saved_query = read_missing_payload_query()

    if saved_query != normalized_query:
        raise RuntimeError(
            "누락건 payload QUERY 저장 후 검증에 실패했습니다."
        )

    _emit_log(
        "✅ 누락건 payload QUERY 업데이트 완료",
        log_callback=log_callback,
    )
    _emit_log(
        f"Payload: {MISSING_PAYLOAD}",
        log_callback=log_callback,
    )
    _emit_log(
        f"QUERY: {normalized_query}",
        log_callback=log_callback,
    )

    return normalized_query


def validate_missing_cases_environment() -> tuple[
    ModuleDependencyCheck,
    ...,
]:
    """기존 `누락/` 파이프라인을 UI에서 실행하기 위한 필수 파일을 검증한다."""

    checks: list[ModuleDependencyCheck] = []

    for required_path in MISSING_REQUIRED_PATHS:
        try:
            display_path = str(
                required_path.relative_to(
                    PROJECT_ROOT
                )
            )
        except ValueError:
            display_path = str(required_path)

        checks.append(
            ModuleDependencyCheck(
                name=display_path,
                passed=required_path.is_file(),
                detail=str(required_path),
            )
        )

    # 파일이 존재할 때 payload schema도 미리 확인한다.
    if MISSING_PAYLOAD.is_file():
        try:
            current_query = read_missing_payload_query()
        except Exception as exc:
            checks.append(
                ModuleDependencyCheck(
                    name="Missing payload QUERY schema",
                    passed=False,
                    detail=(
                        f"{type(exc).__name__}: {exc}"
                    ),
                )
            )
        else:
            query_preview = current_query
            if len(query_preview) > 160:
                query_preview = (
                    query_preview[:157] + "..."
                )

            checks.append(
                ModuleDependencyCheck(
                    name="Missing payload QUERY schema",
                    passed=True,
                    detail=(
                        "QUERY filterType=IN / values=[query] 확인"
                        + (
                            f" · 현재 QUERY: {query_preview}"
                            if query_preview
                            else " · 현재 QUERY는 빈 문자열"
                        )
                    ),
                )
            )

    return tuple(checks)


def run_missing_cases_pipeline(
    start_datetime: datetime | str,
    end_datetime: datetime | str,
    query_text: str,
    *,
    log_callback: LogCallback | None = None,
    progress_callback: ProgressCallback | None = None,
    check_google_adc: bool = True,
) -> MissingCasesRunResult:
    """
    기존 수동 누락건 운영을 그대로 Web UI에서 실행한다.

    기존 수동 절차:
        1. payload_6_1_누락건.json의 QUERY values 수정
        2. cd 누락
        3. python run_pipeline.py
        4. 시작/종료 datetime 입력

    이 함수는 위 네 단계를 자동화하되 `누락/` 내부의 기존
    run_pipeline.py와 1~4 처리 모듈 자체는 수정하지 않는다.
    """

    checks = validate_missing_cases_environment()
    failed_checks = [
        check
        for check in checks
        if not check.passed
    ]

    if failed_checks:
        details = "\n".join(
            f"- {check.name}: {check.detail}"
            for check in failed_checks
        )
        raise RuntimeError(
            "누락건 실행 환경 검증에 실패했습니다.\n"
            + details
        )

    if check_google_adc:
        ensure_google_adc(
            log_callback=log_callback,
        )

    (
        _start_datetime_obj,
        end_datetime_obj,
        normalized_start_datetime,
        normalized_end_datetime,
    ) = normalize_pipeline_datetimes(
        start_datetime=start_datetime,
        end_datetime=end_datetime,
    )

    input_date = end_datetime_obj.strftime(
        "%y%m%d"
    )

    normalized_query = update_missing_payload_query(
        query_text,
        log_callback=log_callback,
    )

    output_dir = (
        MISSING_OUTPUT_ROOT
        / f"{input_date}_누락"
    )
    media_dir = (
        MISSING_MEDIA_ROOT
        / f"{input_date}_누락"
    )

    _emit_log(
        log_callback=log_callback,
    )
    _emit_log(
        "=" * 70,
        log_callback=log_callback,
    )
    _emit_log(
        "누락건 Pipeline 실행",
        log_callback=log_callback,
    )
    _emit_log(
        "기존 운영과 동일하게 `누락/` 폴더의 run_pipeline.py를 실행합니다.",
        log_callback=log_callback,
    )
    _emit_log(
        f"조회 시작: {normalized_start_datetime}",
        log_callback=log_callback,
    )
    _emit_log(
        f"조회 종료: {normalized_end_datetime}",
        log_callback=log_callback,
    )
    _emit_log(
        f"작업 날짜: {input_date}",
        log_callback=log_callback,
    )
    _emit_log(
        f"누락 Output: {output_dir}",
        log_callback=log_callback,
    )
    _emit_log(
        f"누락 Media: {media_dir}",
        log_callback=log_callback,
    )
    _emit_log(
        "=" * 70,
        log_callback=log_callback,
    )

    stdin_text = (
        normalized_start_datetime
        + "\n"
        + normalized_end_datetime
        + "\n"
    )

    environment = os.environ.copy()
    environment["PYTHONUNBUFFERED"] = "1"

    process = subprocess.Popen(
        [
            sys.executable,
            str(MISSING_RUN_PIPELINE),
        ],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1,
        cwd=MISSING_ROOT,
        env=environment,
    )

    if process.stdin is None:
        raise RuntimeError(
            "누락건 run_pipeline.py stdin pipe를 생성하지 못했습니다."
        )

    process.stdin.write(
        stdin_text
    )
    process.stdin.close()

    module_index = {
        module_name: index
        for index, module_name in enumerate(
            MISSING_PIPELINE_MODULES,
            start=1,
        )
    }
    started_modules: set[str] = set()
    completed_modules: set[str] = set()

    if process.stdout is not None:
        for raw_line in process.stdout:
            line = raw_line.rstrip(
                "\r\n"
            )

            _emit_log(
                line,
                log_callback=log_callback,
            )

            for module_name, current_step in module_index.items():
                if (
                    module_name in line
                    and "실행 시작" in line
                    and module_name not in started_modules
                ):
                    started_modules.add(
                        module_name
                    )
                    _emit_progress(
                        current_step=current_step,
                        total_steps=len(MISSING_PIPELINE_MODULES),
                        module_name=module_name,
                        status="started",
                        progress_callback=progress_callback,
                    )

                if (
                    module_name in line
                    and "실행 완료" in line
                    and module_name not in completed_modules
                ):
                    completed_modules.add(
                        module_name
                    )
                    _emit_progress(
                        current_step=current_step,
                        total_steps=len(MISSING_PIPELINE_MODULES),
                        module_name=module_name,
                        status="completed",
                        progress_callback=progress_callback,
                    )

    return_code = process.wait()

    if return_code != 0:
        raise RuntimeError(
            "누락건 run_pipeline.py 실행 실패 "
            f"(return code: {return_code})\n"
            f"작업 날짜: {input_date}\n"
            f"Payload: {MISSING_PAYLOAD}"
        )

    # 기존 run_pipeline 로그 형식이 달라 progress를 파싱하지 못했더라도,
    # 프로세스 전체가 성공했다면 네 단계는 모두 완료된 것으로 표시한다.
    for module_name, current_step in module_index.items():
        if module_name not in started_modules:
            _emit_progress(
                current_step=current_step,
                total_steps=len(MISSING_PIPELINE_MODULES),
                module_name=module_name,
                status="started",
                progress_callback=progress_callback,
            )

        if module_name not in completed_modules:
            _emit_progress(
                current_step=current_step,
                total_steps=len(MISSING_PIPELINE_MODULES),
                module_name=module_name,
                status="completed",
                progress_callback=progress_callback,
            )

    if not output_dir.is_dir():
        raise RuntimeError(
            "누락건 Pipeline은 성공 코드로 종료했지만 예상 Output 폴더를 "
            "찾지 못했습니다.\n"
            f"예상 경로: {output_dir}"
        )

    if not media_dir.is_dir():
        raise RuntimeError(
            "누락건 Pipeline은 성공 코드로 종료했지만 예상 Media 폴더를 "
            "찾지 못했습니다.\n"
            f"예상 경로: {media_dir}"
        )

    _emit_log(
        log_callback=log_callback,
    )
    _emit_log(
        "✅ 누락건 Pipeline 실행 완료",
        log_callback=log_callback,
    )
    _emit_log(
        f"Output: {output_dir}",
        log_callback=log_callback,
    )
    _emit_log(
        f"Media: {media_dir}",
        log_callback=log_callback,
    )

    return MissingCasesRunResult(
        input_date=input_date,
        normalized_start_datetime=(
            normalized_start_datetime
        ),
        normalized_end_datetime=(
            normalized_end_datetime
        ),
        query_text=normalized_query,
        payload_path=MISSING_PAYLOAD,
        missing_root=MISSING_ROOT,
        output_dir=output_dir,
        media_dir=media_dir,
    )

