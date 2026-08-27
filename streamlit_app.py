from __future__ import annotations

from datetime import datetime, timedelta
from pathlib import Path
import shutil
import time
import traceback
from typing import Any, Callable

import streamlit as st


# =============================================================================
# Page / project configuration
# =============================================================================

st.set_page_config(
    page_title="Local Campaign Automation",
    page_icon="📊",
    layout="wide",
)

PROJECT_ROOT = Path(__file__).resolve().parent

REQUIRED_PROJECT_FILES = (
    "pipeline_service.py",
    "pipeline_run_paths.py",
    "sprinklr_export_excel.py",
    "raw_to_processed.py",
    "comment_extractor.py",
    "media_extractor.py",
    "llm_analysis_pipeline.py",
)

BUZZ_VOLUME_REQUIRED_FILES = (
    "buzz_volume_adaptor.py",
    "payload/buzz_volume_base_payload.json",
)

BUZZ_VOLUME_MODULE = "buzz_volume_adaptor.py"

MODULE_LABELS = {
    "sprinklr_export_excel.py": "Sprinklr Export",
    "raw_to_processed.py": "Raw → Processed",
    "media_extractor.py": "Media Extraction",
    "llm_analysis_pipeline.py": "LLM Analysis",
}

MODULE_ORDER = tuple(MODULE_LABELS)


# pipeline_service import 자체가 실패해도 UI에서 원인을 보여주기 위해
# 앱 전체 import 실패로 끝내지 않는다.
SERVICE_IMPORT_ERROR: Exception | None = None

try:
    from pipeline_service import (
        PipelineProgress,
        build_buzz_volume_paths,
        build_expected_run_artifacts,
        list_existing_run_paths,
        read_missing_payload_query,
        run_buzz_volume,
        run_local_campaign_pipeline,
        run_missing_cases_pipeline,
        run_single_module,
        save_buzz_volume_input_file,
        validate_buzz_volume_input,
        validate_missing_cases_environment,
        validate_module_dependencies,
    )
except Exception as exc:  # pragma: no cover - UI에서 직접 표시
    SERVICE_IMPORT_ERROR = exc
    PipelineProgress = Any  # type: ignore[assignment,misc]
    build_buzz_volume_paths = None  # type: ignore[assignment]
    build_expected_run_artifacts = None  # type: ignore[assignment]
    list_existing_run_paths = None  # type: ignore[assignment]
    read_missing_payload_query = None  # type: ignore[assignment]
    run_buzz_volume = None  # type: ignore[assignment]
    run_local_campaign_pipeline = None  # type: ignore[assignment]
    run_missing_cases_pipeline = None  # type: ignore[assignment]
    run_single_module = None  # type: ignore[assignment]
    save_buzz_volume_input_file = None  # type: ignore[assignment]
    validate_buzz_volume_input = None  # type: ignore[assignment]
    validate_missing_cases_environment = None  # type: ignore[assignment]
    validate_module_dependencies = None  # type: ignore[assignment]


# =============================================================================
# Session state
# =============================================================================

def _initialize_session_state() -> None:
    now = datetime.now().replace(
        second=0,
        microsecond=0,
    )
    default_start = now - timedelta(days=1)

    defaults: dict[str, Any] = {
        # Full Pipeline
        "start_date": default_start.date(),
        "start_time": default_start.time(),
        "end_date": now.date(),
        "end_time": now.time(),
        "last_logs": [],
        "last_result": None,
        "last_error": None,
        "last_traceback": None,
        # Missing Cases
        "missing_query": "",
        "missing_query_initialized": False,
        "missing_start_date": default_start.date(),
        "missing_start_time": default_start.time(),
        "missing_end_date": now.date(),
        "missing_end_time": now.time(),
        "missing_last_logs": [],
        "missing_last_result": None,
        "missing_last_error": None,
        "missing_last_traceback": None,
        # Individual Module
        "individual_date": now.date(),
        "individual_module": "raw_to_processed.py",
        "individual_overwrite": False,
        "individual_start_date": default_start.date(),
        "individual_start_time": default_start.time(),
        "individual_end_date": now.date(),
        "individual_end_time": now.time(),
        "individual_last_logs": [],
        "individual_last_result": None,
        "individual_last_error": None,
        "individual_last_traceback": None,
        # Buzz Volume
        "buzz_cut_type": "daily",
        "buzz_reference_date": now.date(),
        "buzz_replace_input": False,
        "buzz_overwrite_result": False,
        "buzz_last_logs": [],
        "buzz_last_result": None,
        "buzz_last_error": None,
        "buzz_last_traceback": None,
    }

    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


_initialize_session_state()


# =============================================================================
# Helpers
# =============================================================================

def _combine_datetime(date_value, time_value) -> datetime:
    return datetime.combine(
        date_value,
        time_value,
    ).replace(microsecond=0)


def _format_datetime(value: datetime) -> str:
    return value.strftime(
        "%Y-%m-%d %H:%M:%S"
    )


def _check_environment() -> tuple[
    bool,
    bool,
    list[tuple[str, bool, str]],
]:
    """
    반환:
        full_pipeline_ok
        core_files_ok
        checks

    Full Pipeline은 Google ADC 확인을 수행하므로 gcloud까지 필요하고,
    Individual Module 중 LLM 이외의 모듈은 core files만 있으면 된다.
    """

    checks: list[tuple[str, bool, str]] = []
    core_results: list[bool] = []

    for file_name in REQUIRED_PROJECT_FILES:
        file_path = PROJECT_ROOT / file_name
        exists = file_path.is_file()
        core_results.append(exists)
        checks.append(
            (
                file_name,
                exists,
                str(file_path),
            )
        )

    service_ok = SERVICE_IMPORT_ERROR is None
    core_results.append(service_ok)

    if SERVICE_IMPORT_ERROR is not None:
        checks.append(
            (
                "pipeline_service import",
                False,
                (
                    f"{type(SERVICE_IMPORT_ERROR).__name__}: "
                    f"{SERVICE_IMPORT_ERROR}"
                ),
            )
        )
    else:
        checks.append(
            (
                "pipeline_service import",
                True,
                "정상",
            )
        )

    gcloud_path = shutil.which("gcloud")
    gcloud_ok = gcloud_path is not None
    checks.append(
        (
            "Google Cloud CLI (gcloud)",
            gcloud_ok,
            gcloud_path or "PATH에서 찾지 못함",
        )
    )

    core_files_ok = all(core_results)
    full_pipeline_ok = core_files_ok and gcloud_ok

    return (
        full_pipeline_ok,
        core_files_ok,
        checks,
    )


def _render_environment_check(
    checks: list[tuple[str, bool, str]],
) -> None:
    with st.expander(
        "System Check",
        expanded=False,
    ):
        st.caption(
            "프로젝트 핵심 파일과 Google Cloud CLI 상태를 확인합니다. "
            "LLM Analysis 및 Full Pipeline에는 gcloud가 필요합니다."
        )

        for name, passed, detail in checks:
            icon = "✅" if passed else "❌"
            st.markdown(
                f"{icon} **{name}**  \n"
                f"`{detail}`"
            )


def _step_markdown(
    step_states: dict[str, str],
) -> str:
    icons = {
        "pending": "○",
        "running": "🔄",
        "complete": "✅",
        "error": "❌",
    }

    lines: list[str] = []

    for index, module_name in enumerate(
        MODULE_ORDER,
        start=1,
    ):
        state = step_states.get(
            module_name,
            "pending",
        )
        icon = icons.get(state, "○")
        label = MODULE_LABELS[module_name]
        lines.append(
            f"{icon} **{index}. {label}**"
        )

    return "  \n".join(lines)


def _find_excel_results(
    output_dir: Path,
) -> list[Path]:
    if not output_dir.is_dir():
        return []

    files = [
        path
        for path in output_dir.iterdir()
        if (
            path.is_file()
            and path.suffix.casefold()
            in {".xlsx", ".xlsm"}
        )
    ]

    return sorted(
        files,
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )


def _render_excel_downloads(
    output_dir: Path,
    *,
    key_prefix: str,
) -> None:
    excel_files = _find_excel_results(
        output_dir
    )

    if not excel_files:
        st.info(
            "현재 Output 폴더에서 다운로드 가능한 Excel 파일을 찾지 못했습니다."
        )
        return

    st.markdown("**Output Excel 다운로드**")

    file_names = [
        path.name
        for path in excel_files
    ]

    selected_name = st.selectbox(
        "다운로드할 파일",
        options=file_names,
        key=f"{key_prefix}_excel_selection",
    )

    selected_path = next(
        path
        for path in excel_files
        if path.name == selected_name
    )

    try:
        file_bytes = selected_path.read_bytes()
    except OSError as exc:
        st.error(
            "파일을 읽지 못했습니다: "
            f"{exc}"
        )
        return

    st.download_button(
        label="Download Excel",
        data=file_bytes,
        file_name=selected_path.name,
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        key=f"{key_prefix}_download_button",
    )


def _result_to_state(result) -> dict[str, str]:
    return {
        "input_date": result.input_date,
        "start_datetime": result.normalized_start_datetime,
        "end_datetime": result.normalized_end_datetime,
        "run_label": result.run_paths.run_label,
        "output_dir": str(result.run_paths.output_dir),
        "media_dir": str(result.run_paths.media_dir),
        "buzz_volume_root": str(result.buzz_volume_root),
        "buzz_volume_completed_dir": str(
            result.buzz_volume_completed_dir
        ),
    }


def _single_result_to_state(result) -> dict[str, str]:
    return {
        "module_name": result.module_name,
        "module_label": MODULE_LABELS.get(
            result.module_name,
            result.module_name,
        ),
        "input_date": result.input_date,
        "run_label": result.run_paths.run_label,
        "output_dir": str(result.run_paths.output_dir),
        "media_dir": str(result.run_paths.media_dir),
        "overwrite_existing": str(
            result.overwrite_existing
        ),
        "start_datetime": (
            result.normalized_start_datetime
            or ""
        ),
        "end_datetime": (
            result.normalized_end_datetime
            or ""
        ),
    }


def _render_result(
    result_state: dict[str, str],
) -> None:
    st.subheader("Result")

    col1, col2, col3 = st.columns(3)

    col1.metric(
        "작업 날짜",
        result_state["input_date"],
    )
    col2.metric(
        "실행 차수",
        result_state["run_label"],
    )
    col3.metric(
        "상태",
        "Completed",
    )

    st.markdown("**Output Folder**")
    st.code(
        result_state["output_dir"],
        language="text",
    )

    st.markdown("**Media Folder**")
    st.code(
        result_state["media_dir"],
        language="text",
    )

    with st.expander(
        "Buzz Volume 경로",
        expanded=False,
    ):
        st.markdown("**입력 폴더**")
        st.code(
            result_state["buzz_volume_root"],
            language="text",
        )
        st.markdown("**완료 폴더**")
        st.code(
            result_state[
                "buzz_volume_completed_dir"
            ],
            language="text",
        )

    _render_excel_downloads(
        Path(result_state["output_dir"]),
        key_prefix="full_result",
    )


def _render_single_result(
    result_state: dict[str, str],
) -> None:
    st.subheader("Individual Module Result")

    col1, col2, col3 = st.columns(3)
    col1.metric(
        "Module",
        result_state["module_label"],
    )
    col2.metric(
        "Run",
        result_state["run_label"],
    )
    col3.metric(
        "Status",
        "Completed",
    )

    st.markdown("**Output Folder**")
    st.code(
        result_state["output_dir"],
        language="text",
    )

    st.markdown("**Media Folder**")
    st.code(
        result_state["media_dir"],
        language="text",
    )

    _render_excel_downloads(
        Path(result_state["output_dir"]),
        key_prefix="individual_result",
    )


def _create_live_log_callback(
    log_placeholder,
    logs: list[str],
) -> Callable[[str], None]:
    last_log_render_time = [0.0]

    def log_callback(
        message: str,
    ) -> None:
        logs.append(message)

        if len(logs) > 5_000:
            del logs[:-5_000]

        now = time.monotonic()
        should_render = (
            now
            - last_log_render_time[0]
            >= 0.25
            or message.startswith("✅")
            or "실행 실패" in message
            or "ERROR" in message
        )

        if not should_render:
            return

        log_placeholder.code(
            "\n".join(logs[-300:]),
            language="text",
        )
        last_log_render_time[0] = now

    return log_callback


def _render_dependency_checks(
    checks,
) -> bool:
    if not checks:
        st.warning(
            "선행조건 검사 결과를 가져오지 못했습니다."
        )
        return False

    all_passed = True

    st.markdown("**Dependencies**")

    for check in checks:
        icon = "✅" if check.passed else "❌"
        st.markdown(
            f"{icon} **{check.name}**  \n"
            f"`{check.detail}`"
        )
        all_passed = all_passed and check.passed

    return all_passed



# =============================================================================
# Buzz Volume helpers / tab
# =============================================================================

def _check_buzz_environment() -> tuple[
    bool,
    list[tuple[str, bool, str]],
]:
    checks: list[tuple[str, bool, str]] = []

    for relative_path in BUZZ_VOLUME_REQUIRED_FILES:
        file_path = PROJECT_ROOT / relative_path
        checks.append(
            (
                relative_path,
                file_path.is_file(),
                str(file_path),
            )
        )

    service_functions_ok = all(
        function is not None
        for function in (
            build_buzz_volume_paths,
            run_buzz_volume,
            save_buzz_volume_input_file,
            validate_buzz_volume_input,
        )
    )
    checks.append(
        (
            "Buzz Volume service functions",
            service_functions_ok,
            "정상" if service_functions_ok else "pipeline_service import 실패 또는 함수 누락",
        )
    )

    return (
        all(passed for _, passed, _ in checks),
        checks,
    )


def _buzz_result_to_state(result) -> dict[str, str]:
    return {
        "data_cut_type": result.data_cut_type,
        "reference_date": result.reference_date,
        "input_excel_path": str(result.input_excel_path),
        "output_excel_path": str(result.output_excel_path),
        "buzz_volume_root": str(result.buzz_volume_root),
        "buzz_volume_completed_dir": str(
            result.buzz_volume_completed_dir
        ),
        "overwrite_existing_result": str(
            result.overwrite_existing_result
        ),
    }


def _render_buzz_result(
    result_state: dict[str, str],
) -> None:
    st.subheader("Buzz Volume Result")

    col1, col2, col3 = st.columns(3)
    col1.metric(
        "Data Cut",
        result_state["data_cut_type"].upper(),
    )
    col2.metric(
        "Reference Date",
        result_state["reference_date"],
    )
    col3.metric(
        "상태",
        "Completed",
    )

    st.markdown("**Input Excel**")
    st.code(
        result_state["input_excel_path"],
        language="text",
    )

    st.markdown("**Completed Excel**")
    st.code(
        result_state["output_excel_path"],
        language="text",
    )

    output_path = Path(
        result_state["output_excel_path"]
    )

    if not output_path.is_file():
        st.warning(
            "완료 Excel을 현재 경로에서 찾지 못했습니다."
        )
        return

    try:
        file_bytes = output_path.read_bytes()
    except OSError as exc:
        st.error(
            f"완료 Excel을 읽지 못했습니다: {exc}"
        )
        return

    st.download_button(
        label="Download Buzz Volume Excel",
        data=file_bytes,
        file_name=output_path.name,
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        key="buzz_volume_result_download",
    )


def _render_buzz_volume_tab() -> None:
    st.subheader("Buzz Volume")
    st.write(
        "1차·2차 등 여러 실행 결과를 통합하고 수동 정제한 최종 Excel에 "
        "Sprinklr Mentions(Buzz Volume)을 채웁니다."
    )
    st.info(
        "Buzz Volume은 Full Pipeline 1~4단계와 별도 실행입니다. "
        "최종 통합·정제 Excel을 업로드하거나 output/Buzz_Volume에 미리 준비하세요."
    )

    buzz_environment_ok, buzz_checks = (
        _check_buzz_environment()
    )

    with st.expander(
        "Buzz Volume System Check",
        expanded=False,
    ):
        for name, passed, detail in buzz_checks:
            icon = "✅" if passed else "❌"
            st.markdown(
                f"{icon} **{name}**  \n`{detail}`"
            )

    if not buzz_environment_ok:
        st.error(
            "Buzz Volume 실행에 필요한 파일이 준비되지 않았습니다. "
            "위 ❌ 항목을 먼저 해결하세요."
        )

    control_col1, control_col2 = st.columns(2)

    with control_col1:
        data_cut_type = st.radio(
            "Data Cut",
            options=["daily", "weekly"],
            horizontal=True,
            key="buzz_cut_type",
        )

    with control_col2:
        reference_date = st.date_input(
            "Reference Date",
            key="buzz_reference_date",
            format="YYYY-MM-DD",
        )

    date_rule_ok = True
    date_rule_message = ""

    if data_cut_type == "weekly" and reference_date.weekday() != 0:
        date_rule_ok = False
        date_rule_message = (
            "Weekly 기준 날짜는 반드시 월요일이어야 합니다."
        )
    elif (
        data_cut_type == "daily"
        and reference_date < datetime(2026, 7, 1).date()
    ):
        date_rule_ok = False
        date_rule_message = (
            "Daily 기준 날짜는 2026-07-01보다 이전일 수 없습니다."
        )

    if date_rule_ok:
        st.success(
            "Data Cut / Reference Date 검증 완료"
        )
    else:
        st.error(date_rule_message)

    expected_input_path: Path | None = None
    expected_output_path: Path | None = None

    if build_buzz_volume_paths is not None:
        try:
            (
                expected_input_path,
                expected_output_path,
            ) = build_buzz_volume_paths(
                reference_date
            )
        except Exception as exc:
            st.error(
                "Buzz Volume 경로 계산 실패: "
                f"{type(exc).__name__}: {exc}"
            )

    if expected_input_path is not None:
        with st.expander(
            "Expected File Paths",
            expanded=False,
        ):
            st.markdown("**Input**")
            st.code(
                str(expected_input_path),
                language="text",
            )
            st.markdown("**Completed Output**")
            st.code(
                str(expected_output_path),
                language="text",
            )

    st.markdown("### Final Excel")
    uploaded_file = st.file_uploader(
        "최종 통합·정제 Excel (.xlsx)",
        type=["xlsx"],
        key="buzz_volume_upload",
        help=(
            "업로드한 파일명과 관계없이 Reference Date에 맞는 canonical 파일명으로 "
            "output/Buzz_Volume에 저장됩니다."
        ),
    )

    existing_input = bool(
        expected_input_path is not None
        and expected_input_path.is_file()
    )
    existing_output = bool(
        expected_output_path is not None
        and expected_output_path.is_file()
    )

    if existing_input:
        st.success(
            "기존 Buzz Volume 입력 Excel이 준비되어 있습니다."
        )
    elif uploaded_file is not None:
        st.success(
            "업로드 파일이 준비되었습니다. Run 시 canonical 파일명으로 저장합니다."
        )
    else:
        st.warning(
            "Buzz Volume 입력 Excel이 없습니다. 파일을 업로드하세요."
        )

    replace_input = st.checkbox(
        "업로드 파일로 기존 Buzz Volume 입력 Excel 교체",
        key="buzz_replace_input",
        disabled=(
            uploaded_file is None
            or not existing_input
        ),
    )

    overwrite_result = st.checkbox(
        "기존 completed 결과 덮어쓰기 허용",
        key="buzz_overwrite_result",
        disabled=not existing_output,
    )

    if existing_output and not overwrite_result:
        st.info(
            "동일 기준 날짜의 completed 결과가 이미 있습니다. "
            "재실행하려면 덮어쓰기를 허용하세요."
        )
    elif existing_output and overwrite_result:
        st.warning(
            "기존 completed 결과가 재실행 결과로 갱신됩니다."
        )

    # Service dependency check. 업로드 파일이 있으면 input 파일은 Run 직전에 저장되므로
    # require_input_file=False로 검사한다.
    service_checks = ()
    service_checks_ok = False

    if validate_buzz_volume_input is not None and date_rule_ok:
        try:
            service_checks = validate_buzz_volume_input(
                data_cut_type=data_cut_type,
                reference_date=reference_date,
                require_input_file=(uploaded_file is None),
            )
        except Exception as exc:
            st.error(
                "Buzz Volume 선행조건 검사 실패: "
                f"{type(exc).__name__}: {exc}"
            )
        else:
            with st.expander(
                "Execution Validation",
                expanded=True,
            ):
                for check in service_checks:
                    icon = "✅" if check.passed else "❌"
                    st.markdown(
                        f"{icon} **{check.name}**  \n`{check.detail}`"
                    )
            service_checks_ok = all(
                check.passed
                for check in service_checks
            )

    input_ready = (
        existing_input
        or uploaded_file is not None
    )

    upload_conflict = (
        uploaded_file is not None
        and existing_input
        and not replace_input
    )

    completed_conflict = (
        existing_output
        and not overwrite_result
    )

    run_disabled = (
        not buzz_environment_ok
        or not date_rule_ok
        or not service_checks_ok
        or not input_ready
        or upload_conflict
        or completed_conflict
        or run_buzz_volume is None
    )

    if upload_conflict:
        st.error(
            "동일 기준 날짜의 입력 Excel이 이미 있습니다. "
            "새 업로드 파일을 사용하려면 입력 Excel 교체를 허용하세요."
        )

    run_clicked = st.button(
        "▶ Fill Buzz Volume",
        type="primary",
        disabled=run_disabled,
        key="run_buzz_volume_button",
    )

    if run_clicked:
        st.session_state["buzz_last_result"] = None
        st.session_state["buzz_last_error"] = None
        st.session_state["buzz_last_traceback"] = None
        st.session_state["buzz_last_logs"] = []

        if uploaded_file is not None:
            if save_buzz_volume_input_file is None:
                st.error(
                    "Buzz Volume 업로드 저장 함수를 불러오지 못했습니다."
                )
                return

            try:
                saved_path = save_buzz_volume_input_file(
                    uploaded_file.getvalue(),
                    reference_date,
                    overwrite_existing=(
                        replace_input
                        or not existing_input
                    ),
                )
            except Exception as exc:
                st.error(
                    "Buzz Volume 입력 Excel 저장 실패: "
                    f"{type(exc).__name__}: {exc}"
                )
                return
            else:
                st.success(
                    f"입력 Excel 저장 완료: {saved_path.name}"
                )

        progress_bar = st.progress(
            0,
            text="Buzz Volume 준비 중...",
        )
        status_box = st.status(
            "Buzz Volume 준비 중...",
            expanded=True,
            state="running",
        )

        st.markdown("**Execution Log**")
        log_placeholder = st.empty()
        log_placeholder.code(
            "로그 수집을 시작합니다...",
            language="text",
        )

        logs: list[str] = []
        log_callback = _create_live_log_callback(
            log_placeholder,
            logs,
        )

        def progress_callback(
            progress: PipelineProgress,
        ) -> None:
            if progress.status == "started":
                status_box.update(
                    label="Buzz Volume 실행 중",
                    state="running",
                    expanded=True,
                )

            percent = int(
                round(progress.fraction * 100)
            )
            progress_bar.progress(
                percent,
                text=f"Buzz Volume · {percent}%",
            )

        try:
            result = run_buzz_volume(
                data_cut_type=data_cut_type,
                reference_date=reference_date,
                overwrite_existing_result=overwrite_result,
                log_callback=log_callback,
                progress_callback=progress_callback,
            )

        except Exception as exc:
            error_message = (
                f"{type(exc).__name__}: {exc}"
            )

            logs.append("")
            logs.append(
                "[BUZZ VOLUME ERROR] "
                + error_message
            )

            log_placeholder.code(
                "\n".join(logs[-300:]),
                language="text",
            )

            status_box.update(
                label="Buzz Volume 실행 실패",
                state="error",
                expanded=True,
            )

            st.session_state["buzz_last_logs"] = logs
            st.session_state["buzz_last_error"] = error_message
            st.session_state["buzz_last_traceback"] = traceback.format_exc()

            st.error(
                "Buzz Volume 실행 중 오류가 발생했습니다. "
                "Execution Log를 확인하세요."
            )
            st.code(
                error_message,
                language="text",
            )

            with st.expander(
                "상세 Traceback",
                expanded=False,
            ):
                st.code(
                    st.session_state[
                        "buzz_last_traceback"
                    ],
                    language="text",
                )

        else:
            progress_bar.progress(
                100,
                text="Buzz Volume 완료 · 100%",
            )
            log_placeholder.code(
                "\n".join(logs[-300:]),
                language="text",
            )

            status_box.update(
                label="Buzz Volume 실행 완료",
                state="complete",
                expanded=False,
            )

            result_state = _buzz_result_to_state(
                result
            )
            st.session_state["buzz_last_logs"] = logs
            st.session_state["buzz_last_result"] = result_state

            st.success(
                "Buzz Volume 적재가 정상 완료되었습니다."
            )
            _render_buzz_result(
                result_state
            )

    else:
        last_result = st.session_state.get(
            "buzz_last_result"
        )
        last_error = st.session_state.get(
            "buzz_last_error"
        )
        last_logs = st.session_state.get(
            "buzz_last_logs",
            [],
        )

        if last_result:
            st.divider()
            _render_buzz_result(
                last_result
            )

        if last_error:
            st.divider()
            st.subheader("Last Buzz Volume Error")
            st.error(last_error)

        if last_logs:
            with st.expander(
                "Last Buzz Volume Log",
                expanded=False,
            ):
                st.code(
                    "\n".join(last_logs[-300:]),
                    language="text",
                )




# =============================================================================
# Missing Cases helpers / tab
# =============================================================================

def _missing_result_to_state(result) -> dict[str, str]:
    return {
        "input_date": result.input_date,
        "normalized_start_datetime": result.normalized_start_datetime,
        "normalized_end_datetime": result.normalized_end_datetime,
        "query_text": result.query_text,
        "payload_path": str(result.payload_path),
        "missing_root": str(result.missing_root),
        "output_dir": str(result.output_dir),
        "media_dir": str(result.media_dir),
    }


def _render_missing_result(
    result_state: dict[str, str],
) -> None:
    st.subheader("Missing Cases Result")

    col1, col2, col3 = st.columns(3)
    col1.metric(
        "Work Date",
        result_state["input_date"],
    )
    col2.metric(
        "Mode",
        "Missing Cases",
    )
    col3.metric(
        "Status",
        "Completed",
    )

    st.markdown("**Payload QUERY**")
    st.code(
        result_state["query_text"],
        language="text",
    )

    st.markdown("**Payload File**")
    st.code(
        result_state["payload_path"],
        language="text",
    )

    output_col, media_col = st.columns(2)

    with output_col:
        st.markdown("**Missing Output Folder**")
        st.code(
            result_state["output_dir"],
            language="text",
        )

    with media_col:
        st.markdown("**Missing Media Folder**")
        st.code(
            result_state["media_dir"],
            language="text",
        )

    _render_excel_downloads(
        Path(result_state["output_dir"]),
        key_prefix="missing_result",
    )


def _render_missing_cases_tab() -> None:
    st.subheader("Missing Cases")
    st.write(
        "기존 누락건 운영 절차를 그대로 Web UI에서 실행합니다. "
        "QUERY를 입력하면 payload_6_1_누락건.json의 QUERY values만 교체한 뒤 "
        "`누락/run_pipeline.py`를 실행합니다."
    )
    st.info(
        "기존처럼 별도로 `cd 누락`하거나 JSON을 직접 열어 수정할 필요가 없습니다. "
        "누락 폴더 내부의 1~4 처리 모듈은 그대로 사용합니다."
    )

    if validate_missing_cases_environment is None:
        st.error(
            "Missing Cases 서비스 함수를 불러오지 못했습니다. "
            "pipeline_service.py가 최신 버전인지 확인하세요."
        )
        return

    try:
        checks = validate_missing_cases_environment()
    except Exception as exc:
        st.error(
            "Missing Cases System Check 중 오류가 발생했습니다."
        )
        st.code(
            f"{type(exc).__name__}: {exc}",
            language="text",
        )
        return

    with st.expander(
        "Missing Cases System Check",
        expanded=False,
    ):
        missing_environment_ok = _render_dependency_checks(
            checks
        )

    if not missing_environment_ok:
        st.error(
            "누락건 실행에 필요한 파일 또는 payload 구조가 준비되지 않았습니다. "
            "위 ❌ 항목을 먼저 해결하세요."
        )

    # 최초 한 번만 현재 JSON의 QUERY를 text_area 기본값으로 읽는다.
    if (
        not st.session_state.get(
            "missing_query_initialized",
            False,
        )
        and read_missing_payload_query is not None
    ):
        try:
            current_query = read_missing_payload_query()
        except Exception:
            current_query = ""

        st.session_state["missing_query"] = current_query
        st.session_state["missing_query_initialized"] = True

    with st.form(
        "missing_cases_form",
        clear_on_submit=False,
    ):
        query_text = st.text_area(
            "Missing Query",
            key="missing_query",
            height=140,
            placeholder=(
                "예: Query_A OR Query_B OR Query_C\n"
                "기존에 payload에 넣던 최종 Query 문자열을 그대로 입력하세요."
            ),
        )

        (
            start_date_col,
            start_time_col,
            end_date_col,
            end_time_col,
        ) = st.columns(4)

        with start_date_col:
            start_date = st.date_input(
                "Start Date",
                key="missing_start_date",
                format="YYYY-MM-DD",
            )

        with start_time_col:
            start_time = st.time_input(
                "Start Time",
                key="missing_start_time",
                step=timedelta(minutes=1),
            )

        with end_date_col:
            end_date = st.date_input(
                "End Date",
                key="missing_end_date",
                format="YYYY-MM-DD",
            )

        with end_time_col:
            end_time = st.time_input(
                "End Time",
                key="missing_end_time",
                step=timedelta(minutes=1),
            )

        submitted = st.form_submit_button(
            "▶ Run Missing Cases Pipeline",
            type="primary",
            disabled=(
                not missing_environment_ok
                or run_missing_cases_pipeline is None
            ),
        )

    if (
        start_date is not None
        and start_time is not None
        and end_date is not None
        and end_time is not None
    ):
        preview_start = _combine_datetime(
            start_date,
            start_time,
        )
        preview_end = _combine_datetime(
            end_date,
            end_time,
        )
        input_date = preview_end.strftime(
            "%y%m%d"
        )

        predicted_output = (
            PROJECT_ROOT
            / "누락"
            / "output_누락"
            / f"{input_date}_누락"
        )
        predicted_media = (
            PROJECT_ROOT
            / "누락"
            / "media_누락"
            / f"{input_date}_누락"
        )

        with st.expander(
            "실행값 확인",
            expanded=False,
        ):
            st.write(
                "조회 시작:",
                _format_datetime(preview_start),
            )
            st.write(
                "조회 종료:",
                _format_datetime(preview_end),
            )
            st.write(
                "작업 기준 날짜:",
                input_date,
            )
            st.markdown("**예상 Output**")
            st.code(
                str(predicted_output),
                language="text",
            )
            st.markdown("**예상 Media**")
            st.code(
                str(predicted_media),
                language="text",
            )

        if predicted_output.exists() or predicted_media.exists():
            st.warning(
                "같은 작업 날짜의 누락 Output/Media 폴더가 이미 존재합니다. "
                "기존 누락 run_pipeline.py의 기존 결과 보호/덮어쓰기 규칙이 그대로 적용됩니다."
            )

    if submitted:
        st.session_state["missing_last_result"] = None
        st.session_state["missing_last_error"] = None
        st.session_state["missing_last_traceback"] = None
        st.session_state["missing_last_logs"] = []

        if not query_text.strip():
            st.error(
                "Missing Query를 입력하세요."
            )
            return

        if (
            start_date is None
            or start_time is None
            or end_date is None
            or end_time is None
        ):
            st.error(
                "시작/종료 날짜와 시간을 모두 입력하세요."
            )
            return

        start_datetime = _combine_datetime(
            start_date,
            start_time,
        )
        end_datetime = _combine_datetime(
            end_date,
            end_time,
        )

        if end_datetime <= start_datetime:
            st.error(
                "종료 날짜와 시간은 시작 날짜와 시간보다 늦어야 합니다."
            )
            return

        progress_bar = st.progress(
            0,
            text="Missing Cases Pipeline 준비 중...",
        )

        step_states = {
            module_name: "pending"
            for module_name in MODULE_ORDER
        }

        steps_placeholder = st.empty()
        steps_placeholder.markdown(
            _step_markdown(step_states)
        )

        status_box = st.status(
            "Missing Cases Pipeline 준비 중...",
            expanded=True,
            state="running",
        )

        st.markdown("**Execution Log**")
        log_placeholder = st.empty()
        log_placeholder.code(
            "로그 수집을 시작합니다...",
            language="text",
        )

        logs: list[str] = []
        log_callback = _create_live_log_callback(
            log_placeholder,
            logs,
        )

        def progress_callback(
            progress: PipelineProgress,
        ) -> None:
            module_name = progress.module_name
            label = MODULE_LABELS.get(
                module_name,
                module_name,
            )

            if progress.status == "started":
                step_states[module_name] = "running"
                status_box.update(
                    label=(
                        f"{progress.current_step}/"
                        f"{progress.total_steps} · {label} 실행 중"
                    ),
                    state="running",
                    expanded=True,
                )

            elif progress.status == "completed":
                step_states[module_name] = "complete"

            steps_placeholder.markdown(
                _step_markdown(step_states)
            )

            percent = int(
                round(
                    progress.fraction * 100
                )
            )
            progress_bar.progress(
                percent,
                text=(
                    f"Missing Cases · {label} · {percent}%"
                ),
            )

        try:
            result = run_missing_cases_pipeline(
                start_datetime=start_datetime,
                end_datetime=end_datetime,
                query_text=query_text,
                log_callback=log_callback,
                progress_callback=progress_callback,
            )

        except Exception as exc:
            for module_name in MODULE_ORDER:
                if step_states[module_name] == "running":
                    step_states[module_name] = "error"
                    break

            steps_placeholder.markdown(
                _step_markdown(step_states)
            )

            error_message = (
                f"{type(exc).__name__}: {exc}"
            )
            logs.append("")
            logs.append(
                "[MISSING PIPELINE ERROR] "
                + error_message
            )
            log_placeholder.code(
                "\n".join(logs[-300:]),
                language="text",
            )
            status_box.update(
                label="Missing Cases Pipeline 실행 실패",
                state="error",
                expanded=True,
            )

            st.session_state["missing_last_logs"] = logs
            st.session_state["missing_last_error"] = error_message
            st.session_state["missing_last_traceback"] = traceback.format_exc()

            st.error(
                "누락건 Pipeline 실행 중 오류가 발생했습니다. "
                "아래 오류 정보와 Execution Log를 확인하세요."
            )
            st.code(
                error_message,
                language="text",
            )

            with st.expander(
                "상세 Traceback",
                expanded=False,
            ):
                st.code(
                    st.session_state["missing_last_traceback"],
                    language="text",
                )

        else:
            progress_bar.progress(
                100,
                text="Missing Cases Pipeline 완료 · 100%",
            )
            steps_placeholder.markdown(
                _step_markdown(step_states)
            )
            log_placeholder.code(
                "\n".join(logs[-300:]),
                language="text",
            )
            status_box.update(
                label="Missing Cases Pipeline 실행 완료",
                state="complete",
                expanded=False,
            )

            result_state = _missing_result_to_state(
                result
            )
            st.session_state["missing_last_logs"] = logs
            st.session_state["missing_last_result"] = result_state
            st.session_state["missing_last_error"] = None
            st.session_state["missing_last_traceback"] = None

            st.success(
                "누락건 Pipeline이 정상 완료되었습니다."
            )
            _render_missing_result(
                result_state
            )

    if not submitted:
        last_result = st.session_state.get(
            "missing_last_result"
        )
        last_error = st.session_state.get(
            "missing_last_error"
        )
        last_logs = st.session_state.get(
            "missing_last_logs",
            [],
        )

        if last_result:
            st.divider()
            _render_missing_result(
                last_result
            )

        if last_error:
            st.divider()
            st.subheader("Last Missing Cases Error")
            st.error(last_error)

        if last_logs:
            with st.expander(
                "Last Missing Cases Execution Log",
                expanded=False,
            ):
                st.code(
                    "\n".join(last_logs[-300:]),
                    language="text",
                )


# =============================================================================
# Full Pipeline tab
# =============================================================================

def _render_full_pipeline_tab(
    *,
    full_pipeline_ok: bool,
) -> None:
    st.subheader("Full Pipeline")
    st.write(
        "조회 시작/종료 시간을 선택한 뒤 버튼을 누르면 1~4단계를 순서대로 실행합니다."
    )
    st.info(
        "실행 도중 터미널 입력을 받을 수 없습니다. "
        "Social Login이 필요한 환경이라면 로그인 세션을 먼저 준비한 뒤 실행하세요."
    )

    if not full_pipeline_ok:
        st.error(
            "Full Pipeline 실행 환경이 준비되지 않았습니다. "
            "System Check의 ❌ 항목을 먼저 해결하세요."
        )

    with st.form(
        "full_pipeline_form",
        clear_on_submit=False,
    ):
        (
            start_date_col,
            start_time_col,
            end_date_col,
            end_time_col,
        ) = st.columns(4)

        with start_date_col:
            start_date = st.date_input(
                "Start Date",
                key="start_date",
                format="YYYY-MM-DD",
            )

        with start_time_col:
            start_time = st.time_input(
                "Start Time",
                key="start_time",
                step=timedelta(minutes=1),
            )

        with end_date_col:
            end_date = st.date_input(
                "End Date",
                key="end_date",
                format="YYYY-MM-DD",
            )

        with end_time_col:
            end_time = st.time_input(
                "End Time",
                key="end_time",
                step=timedelta(minutes=1),
            )

        submitted = st.form_submit_button(
            "▶ Run Full Pipeline",
            type="primary",
            disabled=not full_pipeline_ok,
        )

    if (
        start_date is not None
        and start_time is not None
        and end_date is not None
        and end_time is not None
    ):
        preview_start = _combine_datetime(
            start_date,
            start_time,
        )
        preview_end = _combine_datetime(
            end_date,
            end_time,
        )

        with st.expander(
            "입력값 확인",
            expanded=False,
        ):
            st.write(
                "조회 시작:",
                _format_datetime(preview_start),
            )
            st.write(
                "조회 종료:",
                _format_datetime(preview_end),
            )
            st.write(
                "작업 기준 날짜:",
                preview_end.strftime("%y%m%d"),
            )

    if submitted:
        st.session_state["last_result"] = None
        st.session_state["last_error"] = None
        st.session_state["last_traceback"] = None
        st.session_state["last_logs"] = []

        if (
            start_date is None
            or start_time is None
            or end_date is None
            or end_time is None
        ):
            st.error(
                "시작/종료 날짜와 시간을 모두 입력하세요."
            )
            return

        start_datetime = _combine_datetime(
            start_date,
            start_time,
        )
        end_datetime = _combine_datetime(
            end_date,
            end_time,
        )

        if end_datetime <= start_datetime:
            st.error(
                "종료 날짜와 시간은 시작 날짜와 시간보다 늦어야 합니다."
            )
            return

        if run_local_campaign_pipeline is None:
            st.error(
                "pipeline_service를 불러오지 못해 실행할 수 없습니다."
            )
            return

        progress_bar = st.progress(
            0,
            text="Pipeline 준비 중...",
        )

        step_states = {
            module_name: "pending"
            for module_name in MODULE_ORDER
        }

        steps_placeholder = st.empty()
        steps_placeholder.markdown(
            _step_markdown(step_states)
        )

        status_box = st.status(
            "Pipeline 준비 중...",
            expanded=True,
            state="running",
        )

        st.markdown("**Execution Log**")
        log_placeholder = st.empty()
        log_placeholder.code(
            "로그 수집을 시작합니다...",
            language="text",
        )

        logs: list[str] = []
        log_callback = _create_live_log_callback(
            log_placeholder,
            logs,
        )

        def progress_callback(
            progress: PipelineProgress,
        ) -> None:
            module_name = progress.module_name
            label = MODULE_LABELS.get(
                module_name,
                module_name,
            )

            if progress.status == "started":
                step_states[module_name] = "running"
                status_box.update(
                    label=(
                        f"{progress.current_step}/"
                        f"{progress.total_steps} · {label} 실행 중"
                    ),
                    state="running",
                    expanded=True,
                )

            elif progress.status == "completed":
                step_states[module_name] = "complete"

            steps_placeholder.markdown(
                _step_markdown(step_states)
            )

            percent = int(
                round(
                    progress.fraction * 100
                )
            )

            progress_bar.progress(
                percent,
                text=f"{label} · {percent}%",
            )

        try:
            result = run_local_campaign_pipeline(
                start_datetime=start_datetime,
                end_datetime=end_datetime,
                log_callback=log_callback,
                progress_callback=progress_callback,
            )

        except Exception as exc:
            for module_name in MODULE_ORDER:
                if step_states[module_name] == "running":
                    step_states[module_name] = "error"
                    break

            steps_placeholder.markdown(
                _step_markdown(step_states)
            )

            error_message = (
                f"{type(exc).__name__}: {exc}"
            )

            logs.append("")
            logs.append(
                "[PIPELINE ERROR] "
                + error_message
            )

            log_placeholder.code(
                "\n".join(logs[-300:]),
                language="text",
            )

            status_box.update(
                label="Pipeline 실행 실패",
                state="error",
                expanded=True,
            )

            st.session_state["last_logs"] = logs
            st.session_state["last_error"] = error_message
            st.session_state["last_traceback"] = traceback.format_exc()

            st.error(
                "Pipeline 실행 중 오류가 발생했습니다. "
                "아래 오류 정보와 Execution Log를 확인하세요."
            )
            st.code(
                error_message,
                language="text",
            )

            with st.expander(
                "상세 Traceback",
                expanded=False,
            ):
                st.code(
                    st.session_state[
                        "last_traceback"
                    ],
                    language="text",
                )

        else:
            progress_bar.progress(
                100,
                text="Pipeline 완료 · 100%",
            )
            steps_placeholder.markdown(
                _step_markdown(step_states)
            )
            log_placeholder.code(
                "\n".join(logs[-300:]),
                language="text",
            )

            status_box.update(
                label="Pipeline 실행 완료",
                state="complete",
                expanded=False,
            )

            result_state = _result_to_state(
                result
            )

            st.session_state["last_logs"] = logs
            st.session_state["last_result"] = result_state

            st.success(
                "전체 Local Campaign Pipeline이 정상 완료되었습니다."
            )
            _render_result(
                result_state
            )

    else:
        last_result = st.session_state.get(
            "last_result"
        )
        last_error = st.session_state.get(
            "last_error"
        )
        last_logs = st.session_state.get(
            "last_logs",
            [],
        )

        if last_result:
            st.divider()
            _render_result(
                last_result
            )

        if last_error:
            st.divider()
            st.subheader("Last Error")
            st.error(last_error)

        if last_logs:
            with st.expander(
                "Last Execution Log",
                expanded=False,
            ):
                st.code(
                    "\n".join(last_logs[-300:]),
                    language="text",
                )


# =============================================================================
# Individual Module tab
# =============================================================================

def _render_individual_module_tab(
    *,
    core_files_ok: bool,
    gcloud_ok: bool,
) -> None:
    st.subheader("Individual Module")
    st.write(
        "기존 실행 차수의 결과를 그대로 사용하면서 필요한 모듈 하나만 재실행합니다."
    )
    st.warning(
        "Individual Module은 새 차수를 만들지 않습니다. "
        "선택한 기존 Output/Media 폴더에 결과를 다시 생성할 수 있으므로 날짜와 차수를 반드시 확인하세요."
    )

    if not core_files_ok:
        st.error(
            "개별 모듈 실행에 필요한 프로젝트 파일이 준비되지 않았습니다."
        )

    selector_col1, selector_col2 = st.columns(2)

    with selector_col1:
        selected_module = st.selectbox(
            "Module",
            options=list(MODULE_LABELS),
            format_func=lambda value: MODULE_LABELS[value],
            key="individual_module",
        )

    with selector_col2:
        selected_date = st.date_input(
            "Work Date",
            key="individual_date",
            format="YYYY-MM-DD",
        )

    input_date = selected_date.strftime(
        "%y%m%d"
    )

    existing_runs = ()
    existing_runs_error: str | None = None

    if list_existing_run_paths is not None:
        try:
            existing_runs = list_existing_run_paths(
                input_date
            )
        except Exception as exc:
            existing_runs_error = (
                f"{type(exc).__name__}: {exc}"
            )
    else:
        existing_runs_error = (
            "pipeline_service의 기존 실행 조회 함수를 사용할 수 없습니다."
        )

    if existing_runs_error:
        st.error(existing_runs_error)

    if not existing_runs:
        st.info(
            f"{input_date}의 기존 실행 차수를 찾지 못했습니다. "
            "먼저 Full Pipeline을 실행하세요."
        )
        selected_run = None
    else:
        selected_run = st.selectbox(
            "Existing Run",
            options=list(existing_runs),
            format_func=lambda run: (
                f"{run.run_label} · {run.output_dir.name}"
            ),
            key="individual_run_selection",
        )

        target_col1, target_col2 = st.columns(2)
        with target_col1:
            st.markdown("**Target Output**")
            st.code(
                str(selected_run.output_dir),
                language="text",
            )
        with target_col2:
            st.markdown("**Target Media**")
            st.code(
                str(selected_run.media_dir),
                language="text",
            )

    # Sprinklr만 시작/종료 datetime 두 개를 받는다.
    individual_start_datetime = None
    individual_end_datetime = None

    if selected_module == "sprinklr_export_excel.py":
        st.markdown("**Sprinklr Query Range**")
        st.caption(
            "종료 날짜는 선택한 Work Date와 같아야 합니다."
        )

        (
            ind_start_date_col,
            ind_start_time_col,
            ind_end_date_col,
            ind_end_time_col,
        ) = st.columns(4)

        with ind_start_date_col:
            ind_start_date = st.date_input(
                "Start Date",
                key="individual_start_date",
                format="YYYY-MM-DD",
            )
        with ind_start_time_col:
            ind_start_time = st.time_input(
                "Start Time",
                key="individual_start_time",
                step=timedelta(minutes=1),
            )
        with ind_end_date_col:
            ind_end_date = st.date_input(
                "End Date",
                key="individual_end_date",
                format="YYYY-MM-DD",
            )
        with ind_end_time_col:
            ind_end_time = st.time_input(
                "End Time",
                key="individual_end_time",
                step=timedelta(minutes=1),
            )

        individual_start_datetime = _combine_datetime(
            ind_start_date,
            ind_start_time,
        )
        individual_end_datetime = _combine_datetime(
            ind_end_date,
            ind_end_time,
        )

    dependency_checks = ()
    dependencies_ok = False

    if (
        selected_run is not None
        and validate_module_dependencies is not None
    ):
        dependency_checks = validate_module_dependencies(
            module_name=selected_module,
            input_date=input_date,
            run_number=selected_run.run_number,
        )
        dependencies_ok = _render_dependency_checks(
            dependency_checks
        )

    if selected_module == "llm_analysis_pipeline.py" and not gcloud_ok:
        st.error(
            "LLM Analysis 개별 실행에는 Google Cloud CLI(gcloud)가 필요합니다."
        )
        dependencies_ok = False

    existing_artifact_detected = False

    if (
        selected_run is not None
        and build_expected_run_artifacts is not None
    ):
        artifacts = build_expected_run_artifacts(
            selected_run
        )

        output_key_by_module = {
            "sprinklr_export_excel.py": "raw_excel",
            "raw_to_processed.py": "formatted_excel",
            "media_extractor.py": "media_result_excel",
        }

        artifact_key = output_key_by_module.get(
            selected_module
        )

        if artifact_key is not None:
            existing_artifact_detected = artifacts[
                artifact_key
            ].exists()

        elif selected_module == "llm_analysis_pipeline.py":
            existing_artifact_detected = (
                artifacts["llm_log_excel"].exists()
                or artifacts["llm_completed_excel"].exists()
            )

        if selected_module == "media_extractor.py":
            try:
                media_has_contents = any(
                    selected_run.media_dir.iterdir()
                )
            except OSError:
                media_has_contents = False

            existing_artifact_detected = (
                existing_artifact_detected
                or media_has_contents
            )

    overwrite_existing = st.checkbox(
        "기존 결과가 있으면 덮어쓰기 허용",
        key="individual_overwrite",
        help=(
            "1~3단계는 --overwrite, 4단계는 --overwrite-results 옵션을 사용합니다."
        ),
    )

    if existing_artifact_detected:
        if overwrite_existing:
            st.warning(
                "기존 결과가 감지되었습니다. 실행 시 선택한 모듈의 기존 결과가 갱신될 수 있습니다."
            )
        else:
            st.info(
                "기존 결과가 감지되었습니다. 모듈의 중복 보호 정책에 따라 실행이 중단될 수 있습니다. "
                "의도적인 재실행이라면 덮어쓰기를 허용하세요."
            )

    run_disabled = (
        not core_files_ok
        or selected_run is None
        or not dependencies_ok
        or run_single_module is None
    )

    if selected_module == "llm_analysis_pipeline.py":
        run_disabled = run_disabled or not gcloud_ok

    run_clicked = st.button(
        f"▶ Run {MODULE_LABELS[selected_module]}",
        type="primary",
        disabled=run_disabled,
        key="run_individual_module_button",
    )

    if run_clicked and selected_run is not None:
        st.session_state["individual_last_result"] = None
        st.session_state["individual_last_error"] = None
        st.session_state["individual_last_traceback"] = None
        st.session_state["individual_last_logs"] = []

        progress_bar = st.progress(
            0,
            text=f"{MODULE_LABELS[selected_module]} 준비 중...",
        )
        status_box = st.status(
            f"{MODULE_LABELS[selected_module]} 준비 중...",
            expanded=True,
            state="running",
        )

        st.markdown("**Execution Log**")
        log_placeholder = st.empty()
        log_placeholder.code(
            "로그 수집을 시작합니다...",
            language="text",
        )

        logs: list[str] = []
        log_callback = _create_live_log_callback(
            log_placeholder,
            logs,
        )

        def progress_callback(
            progress: PipelineProgress,
        ) -> None:
            label = MODULE_LABELS.get(
                progress.module_name,
                progress.module_name,
            )

            if progress.status == "started":
                status_box.update(
                    label=f"{label} 실행 중",
                    state="running",
                    expanded=True,
                )

            percent = int(
                round(
                    progress.fraction * 100
                )
            )
            progress_bar.progress(
                percent,
                text=f"{label} · {percent}%",
            )

        try:
            result = run_single_module(
                module_name=selected_module,
                input_date=input_date,
                run_number=selected_run.run_number,
                start_datetime=(
                    individual_start_datetime
                ),
                end_datetime=(
                    individual_end_datetime
                ),
                overwrite_existing=(
                    overwrite_existing
                ),
                log_callback=log_callback,
                progress_callback=progress_callback,
            )

        except Exception as exc:
            error_message = (
                f"{type(exc).__name__}: {exc}"
            )

            logs.append("")
            logs.append(
                "[INDIVIDUAL MODULE ERROR] "
                + error_message
            )

            log_placeholder.code(
                "\n".join(logs[-300:]),
                language="text",
            )

            status_box.update(
                label=(
                    f"{MODULE_LABELS[selected_module]} 실행 실패"
                ),
                state="error",
                expanded=True,
            )

            st.session_state[
                "individual_last_logs"
            ] = logs
            st.session_state[
                "individual_last_error"
            ] = error_message
            st.session_state[
                "individual_last_traceback"
            ] = traceback.format_exc()

            st.error(
                "개별 모듈 실행 중 오류가 발생했습니다."
            )
            st.code(
                error_message,
                language="text",
            )

            with st.expander(
                "상세 Traceback",
                expanded=False,
            ):
                st.code(
                    st.session_state[
                        "individual_last_traceback"
                    ],
                    language="text",
                )

        else:
            progress_bar.progress(
                100,
                text=(
                    f"{MODULE_LABELS[selected_module]} 완료 · 100%"
                ),
            )
            log_placeholder.code(
                "\n".join(logs[-300:]),
                language="text",
            )

            status_box.update(
                label=(
                    f"{MODULE_LABELS[selected_module]} 실행 완료"
                ),
                state="complete",
                expanded=False,
            )

            result_state = _single_result_to_state(
                result
            )

            st.session_state[
                "individual_last_logs"
            ] = logs
            st.session_state[
                "individual_last_result"
            ] = result_state

            st.success(
                f"{MODULE_LABELS[selected_module]} 실행이 정상 완료되었습니다."
            )
            _render_single_result(
                result_state
            )

    else:
        last_result = st.session_state.get(
            "individual_last_result"
        )
        last_error = st.session_state.get(
            "individual_last_error"
        )
        last_logs = st.session_state.get(
            "individual_last_logs",
            [],
        )

        if last_result:
            st.divider()
            _render_single_result(
                last_result
            )

        if last_error:
            st.divider()
            st.subheader("Last Individual Module Error")
            st.error(last_error)

        if last_logs:
            with st.expander(
                "Last Individual Module Log",
                expanded=False,
            ):
                st.code(
                    "\n".join(last_logs[-300:]),
                    language="text",
                )


# =============================================================================
# Main UI
# =============================================================================

st.title("Local Campaign Automation")
st.caption(
    "정규 Pipeline · 누락건 Pipeline · Individual Module · Buzz Volume"
)

(
    full_pipeline_ok,
    core_files_ok,
    environment_checks,
) = _check_environment()
_render_environment_check(
    environment_checks
)

gcloud_ok = shutil.which("gcloud") is not None

full_tab, missing_tab, individual_tab, buzz_tab = st.tabs(
    [
        "Full Pipeline",
        "Missing Cases",
        "Individual Module",
        "Buzz Volume",
    ]
)

with full_tab:
    _render_full_pipeline_tab(
        full_pipeline_ok=full_pipeline_ok,
    )

with missing_tab:
    _render_missing_cases_tab()

with individual_tab:
    _render_individual_module_tab(
        core_files_ok=core_files_ok,
        gcloud_ok=gcloud_ok,
    )

with buzz_tab:
    _render_buzz_volume_tab()


st.divider()
st.caption(
    "Full Pipeline은 새로운 정규 실행 차수를 생성합니다. "
    "Missing Cases는 payload QUERY를 갱신한 뒤 기존 `누락/` 서브 파이프라인을 실행합니다. "
    "Individual Module은 기존 정규 실행 차수를 재사용하고, "
    "Buzz Volume은 통합·정제 완료된 최종 Excel을 대상으로 별도 실행합니다."
)
