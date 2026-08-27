from datetime import datetime
from pathlib import Path
import subprocess
import sys


MODULES = [
    "sprinklr_export_excel.py",
    "raw_to_processed.py",
    "media_extractor.py",
    "llm_analysis_pipeline.py",
]


def parse_datetime(
    datetime_text: str,
    input_name: str,
) -> datetime:
    """
    YYYY-MM-DD HH:MM:SS 형식의 문자열을 datetime으로 변환합니다.
    """

    datetime_text = datetime_text.strip()

    try:
        return datetime.strptime(
            datetime_text,
            "%Y-%m-%d %H:%M:%S",
        )

    except ValueError as exc:
        raise ValueError(
            f"{input_name}은 YYYY-MM-DD HH:MM:SS 형식이어야 합니다.\n"
            f"입력값: {datetime_text}\n"
            "예시: 2026-07-27 19:00:00"
        ) from exc


def run_module(
    module_name: str,
    module_inputs: list[str],
) -> None:
    """
    Python 모듈을 실행하고, 해당 모듈의 input()에 값을 순서대로 전달합니다.

    예:
        module_inputs=[
            "2026-07-26 19:00:00",
            "2026-07-27 19:00:00",
        ]

    전달되는 입력:
        2026-07-26 19:00:00
        2026-07-27 19:00:00
    """

    module_path = Path(__file__).parent / module_name

    if not module_path.exists():
        raise FileNotFoundError(
            f"모듈 파일을 찾을 수 없습니다: {module_path}"
        )

    # input() 호출 순서에 맞게 줄바꿈으로 연결
    stdin_text = "\n".join(module_inputs) + "\n"

    print()
    print("=" * 70)
    print(f"실행 시작: {module_name}")

    for index, input_value in enumerate(
        module_inputs,
        start=1,
    ):
        print(f"전달 입력값 {index}: {input_value}")

    print("=" * 70)

    result = subprocess.run(
        [sys.executable, str(module_path)],
        input=stdin_text,
        text=True,
        check=False,
    )

    if result.returncode != 0:
        raise RuntimeError(
            f"{module_name} 실행 실패 "
            f"(return code: {result.returncode})"
        )

    print(f"[OK] 실행 완료: {module_name}")


def main() -> None:
    print("=" * 70)
    print("Local Campaign 전체 자동화 파이프라인")
    print("=" * 70)

    start_datetime_text = input(
        "\nSprinklr 조회 시작 날짜와 시간을 입력하세요.\n"
        "형식: YYYY-MM-DD HH:MM:SS\n"
        "예시: 2026-07-26 19:00:00\n"
        "입력: "
    ).strip()

    end_datetime_text = input(
        "\nSprinklr 조회 종료 날짜와 시간을 입력하세요.\n"
        "형식: YYYY-MM-DD HH:MM:SS\n"
        "예시: 2026-07-27 19:00:00\n"
        "입력: "
    ).strip()

    start_datetime = parse_datetime(
        datetime_text=start_datetime_text,
        input_name="시작 날짜와 시간",
    )

    end_datetime = parse_datetime(
        datetime_text=end_datetime_text,
        input_name="종료 날짜와 시간",
    )

    if end_datetime <= start_datetime:
        raise ValueError(
            "종료 날짜와 시간은 시작 날짜와 시간보다 늦어야 합니다.\n"
            f"시작: {start_datetime}\n"
            f"종료: {end_datetime}"
        )

    normalized_start_datetime = start_datetime.strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    normalized_end_datetime = end_datetime.strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    # 종료 시각의 날짜를 작업 기준 날짜로 사용
    input_date = end_datetime.strftime("%y%m%d")

    print()
    print("입력값 확인")
    print(f"- 조회 시작 시각: {normalized_start_datetime}")
    print(f"- 조회 종료 시각: {normalized_end_datetime}")
    print(f"- 후속 모듈 작업 날짜: {input_date}")

    # =========================================================
    # 1단계
    # sprinklr_export_excel.py의 input() 호출 순서:
    #   1. start_datetime
    #   2. end_datetime
    # =========================================================
    run_module(
        module_name="sprinklr_export_excel.py",
        module_inputs=[
            normalized_start_datetime,
            normalized_end_datetime,
        ],
    )

    # =========================================================
    # 2~4단계
    # 각 모듈의 input()에 YYMMDD 한 번 전달
    # =========================================================
    for module_name in [
        "raw_to_processed.py",
        "media_extractor.py",
        "llm_analysis_pipeline.py",
    ]:
        run_module(
            module_name=module_name,
            module_inputs=[input_date],
        )

    print()
    print("=" * 70)
    print("[OK] 전체 Local Campaign 자동화 파이프라인 (누락건) 실행 완료")
    print("=" * 70)


if __name__ == "__main__":
    main()