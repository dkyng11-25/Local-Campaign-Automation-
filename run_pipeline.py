from __future__ import annotations

from pipeline_service import (
    run_local_campaign_pipeline,
)


def main() -> None:
    print("=" * 70)
    print(
        "Local Campaign 전체 자동화 파이프라인"
    )
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

    run_local_campaign_pipeline(
        start_datetime=start_datetime_text,
        end_datetime=end_datetime_text,
    )


if __name__ == "__main__":
    main()
