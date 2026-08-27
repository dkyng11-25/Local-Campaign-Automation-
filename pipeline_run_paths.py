from __future__ import annotations

import argparse
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path


RUN_FOLDER_SUFFIX = "차"


@dataclass(frozen=True)
class PipelineRunPaths:
    """
    한 번의 파이프라인 실행에서 공통으로 사용할 경로 정보.

    최초 실행:
        run_number = 1
        output_dir = output/{YYMMDD}
        media_dir = media/{YYMMDD}

    두 번째 실행부터:
        run_number = 2, 3, ...
        output_dir = output/{YYMMDD}_{run_number}차
        media_dir = media/{YYMMDD}_{run_number}차
    """

    input_date: str
    run_number: int
    output_dir: Path
    media_dir: Path

    @property
    def run_label(self) -> str:
        return f"{self.run_number}{RUN_FOLDER_SUFFIX}"


def validate_input_date(input_date: str) -> str:
    """
    입력 날짜가 YYMMDD 형식의 실제 날짜인지 검증한다.

    예:
        260805 -> 정상
        20260805 -> 오류
        260231 -> 오류
    """

    normalized_date = input_date.strip()

    try:
        datetime.strptime(normalized_date, "%y%m%d")
    except ValueError as exc:
        raise ValueError(
            "날짜는 YYMMDD 형식의 실제 날짜여야 합니다. "
            "예시: 260805"
        ) from exc

    return normalized_date


def _ensure_directory_or_missing(path: Path) -> None:
    """
    경로가 없거나 디렉터리인 경우만 허용한다.

    같은 이름의 일반 파일이 존재하면 폴더를 만들거나 이름을
    변경할 수 없으므로 명확한 오류를 발생시킨다.
    """

    if path.exists() and not path.is_dir():
        raise FileExistsError(
            "디렉터리로 사용해야 하는 경로에 일반 파일이 존재합니다: "
            f"{path}"
        )


def _build_run_pattern(input_date: str) -> re.Pattern[str]:
    """
    예: 260805_1차, 260805_2차 형태만 인식하는 정규식 생성.
    """

    return re.compile(
        rf"^{re.escape(input_date)}_(\d+){re.escape(RUN_FOLDER_SUFFIX)}$"
    )


def find_numbered_run_directories(
    root_dir: Path,
    input_date: str,
) -> dict[int, Path]:
    """
    root_dir 아래에서 해당 날짜의 차수 폴더를 찾는다.

    반환 예:
        {
            1: Path("output/260805_1차"),
            2: Path("output/260805_2차"),
        }
    """

    if not root_dir.exists():
        return {}

    pattern = _build_run_pattern(input_date)
    found: dict[int, Path] = {}

    for child_path in root_dir.iterdir():
        if not child_path.is_dir():
            continue

        matched = pattern.fullmatch(child_path.name)
        if matched is None:
            continue

        run_number = int(matched.group(1))

        if run_number < 1:
            raise ValueError(
                f"실행 차수는 1 이상이어야 합니다: {child_path}"
            )

        found[run_number] = child_path

    return found


def _validate_contiguous_run_numbers(
    run_numbers: set[int],
    *,
    root_name: str,
    input_date: str,
) -> None:
    """
    차수 폴더가 1차부터 끊김 없이 이어지는지 검증한다.

    예:
        {1, 2, 3} -> 정상
        {1, 3}    -> 오류
        {2}       -> 오류
    """

    if not run_numbers:
        return

    expected = set(range(1, max(run_numbers) + 1))

    if run_numbers != expected:
        missing = sorted(expected - run_numbers)
        raise RuntimeError(
            f"{root_name}의 {input_date} 실행 차수 폴더가 연속적이지 않습니다. "
            f"현재 차수={sorted(run_numbers)}, 누락 차수={missing}"
        )


def _create_directory_pair(
    output_dir: Path,
    media_dir: Path,
) -> None:
    """
    output/media 폴더를 한 쌍으로 생성한다.

    두 번째 폴더 생성에 실패하면 먼저 만든 빈 output 폴더를
    제거하여 한쪽만 남는 상태를 최대한 방지한다.
    """

    _ensure_directory_or_missing(output_dir)
    _ensure_directory_or_missing(media_dir)

    if output_dir.exists():
        raise FileExistsError(
            f"생성할 output 폴더가 이미 존재합니다: {output_dir}"
        )

    if media_dir.exists():
        raise FileExistsError(
            f"생성할 media 폴더가 이미 존재합니다: {media_dir}"
        )

    output_created = False

    try:
        output_dir.mkdir(parents=True, exist_ok=False)
        output_created = True

        media_dir.mkdir(parents=True, exist_ok=False)

    except Exception:
        if output_created and output_dir.exists():
            try:
                output_dir.rmdir()
            except OSError:
                pass
        raise


def _rename_directory_pair(
    source_output_dir: Path,
    target_output_dir: Path,
    source_media_dir: Path,
    target_media_dir: Path,
) -> None:
    """
    output/media 기본 날짜 폴더를 각각 1차 폴더로 변경한다.

    media 폴더 이름 변경에 실패하면 output 폴더 이름을 원래대로
    되돌리는 rollback을 시도한다.
    """

    if not source_output_dir.is_dir():
        raise FileNotFoundError(
            f"이름을 변경할 output 폴더가 없습니다: {source_output_dir}"
        )

    if not source_media_dir.is_dir():
        raise FileNotFoundError(
            f"이름을 변경할 media 폴더가 없습니다: {source_media_dir}"
        )

    _ensure_directory_or_missing(target_output_dir)
    _ensure_directory_or_missing(target_media_dir)

    if target_output_dir.exists():
        raise FileExistsError(
            f"변경 대상 output 폴더가 이미 존재합니다: {target_output_dir}"
        )

    if target_media_dir.exists():
        raise FileExistsError(
            f"변경 대상 media 폴더가 이미 존재합니다: {target_media_dir}"
        )

    output_renamed = False

    try:
        source_output_dir.rename(target_output_dir)
        output_renamed = True

        source_media_dir.rename(target_media_dir)

    except Exception:
        if (
            output_renamed
            and target_output_dir.exists()
            and not source_output_dir.exists()
        ):
            try:
                target_output_dir.rename(source_output_dir)
            except OSError:
                pass
        raise


def prepare_pipeline_run_paths(
    input_date: str,
    output_root: Path,
    media_root: Path,
) -> PipelineRunPaths:
    """
    날짜별 실행 폴더를 검색하고 이번 실행 차수를 확정한다.

    규칙
    ----
    1. 최초 실행
       output/{date}
       media/{date}

    2. 두 번째 실행
       기존 output/{date} -> output/{date}_1차
       기존 media/{date}  -> media/{date}_1차
       신규 output/{date}_2차
       신규 media/{date}_2차

    3. 세 번째 이후
       기존 최대 차수 + 1 폴더를 output/media에 함께 생성

    안전 정책
    ---------
    - output과 media의 상태 또는 차수가 다르면 중단한다.
    - 기본 날짜 폴더와 차수 폴더가 동시에 있으면 중단한다.
    - 차수 폴더가 1차부터 연속적이지 않으면 중단한다.
    - 불명확한 상태를 자동 추정하거나 덮어쓰지 않는다.
    """

    normalized_date = validate_input_date(input_date)

    output_root = Path(output_root).resolve()
    media_root = Path(media_root).resolve()

    _ensure_directory_or_missing(output_root)
    _ensure_directory_or_missing(media_root)

    output_root.mkdir(parents=True, exist_ok=True)
    media_root.mkdir(parents=True, exist_ok=True)

    base_output_dir = output_root / normalized_date
    base_media_dir = media_root / normalized_date

    _ensure_directory_or_missing(base_output_dir)
    _ensure_directory_or_missing(base_media_dir)

    base_output_exists = base_output_dir.is_dir()
    base_media_exists = base_media_dir.is_dir()

    # output과 media는 항상 같은 실행 상태여야 한다.
    if base_output_exists != base_media_exists:
        raise RuntimeError(
            "기본 날짜 폴더의 상태가 서로 다릅니다. "
            "자동으로 차수를 계산하면 결과가 잘못 연결될 수 있으므로 중단합니다.\n"
            f"output 기본 폴더 존재={base_output_exists}: {base_output_dir}\n"
            f"media 기본 폴더 존재={base_media_exists}: {base_media_dir}"
        )

    output_numbered = find_numbered_run_directories(
        output_root,
        normalized_date,
    )
    media_numbered = find_numbered_run_directories(
        media_root,
        normalized_date,
    )

    output_run_numbers = set(output_numbered)
    media_run_numbers = set(media_numbered)

    _validate_contiguous_run_numbers(
        output_run_numbers,
        root_name="output",
        input_date=normalized_date,
    )
    _validate_contiguous_run_numbers(
        media_run_numbers,
        root_name="media",
        input_date=normalized_date,
    )

    if output_run_numbers != media_run_numbers:
        raise RuntimeError(
            "output과 media의 실행 차수가 서로 다릅니다. "
            "어느 결과끼리 연결되는지 확정할 수 없어 중단합니다.\n"
            f"output 차수={sorted(output_run_numbers)}\n"
            f"media 차수={sorted(media_run_numbers)}"
        )

    # 기본 폴더와 차수 폴더가 동시에 존재하면 실행 이력이 모호하다.
    if base_output_exists and output_run_numbers:
        raise RuntimeError(
            "기본 날짜 폴더와 차수 폴더가 동시에 존재합니다. "
            "기존 폴더를 수동으로 확인한 뒤 다시 실행하세요.\n"
            f"기본 output 폴더: {base_output_dir}\n"
            f"기존 차수: {sorted(output_run_numbers)}"
        )

    # 1차: 아무 실행 폴더도 없는 경우 기존 규칙대로 기본 날짜 폴더 생성
    if not base_output_exists and not output_run_numbers:
        _create_directory_pair(
            output_dir=base_output_dir,
            media_dir=base_media_dir,
        )

        return PipelineRunPaths(
            input_date=normalized_date,
            run_number=1,
            output_dir=base_output_dir,
            media_dir=base_media_dir,
        )

    # 2차: 기본 날짜 폴더가 존재하면 기존 결과를 1차로 보존
    if base_output_exists:
        first_output_dir = output_root / f"{normalized_date}_1차"
        first_media_dir = media_root / f"{normalized_date}_1차"

        second_output_dir = output_root / f"{normalized_date}_2차"
        second_media_dir = media_root / f"{normalized_date}_2차"

        _rename_directory_pair(
            source_output_dir=base_output_dir,
            target_output_dir=first_output_dir,
            source_media_dir=base_media_dir,
            target_media_dir=first_media_dir,
        )

        try:
            _create_directory_pair(
                output_dir=second_output_dir,
                media_dir=second_media_dir,
            )
        except Exception:
            # 2차 폴더 생성 실패 시 기존 1차 폴더명을 원래대로 복구 시도
            try:
                _rename_directory_pair(
                    source_output_dir=first_output_dir,
                    target_output_dir=base_output_dir,
                    source_media_dir=first_media_dir,
                    target_media_dir=base_media_dir,
                )
            except Exception as rollback_exc:
                raise RuntimeError(
                    "2차 실행 폴더 생성에 실패했고 기존 폴더명 복구도 실패했습니다. "
                    "output/media 폴더 상태를 수동으로 확인해야 합니다."
                ) from rollback_exc
            raise

        return PipelineRunPaths(
            input_date=normalized_date,
            run_number=2,
            output_dir=second_output_dir,
            media_dir=second_media_dir,
        )

    # 3차 이후: 기존 차수 중 최대값 + 1
    next_run_number = max(output_run_numbers) + 1

    next_output_dir = (
        output_root
        / f"{normalized_date}_{next_run_number}{RUN_FOLDER_SUFFIX}"
    )
    next_media_dir = (
        media_root
        / f"{normalized_date}_{next_run_number}{RUN_FOLDER_SUFFIX}"
    )

    _create_directory_pair(
        output_dir=next_output_dir,
        media_dir=next_media_dir,
    )

    return PipelineRunPaths(
        input_date=normalized_date,
        run_number=next_run_number,
        output_dir=next_output_dir,
        media_dir=next_media_dir,
    )


def parse_arguments() -> argparse.Namespace:
    project_root = Path(__file__).resolve().parent

    parser = argparse.ArgumentParser(
        description=(
            "로컬 캠페인 파이프라인의 날짜별 output/media 실행 폴더를 "
            "검색하고 다음 실행 차수를 생성합니다."
        )
    )

    parser.add_argument(
        "--date",
        dest="input_date",
        help="작업 날짜(YYMMDD). 생략하면 터미널에서 입력받습니다.",
    )
    parser.add_argument(
        "--output-root",
        type=Path,
        default=project_root / "output",
        help="output 루트 폴더 경로",
    )
    parser.add_argument(
        "--media-root",
        type=Path,
        default=project_root / "media",
        help="media 루트 폴더 경로",
    )

    return parser.parse_args()


def main() -> None:
    args = parse_arguments()

    input_date = args.input_date
    if input_date is None:
        input_date = input(
            "작업 날짜를 YYMMDD 형식으로 입력하세요: "
        ).strip()

    run_paths = prepare_pipeline_run_paths(
        input_date=input_date,
        output_root=args.output_root,
        media_root=args.media_root,
    )

    print("=" * 72)
    print("[DONE] 파이프라인 실행 폴더 준비 완료")
    print(f"[INFO] 작업 날짜: {run_paths.input_date}")
    print(f"[INFO] 실행 차수: {run_paths.run_label}")
    print(f"[INFO] Output 폴더: {run_paths.output_dir}")
    print(f"[INFO] Media 폴더: {run_paths.media_dir}")
    print("=" * 72)


if __name__ == "__main__":
    main()
