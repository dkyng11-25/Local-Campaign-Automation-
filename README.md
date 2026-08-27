**# Local Campaign Analysis Automation Pipeline**

**## 1. 프로젝트 개요**

본 프로젝트는 Sprinklr에서 Local Campaign 데이터를 추출하고, 데이터를 정제한 뒤 게시물의 이미지·영상 미디어를 확보하여 LLM 분석 결과까지 생성하는 자동화 파이프라인입니다.

\`run\_pipeline.py\`는 다음 4개 모듈을 순서대로 실행합니다. 사용자는 Sprinklr 조회 시작 시각과 종료 시각만 입력하며, 후속 모듈에서 사용하는 \`YYMMDD\` 작업 날짜는 종료 시각을 기준으로 자동 생성됩니다.

\`\`\`text
run\_pipeline.py
    │
    ├─ [1] sprinklr\_export\_excel.py
    ├─ [2] raw\_to\_processed.py
    ├─ [3] media\_extractor.py
    └─ [4] llm\_analysis\_pipeline.py
\`\`\`

각 단계의 주요 역할은 다음과 같습니다.

1\. Sprinklr API에서 Local Campaign Raw 데이터 추출
2\. Raw Excel을 분석 가능한 형태로 전처리
3\. 게시물별 이미지 및 영상 파일 추출
4\. 프롬프트와 미디어를 LLM에 전달하여 분석 결과 생성

\`buzz\_volume\_adaptor.py\`는 위 1\~4단계와 별도로 실행합니다. 1차·2차 등 여러 실행 결과를 통합하고 수동 정제까지 완료한 최종 Excel을 \`output/Buzz\_Volume\` 폴더에 넣은 뒤 실행합니다.

\`\`\`text
[5] buzz\_volume\_adaptor.py
    ├─ Daily 또는 Weekly 데이터 컷 설정
    ├─ 캠페인별 Query로 Sprinklr API 호출
    ├─ Buzz Volume 값 적재
    └─ output/Buzz\_Volume/completed에 최종 결과 저장
\`\`\`

정상 운영 시 1\~4단계는 \`run\_pipeline.py\`로 실행합니다. 개별 모듈 명령어는 특정 단계에서 오류가 발생했거나, 이미 만들어진 실행 차수 폴더에서 해당 단계만 다시 실행할 때 사용합니다.

**### 반드시 확인해야 하는 운영 규칙**

파이프라인을 실행하기 전에 다음 항목을 반드시 확인합니다.

1\. **\*\*조회 종료 시각의 초는 반드시 \`59\`로 입력합니다.\*\***
   - 예: \`2026-07-27 16:30:59\`
   - 종료 초를 \`00\`으로 입력하면 해당 종료 분의 데이터가 누락될 수 있습니다.

2\. **\*\*개별 모듈을 다시 실행할 때는 \`run\_pipeline.py\` 로그에 표시된 정확한 Output/Media 폴더를 사용합니다.\*\***
   - 예: \`output\260805\_2차\`
   - 예: \`media\260805\_2차\`

3\. **\*\*Buzz Volume 입력 파일은 날짜별 실행 폴더가 아니라 공용 폴더에 넣습니다.\*\***
   - 입력: \`output\Buzz\_Volume\`
   - 완료 결과: \`output\Buzz\_Volume\completed\`

**## 2. 전체 Workflow**

\`\`\`text
사용자 입력
├─ 조회 시작 시각: YYYY-MM-DD HH\:MM\:SS
└─ 조회 종료 시각: YYYY-MM-DD HH\:MM\:SS
        │
        ▼
run\_pipeline.py
        │
        ├─ 종료 시각에서 YYMMDD 자동 생성
        ├─ 실행 차수별 output/media 폴더 확정
        └─ output/Buzz\_Volume 및 completed 폴더 생성·확인
        │
        ▼
[1] sprinklr\_export\_excel.py
        │
        ├─ Sprinklr API 호출
        ├─ 원문 및 전략법인 데이터 조회
        ├─ 게시물 URL, 본문, 플랫폼, 미디어 URL 정리
        ├─ 게시 계정명, Screen Name, 위치, Bio, Website, 인증 정보, 팔로워 수 정리
        └─ Raw Excel 생성
        │
        ▼
[2] raw\_to\_processed.py
        │
        ├─ 필요한 컬럼 선택 및 정리
        ├─ 데이터 형식 표준화
        ├─ 분석 대상 행 생성
        └─ Processed Excel 생성
        │
        ▼
[3] media\_extractor.py
        │
        ├─ 플랫폼별 미디어 추출 전략 적용
        ├─ 이미지 및 영상 다운로드
        ├─ 게시물과 로컬 미디어 파일 연결
        └─ Media Result Excel 생성
        │
        ▼
[4] llm\_analysis\_pipeline.py
        │
        ├─ 분석용 Input Set 생성
        ├─ Prompt 및 JSON Schema 로드
        ├─ 미디어와 게시물 정보를 LLM에 전달
        ├─ 게시자 유형·국가 판정 및 Country–Subsidiary 정확 매핑
        ├─ LLM 응답 검증 및 구조화
        └─ 분석 결과 저장
        │
        ▼
1차·2차 등 실행 결과 통합 및 사용자 최종 정제
        │
        ├─ Buzz Volume 이외 항목 완료
        └─ 지정 파일명으로 output/Buzz\_Volume에 저장
        │
        ▼
[5] buzz\_volume\_adaptor.py — 별도 실행
        │
        ├─ 기준 날짜로 입력 파일명 자동 계산
        ├─ Daily 또는 Weekly 대상 행 선택
        ├─ 캠페인별 Query 병렬 API 호출
        ├─ Buzz Volume 적재
        └─ output/Buzz\_Volume/completed에 결과 저장
\`\`\`

1\~4단계 중 앞 단계가 실패하면 \`run\_pipeline.py\`는 즉시 중단됩니다. 이는 불완전하거나 이전에 생성된 파일을 후속 모듈이 잘못 처리하는 것을 방지하기 위한 동작입니다.

Buzz Volume 모듈은 1\~4단계가 완료된 직후 자동 실행되지 않습니다. 여러 차수의 데이터를 통합하고 수동 정제한 최종 파일을 준비한 후 별도로 실행합니다.

**## 3. 파이프라인 실행 방법**

**### 3.1 전체 파이프라인 실행 — 기본 권장 방식**

프로젝트 루트에서 다음 명령어를 실행합니다.

\`\`\`powershell
python run\_pipeline.py
\`\`\`

\`uv\`를 사용하는 경우:

\`\`\`powershell
uv run python run\_pipeline.py
\`\`\`

실행 시 다음 두 값을 한 번씩 입력합니다.

\`\`\`text
조회 시작 시각: YYYY-MM-DD HH\:MM\:SS
조회 종료 시각: YYYY-MM-DD HH\:MM\:SS
\`\`\`

\> **\*\*중요:\*\*** 조회 종료 시각의 초는 반드시 \`59\`로 입력합니다.  
\> 종료 초가 \`00\`이면 해당 종료 분에 생성된 데이터가 누락될 수 있습니다.

예시:

\`\`\`text
조회 시작 시각: 2026-08-05 09:00:00
조회 종료 시각: 2026-08-05 12:00:59
\`\`\`

후속 모듈에서 사용하는 작업 날짜는 종료 시각을 기준으로 자동 생성됩니다.

\`\`\`text
2026-08-05 12:00:59 → 260805
\`\`\`

\`run\_pipeline.py\`는 파이프라인 전체에서 한 번만 실행 차수를 확정합니다. 같은 날짜의 실행이 반복되면 Output과 Media 폴더가 동일한 차수로 관리됩니다.

\`\`\`text
첫 실행:
output\260805
media\260805

같은 날짜의 두 번째 실행을 시작한 이후:
output\260805\_1차
media\260805\_1차

새 실행:
output\260805\_2차
media\260805\_2차
\`\`\`

실제 생성된 경로는 실행 로그의 다음 항목을 기준으로 확인합니다.

\`\`\`text
실행 차수
Output 폴더
Media 폴더
\`\`\`

**### 3.2 개별 모듈 실행 — 오류 복구 및 단계별 재실행용**

개별 실행 시에는 사용할 실행 차수의 디렉터리를 명령어에 직접 지정해야 합니다. 사용자는 폴더 이름을 추측하지 않고 \`run\_pipeline.py\`가 출력한 실제 Output/Media 경로를 사용합니다.

아래 예시는 \`260805\_2차\`를 다시 실행하는 경우입니다.

**#### 1단계 단독 실행**

\`\`\`powershell
python sprinklr\_export\_excel.py \`
  --output-dir "output\260805\_2차"
\`\`\`

\`uv\` 환경:

\`\`\`powershell
uv run python sprinklr\_export\_excel.py \`
  --output-dir "output\260805\_2차"
\`\`\`

실행 후 Sprinklr 조회 시작 시각과 종료 시각을 입력합니다.

**#### 2단계 단독 실행**

\`\`\`powershell
python raw\_to\_processed.py \`
  --output-dir "output\260805\_2차"
\`\`\`

\`uv\` 환경:

\`\`\`powershell
uv run python raw\_to\_processed.py \`
  --output-dir "output\260805\_2차"
\`\`\`

실행 후 작업 날짜를 \`YYMMDD\` 형식으로 입력합니다.

**#### 3단계 단독 실행**

\`\`\`powershell
python media\_extractor.py \`
  --output-dir "output\260805\_2차" \`
  --media-dir "media\260805\_2차"
\`\`\`

\`uv\` 환경:

\`\`\`powershell
uv run python media\_extractor.py \`
  --output-dir "output\260805\_2차" \`
  --media-dir "media\260805\_2차"
\`\`\`

실행 후 작업 날짜를 \`YYMMDD\` 형식으로 입력합니다.

**#### 4단계 단독 실행**

\`\`\`powershell
python llm\_analysis\_pipeline.py \`
  --output-dir "output\260805\_2차"
\`\`\`

\`uv\` 환경:

\`\`\`powershell
uv run python llm\_analysis\_pipeline.py \`
  --output-dir "output\260805\_2차"
\`\`\`

실행 후 작업 날짜를 \`YYMMDD\` 형식으로 입력합니다.

**#### 기존 결과를 의도적으로 다시 생성하는 경우**

모듈이 기존 결과 파일을 보호하도록 구현된 경우 다음 옵션이 추가로 필요할 수 있습니다.

\`\`\`powershell
\# 1\~3단계
\--overwrite

\# 4단계
\--overwrite-results
\`\`\`

예시:

\`\`\`powershell
python media\_extractor.py \`
  --output-dir "output\260805\_2차" \`
  --media-dir "media\260805\_2차" \`
  --overwrite
\`\`\`

실행 옵션은 각 모듈의 도움말에서 확인할 수 있습니다.

\`\`\`powershell
python sprinklr\_export\_excel.py --help
python raw\_to\_processed.py --help
python media\_extractor.py --help
python llm\_analysis\_pipeline.py --help
\`\`\`

개별 실행 시에도 반드시 앞 단계의 출력 파일이 정상인지 확인한 후 다음 단계를 실행합니다. 실패한 실행을 재개할 때는 새로운 차수 폴더를 만들지 않고, 실패했던 정확한 Output/Media 폴더를 재사용합니다.

**### 3.3 Buzz Volume 별도 실행**

\`buzz\_volume\_adaptor.py\`는 날짜별·차수별 Output 폴더를 직접 사용하지 않습니다. 여러 차수의 결과를 통합하고 수동 정제한 최종 Excel을 다음 위치에 넣습니다.

\`\`\`text
output\Buzz\_Volume\\
{YYMMDD}\_SLCC\_SOV\_Local Campaign Tracking\_{월}월\_v01.xlsx
\`\`\`

예시:

\`\`\`text
output\Buzz\_Volume\\
260805\_SLCC\_SOV\_Local Campaign Tracking\_8월\_v01.xlsx
\`\`\`

그다음 프로젝트 루트에서 실행합니다.

\`\`\`powershell
python buzz\_volume\_adaptor.py
\`\`\`

\`uv\` 환경:

\`\`\`powershell
uv run python buzz\_volume\_adaptor.py
\`\`\`

실행 시 다음 값을 입력합니다.

\`\`\`text
데이터 컷 유형: daily 또는 weekly
기준 날짜: YYYY-MM-DD
\`\`\`

예시:

\`\`\`text
데이터 컷 유형: daily
기준 날짜: 2026-08-05
\`\`\`

모듈은 기준 날짜를 이용해 입력 파일명을 자동으로 계산하므로 파일 경로나 파일명을 명령어로 지정하지 않습니다.

결과는 다음 위치에 저장됩니다.

\`\`\`text
output\Buzz\_Volume\completed\\
{YYMMDD}\_SLCC\_SOV\_Local Campaign Tracking\_{월}월\_v01\_mentions\_updated.xlsx
\`\`\`

**### 3.4 누락건 별도 처리**

기존 1~4단계 파이프라인 실행 이후 누락된 캠페인 건을 별도로 처리해야 하는 경우, 프로젝트 루트의 `누락` 폴더에 구성된 별도 파이프라인을 실행합니다.

누락건은 다음 순서로 처리합니다.

**#### Step 1. 누락건 수기 검수 및 Query 준비**

누락된 캠페인 건을 수기로 검수한 뒤, 처리 대상 Query들을 `OR` 조건으로 하나의 Query로 연결합니다.

예시:

```text
Query_A OR Query_B OR Query_C
```

연결한 Query를 프로젝트 루트의 다음 payload 파일에 반영합니다.

```text
payload\payload_6_1_누락건.json
```

해당 JSON 파일의 `filters` 내부에서 누락건 조회에 사용하는 Query 필터의 `values` 값을 위에서 연결한 Query로 업데이트합니다.

```text
누락건 수기 검수
    ↓
대상 Query 확인
    ↓
Query들을 OR로 연결
    ↓
payload\payload_6_1_누락건.json
    ↓
filters의 values 값을 연결된 Query로 업데이트
```

누락건 실행 전에는 반드시 `payload_6_1_누락건.json`에 이번 실행 대상 Query가 정확히 반영되었는지 확인합니다.

**#### Step 2. 터미널에서 누락 폴더로 이동**

PowerShell의 현재 위치가 프로젝트 루트인 상태에서 다음 명령어를 실행합니다.

```powershell
cd "누락"
```

이동 후 현재 위치를 확인합니다.

```powershell
pwd
```

현재 경로가 프로젝트 루트 하위의 `누락` 폴더인지 확인합니다.

```text
Local_Campaign_Automation
└─ 누락
```

**#### Step 3. 누락 파이프라인 실행**

터미널에서 `누락` 폴더로 이동한 것이 확인되면 해당 폴더 안의 `run_pipeline.py`를 실행합니다.

```powershell
python run_pipeline.py
```

`uv`를 사용하는 경우:

```powershell
uv run python run_pipeline.py
```

누락건 처리 순서를 요약하면 다음과 같습니다.

```text
[1] 누락건 수기 검수
        ↓
[2] 대상 Query들을 OR로 연결
        ↓
[3] payload\payload_6_1_누락건.json
    filters의 values 업데이트
        ↓
[4] 프로젝트 루트에서 cd "누락"
        ↓
[5] 누락 폴더에서 run_pipeline.py 실행
```

누락건 파이프라인은 일반 실행용 `run_pipeline.py`가 아니라 `누락` 폴더 내부의 `run_pipeline.py`를 실행해야 합니다.

**## 4. 시스템 요구사항**

**### 4.1 운영체제**

현재 프로젝트는 다음 환경을 기준으로 작성되었습니다.

\`\`\`text
Operating System: Windows
Shell: PowerShell
\`\`\`

macOS 및 Linux 환경에서의 실행은 별도 검증이 필요합니다.

**### 4.2 필수 프로그램**

다음 프로그램이 설치되어 있어야 합니다.

\* Python
\* uv — 선택 사항
\* Google Cloud SDK
\* gcloud CLI
\* gallery-dl — \`requirements.txt\`를 통해 설치
\* Git 또는 프로젝트 파일을 전달받을 수 있는 환경

설치 여부는 PowerShell에서 다음 명령어로 확인할 수 있습니다.

\`\`\`powershell
python --version
uv --version
gcloud --version
gallery-dl --version
\`\`\`

**#### 권장 Python 버전**

\`\`\`text
Python 3.11 이상
\`\`\`

\`media\_extractor.py\`에서 \`StrEnum\`을 사용하므로 Python 3.11 이상이 필요합니다.

**---**

**## 5. 프로젝트 폴더 구조**

아래는 현재 권장 프로젝트 구조입니다.

\`\`\`text
Local\_Campaign\_Automation
│
├─ pipeline\_run\_paths.py           # 날짜·실행 차수별 공통 경로 관리
├─ run\_pipeline.py                 # 1\~4단계 전체 파이프라인 실행
├─ sprinklr\_export\_excel.py
├─ raw\_to\_processed.py
├─ media\_extractor.py
├─ llm\_analysis\_pipeline.py
├─ buzz\_volume\_adaptor.py          # 통합·정제 파일의 Buzz Volume 별도 적재
│
├─ payload/
│  └─ buzz\_volume\_base\_payload.json
\|  └─ payload\_{}.json              # 필요한 sprinklr dashboard widget에서 다운로드 한 payload 
│
├─ prompts/
│  ├─ user\_prompt.txt
│  └─ response\_schema.json
│
├─ config/
│  └─ country\_subsidiary\_mapping.json
│
├─ output/
│  ├─ YYMMDD/                      # 해당 날짜의 첫 실행 폴더
│  │  ├─ {YYMMDD}\_SLCC\_SOV\_Local Campaign Tracking\_{월}월\_v01.xlsx
│  │  ├─ {YYMMDD}\_SLCC\_SOV\_Local Campaign Tracking\_{월}월\_v01\_formatted.xlsx
│  │  ├─ {YYMMDD}\_campaign\_media\_result.xlsx
│  │  └─ {YYMMDD}\_campaign\_media\_result\_llm\_result.xlsx
│  │
│  ├─ YYMMDD\_1차/                  # 같은 날짜에 추가 실행 시 기존 첫 실행 폴더
│  ├─ YYMMDD\_2차/                  # 같은 날짜의 두 번째 실행 폴더
│  ├─ YYMMDD\_3차/
│  │
│  └─ Buzz\_Volume/                 # 차수와 독립적인 공용 Buzz Volume 작업 폴더
│     ├─ {YYMMDD}\_SLCC\_SOV\_Local Campaign Tracking\_{월}월\_v01.xlsx
│     ├─ completed/
│     │  └─ {YYMMDD}\_SLCC\_SOV\_Local Campaign Tracking\_{월}월\_v01\_mentions\_updated.xlsx
│     └─ artifacts/
│        ├─ failed\_results/
│        └─ response\_samples/
│
├─ media/
│  ├─ YYMMDD/
│  │  ├─ Raw Data\_원문/
│  │  └─ Raw Data\_전략법인/
│  ├─ YYMMDD\_1차/
│  └─ YYMMDD\_2차/
│
├─ logs/
├─ .env
├─ requirements.txt
├─ pyproject.toml                  # uv 프로젝트를 사용할 경우
├─ uv.lock                         # uv 프로젝트를 사용할 경우
└─ README.md
\`\`\`

\`YYMMDD\`는 실행 대상 날짜를 의미합니다. 예를 들어 \`2026-08-05\`는 \`260805\`입니다.

같은 날짜에 파이프라인을 여러 번 실행하면 Output과 Media가 같은 실행 차수를 사용합니다.

\`\`\`text
output\260805\_2차
media\260805\_2차
\`\`\`

\`output\Buzz\_Volume\`은 날짜별 실행 차수 폴더가 아닙니다. 1차·2차 등 여러 결과를 합치고 최종 정제한 파일을 사용자가 직접 넣는 공용 작업 폴더입니다.

**## 6. 최초 1회 환경 설정**

**### 6.1 프로젝트 폴더로 이동**

PowerShell에서 프로젝트가 저장된 폴더로 이동합니다.

\`\`\`powershell
cd "C:\프로젝트가\_저장된\_경로"
\`\`\`

예시:

\`\`\`powershell
cd "C:\Users\사용자명\Desktop\Local\_Campaign\_Automation"
\`\`\`

명령어를 실행하기 전에 현재 PowerShell 위치가 프로젝트 루트인지 확인합니다.

\`\`\`powershell
pwd
\`\`\`

**### 6.2 Python 가상환경 생성 및 라이브러리 설치**

배포받은 사용자는 아래 두 방법 중 하나를 선택합니다. \`pip\` 방식이 가장 보편적이며, \`uv\`는 선택 사항입니다.

**#### 방법 A. Python + pip 사용 — 기본 권장**

가상환경 생성:

\`\`\`powershell
python -m venv .venv
\`\`\`

가상환경 활성화:

\`\`\`powershell
.\\.venv\Scripts\Activate.ps1
\`\`\`

필요한 패키지 일괄 설치:

\`\`\`powershell
python -m pip install -r requirements.txt
\`\`\`

\`No module named pip\` 오류가 발생하면:

\`\`\`powershell
python -m ensurepip --upgrade
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
\`\`\`

**#### 방법 B. uv 사용 — 선택 사항**

가상환경 생성:

\`\`\`powershell
uv venv
\`\`\`

\`requirements.txt\` 설치:

\`\`\`powershell
uv pip install -r requirements.txt
\`\`\`

가상환경을 활성화하지 않고 실행하려면:

\`\`\`powershell
uv run python run\_pipeline.py
\`\`\`

**### 6.3 requirements.txt**

\`requirements.txt\`는 프로젝트 루트, 즉 \`run\_pipeline.py\`와 같은 위치에 둡니다.

현재 파이프라인의 주요 외부 패키지는 다음과 같습니다.

\`\`\`text
pandas
requests
openpyxl
python-dotenv
google-genai
gallery-dl
\`\`\`

Python 기본 라이브러리인 \`pathlib\`, \`datetime\`, \`json\`, \`subprocess\`, \`sys\`, \`re\`, \`os\` 등은 \`requirements.txt\`에 작성하지 않습니다.

**### 6.4 Google Cloud 인증**

Google Cloud 인증은 다음 두 종류로 구분합니다.

**#### 최초 계정 로그인 또는 계정 인증 만료 시**

\`\`\`powershell
gcloud auth login
\`\`\`

현재 로그인된 계정 확인:

\`\`\`powershell
gcloud auth list
\`\`\`

현재 설정된 Google Cloud 프로젝트 확인:

\`\`\`powershell
gcloud config get-value project
\`\`\`

예상 출력:

\`\`\`text
slcc-buzz-agent-dev
\`\`\`

필요한 경우 프로젝트를 설정합니다.

\`\`\`powershell
gcloud config set project slcc-buzz-agent-dev
\`\`\`

**#### PowerShell에서 \`gcloud.ps1\` 실행 오류가 발생하는 경우**

다음과 같은 오류가 발생할 수 있습니다.

\`\`\`text
이 시스템에서 스크립트를 실행할 수 없으므로 gcloud.ps1 파일을 로드할 수 없습니다.
\`\`\`

이 경우 \`gcloud.cmd\`를 직접 실행합니다.

계정 로그인:

\`\`\`powershell
& "C:\Users\사용자명\AppData\Local\Google\Cloud SDK\google-cloud-sdk\bin\gcloud.cmd" auth login
\`\`\`

매일 Application Default Credentials 갱신:

\`\`\`powershell
& "C:\Users\사용자명\AppData\Local\Google\Cloud SDK\google-cloud-sdk\bin\gcloud.cmd" auth application-default login
\`\`\`

로그인 상태 확인:

\`\`\`powershell
& "C:\Users\사용자명\AppData\Local\Google\Cloud SDK\google-cloud-sdk\bin\gcloud.cmd" auth list
\`\`\`

다음 오류가 발생하면 계정 로그인과 Application Default Credentials 인증을 모두 다시 수행합니다.

\`\`\`text
Reauthentication failed
cannot prompt during non-interactive execution
DefaultCredentialsError
\`\`\`

**---**

**## 7. 환경변수 및 설정값**

프로젝트에서 \`.env\` 파일을 사용하는 경우 프로젝트 루트에 \`.env\` 파일을 생성합니다.

환경변수는 다음과 같습니다.

\`\`\`env
SPRINKLR\_API\_KEY=
SPRINKLR\_ACCESS\_TOKEN=
\`\`\`

다음 항목은 실제 코드에서 사용하는 변수명을 확인한 후 수정해야 합니다.

**### 7.1 보안 주의사항**

다음 파일과 정보는 Git, 메일 또는 메신저를 통해 공유하지 않습니다.

\* \`.env\`
\* Sprinklr API Key
\* Sprinklr Access Token
\* Google Cloud 서비스 계정 JSON Key
\* 고객 데이터가 포함된 Excel
\* 다운로드된 고객 캠페인 미디어
\* 인증 토큰이 포함된 로그

권장 \`.gitignore\` 예시는 다음과 같습니다.

\`\`\`gitignore
.env
.venv/
\_\_pycache\_\_/
\*.pyc
logs/
output/
service-account\*.json
credentials\*.json
\`\`\`

JSON Schema 파일까지 제외되지 않도록 \`\*.json\` 전체를 \`.gitignore\`에 등록하지 않는 것을 권장합니다.

**---**

**## 8. 실행 전 공통 확인사항**

각 모듈을 실행하기 전에 다음 사항을 확인합니다.

\* PowerShell의 현재 위치가 프로젝트 루트인지 확인
\* Python 환경이 정상적으로 구성되어 있는지 확인
\* \`requirements.txt\`의 패키지가 설치되어 있는지 확인
\* 1절의 필수 운영 규칙을 완료했는지 확인
\* 실행 대상 날짜가 올바른지 확인
\* 입력 Excel 파일이 지정된 경로에 존재하는지 확인
\* Prompt, JSON Schema 및 \`config/country\_subsidiary\_mapping.json\` 파일이 존재하는지 확인
\* Google Cloud 계정과 프로젝트가 올바른지 확인
\* 기존 결과 파일이 있는 경우 덮어쓰기 여부 확인

**---**

**## 9. 모듈별 실행 방법**

**### 9.1 \`sprinklr\_export\_excel.py\`**

**#### 목적**

Sprinklr API를 호출하여 Local Campaign Raw 데이터를 조회하고 Raw Excel 파일로 저장합니다.

**#### 주요 역할**

\* Sprinklr Reporting API 호출
\* 원문 데이터 조회
\* 전략법인 데이터 조회
\* 게시물 본문 추출
\* 게시물 URL 추출
\* 생성 시간 추출
\* 플랫폼 정보 추출
\* 작성자와 Sender Profile의 Screen Name, 팔로워 수, 위치, Bio, Website, 인증 정보 및 Profile Tags 추출
\* 미디어 URL 및 미디어 타입 추출
\* Raw Excel 생성

**#### 주요 출력 시트**

\`\`\`text
Raw Data\_원문
Raw Data\_전략법인
\`\`\`

각 시트에서 사용되는 주요 컬럼은 다음과 같습니다.

\`\`\`text
Conversation Stream
Campaign ID
Profile URL
User Name
Permalink
Created Time
snType column 
Media Type
Media URL
source\_widget 
data\_cut\_start
data\_cut\_end 
extracted\_at
Sender Profile Available
Sender Screen Name
Sender Follower Count
Sender Location
Sender Detailed Location
Sender Bio
Sender Website
Sender Verified
Sender Verified Type
Sender Profile Tags
\`\`\`

\`Sender SN ID\`와 \`Sender Universal Profile ID\`는 현재 캠페인 분석 및 Country–Subsidiary 판정에 사용하지 않으므로 Raw Excel 추출 대상에서 제외합니다.

\`Raw Data\_전략법인\` 시트에는 다음 컬럼이 추가로 포함될 수 있습니다.

\`\`\`text
Author Screen Name
\`\`\`

**#### 실행 전 설정**

입력 형식과 전체 실행 방법은 **\*\*3.1절\*\***을 따릅니다. 조회 종료 시각의 초는 반드시 \`59\`로 입력합니다.

\`\`\`text
예: 2026-07-24 18:00:59
\`\`\`

**#### 개별 실행 명령어 — 1단계 오류 복구용**

사용할 실행 차수의 Output 폴더를 명시합니다.

\`\`\`powershell
python sprinklr\_export\_excel.py \`
  --output-dir "output\260805\_2차"
\`\`\`

\`uv\` 환경:

\`\`\`powershell
uv run python sprinklr\_export\_excel.py \`
  --output-dir "output\260805\_2차"
\`\`\`

**#### 출력 결과**

출력 구조:

\`\`\`text
output/
└─ YYMMDD 또는 YYMMDD\_N차/
   └─ {YYMMDD}\_SLCC\_SOV\_Local Campaign Tracking\_{월}월\_v01.xlsx
\`\`\`

**#### 정상 실행 확인**

모듈별 결과 검증 항목은 **\*\*15절 실행 완료 후 검증 체크리스트\*\***를 확인합니다.

**#### 주의사항**

\* Sprinklr Access Token이 만료된 경우 API 호출이 실패할 수 있습니다.
\* 조회 기간이 너무 길면 API 응답 시간이 길어질 수 있습니다.
\* Sprinklr 응답 구조가 변경되면 일부 컬럼이 비어 있을 수 있습니다.
\* API 응답의 일부 미디어 URL은 일정 시간이 지나면 만료될 수 있습니다.

**---**

**### 9.2 \`raw\_to\_processed.py\`**

**#### 목적**

Sprinklr에서 생성된 Raw Excel을 읽고, 최종 Local Campaign 정리본에 필요한 컬럼 생성 및 1차 정제 후 데이터를 기입합니다.

**#### 주요 역할**

\* Raw Excel 파일 읽기
\* 필요한 Sheet 및 컬럼 검증
\* 컬럼명 및 값 정규화
\* Raw Data 기반으로 데이터 정제 (Hashtags, URL, Query 컬럼 생성)
\* 불필요하거나 유효하지 않은 행 제외
\* {YYMMDD}\_SLCC\_SOV\_Local Campaign Tracking\_{월}월\_v01.xlsx에 새로운 Sheet (로컬 캠페인 리스트\_QHB8) 생성 

**#### 주요 입력**

이전 단계에서 생성된 Raw Excel입니다.

\`\`\`text
sprinklr\_export\_excel.py 결과 Excel:
{YYMMDD}\_SLCC\_SOV\_Local Campaign Tracking\_{월}월\_v01.xlsx
\`\`\`

**#### 주요 출력 컬럼**

\`\`\`text
\#
Campaign Date 
Campaign Image
Subsidiary (Country) / Influencer (Subsidiary)
Campaign Name
Product
CXP Product Feature 
Description
Buzz Volume
Channel
Giveaway 
Influencer 
HTR/DE
Conv.Card
Hashtags
URL
비고
Query 
\`\`\`

이 단계에서 자동으로 값이 채워지는 컬럼은 다음과 같습니다.

\`\`\`
\#
Campaign Date
Channel
Influencer
Hashtags
URL
Query
\`\`\`

**#### 개별 실행 명령어 — 2단계 오류 복구용**

사용할 실행 차수의 Output 폴더를 명시합니다.

\`\`\`powershell
python raw\_to\_processed.py \`
  --output-dir "output\260805\_2차"
\`\`\`

\`uv\` 환경:

\`\`\`powershell
uv run python raw\_to\_processed.py \`
  --output-dir "output\260805\_2차"
\`\`\`

**#### 출력 결과**

출력 구조:

\`\`\`text
output/
└─ YYMMDD 또는 YYMMDD\_N차/
   └─ {YYMMDD}\_SLCC\_SOV\_Local Campaign Tracking\_{월}월\_v01\_formatted.xlsx
\`\`\`

**#### 정상 실행 확인**

모듈별 결과 검증 항목은 **\*\*15절 실행 완료 후 검증 체크리스트\*\***를 확인합니다.

**#### 주의사항**

\* Excel 시트명은 코드에서 사용하는 이름과 정확히 일치해야 합니다.
\* 필수 컬럼의 공백이나 오탈자로 인해 실행이 실패할 수 있습니다.
\* 이 단계에서 생성된 결과 파일은 다음 단계에서 직접 사용되므로 파일명을 임의로 변경하지 않는 것을 권장합니다.

**---**

**### 9.3 \`media\_extractor.py\`**

**#### 목적**

Raw Excel에 저장된 게시물 URL과 미디어 정보를 이용하여 게시물별 이미지 및 영상 파일을 로컬 환경에 저장합니다.

**#### 주요 역할**

\* Raw Excel 읽기
\* 분석 대상 게시물 구조화
\* 플랫폼 식별
\* 미디어 타입 식별
\* 미디어 URL 정규화
\* 플랫폼별 추출 전략 선택
\* 이미지 및 영상 다운로드
\* 게시물과 로컬 파일 연결
\* 다운로드 결과 및 실패 사유 기록

**#### 지원 대상 플랫폼**

현재 대상 플랫폼은 다음과 같습니다.

\| 플랫폼       | 주요 처리 방식                             | 주요 제한사항                 |
\| --------- | ------------------------------------ | ----------------------- |
\| YouTube   | 게시물 URL 유지 또는 분석 단계 직접 전달            | 삭제·비공개 영상 처리 불가         |
\| Instagram | Sprinklr Media URL 또는 Child Media 사용 | Media URL 만료 가능 (특히 Reels)         |
\| Twitter/X | Media URL 또는 \`gallery-dl\` 사용         | 게시물 상태와 접근 권한에 따라 실패 가능 |
\| Facebook  | 공개 게시물 및 Media URL 기반                | URL 구조와 권한에 따라 실패 가능    |

**#### Twitter/X 처리 원칙**

Twitter/X 게시물에서 기존 Media URL만으로 파일을 확보할 수 없는 경우 \`gallery-dl\`을 사용합니다.

설치 확인:

\`\`\`powershell
gallery-dl --version
\`\`\`

수동 테스트:

\`\`\`powershell
gallery-dl -v "게시물\_URL"
\`\`\`

**#### 미디어 저장 폴더 규칙**

최상위 폴더는 사용자가 입력한 날짜로 생성됩니다.

예시:

\`\`\`text
media/
└─ 260724/
\`\`\`

날짜 폴더 내부에 원본 Sheet 이름을 기준으로 하위 폴더가 생성됩니다.

\`\`\`text
media/
└─ 260724/
   ├─ Raw Data\_원문/
   └─ Raw Data\_전략법인/
\`\`\`

**#### 파일명 규칙**

기본 파일명은 다음 형식을 사용합니다.

\`\`\`text
{source\_sheet\_name}\_{raw\_row\_number}.{extension}
\`\`\`

예시:

\`\`\`text
Raw Data\_원문\_15.mp4
Raw Data\_전략법인\_32.jpg
\`\`\`

한 게시물에서 여러 개의 미디어가 추출되는 경우 첫 번째 미디어에는 별도의 번호를 붙이지 않고, 두 번째 미디어부터 \`\_02\`, \`\_03\` 순서로 번호를 붙입니다.

파일명 규칙:

\`\`\`text
첫 번째 미디어:
{source\_sheet\_name}\_{raw\_row\_number}.{extension}

두 번째 이후 미디어:
{source\_sheet\_name}\_{raw\_row\_number}\_{media\_index\_two\_digit}.{extension}
\`\`\`

예시:

\`\`\`text
Raw Data\_원문\_15.jpg
Raw Data\_원문\_15\_02.jpg
Raw Data\_원문\_15\_03.mp4
\`\`\`

**#### 개별 실행 명령어 — 3단계 오류 복구용**

Output과 Media에 동일한 실행 차수를 명시합니다.

\`\`\`powershell
python media\_extractor.py \`
  --output-dir "output\260805\_2차" \`
  --media-dir "media\260805\_2차"
\`\`\`

\`uv\` 환경:

\`\`\`powershell
uv run python media\_extractor.py \`
  --output-dir "output\260805\_2차" \`
  --media-dir "media\260805\_2차"
\`\`\`

**#### 출력 결과**

\`\`\`text
media/
└─ YYMMDD 또는 YYMMDD\_N차/
   ├─ Raw Data\_원문/
   │  ├─ Raw Data\_원문\_행번호.jpg
   │  └─ Raw Data\_원문\_행번호.mp4
   │
   └─ Raw Data\_전략법인/
      ├─ Raw Data\_전략법인\_행번호.jpg
      └─ Raw Data\_전략법인\_행번호.mp4
\`\`\`

추출된 미디어와 결과 Excel은 명령어로 지정한 동일 실행 차수의 \`media\` 및 \`output\` 폴더에 생성됩니다.

파일 이름:

\`\`\`text
{YYMMDD}\_campaign\_media\_result.xlsx
\`\`\`

**#### 정상 실행 확인**

모듈별 결과 검증 항목은 **\*\*15절 실행 완료 후 검증 체크리스트\*\***를 확인합니다.

**#### 주의사항**

플랫폼별 접근 제한과 URL 만료 이슈는 **\*\*13절 플랫폼별 알려진 제한사항\*\***을 참고합니다.

\* 다운로드된 파일이 존재하더라도 실제 이미지 또는 영상 파일이 손상되었을 수 있습니다.
\* 파일 크기가 지나치게 작은 경우 오류 응답이 파일로 저장된 것은 아닌지 확인해야 합니다.
\* 같은 날짜와 같은 행 번호로 다시 실행할 경우 기존 파일이 덮어써질 수 있으므로 재실행 전에 기존 파일 처리 정책을 확인합니다.

**---**

**### 9.4 \`llm\_analysis\_pipeline.py\`**

**#### 목적**

Processed Excel과 로컬 미디어 파일을 연결하고, 사전에 정의된 Prompt 및 JSON Schema와 함께 LLM에 전달하여 구조화된 캠페인 분석 결과를 생성합니다 (Processed Excel에 비어있던 컬럼들이 다수 채워집니다).

**#### 주요 역할**

\* Processed Excel 읽기
\* 게시물별 LLM Input Set 생성
\* 게시물과 미디어 파일 연결
\* Prompt 파일 읽기
\* JSON Schema 파일 읽기
\* Gemini API 호출
\* 응답 JSON 파싱
\* 결과 Schema 검증
\* 성공 및 실패 결과 정리
\* 최종 분석 결과 저장

**#### LLM Input Set 구조**

\`\`\`text
campaign\_id
source\_sheet\_name
raw\_row\_number
platform
conversation\_stream
profile\_name
profile\_url
sender\_profile\_available
sender\_screen\_name
sender\_follower\_count
sender\_location
sender\_detailed\_location
sender\_bio
sender\_website
sender\_verified
sender\_verified\_type
sender\_profile\_tags
original\_post\_url
media\_paths
\`\`\`

\`sender\_follower\_count\`는 **\*\*현재 게시 플랫폼에서 게시물을 업로드한 계정의 팔로워 수\*\***입니다. 콘텐츠에 등장하는 배우나 인플루언서의 팔로워 수로 자동 사용하지 않습니다.

\`sender\_sn\_id\`와 \`sender\_universal\_profile\_id\`는 LLM Input Set 및 Gemini 입력에서 제거합니다.

각 게시물의 미디어는 다음 조합을 기준으로 연결됩니다.

\`\`\`text
campaign\_id + source\_sheet\_name + raw\_row\_number
\`\`\`

**#### Prompt 파일**

Prompt 경로:

\`\`\`text
prompts/user\_prompt.txt
\`\`\`

**#### JSON Schema 파일**

Schema 경로:

\`\`\`text
prompts/response\_schema.json
\`\`\`

LLM은 다음 9개 필드를 반환합니다.

\`\`\`text
Campaign Name
Product
CXP Product Feature
Description
Publisher Type
Publisher Country
Publisher Classification Reason
Publisher Classification Confidence
Requires Manual Review
\`\`\`

**#### 게시자 및 Subsidiary 확정 규칙**

게시자 유형은 콘텐츠에 등장하는 인물이 아니라 **\*\*게시물을 업로드한 계정\*\***을 기준으로 판단합니다.

\`\`\`text
SamsungGulf 계정
→ Publisher Type: OWNED
→ Publisher Country: Dubai
→ Subsidiary: SGE
→ URL 헤더: [당사 게시글]
→ Influencer: No

Samsung Korea / SamsungKorea / Samsung Korea 삼성전자 계정
→ Publisher Type: OWNED
→ Publisher Country: Korea
→ Subsidiary: KOREA
→ URL 헤더: [당사 게시글]
→ Influencer: No
\`\`\`

Country–Subsidiary 매핑은 다음 파일의 정확한 Country 문자열을 기준으로 처리합니다.

\`\`\`text
config/country\_subsidiary\_mapping.json
\`\`\`

예:

\`\`\`text
Dubai → SGE
Korea → KOREA
\`\`\`

\`South Korea\`는 현재 매핑 key가 아니므로 사용하지 않으며, Samsung Korea 계정은 반드시 \`Korea\`를 반환합니다.

공식 Samsung 계정의 캠페인 미디어에 배우·가수·크리에이터·인플루언서가 등장하고 신원과 직업의 내부 판단 신뢰도가 99 이상이면 Description에 해당 인물 정보를 포함할 수 있습니다. 그러나 게시 계정이 공식 Samsung 계정이면 다음 결과는 변하지 않습니다.

\`\`\`text
Publisher Type = OWNED
URL 헤더 = [당사 게시글]
Influencer = No
\`\`\`

인물의 직업이 배우·가수·모델·스포츠 선수 등으로 확인되면 일반적인 \`인플루언서\`보다 구체적인 직업명을 우선합니다.

팔로워 수는 현재 게시 플랫폼과 일치하는 경우에만 사용합니다.

\`\`\`text
Twitter/X 게시물 → X 팔로워
Instagram 게시물 → IG 팔로워
YouTube 게시물 → YouTube 팔로워
TikTok 게시물 → TikTok 팔로워
Facebook 게시물 → Facebook 팔로워
\`\`\`

다른 플랫폼의 팔로워 수를 현재 게시 플랫폼의 팔로워 수로 바꾸거나 대체하지 않습니다. 공식 Samsung 계정의 \`sender\_follower\_count\`를 콘텐츠 속 등장 인물의 팔로워 수로 사용하지 않습니다.

**#### 실행 전 인증**

매일 실행 전에 다음 명령을 수행합니다.

\`\`\`powershell
gcloud auth application-default login
\`\`\`

계정 로그인, 프로젝트 설정, 인증 만료 및 PowerShell 오류 대응은 **\*\*6.4절 Google Cloud 인증\*\***을 참고합니다.

**#### 개별 실행 명령어 — 4단계 오류 복구용**

사용할 실행 차수의 Output 폴더를 명시합니다.

\`\`\`powershell
python llm\_analysis\_pipeline.py \`
  --output-dir "output\260805\_2차"
\`\`\`

\`uv\` 환경:

\`\`\`powershell
uv run python llm\_analysis\_pipeline.py \`
  --output-dir "output\260805\_2차"
\`\`\`

**#### 출력 결과**

저장 구조:

\`\`\`text
output/
└─ YYMMDD 또는 YYMMDD\_N차/
   └─ LLM 분석 결과 Excel
\`\`\`

**#### 정상 실행 확인**

모듈별 결과 검증 항목은 **\*\*15절 실행 완료 후 검증 체크리스트\*\***를 확인합니다.

**#### 주의사항**

\* Google Cloud 프로젝트 또는 Region 설정이 잘못되면 모델을 찾지 못할 수 있습니다.
\* Prompt와 JSON Schema의 구조가 서로 맞지 않으면 응답 파싱이 실패할 수 있습니다.
\* 이미지나 영상 파일이 손상되면 해당 게시물 분석이 실패할 수 있습니다.
\* 한 게시물에 여러 미디어가 존재하는 경우 모든 파일이 하나의 Input Set에 포함되는지 확인해야 합니다.
\* 기존 성공 결과를 자동으로 건너뛰지 않으면 전체 행이 다시 호출될 수 있습니다.
\* LLM 응답은 항상 완전히 동일하지 않을 수 있으므로 **\*\*14절의 소량 테스트 절차\*\***와 **\*\*15절의 검증 체크리스트\*\***를 따릅니다.

**---**

**---**

**### 9.5 \`buzz\_volume\_adaptor.py\`**

**#### 목적**

1차·2차 등 여러 실행 결과를 통합하고 수동 정제한 최종 Excel에 캠페인별 Buzz Volume을 적재합니다.

이 모듈은 \`run\_pipeline.py\`의 1\~4단계에 포함되지 않으며 별도로 실행합니다.

**#### 입력 파일 준비**

Buzz Volume 이외의 항목이 모두 완료된 최종 통합·정제 Excel을 다음 폴더에 넣습니다.

\`\`\`text
output\Buzz\_Volume\\
{YYMMDD}\_SLCC\_SOV\_Local Campaign Tracking\_{월}월\_v01.xlsx
\`\`\`

예시:

\`\`\`text
output\Buzz\_Volume\\
260805\_SLCC\_SOV\_Local Campaign Tracking\_8월\_v01.xlsx
\`\`\`

파일명은 기준 날짜와 일치해야 합니다. 모듈은 사용자가 입력한 \`YYYY-MM-DD\`를 기준으로 \`YYMMDD\`와 월을 계산하여 파일을 자동으로 찾습니다.

**#### 개별 실행 명령어**

\`\`\`powershell
python buzz\_volume\_adaptor.py
\`\`\`

\`uv\` 환경:

\`\`\`powershell
uv run python buzz\_volume\_adaptor.py
\`\`\`

Buzz Volume 모듈은 공용 \`output\Buzz\_Volume\` 경로를 사용하므로 \`--output-dir\` 또는 \`--media-dir\`을 입력하지 않습니다.

**#### 실행 입력**

\`\`\`text
데이터 컷 유형을 입력하세요 (daily / weekly):
기준 날짜를 입력하세요 (YYYY-MM-DD):
\`\`\`

Daily 예시:

\`\`\`text
daily
2026-08-05
\`\`\`

Weekly 실행 시 기준 날짜는 반드시 월요일이어야 하며, 직전 월요일부터 일요일까지의 캠페인을 처리합니다.

**#### 처리 방식**

\* Excel에서 Query가 있는 행만 처리
\* Weekly 실행 시 직전 주 Campaign Date 범위로 추가 필터링
\* 최대 5개 Worker를 이용한 병렬 API 호출
\* 초당 요청 시작 횟수 제한
\* 요청별 자동 재시도
\* 배치별 체크포인트 저장
\* 최종 실패 행은 \`API\_Failed\`로 기록
\* 기존 Buzz Volume 값은 새 API 결과로 덮어쓰기

**#### 출력 결과**

\`\`\`text
output\Buzz\_Volume\completed\\
{YYMMDD}\_SLCC\_SOV\_Local Campaign Tracking\_{월}월\_v01\_mentions\_updated.xlsx
\`\`\`

실패 상세 JSON:

\`\`\`text
output\Buzz\_Volume\artifacts\failed\_results\\
\`\`\`

Mention 값을 파싱하지 못한 API 응답 샘플:

\`\`\`text
output\Buzz\_Volume\artifacts\response\_samples\\
\`\`\`

**#### 주의사항**

\* 입력 파일은 날짜별 실행 폴더가 아니라 \`output\Buzz\_Volume\`에 넣습니다.
\* 입력 원본은 수정하지 않고 \`completed\` 폴더에 별도 결과를 생성합니다.
\* 기준 날짜와 입력 파일명의 \`YYMMDD\` 및 월이 일치해야 합니다.
\* Daily와 Weekly가 사용하는 대상 시트명이 다릅니다.
\* 일부 캠페인이 최종 실패하면 성공값과 \`API\_Failed\`를 Excel에 저장한 뒤 오류로 종료할 수 있습니다.

**## 10. 단계별 입력 및 출력 연결 관계**

\| 단계 | 실행 파일 | 입력 | 출력 | 다음 단계 사용 여부 |
\| --- | --- | --- | --- | --- |
\| 1 | \`sprinklr\_export\_excel.py\` | Sprinklr API 및 조회 기간 | Raw Excel | 2·3단계 입력 |
\| 2 | \`raw\_to\_processed.py\` | 같은 실행 차수의 Raw Excel | Processed Excel | 4단계 입력 |
\| 3 | \`media\_extractor.py\` | 같은 실행 차수의 Raw Excel 및 게시물 URL | Media Result Excel 및 로컬 이미지·영상 | 4단계 입력 |
\| 4 | \`llm\_analysis\_pipeline.py\` | Processed Excel, Media Result, 미디어, Prompt, Schema | LLM 분석 결과 | 차수별 분석 산출물 |
\| 수동 통합·정제 | 사용자 작업 | 1차·2차 등 여러 결과 | Buzz Volume 이외 항목이 완료된 통합 Excel | 5단계 입력 |
\| 5 | \`buzz\_volume\_adaptor.py\` | \`output/Buzz\_Volume\`의 통합·정제 Excel | \`completed\`의 Buzz Volume 적재 완료 Excel | 최종 산출물 |

1\~4단계는 실행 차수별 Output/Media 폴더를 사용합니다. 5단계는 실행 차수와 독립적인 공용 \`output/Buzz\_Volume\` 폴더를 사용합니다.

**## 11. 일일 실행 예시**

다음은 \`2026-08-05\` 데이터를 처리하는 실행 예시입니다. 최초 환경 구성은 **\*\*6절\*\***을 참고합니다.

**### Step 1. 프로젝트 폴더로 이동**

\`\`\`powershell
cd "C:\Users\사용자명\Desktop\Local\_Campaign\_Automation
\`\`\`

가상환경을 사용하는 경우 활성화합니다.

\`\`\`powershell
.\\.venv\Scripts\Activate.ps1
\`\`\`

**### Step 2. Application Default Credentials 갱신**

\`\`\`powershell
gcloud auth application-default login
\`\`\`

계정 인증이 만료된 경우에는 \`gcloud auth login\`도 수행합니다.

**### Step 3. 전체 파이프라인 1\~4단계 실행**

\`\`\`powershell
python run\_pipeline.py
\`\`\`

\`uv\`를 사용하는 경우:

\`\`\`powershell
uv run python run\_pipeline.py
\`\`\`

입력 예시:

\`\`\`text
조회 시작 시각: 2026-08-05 09:00:00
조회 종료 시각: 2026-08-05 12:00:59
자동 생성 작업 날짜: 260805
\`\`\`

실행 로그에 표시되는 경로를 기록합니다.

\`\`\`text
실행 차수: 2차
Output 폴더: ...\output\260805\_2차
Media 폴더: ...\media\260805\_2차
\`\`\`

**### Step 4. 1\~4단계 결과 확인**

생성 파일과 성공·실패 건수는 **\*\*15절 실행 완료 후 검증 체크리스트\*\***에 따라 확인합니다.

특정 단계가 실패하면 로그의 정확한 Output/Media 경로를 사용해 해당 단계부터 다시 실행합니다.

예시:

\`\`\`powershell
python media\_extractor.py \`
  --output-dir "output\260805\_2차" \`
  --media-dir "media\260805\_2차"
\`\`\`

**### Step 5. 여러 차수 결과 통합 및 최종 정제**

필요한 1차·2차 결과를 통합하고 Buzz Volume 이외의 항목을 최종 검수합니다.

완료된 파일을 다음 이름으로 저장합니다.

\`\`\`text
260805\_SLCC\_SOV\_Local Campaign Tracking\_8월\_v01.xlsx
\`\`\`

그리고 다음 공용 폴더에 넣습니다.

\`\`\`text
output\Buzz\_Volume\\
\`\`\`

**### Step 6. Buzz Volume 별도 실행**

\`\`\`powershell
python buzz\_volume\_adaptor.py
\`\`\`

입력 예시:

\`\`\`text
데이터 컷 유형: daily
기준 날짜: 2026-08-05
\`\`\`

**### Step 7. 최종 결과 확인**

\`\`\`text
output\Buzz\_Volume\completed\\
260805\_SLCC\_SOV\_Local Campaign Tracking\_8월\_v01\_mentions\_updated.xlsx
\`\`\`

일부 행이 실패한 경우 \`Buzz Volume\`에 \`API\_Failed\`가 기록되었는지 확인하고, \`output\Buzz\_Volume\artifacts\`의 실패 자료를 검토합니다.

**## 12. Troubleshooting**

**### 12.1 \`python\` 명령어를 찾을 수 없는 경우**

\`\`\`powershell
python --version
\`\`\`

Python이 인식되지 않는 경우 Python 설치 상태와 환경변수 PATH를 확인합니다.

\`uv\`가 설치되어 있다면 다음 명령어도 확인합니다.

\`\`\`powershell
uv python list
\`\`\`

**### 12.2 \`uv\` 명령어를 찾을 수 없는 경우**

\`\`\`powershell
uv --version
\`\`\`

설치 후 PowerShell 또는 VS Code를 완전히 종료하고 다시 실행합니다.

**### 12.3 Python 모듈을 찾을 수 없는 경우**

예시 오류:

\`\`\`text
ModuleNotFoundError: No module named '패키지명'
\`\`\`

pip 사용자:

\`\`\`powershell
python -m pip install -r requirements.txt
\`\`\`

uv 사용자:

\`\`\`powershell
uv pip install -r requirements.txt
\`\`\`

특정 패키지만 설치해야 하는 경우:

\`\`\`powershell
python -m pip install 패키지명
\`\`\`

또는:

\`\`\`powershell
uv pip install 패키지명
\`\`\`

**### 12.4 조회 종료 분의 데이터가 누락되는 경우**

조회 종료 시각의 초가 \`59\`인지 확인합니다.

\`\`\`text
올바른 입력: 2026-07-27 16:30:59
잘못된 입력: 2026-07-27 16:30:00
\`\`\`

종료 초를 \`00\`으로 입력하면 \`16:30:01\`부터 \`16:30:59\` 사이에 생성된 데이터가 조회 범위에서 제외될 수 있습니다.

**### 12.5 입력 파일을 찾을 수 없는 경우**

예시 오류:

\`\`\`text
FileNotFoundError
\`\`\`

다음 항목을 확인합니다.

\* 1\~4단계 단독 실행에서는 \`--output-dir\`과 필요한 경우 \`--media-dir\`이 실제 실행 차수 폴더를 가리키는지 확인
\* \`run\_pipeline.py\` 로그에 출력된 Output/Media 경로와 명령어 경로가 일치하는지 확인
\* 입력 Excel 파일명이 코드의 예상 파일명과 일치하는지 확인
\* Buzz Volume 입력은 \`output\Buzz\_Volume\`에 존재하는지 확인
\* Buzz Volume 기준 날짜와 파일명의 \`YYMMDD\` 및 월이 일치하는지 확인
\* PowerShell의 실행 위치가 프로젝트 루트인지 확인
\* 경로에 오탈자가 없는지 확인
\* 확장자가 \`.xlsx\`인지 확인

**### 12.6 PowerShell에서 \`gcloud\` 실행 오류가 발생하는 경우**

다음 명령어로 \`gcloud.cmd\`를 직접 실행합니다.

\`\`\`powershell
& "C:\Users\사용자명\AppData\Local\Google\Cloud SDK\google-cloud-sdk\bin\gcloud.cmd" auth login
\`\`\`

**### 12.7 Google Cloud 인증 오류**

예시 오류:

\`\`\`text
Reauthentication failed
cannot prompt during non-interactive execution
\`\`\`

계정 로그인과 Application Default Credentials 인증을 다시 수행합니다.

\`\`\`powershell
gcloud auth login
gcloud auth application-default login
\`\`\`

계정을 확인합니다.

\`\`\`powershell
gcloud auth list
\`\`\`

필요한 계정을 활성화합니다.

\`\`\`powershell
gcloud config set account 계정이메일
\`\`\`

**### 12.8 Google Cloud 프로젝트 오류**

현재 프로젝트 확인:

\`\`\`powershell
gcloud config get-value project
\`\`\`

프로젝트 설정:

\`\`\`powershell
gcloud config set project PROJECT\_ID
\`\`\`

**### 12.9 미디어 추출 실패**

다음 항목을 확인합니다.

\* 게시물 URL이 정상적으로 열리는지 확인
\* 게시물이 삭제 또는 비공개 상태인지 확인
\* Instagram Media URL이 만료되었는지 확인
\* \`gallery-dl\`이 정상 설치되어 있는지 확인
\* Twitter/X 게시물 URL을 \`gallery-dl\`로 직접 테스트
\* 저장 폴더에 쓰기 권한이 있는지 확인
\* 파일명이 Windows에서 허용되는 형식인지 확인

Twitter/X 수동 테스트:

\`\`\`powershell
gallery-dl -v "게시물\_URL"
\`\`\`

**### 12.10 LLM 응답 JSON 파싱 실패**

다음 항목을 확인합니다.

\* Prompt 파일 경로
\* JSON Schema 파일 경로
\* Prompt가 JSON 형식 응답을 명확하게 요구하는지 확인
\* LLM 응답에 JSON 외의 설명 문장이 포함되었는지 확인
\* 필수 필드가 누락되었는지 확인
\* 응답 값의 타입이 Schema와 일치하는지 확인
\* 응답 원문을 별도 로그에 저장했는지 확인

**### 12.11 LLM 입력 미디어를 찾지 못하는 경우**

다음 세 값이 일치하는지 확인합니다.

\`\`\`text
campaign\_id
source\_sheet\_name
raw\_row\_number
\`\`\`

또한 다음 항목을 확인합니다.

\* 미디어 폴더 경로가 올바른지 확인
\* 미디어 파일명이 규칙에 맞는지 확인
\* 파일 확장자가 지원 대상인지 확인
\* Processed Excel의 행 번호와 실제 파일명의 행 번호가 같은지 확인

**### 12.12 전체 파이프라인이 중간에 중단되는 경우**

\`run\_pipeline.py\`는 앞 단계의 return code가 0이 아니면 즉시 중단됩니다.

\`\`\`text
RuntimeError: 모듈명 실행 실패 (return code: 1)
\`\`\`

이 경우 마지막으로 표시된 실패 모듈과 해당 모듈의 원본 오류 메시지를 확인합니다. 문제를 해결한 뒤, 이미 생성된 앞 단계 결과가 정상이라면 실패한 모듈부터 개별 실행할 수 있습니다.

예를 들어 \`260805\_2차\`의 3단계가 실패했다면 다음처럼 재실행합니다.

\`\`\`powershell
python media\_extractor.py \`
  --output-dir "output\260805\_2차" \`
  --media-dir "media\260805\_2차"
\`\`\`

3단계가 정상 완료된 뒤 4단계를 실행합니다.

\`\`\`powershell
python llm\_analysis\_pipeline.py \`
  --output-dir "output\260805\_2차"
\`\`\`

기존 결과 파일이 남아 있다면 재실행 전에 덮어쓰기 여부를 확인합니다. 현재 저장 방식은 동일 경로·동일 파일명에 새 결과를 저장할 경우 기존 파일 전체가 교체될 수 있습니다. 이는 기존 데이터와 신규 데이터를 비교해 추가하는 append 또는 upsert 방식이 아닙니다.

**---**

**## 13. 플랫폼별 알려진 제한사항**

**### YouTube**

\* 삭제된 영상은 분석할 수 없습니다.
\* 비공개 영상은 분석할 수 없습니다.
\* 지역 또는 연령 제한 영상은 접근이 제한될 수 있습니다.
\* 긴 영상은 처리 시간과 API 비용이 증가할 수 있습니다.
\* URL 직접 전달 방식을 사용합니다.

**### Instagram**

\* Sprinklr에서 제공된 Media URL이 만료될 수 있습니다.
\* Reels URL에서 \`URL is expired\` 오류가 발생할 수 있습니다.
\* Carousel 게시물은 \`childMedias\`에 여러 이미지가 포함될 수 있습니다.
\* 게시물 URL이 존재하더라도 원본 미디어를 항상 다운로드할 수 있는 것은 아닙니다.

**### Twitter/X**

\* Twitter/X 공식 API는 사용하지 않습니다.
\* 필요한 경우 \`gallery-dl\`을 사용합니다.
\* 게시물이 삭제되거나 접근이 제한되면 다운로드할 수 없습니다.
\* 로그인이나 쿠키가 필요한 게시물은 자동 추출이 실패할 수 있습니다.
\* 게시물에 미디어가 없고 외부 링크만 있는 경우 별도 처리가 필요할 수 있습니다.

**### Facebook**

\* 공개 게시물을 기준으로 처리합니다.
\* 게시물 권한에 따라 미디어 다운로드가 실패할 수 있습니다.
\* Facebook URL 구조 변경 시 추출 로직 수정이 필요할 수 있습니다.

**---**

**## 14. 비용 및 실행량 주의사항**

LLM 분석 단계에서는 Google Cloud 및 Gemini 또는 Vertex AI 사용 비용이 발생할 수 있습니다.

대량 실행 전 다음 절차를 권장합니다.

1\. 1개 행으로 테스트
2\. 5\~10개 행으로 Prompt 및 결과 구조 검증
3\. 미디어가 여러 개인 게시물 테스트
4\. 이미지 게시물과 영상 게시물 각각 테스트
5\. 실패 로그 확인
6\. 전체 데이터 실행

**---**

**## 15. 실행 완료 후 검증 체크리스트**

**### Sprinklr 추출 결과**

\`\`\`text
[ ] Raw Excel 파일이 생성되었다.
[ ] Raw Data\_원문 시트가 존재한다.
[ ] Raw Data\_전략법인 시트가 존재한다.
[ ] 각 시트에 예상 데이터가 존재한다.
[ ] Permalink가 정상적으로 저장되었다.
[ ] Conversation Stream이 정상적으로 저장되었다.
[ ] Created Time이 정상적으로 저장되었다.
[ ] Sender Screen Name과 Sender Follower Count가 정상적으로 저장되었다.
[ ] Sender Location, Bio, Website, 인증 정보가 필요한 행에 저장되었다.
[ ] Sender SN ID와 Sender Universal Profile ID가 생성되지 않았다.
\`\`\`

**### 전처리 결과**

\`\`\`text
[ ] Processed Excel 파일이 생성되었다.
[ ] 로컬 캠페인 리스트\_QHB8 시트가 정상적으로 생성되었다.
[ ] 필수 컬럼들이 정상적으로 생성되었다.
[ ] Campaign Date가 정상적으로 변환되었다. 
[ ] Channel이 정상적으로 기입되었다.
[ ] Influencer가 정상적으로 기입되었다.
[ ] Hashtags가 정상적으로 기입되었다.
[ ] URL이 정상적으로 기입되었다.
[ ] Query가 정상적으로 기입되었다.
\`\`\`

**### 미디어 추출 결과**

\`\`\`text
[ ] 날짜 폴더가 생성되었다.
[ ] Raw Data\_원문 폴더가 생성되었다.
[ ] Raw Data\_전략법인 폴더가 생성되었다.
[ ] 이미지 파일이 정상적으로 열린다.
[ ] 영상 파일이 정상적으로 재생된다.
[ ] 파일명이 시트명과 행 번호 규칙을 따른다.
[ ] 성공 및 실패 건수를 확인했다.
[ ] 다중 미디어 게시물의 파일 누락 여부를 확인했다.
\`\`\`

**### LLM 분석 결과**

\`\`\`text
[ ] Google Cloud 인증이 정상적으로 완료되었다.
[ ] Prompt 파일이 정상적으로 로드되었다.
[ ] JSON Schema가 정상적으로 로드되었다.
[ ] LLM Input Set 개수가 예상 행 수와 일치한다.
[ ] 미디어 파일이 올바른 게시물에 연결되었다.
[ ] LLM 호출 성공 및 실패 건수를 확인했다.
[ ] JSON 파싱 실패 건을 확인했다.
[ ] 공식 Samsung 계정의 URL에 [당사 게시글] 헤더가 적용되었다.
[ ] 공식 Samsung 계정의 Influencer 값이 No로 기록되었다.
[ ] OWNED 계정의 팔로워 수가 콘텐츠 속 인물의 팔로워 수로 잘못 사용되지 않았다.
[ ] Description의 인물 직업과 팔로워 플랫폼이 입력 근거와 일치한다.
[ ] 최종 결과 파일이 생성되었다.
[ ] 입력 행 수와 성공·실패·제외 건수의 합이 일치한다.
\`\`\`


**### Buzz Volume 결과**

\`\`\`text
[ ] output/Buzz\_Volume 폴더가 존재한다.
[ ] 통합·정제 완료 입력 Excel이 기준 날짜와 일치하는 파일명으로 저장되었다.
[ ] Daily 또는 Weekly 데이터 컷을 올바르게 선택했다.
[ ] Weekly 기준 날짜가 월요일인지 확인했다.
[ ] 대상 시트가 정상적으로 선택되었다.
[ ] Query가 있는 예상 행 수와 API 호출 대상 수가 일치한다.
[ ] 배치별 체크포인트가 completed 결과 Excel에 저장되었다.
[ ] 기존 Buzz Volume 값이 최신 API 결과로 덮어써졌다.
[ ] 최종 결과가 output/Buzz\_Volume/completed에 생성되었다.
[ ] 최종 실패 행의 Buzz Volume에 API\_Failed가 기록되었다.
[ ] 실패가 있는 경우 artifacts/failed\_results와 response\_samples를 확인했다.
\`\`\`

**---**

**## 16. 운영 시 권장사항**

\* 원본 Raw Excel을 직접 수정하지 않습니다.
\* 1\~4단계는 날짜와 실행 차수가 일치하는 Output/Media 폴더를 사용합니다.
\* 실행 전에 기존 결과 파일을 백업합니다.
\* 파일명을 임의로 변경하지 않으며, Buzz Volume 입력 파일명은 기준 날짜와 정확히 일치시킵니다.
\* Prompt 또는 Schema를 변경하면 버전을 함께 기록합니다.
\* LLM 분석 재실행 전 중복 API 비용 여부를 확인합니다.
\* 실패한 게시물은 오류 사유를 확인한 뒤 별도로 재실행합니다.
\* API Key와 인증 파일은 코드 저장소에 업로드하지 않습니다.

**---**