# Confidential Sheet Cover

[English](README.md) | [한국어](README.ko.md)

`confidential-sheet-cover`는 로컬 `.xlsx` 스프레드시트에서 급여, 연봉, 인건비, 사업비, 프로젝트 비용, 비용 배분 등 민감정보 관련 키워드를 검색하고, 매칭된 파일의 맨 첫 워크시트로 `Confidential` 커버 시트를 추가하는 CLI 도구입니다.

Google Drive for Desktop, OneDrive, Dropbox로 동기화된 폴더나 일반 로컬 폴더처럼 파일시스템에서 접근 가능한 폴더를 대상으로 사용할 수 있습니다.

## 주요 기능

- 하나 이상의 폴더 아래에 있는 `.xlsx` 파일 검색
- 파일 경로, 시트명, 셀 텍스트 검사
- 영어와 한국어 민감정보 키워드 기본 제공
- 첫 번째 워크시트로 `Confidential` 시트 추가
- CSV 매니페스트와 JSON 요약 파일 생성
- 파일 수정 전 백업 생성
- 실제 수정 전 `--dry-run`으로 대상 파일 사전 확인

## 기본 키워드

기본 키워드 목록에는 다음과 같은 영어/한국어 용어가 포함되어 있습니다.

```text
payroll, salary, salaries, wage, compensation, bonus, severance,
personnel cost, labor cost, headcount cost, project cost,
business cost, business expense, operating expense, grant budget,
인건비, 연봉, 임금, 급여, 보상, 상여금, 퇴직금, 포괄임금, 사업비, 운영비, 예산
```

`--keyword`로 키워드를 추가하거나, `--no-default-keywords --keywords-file` 조합으로 기본 키워드를 사용하지 않고 별도 키워드 파일만 사용할 수 있습니다.

## 설치

```bash
python3 -m pip install .
```

저장소에서 바로 실행할 수도 있습니다.

```bash
python3 -m confidential_sheet_cover --help
```

## Google Drive 준비

이 도구는 Google Drive API를 직접 호출하지 않습니다. 로컬 `.xlsx` 파일을 처리하는 도구입니다. Google Drive에 있는 파일을 대상으로 사용하려면 먼저 대상 워크북을 로컬 파일시스템에서 접근할 수 있게 만들어야 합니다.

### 선택지 1: Google Drive for Desktop

가장 간단한 방식입니다.

1. Google Drive for Desktop을 설치합니다.
2. 대상 Drive 파일에 접근 권한이 있는 Google 계정으로 로그인합니다.
3. 컴퓨터에서 동기화된 Drive 폴더를 엽니다.
4. Drive 클라이언트가 스트리밍 모드라면 대상 `.xlsx` 파일이 오프라인 사용 가능 상태인지 확인합니다.
5. 동기화된 로컬 폴더 경로를 대상으로 이 도구를 실행합니다.

예시:

```bash
confidential-sheet-cover "$HOME/Google Drive/My Drive/Finance" --dry-run
confidential-sheet-cover "$HOME/Google Drive/My Drive/Finance"
```

정확한 로컬 Drive 경로는 운영체제와 Google Drive for Desktop 설정에 따라 달라집니다.

### 선택지 2: rclone + Google OAuth

Drive 폴더를 CLI로 반복적으로 복사하거나 마운트해야 한다면 이 방식을 사용할 수 있습니다.

상위 절차:

1. Google Cloud 프로젝트를 생성합니다.
2. 해당 프로젝트에서 Google Drive API를 활성화합니다.
3. OAuth 동의 화면을 설정합니다.
4. OAuth Client ID를 생성합니다.
5. 로컬 CLI 인증 용도라면 OAuth 클라이언트 유형은 보통 **Desktop app**을 선택합니다.
6. `rclone`을 설치하고 설정합니다.
7. `rclone config`에서 OAuth client ID와 secret을 사용해 Google Drive remote를 만듭니다.
8. 대상 Drive 폴더를 로컬로 복사, 동기화, 또는 마운트합니다.
9. 로컬로 복사되거나 마운트된 폴더를 대상으로 이 도구를 실행합니다.

예시 rclone 흐름:

```bash
rclone config
rclone copy mydrive:"Folder/With/Workbooks" ./drive-workbooks
confidential-sheet-cover ./drive-workbooks --dry-run
confidential-sheet-cover ./drive-workbooks
```

로컬 OAuth redirect 흐름을 사용한다면 OAuth 클라이언트 유형은 보통 **Web application**이 아니라 **Desktop app**이어야 합니다. Web application 클라이언트를 사용하면 정확한 로컬 redirect URI가 등록되어 있지 않은 경우 redirect URI mismatch 오류가 날 수 있습니다.

### Native Google Sheets

Google Sheets 네이티브 파일은 `.xlsx` 파일이 아닙니다. 이 도구는 네이티브 Google Sheets를 직접 수정하지 않습니다.

권장 선택지:

- 네이티브 Google Sheets를 `.xlsx`로 내보낸 뒤 이 도구를 실행하고, 필요하면 수정된 워크북을 다시 업로드합니다.
- Google Drive에 저장된 Office `.xlsx` 파일은 Google Drive for Desktop이나 rclone을 통해 로컬 파일로 접근해 처리합니다.
- 네이티브 Google Sheets를 원본 그대로 수정해야 한다면 별도의 Google Sheets API 워크플로를 만들어야 합니다.

### Service Account

이 도구를 사용하는 데 Google service account는 필요하지 않습니다.

Service account는 domain-wide delegation 등 Google Workspace 관리자 설정을 갖춘 조직 단위 자동화를 만들 때만 고려하면 됩니다. 개인 또는 팀 단위 로컬 작업에는 Google Drive for Desktop이나 OAuth 기반 rclone remote가 보통 더 단순합니다.

## 사용법

항상 dry run부터 실행하는 것을 권장합니다.

```bash
python3 -m confidential_sheet_cover "/path/to/folder" --dry-run
```

`confidential-cover-results/manifest.csv`를 검토한 뒤 실제 변경을 적용합니다.

```bash
python3 -m confidential_sheet_cover "/path/to/folder"
```

`pip install .`로 설치했다면 콘솔 명령을 사용할 수 있습니다.

```bash
confidential-sheet-cover "/path/to/folder" --dry-run
confidential-sheet-cover "/path/to/folder"
```

## 주요 옵션

```bash
confidential-sheet-cover "/path/to/folder" \
  --dry-run \
  --output-dir confidential-cover-results \
  --max-cells-per-sheet 5000
```

사용자 지정 키워드 추가:

```bash
confidential-sheet-cover "/path/to/folder" \
  --keyword "equity compensation" \
  --keyword "contractor fee"
```

키워드 파일 사용:

```bash
confidential-sheet-cover "/path/to/folder" \
  --keywords-file keywords.txt
```

파일 경로와 시트명만 검사하고 셀 내용은 검사하지 않기:

```bash
confidential-sheet-cover "/path/to/folder" --no-content-scan
```

커버 시트명과 문구 변경:

```bash
confidential-sheet-cover "/path/to/folder" \
  --sheet-name "Confidential" \
  --cover-text "Confidential Document - details start on the next sheet"
```

## 산출물

기본적으로 결과는 `confidential-cover-results/`에 저장됩니다.

```text
confidential-cover-results/
  manifest.csv
  summary.json
  backups/
```

`manifest.csv`에는 다음 정보가 포함됩니다.

- 파일 경로
- 처리 상태
- 매칭된 키워드
- 매칭 위치
- 처리 전/후 첫 시트명
- 백업 파일 경로
- 처리 전/후 SHA-256 해시
- 오류 메시지

주요 상태값:

| 상태 | 의미 |
| --- | --- |
| `dry_run_target` | 키워드가 매칭되었지만 `--dry-run`이라 실제 수정하지 않음 |
| `updated` | 새 `Confidential` 커버 시트가 추가됨 |
| `moved_existing_cover` | 기존 커버 시트를 맨 앞으로 이동함 |
| `skipped_already_covered` | 이미 기대한 커버 시트가 첫 시트에 있음 |
| `skipped_no_match` | 민감 키워드가 발견되지 않음 |
| `error` | 파일을 검사하거나 수정하지 못함 |

## 안전 주의사항

- 항상 `--dry-run`을 먼저 실행합니다.
- 실제 적용 전 `manifest.csv`를 검토합니다.
- 수정되는 각 워크북은 저장 전 백업이 생성됩니다.
- 이 도구는 `.xlsx` 파일을 지원합니다. 네이티브 Google Sheets는 처리 전에 `.xlsx`로 내보내야 합니다.
- 오래된 `.xls` 파일은 처리 전에 `.xlsx`로 변환하는 것을 권장합니다.
- `openpyxl`은 Excel의 모든 고급 기능을 완벽하게 보존하지 않을 수 있습니다. 중요한 워크북에는 먼저 사본으로 테스트하세요.

## 라이선스

MIT
