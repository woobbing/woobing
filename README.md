# NetSuite to Google Sheets Automation

이 스크립트는 NetSuite에서 특정 리포트(또는 Saved Search)를 Excel 파일로 다운로드한 후, 지정된 Google Sheets에 자동으로 업로드하는 자동화 도구입니다.

## 주요 기능

- Playwright를 사용하여 NetSuite에 로그인하고 리포트를 다운로드합니다.
- 보안 질문 인증을 지원합니다.
- 다운로드한 Excel 파일(xlsx, xls, xml)을 파싱합니다.
- Google Sheets API를 사용하여 데이터를 특정 시트에 업로드합니다.
- GitHub Actions를 통해 주기적으로 실행할 수 있습니다.

## 설치 방법

1.  **저장소 복제**
    ```bash
    git clone <your-repository-url>
    cd ohoraAutomation
    ```

2.  **가상 환경 생성 및 활성화**
    ```bash
    python -m venv venv
    source venv/bin/activate  # macOS/Linux
    .\venv\Scripts\activate   # Windows
    ```

3.  **필요한 라이브러리 설치**
    ```bash
    pip install -r requirements.txt
    playwright install
    ```

## 사용 방법

1.  `.env` 파일을 생성하고 아래 환경 변수를 설정합니다.

    ```env
    NETSUITE_EMAIL="your-netsuite-email@example.com"
    NETSUITE_PASSWORD="your-netsuite-password"
    NETSUITE_ACCOUNT_ID="1234567"
    NETSUITE_REPORT_URL="https://..."
    NETSUITE_SECURITY_ANSWERS="answer1,answer2" # 보안 질문 답변 (쉼표로 구분)
    GOOGLE_SPREADSHEET_ID="your-google-spreadsheet-id"
    GOOGLE_CREDENTIALS_PATH="path/to/your/credentials.json" # 로컬 실행 시
    # GOOGLE_CREDENTIALS_JSON='{"type": "service_account", ...}' # GitHub Actions Secret 용
    GOOGLE_WORKSHEET_NAME="Sheet1" # 선택 사항
    SYNC_STATUS_CELL="A1" # 동기화 상태 표시 셀 (선택 사항, 기본값: A1)
    ```

    **동기화 상태 셀 설정 (`SYNC_STATUS_CELL`)**
    - 데이터 동기화 상태를 표시할 셀 위치를 지정합니다.
    - 예: `A1`, `Z1`, `AA1`, `AB1` 등
    - 데이터와 겹치지 않도록 빈 셀을 선택하세요.
    - 설정하지 않으면 기본값 `A1`이 사용됩니다.
    - **리포트별로 다른 셀을 사용하려면** `reports_config.json`에서 각 리포트의 `sync_status_cell`을 설정하세요.

2.  **리포트별 동기화 상태 셀 설정** (선택 사항)

    `reports_config.json` 파일을 수정하여 각 리포트마다 다른 셀에 동기화 상태를 표시할 수 있습니다.

    ```json
    {
      "reports": [
        {
          "name": "Item List Export",
          "netsuite_url": "https://...",
          "spreadsheet_id": "your-spreadsheet-id",
          "worksheet_name": "Item_DB",
          "enabled": true,
          "sync_status_cell": "A1"
        },
        {
          "name": "BOM Revision List Export",
          "netsuite_url": "https://...",
          "spreadsheet_id": "your-spreadsheet-id",
          "worksheet_name": "BOM Revision_DB",
          "enabled": true,
          "sync_status_cell": "A1"
        }
      ]
    }
    ```

    **우선순위**: 리포트 설정 (`sync_status_cell`) > 환경변수 (`SYNC_STATUS_CELL`) > 기본값 (`A1`)

3.  **스크립트 실행**
    ```bash
    python main.py
    ```