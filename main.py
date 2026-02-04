"""
NetSuite → Google Sheets 자동화 메인 스크립트
GitHub Actions 또는 로컬에서 실행
다중 리포트 지원
"""

import os
import sys
import json
import time
import logging
from dataclasses import dataclass
from datetime import datetime
from dotenv import load_dotenv

from netsuite_export import NetSuiteExporter
from upload_to_sheets import upload_excel_to_google_sheets, GoogleSheetsUploader
from report_config import ReportConfig, ReportConfigManager, load_reports_from_env
from slack_notifier import SlackNotifier

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)


def get_env_or_fail(key: str) -> str:
    """환경변수 가져오기 (없으면 에러)"""
    value = os.getenv(key)
    if not value:
        raise ValueError(f"환경변수 '{key}'가 설정되지 않았습니다.")
    return value


@dataclass
class ProcessResult:
    """리포트 처리 결과"""
    report_name: str
    download_success: bool
    upload_success: bool
    file_path: str = None
    error: str = None


def process_reports(
    reports: list[ReportConfig],
    netsuite_email: str,
    netsuite_password: str,
    netsuite_account_id: str,
    netsuite_base_url: str,
    security_answer: str,
    google_credentials_path: str,
    google_credentials_json: str,
    headless: bool = True,
    default_sync_status_cell: str = "A1"
) -> list[ProcessResult]:
    """
    여러 리포트를 처리 (한 번 로그인 후 순차 처리)

    Returns:
        list[ProcessResult]: 각 리포트별 처리 결과
    """
    results = []

    if not reports:
        logger.warning("처리할 리포트가 없습니다.")
        return results

    # Google 인증 정보 처리
    credentials_dict = None
    credentials_json_path = None

    if google_credentials_json:
        credentials_dict = json.loads(google_credentials_json)
        credentials_json_path = None  # 명시적으로 None으로 설정
    elif google_credentials_path:
        credentials_json_path = google_credentials_path
        credentials_dict = None  # 명시적으로 None으로 설정
    else:
        raise ValueError("Google 인증 정보가 없습니다.")

    # Google Sheets Uploader 초기화 (상태 기록용)
    sheets_uploader = GoogleSheetsUploader(
        credentials_json=credentials_json_path,
        credentials_dict=credentials_dict
    )

    # NetSuite Exporter 초기화 (한 번만 로그인)
    exporter = NetSuiteExporter(
        email=netsuite_email,
        password=netsuite_password,
        account_id=netsuite_account_id,
        base_url=netsuite_base_url,
        security_answer=security_answer
    )

    try:
        exporter.start_browser(headless=headless)

        if not exporter.login():
            raise Exception("NetSuite 로그인 실패")

        # 각 리포트 처리
        total = len(reports)
        for i, report in enumerate(reports, 1):
            logger.info("=" * 50)
            logger.info(f"[{i}/{total}] 리포트 처리 중: {report.name}")
            logger.info("=" * 50)

            result = ProcessResult(report_name=report.name, download_success=False, upload_success=False)

            # 리포트별 동기화 상태 셀 결정 (우선순위: 리포트 설정 > 환경변수 > 기본값)
            sync_cell = report.sync_status_cell or os.getenv("SYNC_STATUS_CELL") or default_sync_status_cell

            # 동기화 시작 상태 기록
            sheets_uploader.update_sync_status(
                spreadsheet_id=report.spreadsheet_id,
                worksheet_name=report.worksheet_name,
                status="동기화 진행 중",
                cell=sync_cell
            )

            # 1. 다운로드
            try:
                logger.info(f"다운로드 중: {report.netsuite_url[:60]}...")
                file_path = exporter.export_report(report.netsuite_url)

                if file_path:
                    result.download_success = True
                    result.file_path = file_path
                    logger.info(f"다운로드 완료: {file_path}")
                else:
                    result.error = "Export 실패"
                    logger.error("다운로드 실패: Export 실패")
                    results.append(result)
                    continue

            except Exception as e:
                result.error = str(e)
                logger.error(f"다운로드 실패: {e}")
                results.append(result)
                continue

            # 2. 업로드
            try:
                logger.info(f"업로드 중: Spreadsheet ID={report.spreadsheet_id}")
                success = upload_excel_to_google_sheets(
                    excel_path=file_path,
                    spreadsheet_id=report.spreadsheet_id,
                    credentials_json=credentials_json_path,
                    credentials_dict=credentials_dict,
                    worksheet_name=report.worksheet_name
                )

                if success:
                    result.upload_success = True
                    logger.info("업로드 완료!")

                    # 동기화 완료 상태 기록
                    sheets_uploader.update_sync_status(
                        spreadsheet_id=report.spreadsheet_id,
                        worksheet_name=report.worksheet_name,
                        status="동기화 완료",
                        cell=sync_cell
                    )
                else:
                    result.error = "업로드 실패"
                    logger.error("업로드 실패")

                    # 동기화 실패 상태 기록
                    sheets_uploader.update_sync_status(
                        spreadsheet_id=report.spreadsheet_id,
                        worksheet_name=report.worksheet_name,
                        status="동기화 실패",
                        cell=sync_cell
                    )

            except Exception as e:
                result.error = str(e)
                logger.error(f"업로드 실패: {e}")

                # 동기화 실패 상태 기록
                sheets_uploader.update_sync_status(
                    spreadsheet_id=report.spreadsheet_id,
                    worksheet_name=report.worksheet_name,
                    status="동기화 실패",
                    cell=sync_cell
                )

            results.append(result)

    finally:
        exporter.close()

    return results


def print_summary(results: list[ProcessResult]):
    """처리 결과 요약 출력"""
    logger.info("=" * 50)
    logger.info("처리 결과 요약")
    logger.info("=" * 50)

    success_count = sum(1 for r in results if r.download_success and r.upload_success)
    fail_count = len(results) - success_count

    for r in results:
        if r.download_success and r.upload_success:
            status = "[OK]"
        elif r.download_success:
            status = "[UPLOAD FAIL]"
        else:
            status = "[DOWNLOAD FAIL]"

        logger.info(f"  {status} {r.report_name}")
        if r.error:
            logger.info(f"       Error: {r.error}")

    logger.info(f"성공: {success_count}개 / 실패: {fail_count}개 / 총: {len(results)}개")

    return fail_count == 0


def main():
    """메인 실행 함수"""
    # .env 파일 로드 (로컬 실행 시)
    load_dotenv()

    # --headless 인자 확인 (GitHub Actions에서 사용)
    is_headless = "--headless" in sys.argv

    # Slack Notifier 초기화
    slack = SlackNotifier()

    # 시작 시간 기록
    start_time = time.time()

    logger.info("=" * 50)
    logger.info("NetSuite -> Google Sheets 자동화 시작")
    logger.info("=" * 50)

    # 환경변수 로드
    try:
        netsuite_email = get_env_or_fail("NETSUITE_EMAIL")
        netsuite_password = get_env_or_fail("NETSUITE_PASSWORD")
        netsuite_account_id = get_env_or_fail("NETSUITE_ACCOUNT_ID")

        # Google 인증 정보 (JSON 파일 경로 또는 직접 JSON 문자열)
        google_credentials_path = os.getenv("GOOGLE_CREDENTIALS_PATH")
        google_credentials_json = os.getenv("GOOGLE_CREDENTIALS_JSON")

        # 선택적 환경변수
        netsuite_base_url = os.getenv("NETSUITE_BASE_URL")
        # 보안 질문 답변 (쉼표로 구분된 여러 답변 지원)
        security_answer = os.getenv("NETSUITE_SECURITY_ANSWERS") or os.getenv("NETSUITE_SECURITY_ANSWER")

    except ValueError as e:
        logger.error(f"환경변수 오류: {e}")
        sys.exit(1)

    # 리포트 설정 로드 (우선순위: 설정파일 > 환경변수)
    config_manager = ReportConfigManager()
    reports = config_manager.get_enabled_reports()

    if not reports:
        # 설정 파일에 리포트가 없으면 환경변수에서 로드 (하위 호환)
        logger.info("설정 파일에 리포트가 없습니다. 환경변수에서 로드 시도...")
        reports = load_reports_from_env()

    if not reports:
        logger.error("처리할 리포트가 없습니다.")
        logger.error("  - reports_config.json에 리포트를 추가하거나")
        logger.error("  - NETSUITE_REPORT_URL, GOOGLE_SPREADSHEET_ID 환경변수를 설정하세요.")
        sys.exit(1)

    logger.info(f"처리할 리포트: {len(reports)}개")
    for r in reports:
        logger.info(f"  - {r.name}: {r.netsuite_url[:50]}...")

    # 리포트 처리
    results = []
    try:
        results = process_reports(
            reports=reports,
            netsuite_email=netsuite_email,
            netsuite_password=netsuite_password,
            netsuite_account_id=netsuite_account_id,
            netsuite_base_url=netsuite_base_url,
            security_answer=security_answer,
            google_credentials_path=google_credentials_path,
            google_credentials_json=google_credentials_json,
            headless=is_headless
        )

        # 결과 요약
        all_success = print_summary(results)

        # 소요 시간 계산
        duration = time.time() - start_time

        # Slack 알림
        if all_success:
            slack.send_success_notification(results, duration)
            logger.info("모든 자동화 완료!")
            sys.exit(0)
        else:
            slack.send_failure_notification(results, duration=duration)
            logger.warning("일부 리포트 처리 실패")
            sys.exit(1)

    except Exception as e:
        # 치명적 오류 발생 시
        duration = time.time() - start_time
        error_msg = str(e)
        logger.error(f"치명적 오류 발생: {error_msg}")

        slack.send_failure_notification(results, error=error_msg, duration=duration)
        sys.exit(1)


if __name__ == "__main__":
    main()
