"""
NetSuite Report Export 자동화 모듈
Playwright를 사용하여 NetSuite에 로그인하고 Report를 Excel로 다운로드
"""

import os
import time
import logging
from pathlib import Path
from playwright.sync_api import sync_playwright, Page, Browser

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)


class NetSuiteExporter:
    # 타임아웃 기본값 (환경변수로 오버라이드 가능)
    DEFAULT_PAGE_TIMEOUT = int(os.getenv("NS_PAGE_TIMEOUT", "60000"))  # 60초
    DEFAULT_ELEMENT_TIMEOUT = int(os.getenv("NS_ELEMENT_TIMEOUT", "5000"))  # 5초
    DEFAULT_DOWNLOAD_TIMEOUT = int(os.getenv("NS_DOWNLOAD_TIMEOUT", "120000"))  # 120초
    DEFAULT_NETWORK_IDLE_TIMEOUT = int(os.getenv("NS_NETWORK_IDLE_TIMEOUT", "60000"))  # 60초

    def __init__(
        self,
        email: str,
        password: str,
        account_id: str,
        base_url: str = None,
        download_dir: str = None,
        security_answer: str = None
    ):
        """
        NetSuite Exporter 초기화

        Args:
            email: NetSuite 로그인 이메일
            password: NetSuite 비밀번호
            account_id: NetSuite Account ID (예: 1234567)
            base_url: 회사 전용 URL (예: https://1234567.app.netsuite.com)
            download_dir: 다운로드 디렉토리 경로
            security_answer: 보안 질문 답변
        """
        self.email = email
        self.password = password
        self.account_id = account_id
        self.base_url = base_url or f"https://{account_id}.app.netsuite.com"
        self.download_dir = download_dir or os.path.join(os.getcwd(), "downloads")
        self.security_answer = security_answer

        # 다운로드 디렉토리 생성
        Path(self.download_dir).mkdir(parents=True, exist_ok=True)

        self.browser: Browser = None
        self.page: Page = None
        self.playwright = None
        self.logger = logger

    def start_browser(self, headless: bool = True):
        """브라우저 시작"""
        try:
            self.logger.info("Playwright 시작 중...")
            self.playwright = sync_playwright().start()
            self.logger.info("Chromium 브라우저 실행 중...")
            self.browser = self.playwright.chromium.launch(
                headless=headless,
                args=['--disable-blink-features=AutomationControlled']
            )
            self.logger.info("브라우저 컨텍스트 생성 중...")
            self.context = self.browser.new_context(
                accept_downloads=True,
                viewport={'width': 1920, 'height': 1080}
            )
            self.page = self.context.new_page()
            # 자동화 탐지 우회
            self.page.add_init_script("""
                Object.defineProperty(navigator, 'webdriver', {
                    get: () => undefined
                });
            """)
        except Exception as e:
            self.logger.critical(f"브라우저 시작 실패: {e}")
            raise

    def _save_debug_artifacts(self, prefix: str) -> None:
        """디버깅용 스크린샷과 HTML 저장"""
        if not self.page:
            return
        try:
            screenshot_path = os.path.join(self.download_dir, f"{prefix}.png")
            html_path = os.path.join(self.download_dir, f"{prefix}.html")
            self.page.screenshot(path=screenshot_path)
            with open(html_path, "w", encoding="utf-8") as f:
                f.write(self.page.content())
            self.logger.debug(f"디버그 아티팩트 저장: {prefix}")
        except Exception as e:
            self.logger.warning(f"디버그 아티팩트 저장 실패: {e}")

    def _click_first_visible(self, selectors: list[str], description: str, timeout: int = None) -> bool:
        """보이는 첫 번째 요소를 찾아 클릭합니다."""
        timeout = timeout or self.DEFAULT_ELEMENT_TIMEOUT
        for selector in selectors:
            try:
                locator = self.page.locator(selector).first
                if locator.is_visible(timeout=1000):
                    locator.click(timeout=timeout)
                    self.logger.info(f"{description} 클릭 완료: {selector}")
                    return True
            except Exception:
                continue
        self.logger.warning(f"{description} 버튼을 찾을 수 없습니다")
        self._save_debug_artifacts(f"{description.replace(' ', '_').lower()}_click_error")
        return False

    def _fill_first_visible(self, selectors: list[str], value: str, description: str) -> bool:
        """보이는 첫 번째 요소를 찾아 값을 채웁니다."""
        for selector in selectors:
            try:
                locator = self.page.locator(selector).first
                if locator.is_visible(timeout=1000):
                    locator.fill(value)
                    self.logger.info(f"{description} 입력 완료: {selector}")
                    return True
            except Exception:
                continue
        self.logger.warning(f"{description} 필드를 찾거나 채울 수 없습니다")
        self._save_debug_artifacts(f"{description.replace(' ', '_').lower()}_fill_error")
        return False


    def login(self) -> bool:
        """
        NetSuite 로그인

        Returns:
            bool: 로그인 성공 여부
        """
        try:
            self.logger.info("NetSuite 로그인 시도...")

            # system.netsuite.com을 통한 로그인 (표준 방식)
            login_url = "https://system.netsuite.com/pages/customerlogin.jsp"
            self.logger.info(f"로그인 페이지로 이동: {login_url}")
            self.page.goto(login_url, wait_until="networkidle", timeout=self.DEFAULT_NETWORK_IDLE_TIMEOUT)

            # 로그인 페이지 스크린샷
            self._save_debug_artifacts("login_page")

            # 이메일 입력 필드 대기 및 입력
            email_selectors = [
                'input[name="email"]',
                'input#email',
                'input[type="email"]',
                'input[placeholder*="email" i]',
                'input[placeholder*="Email" i]',
            ]
            if not self._fill_first_visible(email_selectors, self.email, "이메일"):
                return False

            # 비밀번호 입력
            password_selectors = [
                'input[name="password"]',
                'input#password',
                'input[type="password"]',
            ]
            if not self._fill_first_visible(password_selectors, self.password, "비밀번호"):
                return False

            # 로그인 버튼 클릭
            login_button_selectors = [
                'input[type="submit"]',
                'button[type="submit"]',
                'button:has-text("Log In")',
                '#login-submit',
            ]

            if not self._click_first_visible(login_button_selectors, "로그인 버튼"):
                return False

            # 로그인 완료 대기
            self.logger.info("로그인 처리 대기 중...")
            self.page.wait_for_load_state("networkidle", timeout=self.DEFAULT_NETWORK_IDLE_TIMEOUT)
            time.sleep(3)

            # 로그인 후 스크린샷
            self._save_debug_artifacts("after_login")

            # 로그인 에러 확인
            error_selectors = ['.error', '.alert', '[class*="error"]', '[role="alert"]']
            for selector in error_selectors:
                if self.page.locator(selector).count() > 0:
                    error_text = self.page.locator(selector).first.text_content()
                    if error_text and error_text.strip():
                        self.logger.error(f"로그인 에러 감지: {error_text}")
                        return False

            # 보안 질문 페이지 확인 및 처리
            current_url = self.page.url
            self.logger.info(f"현재 URL: {current_url}")

            if "securityquestions" in current_url.lower():
                self.logger.info("보안 질문 페이지 감지됨")
                if not self.security_answer:
                    self.logger.error("보안 질문 답변이 설정되지 않았습니다. NETSUITE_SECURITY_ANSWERS 환경변수를 설정하세요.")
                    return False

                # security_answer가 쉼표로 구분된 여러 답변인 경우 리스트로 변환
                answers = [a.strip() for a in self.security_answer.split(',')] if ',' in self.security_answer else [self.security_answer]
                self.logger.info(f"시도할 답변 개수: {len(answers)}")

                # 각 답변을 순차적으로 시도
                for attempt, answer in enumerate(answers):
                    self.logger.info(f"답변 시도 {attempt + 1}/{len(answers)}")

                    # 현재 URL이 보안 질문 페이지가 아니면 이미 성공한 것
                    if "securityquestions" not in self.page.url.lower():
                        self.logger.info("보안 질문 통과!")
                        break

                    # 보안 질문 답변 입력
                    answer_selectors = [
                        'input[name="answer"]',
                        'input#answer',
                        'input[type="text"]',
                    ]
                    if not self._fill_first_visible(answer_selectors, answer, "보안 질문 답변"):
                        continue

                    # 제출 버튼 클릭
                    submit_selectors = [
                        'input[type="submit"]',
                        'button[type="submit"]',
                        'input[value*="제출"]',
                        'button:has-text("제출")',
                        'input[value*="Submit"]',
                    ]

                    if not self._click_first_visible(submit_selectors, "보안 질문 제출"):
                        continue

                    # 보안 질문 처리 후 대기
                    self.logger.info("보안 질문 처리 대기 중...")
                    self.page.wait_for_load_state("networkidle", timeout=self.DEFAULT_NETWORK_IDLE_TIMEOUT)
                    time.sleep(2)

                    # 성공 여부 확인
                    current_url = self.page.url
                    if "securityquestions" not in current_url.lower():
                        self.logger.info("보안 질문 답변 정답!")
                        break
                    else:
                        self.logger.warning("보안 질문 답변이 틀렸습니다. 다음 답변 시도...")

                # 보안 질문 후 스크린샷
                self._save_debug_artifacts("after_security")
                current_url = self.page.url
                self.logger.info(f"보안 질문 후 URL: {current_url}")

            # 로그인 성공 여부 확인
            if "customerlogin" not in current_url.lower() and "securityquestions" not in current_url.lower():
                self.logger.info("로그인 성공!")
                self._establish_account_session()
                return True
            elif "app.netsuite.com" in current_url and "login" not in current_url.lower():
                self.logger.info("로그인 성공! (앱 페이지로 이동됨)")
                self._establish_account_session()
                return True
            else:
                error_text = self.page.locator('.error, .alert, [class*="error"]').text_content() if self.page.locator('.error, .alert, [class*="error"]').count() > 0 else ""
                self.logger.error(f"로그인 실패 - 현재 URL: {current_url}, 에러: {error_text}")
                self._save_debug_artifacts("login_final_failure")
                return False

        except Exception as e:
            self.logger.error(f"로그인 중 오류 발생: {str(e)}")
            self._save_debug_artifacts("login_error")
            return False

    def _establish_account_session(self) -> None:
        """Account 홈페이지로 이동하여 세션을 확립합니다."""
        self.logger.info(f"Account 홈페이지로 이동 중: {self.base_url}")
        try:
            self.page.goto(self.base_url, wait_until="networkidle", timeout=self.DEFAULT_PAGE_TIMEOUT)
            time.sleep(2)
            self.logger.info("Account 세션 확립 완료")
        except Exception as e:
            self.logger.warning(f"Account 홈페이지 이동 실패 (계속 진행): {e}")

    def export_saved_search_results(self, search_url: str) -> str:
        """
        Saved Search 결과를 Excel로 Export

        Args:
            search_url: Saved Search 결과 페이지 URL

        Returns:
            str: 다운로드된 파일 경로
        """
        try:
            self.logger.info(f"Saved Search 페이지로 이동: {search_url}")
            self.page.goto(search_url, wait_until="networkidle", timeout=self.DEFAULT_DOWNLOAD_TIMEOUT)

            # 검색 결과 로딩 대기
            self.logger.info("검색 결과 로딩 대기 중...")
            time.sleep(5)

            # 스크린샷 저장 (디버깅용)
            self._save_debug_artifacts("search_page")
            self.logger.debug("검색 페이지 스크린샷 저장됨")

            # 방법 1: Export 아이콘/링크 직접 클릭
            export_selectors = [
                '[id*="csv"]', '[id*="CSV"]',
                'img[alt*="CSV"]', 'img[alt*="Excel"]',
                '[id*="excel"]',
                'a:has-text("Export")', 'span:has-text("Export")',
                'div[id*="export"]', 'a[id*="export"]',
                'input[value*="Export"]',
                'a:has-text("CSV")', 'a:has-text("Excel")',
            ]

            # Export 버튼 찾기
            for selector in export_selectors:
                try:
                    locator = self.page.locator(selector)
                    if locator.count() > 0:
                        self.logger.info(f"Export 요소 발견: {selector}")
                        with self.page.expect_download(timeout=self.DEFAULT_DOWNLOAD_TIMEOUT) as download_info:
                            locator.first.click()
                        download = download_info.value
                        file_path = os.path.join(self.download_dir, download.suggested_filename)
                        download.save_as(file_path)
                        self.logger.info(f"파일 다운로드 완료: {file_path}")
                        return file_path
                except Exception as e:
                    self.logger.debug(f"선택자 {selector} 시도 실패: {e}")
                    continue

            # 방법 2: JavaScript로 Excel Export 트리거
            self.logger.info("JavaScript로 Excel Export 시도...")
            try:
                with self.page.expect_download(timeout=self.DEFAULT_DOWNLOAD_TIMEOUT) as download_info:
                    self.page.evaluate("""
                        var currentUrl = window.location.href;
                        var xlsUrl = currentUrl.replace('searchresults.nl', 'searchresults.xls');
                        if (xlsUrl.indexOf('csv=') === -1) {
                            xlsUrl += '&csv=Export&OfficeXML=T&size=1000';
                        }
                        window.location.href = xlsUrl;
                    """)
                download = download_info.value
                file_path = os.path.join(self.download_dir, download.suggested_filename)
                download.save_as(file_path)
                self.logger.info(f"Excel 다운로드 완료: {file_path}")
                return file_path
            except Exception as e:
                self.logger.warning(f"JavaScript Export 실패: {e}")
                if "Download is starting" in str(e):
                    self.logger.info("다운로드가 시작됨, 대기 중...")
                    time.sleep(15)

            # 방법 3: 다운로드 폴더에서 파일 확인
            file_path = self._find_latest_downloaded_file()
            if file_path:
                return file_path

            self.logger.error("Export를 찾을 수 없습니다")
            self._save_debug_artifacts("export_not_found")
            return None

        except Exception as e:
            self.logger.error(f"Export 중 오류 발생: {str(e)}")
            self._save_debug_artifacts("export_error")
            return None

    def _find_latest_downloaded_file(self) -> str:
        """다운로드 폴더에서 가장 최근 다운로드된 파일 찾기"""
        import glob
        time.sleep(5)
        csv_files = glob.glob(os.path.join(self.download_dir, "*.csv"))
        xlsx_files = glob.glob(os.path.join(self.download_dir, "*.xlsx"))
        xls_files = glob.glob(os.path.join(self.download_dir, "*.xls"))
        all_files = csv_files + xlsx_files + xls_files

        if all_files:
            latest_file = max(all_files, key=os.path.getctime)
            self.logger.info(f"다운로드 폴더에서 파일 발견: {latest_file}")
            return latest_file
        return None

    def export_report(self, report_url: str, wait_time: int = 5) -> str:
        """
        Report 또는 Saved Search를 Excel/CSV로 Export

        Args:
            report_url: Report 또는 Saved Search 페이지 URL
            wait_time: 다운로드 대기 시간 (초)

        Returns:
            str: 다운로드된 파일 경로
        """
        # Saved Search URL인 경우 전용 메서드 사용
        if "searchresults.nl" in report_url or "searchid=" in report_url:
            return self.export_saved_search_results(report_url)

        # 일반 Report Export
        try:
            self.logger.info(f"Report 페이지로 이동: {report_url}")
            self.page.goto(report_url, wait_until="networkidle", timeout=self.DEFAULT_PAGE_TIMEOUT)

            # Report 로딩 대기
            time.sleep(3)

            # Export 버튼 찾기 및 클릭
            export_selectors = [
                'text="Export"',
                'button:has-text("Export")',
                '[id*="export"]',
                '[class*="export"]',
                'text="Excel"',
                'a:has-text("Excel")',
            ]

            export_clicked = False
            for selector in export_selectors:
                try:
                    if self.page.locator(selector).count() > 0:
                        self.page.click(selector, timeout=self.DEFAULT_ELEMENT_TIMEOUT)
                        export_clicked = True
                        self.logger.info(f"Export 버튼 클릭: {selector}")
                        break
                except:
                    continue

            if not export_clicked:
                try:
                    self.page.click('[id*="menu"], [class*="dropdown"]', timeout=self.DEFAULT_ELEMENT_TIMEOUT)
                    time.sleep(1)
                    self.page.click('text="Excel"', timeout=self.DEFAULT_ELEMENT_TIMEOUT)
                    export_clicked = True
                except:
                    pass

            if not export_clicked:
                self.logger.error("Export 버튼을 찾을 수 없습니다")
                self._save_debug_artifacts("export_not_found")
                return None

            # 다운로드 대기 및 처리
            with self.page.expect_download(timeout=self.DEFAULT_DOWNLOAD_TIMEOUT) as download_info:
                try:
                    self.page.click('text="Excel"', timeout=self.DEFAULT_ELEMENT_TIMEOUT)
                except:
                    pass

            download = download_info.value
            file_path = os.path.join(self.download_dir, download.suggested_filename)
            download.save_as(file_path)

            self.logger.info(f"파일 다운로드 완료: {file_path}")
            return file_path

        except Exception as e:
            self.logger.error(f"Export 중 오류 발생: {str(e)}")
            self._save_debug_artifacts("export_error")
            return None

    def export_saved_search(self, search_id: str) -> str:
        """
        Saved Search 결과를 Excel로 Export

        Args:
            search_id: Saved Search ID

        Returns:
            str: 다운로드된 파일 경로
        """
        search_url = f"{self.base_url}/app/common/search/searchresults.nl?searchid={search_id}&whence="
        return self.export_report(search_url)

    def close(self):
        """브라우저 종료"""
        if self.browser:
            self.browser.close()
        if self.playwright:
            self.playwright.stop()
        self.logger.info("브라우저 종료")


def download_netsuite_report(
    email: str,
    password: str,
    account_id: str,
    report_url: str,
    base_url: str = None,
    download_dir: str = None,
    headless: bool = True,
    security_answer: str = None
) -> str:
    """
    NetSuite Report 다운로드 헬퍼 함수 (단일 리포트)

    Returns:
        str: 다운로드된 파일 경로
    """
    exporter = NetSuiteExporter(
        email=email,
        password=password,
        account_id=account_id,
        base_url=base_url,
        download_dir=download_dir,
        security_answer=security_answer
    )

    try:
        exporter.start_browser(headless=headless)

        if not exporter.login():
            raise Exception("NetSuite 로그인 실패")

        file_path = exporter.export_report(report_url)

        if not file_path:
            raise Exception("Report Export 실패")

        return file_path

    finally:
        exporter.close()


def download_netsuite_reports(
    email: str,
    password: str,
    account_id: str,
    report_urls: list[str],
    base_url: str = None,
    download_dir: str = None,
    headless: bool = True,
    security_answer: str = None
) -> dict[str, str]:
    """
    NetSuite 다중 Report 다운로드 헬퍼 함수
    한 번 로그인 후 여러 리포트를 순차적으로 다운로드

    Args:
        report_urls: 다운로드할 리포트 URL 목록

    Returns:
        dict: {report_url: file_path} 매핑 (실패 시 file_path는 None)
    """
    exporter = NetSuiteExporter(
        email=email,
        password=password,
        account_id=account_id,
        base_url=base_url,
        download_dir=download_dir,
        security_answer=security_answer
    )

    results = {}

    try:
        exporter.start_browser(headless=headless)

        if not exporter.login():
            raise Exception("NetSuite 로그인 실패")

        for i, report_url in enumerate(report_urls, 1):
            logger.info(f"[{i}/{len(report_urls)}] 리포트 다운로드 중...")
            try:
                file_path = exporter.export_report(report_url)
                results[report_url] = file_path
                if file_path:
                    logger.info(f"  성공: {file_path}")
                else:
                    logger.warning(f"  실패: Export 실패")
            except Exception as e:
                logger.error(f"  실패: {e}")
                results[report_url] = None

        return results

    finally:
        exporter.close()


if __name__ == "__main__":
    # 테스트용 (환경변수에서 값 로드)
    from dotenv import load_dotenv
    load_dotenv()

    email = os.getenv("NETSUITE_EMAIL")
    password = os.getenv("NETSUITE_PASSWORD")
    account_id = os.getenv("NETSUITE_ACCOUNT_ID")
    report_url = os.getenv("NETSUITE_REPORT_URL")
    # 쉼표로 구분된 여러 답변 지원 (NETSUITE_SECURITY_ANSWERS) 또는 단일 답변 (NETSUITE_SECURITY_ANSWER)
    security_answer = os.getenv("NETSUITE_SECURITY_ANSWERS") or os.getenv("NETSUITE_SECURITY_ANSWER")

    if all([email, password, account_id, report_url]):
        file_path = download_netsuite_report(
            email=email,
            password=password,
            account_id=account_id,
            report_url=report_url,
            security_answer=security_answer,
            headless=False  # 테스트 시에는 브라우저 표시
        )
        print(f"다운로드 완료: {file_path}")
    else:
        print("환경변수가 설정되지 않았습니다. .env 파일을 확인하세요.")
