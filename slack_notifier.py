"""
Slack 알림 모듈
작업 성공/실패 시 Slack으로 메시지 전송
"""

import os
import requests
from datetime import datetime, timezone, timedelta
from typing import Optional


class SlackNotifier:
    # 한국 시간대 (UTC+9)
    KST = timezone(timedelta(hours=9))

    def __init__(self, webhook_url: str = None):
        """
        Slack Notifier 초기화

        Args:
            webhook_url: Slack Webhook URL (환경변수 SLACK_WEBHOOK_URL에서 자동 로드)
        """
        self.webhook_url = webhook_url or os.getenv("SLACK_WEBHOOK_URL")
        if not self.webhook_url:
            print("[WARNING] SLACK_WEBHOOK_URL이 설정되지 않았습니다. Slack 알림이 비활성화됩니다.")

    def send_message(self, text: str, blocks: list = None) -> bool:
        """
        Slack 메시지 전송

        Args:
            text: 메시지 텍스트 (fallback용)
            blocks: Slack Block Kit 블록 (선택사항)

        Returns:
            bool: 전송 성공 여부
        """
        if not self.webhook_url:
            print("[INFO] Slack Webhook URL이 없어 메시지를 전송하지 않습니다.")
            return False

        payload = {"text": text}
        if blocks:
            payload["blocks"] = blocks

        try:
            response = requests.post(self.webhook_url, json=payload, timeout=10)
            if response.status_code == 200:
                print("Slack 메시지 전송 성공")
                return True
            else:
                print(f"Slack 메시지 전송 실패: {response.status_code} - {response.text}")
                return False
        except Exception as e:
            print(f"Slack 메시지 전송 중 오류: {e}")
            return False

    def send_success_notification(self, results: list, duration: float) -> bool:
        """
        작업 성공 알림 (간단하게 한 줄)

        Args:
            results: ProcessResult 리스트
            duration: 작업 소요 시간 (초)
        """
        success_count = sum(1 for r in results if r.download_success and r.upload_success)
        total_count = len(results)

        # 완료 시간 (한국 시간)
        complete_time = datetime.now(self.KST).strftime('%Y-%m-%d %H:%M:%S KST')

        # 간단하게 한 줄로 메시지 작성
        text = f"✅ 동기화 완료 ({success_count}/{total_count}) - {complete_time}"

        return self.send_message(text)

    def send_failure_notification(self, results: list, error: str = None, duration: float = None) -> bool:
        """
        작업 실패 알림 (간단하게 한 줄)

        Args:
            results: ProcessResult 리스트 (비어있을 수 있음)
            error: 에러 메시지
            duration: 작업 소요 시간 (초)
        """
        # 실패 시간 (한국 시간)
        fail_time = datetime.now(self.KST).strftime('%Y-%m-%d %H:%M:%S KST')

        if results:
            fail_count = sum(1 for r in results if not (r.download_success and r.upload_success))
            total_count = len(results)

            # 실패한 리포트 이름 수집
            failed_names = [r.report_name for r in results if not (r.download_success and r.upload_success)]

            if failed_names:
                text = f"❌ 동기화 실패 ({', '.join(failed_names)}) - {fail_time}"
            else:
                text = f"❌ 동기화 실패 ({fail_count}/{total_count}) - {fail_time}"
        else:
            text = f"❌ 동기화 실패 - {fail_time}"

        # 에러 메시지가 있으면 추가
        if error:
            text += f" | {error[:100]}"  # 에러 메시지는 최대 100자로 제한

        return self.send_message(text)


if __name__ == "__main__":
    # 테스트
    from dataclasses import dataclass

    @dataclass
    class TestResult:
        report_name: str
        download_success: bool
        upload_success: bool
        error: str = None

    notifier = SlackNotifier()

    # 성공 알림 테스트
    print("\n=== 성공 알림 테스트 ===")
    results = [
        TestResult("Item List Export", True, True),
        TestResult("BOM Revision List Export", True, True),
    ]
    notifier.send_success_notification(results, 45.3)

    # 실패 알림 테스트
    print("\n=== 실패 알림 테스트 ===")
    results_fail = [
        TestResult("Item List Export", True, True),
        TestResult("BOM Revision List Export", False, False, "Export 실패"),
    ]
    notifier.send_failure_notification(results_fail, duration=30.5)
