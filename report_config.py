"""
리포트 설정 관리 모듈
여러 NetSuite 리포트와 Google Sheets 매핑을 관리
"""

import os
import json
from dataclasses import dataclass, field, asdict
from typing import Optional
from pathlib import Path


@dataclass
class ReportConfig:
    """개별 리포트 설정"""
    name: str  # 리포트 식별 이름
    netsuite_url: str  # NetSuite 리포트/Saved Search URL
    spreadsheet_id: str  # Google Spreadsheet ID
    worksheet_name: Optional[str] = None  # Google Sheets 워크시트 이름
    enabled: bool = True  # 활성화 여부
    sync_status_cell: Optional[str] = None  # 동기화 상태 표시 셀 (예: "A1", "Z1" 등)

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict) -> "ReportConfig":
        return cls(**data)


class ReportConfigManager:
    """리포트 설정 관리자"""

    def __init__(self, config_path: str = None):
        """
        Args:
            config_path: 설정 파일 경로 (기본: reports_config.json)
        """
        self.config_path = config_path or os.path.join(
            os.path.dirname(__file__), "reports_config.json"
        )
        self.reports: list[ReportConfig] = []
        self._load_config()

    def _load_config(self):
        """설정 파일에서 리포트 목록 로드"""
        if os.path.exists(self.config_path):
            with open(self.config_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                self.reports = [
                    ReportConfig.from_dict(r) for r in data.get("reports", [])
                ]
            print(f"설정 로드 완료: {len(self.reports)}개 리포트")
        else:
            self.reports = []
            print(f"설정 파일이 없습니다. 새로 생성됩니다: {self.config_path}")

    def _save_config(self):
        """설정 파일에 리포트 목록 저장"""
        data = {"reports": [r.to_dict() for r in self.reports]}
        with open(self.config_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"설정 저장 완료: {self.config_path}")

    def add_report(self, report: ReportConfig) -> bool:
        """
        리포트 추가

        Args:
            report: 추가할 리포트 설정

        Returns:
            bool: 성공 여부
        """
        # 중복 이름 체크
        if any(r.name == report.name for r in self.reports):
            print(f"이미 존재하는 리포트 이름입니다: {report.name}")
            return False

        self.reports.append(report)
        self._save_config()
        print(f"리포트 추가됨: {report.name}")
        return True

    def remove_report(self, name: str) -> bool:
        """
        리포트 삭제

        Args:
            name: 삭제할 리포트 이름

        Returns:
            bool: 성공 여부
        """
        for i, r in enumerate(self.reports):
            if r.name == name:
                del self.reports[i]
                self._save_config()
                print(f"리포트 삭제됨: {name}")
                return True

        print(f"리포트를 찾을 수 없습니다: {name}")
        return False

    def update_report(self, name: str, **kwargs) -> bool:
        """
        리포트 설정 업데이트

        Args:
            name: 업데이트할 리포트 이름
            **kwargs: 업데이트할 필드들

        Returns:
            bool: 성공 여부
        """
        for r in self.reports:
            if r.name == name:
                for key, value in kwargs.items():
                    if hasattr(r, key):
                        setattr(r, key, value)
                self._save_config()
                print(f"리포트 업데이트됨: {name}")
                return True

        print(f"리포트를 찾을 수 없습니다: {name}")
        return False

    def get_report(self, name: str) -> Optional[ReportConfig]:
        """리포트 설정 조회"""
        for r in self.reports:
            if r.name == name:
                return r
        return None

    def get_enabled_reports(self) -> list[ReportConfig]:
        """활성화된 리포트 목록 반환"""
        return [r for r in self.reports if r.enabled]

    def get_all_reports(self) -> list[ReportConfig]:
        """모든 리포트 목록 반환"""
        return self.reports.copy()

    def enable_report(self, name: str) -> bool:
        """리포트 활성화"""
        return self.update_report(name, enabled=True)

    def disable_report(self, name: str) -> bool:
        """리포트 비활성화"""
        return self.update_report(name, enabled=False)

    def list_reports(self):
        """리포트 목록 출력"""
        if not self.reports:
            print("등록된 리포트가 없습니다.")
            return

        print("\n=== 등록된 리포트 목록 ===")
        for i, r in enumerate(self.reports, 1):
            status = "✓" if r.enabled else "✗"
            print(f"{i}. [{status}] {r.name}")
            print(f"   URL: {r.netsuite_url[:50]}...")
            print(f"   Spreadsheet: {r.spreadsheet_id}")
            if r.worksheet_name:
                print(f"   Worksheet: {r.worksheet_name}")
        print()


def load_reports_from_env() -> list[ReportConfig]:
    """
    환경변수에서 리포트 설정 로드 (기존 단일 리포트 호환)

    환경변수 형식:
    - 단일 리포트: NETSUITE_REPORT_URL, GOOGLE_SPREADSHEET_ID
    - 다중 리포트: NETSUITE_REPORTS (JSON 배열)
    """
    reports = []

    # 방법 1: JSON 배열로 다중 리포트 설정
    reports_json = os.getenv("NETSUITE_REPORTS")
    if reports_json:
        try:
            reports_data = json.loads(reports_json)
            for r in reports_data:
                reports.append(ReportConfig.from_dict(r))
            print(f"환경변수에서 {len(reports)}개 리포트 로드됨")
            return reports
        except json.JSONDecodeError as e:
            print(f"NETSUITE_REPORTS JSON 파싱 오류: {e}")

    # 방법 2: 기존 단일 리포트 환경변수 (하위 호환)
    report_url = os.getenv("NETSUITE_REPORT_URL")
    spreadsheet_id = os.getenv("GOOGLE_SPREADSHEET_ID")
    worksheet_name = os.getenv("GOOGLE_WORKSHEET_NAME")

    if report_url and spreadsheet_id:
        reports.append(ReportConfig(
            name="default",
            netsuite_url=report_url,
            spreadsheet_id=spreadsheet_id,
            worksheet_name=worksheet_name
        ))
        print("환경변수에서 단일 리포트 로드됨 (레거시 모드)")

    return reports


if __name__ == "__main__":
    # CLI로 리포트 관리
    import sys

    manager = ReportConfigManager()

    if len(sys.argv) < 2:
        print("사용법:")
        print("  python report_config.py list              - 리포트 목록 보기")
        print("  python report_config.py add <name> <url> <spreadsheet_id> [worksheet]")
        print("  python report_config.py remove <name>     - 리포트 삭제")
        print("  python report_config.py enable <name>     - 리포트 활성화")
        print("  python report_config.py disable <name>    - 리포트 비활성화")
        sys.exit(0)

    command = sys.argv[1]

    if command == "list":
        manager.list_reports()

    elif command == "add" and len(sys.argv) >= 5:
        name = sys.argv[2]
        url = sys.argv[3]
        spreadsheet_id = sys.argv[4]
        worksheet = sys.argv[5] if len(sys.argv) > 5 else None

        report = ReportConfig(
            name=name,
            netsuite_url=url,
            spreadsheet_id=spreadsheet_id,
            worksheet_name=worksheet
        )
        manager.add_report(report)

    elif command == "remove" and len(sys.argv) >= 3:
        manager.remove_report(sys.argv[2])

    elif command == "enable" and len(sys.argv) >= 3:
        manager.enable_report(sys.argv[2])

    elif command == "disable" and len(sys.argv) >= 3:
        manager.disable_report(sys.argv[2])

    else:
        print("잘못된 명령어입니다. 'python report_config.py' 로 사용법을 확인하세요.")
