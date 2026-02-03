"""
메인 윈도우 UI 컴포넌트

메인 페이지와 수업준비 화면을 관리하는 윈도우
"""
import json
import os
import sys
from datetime import datetime

from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QScrollArea, QStackedWidget, QMessageBox, QDialog)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QCloseEvent
from database.sqlite_connection import SQLiteConnection
from core.models import Worksheet
from database.repositories import WorksheetRepository
from ui.components.header import Header
from ui.components.sidebar import Sidebar
from ui.screens.login_screen import LoginScreen
from services.login_api import load_session, save_session, clear_session


def _get_app_root():
    """config·DB 경로의 기준 폴더. exe 실행 시에는 exe가 있는 폴더, 개발 시에는 프로젝트 루트."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


class MainWindow(QMainWindow):
    """메인 윈도우"""
    
    def __init__(self):
        super().__init__()
        root = _get_app_root()
        # DB 연결 (SQLite 단일 파일). config는 root 기준으로 로드
        config_path = os.path.join(root, "config", "config.json")
        db_path = "./db/ch_lms.db"
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            db_path = (cfg.get("database") or {}).get("path") or db_path
        except Exception:
            pass
        if not os.path.isabs(db_path):
            db_path = os.path.join(root, db_path)
        self.db_connection = SQLiteConnection(db_path)
        self.db_connection.connect()

        # 기본 경로 실패 시 쓰기 가능한 대체 경로로 재시도 (테스트 앱/권한 이슈 대응)
        if not self.db_connection.is_connected():
            fallback_base = os.environ.get("LOCALAPPDATA") or os.environ.get("TEMP") or os.path.expanduser("~")
            fallback_path = os.path.join(fallback_base, "HanQ", "db", "ch_lms.db")
            self.db_connection = SQLiteConnection(fallback_path)
            self.db_connection.connect()

        if not self.db_connection.is_connected():
            QMessageBox.information(
                self,
                "DB 연결 실패",
                "로컬 DB 파일에 연결할 수 없습니다.\n"
                "로그인은 가능하나, 워크시트·문제은행 등 일부 기능이 제한됩니다."
            )

        self.init_ui()
    
    def init_ui(self):
        """UI 초기화: 로그인 화면만 먼저 표시. 메인 콘텐츠(win32com 등)는 로그인 성공 시 로드."""
        self.setWindowTitle("HanQ")
        self.setMinimumSize(1200, 700)
        self.resize(1600, 900)

        self.main_stack = QStackedWidget()
        self.setCentralWidget(self.main_stack)

        self.login_screen = LoginScreen(self)
        self.login_screen.login_succeeded.connect(self._on_login_succeeded)
        self.main_stack.addWidget(self.login_screen)

        session = load_session()
        if session:
            self._build_main_content()
            self.header.set_user_display(session.get("name") or session.get("user_id") or "")
            self.header.set_admin_mode((session.get("user_id") or "").strip().lower() == "admin")
            self.main_stack.setCurrentIndex(1)
            self.stacked_widget.setCurrentIndex(0)
        else:
            self.main_stack.setCurrentIndex(0)

    def closeEvent(self, event: QCloseEvent) -> None:
        """앱 종료 시 세션 삭제 → 다음 실행 시 로그인 화면부터 표시."""
        clear_session()
        event.accept()

    def _build_main_content(self):
        """로그인 성공 시 또는 세션 있을 때만 호출. 메인 콘텐츠(헤더+화면들) 생성 및 win32com 등 로드."""
        if self.main_stack.count() >= 2:
            return
        from ui.screens.main_page import MainPageScreen
        from ui.screens.worksheet_list import WorksheetListScreen
        from ui.screens.worksheet_create import WorksheetCreateScreen
        from ui.screens.worksheet_edit import WorksheetEditScreen
        from ui.screens.textbook_db import TextbookDBScreen
        from ui.screens.exam_db import ExamDBScreen
        from ui.screens.class_worksheet import ClassWorksheetScreen
        from ui.screens.admin import AdminScreen, MemberManagementView

        main_content = QWidget()
        main_layout = QVBoxLayout(main_content)
        main_layout.setSpacing(0)
        main_layout.setContentsMargins(0, 0, 0, 0)

        self.header = Header()
        self.header.tab_changed.connect(self.on_header_tab_changed)
        self.header.logo_clicked.connect(self.on_logo_clicked)
        self.header.profile_edit_clicked.connect(self._on_profile_edit_clicked)
        self.header.logout_btn.clicked.connect(self._on_logout)
        main_layout.addWidget(self.header)
        # 헤더 생성 직후 관리자 여부 설정 (비관리자일 때 회원관리 탭 숨김)
        session = load_session()
        is_admin = bool(session and (session.get("user_id") or "").strip().lower() == "admin")
        self.header.set_admin_mode(is_admin)

        bottom_layout = QHBoxLayout()
        bottom_layout.setSpacing(0)
        bottom_layout.setContentsMargins(0, 0, 0, 0)

        self.sidebar = Sidebar()
        self.sidebar.setMinimumWidth(180)
        self.sidebar.setMaximumWidth(180)
        self.sidebar.hide()
        self.sidebar.menu_clicked.connect(self.on_sidebar_menu_clicked)
        bottom_layout.addWidget(self.sidebar)

        self.stacked_widget = QStackedWidget()

        self.main_page = MainPageScreen(self.db_connection)
        self.main_page.start_requested.connect(self._go_to_class_prep)
        main_page_scroll = QScrollArea()
        main_page_scroll.setWidget(self.main_page)
        main_page_scroll.setWidgetResizable(True)
        main_page_scroll.setFrameShape(QScrollArea.NoFrame)
        main_page_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        main_page_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.stacked_widget.addWidget(main_page_scroll)

        self.worksheet_list_screen = WorksheetListScreen(self.db_connection)
        self.worksheet_list_screen.create_requested.connect(self.on_create_worksheet_requested)
        worksheet_scroll = QScrollArea()
        worksheet_scroll.setWidget(self.worksheet_list_screen)
        worksheet_scroll.setWidgetResizable(True)
        worksheet_scroll.setFrameShape(QScrollArea.NoFrame)
        worksheet_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        worksheet_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.stacked_widget.addWidget(worksheet_scroll)

        self.worksheet_create_screen = WorksheetCreateScreen(self.db_connection)
        self.worksheet_create_screen.close_requested.connect(self.on_create_screen_closed)
        self.worksheet_create_screen.preview_requested.connect(self.on_preview_requested)
        self.stacked_widget.addWidget(self.worksheet_create_screen)

        self.worksheet_edit_screen = WorksheetEditScreen(self.db_connection)
        self.worksheet_edit_screen.back_requested.connect(self.on_edit_back_requested)
        self.worksheet_edit_screen.finalized.connect(self.on_edit_finalized)
        self.stacked_widget.addWidget(self.worksheet_edit_screen)

        self.textbook_db_screen = TextbookDBScreen(self.db_connection)
        textbook_db_scroll = QScrollArea()
        textbook_db_scroll.setWidget(self.textbook_db_screen)
        textbook_db_scroll.setWidgetResizable(True)
        textbook_db_scroll.setFrameShape(QScrollArea.NoFrame)
        textbook_db_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        textbook_db_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.stacked_widget.addWidget(textbook_db_scroll)

        self.exam_db_screen = ExamDBScreen(self.db_connection)
        exam_db_scroll = QScrollArea()
        exam_db_scroll.setWidget(self.exam_db_screen)
        exam_db_scroll.setWidgetResizable(True)
        exam_db_scroll.setFrameShape(QScrollArea.NoFrame)
        exam_db_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        exam_db_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.stacked_widget.addWidget(exam_db_scroll)

        self.class_worksheet_screen = ClassWorksheetScreen(self.db_connection)
        self.stacked_widget.addWidget(self.class_worksheet_screen)

        self.admin_screen = AdminScreen(self.db_connection)
        self.stacked_widget.addWidget(self.admin_screen)

        # 회원관리 전용 화면 (상단 '회원관리' 탭 클릭 시: 학생관리/반관리 사이드바 없이 회원 목록만)
        member_only_widget = QWidget()
        member_only_layout = QVBoxLayout(member_only_widget)
        member_only_layout.setContentsMargins(24, 16, 24, 24)
        member_only_layout.setSpacing(0)
        self.member_management_standalone = MemberManagementView(member_only_widget)
        member_only_layout.addWidget(self.member_management_standalone)
        self.stacked_widget.addWidget(member_only_widget)

        bottom_layout.addWidget(self.stacked_widget)
        main_layout.addLayout(bottom_layout)

        self.main_stack.addWidget(main_content)
        self.stacked_widget.setCurrentIndex(0)
    
    def on_header_tab_changed(self, tab_name):
        """헤더 탭 변경 시 호출"""
        if tab_name == "수업준비":
            # 사이드바 표시
            self.sidebar.show()
            # 학습지 목록 화면으로 전환
            self.stacked_widget.setCurrentIndex(1)
            try:
                self.worksheet_list_screen.reload_from_db()
            except Exception:
                pass
        elif tab_name == "수업":
            # 수업 탭: 메인(수업준비) 사이드바는 숨기고, 탭 내부 UI 사용
            self.sidebar.hide()
            # 수업 화면으로 전환 (인덱스: 6) + 관리에서 등록한 학생/반이 바로 보이도록 새로고침
            self.stacked_widget.setCurrentIndex(6)
            try:
                self.class_worksheet_screen.refresh_from_db()
            except Exception:
                pass
        elif tab_name == "관리":
            # 관리 탭: 메인(수업준비) 사이드바는 숨기고, 탭 내부 UI 사용
            self.sidebar.hide()
            # 관리 화면으로 전환 (인덱스: 7)
            self.stacked_widget.setCurrentIndex(7)
        elif tab_name == "회원관리":
            # 회원관리 탭: 회원관리 전용 화면만 표시 (학생관리/반관리 사이드바 없음, 인덱스 8)
            self.sidebar.hide()
            self.stacked_widget.setCurrentIndex(8)
            self.member_management_standalone.load_users()
        else:
            # 사이드바 숨김
            self.sidebar.hide()
            # 메인 페이지로 전환
            self.stacked_widget.setCurrentIndex(0)
    
    def on_sidebar_menu_clicked(self, menu_name):
        """사이드바 메뉴 클릭 시 호출"""
        if menu_name == "학습지":
            self.stacked_widget.setCurrentIndex(1)
            try:
                self.worksheet_list_screen.reload_from_db()
            except Exception:
                pass
        elif menu_name == "교재DB":
            self.stacked_widget.setCurrentIndex(4)
            # 목록 새로고침
            self.textbook_db_screen.load_textbooks()
        elif menu_name == "기출DB":
            self.stacked_widget.setCurrentIndex(5)
            # 목록 새로고침
            self.exam_db_screen.load_exams()
        # 추후 다른 메뉴 구현 시 여기에 추가
    
    def on_create_worksheet_requested(self):
        """학습지 생성 요청 시 호출"""
        # 요구사항: 생성 화면에서는 상단바 메뉴가 선택되지 않은 상태 유지
        try:
            self.header.clear_tab_selection()
        except Exception:
            pass
        # 학습지 생성 화면으로 전환 (인덱스 2)
        self.stacked_widget.setCurrentIndex(2)
    
    def on_create_screen_closed(self):
        """학습지 생성 화면 닫기 요청 시 호출"""
        # 학습지 목록 화면으로 돌아가기 (인덱스 1)
        self.stacked_widget.setCurrentIndex(1)

    def on_preview_requested(self, payload: dict):
        """학습지 생성 → 문항 편집 화면으로 이동"""
        from ui.screens.worksheet_edit import WorksheetDraft
        try:
            draft = WorksheetDraft(**payload["draft"])
            self.worksheet_edit_screen.load_draft(draft)
            self.stacked_widget.setCurrentIndex(3)
        except Exception as e:
            QMessageBox.critical(self, "오류", f"문항 편집 화면을 열 수 없습니다.\n\n{e}")

    def on_edit_back_requested(self):
        """문항 편집 → 학습지 생성으로 돌아가기"""
        self.stacked_widget.setCurrentIndex(2)

    def on_edit_finalized(self, payload: dict):
        """최종 생성 완료 후 처리 (번호 매핑 포함)."""
        try:
            output_path = str(payload.get("output_path") or "").strip()
            numbered = list(payload.get("numbered") or [])
            title = str(payload.get("title") or "").strip()
            creator = str(payload.get("creator") or "").strip()
            grade = str(payload.get("grade") or "").strip()
            type_text = str(payload.get("type_text") or "").strip()
            problem_ids = list(payload.get("problem_ids") or []) or [x.get("problem_id") for x in numbered if isinstance(x, dict) and x.get("problem_id")]
        except Exception:
            QMessageBox.warning(self, "저장 실패", "완료 데이터를 해석할 수 없습니다.")
            return

        # 로컬 저장(HWP)은 이미 완료된 상태. DB 저장은 가능할 때만 추가로 수행.
        if not output_path or not os.path.exists(output_path):
            QMessageBox.warning(self, "저장 실패", "생성된 HWP 파일을 찾을 수 없습니다.")
            return

        if not self.db_connection.is_connected():
            # 오프라인 모드: 로컬 파일은 남아있으므로 별도 경고만
            QMessageBox.information(self, "오프라인", "DB에 연결할 수 없어 학습지 목록에 저장하지 못했습니다.\n로컬 HWP 파일은 생성되었습니다.")
            return

        try:
            with open(output_path, "rb") as f:
                hwp_bytes = f.read()
        except Exception as e:
            QMessageBox.critical(self, "저장 실패", f"HWP 파일을 읽을 수 없습니다.\n\n{e}")
            return

        try:
            ws = Worksheet(
                title=title,
                grade=grade,
                type_text=type_text,
                creator=creator,
                created_at=datetime.now(),
                problem_ids=[str(x) for x in (problem_ids or []) if x],
                numbered=numbered,
            )
            repo = WorksheetRepository(self.db_connection)
            _ws_id = repo.create(
                ws,
                hwp_bytes=hwp_bytes,
                hwp_filename=os.path.basename(output_path) or None,
            )
        except Exception as e:
            QMessageBox.warning(self, "저장 실패", f"학습지 목록(DB)에 저장하지 못했습니다.\n\n{e}")
            return

        # 목록 갱신 + 목록 화면으로 이동
        try:
            self.worksheet_list_screen.reload_from_db()
        except Exception:
            pass
        self.stacked_widget.setCurrentIndex(1)
    
    def on_logo_clicked(self):
        """로고 클릭 시 호출 - 메인 페이지로 이동"""
        # 모든 탭 선택 해제(Exclusive 그룹에서도 확실히)
        try:
            self.header.clear_tab_selection()
        except Exception:
            for btn in self.header.tab_group.buttons():
                btn.setChecked(False)
        
        # 사이드바 메뉴 선택 해제
        self.sidebar.clear_selection()
        
        # 사이드바 숨김
        self.sidebar.hide()
        
        # 메인 페이지로 전환
        self.stacked_widget.setCurrentIndex(0)

    def _on_login_succeeded(self, user_id: str, name: str):
        """로그인 성공 시: 세션 저장, 메인 콘텐츠 로드(최초 1회), 헤더 이름 표시, 메인 화면으로 전환"""
        save_session(user_id, name)
        self._build_main_content()
        self.header.set_user_display(name)
        self.header.set_admin_mode((user_id or "").strip().lower() == "admin")
        self.main_stack.setCurrentIndex(1)

    def _on_profile_edit_clicked(self):
        """정보수정 클릭: 다이얼로그 표시 후 성공 시 세션·헤더 갱신"""
        session = load_session()
        if not session:
            QMessageBox.information(self, "알림", "로그인된 상태가 아닙니다.")
            return
        from ui.components.profile_edit_dialog import ProfileEditDialog
        dlg = ProfileEditDialog(
            session.get("user_id") or "",
            session.get("name") or "",
            self,
        )
        if dlg.exec_() == QDialog.Accepted:
            user_id, name = dlg.get_updated_profile()
            save_session(user_id, name)
            self.header.set_user_display(name)
            self.header.set_admin_mode((user_id or "").strip().lower() == "admin")

    def _on_logout(self):
        """로그아웃: 세션 삭제 후 로그인 화면으로"""
        clear_session()
        if hasattr(self, "header") and self.header is not None:
            self.header.set_admin_mode(False)
        self.main_stack.setCurrentIndex(0)


    def _go_to_class_prep(self):
        """메인 페이지 '시작하기' → 수업준비(학습지)로 이동"""
        try:
            # 헤더 탭 상태도 함께 맞춤
            for btn in self.header.tab_group.buttons():
                if btn.text() == "수업준비":
                    btn.setChecked(True)
                else:
                    btn.setChecked(False)
        except Exception:
            pass
        self.sidebar.show()
        self.stacked_widget.setCurrentIndex(1)
        try:
            self.worksheet_list_screen.reload_from_db()
        except Exception:
            pass
