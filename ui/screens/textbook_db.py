"""
교재DB 화면

교재 메타데이터 관리 및 파싱 결과 조회 화면
"""
from typing import List, Optional, Dict, Any

from PyQt5.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QHeaderView,
    QFileDialog,
    QMessageBox,
    QProgressDialog,
    QLabel,
    QSplitter,
    QComboBox,
    QCheckBox,
    QLineEdit,
    QFrame,
    QSizePolicy,
    QStyledItemDelegate,
    QStyle,
)
from PyQt5.QtCore import Qt, QEvent
from PyQt5.QtGui import QFont, QColor, QPainter
from database.sqlite_connection import SQLiteConnection
from database.repositories import TextbookRepository
from services.parsing import ParsingService, ReparseMode
from services.problem import ProblemService
from core.models import Textbook, SourceType
from ui.components.metadata_input_dialog import TextbookMetadataDialog
from ui.components.problem_detail_dialog import ProblemDetailDialog
from utils.hwp_restore import HWPRestore
from processors.hwp.hwp_reader import HWPNotInstalledError, HWPInitializationError


class TextbookDBScreen(QWidget):
    """교재DB 화면"""
    
    def __init__(self, db_connection: SQLiteConnection, parent=None):
        """
        TextbookDBScreen 초기화
        
        Args:
            db_connection: DB 연결 인스턴스
            parent: 부모 위젯
        """
        super().__init__(parent)
        self.db_connection = db_connection
        self.textbook_repo = TextbookRepository(db_connection)
        self.parsing_service = ParsingService(db_connection)
        self.problem_service = ProblemService(db_connection)
        self.hwp_restore = HWPRestore(db_connection)
        self.current_textbook_id = None
        self.init_ui()
        self.load_textbooks()
    
    def init_ui(self):
        """UI 초기화"""
        self.setObjectName("TextbookDBRoot")
        self._textbooks_cache: List[Textbook] = []

        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(10, 16, 24, 24)  # 좌측 여백 미세 조정, 우측 여백 유지
        
        # 상단 컨트롤 영역
        control_layout = QHBoxLayout()
        control_layout.setSpacing(10)
        # ✅ 검색창 하단 여백 최소화(표를 더 위로 끌어올림)
        control_layout.setContentsMargins(0, 0, 0, 5)

        # 검색창
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("교재명 또는 단원명을 검색하세요...")
        self.search_input.setClearButtonEnabled(True)
        self.search_input.setFixedHeight(40)
        self.search_input.setMinimumWidth(520)
        self.search_input.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.search_input.textChanged.connect(self._apply_textbook_filters)
        control_layout.addWidget(self.search_input)

        control_layout.addStretch(1)

        # 우측 상단 버튼(주요/보조)
        self.btn_generate_preview = QPushButton("미리보기 일괄생성")
        self.btn_generate_preview.setObjectName("secondary")
        self.btn_generate_preview.setFixedHeight(40)
        self.btn_generate_preview.setCursor(Qt.PointingHandCursor)
        self.btn_generate_preview.setFocusPolicy(Qt.NoFocus)
        self.btn_generate_preview.clicked.connect(self.on_generate_previews_clicked)
        self.btn_generate_preview.setEnabled(False)
        control_layout.addWidget(self.btn_generate_preview)

        btn_create = QPushButton("DB 생성")
        btn_create.setObjectName("primary")
        btn_create.setFixedHeight(40)
        btn_create.setMinimumWidth(110)
        btn_create.setCursor(Qt.PointingHandCursor)
        btn_create.setFocusPolicy(Qt.NoFocus)
        btn_create.clicked.connect(self.on_create_db)
        control_layout.addWidget(btn_create)

        main_layout.addLayout(control_layout)
        # ✅ 검색창 하단 ↔ 표 헤더 사이 여백(10px 이내로 압축)
        main_layout.addSpacing(8)
        
        # 스플리터 (테이블 + Problem 목록)
        splitter = QSplitter(Qt.Horizontal)
        
        # 좌측: 교재 테이블
        table_widget = QFrame()
        table_widget.setObjectName("leftCard")
        table_layout = QVBoxLayout(table_widget)
        table_layout.setContentsMargins(14, 14, 14, 14)
        table_layout.setSpacing(8)

        left_title = QLabel("교재 목록")
        left_title.setObjectName("panelTitle")
        table_layout.addWidget(left_title)
        
        self.table = QTableWidget()
        self.table.setObjectName("DBTable")
        self.table.setColumnCount(9)
        self.table.setHorizontalHeaderLabels([
            "출처", "과목", "대단원", "소단원", "교재명", "생성일", "문제수", "상태", "⋯"
        ])
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)  # 교재명 자동 확장
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        # ✅ 여러 행 동시 선택(멀티 삭제용)
        self.table.setSelectionMode(QTableWidget.ExtendedSelection)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setAlternatingRowColors(False)
        self.table.setShowGrid(True)
        self.table.verticalHeader().setVisible(False)
        self.table.setMouseTracking(True)
        # ✅ 테이블 수직 밀도(행 높이) 강제
        try:
            self.table.verticalHeader().setDefaultSectionSize(38)
        except Exception:
            pass

        # 선택 효과(배경 + 좌측 블루 라인) - delegate로 구현
        self.table.setItemDelegate(_RowSelectDelegate(self.table))
        # 가독성: 컬럼 너비에 여유를 줌(학습지 화면 스타일)
        self.table.setColumnWidth(0, 90)   # 출처
        self.table.setColumnWidth(1, 140)  # 과목
        self.table.setColumnWidth(2, 150)  # 대단원
        self.table.setColumnWidth(3, 150)  # 소단원
        self.table.setColumnWidth(5, 130)  # 생성일
        self.table.setColumnWidth(6, 90)   # 문제수
        self.table.setColumnWidth(7, 90)   # 상태
        self.table.setColumnWidth(8, 70)   # ⋯
        
        # 헤더 가운데 정렬(전체 컬럼)
        for col in range(9):
            header_item = self.table.horizontalHeaderItem(col)
            if header_item:
                header_item.setTextAlignment(Qt.AlignCenter)
        
        # 행 선택 시 Problem 목록 조회
        self.table.selectionModel().selectionChanged.connect(self.on_table_selection_changed)
        
        # 더보기 메뉴 (우클릭)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.on_table_context_menu)
        
        table_layout.addWidget(self.table)
        splitter.addWidget(table_widget)
        
        # 우측: Problem 목록
        problem_widget = QFrame()
        problem_widget.setObjectName("PreviewPanel")
        problem_layout = QVBoxLayout(problem_widget)
        problem_layout.setContentsMargins(14, 14, 14, 14)
        problem_layout.setSpacing(8)
        
        problem_header_layout = QHBoxLayout()
        problem_header_layout.setContentsMargins(0, 0, 0, 0)

        problem_label = QLabel("문제 목록")
        problem_label.setObjectName("panelTitle")
        problem_header_layout.addWidget(problem_label)
        problem_header_layout.addStretch()

        self.only_untagged_checkbox = QCheckBox("미지정만")
        self.only_untagged_checkbox.setObjectName("onlyUntagged")
        self.only_untagged_checkbox.stateChanged.connect(self.on_only_untagged_changed)
        problem_header_layout.addWidget(self.only_untagged_checkbox)

        problem_layout.addLayout(problem_header_layout)
        
        self.problem_table = QTableWidget()
        self.problem_table.setObjectName("ProblemTable")
        self.problem_table.setColumnCount(4)
        self.problem_table.setHorizontalHeaderLabels([
            "#", "미리보기", "난이도", "원본"
        ])
        self.problem_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)  # 미리보기 자동 확장
        self.problem_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.problem_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.problem_table.setAlternatingRowColors(False)
        self.problem_table.setShowGrid(True)
        self.problem_table.verticalHeader().setVisible(False)
        # ✅ 테이블 수직 밀도(행 높이) 강제
        try:
            self.problem_table.verticalHeader().setDefaultSectionSize(38)
        except Exception:
            pass
        self.problem_table.setColumnWidth(0, 50)   # #
        self.problem_table.setColumnWidth(2, 85)   # 난이도(넓게)
        self.problem_table.setColumnWidth(3, 55)   # 원본(좁게)
        self.problem_table.setWordWrap(False)

        # 헤더 정렬(미리보기 컬럼 제외)
        for col in range(4):
            if col != 1:
                header_item = self.problem_table.horizontalHeaderItem(col)
                if header_item:
                    header_item.setTextAlignment(Qt.AlignCenter)

        # 셀 위젯(난이도 콤보) 선택 배경 동기화용
        self._problem_cell_widgets = {}
        self.problem_table.selectionModel().selectionChanged.connect(self.on_problem_selection_changed)
        
        # Problem 행 더블클릭 시 상세 보기
        self.problem_table.itemDoubleClicked.connect(self.on_problem_double_clicked)
        # 난이도 단축키(0~4)
        self.problem_table.installEventFilter(self)
        
        problem_layout.addWidget(self.problem_table)
        splitter.addWidget(problem_widget)
        
        # 스플리터 비율 설정 (좌측 60%, 우측 40%)
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 2)
        splitter.setSizes([600, 400])
        
        main_layout.addWidget(splitter)
        
        # ✅ 전면 덮어쓰기: Modern Dashboard 스타일
        self.setStyleSheet(
            """
            QWidget#TextbookDBRoot {
                background: transparent;
                font-family: 'Pretendard','Malgun Gothic','맑은 고딕';
            }
            QWidget#TextbookDBRoot * {
                outline: none;
            }

            QLabel#panelTitle {
                color: #222222;
                font-size: 12pt;
                font-weight: 800;
            }

            QLineEdit {
                background-color: #FFFFFF;
                border: 1.5px solid #CBD5E1;
                border-radius: 8px;
                padding: 8px 15px;
                color: #222222;
                font-weight: 600;
            }
            QLineEdit::placeholder {
                color: #64748B;
            }
            QLineEdit:focus {
                border: 2px solid #2563EB;
                padding: 7px 14px;
            }

            QPushButton#primary {
                background-color: #2563EB;
                color: #FFFFFF;
                border: none;
                border-radius: 10px;
                padding: 9px 16px;
                font-weight: 800;
            }
            QPushButton#primary:hover { background-color: #1D4ED8; }

            QPushButton#secondary {
                background-color: #FFFFFF;
                color: #2563EB;
                border: 1px solid #2563EB;
                border-radius: 10px;
                padding: 9px 16px;
                font-weight: 800;
            }
            QPushButton#secondary:hover { background-color: #EFF6FF; }
            QPushButton#secondary:disabled { color: #94A3B8; border-color: #CBD5E1; }

            QFrame#leftCard, QFrame#PreviewPanel {
                background-color: #FFFFFF;
                border: 1px solid #E2E8F0;
                border-radius: 12px;
            }

            QCheckBox#onlyUntagged {
                color: #222222;
                font-weight: 800;
                background: none;
            }
            QCheckBox#onlyUntagged::indicator {
                width: 18px;
                height: 18px;
                background: #FFFFFF;
                border: 2px solid #334155;
                border-radius: 4px;
            }
            QCheckBox#onlyUntagged::indicator:checked {
                background: #2563EB;
                border-color: #2563EB;
            }
            QCheckBox#onlyUntagged::indicator:disabled {
                background: #F1F5F9;
                border-color: #CBD5E1;
            }

            QTableWidget#DBTable, QTableWidget#ProblemTable {
                background-color: #FFFFFF;
                border: 1px solid #E2E8F0;
                border-radius: 12px;
                gridline-color: #E0E0E0;
                selection-color: #222222;
                color: #222222;
                outline: none;
            }
            QTableWidget#ProblemTable {
                selection-background-color: #E8F2FF;
            }
            QTableWidget::item {
                padding-top: 2px;
                padding-bottom: 2px;
                padding-left: 12px;
                padding-right: 12px;
                color: #222222;
                background: none;
                font-weight: 700;
                border: none;
                border-bottom: 1px solid #F1F5F9;
            }
            QTableWidget::item:selected {
                background: none;
                color: #222222;
                font-weight: 700;
            }
            QTableWidget#ProblemTable::item:selected {
                background-color: #E8F2FF;
                color: #222222;
                border: none;
                border-bottom: 1px solid #D6E8FF;
            }
            QTableWidget::item:hover {
                background: none;
            }

            QHeaderView::section {
                background-color: #F8FAFC;
                color: #222222;
                font-weight: 800;
                padding: 0px 10px;
                border: none;
                border-right: 1px solid #F1F5F9;
                border-bottom: 1px solid #F1F5F9;
            }
            QHeaderView::section:last {
                border-right: none;
            }

            QLabel#StatusBadgeOk {
                background-color: #F0FDF4;
                color: #166534;
                border-radius: 8px;
                padding: 4px 10px;
                font-weight: 800;
            }
            QLabel#StatusBadgeWarn {
                background-color: #FFFBEB;
                color: #92400E;
                border-radius: 8px;
                padding: 4px 10px;
                font-weight: 800;
            }
            QLabel#StatusBadgeFail {
                background-color: #FEF2F2;
                color: #991B1B;
                border-radius: 8px;
                padding: 4px 10px;
                font-weight: 800;
            }

            QComboBox#DifficultyCombo {
                background: transparent;
                border: none;
                border-radius: 0;
                padding: 2px 4px;
                min-width: 0;
                max-width: 50px;
                font-size: 9pt;
                color: #222222;
                font-weight: bold;
            }
            QComboBox#DifficultyCombo::drop-down { border: none; }
            QComboBox#DifficultyCombo QAbstractItemView {
                background-color: #FFFFFF;
                border: 1px solid #E0E0E0;
                selection-background-color: #F0F7FF;
                selection-color: #222222;
                outline: 0;
            }
            """
        )

        # 헤더 높이(가독성)
        try:
            # ✅ 헤더 높이도 슬림하게
            self.table.horizontalHeader().setFixedHeight(32)
            self.problem_table.horizontalHeader().setFixedHeight(32)
        except Exception:
            pass
    
    def load_textbooks(self):
        """교재 목록 로드"""
        try:
            self._textbooks_cache = self.textbook_repo.list_all()
            self._apply_textbook_filters()
        except ConnectionError as e:
            QMessageBox.warning(self, "연결 오류", f"DB에 연결할 수 없습니다.\n\n{str(e)}")
        except Exception as e:
            QMessageBox.critical(self, "오류", f"교재 목록을 불러올 수 없습니다.\n\n{str(e)}")

    def _apply_textbook_filters(self):
        textbooks = list(self._textbooks_cache or [])
        query = (self.search_input.text() if getattr(self, "search_input", None) else "") or ""
        q = query.strip().lower()

        filtered: List[Textbook] = []
        for tb in textbooks:
            if q:
                hay = f"{tb.subject} {tb.major_unit} {tb.sub_unit or ''} {tb.name}".lower()
                if q not in hay:
                    continue
            filtered.append(tb)

        self.table.setRowCount(len(filtered))
            
        for row, textbook in enumerate(filtered):
                # 출처
                item = QTableWidgetItem("교재")
                item.setTextAlignment(Qt.AlignCenter)
                self.table.setItem(row, 0, item)
                
                # 과목
                item = QTableWidgetItem(textbook.subject)
                item.setTextAlignment(Qt.AlignCenter)
                self.table.setItem(row, 1, item)
                
                # 대단원
                item = QTableWidgetItem(textbook.major_unit)
                item.setTextAlignment(Qt.AlignCenter)
                self.table.setItem(row, 2, item)
                
                # 소단원
                item = QTableWidgetItem(textbook.sub_unit or "")
                item.setTextAlignment(Qt.AlignCenter)
                self.table.setItem(row, 3, item)
                
                # 교재명
                item = QTableWidgetItem(textbook.name)
                item.setTextAlignment(Qt.AlignVCenter | Qt.AlignLeft)
                self.table.setItem(row, 4, item)
                # ID 저장 (더보기 메뉴용)
                item.setData(Qt.UserRole, textbook.id)
                
                # 생성일
                if textbook.created_at:
                    date_str = textbook.created_at.strftime("%Y.%m.%d")
                else:
                    date_str = ""
                item = QTableWidgetItem(date_str)
                item.setTextAlignment(Qt.AlignCenter)
                self.table.setItem(row, 5, item)
                
                # 문제수
                item = QTableWidgetItem(str(textbook.problem_count))
                item.setTextAlignment(Qt.AlignCenter)
                self.table.setItem(row, 6, item)
                
                # 상태
                if textbook.is_parsed:
                    status = "완료" if (textbook.problem_count or 0) > 0 else "부분"
                else:
                    status = "실패"

                badge = QLabel(status)
                if status == "완료":
                    badge.setObjectName("StatusBadgeOk")
                elif status == "부분":
                    badge.setObjectName("StatusBadgeWarn")
                else:
                    badge.setObjectName("StatusBadgeFail")
                badge.setAlignment(Qt.AlignCenter)
                badge_wrap = QWidget()
                wrap_l = QHBoxLayout(badge_wrap)
                wrap_l.setContentsMargins(0, 0, 0, 0)
                wrap_l.addStretch()
                wrap_l.addWidget(badge)
                wrap_l.addStretch()
                self.table.setCellWidget(row, 7, badge_wrap)
                
                # 더보기 버튼
                btn_more = QPushButton("⋯")
                btn_more.setMinimumWidth(40)
                btn_more.setMinimumHeight(28)
                btn_more.setFont(QFont("맑은 고딕", 14, QFont.Bold))
                btn_more.setStyleSheet("""
                    QPushButton {
                        background-color: transparent;
                        color: #94A3B8;
                        border: none;
                        border-radius: 8px;
                    }
                    QPushButton:hover {
                        background-color: #F1F5F9;
                        color: #475569;
                    }
                """)
                btn_more.clicked.connect(lambda checked, tid=textbook.id: self.on_more_clicked(tid))
                self.table.setCellWidget(row, 8, btn_more)
                
                # 행 높이(가독성)
                self.table.setRowHeight(row, 38)
    
    def on_create_db(self):
        """DB 생성 버튼 클릭 처리"""
        # 1. HWP 파일 선택
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "HWP 파일 선택",
            "",
            "HWP 파일 (*.hwp);;모든 파일 (*.*)"
        )
        
        if not file_path:
            return
        
        # 2. 메타데이터 입력 다이얼로그
        dialog = TextbookMetadataDialog(self)
        
        def on_metadata_submitted(metadata):
            dialog.close()
            self.process_textbook_creation(file_path, metadata)
        
        dialog.submitted.connect(on_metadata_submitted)
        dialog.exec_()
    
    def process_textbook_creation(self, hwp_path: str, metadata: dict):
        """교재 생성 및 파싱 처리"""
        try:
            # 3. Textbook 생성
            textbook = Textbook(
                name=metadata['name'],
                subject=metadata['subject'],
                major_unit=metadata['major_unit'],
                sub_unit=metadata.get('sub_unit')
            )
            
            textbook_id = self.textbook_repo.create(textbook)
            if not textbook_id:
                QMessageBox.critical(self, "오류", "교재를 생성할 수 없습니다.")
                return
            
            # 4. 파싱 진행 다이얼로그
            progress = QProgressDialog("HWP 파일을 파싱하는 중...", "취소", 0, 100, self)
            progress.setWindowModality(Qt.WindowModal)
            progress.setAutoClose(True)
            progress.setAutoReset(True)
            
            def progress_callback(current, total):
                if total > 0:
                    progress.setMaximum(total)
                    progress.setValue(current)
                if progress.wasCanceled():
                    return False
                return True
            
            # 5. 파싱 실행 (저장되는 한 문제짜리 문서에 스타일 적용)
            result = self.parsing_service.reparse_textbook(
                textbook_id=textbook_id,
                hwp_path=hwp_path,
                mode=ReparseMode.REPLACE,
                creator="",  # 추후 사용자 정보에서 가져오기
                progress_callback=progress_callback,
                apply_style_to_blocks=True
            )
            
            progress.close()
            
            if result['success']:
                QMessageBox.information(
                    self,
                    "완료",
                    f"파싱이 완료되었습니다.\n\n"
                    f"생성된 문제: {result['created_count']}개\n"
                    f"총 문제: {result['total_problems']}개"
                )
                # 목록 새로고침
                self.load_textbooks()
            else:
                QMessageBox.warning(
                    self,
                    "파싱 실패",
                    f"파싱 중 오류가 발생했습니다.\n\n{result.get('error', '알 수 없는 오류')}"
                )
        
        except (HWPNotInstalledError, HWPInitializationError) as e:
            QMessageBox.critical(
                self,
                "한글 프로그램 오류",
                f"한글 프로그램을 사용할 수 없습니다.\n\n{str(e)}\n\n"
                "한글과컴퓨터의 한글 프로그램을 설치한 후 다시 시도해주세요."
            )
        except ConnectionError as e:
            QMessageBox.warning(
                self,
                "연결 오류",
                f"DB에 연결할 수 없습니다.\n\n{str(e)}\n\n"
                "오프라인 모드로 동작 중입니다."
            )
        except Exception as e:
            QMessageBox.critical(self, "오류", f"처리 중 오류가 발생했습니다.\n\n{str(e)}")
    
    def on_table_selection_changed(self, *args):
        """테이블 선택 변경 시 Problem 목록 조회"""
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            self.problem_table.setRowCount(0)
            self.current_textbook_id = None
            try:
                self.btn_generate_preview.setEnabled(False)
            except Exception:
                pass
            return
        
        row = selected_rows[0].row()
        textbook_id_item = self.table.item(row, 4)  # 교재명 셀에 ID 저장
        if not textbook_id_item:
            return
        
        textbook_id = textbook_id_item.data(Qt.UserRole)
        if not textbook_id:
            return
        
        self.current_textbook_id = textbook_id
        try:
            self.btn_generate_preview.setEnabled(True)
        except Exception:
            pass
        self.load_problems(textbook_id)
    
    def load_problems(self, textbook_id: str):
        """Problem 목록 로드"""
        try:
            problems = self.problem_service.get_problems_by_source(
                source_id=textbook_id,
                source_type=SourceType.TEXTBOOK
            )

            # 필터: 미지정만
            if getattr(self, "only_untagged_checkbox", None) and self.only_untagged_checkbox.isChecked():
                problems = [p for p in problems if not p.get("difficulty")]
            
            self.problem_table.setRowCount(len(problems))
            self._problem_cell_widgets = {}
            
            for row, problem in enumerate(problems):
                # 문제 번호
                problem_index = problem.get('problem_index', '')
                if isinstance(problem_index, list):
                    problem_index = problem_index[0] if problem_index else ''
                item = QTableWidgetItem(str(problem_index))
                item.setTextAlignment(Qt.AlignCenter)
                item.setData(Qt.UserRole, problem.get('problem_id'))  # ID 저장
                self.problem_table.setItem(row, 0, item)
                
                # 미리보기
                preview = problem.get('content_text_preview', '')
                if isinstance(preview, list):
                    preview = ' '.join(str(x) for x in preview) if preview else ''
                elif preview is None:
                    preview = ''
                preview_text = str(preview)
                item = QTableWidgetItem(preview_text)
                item.setTextAlignment(Qt.AlignLeft)
                # 툴팁으로 전체 텍스트 확인 가능
                if preview_text:
                    item.setToolTip(preview_text)
                self.problem_table.setItem(row, 1, item)
                
                # 난이도(드롭다운): 고정 너비 50px, 중앙 정렬, 슬림 스타일
                difficulty = problem.get("difficulty")
                combo = QComboBox()
                combo.setObjectName("DifficultyCombo")
                combo.addItems(["미지정", "하", "중", "상", "킬"])
                combo.setFont(QFont("맑은 고딕", 9))
                combo.setFixedWidth(50)
                combo.setMinimumHeight(24)
                try:
                    combo.view().setMinimumWidth(70)
                except Exception:
                    pass

                current_text = difficulty if difficulty else "미지정"
                if current_text in ["하", "중", "상", "킬"]:
                    combo.setCurrentText(current_text)
                else:
                    combo.setCurrentText("미지정")

                problem_id = problem.get("problem_id")
                combo.currentTextChanged.connect(lambda val, pid=problem_id: self.on_difficulty_changed(pid, val))
                wrapper = QWidget()
                wrapper.setStyleSheet("QWidget { background-color: transparent; border: none; }")
                w_layout = QHBoxLayout()
                w_layout.setContentsMargins(4, 0, 4, 0)
                w_layout.setSpacing(0)
                w_layout.addStretch()
                w_layout.addWidget(combo, 0, Qt.AlignCenter)
                w_layout.addStretch()
                wrapper.setLayout(w_layout)
                self.problem_table.setCellWidget(row, 2, wrapper)
                self._problem_cell_widgets[row] = {2: wrapper}
                
                # 원본
                has_raw = problem.get('has_content_raw', False)
                raw_text = "✓" if has_raw else "✗"
                item = QTableWidgetItem(str(raw_text))
                item.setTextAlignment(Qt.AlignCenter)
                self.problem_table.setItem(row, 3, item)

                self.problem_table.setRowHeight(row, 38)
        
        except ConnectionError as e:
            QMessageBox.warning(self, "연결 오류", f"DB에 연결할 수 없습니다.\n\n{str(e)}")
        except Exception as e:
            QMessageBox.critical(self, "오류", f"문제 목록을 불러올 수 없습니다.\n\n{str(e)}")

    def on_problem_selection_changed(self, selected, deselected):
        """문제 테이블 선택 변경 시 난이도 셀 위젯 배경을 행 하이라이트(#E8F2FF)와 동일하게"""
        try:
            deselected_rows = {idx.row() for idx in deselected.indexes()}
            selected_rows = {idx.row() for idx in selected.indexes()}

            for r in deselected_rows:
                widgets = self._problem_cell_widgets.get(r, {})
                for w in widgets.values():
                    w.setStyleSheet("QWidget { background-color: transparent; border: none; }")

            for r in selected_rows:
                widgets = self._problem_cell_widgets.get(r, {})
                for w in widgets.values():
                    w.setStyleSheet("QWidget { background-color: #E8F2FF; border: none; }")
        except Exception:
            pass

    def on_difficulty_changed(self, problem_id: str, value: str):
        """난이도 변경 처리"""
        if not problem_id:
            return
        try:
            difficulty = None if value == "미지정" else value
            self.problem_service.set_problem_difficulty(problem_id, difficulty)

            # 미지정만 필터가 켜져 있으면, 태깅 후 목록 갱신(방금 태깅한 문제를 리스트에서 제거)
            if getattr(self, "only_untagged_checkbox", None) and self.only_untagged_checkbox.isChecked():
                if self.current_textbook_id:
                    self.load_problems(self.current_textbook_id)
        except ConnectionError as e:
            QMessageBox.warning(self, "연결 오류", f"DB에 연결할 수 없습니다.\n\n{str(e)}")
        except Exception as e:
            QMessageBox.critical(self, "오류", f"난이도를 저장할 수 없습니다.\n\n{str(e)}")

    def on_only_untagged_changed(self):
        """미지정만 보기 토글"""
        if self.current_textbook_id:
            self.load_problems(self.current_textbook_id)

    def on_generate_previews_clicked(self):
        """현재 교재의 미리보기 텍스트를 일괄 생성"""
        if not self.current_textbook_id:
            return
        try:
            progress = QProgressDialog("미리보기 텍스트를 생성하는 중...", "취소", 0, 100, self)
            progress.setWindowModality(Qt.WindowModal)
            progress.setAutoClose(True)
            progress.setAutoReset(True)

            def progress_callback(current, total):
                if total <= 0:
                    progress.setMaximum(1)
                    progress.setValue(0)
                else:
                    progress.setMaximum(total)
                    progress.setValue(current)
                if progress.wasCanceled():
                    return False
                return True

            result = self.problem_service.generate_previews_for_source(
                source_id=self.current_textbook_id,
                source_type=SourceType.TEXTBOOK,
                only_missing=True,
                progress_callback=progress_callback,
            )
            progress.close()
            QMessageBox.information(
                self,
                "완료",
                f"미리보기 생성이 완료되었습니다.\n\n"
                f"- 대상: {result.get('total', 0)}개\n"
                f"- 생성: {result.get('updated', 0)}개\n"
                f"- 건너뜀: {result.get('skipped', 0)}개\n"
                f"- 실패: {result.get('failed', 0)}개"
            )
            self.load_problems(self.current_textbook_id)
        except ConnectionError as e:
            QMessageBox.warning(self, "연결 오류", f"DB에 연결할 수 없습니다.\n\n{str(e)}")
        except Exception as e:
            QMessageBox.critical(self, "오류", f"미리보기를 생성할 수 없습니다.\n\n{str(e)}")

    def eventFilter(self, obj, event):
        """난이도 단축키: 1=하, 2=중, 3=상, 4=킬, 0=미지정"""
        if obj is self.problem_table and event.type() == QEvent.KeyPress:
            key = event.key()
            mapping = {
                Qt.Key_1: "하",
                Qt.Key_2: "중",
                Qt.Key_3: "상",
                Qt.Key_4: "킬",
                Qt.Key_0: "미지정",
            }
            if key in mapping:
                selected = self.problem_table.selectionModel().selectedRows()
                if selected:
                    row = selected[0].row()
                    widget = self.problem_table.cellWidget(row, 2)
                    combo: Optional[QComboBox] = None
                    if isinstance(widget, QComboBox):
                        combo = widget
                    elif isinstance(widget, QWidget):
                        try:
                            combo = widget.findChild(QComboBox)
                        except Exception:
                            combo = None
                    if combo is not None:
                        combo.setCurrentText(mapping[key])
                        # 다음 행 자동 이동(필터로 인해 행이 재정렬/삭제될 수 있으니, 단순히 +1만 시도)
                        next_row = min(row + 1, self.problem_table.rowCount() - 1)
                        if next_row != row:
                            self.problem_table.selectRow(next_row)
                return True
        return super().eventFilter(obj, event)
 
    
    def on_problem_double_clicked(self, item):
        """Problem 더블클릭 시 상세 보기"""
        row = item.row()
        problem_id_item = self.problem_table.item(row, 0)
        if not problem_id_item:
            return
        
        problem_id = problem_id_item.data(Qt.UserRole)
        if not problem_id:
            return
        
        try:
            problem_detail = self.problem_service.get_problem_detail(problem_id)
            if not problem_detail:
                QMessageBox.warning(self, "오류", "문제 정보를 찾을 수 없습니다.")
                return
            
            dialog = ProblemDetailDialog(problem_detail, self)
            dialog.set_hwp_restore(self.hwp_restore)
            dialog.exec_()
        
        except ConnectionError as e:
            QMessageBox.warning(self, "연결 오류", f"DB에 연결할 수 없습니다.\n\n{str(e)}")
        except Exception as e:
            QMessageBox.critical(self, "오류", f"문제 상세 정보를 불러올 수 없습니다.\n\n{str(e)}")
    
    def on_more_clicked(self, textbook_id: str):
        """더보기 버튼 클릭 처리"""
        from PyQt5.QtWidgets import QMenu
        
        selected_ids = self._get_selected_textbook_ids()
        target_ids = selected_ids if (textbook_id in selected_ids and len(selected_ids) > 1) else [textbook_id]

        menu = QMenu(self)
        
        action_edit = menu.addAction("수정")
        action_delete = menu.addAction("삭제")
        action_reparse = menu.addAction("재파싱")

        if len(target_ids) != 1:
            action_edit.setEnabled(False)
            action_reparse.setEnabled(False)
        
        action = menu.exec_(self.table.mapToGlobal(self.table.viewport().mapFromGlobal(
            self.table.cursor().pos()
        )))
        
        if action == action_edit:
            self.on_edit_textbook(target_ids[0])
        elif action == action_delete:
            self.on_delete_textbooks(target_ids)
        elif action == action_reparse:
            self.on_reparse_textbook(target_ids[0])
    
    def on_edit_textbook(self, textbook_id: str):
        """교재 수정"""
        try:
            textbook = self.textbook_repo.find_by_id(textbook_id)
            if not textbook:
                QMessageBox.warning(self, "오류", "교재 정보를 찾을 수 없습니다.")
                return

            dialog = TextbookMetadataDialog(self)
            dialog.setWindowTitle("교재 수정")

            # 기존 값 채우기
            try:
                dialog.name_input.setText(textbook.name or "")
            except Exception:
                pass

            try:
                # 과목 → 대단원/소단원 콤보가 연쇄 갱신되므로 순서가 중요
                if textbook.subject:
                    dialog.subject_combo.setCurrentText(textbook.subject)
                if textbook.major_unit:
                    dialog.major_unit_combo.setCurrentText(textbook.major_unit)
                sub = textbook.sub_unit if textbook.sub_unit else "(없음)"
                dialog.sub_unit_combo.setCurrentText(sub)
            except Exception:
                pass

            submitted: Dict = {}

            def _on_submitted(md: dict):
                submitted.clear()
                submitted.update(md)

            dialog.submitted.connect(_on_submitted)
            ok = dialog.exec_()
            if ok != 1:
                return

            md = submitted or {}
            new_name = (md.get("name") or "").strip()
            new_subject = (md.get("subject") or "").strip()
            new_major = (md.get("major_unit") or "").strip()
            new_sub = md.get("sub_unit")

            changed = (
                new_name != (textbook.name or "")
                or new_subject != (textbook.subject or "")
                or new_major != (textbook.major_unit or "")
                or (new_sub or None) != (textbook.sub_unit or None)
            )

            textbook.name = new_name
            textbook.subject = new_subject
            textbook.major_unit = new_major
            textbook.sub_unit = new_sub

            updated = self.textbook_repo.update(textbook)
            if (not updated) and (not changed):
                QMessageBox.information(self, "완료", "변경 사항이 없습니다.")
            elif updated or changed:
                QMessageBox.information(self, "완료", "교재가 수정되었습니다.")
            else:
                QMessageBox.warning(self, "오류", "교재를 수정할 수 없습니다.")

            self.load_textbooks()
            self._select_textbook_row_by_id(textbook_id)
        except ConnectionError as e:
            QMessageBox.warning(self, "연결 오류", f"DB에 연결할 수 없습니다.\n\n{str(e)}")
        except Exception as e:
            QMessageBox.critical(self, "오류", f"수정 중 오류가 발생했습니다.\n\n{str(e)}")
    
    def on_delete_textbook(self, textbook_id: str):
        """교재 삭제"""
        self.on_delete_textbooks([textbook_id])

    def on_delete_textbooks(self, textbook_ids: List[str]):
        """교재 여러 개 삭제 (연결된 문제도 함께 삭제)"""
        ids = [x for x in (textbook_ids or []) if x]
        if not ids:
            return

        msg = (
            f"선택한 교재 {len(ids)}개를 삭제하시겠습니까?\n"
            f"연결된 문제도 함께 삭제됩니다."
        )
        reply = QMessageBox.question(
            self,
            "삭제 확인",
            msg,
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )

        if reply != QMessageBox.Yes:
            return

        try:
            deleted_sources = 0
            deleted_problems = 0
            failed: List[str] = []

            for tid in ids:
                # 연결된 Problem 삭제 (GridFS 포함)
                deleted_problems += self.problem_service.delete_problems_by_source(
                    tid, SourceType.TEXTBOOK
                )
                if self.textbook_repo.delete(tid):
                    deleted_sources += 1
                else:
                    failed.append(tid)

            self.load_textbooks()
            if self.current_textbook_id in ids:
                self.current_textbook_id = None
                self.problem_table.setRowCount(0)

            if failed:
                QMessageBox.warning(
                    self,
                    "부분 완료",
                    f"교재 {deleted_sources}개 삭제, 연결 문제 {deleted_problems}개 삭제 완료.\n\n"
                    f"삭제 실패: {len(failed)}개",
                )
            else:
                QMessageBox.information(
                    self,
                    "완료",
                    f"교재 {deleted_sources}개 삭제, 연결 문제 {deleted_problems}개 삭제 완료.",
                )
        except ConnectionError as e:
            QMessageBox.warning(self, "연결 오류", f"DB에 연결할 수 없습니다.\n\n{str(e)}")
        except Exception as e:
            QMessageBox.critical(self, "오류", f"삭제 중 오류가 발생했습니다.\n\n{str(e)}")
    
    def on_reparse_textbook(self, textbook_id: str):
        """교재 재파싱"""
        # HWP 파일 선택
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "HWP 파일 선택",
            "",
            "HWP 파일 (*.hwp);;모든 파일 (*.*)"
        )
        
        if not file_path:
            return
        
        try:
            # 파싱 진행 다이얼로그
            progress = QProgressDialog("HWP 파일을 재파싱하는 중...", "취소", 0, 100, self)
            progress.setWindowModality(Qt.WindowModal)
            progress.setAutoClose(True)
            progress.setAutoReset(True)
            
            def progress_callback(current, total):
                if total > 0:
                    progress.setMaximum(total)
                    progress.setValue(current)
                if progress.wasCanceled():
                    return False
                return True
            
            # 재파싱 실행 (저장되는 한 문제짜리 문서에 스타일 적용)
            result = self.parsing_service.reparse_textbook(
                textbook_id=textbook_id,
                hwp_path=file_path,
                mode=ReparseMode.REPLACE,
                creator="",
                progress_callback=progress_callback,
                apply_style_to_blocks=True
            )
            
            progress.close()
            
            if result['success']:
                QMessageBox.information(
                    self,
                    "완료",
                    f"재파싱이 완료되었습니다.\n\n"
                    f"생성된 문제: {result['created_count']}개\n"
                    f"총 문제: {result['total_problems']}개"
                )
                # 목록 새로고침
                self.load_textbooks()
                if self.current_textbook_id == textbook_id:
                    self.load_problems(textbook_id)
            else:
                QMessageBox.warning(
                    self,
                    "파싱 실패",
                    f"재파싱 중 오류가 발생했습니다.\n\n{result.get('error', '알 수 없는 오류')}"
                )
        
        except (HWPNotInstalledError, HWPInitializationError) as e:
            QMessageBox.critical(
                self,
                "한글 프로그램 오류",
                f"한글 프로그램을 사용할 수 없습니다.\n\n{str(e)}"
            )
        except ConnectionError as e:
            QMessageBox.warning(self, "연결 오류", f"DB에 연결할 수 없습니다.\n\n{str(e)}")
        except Exception as e:
            QMessageBox.critical(self, "오류", f"재파싱 중 오류가 발생했습니다.\n\n{str(e)}")
    
    def on_table_context_menu(self, position):
        """테이블 우클릭 메뉴"""
        item = self.table.itemAt(position)
        if not item:
            return
        
        row = item.row()
        # 우클릭한 행이 현재 선택에 없으면, 해당 행만 선택(멀티 선택 유지 케이스 고려)
        try:
            if not self.table.selectionModel().isRowSelected(row, self.table.rootIndex()):
                self.table.clearSelection()
                self.table.selectRow(row)
        except Exception:
            pass

        textbook_id_item = self.table.item(row, 4)
        if not textbook_id_item:
            return
        
        textbook_id = textbook_id_item.data(Qt.UserRole)
        if not textbook_id:
            return

        selected_ids = self._get_selected_textbook_ids()
        target_ids = selected_ids if (textbook_id in selected_ids and len(selected_ids) > 1) else [textbook_id]
        
        from PyQt5.QtWidgets import QMenu
        
        menu = QMenu(self)
        action_edit = menu.addAction("수정")
        action_delete = menu.addAction("삭제")
        action_reparse = menu.addAction("재파싱")

        # 수정/재파싱은 1개씩만
        if len(target_ids) != 1:
            action_edit.setEnabled(False)
            action_reparse.setEnabled(False)
        
        action = menu.exec_(self.table.viewport().mapToGlobal(position))
        
        if action == action_edit:
            self.on_edit_textbook(target_ids[0])
        elif action == action_delete:
            self.on_delete_textbooks(target_ids)
        elif action == action_reparse:
            self.on_reparse_textbook(target_ids[0])

    def _get_selected_textbook_ids(self) -> List[str]:
        """교재 테이블에서 선택된 행들의 textbook_id 리스트"""
        ids: List[str] = []
        try:
            selected = self.table.selectionModel().selectedRows()
            for idx in selected:
                r = idx.row()
                item = self.table.item(r, 4)
                if item:
                    tid = item.data(Qt.UserRole)
                    if tid:
                        ids.append(str(tid))
        except Exception:
            return []
        # 중복 제거(순서 유지)
        uniq: List[str] = []
        seen = set()
        for x in ids:
            if x not in seen:
                seen.add(x)
                uniq.append(x)
        return uniq

    def _select_textbook_row_by_id(self, textbook_id: str) -> None:
        """목록 새로고침 후 특정 교재 행을 다시 선택"""
        if not textbook_id:
            return
        try:
            for r in range(self.table.rowCount()):
                item = self.table.item(r, 4)
                if item and item.data(Qt.UserRole) == textbook_id:
                    self.table.selectRow(r)
                    break
        except Exception:
            pass


class _RowSelectDelegate(QStyledItemDelegate):
    """선택 행 강조: 배경 + 좌측 블루 포인트 라인."""

    def paint(self, painter: QPainter, option, index):  # type: ignore[override]
        selected = bool(option.state & QStyle.State_Selected)
        # hover 하이라이트는 "배경 노이즈"가 될 수 있어 사용하지 않음

        # 기본 옵션 복제
        opt = option
        if selected:
            # 기본 selection 제거(직접 그리기)
            opt.state = opt.state & ~QStyle.State_Selected
            opt.palette.setColor(opt.palette.Text, QColor("#000000"))
            f = opt.font
            f.setWeight(QFont.Bold)
            opt.font = f
        else:
            opt.palette.setColor(opt.palette.Text, QColor("#000000"))

        if selected:
            painter.save()
            painter.setPen(Qt.NoPen)
            # 선택 배경은 연한 회색만(요구사항)
            painter.setBrush(QColor("#E2E8F0"))
            painter.drawRect(option.rect)
            if index.column() == 0:
                painter.setBrush(QColor("#2563EB"))
                painter.drawRect(option.rect.left(), option.rect.top(), 4, option.rect.height())
            painter.restore()

        super().paint(painter, opt, index)
