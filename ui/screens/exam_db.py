"""
기출DB 화면

기출 메타데이터 관리 및 파싱 결과 조회 화면
"""
from typing import List, Dict

from PyQt5.QtCore import Qt, QEvent, QSize
from PyQt5.QtGui import QFont, QColor, QPainter, QBrush
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
from database.sqlite_connection import SQLiteConnection
from database.repositories import ExamRepository
from services.parsing import ParsingService, ReparseMode
from services.problem import ProblemService
from core.models import Exam, SourceType
from ui.components.metadata_input_dialog import ExamMetadataDialog
from ui.components.problem_detail_dialog import ProblemDetailDialog
from utils.hwp_restore import HWPRestore
from processors.hwp.hwp_reader import HWPNotInstalledError, HWPInitializationError
from core.unit_catalog import list_subjects, list_major_units, list_sub_units


class ExamDBScreen(QWidget):
    """기출DB 화면"""
    
    def __init__(self, db_connection: SQLiteConnection, parent=None):
        """
        ExamDBScreen 초기화
        
        Args:
            db_connection: DB 연결 인스턴스
            parent: 부모 위젯
        """
        super().__init__(parent)
        self.db_connection = db_connection
        self.exam_repo = ExamRepository(db_connection)
        self.parsing_service = ParsingService(db_connection)
        self.problem_service = ProblemService(db_connection)
        self.hwp_restore = HWPRestore(db_connection)
        self.current_exam_id = None
        self.current_exam_grade = None
        self.init_ui()
        self.load_exams()
    
    def init_ui(self):
        """UI 초기화"""
        self.setObjectName("ExamDBRoot")
        self._exams_cache: List[Exam] = []

        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(10, 16, 24, 24)  # 좌측 여백 미세 조정, 우측 여백 유지

        # 상단 컨트롤 영역(검색창 단일화)
        control_layout = QHBoxLayout()
        control_layout.setSpacing(10)
        control_layout.setContentsMargins(0, 0, 0, 5)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("연도 또는 시험명을 검색하세요...")
        self.search_input.setClearButtonEnabled(True)
        self.search_input.setFixedHeight(40)
        self.search_input.setMinimumWidth(520)
        self.search_input.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.search_input.textChanged.connect(self._apply_exam_filters)
        control_layout.addWidget(self.search_input)

        control_layout.addStretch(1)

        # 우측 상단 버튼(교재DB 스타일과 동일)
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
        # ✅ 검색창 ↔ 표 사이 간격 최소화(5~10px)
        main_layout.addSpacing(8)

        # ✅ 사용자가 좌/우 패널 폭을 조절할 수 있도록 Splitter 사용
        splitter = QSplitter(Qt.Horizontal)
        splitter.setChildrenCollapsible(False)
        try:
            splitter.setHandleWidth(4)
        except Exception:
            pass

        # 좌측: 기출 테이블(카드)
        table_widget = QFrame()
        table_widget.setObjectName("leftCard")
        table_widget.setMinimumWidth(300)
        table_layout = QVBoxLayout(table_widget)
        table_layout.setContentsMargins(14, 14, 14, 14)
        table_layout.setSpacing(8)

        left_title = QLabel("기출 목록")
        left_title.setObjectName("panelTitle")
        table_layout.addWidget(left_title)

        self.table = QTableWidget()
        self.table.setObjectName("DBTable")
        self.table.setColumnCount(10)
        self.table.setHorizontalHeaderLabels(
            ["출처", "연도", "학년", "학기", "유형", "학교명", "생성일", "문제수", "상태", "⋯"]
        )
        # ✅ 유형(컬럼 4)을 가장 넓게(Stretch)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.ExtendedSelection)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setAlternatingRowColors(False)
        self.table.setShowGrid(True)
        self.table.verticalHeader().setVisible(False)
        self.table.setMouseTracking(True)
        # ✅ 밀도: 행 높이 강제
        try:
            self.table.verticalHeader().setDefaultSectionSize(38)
        except Exception:
            pass

        # ✅ 선택 효과(교재DB 동일): delegate로 배경+좌측 라인
        self.table.setItemDelegate(_RowSelectDelegate(self.table))

        # 컬럼 너비(유형 4는 Stretch)
        self.table.setColumnWidth(0, 110)  # 출처
        self.table.setColumnWidth(1, 90)  # 연도
        self.table.setColumnWidth(2, 90)  # 학년
        self.table.setColumnWidth(3, 90)  # 학기
        self.table.setColumnWidth(5, 190)  # 학교명(고정)
        self.table.setColumnWidth(6, 130)  # 생성일
        self.table.setColumnWidth(7, 90)  # 문제수
        self.table.setColumnWidth(8, 90)  # 상태
        self.table.setColumnWidth(9, 70)  # ⋯

        # 헤더 가운데 정렬(전체 컬럼)
        for col in range(10):
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

        # 우측: Problem 목록(카드)
        problem_widget = QFrame()
        problem_widget.setObjectName("PreviewPanel")
        problem_widget.setMinimumWidth(300)
        problem_layout = QVBoxLayout(problem_widget)
        problem_layout.setContentsMargins(14, 14, 14, 14)
        problem_layout.setSpacing(8)

        problem_header_layout = QHBoxLayout()
        problem_header_layout.setContentsMargins(0, 0, 0, 0)

        problem_label = QLabel("문제 목록")
        problem_label.setObjectName("panelTitle")
        problem_header_layout.addWidget(problem_label)
        problem_header_layout.addStretch()

        self.only_untagged_difficulty_checkbox = QCheckBox("난이도 미지정만")
        self.only_untagged_difficulty_checkbox.setObjectName("onlyUntagged")
        self.only_untagged_difficulty_checkbox.stateChanged.connect(self.on_only_untagged_changed)
        problem_header_layout.addWidget(self.only_untagged_difficulty_checkbox)

        self.only_untagged_unit_checkbox = QCheckBox("단원 미지정만")
        self.only_untagged_unit_checkbox.setObjectName("onlyUntagged")
        self.only_untagged_unit_checkbox.stateChanged.connect(self.on_only_untagged_changed)
        problem_header_layout.addWidget(self.only_untagged_unit_checkbox)

        problem_layout.addLayout(problem_header_layout)

        # 빠른 단원 태깅: spacing(10), 콤보 110px 통일로 한 줄 배치
        unit_apply_layout = QHBoxLayout()
        unit_apply_layout.setContentsMargins(5, 5, 5, 5)
        unit_apply_layout.setSpacing(10)

        l = QLabel("과목")
        l.setObjectName("InlineLabel")
        unit_apply_layout.addWidget(l)
        self.unit_subject_combo = QComboBox()
        self.unit_subject_combo.setMinimumHeight(30)
        self.unit_subject_combo.setFixedWidth(110)
        self.unit_subject_combo.addItem("선택")
        self.unit_subject_combo.addItems(list_subjects())
        self.unit_subject_combo.currentTextChanged.connect(self.on_unit_subject_changed)
        unit_apply_layout.addWidget(self.unit_subject_combo)

        l = QLabel("대단원")
        l.setObjectName("InlineLabel")
        unit_apply_layout.addWidget(l)
        self.major_unit_combo = QComboBox()
        self.major_unit_combo.setMinimumHeight(30)
        self.major_unit_combo.setFixedWidth(110)
        self.major_unit_combo.addItem("선택")
        self.major_unit_combo.currentTextChanged.connect(self.on_unit_major_changed)
        unit_apply_layout.addWidget(self.major_unit_combo)

        l = QLabel("소단원")
        l.setObjectName("InlineLabel")
        unit_apply_layout.addWidget(l)
        self.sub_unit_combo = QComboBox()
        self.sub_unit_combo.setMinimumHeight(30)
        self.sub_unit_combo.setFixedWidth(110)
        self.sub_unit_combo.addItem("(없음)")
        unit_apply_layout.addWidget(self.sub_unit_combo)

        unit_apply_layout.addStretch()
        btn_apply_unit = QPushButton("선택에 적용(Enter)")
        btn_apply_unit.setObjectName("filterApplyBtn")
        btn_apply_unit.setMinimumHeight(30)
        btn_apply_unit.setMinimumWidth(115)
        btn_apply_unit.setCursor(Qt.PointingHandCursor)
        btn_apply_unit.setFocusPolicy(Qt.NoFocus)
        btn_apply_unit.clicked.connect(self.apply_unit_to_selection)
        unit_apply_layout.addWidget(btn_apply_unit)

        problem_layout.addLayout(unit_apply_layout)

        # 초기 콤보 상태
        self.on_unit_subject_changed(self.unit_subject_combo.currentText())
        
        self.problem_table = QTableWidget()
        self.problem_table.setObjectName("ProblemTable")
        self.problem_table.setColumnCount(5)
        self.problem_table.setHorizontalHeaderLabels([
            "#", "미리보기", "난이도", "단원", "원본"
        ])
        # ✅ 폭이 좁아져도 사용자가 컬럼을 조절할 수 있게(기본 Interactive)
        try:
            self.problem_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        except Exception:
            pass
        self.problem_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)  # 미리보기 자동 확장
        self.problem_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.problem_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.problem_table.setAlternatingRowColors(False)
        self.problem_table.setShowGrid(True)
        self.problem_table.verticalHeader().setVisible(False)
        # ✅ 밀도: 행 높이 강제
        try:
            self.problem_table.verticalHeader().setDefaultSectionSize(38)
        except Exception:
            pass
        self.problem_table.setColumnWidth(0, 50)   # #
        self.problem_table.setColumnWidth(2, 62)   # 난이도(55px 위젯 + 여유)
        self.problem_table.setColumnWidth(3, 220)  # 단원(가독성)
        self.problem_table.setColumnWidth(4, 80)  # 원본
        self.problem_table.setWordWrap(False)

        # 헤더 정렬(미리보기 컬럼 제외)
        for col in range(5):
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
        
        # ✅ 교재DB와 동일한 비율(좌 60% : 우 40%)
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 2)
        splitter.setSizes([600, 400])
        
        main_layout.addWidget(splitter)

        # ✅ 교재DB와 동일한 QSS 규칙(색/대비/테이블/간격)
        self.setStyleSheet(
            """
            QWidget#ExamDBRoot {
                background: transparent;
                font-family: 'Pretendard','Malgun Gothic','맑은 고딕';
            }
            QWidget#ExamDBRoot * {
                outline: none;
            }

            QLabel#panelTitle {
                color: #222222;
                font-size: 12pt;
                font-weight: 800;
            }
            QLabel#InlineLabel {
                color: #222222;
                font-weight: 700;
                background: transparent;
            }

            QPushButton#primary {
                background-color: #2563EB;
                color: #FFFFFF;
                border: 1px solid #2563EB;
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
                padding: 8px 14px;
                min-height: 22px;
                font-weight: 800;
            }
            QPushButton#secondary:hover { background-color: #EFF6FF; }
            QPushButton#secondary:disabled { color: #94A3B8; border-color: #CBD5E1; }
            QPushButton#filterApplyBtn {
                background-color: #FFFFFF;
                color: #2563EB;
                border: 1px solid #2563EB;
                border-radius: 8px;
                padding: 4px 10px;
                min-height: 22px;
                font-weight: 700;
            }
            QPushButton#filterApplyBtn:hover { background-color: #EFF6FF; }

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

            QComboBox {
                background-color: #FFFFFF;
                border: 1.5px solid #CBD5E1;
                border-radius: 10px;
                padding: 6px 10px;
                color: #222222;
                font-weight: 600;
            }
            QComboBox:hover { background-color: #F8FAFC; }
            QComboBox::drop-down { border: none; }
            QComboBox QAbstractItemView {
                background-color: #FFFFFF;
                border: 1px solid #E0E0E0;
                selection-background-color: #E2E8F0;
                selection-color: #222222;
                outline: 0;
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
                background-color: #ffffff;
                border: 1px solid #dcdcdc;
                border-radius: 4px;
                padding: 0px;
                min-width: 55px;
                font-size: 10pt;
                color: #000000;
                font-weight: bold;
            }
            QComboBox#DifficultyCombo::drop-down { border: none; }
            QComboBox#DifficultyCombo QAbstractItemView {
                background-color: #FFFFFF;
                border: 1px solid #E0E0E0;
                selection-background-color: #F0F7FF;
                selection-color: #000000;
                outline: 0;
            }
            """
        )

        # 헤더 높이(밀도 최적화)
        try:
            self.table.horizontalHeader().setFixedHeight(32)
            self.problem_table.horizontalHeader().setFixedHeight(32)
        except Exception:
            pass


    def load_exams(self):
        """기출 목록 로드"""
        try:
            self._exams_cache = self.exam_repo.list_all()
            self._apply_exam_filters()
        
        except ConnectionError as e:
            QMessageBox.warning(self, "연결 오류", f"DB에 연결할 수 없습니다.\n\n{str(e)}")
        except Exception as e:
            QMessageBox.critical(self, "오류", f"기출 목록을 불러올 수 없습니다.\n\n{str(e)}")

    def _apply_exam_filters(self):
        """상단 검색창 기준으로 기출 목록 필터링(데이터 구조 변경 없음)"""
        exams = list(getattr(self, "_exams_cache", []) or [])
        query = (self.search_input.text() if getattr(self, "search_input", None) else "") or ""
        q = query.strip().lower()

        if q:
            filtered: List[Exam] = []
            for ex in exams:
                hay = f"{ex.year or ''} {ex.grade or ''} {ex.semester or ''} {ex.exam_type or ''} {ex.school_name or ''}".lower()
                if q in hay:
                    filtered.append(ex)
            exams = filtered

        self.table.setRowCount(len(exams))

        for row, exam in enumerate(exams):
            # 출처
            item = QTableWidgetItem("내신기출")
            item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, 0, item)

            # 연도
            item = QTableWidgetItem(exam.year or "")
            item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, 1, item)

            # 학년
            item = QTableWidgetItem(exam.grade or "")
            item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, 2, item)

            # 학기
            item = QTableWidgetItem(exam.semester or "")
            item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, 3, item)

            # 유형(중간고사, 기말고사 등) — 컬럼 인덱스 4, 글자색·중앙 정렬 명시
            type_text = getattr(exam, "exam_type", None) or ""
            item = QTableWidgetItem(str(type_text) if type_text else "")
            item.setTextAlignment(Qt.AlignCenter)
            item.setForeground(QBrush(QColor("#222222")))
            if type_text:
                item.setToolTip(str(type_text))
            self.table.setItem(row, 4, item)

            # 학교명(좌측 정렬) + ID 저장(기존 구조 유지: col 5 item에 UserRole)
            item = QTableWidgetItem(exam.school_name or "")
            item.setTextAlignment(Qt.AlignVCenter | Qt.AlignLeft)
            self.table.setItem(row, 5, item)
            item.setData(Qt.UserRole, exam.id)

            # 생성일
            date_str = exam.created_at.strftime("%Y.%m.%d") if getattr(exam, "created_at", None) else ""
            item = QTableWidgetItem(date_str)
            item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, 6, item)

            # 문제수
            item = QTableWidgetItem(str(exam.problem_count))
            item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, 7, item)

            # 상태(배지)
            if exam.is_parsed:
                status = "완료" if (exam.problem_count or 0) > 0 else "부분"
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
            self.table.setCellWidget(row, 8, badge_wrap)

            # 더보기 버튼
            btn_more = QPushButton("⋯")
            btn_more.setMinimumWidth(40)
            btn_more.setMinimumHeight(28)
            btn_more.setFont(QFont("맑은 고딕", 14, QFont.Bold))
            btn_more.setStyleSheet(
                """
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
                """
            )
            btn_more.clicked.connect(lambda checked, eid=exam.id: self.on_more_clicked(eid))
            self.table.setCellWidget(row, 9, btn_more)

            # ✅ 행 높이(밀도 최적화)
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
        dialog = ExamMetadataDialog(self)
        
        def on_metadata_submitted(metadata):
            dialog.close()
            self.process_exam_creation(file_path, metadata)
        
        dialog.submitted.connect(on_metadata_submitted)
        dialog.exec_()
    
    def process_exam_creation(self, hwp_path: str, metadata: dict):
        """기출 생성 및 파싱 처리"""
        try:
            # 3. Exam 생성
            exam = Exam(
                year=metadata['year'],
                grade=metadata['grade'],
                semester=metadata['semester'],
                exam_type=metadata['exam_type'],
                school_name=metadata['school_name']
            )
            
            exam_id = self.exam_repo.create(exam)
            if not exam_id:
                QMessageBox.critical(self, "오류", "기출을 생성할 수 없습니다.")
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
            result = self.parsing_service.reparse_exam(
                exam_id=exam_id,
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
                self.load_exams()
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
            self.current_exam_id = None
            self.current_exam_grade = None
            try:
                if getattr(self, "btn_generate_preview", None):
                    self.btn_generate_preview.setEnabled(False)
            except Exception:
                pass
            return
        
        row = selected_rows[0].row()
        exam_id_item = self.table.item(row, 5)  # 학교명 셀에 ID 저장
        if not exam_id_item:
            return
        
        exam_id = exam_id_item.data(Qt.UserRole)
        if not exam_id:
            return
        
        self.current_exam_id = exam_id
        # 선택된 exam의 학년은 단원 태그 저장 시 grade로 함께 저장(빠른 실무 태깅)
        grade_item = self.table.item(row, 2)
        self.current_exam_grade = grade_item.text() if grade_item else None
        try:
            if getattr(self, "btn_generate_preview", None):
                self.btn_generate_preview.setEnabled(True)
        except Exception:
            pass

        self.load_problems(exam_id)
    
    def load_problems(self, exam_id: str):
        """Problem 목록 로드"""
        try:
            problems = self.problem_service.get_problems_by_source(
                source_id=exam_id,
                source_type=SourceType.EXAM
            )

            # 필터: 난이도 미지정만
            if getattr(self, "only_untagged_difficulty_checkbox", None) and self.only_untagged_difficulty_checkbox.isChecked():
                problems = [p for p in problems if not p.get("difficulty")]
            # 필터: 단원 미지정만 (대단원이 비어있으면 미지정으로 판단)
            if getattr(self, "only_untagged_unit_checkbox", None) and self.only_untagged_unit_checkbox.isChecked():
                problems = [p for p in problems if not p.get("major_unit")]
            
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
                if preview_text:
                    item.setToolTip(preview_text)
                self.problem_table.setItem(row, 1, item)
                
                # 난이도(드롭다운): 최소 55px 너비로 글자 가시성 확보, 중앙 정렬
                difficulty = problem.get("difficulty")
                combo = QComboBox()
                combo.setObjectName("DifficultyCombo")
                combo.addItems(["미지정", "하", "중", "상", "킬"])
                combo.setFont(QFont("맑은 고딕", 10))
                combo.setFixedWidth(55)
                combo.setMinimumHeight(26)
                try:
                    combo.view().setMinimumWidth(80)
                except Exception:
                    pass

                current_text = difficulty if difficulty else "미지정"
                if current_text in ["하", "중", "상", "킬"]:
                    combo.setCurrentText(current_text)
                else:
                    combo.setCurrentText("미지정")

                problem_id = problem.get("problem_id")
                combo.currentTextChanged.connect(lambda val, pid=problem_id: self.on_difficulty_changed(pid, val))
                container = QWidget()
                container.setStyleSheet("QWidget { background-color: transparent; border: none; }")
                w_layout = QHBoxLayout(container)
                w_layout.setContentsMargins(0, 0, 0, 0)
                w_layout.setSpacing(0)
                w_layout.setAlignment(Qt.AlignCenter)
                w_layout.addWidget(combo)
                self.problem_table.setCellWidget(row, 2, container)
                self._problem_cell_widgets[row] = {2: container}

                # 단원 표시
                unit_display = problem.get("unit_display") or ""
                unit_text = unit_display if unit_display else "미지정"
                item = QTableWidgetItem(str(unit_text))
                item.setTextAlignment(Qt.AlignCenter)
                if unit_display:
                    item.setToolTip(str(unit_display))
                self.problem_table.setItem(row, 3, item)
                
                # 원본
                has_raw = problem.get('has_content_raw', False)
                raw_text = "✓" if has_raw else "✗"
                item = QTableWidgetItem(str(raw_text))
                item.setTextAlignment(Qt.AlignCenter)
                self.problem_table.setItem(row, 4, item)

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

            # 미지정만 필터가 켜져 있으면, 태깅 후 목록 갱신
            if getattr(self, "only_untagged_difficulty_checkbox", None) and self.only_untagged_difficulty_checkbox.isChecked():
                if self.current_exam_id:
                    self.load_problems(self.current_exam_id)
        except ConnectionError as e:
            QMessageBox.warning(self, "연결 오류", f"DB에 연결할 수 없습니다.\n\n{str(e)}")
        except Exception as e:
            QMessageBox.critical(self, "오류", f"난이도를 저장할 수 없습니다.\n\n{str(e)}")

    def on_only_untagged_changed(self):
        """미지정만 보기 토글(난이도/단원)"""
        if self.current_exam_id:
            self.load_problems(self.current_exam_id)

    def on_generate_previews_clicked(self):
        """현재 기출의 미리보기 텍스트를 일괄 생성"""
        if not self.current_exam_id:
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
                source_id=self.current_exam_id,
                source_type=SourceType.EXAM,
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
            self.load_problems(self.current_exam_id)
        except ConnectionError as e:
            QMessageBox.warning(self, "연결 오류", f"DB에 연결할 수 없습니다.\n\n{str(e)}")
        except Exception as e:
            QMessageBox.critical(self, "오류", f"미리보기를 생성할 수 없습니다.\n\n{str(e)}")

    def on_unit_subject_changed(self, subject: str):
        """과목 변경 → 대단원/소단원 갱신"""
        self.major_unit_combo.blockSignals(True)
        self.major_unit_combo.clear()
        self.major_unit_combo.addItem("선택")
        if subject and subject != "선택":
            self.major_unit_combo.addItems(list_major_units(subject))
        self.major_unit_combo.blockSignals(False)

        self.sub_unit_combo.clear()
        self.sub_unit_combo.addItem("(없음)")

    def on_unit_major_changed(self, major: str):
        """대단원 변경 → 소단원 갱신"""
        subject = self.unit_subject_combo.currentText()
        self.sub_unit_combo.clear()
        self.sub_unit_combo.addItem("(없음)")
        if subject and subject != "선택" and major and major != "선택":
            self.sub_unit_combo.addItems(list_sub_units(subject, major))

    def apply_unit_to_selection(self):
        """현재 입력된 과목/대단원/소단원을 선택된 문제들에 일괄 적용"""
        if not self.current_exam_id:
            return

        subject = (self.unit_subject_combo.currentText() or "").strip()
        major = (self.major_unit_combo.currentText() or "").strip()
        sub_raw = (self.sub_unit_combo.currentText() or "").strip()
        sub = None if (not sub_raw or sub_raw == "(없음)") else sub_raw

        if subject == "선택":
            QMessageBox.warning(self, "입력 필요", "과목은 필수입니다.")
            return
        if major == "선택":
            QMessageBox.warning(self, "입력 필요", "대단원은 필수입니다.")
            return

        selected = self.problem_table.selectionModel().selectedRows()
        if not selected:
            return

        try:
            for idx in selected:
                row = idx.row()
                problem_id_item = self.problem_table.item(row, 0)
                if not problem_id_item:
                    continue
                problem_id = problem_id_item.data(Qt.UserRole)
                if not problem_id:
                    continue
                self.problem_service.set_problem_unit(
                    problem_id=problem_id,
                    subject=subject,
                    major_unit=major,
                    sub_unit=sub,
                    grade=self.current_exam_grade
                )
                # 화면 즉시 반영
                unit_text = f"{major} > {sub}" if sub else major
                unit_item = self.problem_table.item(row, 3)
                if unit_item:
                    unit_item.setText(unit_text)
                else:
                    unit_item = QTableWidgetItem(unit_text)
                    unit_item.setTextAlignment(Qt.AlignCenter)
                    self.problem_table.setItem(row, 3, unit_item)

            # 단원 미지정만 필터가 켜져 있으면, 태깅 후 목록 갱신(방금 태깅한 문제들 제거)
            if getattr(self, "only_untagged_unit_checkbox", None) and self.only_untagged_unit_checkbox.isChecked():
                self.load_problems(self.current_exam_id)
                if self.problem_table.rowCount() > 0:
                    self.problem_table.selectRow(0)
                return

            # 다음 행 이동(빠른 전수 태깅)
            current_row = selected[0].row()
            if not self._select_next_untagged_unit(current_row):
                next_row = min(current_row + 1, self.problem_table.rowCount() - 1)
                if next_row != current_row:
                    self.problem_table.selectRow(next_row)
        except ConnectionError as e:
            QMessageBox.warning(self, "연결 오류", f"DB에 연결할 수 없습니다.\n\n{str(e)}")
        except Exception as e:
            QMessageBox.critical(self, "오류", f"단원을 저장할 수 없습니다.\n\n{str(e)}")

    def _select_next_untagged_unit(self, from_row: int) -> bool:
        """현재 행 다음부터 '단원 미지정'인 문제를 찾아 이동"""
        for r in range(from_row + 1, self.problem_table.rowCount()):
            item = self.problem_table.item(r, 3)
            if item and item.text() == "미지정":
                self.problem_table.selectRow(r)
                return True
        return False

    def eventFilter(self, obj, event):
        """단축키: 난이도(0~4), 단원 적용(Enter)"""
        if obj is self.problem_table and event.type() == QEvent.KeyPress:
            key = event.key()
            # Enter/Return: 현재 단원을 선택 문제에 적용
            if key in (Qt.Key_Return, Qt.Key_Enter):
                self.apply_unit_to_selection()
                return True

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
                    combo = None
                    if isinstance(widget, QComboBox):
                        combo = widget
                    elif isinstance(widget, QWidget):
                        try:
                            combo = widget.findChild(QComboBox)
                        except Exception:
                            combo = None
                    if combo:
                        combo.setCurrentText(mapping[key])
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
    
    def on_more_clicked(self, exam_id: str):
        """더보기 버튼 클릭 처리"""
        from PyQt5.QtWidgets import QMenu
        
        selected_ids = self._get_selected_exam_ids()
        target_ids = selected_ids if (exam_id in selected_ids and len(selected_ids) > 1) else [exam_id]

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
            self.on_edit_exam(target_ids[0])
        elif action == action_delete:
            self.on_delete_exams(target_ids)
        elif action == action_reparse:
            self.on_reparse_exam(target_ids[0])
    
    def on_edit_exam(self, exam_id: str):
        """기출 수정"""
        try:
            exam = self.exam_repo.find_by_id(exam_id)
            if not exam:
                QMessageBox.warning(self, "오류", "기출 정보를 찾을 수 없습니다.")
                return

            dialog = ExamMetadataDialog(self)
            dialog.setWindowTitle("기출 수정")

            # 기존 값 채우기
            try:
                dialog.year_input.setText(exam.year or "")
                dialog.grade_input.setText(exam.grade or "")
                dialog.semester_input.setText(exam.semester or "")
                dialog.exam_type_input.setText(exam.exam_type or "")
                dialog.school_name_input.setText(exam.school_name or "")
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
            new_year = (md.get("year") or "").strip()
            new_grade = (md.get("grade") or "").strip()
            new_semester = (md.get("semester") or "").strip()
            new_exam_type = (md.get("exam_type") or "").strip()
            new_school = (md.get("school_name") or "").strip()

            changed = (
                new_year != (exam.year or "")
                or new_grade != (exam.grade or "")
                or new_semester != (exam.semester or "")
                or new_exam_type != (exam.exam_type or "")
                or new_school != (exam.school_name or "")
            )

            exam.year = new_year
            exam.grade = new_grade
            exam.semester = new_semester
            exam.exam_type = new_exam_type
            exam.school_name = new_school

            updated = self.exam_repo.update(exam)
            if (not updated) and (not changed):
                QMessageBox.information(self, "완료", "변경 사항이 없습니다.")
            elif updated or changed:
                QMessageBox.information(self, "완료", "기출이 수정되었습니다.")
            else:
                QMessageBox.warning(self, "오류", "기출을 수정할 수 없습니다.")

            self.load_exams()
            self._select_exam_row_by_id(exam_id)
        except ConnectionError as e:
            QMessageBox.warning(self, "연결 오류", f"DB에 연결할 수 없습니다.\n\n{str(e)}")
        except Exception as e:
            QMessageBox.critical(self, "오류", f"수정 중 오류가 발생했습니다.\n\n{str(e)}")
    
    def on_delete_exam(self, exam_id: str):
        """기출 삭제"""
        self.on_delete_exams([exam_id])

    def on_delete_exams(self, exam_ids: List[str]):
        """기출 여러 개 삭제 (연결된 문제도 함께 삭제)"""
        ids = [x for x in (exam_ids or []) if x]
        if not ids:
            return

        msg = (
            f"선택한 기출 {len(ids)}개를 삭제하시겠습니까?\n"
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

            for eid in ids:
                deleted_problems += self.problem_service.delete_problems_by_source(
                    eid, SourceType.EXAM
                )
                if self.exam_repo.delete(eid):
                    deleted_sources += 1
                else:
                    failed.append(eid)

            self.load_exams()
            if self.current_exam_id in ids:
                self.current_exam_id = None
                self.current_exam_grade = None
                self.problem_table.setRowCount(0)

            if failed:
                QMessageBox.warning(
                    self,
                    "부분 완료",
                    f"기출 {deleted_sources}개 삭제, 연결 문제 {deleted_problems}개 삭제 완료.\n\n"
                    f"삭제 실패: {len(failed)}개",
                )
            else:
                QMessageBox.information(
                    self,
                    "완료",
                    f"기출 {deleted_sources}개 삭제, 연결 문제 {deleted_problems}개 삭제 완료.",
                )
        except ConnectionError as e:
            QMessageBox.warning(self, "연결 오류", f"DB에 연결할 수 없습니다.\n\n{str(e)}")
        except Exception as e:
            QMessageBox.critical(self, "오류", f"삭제 중 오류가 발생했습니다.\n\n{str(e)}")
    
    def on_reparse_exam(self, exam_id: str):
        """기출 재파싱"""
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
            result = self.parsing_service.reparse_exam(
                exam_id=exam_id,
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
                self.load_exams()
                if self.current_exam_id == exam_id:
                    self.load_problems(exam_id)
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

        exam_id_item = self.table.item(row, 5)
        if not exam_id_item:
            return
        
        exam_id = exam_id_item.data(Qt.UserRole)
        if not exam_id:
            return

        selected_ids = self._get_selected_exam_ids()
        target_ids = selected_ids if (exam_id in selected_ids and len(selected_ids) > 1) else [exam_id]
        
        from PyQt5.QtWidgets import QMenu
        
        menu = QMenu(self)
        action_edit = menu.addAction("수정")
        action_delete = menu.addAction("삭제")
        action_reparse = menu.addAction("재파싱")

        if len(target_ids) != 1:
            action_edit.setEnabled(False)
            action_reparse.setEnabled(False)
        
        action = menu.exec_(self.table.viewport().mapToGlobal(position))
        
        if action == action_edit:
            self.on_edit_exam(target_ids[0])
        elif action == action_delete:
            self.on_delete_exams(target_ids)
        elif action == action_reparse:
            self.on_reparse_exam(target_ids[0])

    def _get_selected_exam_ids(self) -> List[str]:
        """기출 테이블에서 선택된 행들의 exam_id 리스트"""
        ids: List[str] = []
        try:
            selected = self.table.selectionModel().selectedRows()
            for idx in selected:
                r = idx.row()
                item = self.table.item(r, 5)
                if item:
                    eid = item.data(Qt.UserRole)
                    if eid:
                        ids.append(str(eid))
        except Exception:
            return []
        uniq: List[str] = []
        seen = set()
        for x in ids:
            if x not in seen:
                seen.add(x)
                uniq.append(x)
        return uniq

    def _select_exam_row_by_id(self, exam_id: str) -> None:
        """목록 새로고침 후 특정 기출 행을 다시 선택"""
        if not exam_id:
            return
        try:
            for r in range(self.table.rowCount()):
                item = self.table.item(r, 5)
                if item and item.data(Qt.UserRole) == exam_id:
                    self.table.selectRow(r)
                    break
        except Exception:
            pass


class _RowSelectDelegate(QStyledItemDelegate):
    """선택 행 강조: 배경 + 좌측 블루 포인트 라인."""

    def paint(self, painter: QPainter, option, index):  # type: ignore[override]
        selected = bool(option.state & QStyle.State_Selected)
        # hover 하이라이트는 "배경 노이즈"가 될 수 있어 사용하지 않음

        opt = option
        if selected:
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
