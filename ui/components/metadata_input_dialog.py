"""
메타데이터 입력 다이얼로그

교재/기출 DB 생성 시 메타데이터를 입력받는 다이얼로그
"""
from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel,
                             QLineEdit, QPushButton, QFormLayout, QComboBox)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QFont
from core.unit_catalog import list_subjects, list_major_units, list_sub_units


class TextbookMetadataDialog(QDialog):
    """교재 메타데이터 입력 다이얼로그"""
    
    # 시그널 정의
    submitted = pyqtSignal(dict)  # 메타데이터 제출 (name, subject, major_unit, sub_unit)
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("교재 메타데이터 입력")
        self.setMinimumWidth(400)
        self.init_ui()
    
    def init_ui(self):
        """UI 초기화"""
        layout = QVBoxLayout()
        layout.setSpacing(16)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # 폼 레이아웃
        form_layout = QFormLayout()
        form_layout.setSpacing(12)
        
        # 교재명
        self.name_input = QLineEdit()
        self.name_input.setMinimumHeight(36)
        self.name_input.setFont(QFont("맑은 고딕", 11))
        self.name_input.setPlaceholderText("예: 쎈수학")
        form_layout.addRow("교재명 *:", self.name_input)
        
        # 과목
        self.subject_combo = QComboBox()
        self.subject_combo.setMinimumHeight(36)
        self.subject_combo.setFont(QFont("맑은 고딕", 11))
        self.subject_combo.addItem("선택")
        self.subject_combo.addItems(list_subjects())
        self.subject_combo.currentTextChanged.connect(self.on_subject_changed)
        form_layout.addRow("과목 *:", self.subject_combo)
        
        # 대단원
        self.major_unit_combo = QComboBox()
        self.major_unit_combo.setMinimumHeight(36)
        self.major_unit_combo.setFont(QFont("맑은 고딕", 11))
        self.major_unit_combo.addItem("선택")
        self.major_unit_combo.currentTextChanged.connect(self.on_major_unit_changed)
        form_layout.addRow("대단원 *:", self.major_unit_combo)
        
        # 소단원
        self.sub_unit_combo = QComboBox()
        self.sub_unit_combo.setMinimumHeight(36)
        self.sub_unit_combo.setFont(QFont("맑은 고딕", 11))
        self.sub_unit_combo.addItem("(없음)")
        form_layout.addRow("소단원:", self.sub_unit_combo)
        
        layout.addLayout(form_layout)
        
        # 버튼 영역
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        btn_cancel = QPushButton("취소")
        btn_cancel.setMinimumWidth(80)
        btn_cancel.setMinimumHeight(36)
        btn_cancel.setFont(QFont("맑은 고딕", 11))
        btn_cancel.clicked.connect(self.reject)
        button_layout.addWidget(btn_cancel)
        
        btn_submit = QPushButton("확인")
        btn_submit.setMinimumWidth(80)
        btn_submit.setMinimumHeight(36)
        btn_submit.setFont(QFont("맑은 고딕", 11, QFont.Bold))
        btn_submit.setStyleSheet("""
            QPushButton {
                background-color: #1967D2;
                color: #FFFFFF;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #1557B0;
            }
        """)
        btn_submit.clicked.connect(self.on_submit)
        button_layout.addWidget(btn_submit)
        
        layout.addLayout(button_layout)
        self.setLayout(layout)
        
        # 스타일 설정
        self.setStyleSheet("""
            QDialog {
                background-color: #FFFFFF;
            }
            QLineEdit {
                border: 1px solid #DADCE0;
                border-radius: 4px;
                padding: 8px 12px;
            }
            QLineEdit:focus {
                border-color: #1967D2;
            }
            QComboBox {
                border: 1px solid #DADCE0;
                border-radius: 4px;
                padding: 6px 12px;
                background-color: #FFFFFF;
            }
            QComboBox:focus {
                border-color: #1967D2;
            }
            QPushButton {
                border: 1px solid #DADCE0;
                border-radius: 4px;
                background-color: #FFFFFF;
                color: #3C4043;
                padding: 8px 16px;
            }
            QPushButton:hover {
                background-color: #F8F9FA;
            }
        """)

        # 초기 상태
        self.on_subject_changed(self.subject_combo.currentText())
    
    def on_subject_changed(self, subject: str):
        """과목 변경 → 대단원/소단원 갱신"""
        self.major_unit_combo.blockSignals(True)
        self.major_unit_combo.clear()
        self.major_unit_combo.addItem("선택")
        if subject and subject != "선택":
            self.major_unit_combo.addItems(list_major_units(subject))
        self.major_unit_combo.blockSignals(False)

        self.sub_unit_combo.clear()
        self.sub_unit_combo.addItem("(없음)")

    def on_major_unit_changed(self, major_unit: str):
        """대단원 변경 → 소단원 갱신"""
        subject = self.subject_combo.currentText()
        self.sub_unit_combo.clear()
        self.sub_unit_combo.addItem("(없음)")
        if subject and subject != "선택" and major_unit and major_unit != "선택":
            self.sub_unit_combo.addItems(list_sub_units(subject, major_unit))

    def on_submit(self):
        """제출 버튼 클릭 처리"""
        name = self.name_input.text().strip()
        subject = self.subject_combo.currentText().strip()
        major_unit = self.major_unit_combo.currentText().strip()
        sub_unit_raw = self.sub_unit_combo.currentText().strip()
        sub_unit = None if (not sub_unit_raw or sub_unit_raw == "(없음)") else sub_unit_raw
        
        # 필수 필드 검증
        if not name or subject == "선택" or major_unit == "선택":
            from PyQt5.QtWidgets import QMessageBox
            QMessageBox.warning(self, "입력 오류", "교재명, 과목, 대단원은 필수 입력 항목입니다.")
            return
        
        # 메타데이터 전달
        metadata = {
            'name': name,
            'subject': subject,
            'major_unit': major_unit,
            'sub_unit': sub_unit
        }
        self.submitted.emit(metadata)
        self.accept()


class ExamMetadataDialog(QDialog):
    """기출 메타데이터 입력 다이얼로그"""
    
    # 시그널 정의
    submitted = pyqtSignal(dict)  # 메타데이터 제출 (grade, semester, exam_type, school_name, year)
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("기출 메타데이터 입력")
        self.setMinimumWidth(400)
        self.init_ui()
    
    def init_ui(self):
        """UI 초기화"""
        layout = QVBoxLayout()
        layout.setSpacing(16)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # 폼 레이아웃
        form_layout = QFormLayout()
        form_layout.setSpacing(12)
        
        # 연도
        self.year_input = QLineEdit()
        self.year_input.setMinimumHeight(36)
        self.year_input.setFont(QFont("맑은 고딕", 11))
        self.year_input.setPlaceholderText("예: 2025")
        form_layout.addRow("연도 *:", self.year_input)
        
        # 학년
        self.grade_input = QLineEdit()
        self.grade_input.setMinimumHeight(36)
        self.grade_input.setFont(QFont("맑은 고딕", 11))
        self.grade_input.setPlaceholderText("예: 1학년")
        form_layout.addRow("학년 *:", self.grade_input)
        
        # 학기
        self.semester_input = QLineEdit()
        self.semester_input.setMinimumHeight(36)
        self.semester_input.setFont(QFont("맑은 고딕", 11))
        self.semester_input.setPlaceholderText("예: 1학기")
        form_layout.addRow("학기 *:", self.semester_input)
        
        # 유형
        self.exam_type_input = QLineEdit()
        self.exam_type_input.setMinimumHeight(36)
        self.exam_type_input.setFont(QFont("맑은 고딕", 11))
        self.exam_type_input.setPlaceholderText("예: 중간고사")
        form_layout.addRow("유형 *:", self.exam_type_input)
        
        # 학교명
        self.school_name_input = QLineEdit()
        self.school_name_input.setMinimumHeight(36)
        self.school_name_input.setFont(QFont("맑은 고딕", 11))
        self.school_name_input.setPlaceholderText("예: OO고등학교")
        form_layout.addRow("학교명 *:", self.school_name_input)
        
        layout.addLayout(form_layout)
        
        # 버튼 영역
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        btn_cancel = QPushButton("취소")
        btn_cancel.setMinimumWidth(80)
        btn_cancel.setMinimumHeight(36)
        btn_cancel.setFont(QFont("맑은 고딕", 11))
        btn_cancel.clicked.connect(self.reject)
        button_layout.addWidget(btn_cancel)
        
        btn_submit = QPushButton("확인")
        btn_submit.setMinimumWidth(80)
        btn_submit.setMinimumHeight(36)
        btn_submit.setFont(QFont("맑은 고딕", 11, QFont.Bold))
        btn_submit.setStyleSheet("""
            QPushButton {
                background-color: #1967D2;
                color: #FFFFFF;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #1557B0;
            }
        """)
        btn_submit.clicked.connect(self.on_submit)
        button_layout.addWidget(btn_submit)
        
        layout.addLayout(button_layout)
        self.setLayout(layout)
        
        # 스타일 설정
        self.setStyleSheet("""
            QDialog {
                background-color: #FFFFFF;
            }
            QLineEdit {
                border: 1px solid #DADCE0;
                border-radius: 4px;
                padding: 8px 12px;
            }
            QLineEdit:focus {
                border-color: #1967D2;
            }
            QPushButton {
                border: 1px solid #DADCE0;
                border-radius: 4px;
                background-color: #FFFFFF;
                color: #3C4043;
                padding: 8px 16px;
            }
            QPushButton:hover {
                background-color: #F8F9FA;
            }
        """)
    
    def on_submit(self):
        """제출 버튼 클릭 처리"""
        year = self.year_input.text().strip()
        grade = self.grade_input.text().strip()
        semester = self.semester_input.text().strip()
        exam_type = self.exam_type_input.text().strip()
        school_name = self.school_name_input.text().strip()
        
        # 필수 필드 검증
        if not year or not grade or not semester or not exam_type or not school_name:
            from PyQt5.QtWidgets import QMessageBox
            QMessageBox.warning(self, "입력 오류", "모든 항목은 필수 입력입니다.")
            return
        
        # 메타데이터 전달
        metadata = {
            'year': year,
            'grade': grade,
            'semester': semester,
            'exam_type': exam_type,
            'school_name': school_name
        }
        self.submitted.emit(metadata)
        self.accept()
