"""
Problem 상세 보기 다이얼로그

Problem 상세 정보를 표시하고 원본 HWP를 볼 수 있는 다이얼로그
"""
import os
from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel,
                             QPushButton, QTextEdit, QScrollArea, QWidget)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont


class ProblemDetailDialog(QDialog):
    """Problem 상세 보기 다이얼로그"""
    
    def __init__(self, problem_data: dict, parent=None):
        """
        ProblemDetailDialog 초기화
        
        Args:
            problem_data: ProblemService.get_problem_detail() 반환 데이터
            parent: 부모 위젯
        """
        super().__init__(parent)
        self.problem_data = problem_data
        self.hwp_restore = None  # HWPRestore 인스턴스 (외부에서 주입)
        # problem_index가 리스트일 수 있으므로 문자열로 변환
        problem_index = problem_data.get('problem_index', '?')
        if isinstance(problem_index, list):
            problem_index = problem_index[0] if problem_index else '?'
        self.setWindowTitle(f"문제 상세 - #{problem_index}")
        self.setMinimumSize(700, 600)
        self.init_ui()
    
    def set_hwp_restore(self, hwp_restore):
        """HWPRestore 인스턴스 설정"""
        self.hwp_restore = hwp_restore
    
    def init_ui(self):
        """UI 초기화"""
        layout = QVBoxLayout()
        layout.setSpacing(16)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # 문제 번호 및 메타데이터
        meta_layout = QHBoxLayout()
        meta_layout.setSpacing(12)
        
        # problem_index가 리스트일 수 있으므로 문자열로 변환
        problem_index = self.problem_data.get('problem_index', '?')
        if isinstance(problem_index, list):
            problem_index = problem_index[0] if problem_index else '?'
        index_label = QLabel(f"문제 #{problem_index}")
        index_label.setFont(QFont("맑은 고딕", 12, QFont.Bold))
        meta_layout.addWidget(index_label)
        
        meta_layout.addStretch()
        
        # 원본 보존 여부 표시
        if self.problem_data.get('has_content_raw'):
            raw_status = QLabel("✓ 원본 HWP 보존됨")
            raw_status.setFont(QFont("맑은 고딕", 10))
            raw_status.setStyleSheet("color: #2E7D32;")
            meta_layout.addWidget(raw_status)
        else:
            raw_status = QLabel("✗ 원본 HWP 없음")
            raw_status.setFont(QFont("맑은 고딕", 10))
            raw_status.setStyleSheet("color: #D32F2F;")
            meta_layout.addWidget(raw_status)
        
        layout.addLayout(meta_layout)
        
        # 텍스트 미리보기 영역
        preview_label = QLabel("텍스트 미리보기")
        preview_label.setFont(QFont("맑은 고딕", 11, QFont.Bold))
        layout.addWidget(preview_label)
        
        self.text_preview = QTextEdit()
        self.text_preview.setReadOnly(True)
        self.text_preview.setMinimumHeight(200)
        self.text_preview.setFont(QFont("맑은 고딕", 10))
        # content_text가 리스트일 수 있으므로 문자열로 변환
        content_text = self.problem_data.get('content_text', '')
        if isinstance(content_text, list):
            content_text = ' '.join(str(x) for x in content_text) if content_text else ''
        elif content_text is None:
            content_text = ''
        self.text_preview.setText(str(content_text))
        self.text_preview.setStyleSheet("""
            QTextEdit {
                border: 1px solid #DADCE0;
                border-radius: 4px;
                padding: 12px;
                background-color: #FAFAFA;
            }
        """)
        layout.addWidget(self.text_preview)
        
        # 태그 정보
        tags = self.problem_data.get('tags')
        if tags and len(tags) > 0:
            tags_label = QLabel("태그 정보")
            tags_label.setFont(QFont("맑은 고딕", 11, QFont.Bold))
            layout.addWidget(tags_label)
            
            tags_widget = QWidget()
            tags_layout = QHBoxLayout()
            tags_layout.setSpacing(8)
            tags_layout.setContentsMargins(0, 0, 0, 0)
            
            for tag in tags:
                subject = tag.get('subject', '') or ''
                grade = tag.get('grade', '') or ''
                major_unit = tag.get('major_unit')
                sub_unit = tag.get('sub_unit')
                unit = tag.get('unit')
                difficulty = tag.get('difficulty')

                # 난이도만 있는 태그(예: subject/grade 비어있음)도 보기 좋게 표시
                if subject or grade:
                    tag_text = f"{subject} / {grade}"
                else:
                    tag_text = "난이도"

                # 단원 표시: major/sub 우선, 없으면 레거시 unit
                if major_unit and sub_unit:
                    tag_text += f" / {major_unit} > {sub_unit}"
                elif major_unit:
                    tag_text += f" / {major_unit}"
                elif unit:
                    tag_text += f" / {unit}"
                if difficulty:
                    tag_text += f" / {difficulty}"
                
                tag_label = QLabel(tag_text)
                tag_label.setFont(QFont("맑은 고딕", 9))
                tag_label.setStyleSheet("""
                    QLabel {
                        background-color: #E8F0FE;
                        color: #1967D2;
                        padding: 6px 12px;
                        border-radius: 4px;
                    }
                """)
                tags_layout.addWidget(tag_label)
            
            tags_layout.addStretch()
            tags_widget.setLayout(tags_layout)
            layout.addWidget(tags_widget)
        
        # 버튼 영역
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        # 원본 HWP 보기 버튼
        btn_view_hwp = QPushButton("원본 HWP 보기")
        btn_view_hwp.setMinimumWidth(120)
        btn_view_hwp.setMinimumHeight(36)
        btn_view_hwp.setFont(QFont("맑은 고딕", 11))
        btn_view_hwp.setEnabled(self.problem_data.get('has_content_raw', False))
        btn_view_hwp.setStyleSheet("""
            QPushButton {
                background-color: #1967D2;
                color: #FFFFFF;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover:enabled {
                background-color: #1557B0;
            }
            QPushButton:disabled {
                background-color: #E0E0E0;
                color: #9E9E9E;
            }
        """)
        btn_view_hwp.clicked.connect(self.on_view_hwp)
        button_layout.addWidget(btn_view_hwp)
        
        # 닫기 버튼
        btn_close = QPushButton("닫기")
        btn_close.setMinimumWidth(80)
        btn_close.setMinimumHeight(36)
        btn_close.setFont(QFont("맑은 고딕", 11))
        btn_close.clicked.connect(self.accept)
        button_layout.addWidget(btn_close)
        
        layout.addLayout(button_layout)
        self.setLayout(layout)
        
        # 스타일 설정
        self.setStyleSheet("""
            QDialog {
                background-color: #FFFFFF;
            }
        """)
    
    def on_view_hwp(self):
        """원본 HWP 보기 버튼 클릭 처리"""
        if not self.hwp_restore:
            from PyQt5.QtWidgets import QMessageBox
            QMessageBox.warning(self, "오류", "HWP 복원 기능이 초기화되지 않았습니다.")
            return
        
        problem_id = self.problem_data.get('problem_id')
        if not problem_id:
            from PyQt5.QtWidgets import QMessageBox
            QMessageBox.warning(self, "오류", "Problem ID가 없습니다.")
            return
        
        try:
            # HWP 파일 복원
            hwp_path = self.hwp_restore.restore_to_temp_file(problem_id)
            
            # Windows에서 파일 열기
            os.startfile(hwp_path)
            
        except Exception as e:
            from PyQt5.QtWidgets import QMessageBox
            QMessageBox.critical(
                self,
                "오류",
                f"원본 HWP를 열 수 없습니다.\n\n{str(e)}"
            )
