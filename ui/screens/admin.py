"""
관리 탭 화면

요구사항(업데이트):
- 메인 상단바/메인 사이드바(수업준비용)는 변경하지 않음
- 관리 탭 콘텐츠 영역에서만 내부 측면메뉴 구성
  - 학생관리 / 반관리
- 학생관리(배포 고려):
  - MongoDB 연결 시: students 컬렉션 CRUD
  - 엑셀(xlsx) 업로드/추출 지원(openpyxl)
  - UI 컬럼은 요청된 6개 필드만 표시:
    학년, 상태(재원/휴원/퇴원), 학생이름, 학교명, 학부모 연락처, 학생 연락처
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import List, Optional

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import (
    QWidget,
    QHBoxLayout,
    QVBoxLayout,
    QFrame,
    QPushButton,
    QLabel,
    QLineEdit,
    QComboBox,
    QTableWidget,
    QTableWidgetItem,
    QHeaderView,
    QFileDialog,
    QMessageBox,
    QDialog,
    QFormLayout,
    QListWidget,
    QListWidgetItem,
    QSplitter,
    QInputDialog,
)

from core.models import Student
from database.repositories.student_repository import StudentRepository
from database.repositories.class_repository import ClassRepository
from services.student import export_students_to_xlsx, import_students_from_xlsx, normalize_phone
from services.login_api import (
    list_users,
    add_user as api_add_user,
    admin_update_user as api_admin_update_user,
    delete_user as api_delete_user,
)


def _font(size_pt: int, *, bold: bool = False, extra_bold: bool = False) -> QFont:
    f = QFont("Pretendard")
    if not f.exactMatch():
        f = QFont("맑은 고딕")
    f.setPointSize(int(size_pt))
    if extra_bold:
        f.setWeight(QFont.ExtraBold)
    elif bold:
        f.setBold(True)
    else:
        try:
            f.setWeight(QFont.Medium)
        except Exception:
            pass
    return f


GRADE_ORDER = ["초1", "초2", "초3", "초4", "초5", "초6", "중1", "중2", "중3", "고1", "고2", "고3"]


@dataclass
class StudentRow:
    """
    UI에서 다루기 위한 로컬 표현(필드 최소화).
    - id는 DB 연결 시에만 사용
    """

    id: Optional[str] = None
    grade: str = ""
    status: str = "재원"
    name: str = ""
    school_name: str = ""
    parent_phone: str = ""
    student_phone: str = ""


@dataclass
class ClassRow:
    """반 UI용 로컬 표현. id는 DB 연결 시 사용."""
    id: Optional[str] = None
    grade: str = ""
    name: str = ""  # 반명
    teacher: str = ""
    note: str = ""
    student_ids: List[str] = field(default_factory=list)


class _NavButton(QPushButton):
    def __init__(self, text: str, parent: Optional[QWidget] = None):
        super().__init__(text, parent)
        self.setCheckable(True)
        self.setFocusPolicy(Qt.NoFocus)
        self.setCursor(Qt.PointingHandCursor)
        self.setFixedHeight(44)
        self.setFont(_font(11, bold=True))
        self.setObjectName("AdminNavBtn")


class _StudentDialog(QDialog):
    def __init__(self, *, title: str, row: Optional[StudentRow] = None, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setModal(True)
        self._row = row

        lay = QVBoxLayout(self)
        lay.setContentsMargins(16, 16, 16, 16)
        lay.setSpacing(12)

        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignLeft)
        form.setFormAlignment(Qt.AlignTop)
        form.setHorizontalSpacing(12)
        form.setVerticalSpacing(10)

        self.cmb_grade = QComboBox()
        self.cmb_grade.addItems(GRADE_ORDER)
        self.cmb_status = QComboBox()
        self.cmb_status.addItems(["재원", "휴원", "퇴원"])
        self.inp_name = QLineEdit()
        self.inp_school = QLineEdit()
        self.inp_parent_phone = QLineEdit()
        self.inp_student_phone = QLineEdit()

        for w in (
            self.cmb_grade,
            self.cmb_status,
            self.inp_name,
            self.inp_school,
            self.inp_parent_phone,
            self.inp_student_phone,
        ):
            try:
                w.setFixedHeight(36)
            except Exception:
                pass

        form.addRow("학년", self.cmb_grade)
        form.addRow("상태", self.cmb_status)
        form.addRow("학생 이름", self.inp_name)
        form.addRow("학교명", self.inp_school)
        form.addRow("학부모 연락처", self.inp_parent_phone)
        form.addRow("학생 연락처", self.inp_student_phone)

        lay.addLayout(form)

        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        self.btn_cancel = QPushButton("취소")
        self.btn_cancel.setFocusPolicy(Qt.NoFocus)
        self.btn_cancel.setCursor(Qt.PointingHandCursor)
        self.btn_cancel.clicked.connect(self.reject)
        self.btn_ok = QPushButton("저장")
        self.btn_ok.setFocusPolicy(Qt.NoFocus)
        self.btn_ok.setCursor(Qt.PointingHandCursor)
        self.btn_ok.setObjectName("primary")
        self.btn_ok.clicked.connect(self._on_ok)
        btn_row.addWidget(self.btn_cancel)
        btn_row.addWidget(self.btn_ok)
        lay.addLayout(btn_row)

        if row is not None:
            self._load_row(row)

    def _load_row(self, r: StudentRow) -> None:
        try:
            self.cmb_grade.setCurrentText(r.grade)
        except Exception:
            pass
        try:
            self.cmb_status.setCurrentText(r.status)
        except Exception:
            pass
        self.inp_name.setText(r.name)
        self.inp_school.setText(r.school_name)
        self.inp_parent_phone.setText(r.parent_phone)
        self.inp_student_phone.setText(r.student_phone)

    def _on_ok(self) -> None:
        if not (self.inp_name.text() or "").strip():
            QMessageBox.information(self, "입력 필요", "학생 이름을 입력해 주세요.")
            return
        self.accept()

    def result_row(self) -> StudentRow:
        return StudentRow(
            id=getattr(self._row, "id", None) if self._row is not None else None,
            grade=(self.cmb_grade.currentText() or "").strip(),
            status=(self.cmb_status.currentText() or "").strip(),
            name=(self.inp_name.text() or "").strip(),
            school_name=(self.inp_school.text() or "").strip(),
            parent_phone=normalize_phone((self.inp_parent_phone.text() or "").strip()),
            student_phone=normalize_phone((self.inp_student_phone.text() or "").strip()),
        )


class _ClassDialog(QDialog):
    def __init__(self, *, title: str, row: Optional[ClassRow] = None, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setModal(True)
        self._row = row

        lay = QVBoxLayout(self)
        lay.setContentsMargins(16, 16, 16, 16)
        lay.setSpacing(12)

        form = QFormLayout()
        form.setHorizontalSpacing(12)
        form.setVerticalSpacing(10)

        self.cmb_grade = QComboBox()
        self.cmb_grade.addItems(GRADE_ORDER)
        self.inp_name = QLineEdit()
        self.inp_name.setPlaceholderText("예: 1반, A반")
        self.inp_teacher = QLineEdit()
        self.inp_teacher.setPlaceholderText("예: 이창현T")
        self.inp_note = QLineEdit()
        self.inp_note.setPlaceholderText("선택")
        for w in (self.cmb_grade, self.inp_name, self.inp_teacher, self.inp_note):
            w.setFixedHeight(36)
        form.addRow("학년", self.cmb_grade)
        form.addRow("반명", self.inp_name)
        form.addRow("담당강사", self.inp_teacher)
        form.addRow("비고", self.inp_note)
        lay.addLayout(form)

        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        btn_cancel = QPushButton("취소")
        btn_cancel.setCursor(Qt.PointingHandCursor)
        btn_cancel.setFocusPolicy(Qt.NoFocus)
        btn_cancel.clicked.connect(self.reject)
        btn_ok = QPushButton("저장")
        btn_ok.setCursor(Qt.PointingHandCursor)
        btn_ok.setFocusPolicy(Qt.NoFocus)
        btn_ok.setObjectName("primary")
        btn_ok.clicked.connect(self._on_ok)
        btn_row.addWidget(btn_cancel)
        btn_row.addWidget(btn_ok)
        lay.addLayout(btn_row)

        if row is not None:
            try:
                self.cmb_grade.setCurrentText(row.grade)
            except Exception:
                pass
            self.inp_name.setText(row.name)
            self.inp_teacher.setText(row.teacher)
            self.inp_note.setText(row.note)

    def _on_ok(self) -> None:
        if not (self.inp_name.text() or "").strip():
            QMessageBox.information(self, "입력 필요", "반명을 입력해 주세요.")
            return
        self.accept()

    def result_row(self) -> ClassRow:
        return ClassRow(
            id=getattr(self._row, "id", None) if self._row is not None else None,
            grade=(self.cmb_grade.currentText() or "").strip(),
            name=(self.inp_name.text() or "").strip(),
            teacher=(self.inp_teacher.text() or "").strip(),
            note=(self.inp_note.text() or "").strip(),
            student_ids=list(getattr(self._row, "student_ids", None) or []),
        )


class StudentManagementView(QWidget):
    def __init__(self, db_connection=None, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.db_connection = db_connection
        self.repo: Optional[StudentRepository] = None
        try:
            if self.db_connection is not None and getattr(self.db_connection, "is_connected", None) and self.db_connection.is_connected():
                self.repo = StudentRepository(self.db_connection)
        except Exception:
            self.repo = None

        # 연결 실패(오프라인)에서도 UI는 동작하도록 로컬 목록 유지
        self._rows: List[StudentRow] = []
        self._grade_filter: str = "전체"
        self._search: str = ""

        self._table: Optional[QTableWidget] = None
        self._hint: Optional[QLabel] = None
        self._build_ui()
        self._load_from_db_or_seed()
        self._refresh()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(10)

        # 상단 필터/액션 바
        bar = QHBoxLayout()
        bar.setSpacing(10)

        self.cmb_sort = QComboBox()
        self.cmb_sort.addItems(["최신 등록순", "이름순"])
        self.cmb_sort.setFixedHeight(36)
        bar.addWidget(self.cmb_sort, 0)

        # 학년(전체/초/중/고) 칩
        self.btn_all = QPushButton("전체")
        self.btn_ele = QPushButton("초")
        self.btn_mid = QPushButton("중")
        self.btn_high = QPushButton("고")
        for b in (self.btn_all, self.btn_ele, self.btn_mid, self.btn_high):
            b.setObjectName("FilterChip")
            b.setCheckable(True)
            b.setCursor(Qt.PointingHandCursor)
            b.setFocusPolicy(Qt.NoFocus)
            b.setFont(_font(10, bold=True))
            b.setFixedHeight(34)
        self.btn_all.setChecked(True)

        def _set_grade_filter(v: str):
            self._grade_filter = v
            self.btn_all.setChecked(v == "전체")
            self.btn_ele.setChecked(v == "초")
            self.btn_mid.setChecked(v == "중")
            self.btn_high.setChecked(v == "고")
            self._refresh()

        self.btn_all.clicked.connect(lambda: _set_grade_filter("전체"))
        self.btn_ele.clicked.connect(lambda: _set_grade_filter("초"))
        self.btn_mid.clicked.connect(lambda: _set_grade_filter("중"))
        self.btn_high.clicked.connect(lambda: _set_grade_filter("고"))

        bar.addWidget(self.btn_all, 0)
        bar.addWidget(self.btn_ele, 0)
        bar.addWidget(self.btn_mid, 0)
        bar.addWidget(self.btn_high, 0)

        bar.addStretch(1)

        self.btn_export = QPushButton("엑셀 추출")
        self.btn_export.setCursor(Qt.PointingHandCursor)
        self.btn_export.setFocusPolicy(Qt.NoFocus)
        self.btn_export.setFixedHeight(36)
        self.btn_export.setObjectName("secondary")
        self.btn_export.clicked.connect(self._on_export)
        bar.addWidget(self.btn_export, 0)

        self.btn_import = QPushButton("엑셀 업로드")
        self.btn_import.setCursor(Qt.PointingHandCursor)
        self.btn_import.setFocusPolicy(Qt.NoFocus)
        self.btn_import.setFixedHeight(36)
        self.btn_import.setObjectName("secondary")
        self.btn_import.clicked.connect(self._on_import)
        bar.addWidget(self.btn_import, 0)

        self.btn_add = QPushButton("학생 추가")
        self.btn_add.setCursor(Qt.PointingHandCursor)
        self.btn_add.setFocusPolicy(Qt.NoFocus)
        self.btn_add.setFixedHeight(36)
        self.btn_add.setObjectName("primary")
        self.btn_add.clicked.connect(self._on_add)
        bar.addWidget(self.btn_add, 0)

        self.btn_delete_selected = QPushButton("선택 삭제")
        self.btn_delete_selected.setCursor(Qt.PointingHandCursor)
        self.btn_delete_selected.setFocusPolicy(Qt.NoFocus)
        self.btn_delete_selected.setFixedHeight(36)
        self.btn_delete_selected.setObjectName("secondary")
        self.btn_delete_selected.clicked.connect(self._on_delete_selected)
        bar.addWidget(self.btn_delete_selected, 0)

        self.search = QLineEdit()
        self.search.setObjectName("StudentSearchBox")
        self.search.setPlaceholderText("학생 이름 검색")
        self.search.setClearButtonEnabled(True)
        self.search.setFixedHeight(36)
        self.search.setMinimumWidth(220)
        self.search.textChanged.connect(self._on_search)
        bar.addWidget(self.search, 0)

        root.addLayout(bar)

        # 표
        table = QTableWidget()
        self._table = table
        # ✅ 가독성: 폰트/행 높이/패딩/색상
        table.setFont(_font(10, bold=False))
        table.setColumnCount(8)
        table.setHorizontalHeaderLabels(["학년", "상태", "학생 이름", "학교명", "학부모 연락처", "학생 연락처", "수정", "삭제"])
        table.verticalHeader().setVisible(False)
        # 요청: 표 셀 배경은 모두 흰색(헤더만 배경색)
        table.setAlternatingRowColors(False)
        table.setShowGrid(False)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.setSelectionMode(QTableWidget.ExtendedSelection)  # Ctrl/Shift/드래그로 여러 행 선택
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        try:
            table.verticalHeader().setDefaultSectionSize(46)
        except Exception:
            pass
        try:
            table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
            table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
            table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
            table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
            table.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeToContents)
            table.horizontalHeader().setSectionResizeMode(7, QHeaderView.ResizeToContents)
        except Exception:
            pass
        # ✅ 테이블 스타일(로컬): 셀 패딩 + 진한 글씨 + 헤더 가독성
        table.setStyleSheet(
            """
            QTableWidget {
                background: #FFFFFF;
                border: 1px solid #CBD5E1; /* 더 진한 테두리 */
                border-radius: 12px;
            }
            QHeaderView::section {
                background: #F1F5F9; /* 헤더만 배경색 */
                color: #0F172A;
                font-weight: 900;
                padding: 10px 12px;
                border: none;
                border-bottom: 1px solid #CBD5E1;
            }
            QTableWidget::item {
                padding: 10px 12px;
                color: #0F172A;
                background: #FFFFFF;
            }
            QTableWidget::item:selected {
                background: #E0F2FE;
                color: #0F172A;
            }
            """
        )
        table.setFixedHeight(520)
        root.addWidget(table)

        self._hint = QLabel("")
        self._hint.setFont(_font(9, bold=True))
        self._hint.setStyleSheet("color:#64748B;")
        root.addWidget(self._hint)
        root.addStretch(1)

    def _on_search(self, text: str) -> None:
        self._search = (text or "").strip()
        self._refresh()

    def _load_from_db_or_seed(self) -> None:
        # DB 연결 시: students 컬렉션 로드
        if self.repo is not None:
            try:
                items = self.repo.list_all()
                self._rows = [
                    StudentRow(
                        id=s.id,
                        grade=s.grade,
                        status=s.status,
                        name=s.name,
                        school_name=s.school_name,
                        parent_phone=s.parent_phone,
                        student_phone=s.student_phone,
                    )
                    for s in (items or [])
                ]
                if self._hint is not None:
                    self._hint.setText("")
                return
            except Exception as e:
                # 실패 시 로컬로 폴백
                if self._hint is not None:
                    self._hint.setText("")

        # 오프라인/초기: 목데이터
        self._rows = [
            StudentRow(None, "중3", "재원", "강라희", "OO중학교", normalize_phone("010-9853-1545"), normalize_phone("010-9853-1545")),
            StudentRow(None, "중1", "재원", "임아윤", "OO중학교", normalize_phone("010-2787-0289"), normalize_phone("010-2787-0289")),
            StudentRow(None, "중2", "재원", "심서우", "OO중학교", normalize_phone("010-9079-6599"), normalize_phone("010-9079-6599")),
        ]
        if self._hint is not None:
            self._hint.setText("")

    def _on_export(self) -> None:
        rows = self._filtered_rows()
        if not rows:
            QMessageBox.information(self, "엑셀 추출", "추출할 학생 데이터가 없습니다.")
            return

        path, _ = QFileDialog.getSaveFileName(self, "엑셀로 저장", "students.xlsx", "Excel (*.xlsx)")
        if not path:
            return

        students = [
            Student(
                grade=r.grade,
                status=r.status,
                name=r.name,
                school_name=r.school_name,
                parent_phone=r.parent_phone,
                student_phone=r.student_phone,
            )
            for r in rows
        ]
        try:
            export_students_to_xlsx(students, path)
        except Exception as e:
            QMessageBox.critical(self, "엑셀 추출 실패", str(e))
            return
        QMessageBox.information(self, "엑셀 추출 완료", f"저장 완료:\n{path}")

    def _on_import(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "엑셀 업로드", "", "Excel (*.xlsx)")
        if not path:
            return
        try:
            students, stats = import_students_from_xlsx(path)
        except Exception as e:
            QMessageBox.critical(self, "엑셀 업로드 실패", str(e))
            return

        if not students:
            QMessageBox.information(self, "엑셀 업로드", "가져올 학생 데이터가 없습니다.")
            return

        # DB 연결 시 upsert, 아니면 로컬에 병합
        if self.repo is not None:
            try:
                result = self.repo.bulk_upsert(students)
            except Exception as e:
                QMessageBox.critical(self, "DB 저장 실패", str(e))
                return
            self._load_from_db_or_seed()
            self._refresh()
            QMessageBox.information(
                self,
                "엑셀 업로드 완료",
                f"총 행: {stats.get('rows', 0)}\n"
                f"가져옴: {stats.get('imported', 0)}\n"
                f"스킵: {stats.get('skipped', 0)}\n\n"
                f"DB 반영: 추가 {result.get('inserted', 0)}, 업데이트 {result.get('updated', 0)}, 스킵 {result.get('skipped', 0)}",
            )
            return

        # 로컬 병합(간단): 이름 기준으로 덮어쓰기
        merged = 0
        for s in students:
            key = (s.name or "").strip()
            if not key:
                continue
            found = False
            for i, r in enumerate(self._rows):
                if (r.name or "").strip() == key:
                    self._rows[i] = StudentRow(
                        id=r.id,
                        grade=s.grade,
                        status=s.status,
                        name=s.name,
                        school_name=s.school_name,
                        parent_phone=s.parent_phone,
                        student_phone=s.student_phone,
                    )
                    found = True
                    merged += 1
                    break
            if not found:
                self._rows.append(
                    StudentRow(
                        id=None,
                        grade=s.grade,
                        status=s.status,
                        name=s.name,
                        school_name=s.school_name,
                        parent_phone=s.parent_phone,
                        student_phone=s.student_phone,
                    )
                )
        self._refresh()
        QMessageBox.information(self, "엑셀 업로드 완료", f"가져옴: {len(students)} (로컬 병합: {merged})")

    def _on_add(self) -> None:
        dlg = _StudentDialog(title="학생 등록", parent=self)
        if dlg.exec_() != dlg.Accepted:
            return
        row = dlg.result_row()
        # DB 저장
        if self.repo is not None:
            try:
                _id = self.repo.create(
                    Student(
                        grade=row.grade,
                        status=row.status,
                        name=row.name,
                        school_name=row.school_name,
                        parent_phone=row.parent_phone,
                        student_phone=row.student_phone,
                    )
                )
                row.id = _id
                self._load_from_db_or_seed()
            except Exception as e:
                QMessageBox.critical(self, "저장 실패", str(e))
                return
        else:
            self._rows.insert(0, row)
        self._refresh()

    def _on_edit(self, row_idx: int) -> None:
        if not (0 <= row_idx < len(self._filtered_rows())):
            return
        r = self._filtered_rows()[row_idx]
        dlg = _StudentDialog(title="학생 정보 수정", row=r, parent=self)
        if dlg.exec_() != dlg.Accepted:
            return
        new_r = dlg.result_row()
        # DB update
        if self.repo is not None and new_r.id:
            try:
                ok = self.repo.update(
                    Student(
                        id=new_r.id,
                        grade=new_r.grade,
                        status=new_r.status,
                        name=new_r.name,
                        school_name=new_r.school_name,
                        parent_phone=new_r.parent_phone,
                        student_phone=new_r.student_phone,
                    )
                )
                if not ok:
                    QMessageBox.warning(self, "수정 실패", "학생 정보를 수정하지 못했습니다.")
                self._load_from_db_or_seed()
            except Exception as e:
                QMessageBox.critical(self, "수정 실패", str(e))
                return
        else:
            # 원본 리스트에서 해당 항목을 찾아 교체
            for i, rr in enumerate(self._rows):
                if rr is r:
                    self._rows[i] = new_r
                    break
        self._refresh()

    def _on_delete(self, row_idx: int) -> None:
        rows = self._filtered_rows()
        if not (0 <= row_idx < len(rows)):
            return
        r = rows[row_idx]
        if QMessageBox.question(self, "삭제 확인", f"학생 '{r.name}'을(를) 삭제할까요?") != QMessageBox.Yes:
            return

        if self.repo is not None and r.id:
            try:
                ok = self.repo.soft_delete(r.id)
                if not ok:
                    QMessageBox.warning(self, "삭제 실패", "학생을 삭제하지 못했습니다.")
                self._load_from_db_or_seed()
            except Exception as e:
                QMessageBox.critical(self, "삭제 실패", str(e))
                return
        else:
            self._rows = [x for x in self._rows if x is not r]
        self._refresh()

    def _on_delete_selected(self) -> None:
        """선택된 행(여러 명) 일괄 삭제 — Ctrl/Shift/드래그로 선택 후 사용."""
        table = self._table
        if table is None:
            return
        rows = self._filtered_rows()
        selected_indexes = table.selectedIndexes()
        row_indices = sorted(set(idx.row() for idx in selected_indexes), reverse=True)
        to_delete = [rows[i] for i in row_indices if 0 <= i < len(rows)]
        if not to_delete:
            QMessageBox.information(self, "선택 삭제", "삭제할 학생을 선택해 주세요. (Ctrl+클릭, Shift+클릭, 드래그로 여러 명 선택)")
            return
        n = len(to_delete)
        if QMessageBox.question(
            self,
            "선택 삭제 확인",
            f"선택한 {n}명의 학생을 삭제할까요?\n\n" + "\n".join(f"· {r.name} ({r.grade})" for r in to_delete[:10]) + ("\n..." if n > 10 else ""),
        ) != QMessageBox.Yes:
            return
        failed = 0
        for r in to_delete:
            if self.repo is not None and r.id:
                try:
                    if not self.repo.soft_delete(r.id):
                        failed += 1
                except Exception:
                    failed += 1
            else:
                self._rows = [x for x in self._rows if x is not r]
        if self.repo is not None:
            self._load_from_db_or_seed()
        self._refresh()
        if failed:
            QMessageBox.warning(self, "선택 삭제", f"{n - failed}명 삭제됨. {failed}명 삭제 실패.")
        else:
            QMessageBox.information(self, "선택 삭제", f"{n}명 삭제되었습니다.")

    def _filtered_rows(self) -> List[StudentRow]:
        out = list(self._rows)
        if self._grade_filter != "전체":
            if self._grade_filter == "초":
                out = [r for r in out if (r.grade or "").startswith("초")]
            elif self._grade_filter == "중":
                out = [r for r in out if (r.grade or "").startswith("중")]
            elif self._grade_filter == "고":
                out = [r for r in out if (r.grade or "").startswith("고")]
        if self._search:
            q = self._search.lower()
            out = [r for r in out if (q in (r.name or "").lower()) or (q in (r.school_name or "").lower())]
        # 정렬(간단)
        if (self.cmb_sort.currentText() or "") == "이름순":
            out = sorted(out, key=lambda x: (x.name or ""))
        return out

    def _refresh(self) -> None:
        table = self._table
        if table is None:
            return
        rows = self._filtered_rows()
        table.setRowCount(len(rows))

        for i, r in enumerate(rows):
            for col, text in [
                (0, r.grade),
                (1, r.status),
                (2, r.name),
                (3, r.school_name),
                (4, r.parent_phone),
                (5, r.student_phone),
            ]:
                it = QTableWidgetItem(text or "")
                it.setForeground(Qt.black)  # ✅ 더 진한 텍스트
                it.setFont(_font(10, bold=True if col in (0, 1, 2) else False))
                table.setItem(i, col, it)

            btn_edit = QPushButton("수정")
            btn_edit.setCursor(Qt.PointingHandCursor)
            btn_edit.setFocusPolicy(Qt.NoFocus)
            # 요청: 수정/삭제 버튼 더 작게
            btn_edit.setFixedSize(54, 26)
            btn_edit.setObjectName("RowEditBtn")
            btn_edit.setFont(_font(9, bold=True))
            btn_edit.clicked.connect(lambda _=False, idx=i: self._on_edit(idx))
            table.setCellWidget(i, 6, btn_edit)

            btn_del = QPushButton("삭제")
            btn_del.setCursor(Qt.PointingHandCursor)
            btn_del.setFocusPolicy(Qt.NoFocus)
            btn_del.setFixedSize(54, 26)
            btn_del.setObjectName("RowDeleteBtn")
            btn_del.setFont(_font(9, bold=True))
            btn_del.clicked.connect(lambda _=False, idx=i: self._on_delete(idx))
            table.setCellWidget(i, 7, btn_del)

        try:
            table.resizeRowsToContents()
        except Exception:
            pass


class _ClassStudentsDialog(QDialog):
    """반 소속 학생 추가/제거 다이얼로그. 왼쪽=소속 학생, 오른쪽=추가 가능 학생."""

    def __init__(
        self,
        *,
        class_row: ClassRow,
        student_repo: Optional[StudentRepository],
        class_repo: Optional[ClassRepository],
        parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self.setWindowTitle(f"학생 관리 — {class_row.grade} {class_row.name}")
        self.setModal(True)
        self._class_row = class_row
        self._student_repo = student_repo
        self._class_repo = class_repo
        self._member_ids: List[str] = list(class_row.student_ids or [])

        lay = QVBoxLayout(self)
        lay.setContentsMargins(16, 16, 16, 16)
        lay.setSpacing(12)

        split = QSplitter(Qt.Horizontal)

        # 왼쪽: 이 반 소속 학생 (Ctrl/드래그로 여러 명 선택 후 제거 가능)
        left = QFrame()
        left_lay = QVBoxLayout(left)
        left_lay.setContentsMargins(0, 0, 0, 0)
        left_lay.addWidget(QLabel("소속 학생"))
        self.list_member = QListWidget()
        self.list_member.setMinimumWidth(200)
        self.list_member.setSelectionMode(QListWidget.ExtendedSelection)
        left_lay.addWidget(self.list_member)
        btn_remove = QPushButton("선택 제거")
        btn_remove.setCursor(Qt.PointingHandCursor)
        btn_remove.setFocusPolicy(Qt.NoFocus)
        btn_remove.clicked.connect(self._on_remove)
        left_lay.addWidget(btn_remove)
        split.addWidget(left)

        # 오른쪽: 추가 가능한 학생 (Ctrl/드래그로 여러 명 선택 후 추가 가능)
        right = QFrame()
        right_lay = QVBoxLayout(right)
        right_lay.setContentsMargins(0, 0, 0, 0)
        right_lay.addWidget(QLabel("추가할 학생 (Ctrl/드래그로 여러 명 선택)"))
        self.list_available = QListWidget()
        self.list_available.setMinimumWidth(200)
        self.list_available.setSelectionMode(QListWidget.ExtendedSelection)
        right_lay.addWidget(self.list_available)
        btn_add = QPushButton("선택 추가")
        btn_add.setCursor(Qt.PointingHandCursor)
        btn_add.setFocusPolicy(Qt.NoFocus)
        btn_add.clicked.connect(self._on_add)
        right_lay.addWidget(btn_add)
        split.addWidget(right)

        lay.addWidget(split)

        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        btn_cancel = QPushButton("취소")
        btn_cancel.setCursor(Qt.PointingHandCursor)
        btn_cancel.setFocusPolicy(Qt.NoFocus)
        btn_cancel.clicked.connect(self.reject)
        btn_ok = QPushButton("저장")
        btn_ok.setCursor(Qt.PointingHandCursor)
        btn_ok.setFocusPolicy(Qt.NoFocus)
        btn_ok.setObjectName("primary")
        btn_ok.clicked.connect(self._on_save)
        btn_row.addWidget(btn_cancel)
        btn_row.addWidget(btn_ok)
        lay.addLayout(btn_row)

        self._fill_lists()

    def _student_label(self, s: Optional[Student]) -> str:
        if not s:
            return "(알 수 없음)"
        return f"{s.name or ''} ({s.grade or ''})"

    def _fill_lists(self) -> None:
        self.list_member.clear()
        self.list_available.clear()
        all_students = []
        if self._student_repo:
            try:
                all_students = self._student_repo.list_all()
            except Exception:
                pass
        member_set = set(str(x) for x in self._member_ids if x)
        id_to_student = {str(s.id): s for s in all_students if s and s.id}
        for sid in self._member_ids:
            if not sid:
                continue
            s = id_to_student.get(str(sid))
            it = QListWidgetItem(self._student_label(s))
            it.setData(Qt.UserRole, str(sid))
            self.list_member.addItem(it)
        for s in all_students:
            if not s or not s.id:
                continue
            if str(s.id) in member_set:
                continue
            it = QListWidgetItem(self._student_label(s))
            it.setData(Qt.UserRole, str(s.id))
            self.list_available.addItem(it)

    def _on_add(self) -> None:
        selected = self.list_available.selectedItems()
        if not selected:
            return
        for it in selected:
            sid = it.data(Qt.UserRole)
            if sid and sid not in self._member_ids:
                self._member_ids.append(sid)
        self._fill_lists()

    def _on_remove(self) -> None:
        selected = self.list_member.selectedItems()
        if not selected:
            return
        remove_set = {str(it.data(Qt.UserRole)) for it in selected if it.data(Qt.UserRole)}
        self._member_ids = [x for x in self._member_ids if str(x) not in remove_set]
        self._fill_lists()

    def _on_save(self) -> None:
        self._class_row.student_ids = list(self._member_ids)
        if self._class_repo and self._class_row.id:
            try:
                from core.models import SchoolClass
                sc = SchoolClass(
                    id=self._class_row.id,
                    grade=self._class_row.grade,
                    name=self._class_row.name,
                    teacher=self._class_row.teacher,
                    note=self._class_row.note,
                    student_ids=list(self._member_ids),
                )
                self._class_repo.update(sc)
            except Exception as e:
                QMessageBox.critical(self, "저장 실패", str(e))
                return
        self.accept()


class ClassManagementView(QWidget):
    def __init__(self, db_connection=None, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.db_connection = db_connection
        self.repo: Optional[ClassRepository] = None
        self._student_repo: Optional[StudentRepository] = None
        try:
            if self.db_connection and getattr(self.db_connection, "is_connected", None) and self.db_connection.is_connected():
                self.repo = ClassRepository(self.db_connection)
                self._student_repo = StudentRepository(self.db_connection)
        except Exception:
            self.repo = None
            self._student_repo = None
        self._rows: List[ClassRow] = []
        self._search: str = ""
        self._table: Optional[QTableWidget] = None
        self._build_ui()
        self._load_from_db_or_seed()
        self._refresh()

    def _load_from_db_or_seed(self) -> None:
        if self.repo is None:
            self._rows = []
            return
        try:
            classes = self.repo.list_all()
            self._rows = [
                ClassRow(
                    id=str(c.id),
                    grade=c.grade or "",
                    name=c.name or "",
                    teacher=c.teacher or "",
                    note=c.note or "",
                    student_ids=list(c.student_ids or []),
                )
                for c in classes
            ]
        except Exception:
            self._rows = []

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(10)

        bar = QHBoxLayout()
        bar.setSpacing(10)

        title = QLabel("반 관리")
        title.setFont(_font(14, extra_bold=True))
        bar.addWidget(title)
        bar.addStretch(1)

        self.search = QLineEdit()
        self.search.setPlaceholderText("반명·학년·강사 검색")
        self.search.setClearButtonEnabled(True)
        self.search.setFixedHeight(36)
        self.search.setMinimumWidth(220)
        self.search.textChanged.connect(self._on_search)
        bar.addWidget(self.search, 0)

        self.btn_add = QPushButton("반 추가")
        self.btn_add.setObjectName("primary")
        self.btn_add.setCursor(Qt.PointingHandCursor)
        self.btn_add.setFocusPolicy(Qt.NoFocus)
        self.btn_add.setFixedHeight(36)
        self.btn_add.clicked.connect(self._on_add)
        bar.addWidget(self.btn_add, 0)

        root.addLayout(bar)

        table = QTableWidget()
        self._table = table
        table.setFont(_font(10, bold=False))
        table.setColumnCount(8)
        table.setHorizontalHeaderLabels(["학년", "반명", "담당강사", "학생 수", "비고", "수정", "삭제", "학생 관리"])
        table.verticalHeader().setVisible(False)
        table.setAlternatingRowColors(False)
        table.setShowGrid(False)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.setSelectionMode(QTableWidget.SingleSelection)
        try:
            table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
            table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
            table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
            table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
            table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
            table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)
            table.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeToContents)
            table.horizontalHeader().setSectionResizeMode(7, QHeaderView.ResizeToContents)
        except Exception:
            pass
        try:
            table.verticalHeader().setDefaultSectionSize(46)
        except Exception:
            pass
        table.setStyleSheet(
            """
            QTableWidget {
                background: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 12px;
            }
            QHeaderView::section {
                background: #F1F5F9;
                color: #0F172A;
                font-weight: 900;
                padding: 10px 12px;
                border: none;
                border-bottom: 1px solid #CBD5E1;
            }
            QTableWidget::item {
                padding: 10px 12px;
                color: #0F172A;
                background: #FFFFFF;
            }
            QTableWidget::item:selected {
                background: #E0F2FE;
                color: #0F172A;
            }
            """
        )
        table.setFixedHeight(520)
        root.addWidget(table)

        self._hint = QLabel("")
        self._hint.setFont(_font(9, bold=True))
        self._hint.setStyleSheet("color:#64748B;")
        root.addWidget(self._hint)
        root.addStretch(1)

    def _on_search(self, text: str) -> None:
        self._search = (text or "").strip()
        self._refresh()

    def _on_add(self) -> None:
        dlg = _ClassDialog(title="반 추가", parent=self)
        if dlg.exec_() != dlg.Accepted:
            return
        r = dlg.result_row()
        if self.repo is not None:
            try:
                from core.models import SchoolClass
                sc = SchoolClass(grade=r.grade, name=r.name, teacher=r.teacher, note=r.note, student_ids=[])
                new_id = self.repo.create(sc)
                r.id = new_id
                r.student_ids = []
            except Exception as e:
                QMessageBox.critical(self, "반 추가 실패", str(e))
                return
        self._rows.append(r)
        if self.repo is not None:
            self._load_from_db_or_seed()
        self._refresh()

    def _on_edit(self, idx: int) -> None:
        rows = self._filtered_rows()
        if not (0 <= idx < len(rows)):
            return
        r = rows[idx]
        dlg = _ClassDialog(title="반 수정", row=r, parent=self)
        if dlg.exec_() != dlg.Accepted:
            return
        new_r = dlg.result_row()
        if self.repo is not None and r.id:
            try:
                from core.models import SchoolClass
                sc = SchoolClass(
                    id=r.id,
                    grade=new_r.grade,
                    name=new_r.name,
                    teacher=new_r.teacher,
                    note=new_r.note,
                    student_ids=list(new_r.student_ids or []),
                )
                self.repo.update(sc)
            except Exception as e:
                QMessageBox.warning(self, "수정 실패", str(e))
                return
        for i, rr in enumerate(self._rows):
            if rr is r:
                self._rows[i] = new_r
                if self._rows[i].id is None:
                    self._rows[i].id = r.id
                if not (self._rows[i].student_ids):
                    self._rows[i].student_ids = list(r.student_ids or [])
                break
        self._load_from_db_or_seed()
        self._refresh()

    def _on_delete(self, idx: int) -> None:
        rows = self._filtered_rows()
        if not (0 <= idx < len(rows)):
            return
        r = rows[idx]
        if QMessageBox.question(self, "삭제 확인", f"반 '{r.grade} {r.name}'을(를) 삭제할까요?") != QMessageBox.Yes:
            return
        if self.repo is not None and r.id:
            try:
                ok = self.repo.soft_delete(r.id)
                if not ok:
                    QMessageBox.warning(self, "삭제 실패", "반을 삭제하지 못했습니다.")
                self._load_from_db_or_seed()
            except Exception as e:
                QMessageBox.critical(self, "삭제 실패", str(e))
                return
        else:
            self._rows = [x for x in self._rows if x is not r]
        self._refresh()

    def _on_manage_students(self, idx: int) -> None:
        rows = self._filtered_rows()
        if not (0 <= idx < len(rows)):
            return
        r = rows[idx]
        dlg = _ClassStudentsDialog(
            class_row=r,
            student_repo=self._student_repo,
            class_repo=self.repo,
            parent=self,
        )
        dlg.exec_()
        self._load_from_db_or_seed()
        self._refresh()

    def _filtered_rows(self) -> List[ClassRow]:
        out = list(self._rows)
        if self._search:
            q = self._search.lower()
            out = [
                r for r in out
                if q in (r.name or "").lower()
                or q in (r.grade or "").lower()
                or q in (r.teacher or "").lower()
            ]
        idx = {k: i for i, k in enumerate(GRADE_ORDER)}
        out = sorted(out, key=lambda r: (idx.get(r.grade, 10_000), (r.name or "")))
        return out

    def _refresh(self) -> None:
        table = self._table
        if table is None:
            return
        rows = self._filtered_rows()
        table.setRowCount(len(rows))
        for i, r in enumerate(rows):
            n = len(r.student_ids) if r.student_ids else 0
            for col, text in [
                (0, r.grade),
                (1, r.name),
                (2, r.teacher),
                (3, str(n)),
                (4, r.note),
            ]:
                it = QTableWidgetItem(text or "")
                it.setForeground(Qt.black)
                it.setFont(_font(10, bold=True if col in (0, 1) else False))
                table.setItem(i, col, it)

            btn_edit = QPushButton("수정")
            btn_edit.setCursor(Qt.PointingHandCursor)
            btn_edit.setFocusPolicy(Qt.NoFocus)
            btn_edit.setFixedSize(54, 26)
            btn_edit.setObjectName("RowEditBtn")
            btn_edit.setFont(_font(9, bold=True))
            btn_edit.clicked.connect(lambda _=False, idx=i: self._on_edit(idx))
            table.setCellWidget(i, 5, btn_edit)

            btn_del = QPushButton("삭제")
            btn_del.setCursor(Qt.PointingHandCursor)
            btn_del.setFocusPolicy(Qt.NoFocus)
            btn_del.setFixedSize(54, 26)
            btn_del.setObjectName("RowDeleteBtn")
            btn_del.setFont(_font(9, bold=True))
            btn_del.clicked.connect(lambda _=False, idx=i: self._on_delete(idx))
            table.setCellWidget(i, 6, btn_del)

            btn_member = QPushButton("학생 관리")
            btn_member.setCursor(Qt.PointingHandCursor)
            btn_member.setFocusPolicy(Qt.NoFocus)
            btn_member.setFixedSize(80, 26)
            btn_member.setObjectName("RowEditBtn")
            btn_member.setFont(_font(9, bold=True))
            btn_member.clicked.connect(lambda _=False, idx=i: self._on_manage_students(idx))
            table.setCellWidget(i, 7, btn_member)

        try:
            table.resizeRowsToContents()
        except Exception:
            pass


class _AddMemberDialog(QDialog):
    """회원 추가 다이얼로그 (아이디, 비밀번호, 이름)"""
    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle("회원 추가")
        self.setModal(True)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(16, 16, 16, 16)
        form = QFormLayout()
        self.inp_user_id = QLineEdit()
        self.inp_user_id.setPlaceholderText("로그인 아이디")
        self.inp_password = QLineEdit()
        self.inp_password.setEchoMode(QLineEdit.Password)
        self.inp_password.setPlaceholderText("비밀번호 (제약 없음)")
        self.inp_name = QLineEdit()
        self.inp_name.setPlaceholderText("이름 (선택)")
        for w in (self.inp_user_id, self.inp_password, self.inp_name):
            w.setFixedHeight(36)
        form.addRow("아이디", self.inp_user_id)
        form.addRow("비밀번호", self.inp_password)
        form.addRow("이름", self.inp_name)
        lay.addLayout(form)
        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        btn_cancel = QPushButton("취소")
        btn_cancel.setFocusPolicy(Qt.NoFocus)
        btn_cancel.setCursor(Qt.PointingHandCursor)
        btn_cancel.clicked.connect(self.reject)
        btn_ok = QPushButton("추가")
        btn_ok.setObjectName("primary")
        btn_ok.setFocusPolicy(Qt.NoFocus)
        btn_ok.setCursor(Qt.PointingHandCursor)
        btn_ok.clicked.connect(self._on_ok)
        btn_row.addWidget(btn_cancel)
        btn_row.addWidget(btn_ok)
        lay.addLayout(btn_row)

    def _on_ok(self) -> None:
        if not (self.inp_user_id.text() or "").strip():
            QMessageBox.information(self, "입력 필요", "아이디를 입력해 주세요.")
            return
        self.accept()

    def get_data(self) -> dict:
        return {
            "user_id": (self.inp_user_id.text() or "").strip(),
            "password": self.inp_password.text() or "",
            "name": (self.inp_name.text() or "").strip(),
        }


class _EditMemberDialog(QDialog):
    """관리자용 회원 수정 다이얼로그 (대상 아이디, 이름, 새 비밀번호 + 관리자 비밀번호)"""
    def __init__(self, target_user_id: str, target_name: str, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle("회원 정보 수정")
        self.setModal(True)
        self._target_user_id = target_user_id
        lay = QVBoxLayout(self)
        lay.setContentsMargins(16, 16, 16, 16)
        form = QFormLayout()
        self.inp_user_id = QLineEdit()
        self.inp_user_id.setReadOnly(True)
        self.inp_user_id.setText(target_user_id)
        self.inp_name = QLineEdit()
        self.inp_name.setPlaceholderText("이름 (선택)")
        self.inp_name.setText(target_name or "")
        self.inp_new_password = QLineEdit()
        self.inp_new_password.setEchoMode(QLineEdit.Password)
        self.inp_new_password.setPlaceholderText("변경 시에만 입력")
        self.inp_admin_password = QLineEdit()
        self.inp_admin_password.setEchoMode(QLineEdit.Password)
        self.inp_admin_password.setPlaceholderText("관리자 비밀번호 (필수)")
        for w in (self.inp_user_id, self.inp_name, self.inp_new_password, self.inp_admin_password):
            w.setFixedHeight(36)
        form.addRow("아이디", self.inp_user_id)
        form.addRow("이름", self.inp_name)
        form.addRow("새 비밀번호", self.inp_new_password)
        form.addRow("관리자 비밀번호", self.inp_admin_password)
        lay.addLayout(form)
        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        btn_cancel = QPushButton("취소")
        btn_cancel.setFocusPolicy(Qt.NoFocus)
        btn_cancel.setCursor(Qt.PointingHandCursor)
        btn_cancel.clicked.connect(self.reject)
        btn_ok = QPushButton("저장")
        btn_ok.setObjectName("primary")
        btn_ok.setFocusPolicy(Qt.NoFocus)
        btn_ok.setCursor(Qt.PointingHandCursor)
        btn_ok.clicked.connect(self._on_ok)
        btn_row.addWidget(btn_cancel)
        btn_row.addWidget(btn_ok)
        lay.addLayout(btn_row)

    def _on_ok(self) -> None:
        admin_pw = (self.inp_admin_password.text() or "").strip()
        if not admin_pw:
            QMessageBox.information(self, "입력 필요", "관리자 비밀번호를 입력해 주세요.")
            return
        self.accept()

    def get_data(self) -> dict:
        new_name = (self.inp_name.text() or "").strip()
        new_password = (self.inp_new_password.text() or "").strip() or None
        admin_password = self.inp_admin_password.text() or ""
        return {
            "target_user_id": self._target_user_id,
            "new_name": new_name if new_name else None,
            "new_password": new_password,
            "admin_password": admin_password,
        }


class _DeleteMemberDialog(QDialog):
    """회원 삭제 확인 + 관리자 비밀번호 입력"""
    def __init__(self, target_user_id: str, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle("회원 삭제")
        self.setModal(True)
        self._target_user_id = target_user_id
        lay = QVBoxLayout(self)
        lay.setContentsMargins(16, 16, 16, 16)
        lay.addWidget(QLabel(f"회원 '{target_user_id}'을(를) 삭제합니다. 계속하려면 관리자 비밀번호를 입력하세요."))
        self.inp_admin_password = QLineEdit()
        self.inp_admin_password.setEchoMode(QLineEdit.Password)
        self.inp_admin_password.setPlaceholderText("관리자 비밀번호")
        self.inp_admin_password.setFixedHeight(36)
        lay.addWidget(self.inp_admin_password)
        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        btn_cancel = QPushButton("취소")
        btn_cancel.setFocusPolicy(Qt.NoFocus)
        btn_cancel.setCursor(Qt.PointingHandCursor)
        btn_cancel.clicked.connect(self.reject)
        btn_ok = QPushButton("삭제")
        btn_ok.setObjectName("primary")
        btn_ok.setFocusPolicy(Qt.NoFocus)
        btn_ok.setCursor(Qt.PointingHandCursor)
        btn_ok.clicked.connect(self._on_ok)
        btn_row.addWidget(btn_cancel)
        btn_row.addWidget(btn_ok)
        lay.addLayout(btn_row)

    def _on_ok(self) -> None:
        if not (self.inp_admin_password.text() or "").strip():
            QMessageBox.information(self, "입력 필요", "관리자 비밀번호를 입력해 주세요.")
            return
        self.accept()

    def get_admin_password(self) -> str:
        return self.inp_admin_password.text() or ""


class MemberManagementView(QWidget):
    """회원 관리 뷰 (서버 API: 목록 조회, 회원 추가/수정/삭제)"""
    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self._users: List[dict] = []
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        top = QHBoxLayout()
        top.addWidget(QLabel("회원 목록 (로그인 API 회원)"))
        top.addStretch(1)
        self.btn_add = QPushButton("회원 추가")
        self.btn_add.setObjectName("primary")
        self.btn_add.setCursor(Qt.PointingHandCursor)
        self.btn_add.setFocusPolicy(Qt.NoFocus)
        self.btn_add.clicked.connect(self._on_add)
        self.btn_refresh = QPushButton("새로고침")
        self.btn_refresh.setObjectName("secondary")
        self.btn_refresh.setCursor(Qt.PointingHandCursor)
        self.btn_refresh.setFocusPolicy(Qt.NoFocus)
        self.btn_refresh.clicked.connect(self.load_users)
        top.addWidget(self.btn_add)
        top.addWidget(self.btn_refresh)
        lay.addLayout(top)
        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["아이디", "이름", "등록일", "수정", "삭제"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.table.setAlternatingRowColors(False)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        try:
            self.table.verticalHeader().setDefaultSectionSize(46)
        except Exception:
            pass
        self.table.setStyleSheet(
            """
            QTableWidget {
                background: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 12px;
            }
            QHeaderView::section {
                background: #F1F5F9;
                color: #0F172A;
                font-weight: 900;
                padding: 10px 12px;
                border: none;
                border-bottom: 1px solid #CBD5E1;
            }
            QTableWidget::item {
                padding: 10px 12px;
                color: #0F172A;
                background: #FFFFFF;
            }
            QTableWidget::item:selected {
                background: #E0F2FE;
                color: #0F172A;
            }
            """
        )
        lay.addWidget(self.table)
        self.load_users()

    def _admin_user_id(self) -> Optional[str]:
        from services.login_api import load_session
        session = load_session()
        return (session.get("user_id") or "").strip() or None

    def load_users(self) -> None:
        result = list_users()
        self._users = []
        self.table.setRowCount(0)
        if not result.get("success"):
            return
        for u in result.get("users") or []:
            self._users.append({
                "user_id": u.get("user_id") or "",
                "name": u.get("name") or "",
                "created_at": u.get("created_at") or "",
            })
        for i, u in enumerate(self._users):
            row = self.table.rowCount()
            self.table.insertRow(row)
            self.table.setItem(row, 0, QTableWidgetItem(u.get("user_id") or ""))
            self.table.setItem(row, 1, QTableWidgetItem(u.get("name") or ""))
            self.table.setItem(row, 2, QTableWidgetItem(u.get("created_at") or ""))
            btn_edit = QPushButton("수정")
            btn_edit.setCursor(Qt.PointingHandCursor)
            btn_edit.setFocusPolicy(Qt.NoFocus)
            btn_edit.setFixedSize(54, 26)
            btn_edit.setObjectName("RowEditBtn")
            btn_edit.setFont(_font(9, bold=True))
            btn_edit.clicked.connect(lambda _=False, idx=i: self._on_edit(idx))
            self.table.setCellWidget(row, 3, btn_edit)
            btn_del = QPushButton("삭제")
            btn_del.setCursor(Qt.PointingHandCursor)
            btn_del.setFocusPolicy(Qt.NoFocus)
            btn_del.setFixedSize(54, 26)
            btn_del.setObjectName("RowDeleteBtn")
            btn_del.setFont(_font(9, bold=True))
            btn_del.clicked.connect(lambda _=False, idx=i: self._on_delete(idx))
            self.table.setCellWidget(row, 4, btn_del)

    def _on_add(self) -> None:
        dlg = _AddMemberDialog(self)
        if dlg.exec_() != QDialog.Accepted:
            return
        data = dlg.get_data()
        result = api_add_user(data["user_id"], data["password"], data["name"])
        if result.get("success"):
            QMessageBox.information(self, "완료", "회원이 추가되었습니다.")
            self.load_users()
        else:
            QMessageBox.warning(self, "실패", result.get("message") or "회원 추가에 실패했습니다.")

    def _on_edit(self, row_idx: int) -> None:
        if not (0 <= row_idx < len(self._users)):
            return
        u = self._users[row_idx]
        target_user_id = u.get("user_id") or ""
        target_name = u.get("name") or ""
        admin_id = self._admin_user_id()
        if not admin_id:
            QMessageBox.warning(self, "권한 없음", "로그인된 관리자 정보를 찾을 수 없습니다.")
            return
        dlg = _EditMemberDialog(target_user_id, target_name, self)
        if dlg.exec_() != QDialog.Accepted:
            return
        data = dlg.get_data()
        result = api_admin_update_user(
            admin_id,
            data["admin_password"],
            data["target_user_id"],
            new_name=data.get("new_name"),
            new_password=data.get("new_password"),
        )
        if result.get("success"):
            QMessageBox.information(self, "완료", "회원 정보가 수정되었습니다.")
            self.load_users()
        else:
            QMessageBox.warning(self, "실패", result.get("message") or "회원 수정에 실패했습니다.")

    def _on_delete(self, row_idx: int) -> None:
        if not (0 <= row_idx < len(self._users)):
            return
        u = self._users[row_idx]
        target_user_id = u.get("user_id") or ""
        admin_id = self._admin_user_id()
        if not admin_id:
            QMessageBox.warning(self, "권한 없음", "로그인된 관리자 정보를 찾을 수 없습니다.")
            return
        dlg = _DeleteMemberDialog(target_user_id, self)
        if dlg.exec_() != QDialog.Accepted:
            return
        admin_password = dlg.get_admin_password()
        result = api_delete_user(admin_id, admin_password, target_user_id)
        if result.get("success"):
            QMessageBox.information(self, "완료", "회원이 삭제되었습니다.")
            self.load_users()
        else:
            QMessageBox.warning(self, "실패", result.get("message") or "회원 삭제에 실패했습니다.")


class AdminScreen(QWidget):
    """관리 탭 메인 화면(내부 사이드 메뉴 + 화면 전환)."""

    def __init__(self, db_connection=None, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.db_connection = db_connection
        self._build_ui()

    def _build_ui(self) -> None:
        self.setObjectName("AdminRoot")
        self.setStyleSheet(
            """
            QWidget#AdminRoot {
                background-color: #F8FAFC;
            }

            QFrame#AdminSide {
                background-color: #FFFFFF;
                border: 1px solid #F1F5F9;
                border-radius: 14px;
            }

            QFrame#AdminSide * {
                outline: none;
            }

            QPushButton#AdminNavBtn {
                text-align: left;
                padding-left: 14px;
                border-radius: 12px;
                border: none;
                color: #334155;
                background: transparent;
            }
            QPushButton#AdminNavBtn:hover {
                background: #F8FAFC;
            }
            QPushButton#AdminNavBtn:checked {
                background: #E0F2FE;
                color: #2563EB;
                font-weight: 900;
            }

            /* 학생관리 상단 버튼/검색: 텍스트 진하게 + 테두리 진하게 */
            QPushButton#primary {
                background: #2563EB;
                color: #FFFFFF;
                border: 1px solid #1D4ED8;
                font-weight: 900;
                border-radius: 12px;
                padding: 10px 14px;
            }
            QPushButton#primary:hover {
                background: #1D4ED8;
                border: 1px solid #1E40AF;
            }
            QPushButton#secondary {
                background: #FFFFFF;
                color: #0F172A;
                border: 1px solid #475569;
                font-weight: 900;
                border-radius: 12px;
                padding: 10px 14px;
            }
            QPushButton#secondary:hover {
                background: #F8FAFC;
                border: 1px solid #334155;
            }

            QLineEdit#StudentSearchBox {
                background: #FFFFFF;
                color: #0F172A;
                border: 1px solid #475569;
                border-radius: 12px;
                padding: 10px 12px;
                font-weight: 800;
            }
            QLineEdit#StudentSearchBox:focus {
                border: 2px solid #2563EB;
            }

            /* 학년 필터칩: 기본 텍스트/테두리 더 진하게 */
            QPushButton#FilterChip {
                background: #FFFFFF;
                border: 1px solid #64748B;
                border-radius: 999px;
                padding: 8px 12px;
                color: #334155;
                font-weight: 900;
            }
            QPushButton#FilterChip:hover {
                background: #F8FAFC;
                border: 1px solid #475569;
            }
            QPushButton#FilterChip:checked {
                background: #F0F7FF;
                border: 1px solid #2563EB;
                color: #2563EB;
                font-weight: 900;
            }

            /* 표 내부 액션 버튼(수정/삭제) 가독성 */
            QPushButton#RowEditBtn {
                border: 1px solid #2563EB;
                background: #EFF6FF;
                color: #2563EB;
                border-radius: 10px;
                padding: 4px 8px;
                font-weight: 900;
            }
            QPushButton#RowEditBtn:hover {
                background: #DBEAFE;
            }

            QPushButton#RowDeleteBtn {
                border: 1px solid #EF4444;
                background: #FEF2F2;
                color: #DC2626;
                border-radius: 10px;
                padding: 4px 8px;
                font-weight: 900;
            }
            QPushButton#RowDeleteBtn:hover {
                background: #FEE2E2;
            }

            QWidget#AdminContentCard {
                background-color: #FFFFFF;
                border: 1px solid #F1F5F9;
                border-radius: 14px;
            }
            """
        )

        outer = QHBoxLayout(self)
        outer.setContentsMargins(24, 16, 24, 24)
        outer.setSpacing(16)

        side = QFrame()
        side.setObjectName("AdminSide")
        # ✅ 사이드바 폭을 줄여 본문을 더 넓게 확보
        side.setFixedWidth(180)
        s_lay = QVBoxLayout(side)
        s_lay.setContentsMargins(10, 10, 10, 10)
        s_lay.setSpacing(8)

        self.btn_students = _NavButton("학생 관리")
        self.btn_classes = _NavButton("반 관리")
        self.btn_students.setChecked(True)

        s_lay.addWidget(self.btn_students)
        s_lay.addWidget(self.btn_classes)
        s_lay.addStretch(1)

        outer.addWidget(side, 0)

        content_card = QFrame()
        content_card.setObjectName("AdminContentCard")
        c_lay = QVBoxLayout(content_card)
        c_lay.setContentsMargins(18, 16, 18, 16)
        c_lay.setSpacing(0)

        self.view_students = StudentManagementView(self.db_connection)
        self.view_classes = ClassManagementView(self.db_connection)
        c_lay.addWidget(self.view_students)
        c_lay.addWidget(self.view_classes)
        self.view_classes.hide()

        outer.addWidget(content_card, 1)

        def _show_students():
            self.btn_students.setChecked(True)
            self.btn_classes.setChecked(False)
            self.view_students.show()
            self.view_classes.hide()

        def _show_classes():
            self.btn_students.setChecked(False)
            self.btn_classes.setChecked(True)
            self.view_students.hide()
            self.view_classes.show()

        self.btn_students.clicked.connect(_show_students)
        self.btn_classes.clicked.connect(_show_classes)

