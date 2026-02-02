"""
App-wide modern UI theme (QSS)

Goal: Reduce visual noise (borders), increase padding, unify radius,
and rely on surfaces + soft shadows.
"""

from __future__ import annotations

from PyQt5.QtWidgets import QApplication
from PyQt5.QtGui import QFont


MODERN_QSS = """
/* ===== Base ===== */
* {
  font-family: 'Pretendard','NanumGothic','Malgun Gothic','맑은 고딕',sans-serif;
}

QLabel {
  border: none;
}

QMainWindow, QWidget {
  background-color: #F8FAFC;
  color: #0F172A;
}

QScrollArea, QStackedWidget {
  background: transparent;
  border: none;
}

/* Scrollbars (quiet) */
QScrollBar:vertical {
  background: transparent;
  width: 10px;
  margin: 0px;
}
QScrollBar::handle:vertical {
  background-color: rgba(148,163,184,0.55);
  border-radius: 5px;
  min-height: 28px;
}
QScrollBar::handle:vertical:hover {
  background-color: rgba(100,116,139,0.65);
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
  height: 0px;
}

/* ===== Surfaces (border-less) ===== */
QFrame#ConfigCard,
QFrame#PreviewCanvas,
QFrame#FinalConfig {
  background-color: #FFFFFF;
  border: 1px solid #F1F5F9; /* minimal */
  border-radius: 12px;
}

/* ===== Typography ===== */
QLabel#PageTitle {
  color: #0F172A;
}
QLabel#CardTitle {
  color: #0F172A;
}

/* ===== Inputs ===== */
QLineEdit, QSpinBox, QPlainTextEdit {
  background-color: #F8FAFC;
  border: 1px solid #E2E8F0;
  border-radius: 12px;
  padding: 10px 12px;
}
QLineEdit:focus, QSpinBox:focus, QPlainTextEdit:focus {
  background-color: #FFFFFF;
  border: 2px solid #2563EB;
}

/* ===== Buttons ===== */
QPushButton {
  border-radius: 12px;
  padding: 10px 14px;
  border: 1px solid #E2E8F0; /* outline by default */
  background: transparent;
  color: #0F172A;
}
QPushButton:hover {
  background-color: #F1F5F9;
}

/* Strong outline (source pick) */
QPushButton#SourcePickBtn {
  border: 2px solid #2563EB;
  color: #2563EB;
  background: #FFFFFF;
  font-weight: 800;
}
QPushButton#SourcePickBtn:hover {
  background: #F0F7FF;
}

QPushButton#GenerateBtn,
QPushButton#FinalizeBtn {
  border: none;
  background-color: #2563EB;
  color: #FFFFFF;
  font-weight: 800;
}
QPushButton#GenerateBtn:hover,
QPushButton#FinalizeBtn:hover {
  background-color: #1D4ED8;
}

QPushButton#BackBtn,
QPushButton#CloseBtn {
  background: transparent;
  border: 1px solid #E2E8F0;
  color: #334155;
}

/* Small text button */
QPushButton#OpenHwpTextBtn {
  border: none;
  background: transparent;
  color: #2563EB;
  padding: 4px 6px;
}
QPushButton#OpenHwpTextBtn:hover {
  color: #1D4ED8;
  text-decoration: underline;
}

/* ===== Tree/List ===== */
QTreeWidget {
  background: #FFFFFF;
  border: 1px solid #F1F5F9;
  border-radius: 12px;
}
QTreeWidget::item {
  padding: 10px;
  height: 32px;
}
QTreeWidget::item:hover {
  background: #F8FAFC;
}

QListWidget {
  background: transparent;
  border: none;
}

/* ===== Segmented buttons (Source) ===== */
QFrame#SegmentWrap {
  background: #F1F5F9;
  border: 1px solid #F1F5F9;
  border-radius: 12px;
}
QPushButton#SegmentBtnLeft, QPushButton#SegmentBtnRight {
  border: none;
  border-radius: 12px;
  padding: 10px 0px;
  color: #334155;
  background: transparent;
}
QPushButton#SegmentBtnLeft:checked, QPushButton#SegmentBtnRight:checked {
  background: #2563EB;
  color: #FFFFFF;
  font-weight: 800;
}

/* Chips */
QPushButton#SourceChip {
  background: #FFFFFF;
  border: 1px solid #F1F5F9;
  border-radius: 16px;
  padding: 8px 12px;
  color: #0F172A;
}
QPushButton#SourceChip:hover:enabled {
  background: #F8FAFC;
}
QPushButton#SourceChip:disabled {
  background: #F1F5F9;
  color: #94A3B8;
}

/* Selected tags */
QFrame#SelectedTag {
  background: #2563EB;
  border: none;
  border-radius: 14px;
}
QLabel#SelectedTagText {
  color: #FFFFFF;
  font-weight: 800;
}
QPushButton#SelectedTagClose {
  border: none;
  background: rgba(255,255,255,0.22);
  color: #FFFFFF;
  border-radius: 8px;
  padding: 0px;
}
QPushButton#SelectedTagClose:hover {
  background: rgba(255,255,255,0.32);
}

/* Divider line (very subtle) */
QFrame#DividerLine {
  background: #F1F5F9;
  min-height: 1px;
  max-height: 1px;
}

/* ===== Card 3 (details) clean/minimal ===== */
QFrame#ConfigCard[cardRole="details"] QLabel#SectionLabel {
  color: #475569;
  font-weight: 800;
}
QFrame#ConfigCard[cardRole="details"] QLabel#SectionHint {
  color: #64748B;
}

/* Remove grey fills inside details card */
QFrame#ConfigCard[cardRole="details"] QLineEdit {
  background: #FFFFFF;
  border: 1px solid #F1F5F9;
}

/* Count inline (slider + number) integrated feel */
QFrame#ConfigCard[cardRole="details"] QFrame#CountInline {
  background: transparent;
  border: none;
}
QFrame#ConfigCard[cardRole="details"] QFrame#CountInline QSpinBox {
  background: transparent;
  border: none;
  padding: 0px;
  min-width: 72px;
}
QFrame#ConfigCard[cardRole="details"] QFrame#CountInline QSpinBox::up-button,
QFrame#ConfigCard[cardRole="details"] QFrame#CountInline QSpinBox::down-button {
  width: 0px;
  height: 0px;
}

/* Filter chips (grade/type/order) */
QPushButton#FilterChip {
  background: #FFFFFF;
  border: 1px solid #E2E8F0;
  border-radius: 999px;
  padding: 8px 12px;
  color: #64748B;
}
QPushButton#FilterChip:hover {
  background: #F8FAFC;
}
QPushButton#FilterChip:checked {
  background: #F0F7FF;
  border: 1px solid #2563EB;
  color: #2563EB;
  font-weight: 800;
}

/* Selected list container */
QScrollArea#SelectedContainer {
  background: #F8FAFC;
  border: none;
  border-radius: 12px;
}
QLabel#EmptyHint {
  color: #94A3B8;
}
"""


def apply_theme(app: QApplication) -> None:
    try:
        app.setStyle("Fusion")
    except Exception:
        pass
    app.setStyleSheet(MODERN_QSS)
    # 폰트 렌더링 선명도 (뭉개짐 방지)
    try:
        font = app.font()
        font.setStyleStrategy(QFont.PreferAntialias)
        app.setFont(font)
    except Exception:
        pass

