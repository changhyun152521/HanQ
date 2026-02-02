"""
HanQ 메인 진입점

이 파일은 프로그램의 시작점으로, GUI 애플리케이션을 초기화하고 실행합니다.
- 메인 윈도우 생성 및 표시
- 이벤트 루프 시작
- 전역 설정 로드
"""
import sys
import os
from pathlib import Path
from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
from ui.main_window import MainWindow
from ui.theme import apply_theme


def main():
    """메인 함수"""
    # High DPI 스케일링 및 픽스맵 선명도 (QApplication 생성 전에 설정)
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    # Qt 플러그인 경로 설정 (Windows에서 플러그인을 찾지 못하는 문제 해결)
    if sys.platform == 'win32':
        try:
            import PyQt5
            pyqt5_path = Path(PyQt5.__file__).parent
            plugin_path = pyqt5_path / 'Qt5' / 'plugins'
            if plugin_path.exists():
                os.environ['QT_PLUGIN_PATH'] = str(plugin_path)
        except Exception:
            pass

    app = QApplication(sys.argv)
    apply_theme(app)
    
    # 애플리케이션 정보 설정
    app.setApplicationName("HanQ")
    app.setOrganizationName("HanQ")
    
    # 메인 윈도우 생성 및 표시
    window = MainWindow()
    window.show()
    
    # 이벤트 루프 시작
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
