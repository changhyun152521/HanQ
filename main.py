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
from PyQt5.QtGui import QFont, QFontDatabase
from ui.main_window import MainWindow
from ui.theme import apply_theme


def _app_base_dir() -> str:
    """앱 기준 디렉터리: 배포 시 exe 폴더, 개발 시 프로젝트 루트."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def _load_bundled_fonts() -> None:
    """fonts 폴더의 TTF/OTF를 Qt에 등록해 앱 전역에서 사용 가능하게 합니다."""
    fonts_dir = os.path.join(_app_base_dir(), "fonts")
    if not os.path.isdir(fonts_dir):
        return
    for name in os.listdir(fonts_dir):
        if name.lower().endswith((".ttf", ".otf")):
            path = os.path.join(fonts_dir, name)
            try:
                QFontDatabase.addApplicationFont(path)
            except Exception:
                pass


def main():
    """메인 함수"""
    # Qt 플러그인 경로 설정 (QApplication 생성 전에 반드시 설정)
    # PyInstaller로 빌드된 exe 실행 시: exe와 같은 폴더 기준으로 플러그인 경로 지정
    if getattr(sys, 'frozen', False) and sys.platform == 'win32':
        base_dir = os.path.dirname(sys.executable)
        for sub in ('PyQt5/Qt5/plugins', 'PyQt5/Qt/plugins'):
            plugin_path = os.path.join(base_dir, sub)
            if os.path.isdir(plugin_path):
                plugin_path = os.path.abspath(plugin_path)
                os.environ['QT_PLUGIN_PATH'] = plugin_path
                platforms_path = os.path.join(plugin_path, 'platforms')
                if os.path.isdir(platforms_path):
                    os.environ['QT_QPA_PLATFORM_PLUGIN_PATH'] = platforms_path
                break
    elif sys.platform == 'win32':
        try:
            import PyQt5
            pyqt5_path = Path(PyQt5.__file__).parent
            plugin_path = pyqt5_path / 'Qt5' / 'plugins'
            if not plugin_path.exists():
                plugin_path = pyqt5_path / 'Qt' / 'plugins'
            if plugin_path.exists():
                os.environ['QT_PLUGIN_PATH'] = str(plugin_path)
        except Exception:
            pass

    # High DPI 스케일링 및 픽스맵 선명도 (QApplication 생성 전에 설정)
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    app = QApplication(sys.argv)
    _load_bundled_fonts()
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
