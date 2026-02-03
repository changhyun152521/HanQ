# -*- coding: utf-8 -*-
# 한글 경로에서 PyInstaller가 Qt 플러그인 경로를 잘못 찾는 문제 우회
# 현재 프로세스에서 플러그인 경로를 직접 계산해 QtLibraryInfo에 패치 후 기본 수집 로직 사용

import os

from PyInstaller.utils.hooks.qt import add_qt_dependencies, get_qt_library_info


def _get_pyqt5_plugins_path():
    """한글 경로 대응: 현재 프로세스에서 PyQt5 플러그인 경로 계산"""
    try:
        import PyQt5
        base = os.path.dirname(os.path.abspath(PyQt5.__file__))
        for sub in ('Qt5', 'Qt'):
            path = os.path.join(base, sub, 'plugins')
            if os.path.isdir(path):
                return path
    except Exception:
        pass
    return None


# QtLibraryInfo 로드 유도 후 플러그인 경로 패치 (한글 경로 대응)
qt_info = get_qt_library_info('PyQt5')
_ = getattr(qt_info, 'version', None)  # _load_qt_info 트리거
_plugins_path = _get_pyqt5_plugins_path()
if _plugins_path and getattr(qt_info, 'location', None):
    qt_info.location['PluginsPath'] = _plugins_path

hiddenimports, binaries, datas = add_qt_dependencies(__file__)
