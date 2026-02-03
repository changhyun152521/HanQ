# -*- coding: utf-8 -*-
# HanQ exe 실행 시 Qt 플러그인 경로 + DLL 검색 경로 설정 (main.py보다 먼저 실행됨)
# PyInstaller runtime hook: 메인 스크립트 로드 전에 실행

import os
import sys

if getattr(sys, 'frozen', False) and sys.platform == 'win32':
    base_dir = os.path.dirname(sys.executable)
    base_dir_abs = os.path.abspath(base_dir)

    # DLL 검색 경로에 exe 폴더 및 PyQt5 하위 폴더 추가 (qwindows.dll이 Qt5Core.dll 등을 찾을 수 있도록)
    if hasattr(os, 'add_dll_directory'):
        try:
            os.add_dll_directory(base_dir_abs)
        except OSError:
            pass
        for sub in (('PyQt5', 'Qt5'), ('PyQt5', 'Qt')):
            dll_dir = os.path.join(base_dir_abs, *sub)
            if os.path.isdir(dll_dir):
                try:
                    os.add_dll_directory(dll_dir)
                except OSError:
                    pass

    # Qt 플러그인 경로 설정
    for sub in ('PyQt5', 'Qt5', 'plugins'), ('PyQt5', 'Qt', 'plugins'):
        plugin_path = os.path.join(base_dir, *sub)
        if os.path.isdir(plugin_path):
            os.environ['QT_PLUGIN_PATH'] = os.path.abspath(plugin_path)
            platforms_path = os.path.join(plugin_path, 'platforms')
            if os.path.isdir(platforms_path):
                os.environ['QT_QPA_PLATFORM_PLUGIN_PATH'] = os.path.abspath(platforms_path)
            break
