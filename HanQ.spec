# -*- mode: python ; coding: utf-8 -*-
# HanQ 데스크톱 exe 빌드용 PyInstaller 스펙
# 사용: 빌드 스크립트 실행 후 pyinstaller HanQ.spec

import os
import glob

block_cipher = None

# 한글 경로 대응: 커스텀 훅 사용 (pyinstaller_hooks/hook-PyQt5.QtWidgets.py)
# PyInstaller 실행 시 __file__이 없을 수 있음 → getcwd() 사용 (빌드는 프로젝트 루트에서 실행)
try:
    _script_dir = os.path.dirname(os.path.abspath(__file__))
except NameError:
    _script_dir = os.getcwd()
_hookpath = os.path.join(_script_dir, 'pyinstaller_hooks')

# Qt5 런타임 훅: one-folder 전용 HanQ 훅만 사용
# (공식 pyi_rth_qt5는 one-file용이라 우리가 설정한 QT_PLUGIN_PATH를 덮어써서 제거)
_rth_hanq = os.path.join(_script_dir, 'pyinstaller_hooks', 'pyi_rth_qt5_hanq.py')
_runtime_hooks = [_rth_hanq] if os.path.isfile(_rth_hanq) else []

# PyQt5 플랫폼 플러그인 수동 수집 (exe 실행 시 "no Qt platform plugin" 방지)
_qt5_plugin_binaries = []
try:
    import PyQt5
    _pyqt5_base = os.path.dirname(os.path.abspath(PyQt5.__file__))
    for _sub in ('Qt5', 'Qt'):
        _platforms_src = os.path.join(_pyqt5_base, _sub, 'plugins', 'platforms')
        if os.path.isdir(_platforms_src):
            _dst = os.path.join('PyQt5', _sub, 'plugins', 'platforms')
            for _f in glob.glob(os.path.join(_platforms_src, '*.dll')):
                _qt5_plugin_binaries.append((_f, _dst))
            break
except Exception:
    pass

# 배포용 config는 빌드 스크립트에서 config/config.json으로 복사된 상태로 실행
a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=_qt5_plugin_binaries,
    datas=[
        ('config', 'config'),       # config.json, tag_schema.json
        ('templates', 'templates'), # 워크시트 템플릿 등
        ('fonts', 'fonts'),         # Pretendard, GmarketSans, NanumGothic 등 UI/템플릿용 글꼴
    ],
    hiddenimports=[
        'PyQt5.QtCore',
        'PyQt5.QtGui',
        'PyQt5.QtWidgets',
        'requests',
        'pymongo',
        'openpyxl',
        'matplotlib',
        'matplotlib.backends.backend_qt5agg',
    ],
    hookspath=[_hookpath] if os.path.isdir(_hookpath) else [],
    hooksconfig={},
    runtime_hooks=_runtime_hooks,
    excludes=['auth_api'],  # Node 서버는 번들 제외
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='HanQ',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # GUI 앱이므로 콘솔 숨김
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='HanQ',
)
