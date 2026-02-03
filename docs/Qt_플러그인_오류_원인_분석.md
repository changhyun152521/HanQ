# "no Qt platform plugin could be initialized" 원인 분석

## 1. 오류 의미

- Qt가 **Windows용 플랫폼 플러그인**(`qwindows.dll`)을 **찾지 못했거나**, 찾았지만 **로드/초기화에 실패**했다는 뜻입니다.

## 2. 분석한 원인 (가능성 순)

### 2-1. PyInstaller 공식 `pyi_rth_qt5.py`가 우리 설정을 덮어씀 (가장 유력)

- **현재**: 런타임 훅 순서 = `pyi_rth_qt5_hanq.py`(우리) → `pyi_rth_qt5.py`(PyInstaller 공식).
- **문제**: 공식 훅은 **one-file** 빌드용으로 `sys._MEIPASS` 같은 경로에 `QT_PLUGIN_PATH`를 맞춥니다.  
  우리는 **one-folder**라 exe 옆 `dist\HanQ\` 기준으로 경로를 잡고 있는데, 공식 훅이 **그 다음에** 실행되면서 우리가 설정한 값을 **덮어쓸** 수 있습니다.
- **결과**: exe 실행 시 Qt가 잘못된(또는 존재하지 않는) 경로를 보게 되어 "platform plugin could not be initialized" 발생.

**대응**: 공식 `pyi_rth_qt5.py`를 **runtime_hooks에서 제거**하고, one-folder 전용으로 우리 훅만 사용.

---

### 2-2. `qwindows.dll` 로드 시 의존 DLL을 찾지 못함

- `qwindows.dll`은 **Qt5Core.dll** 등에 의존합니다.
- one-folder에서는 이 DLL들이 `dist\HanQ\PyQt5\Qt5\` 등 **하위 폴더**에 들어갈 수 있는데, Windows는 플러그인 DLL을 로드할 때 **그 DLL이 있는 폴더**와 **exe가 있는 폴더** 위주로 검색합니다.  
  하위 폴더만 있는 경우 검색 순서에 따라 **못 찾을 수** 있습니다.
- **결과**: 플러그인 경로는 맞는데, "플러그인을 초기화할 수 없음"으로 나올 수 있습니다.

**대응**: 런타임 훅에서 **exe 기준 폴더**와 **PyQt5/Qt5** 경로를 `os.add_dll_directory()`로 추가해, DLL 검색 경로에 포함.

---

### 2-3. 플러그인/폴더가 번들에 없음

- 한글 경로 대응 훅이 **빌드 시** 플러그인 **소스** 경로만 패치하고, **수집 대상/목적지**는 PyInstaller 기본 로직(`qt_rel_dir` 등)에 맡기고 있습니다.
- 이론상 `dist\HanQ\PyQt5\Qt5\plugins\platforms\` 에 `qwindows.dll`이 들어가야 하나,  
  PyInstaller/훅 버전·환경에 따라 **수집이 누락**되거나 **다른 경로**에 들어갈 수 있습니다.
- **확인 방법**: 빌드 후 `dist\HanQ\PyQt5\Qt5\plugins\platforms\` (또는 `PyQt5\Qt\plugins\platforms\`) 폴더와 `qwindows.dll` 존재 여부 확인.

**대응**: 위 2-1, 2-2 적용 후에도 오류가 나면, `dist\HanQ` 구조를 캡처해 두고 `--collect-all PyQt5` 등으로 플러그인 강제 수집 여부 검토.

---

### 2-4. 런타임 훅이 실행되지 않음

- `_script_dir`이 `getcwd()`로 잡히는데, 빌드 시 **작업 디렉터리가 프로젝트 루트가 아니면** `pyi_rth_qt5_hanq.py` 경로가 틀려서 훅이 번들에 안 들어갈 수 있습니다.
- **대응**: 빌드는 **항상 프로젝트 루트**(CH_LMS)에서 `.\scripts\build_exe.ps1` 실행.  
  스펙의 `_script_dir` fallback은 이미 `getcwd()`로 되어 있음.

---

## 3. 적용한 수정 요약

1. **HanQ.spec**  
   - PyInstaller 공식 `pyi_rth_qt5.py`를 **runtime_hooks에서 제거**.  
   - one-folder 전용으로 `pyi_rth_qt5_hanq.py`만 사용.

2. **pyinstaller_hooks/pyi_rth_qt5_hanq.py**  
   - `QT_PLUGIN_PATH`, `QT_QPA_PLATFORM_PLUGIN_PATH` 설정은 유지.  
   - **`os.add_dll_directory()`**로 exe 기준 폴더와 `PyQt5\Qt5`(및 `PyQt5\Qt`) 경로를 DLL 검색 경로에 추가.  
   - `qwindows.dll`이 Qt5Core.dll 등을 찾을 수 있도록 함.

3. **검증 순서**  
   - 수정 후 **다시 빌드** → `dist\HanQ\HanQ.exe` 실행.  
   - 여전히 오류 시 `dist\HanQ\PyQt5\Qt5\plugins\platforms\` 존재 여부 확인.
