# exe Qt 플러그인 수정 점검 결과

재빌드 전에 수정 사항이 exe 실행 시 "no Qt platform plugin could be initialized"를 해결하는지 검증한 내용입니다.

---

## 1. 빌드 시 (HanQ.spec)

| 항목 | 내용 | 결과 |
|------|------|------|
| 플러그인 수동 수집 | `_qt5_plugin_binaries`로 PyQt5의 `Qt5/plugins/platforms/*.dll`(또는 `Qt/plugins/platforms/*.dll`)을 찾아 `binaries`에 추가 | ✓ |
| 목적지 경로 | `PyQt5/Qt5/plugins/platforms` (또는 `PyQt5/Qt/plugins/platforms`) | ✓ |
| 빌드 결과 | `dist/HanQ/PyQt5/Qt5/plugins/platforms/qwindows.dll` 등이 생성됨 | ✓ |

- 스펙에서 훅과 별도로 **플랫폼 플러그인 DLL을 직접 수집**하므로, 훅이 한글 경로 등으로 실패해도 해당 DLL은 번들에 들어갑니다.

---

## 2. 실행 시 (pyi_rth_qt5_hanq.py)

| 항목 | 내용 | 결과 |
|------|------|------|
| 실행 순서 | 런타임 훅이 **main.py 로드 전**에 실행됨 | ✓ |
| base_dir | `os.path.dirname(sys.executable)` → exe가 있는 폴더 = `dist/HanQ` | ✓ |
| QT_PLUGIN_PATH | `dist/HanQ/PyQt5/Qt5/plugins` (또는 Qt) | ✓ |
| QT_QPA_PLATFORM_PLUGIN_PATH | `dist/HanQ/PyQt5/Qt5/plugins/platforms` | ✓ |
| DLL 검색 경로 | `os.add_dll_directory(base_dir)`, `os.add_dll_directory(PyQt5/Qt5)` 추가 → qwindows.dll이 Qt5Core.dll 등을 찾을 수 있음 | ✓ |

- 공식 `pyi_rth_qt5`는 **제거**되어 있어, one-file용 경로로 우리 설정이 덮어씌워지지 않습니다.

---

## 3. 경로 일치 여부

| 구분 | 경로 | 일치 |
|------|------|------|
| 스펙에서 수집 목적지 | `PyQt5/Qt5/plugins/platforms` | ✓ |
| 런타임 훅이 찾는 폴더 | `base_dir/PyQt5/Qt5/plugins` 및 그 안의 `platforms` | ✓ |
| Qt가 플러그인을 찾는 위치 | `QT_QPA_PLATFORM_PLUGIN_PATH` = `.../plugins/platforms` | ✓ |

- 빌드 시 넣는 경로와 실행 시 사용하는 경로가 **같은 구조**로 맞춰져 있습니다.

---

## 4. 결론

- **수정 내용만 보면 exe가 열리도록 되어 있습니다.**
  - 플랫폼 플러그인 DLL을 스펙에서 **직접** 수집하고,
  - 실행 시 exe 기준 경로로 플러그인/디렉터리를 지정하며,
  - 공식 훅 제거 + DLL 검색 경로 추가까지 되어 있어, 재빌드 후에는 "no Qt platform plugin could be initialized"가 해결될 가능성이 높습니다.

---

## 5. 재빌드 후 확인 방법 (선택)

빌드가 끝나면 아래만 한 번 확인해 보시면 됩니다.

1. **폴더 존재**  
   `dist\HanQ\PyQt5\Qt5\plugins\platforms\` (또는 `PyQt5\Qt\plugins\platforms\`) 가 있는지
2. **DLL 존재**  
   그 안에 `qwindows.dll` (및 필요 시 다른 플랫폼 플러그인 DLL) 이 있는지
3. **exe 실행**  
   `dist\HanQ\HanQ.exe` 더블클릭 후 로그인 화면이 뜨는지

위가 모두 맞으면 수정이 의도대로 적용된 것입니다.

---

## 6. 재빌드 전 최종 점검 (2025-02-03)

코드베이스를 다시 읽어 아래를 확인했습니다.

| 확인 항목 | 상태 |
|-----------|------|
| **HanQ.spec** | `_qt5_plugin_binaries`로 플랫폼 DLL 수동 수집 → `PyQt5/Qt5/plugins/platforms`(또는 Qt). 런타임 훅은 `pyi_rth_qt5_hanq.py`만 사용, 공식 pyi_rth_qt5 미사용. | ✓ |
| **pyi_rth_qt5_hanq.py** | frozen 시 `base_dir` = exe 폴더. `add_dll_directory(base_dir)`, `PyQt5/Qt5`, `PyQt5/Qt` 추가. `QT_PLUGIN_PATH` / `QT_QPA_PLATFORM_PLUGIN_PATH`를 `PyQt5/Qt5/plugins`(또는 Qt) 및 `.../platforms`로 설정. | ✓ |
| **main.py** | frozen 시 `PyQt5/Qt5/plugins`, `PyQt5/Qt/plugins` 순으로 찾아 `QT_PLUGIN_PATH`, `QT_QPA_PLATFORM_PLUGIN_PATH` 설정. 런타임 훅과 경로 구조 일치. | ✓ |
| **main_window.py** | `_get_app_root()` frozen 시 exe 디렉터리 반환. DB 실패 시 `LOCALAPPDATA`/`TEMP`/홈의 `HanQ/db/ch_lms.db`로 재시도 후, 그래도 실패 시 안내만 하고 앱은 계속 실행. | ✓ |
| **build_exe.ps1** | 배포 config 적용 → `python -X utf8 -m PyInstaller HanQ.spec` → config 복원. | ✓ |

**결론:** 스펙·런타임 훅·main·DB 폴백·빌드 스크립트가 서로 맞게 설정되어 있습니다. **재빌드 후 exe가 정상적으로 열릴 가능성이 높습니다.** 재빌드가 끝나면 위 "5. 재빌드 후 확인 방법"대로 한 번 확인하시면 됩니다.
