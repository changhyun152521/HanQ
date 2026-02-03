# HanQ 데스크톱 exe 빌드 스크립트 (배포용)
# - 배포용 config(config.deploy.json)를 적용한 뒤 PyInstaller 실행
# - 빌드 후 로컬 개발용 config.json 복원
# 사용: 반드시 프로젝트 루트(CH_LMS)에서 실행
#   PowerShell: cd C:\Users\이창현\Desktop\CH_LMS  후  .\scripts\build_exe.ps1
#   CMD: cd /d C:\Users\이창현\Desktop\CH_LMS  후  powershell -ExecutionPolicy Bypass -File scripts\build_exe.ps1

$ErrorActionPreference = "Stop"
# 한글 경로 인코딩 이슈 방지: 스크립트 위치 대신 현재 디렉터리 사용 (반드시 프로젝트 루트에서 실행)
$null = Get-Location | Out-Null

# 상대 경로만 사용 (현재 디렉터리가 프로젝트 루트여야 함)
$ConfigJson = "config\config.json"
$DeployJson = "config\config.deploy.json"
$BackupJson = "config\config.json.bak"

# 1) 로컬 config.json 백업 후 배포용으로 교체
if (-not (Test-Path -LiteralPath $DeployJson)) {
    Write-Error "config.deploy.json 이 없습니다. config 폴더에 config.deploy.json 파일을 넣어 주세요."
}
Copy-Item -LiteralPath $ConfigJson -Destination $BackupJson -Force
Copy-Item -LiteralPath $DeployJson -Destination $ConfigJson -Force
Write-Host "[1/3] 배포용 설정 적용 완료 (config.deploy.json -> config.json)"

try {
    # 2) PyInstaller 실행 (한글 경로 대응: Python UTF-8 모드 사용)
    Write-Host "[2/3] PyInstaller 빌드 중..."
    $env:PYTHONUTF8 = "1"
    & python -X utf8 -m PyInstaller HanQ.spec
    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller 실패 (exit $LASTEXITCODE)"
    }
    Write-Host "[2/3] 빌드 완료. 결과: dist\HanQ\"

    # 배포 시 첫 실행이 로그인 화면이 되도록 세션 파일 제거 (다른 PC에서 관리자 자동 로그인 방지)
    $SessionInDist = "dist\HanQ\config\.hanq_session.json"
    if (Test-Path -LiteralPath $SessionInDist) {
        Remove-Item -LiteralPath $SessionInDist -Force
        Write-Host "배포 패키지에서 세션 파일 제거됨 (첫 실행 시 로그인 화면 표시)"
    }
} finally {
    # 3) 로컬 config.json 복원
    if (Test-Path -LiteralPath $BackupJson) {
        Copy-Item -LiteralPath $BackupJson -Destination $ConfigJson -Force
        Remove-Item -LiteralPath $BackupJson -Force
        Write-Host "[3/3] 로컬 설정 복원 완료 (config.json)"
    }
}

Write-Host ""
Write-Host "exe 위치: dist\HanQ\HanQ.exe"
Write-Host "배포 시 dist\HanQ 폴더 전체를 압축해 배포하면 됩니다."
