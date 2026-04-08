; RG_webview Inno Setup 스크립트
; 사용법:
;   1. PyInstaller 빌드 완료 후 dist\ClassMate\ 폴더가 있어야 합니다.
;      pyinstaller RG_webview.spec
;   2. Inno Setup Compiler에서 이 파일을 열고 Build → Compile
;   3. Output\RG_webview_Setup.exe 가 완성된 인스톨러입니다.
;
; Inno Setup 다운로드: https://jrsoftware.org/isdl.php

#define AppName      "ClassMate"
#define AppVersion   "1.0.0"
#define AppPublisher "ReadingGate"
#define AppExeName   "ClassMate.exe"
#define SourceDir    "dist\ClassMate"

[Setup]
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL=
DefaultDirName={autopf}\{#AppName}
DisableProgramGroupPage=yes
OutputDir=Output
OutputBaseFilename=ClassMate_Setup
SetupIconFile=ClassMate.ico
UninstallDisplayIcon={app}\{#AppExeName}
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
; 한국어 설치 화면
ShowLanguageDialog=no

; 아키텍처 — 64비트 전용
ArchitecturesInstallIn64BitMode=x64compatible
ArchitecturesAllowed=x64compatible

[Languages]
Name: "korean"; MessagesFile: "compiler:Languages\Korean.isl"

[Tasks]
; 바탕화면/시작메뉴 생략

[Files]
; PyInstaller 빌드 결과 폴더 전체 포함
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

; 아이콘 파일 (언인스톨 표시용)
Source: "RG_webview.ico"; DestDir: "{app}"; Flags: ignoreversion

[Run]
; 설치 완료 후 앱 바로 실행 옵션
Filename: "{app}\{#AppExeName}"; Description: "{cm:LaunchProgram,{#AppName}}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; 앱이 생성한 설정/이력 파일 제거 (선택)
; 사용자 데이터는 건드리지 않음 — work_root는 외부 폴더라 안전
Type: files; Name: "{app}\app_config.json"
Type: files; Name: "{app}\run_error.log"
