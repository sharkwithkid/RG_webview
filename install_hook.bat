@echo off
:: install_hook.bat — git post-push hook 설치
:: 프로젝트 루트에서 실행하세요.

echo [훅 설치] post-push hook 설치 중...

if not exist ".git\hooks" (
    echo [오류] .git\hooks 폴더가 없습니다. 프로젝트 루트에서 실행하세요.
    pause
    exit /b 1
)

copy /Y "post-push" ".git\hooks\post-push" > nul
if %errorlevel% neq 0 (
    echo [오류] 훅 파일 복사 실패.
    pause
    exit /b 1
)

echo [완료] .git\hooks\post-push 설치됨
echo        이제 main 브랜치에 push 하면 자동으로 빌드됩니다.
pause
