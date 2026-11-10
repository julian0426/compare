@echo off
chcp 65001 > nul
cd /d "%~dp0"
py -3 -m venv .venv
if errorlevel 1 (
  echo Python 3가 필요합니다. 이 파일은 개발/빌드 PC에서만 실행합니다.
  pause
  exit /b 1
)
call .venv\Scripts\activate
python -m pip install --upgrade pip
pip install pandas openpyxl xlrd pyinstaller
pyinstaller --noconfirm --clean --onefile --windowed --name "회원별_ERP_마스터_금액대조" member_reconciler.py
if errorlevel 1 (
  echo EXE 생성 실패
  pause
  exit /b 1
)
echo 생성 완료: dist\회원별_ERP_마스터_금액대조.exe
pause
