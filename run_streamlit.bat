@echo off
cd /d "%~dp0"  REM 현재 스크립트가 위치한 디렉토리로 이동
python -m streamlit run sm_activity_app.py  REM Streamlit을 사용하여 SM Activity 앱 시작
pause  REM 실행 후 창이 닫히지 않도록 대기
