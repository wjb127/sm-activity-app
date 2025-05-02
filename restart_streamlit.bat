@echo off
echo 현재 실행 중인 Streamlit 서버를 중지합니다...
taskkill /f /im streamlit.exe /t
taskkill /f /im python.exe /fi "WINDOWTITLE eq streamlit" /t
timeout /t 2 /nobreak > nul

echo Streamlit 캐시 지우는 중...
rd /s /q "%USERPROFILE%\.streamlit\cache" 2>nul

echo Streamlit 서버를 재시작합니다...
start cmd /k "cd /d %~dp0 && python -m streamlit run sm_activity_app.py"

echo 잠시 후 브라우저에서 http://localhost:8502 로 접속하세요. 