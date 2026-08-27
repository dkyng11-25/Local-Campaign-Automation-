@echo off
setlocal

cd /d "C:\Users\KEARNEY\Desktop\Local_Campaign_Automation_version7"

echo.
echo ========================================
echo Local Campaign Automation
echo ========================================
echo.
echo Streamlit server starting...
echo URL: http://localhost:8501
echo.

start "Local Campaign Streamlit Server" /MIN ^
".venv\Scripts\python.exe" -m streamlit run "streamlit_app.py" ^
--server.address 0.0.0.0 ^
--server.port 8501 ^
--server.headless true

timeout /t 4 /nobreak >nul

start "" "http://localhost:8501"

echo.
echo Streamlit server started.
echo Browser opened.
echo.
echo Team URL:
echo http://192.168.11.132:8501
echo.
pause