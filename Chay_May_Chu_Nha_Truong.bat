@echo off
title HET HONG DOCX STANDARDIZER (LAN)
color 0A

echo ========================================================
echo   KHOI DONG SERVER CHO MANG NOI BO (MAY NHA TRUONG)
echo ========================================================
echo.
cd /d "E:\docx_standardizer"

echo May chu dang khoi dong... 
echo Nguoi khac co the truy cap qua IP cua may nay (Vi du: http://192.168.x.x:8002)
echo.

:: --host 0.0.0.0 cho phep cac may tinh khac trong mang LAN truy cap
python -m uvicorn app:app --host 0.0.0.0 --port 8002

pause
