@echo off
title Docx Standardizer Server
color 0B

echo ========================================================
echo        KHOI DONG SERVER DOCX STANDARDIZER
echo ========================================================
echo.
echo Dang di chuyen den thu muc du an...
cd /d "E:\docx_standardizer"

echo.
echo Dang chay may chu tai dia chi: http://localhost:8002
echo De tat may chu, hay an dau X hoac bam CTRL + C.
echo.

python -m uvicorn app:app --host 127.0.0.1 --port 8002 --reload

echo.
pause
