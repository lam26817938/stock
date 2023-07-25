@echo off

REM 設定變數
set duration=30

REM 同時執行第一個 Python 檔案
python stock.py %duration%

REM 同時執行第二個 Python 檔案
start /B python tpex.py %duration%
timeout /T 1 /NOBREAK >nul

REM 同時執行第三個 Python 檔案
start /B python yf.py %duration%
timeout /T 1 /NOBREAK >nul

REM 執行第四個 Python 檔案
python comb.py

REM 按下任意鍵繼續...
pause >nul