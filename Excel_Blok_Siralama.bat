@echo off
chcp 65001 > nul
cd /d "%~dp0"
python excel_block_sorter_gui.py
pause
