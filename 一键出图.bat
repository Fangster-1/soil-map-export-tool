@echo off
chcp 65001 >nul
cd /d "%~dp0"

REM ============================================================
REM  必须使用含 arcpy 的 ArcGIS Pro Python，普通 python 会报
REM  "No module named 'arcpy'"。按优先级依次探测可用解释器。
REM ============================================================
set "PYEXE="
for %%P in (
  "D:\GD\arcgispro_clone\python.exe"
  "C:\Program Files\ArcGIS\Pro\bin\Python\envs\arcgispro-py3\python.exe"
  "%LOCALAPPDATA%\Programs\ArcGIS\Pro\bin\Python\envs\arcgispro-py3\python.exe"
) do (
  if not defined PYEXE if exist "%%~P" set "PYEXE=%%~P"
)

if not defined PYEXE (
  echo [错误] 未找到含 arcpy 的 ArcGIS Pro Python 解释器。
  echo 请编辑本 bat，把 PYEXE 指向你的 ArcGIS Pro python.exe。
  pause
  exit /b 1
)

echo 使用解释器: %PYEXE%
"%PYEXE%" soil_map_export_tool.py
pause
