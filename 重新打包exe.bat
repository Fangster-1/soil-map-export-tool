@echo off
chcp 65001 >nul
cd /d "%~dp0"
REM ============================================================
REM  一键重新打包单文件 exe（含图标 + 内嵌主程序）。
REM  修改了 soil_map_export_tool20260616.py 或 launcher.py 后运行本脚本即可。
REM ============================================================
set "PYEXE=D:\GD\arcgispro_clone\python.exe"

echo [1/2] PyInstaller 打包中...
"%PYEXE%" -m PyInstaller --noconfirm --clean --onefile --noconsole ^
  --icon "app.ico" ^
  --add-data "soil_map_export_tool.py;." ^
  --name "土壤类型出图工具" launcher.py
if errorlevel 1 ( echo 打包失败！ & pause & exit /b 1 )

echo [2/2] 拷贝到发布目录...
if not exist "出图工具_发布版(单文件)" mkdir "出图工具_发布版(单文件)"
copy /Y "dist\土壤类型出图工具.exe" "出图工具_发布版(单文件)\" >nul

echo 完成！exe 位于 dist\ 和 出图工具_发布版(单文件)\
echo 注意：GitHub Release 中的 exe 需要手动更新（gh release upload）
pause
