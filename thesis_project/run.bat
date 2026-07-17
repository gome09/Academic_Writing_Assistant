@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

echo ============================================================
echo   论文 Word 草案 ^& 答辩 PPT 草案 生成器
echo ============================================================
echo.

REM ---- 定位 Python：优先 py，其次 python ----
set "PY="
where py >nul 2>nul && set "PY=py"
if not defined PY (
    where python >nul 2>nul && set "PY=python"
)
if not defined PY (
    echo [错误] 未找到 Python，请先安装 Python 3 并加入 PATH。
    echo         下载: https://www.python.org/downloads/
    pause
    exit /b 1
)
echo [信息] 使用解释器: %PY%
echo.

REM ---- 检查并按需安装依赖 ----
%PY% -c "import docx, pptx, pdfplumber" >nul 2>nul
if errorlevel 1 (
    echo [信息] 缺少依赖，正在安装 python-docx / python-pptx / pdfplumber ...
    %PY% -m pip install python-docx python-pptx pdfplumber
    if errorlevel 1 (
        echo [错误] 依赖安装失败，请检查网络或手动执行:
        echo         %PY% -m pip install python-docx python-pptx pdfplumber
        pause
        exit /b 1
    )
)

REM ---- 选择输入：input\ 有文件就用 input\，否则用示例 ----
set "INPUT=input"
dir /b /a-d "input\*" >nul 2>nul
if errorlevel 1 (
    echo [信息] input\ 为空，改用 sample_input\ 示例演示。
    echo         ^(把你的 Word/PDF/TXT/md/json 放进 input\ 后重跑即可^)
    set "INPUT=sample_input"
) else (
    echo [信息] 读取 input\ 下的源文件。
)
echo.

REM ---- 运行主程序 ----
%PY% -u src\main.py --input "%INPUT%" --output output
set "RC=%errorlevel%"
echo.

if "%RC%"=="0" (
    echo ============================================================
    echo   完成！草案已生成到 output\ 目录：
    echo     - output\论文草案.docx
    echo     - output\答辩PPT草案.pptx
    echo   提示：Word 打开后按 F9 更新目录；替换 ^<请填写^> 占位符。
    echo ============================================================
    REM 自动打开 output 目录
    start "" "%cd%\output"
) else (
    echo [错误] 生成失败，退出码 %RC%。请查看上方日志。
)

echo.
pause
endlocal
