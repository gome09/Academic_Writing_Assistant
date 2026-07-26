@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

echo ============================================================
echo   论文 Word 草案 ^& 答辩 PPT 草案 生成器
echo ============================================================
echo.

REM ---- 定位引导 Python，并创建项目隔离环境 ----
set "BOOTSTRAP="
where py >nul 2>nul && set "BOOTSTRAP=py"
if not defined BOOTSTRAP (
    where python >nul 2>nul && set "BOOTSTRAP=python"
)
if not defined BOOTSTRAP (
    echo [错误] 未找到 Python，请先安装 Python 3 并加入 PATH。
    echo         下载: https://www.python.org/downloads/
    pause
    exit /b 1
)
if not exist ".venv\Scripts\python.exe" (
    echo [信息] 正在创建项目虚拟环境 .venv ...
    %BOOTSTRAP% -m venv .venv
    if errorlevel 1 exit /b 1
)
set "PY=.venv\Scripts\python.exe"
echo [信息] 使用解释器: %PY%
echo.

REM ---- 检查并按需安装依赖 ----
%PY% -c "import docx, pptx, pdfplumber, openpyxl, yaml" >nul 2>nul
if errorlevel 1 (
    echo [信息] 缺少核心依赖，正在从 requirements.lock 安装 ...
    %PY% -m pip install -r requirements.lock
    if errorlevel 1 (
        echo [错误] 依赖安装失败，请检查网络或手动执行:
        echo         %PY% -m pip install -r requirements.lock
        pause
        exit /b 1
    )
)
if defined LLM_API_KEY (
    %PY% -c "import openai" >nul 2>nul
    if errorlevel 1 %PY% -m pip install -r requirements-llm.lock
)

REM ---- 选择输入：input\ 有文件就用 input\，否则用示例 ----
set "INPUT=input"
dir /b /a-d "input\*" >nul 2>nul
if errorlevel 1 (
    echo [信息] input\ 为空，改用 sample_input\ 示例演示。
    echo         ^(把 Word/PDF/TXT/Markdown/JSON/Excel/图片放进 input\ 后重跑即可^)
    set "INPUT=sample_input"
) else (
    echo [信息] 读取 input\ 下的源文件。
)
echo.

REM ---- 运行主程序 ----
REM ---- 透传额外参数：编辑本文件在本行后加 --llm 等开关 ----
%PY% -u src\main.py --input "%INPUT%" --output output %*
set "RC=%errorlevel%"
echo.

if "%RC%"=="0" (
    echo ============================================================
    echo   完成！请查看 output\ 目录中的实际产物与运行报告。
    echo   普通模式通常生成 Word + PPT；参考资料模式只生成 Word。
    echo   提示：Word 打开后按提示更新域（或按 F9 更新目录）；替换 ^<请填写^> 占位符。
    echo ============================================================
    REM 自动打开 output 目录
    start "" "%cd%\output"
) else (
    echo [错误] 生成失败，退出码 %RC%。请查看上方日志。
)

echo.
pause
endlocal
