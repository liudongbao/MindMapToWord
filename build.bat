@echo off
chcp 65001 >nul
echo ===============================================
echo  MindMapToWord - 构建与发布脚本
echo ===============================================
echo.

set PROJECT_NAME=MindMapToWord
set OUTPUT_DIR=%PROJECT_NAME%\bin\Release\net8.0-windows\win-x64\publish
set SINGLE_FILE_NAME=MindMapToWord_Single.exe

echo [1/3] 清理旧的发布文件...
if exist "%OUTPUT_DIR%" (
    rmdir /s /q "%OUTPUT_DIR%"
    echo   已清理旧发布目录
)

echo.
echo [2/3] 编译项目...
dotnet build "%PROJECT_NAME%" -c Release
if %errorlevel% neq 0 (
    echo   编译失败！
    pause
    exit /b %errorlevel%
)
echo   编译成功

echo.
echo [3/3] 发布单文件可执行程序...
dotnet publish "%PROJECT_NAME%" -c Release -r win-x64 --self-contained false -p:PublishSingleFile=true -p:PublishTrimmed=true -o "%OUTPUT_DIR%"

if %errorlevel% neq 0 (
    echo   发布失败！
    pause
    exit /b %errorlevel%
)

rem 重命名单文件版本
ren "%OUTPUT_DIR%\%PROJECT_NAME%.exe" "%SINGLE_FILE_NAME%"

echo   发布成功！
echo.
echo ===============================================
echo 输出位置：
echo   普通发布：%PROJECT_NAME%\bin\Release\net8.0-windows\%PROJECT_NAME%.exe
echo   单文件版：%OUTPUT_DIR%\%SINGLE_FILE_NAME%
echo ===============================================

rem 自动打开单文件版本
echo.
echo 正在启动单文件版本...
start "" "%OUTPUT_DIR%\%SINGLE_FILE_NAME%"

pause