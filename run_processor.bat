@echo off
chcp 65001 >nul
title 帮会联赛数据处理程序

:menu
cls
echo ========================================
echo        帮会联赛数据处理程序
echo ========================================
echo.
echo 请选择运行模式：
echo.
echo 1. GUI模式 - 通过文件选择对话框选择CSV文件
echo 2. 命令行模式 - 手动输入CSV文件路径
echo 3. 使用默认文件 (把文件名改为banghuiliansai.csv并放到我们项目的根目录)
echo 4. 退出程序
echo.
set /p choice=请输入选择 (1-4): 

if "%choice%"=="1" goto gui_mode
if "%choice%"=="2" goto cli_mode
if "%choice%"=="3" goto default_mode
if "%choice%"=="4" goto exit
echo 无效选择，请重新输入
pause
goto menu

:gui_mode
echo.
echo 启动GUI模式...
python guild_league_processor_advanced.py
goto end

:cli_mode
echo.
set /p csv_file=请输入CSV文件路径: 
if "%csv_file%"=="" (
    echo 未输入文件路径，返回主菜单
    pause
    goto menu
)
echo.
echo 启动命令行模式，处理文件: %csv_file%
python guild_league_processor_advanced.py "%csv_file%"
goto end

:default_mode
echo.
echo 使用默认文件...
if exist "banghuiliansai.csv" (
    python guild_league_processor_advanced.py "banghuiliansai.csv"
) else (
    echo 错误：找不到默认文件 banghuiliansai.csv
    echo 请确保文件存在于当前目录中
)
goto end

:end
echo.
echo 处理完成！
pause
goto menu

:exit
echo 程序退出
pause 