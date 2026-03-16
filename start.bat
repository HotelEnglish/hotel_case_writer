@echo off
title Streamlit App Launcher
echo ==========================================
echo 正在启动 Streamlit 应用...
echo 目标文件: app_ui.py
echo ==========================================

:: 检查当前目录是否有 app_ui.py
if not exist "app_ui.py" (
    echo [错误] 在当前目录下未找到 'app_ui.py' 文件！
    echo 请确保此 bat 脚本放在与 app_ui.py 相同的文件夹内。
    pause
    exit /b
)

:: 激活虚拟环境 (如果你有虚拟环境，请修改下一行，否则可以删除或注释掉)
:: 例如: call venv\Scripts\activate.bat

:: 运行 Streamlit
streamlit run app_ui.py

:: 防止窗口在程序意外退出时立即关闭
pause