@echo off
chcp 65001 >nul
echo ========================================
echo   AI Copilot - Word 智能写作助手
echo ========================================
echo.
echo 正在启动开发服务器...
echo.
echo 提示：
echo 1. 此窗口将启动开发服务器（保持运行）
echo 2. 请在另一个终端运行: npm start
echo 3. 或者手动打开 Word 并加载插件
echo.
echo ========================================
echo.

npm run dev-server
