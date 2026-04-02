@echo off
cd C:\Users\Administrator\.openclaw\workspace

echo ===== 检查远程仓库配置 =====
git remote -v

echo.
echo ===== 添加今天的新内容 =====
git add -A

echo.
echo ===== 提交更改 =====
git commit -m "2026-04-02 更新: Claw Code架构学习、learning-notes目录、权限配置"

echo.
echo ===== 推送到远程 =====
git push origin main

echo.
echo ===== 完成 =====
pause
