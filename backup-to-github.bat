# 备份 OpenClaw 工作区到 GitHub

# 1. 进入工作区
cd C:\Users\Administrator\.openclaw\workspace

# 2. 检查 git 状态
git status

# 3. 添加所有文件
git add -A

# 4. 提交
git commit -m "Backup 2026-04-02 - 包含Claw Code学习笔记和配置"

# 5. 添加远程仓库（替换为你的GitHub用户名）
git remote add origin https://github.com/你的用户名/openclaw-backup.git

# 6. 推送
git push -u origin main

# 如果推送失败，先创建仓库：
# 访问 https://github.com/new 创建名为 openclaw-backup 的仓库
