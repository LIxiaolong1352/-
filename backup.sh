#!/bin/bash
# OpenClaw Workspace 自动备份脚本

cd /c/Users/Administrator/.openclaw/workspace

# 添加所有文件
git add .

# 提交（带时间戳）
git commit -m "Auto backup: $(date '+%Y-%m-%d %H:%M:%S')"

# 推送到GitHub
git push origin main 2>/dev/null || git push origin master 2>/dev/null

echo "Backup completed at $(date)"
