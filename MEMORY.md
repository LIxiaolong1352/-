# MEMORY.md - 长期记忆

## 已安装的技能

### stock-analysis v6.2.0
- 股票和加密货币分析
- 支持投资组合、观察列表、股息分析
- 8维度评分、热门扫描、传闻检测
- 安装路径: openclaw/skills/stock-analysis

### multi-search-engine v2.0.1
- 17个搜索引擎（8国内+9国际）
- 无需API密钥
- 安装路径: openclaw/skills/multi-search-engine

## 系统配置

### 记忆系统
- Daily notes: `memory/YYYY-MM-DD.md`
- 长期记忆: `MEMORY.md`
- 每次学习新东西都要记录

### 重装系统后恢复方法
- OpenClaw 安装在 `C:\Users\Administrator\AppData\Roaming\npm\node_modules\openclaw`
- 工作目录在 `C:\Users\Administrator\.openclaw\workspace`
- 恢复步骤：
  1. 重装Node.js
  2. 运行 `npm install -g openclaw`
  3. 恢复 workspace 文件夹（从备份或GitHub克隆）
  4. 重新安装 skills: `npx clawhub install find-skills stock-analysis multi-search-engine weather playwright self-improvement skill-creator healthcheck`
- 重要：定期备份 `workspace` 文件夹到GitHub或其他地方

### 用户承诺
- "以后不会忘了你" - 2026-03-13
- 用户重视我的分析，作为投资参考
- 需要持续学习提升专业度

## 用户偏好
- 需要我记住每天学习的内容
- 不要重复工作
