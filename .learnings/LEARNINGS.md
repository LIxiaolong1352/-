## [LRN-20250312-001] best_practice

**Logged**: 2026-03-12T21:50:00+08:00
**Priority**: high
**Status**: resolved
**Area**: config

### Summary
Android App构建：GitHub Actions配置复杂易错，Codemagic更简单可靠

### Details
本次尝试用GitHub Actions构建Android APK，遇到多个问题：
1. 缺少根目录build.gradle和settings.gradle
2. gradlew权限问题（Windows上传后丢失执行权限）
3. 缺少gradle.properties（AndroidX配置）
4. 文件结构多次出错（上传到子文件夹而非根目录）

最终使用Codemagic成功构建，配置更简单，不需要处理gradlew权限问题。

### Suggested Action
- 以后Android项目直接用Codemagic构建
- 使用gradle命令而非gradlew
- settings.gradle必须包含仓库配置（google(), mavenCentral()）
- 项目结构：app文件夹、build.gradle、settings.gradle、gradle.properties缺一不可

### Metadata
- Source: user_feedback
- Related Files: codemagic.yaml, build.gradle, settings.gradle
- Tags: android, build, codemagic, github-actions

### Resolution
- **Resolved**: 2026-03-12T21:50:00+08:00
- **Solution**: 使用Codemagic替代GitHub Actions，配置更简洁
- **Time Saved**: 预计下次可节省3+小时

---

## [LRN-20250312-003] knowledge_gap

**Logged**: 2026-03-12T23:59:00+08:00
**Priority**: high
**Status**: pending
**Area**: finance

### Summary
需要学习更多加密货币技术分析方法

### Details
用户要求分析ETH走势，涉及：
- 支撑阻力位判断
- 成交量分析（CVD指标）
- 资金费率解读
- 持仓量分析
- 多空比判断
- 杠杆仓位风险管理

用户反馈："小弟学习不错，不要忘记了，多学点分析技术"

### Suggested Action
- 学习更多技术分析指标（RSI、MACD、布林带、EMA等）
- 学习加密货币市场特殊指标（资金费率、爆仓数据、持仓量）
- 学习风险管理（杠杆计算、仓位管理、止损策略）
- 记录常见交易场景和应对策略

### Metadata
- Source: user_feedback
- Related Files: scripts/crypto/
- Tags: crypto, trading, technical-analysis, risk-management

---

## [LRN-20250312-002] best_practice

**Logged**: 2026-03-12T21:50:00+08:00
**Priority**: medium
**Status**: resolved
**Area**: infra

### Summary
GitHub文件上传：用户容易把文件传到子文件夹而非根目录

### Details
用户多次尝试上传文件到GitHub，但总是传到子文件夹（如"文件夹2"）或导致文件结构散开。

### Suggested Action
- 以后直接给用户提供zip文件
- 让用户上传zip到GitHub
- 配置workflow自动解压
- 避免手动逐个文件上传

### Metadata
- Source: error
- Related Files: GitHub上传流程
- Tags: github, upload, workflow

---
