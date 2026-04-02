---
name: ai-research-writer
description: AI 学术研究写作助手。提供论文写作全流程的 Prompt 模板和最佳实践，包括中英互译、润色、去 AI 味、逻辑检查、图表生成、实验分析、审稿视角审视等。适用于 NeurIPS/ICML/ICLR/ACL 等顶会论文写作。使用场景：(1) 中文草稿翻译为英文学术论文 (2) 英文论文润色与降重 (3) 去除 AI 写作痕迹 (4) 生成论文架构图和实验图表 (5) 实验结果分析 (6) 投稿前审稿视角检查。
---

# AI 学术研究写作助手

本 Skill 整合了来自 MSRA、Seed、SH AI Lab 等顶尖研究机构以及北大、中科大、上交硕博同学的实战写作技巧。

## 核心功能

### 1. 翻译与转换
- **中转英**: 将中文草稿翻译为符合顶会标准的英文学术论文
  - 参考: `references/zh-to-en.md`
- **英转中**: 将英文 LaTeX 翻译为流畅中文，便于快速理解
  - 参考: `references/en-to-zh.md`
- **中转中**: 将口语化中文重写为规范学术中文（Word 适配）
  - 参考: `references/zh-refine-zh.md`

### 2. 文本优化
- **缩写**: 在不损失信息的前提下压缩文本长度（-5~15 词）
  - 参考: `references/condense.md`
- **扩写**: 增强逻辑连接，使文本更饱满（+5~15 词）
  - 参考: `references/expand.md`
- **表达润色（英文）**: 深度润色，达到零错误出版水准
  - 参考: `references/polish-en.md`
- **表达润色（中文）**: 修复语病，保持作者原有风格
  - 参考: `references/polish-zh.md`

### 3. 质量检查
- **逻辑检查**: 终稿红线审查，检查致命错误
  - 参考: `references/logic-check.md`
- **去 AI 味（LaTeX 英文）**: 去除机械化表达，接近母语水平
  - 参考: `references/humanize-en.md`
- **去 AI 味（Word 中文）**: 去除机器味和翻译腔
  - 参考: `references/humanize-zh.md`

### 4. 图表与可视化
- **论文架构图**: 生成顶会风格的架构图 Prompt
  - 参考: `references/figure-architecture.md`
- **实验绘图推荐**: 从 19 种标准学术图表中选择最佳方案
  - 参考: `references/figure-recommendation.md`
- **图标题生成**: 撰写规范的英文图标题
  - 参考: `references/caption-figure.md`
- **表标题生成**: 撰写规范的英文表标题
  - 参考: `references/caption-table.md`

### 5. 实验与审稿
- **实验分析**: 从实验数据中挖掘关键结论，生成 LaTeX 分析段落
  - 参考: `references/experiment-analysis.md`
- **审稿视角审视**: 以严苛审稿人角度审查论文，发现致命缺陷
  - 参考: `references/reviewer-perspective.md`

## 使用建议

1. **写作流程**: 中转英 → 润色 → 去 AI 味 → 逻辑检查
2. **图表流程**: 架构图 → 实验绘图推荐 → 生成标题
3. **投稿前**: 实验分析 → 审稿视角审视

## 模型推荐

- 日常 idea 交互与论文写作: Gemini-2.5-pro/flash
- 实验代码编写: Claude-4.5 系列

所有 Prompt 均经过一线科研人员实战打磨，开箱即用。
