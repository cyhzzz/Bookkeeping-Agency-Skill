# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

AI代账工具 - 支持初始化定制的代账公司智能化工作台。
管理层初始化时配置品牌标识和县域逻辑，交付员工后直接使用。

## 项目结构

```
代理记账/                          ← 主控Skill（根目录即skill）
├── SKILL.md                       ← 主控Skill（双模式：初始化+工作）
├── config/
│   └── skill-config.yaml        ← 初始化配置和版本追踪
├── assets/
│   └── report-templates/        ← 报告模板（含品牌占位符）
├── parsers/                      ← 财务软件解析器
├── generators/                   ← 报告生成器
├── client-onboarding-helper/    ← 子Skill（可定制）
├── monthly-tax-helper/          ← 子Skill（可定制）
├── complex-business-advisor/      ← 子Skill（可定制）
├── risk-monitor-guard/          ← 子Skill（可定制）
└── customer-success-helper/      ← 子Skill（可定制）
```

## 初始化流程（管理层操作）

说"初始化"启动配置向导，依次完成：

### Step 1: 品牌信息
- 公司名称、简称
- 联系电话、地址
- Logo上传

### Step 2: 县域实情
- 客户行业分布
- 税务特点（查账/核定）
- 特殊优惠政策

### Step 3: 输出格式
- 报告内容定制
- 语言风格（正式/简洁/详细）
- 文件命名规范

### Step 4: 确认节点
- 自动化程度（高/标准/严格）
- 特殊业务复核要求

## 员工使用

初始化完成后，员工直接说需求，AI自动识别意图并调度模块。

## 版本追踪

每个子Skill在frontmatter中标注：
```yaml
version: "1.1.0"
last_updated: "2026-04-07"
```

版本号规则：
- 主版本+1：工作流程/输出格式变更
- 子版本+1：触发词/确认节点变更

## 关键约束

1. **初始化后才能交付** - 首次使用必须完成初始化配置
2. **品牌标识自动更新** - Logo和联系信息自动应用到Word模板
3. **双人复核** - 复杂业务必须人工复核
4. **双格式输出** - MD存档 + Word交付（含品牌标识）
