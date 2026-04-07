# Claude Skills Marketplace

## Marketplace Metadata

**Name**: 大模型驱动的代理记账行业专属SKILL
**Description**: 基于大语言模型的代理记账行业垂直Skill，智能调度5大模块，支持初始化定制
**Version**: 1.1.0
**Author**: cyhzzz
**License**: MIT
**Repository**: https://github.com/cyhzzz/Bookkeeping-Agency-Skill

---

## Available Skills

### 代理记账行业垂直Skill包

#### Bookkeeping-Agency-Skill

**Path**: `/`
**Description**: 大模型驱动的代理记账行业专属SKILL - 智能调度5大模块，支持初始化定制。管理层初始化后交付员工使用。
**Version**: 1.1.0
**Install**: `npx skills add cyhzzz/Bookkeeping-Agency-Skill`

**Features**:
- 初始化向导（品牌标识 + 县域逻辑）
- 意图识别与模块调度
- 5个子Skill协同工作
- 带品牌标识的Word报告生成
- 财务软件数据解析（用友/金蝶/管家婆）

**Sub-Skills**:
- `client-onboarding-helper`: 新客户入职全流程
- `monthly-tax-helper`: 月度做账与报税
- `complex-business-advisor`: 复杂业务咨询（出口退税/高新认定/研发费用）
- `risk-monitor-guard`: 风险监控与政策追踪
- `customer-success-helper`: 客户成功管理

**Trigger**: "初始化" / "新客户入职" / "处理本月票据" / "检查风险"

---

## Installation

### One-Line Install

```bash
npx skills add cyhzzz/Bookkeeping-Agency-Skill
```

### Install via Universal Installer

```bash
npx ai-agent-skills install cyhzzz/Bookkeeping-Agency-Skill
```

---

## Quick Reference

| 场景 | 触发方式 |
|------|---------|
| 初始化配置 | 说"初始化" |
| 新客户入职 | 说"新客户入职" |
| 月度做账 | 说"处理本月票据" |
| 风险检查 | 说"检查风险" |
| 客户价值报告 | 说"给客户出一份价值报告" |

---

## Development

### Adding New Sub-Skills

1. Create sub-skill directory under root
2. Add `SKILL.md` with proper frontmatter
3. Register in main `SKILL.md` dispatch table
4. Update version in `config/skill-config.yaml`
5. Create git tag

### Skill Structure

```
Bookkeeping-Agency-Skill/
├── SKILL.md              # 主控Skill（入口）
├── INSTALLATION.md       # 安装指南
├── MARKETPLACE.md        # 市场文件
├── README.md             # 项目说明
├── CLAUDE.md             # 开发指南
├── config/              # 初始化配置
├── assets/              # 报告模板
├── parsers/             # 财务软件解析器
├── generators/          # 报告生成器
└── [sub-skill]/         # 5个子Skill
```

### SKILL.md Frontmatter Template

```yaml
---
name: skill-name
description: 简短描述（<50字符）
version: "1.1.0"
last_updated: "2026-04-07"
---
```

---

## Dependencies

### Python Dependencies

For report generation:
```bash
pip install python-docx pandas openpyxl
```

---

## Support

- **Issues**: https://github.com/cyhzzz/Bookkeeping-Agency-Skill/issues
- **Documentation**: See README.md

---

## License

MIT License

---

**Last Updated**: 2026-04-07
**Skill Version**: 1.1.0
**Total Sub-Skills**: 5
