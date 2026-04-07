# Installation Guide - 大模型驱动的代理记账行业专属SKILL

基于大语言模型的代理记账行业垂直Skill。完整的安装指南。

---

## Table of Contents

- [Quick Start](#quick-start)
- [Claude Code Native Marketplace](#claude-code-native-marketplace-recommended)
- [Universal Installer](#universal-installer)
- [Manual Installation](#manual-installation)
- [Verification & Testing](#verification--testing)
- [Troubleshooting](#troubleshooting)
- [Uninstallation](#uninstallation)

---

## Quick Start

### For Claude Code Users (Recommended)

```bash
npx skills add cyhzzz/Bookkeeping-Agency-Skill
```

### For All Other Agents (Cursor, VS Code, etc.)

```bash
npx ai-agent-skills install cyhzzz/Bookkeeping-Agency-Skill
```

---

## Claude Code Native Marketplace (Recommended)

### Step 1: Install the Skill

```bash
npx skills add cyhzzz/Bookkeeping-Agency-Skill
```

### Step 2: Verify Installation

```bash
# Skill 已安装，直接说"初始化"开始配置
```

### Update Skill

```bash
npx skills update cyhzzz/Bookkeeping-Agency-Skill
```

### Remove Skill

```bash
npx skills remove Bookkeeping-Agency-Skill
```

---

## Universal Installer

通用安装器使用 [ai-agent-skills](https://github.com/skillcreatorai/Ai-Agent-Skills) 包。

### Install to All Supported Agents

```bash
npx ai-agent-skills install cyhzzz/Bookkeeping-Agency-Skill
```

### Install to Specific Agent

```bash
# Claude Code only
npx ai-agent-skills install cyhzzz/Bookkeeping-Agency-Skill --agent claude

# Cursor only
npx ai-agent-skills install cyhzzz/Bookkeeping-Agency-Skill --agent cursor
```

---

## Manual Installation

用于开发、定制或离线使用。

### Prerequisites

- **Git**
- **Claude Code** (for using skills)
- **Python 3.7+** (for report generation scripts)

### Step 1: Clone Repository

```bash
git clone https://github.com/cyhzzz/Bookkeeping-Agency-Skill.git
cd Bookkeeping-Agency-Skill
```

### Step 2: Copy to Agent Directory

#### For Claude Code

```bash
cp -r Bookkeeping-Agency-Skill ~/.claude/skills/
```

#### For Cursor

```bash
mkdir -p .cursor/skills
cp -r Bookkeeping-Agency-Skill .cursor/skills/
```

#### For VS Code/Copilot

```bash
mkdir -p .github/skills
cp -r Bookkeeping-Agency-Skill .github/skills/
```

---

## Verification & Testing

### Verify Installation

```bash
# Check that skill is present
ls ~/.claude/skills/Bookkeeping-Agency-Skill/
```

### First Use: Run Initialization

```
说"初始化"启动配置向导

AI：欢迎使用代理记账行业SKILL初始化向导！
    请提供公司信息、Logo、当地实情...
```

---

## Troubleshooting

### Issue: "Skills not showing in Claude Code"

**Solution:** Verify installation and restart

```bash
# Check installation
ls -la ~/.claude/skills/

# Verify SKILL.md exists
cat ~/.claude/skills/Bookkeeping-Agency-Skill/SKILL.md

# Restart Claude Code
```

### Issue: "Python module not found"

**Solution:** Install dependencies

```bash
pip install python-docx pandas openpyxl
```

---

## Uninstallation

### Claude Code

```bash
npx skills remove Bookkeeping-Agency-Skill
```

### Universal Installer

```bash
# Remove from Claude Code
rm -rf ~/.claude/skills/Bookkeeping-Agency-Skill/

# Remove from Cursor
rm -rf .cursor/skills/Bookkeeping-Agency-Skill/
```

### Manual Installation

```bash
# Remove from Claude Code
rm -rf ~/.claude/skills/Bookkeeping-Agency-Skill/
```

---

## Support

**Installation Issues?**
- Open issue: https://github.com/cyhzzz/Bookkeeping-Agency-Skill/issues

---

**Last Updated:** 2026-04-07
**Skill Version:** 1.1.0
