# word-revision

Add Word Track Changes revision marks to .docx files while preserving existing formatting.

## 应用场景

`word-revision` 适合需要把文本修改以 Word 修订痕迹（Track Changes）方式写入 `.docx` 的场景，典型用途包括：

- **合同/协议修订**：把审查意见、条款修改直接生成带修订痕迹的 Word 稿，供客户或对方律师逐条确认。
- **法律意见书/律师函改稿**：在保留原有排版、样式、页眉页脚的前提下，标记新增、删除或替换内容。
- **批量文书处理**：对多个 DOCX 执行统一的文本替换、条款插入或段落重写，并统一输出为修订稿。
- **协同审校**：让 Agent 代为执行标准化文字修改，同时保留人工复核所需的修订记录。

核心优势：不重建文档，直接在 OOXML 层操作，保留原始格式、样式、表格、页眉页脚等结构。

## 安装方法

### 方式一：通过 skills 包管理器安装（推荐）

```bash
npx skills install LawyerYuan/word-revision -g
```

`-g` 表示全局安装；若只想在当前项目生效，可去掉 `-g`。

### 方式二：手动克隆到 skills 目录

**Claude Code：**

```bash
git clone https://github.com/LawyerYuan/word-revision.git ~/.claude/skills/word-revision
```

**Codex：**

```bash
git clone https://github.com/LawyerYuan/word-revision.git ~/.codex/skills/word-revision
```

**其他兼容 skills 机制的 Agent：**

将本仓库克隆到对应 Agent 的 skills 目录即可，例如 `~/<agent>/skills/word-revision/`。

## 使用

在支持 skills 的 Agent 中，直接说明你要使用 `word-revision` 处理任务，或显式引用：

```text
Use $word-revision to ...
```

## 文件结构

- `SKILL.md`: skill 的触发说明和执行流程
- `scripts/`: 可复用脚本，如有
- `references/`: 按需读取的参考资料，如有
- `assets/`: 模板或资源文件，如有

## 注意

发布前应检查仓库中是否包含密钥、token、客户材料、个人绝对路径、日志、缓存或临时输出文件。公开仓库尤其需要先完成敏感信息扫描。
