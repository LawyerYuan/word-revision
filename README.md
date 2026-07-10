# word-revision

Add Word Track Changes revision marks to .docx files while preserving existing formatting.

## 用途

该仓库提供一个可复用的 AI Agent skill。将其安装到兼容的 skills 目录后，Agent 可以根据 `SKILL.md` 中的流程执行对应任务。

## 安装

以 Codex 为例：

```bash
git clone https://github.com/LawyerYuan/word-revision.git ~/.codex/skills/word-revision
```

如果你使用其他兼容 skills 机制的 Agent，请放入对应的 skills 目录。

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
