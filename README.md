# outline-to-editable-ppt

生图式可编辑 PPT：一个用于 Codex 的工作流 skill。从共同确认的逐页大纲与资产出发，生成整页视觉稿，再忠实重建可编辑 PowerPoint，并从同一讲稿源同步幻灯片备注和独立 Word 讲稿。

## 适用场景

- 科研汇报、组会、课程展示和内容汇报。
- 需要先确认代表页视觉，再分部分制作与审阅的演示稿。
- 需要原生可编辑文字、图形、表格，以及一致的备注和讲稿。

这是工作流与辅助脚本，不是将任意截图自动转换为可编辑 PPT 的工具。

## 安装

将本仓库放入 Codex 的本地 skills 目录：

```bash
git clone https://github.com/ironman759/outline-to-editable-ppt.git \
  ~/.codex/skills/outline-to-editable-ppt
```

如果目标目录已有同名 skill，请先备份并比较版本。随后在 Codex 中调用 `$outline-to-editable-ppt`，提供用途、受众、源材料及参考风格。

## 工作流程

1. 确认细致逐页大纲、资产清单和验收标准。
2. 生成代表页视觉稿，确认风格并验证可编辑原型。
3. 逐页生成视觉稿，再重建原生文字、数字、图形、箭头和表格；复杂插图保留为独立图片。
4. 按约定部分审阅，维护稳定页 ID 与大纲、资产、讲稿的映射。
5. 合并检查，从单一讲稿源同步备注和 Word，完成视觉与真实 PowerPoint 验证。

默认风格为白底、微软雅黑、黑灰文字、红色强调与 BioArt 插图；当前用户选择和源材料优先。默认交付带备注的可编辑 PPTX 和独立逐页 Word 讲稿。

## 文件结构

```text
SKILL.md
agents/openai.yaml
assets/default-style.json
assets/default-template.pptx
references/page-plan.md
references/style.md
references/rebuild-and-qa.md
scripts/sync_scripts.py
scripts/check_pptx.py
```

`SKILL.md` 是主入口。模板为无内容、无 logo 的 16:9 PPTX。`references/` 提供逐页施工稿、默认风格和验收规则。

## 辅助脚本

运行 Python 脚本需要 `python-pptx` 和 `python-docx`。完整流程还会编排 `jarvis-requirements`、`imagegen`、`pptx` 和 `docx` 等能力，它们不包含在本仓库内。微软雅黑需由运行环境提供。

```bash
python -m pip install python-pptx python-docx
```

对已重建的原生 PPT，同步备注及 Word 讲稿，并检查输出：

```bash
python scripts/sync_scripts.py working/native.pptx working/slides.json \
  --pptx exports/presentation.pptx --docx exports/speaker-notes.docx
python scripts/check_pptx.py exports/presentation.pptx \
  --manifest working/slides.json --docx exports/speaker-notes.docx
```

Manifest 格式、稳定页 ID 和原生标题映射要求见 [逐页施工稿](references/page-plan.md) 与 [重建及验收](references/rebuild-and-qa.md)。同步备注前应将旧备注中必须保留的内容迁入 manifest；脚本会覆盖备注正文。

## 验收边界

结构检查不能代替内容核对、视觉 QA、编辑性抽查或真实 PowerPoint 打开验证。生图中的错字、数据和科学机制必须与源材料核对；概念插图不能作为真实设备或实验结果。未完成真实应用验证时，应标记兼容性待验证。

## 许可证

本仓库采用 [MIT License](LICENSE)。
