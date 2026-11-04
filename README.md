# 雨课堂 PPT Skill

配合通用 PPT 制作 skill，为 PowerPoint 课件添加雨课堂原生习题、投票和互动入口。

先用通用 PPT skill 完成内容、页序和排版，再用本技能从基础稿生成含雨课堂原生对象的交付稿。

| 能力 | Windows 和 macOS 文件脚本 | macOS PowerPoint 插件 |
|---|---|---|
| 在已有课件插入单选、投票 | 四选项单选题、四选项匿名投票 | 单选和投票，支持调整分值、匿名状态、可选数等 |
| 多选、填空、主观、判断、信息页 | 支持全部五类 | `Message.InsertProb` |
| 在当前页添加弹幕、投稿、点名、抢答 | 按原稿页码添加原生按钮 | `Message.InsertInteraction` |
| 重新解析及编辑已有题目 | — | `Message.EditProblem` 和对应编辑宏 |
| 检查保存后的文字和原生标签 | `inspect_deck.py` | 可用编辑面板确认插件识别 |

详细字段见[能力与调用契约](references/capabilities.md)。

## 安装

通过 Skills CLI 安装到支持 Agent Skills 的工具：

```bash
npx skills add Biaoo/rain-classroom-ppt-skill
```

也可将仓库内容放入所用 Agent 的技能目录，目录名使用 `rainclassroom-ppt`。技能入口是 [SKILL.md](SKILL.md)。

## 使用

向 Agent 提出包含雨课堂插件功能的 PPT 制作需求，例如：

> 使用 $rainclassroom-ppt，在这份课件中添加一道原生单选题和一道匿名投票，保留已有版式。

### 在已有课件中插题或添加互动按钮（Windows 和 macOS）

```bash
python3 scripts/insert_questions.py "基础稿.pptx" references/full-plan.json "交付稿.pptx"
```

`scripts/insert_questions.py` 支持 Windows 和 macOS。[双平台 CI](https://github.com/Biaoo/rain-classroom-ppt-skill/actions/workflows/insert-questions.yml)

Windows 如使用 Python 启动器，将 `python3` 换为 `py -3`。标准库即可完成生成。基础稿使用 PowerPoint 标准宽屏尺寸（13⅓ × 7½ 英寸）。按原稿页码定位，`after` 在该页之后插入题页或信息页，`on` 将互动按钮添加到该页。字段说明见 [SKILL.md](SKILL.md)，示例见[完整清单](references/full-plan.json)。输出指定新的文件名。

### 从题目 JSON 新建课件（macOS）

```bash
python3 scripts/create_deck.py references/questions.json /absolute/path/new-lesson.pptx
```

输入为题目 JSON 数组，输出路径所在目录需已存在，文件名使用新的 `.pptx` 路径。脚本完成插题、保存重开及页数和题型检查，并输出结构化结果。

在 Mac PowerPoint 中编辑已有题目时，用 `Message.EditProblem` 定位题目，再调用对应的[编辑宏](references/capabilities.md)。

### 检查课件

```bash
python3 scripts/inspect_deck.py /absolute/path/lesson.pptx
```

检查脚本使用 Python 3 标准库读取 PPTX，可独立于 PowerPoint 运行。在装有 Python 启动器的 Windows 环境中，从技能目录运行：

```powershell
py -3 scripts/inspect_deck.py "课件.pptx"
```

## 文件结构

```text
SKILL.md                     技能入口
agents/openai.yaml           技能展示信息
references/capabilities.md   宏接口与字段契约
references/questions.json    各题型输入示例
references/troubleshooting.md WorkBuddy 权限与 PowerPoint 无响应排查
references/insertion-plan.json  跨平台插题清单示例
references/full-plan.json    七种页面与四种页内按钮示例
assets/native-pages.zip      雨课堂原生题页模板
scripts/insert_questions.py  向已有标准宽屏 PPTX 插题
scripts/create_deck.py       通过 Mac 插件新建题目课件
scripts/inspect_deck.py      读取课件文字与标签
```
