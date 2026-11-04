# 雨课堂 PPT Skill

配合通用 PPT 制作 skill，为 PowerPoint 课件添加雨课堂原生习题、投票和互动入口。

通用 PPT skill 负责内容组织与版式，本技能通过雨课堂插件宏创建原生对象，并检查保存后的题目配置与标签。

**平台支持**：macOS 提供 PowerPoint 原生插题脚本。Windows 可用 Python 3 检查已保存的 PPTX，目前没有 Windows 原生插题脚本。

## 能力

- 创建单选、多选、投票、填空、主观、判断题及信息页。
- 设置正确答案、分值、少选得分、解析、投票匿名状态与可选数量。
- 修改填空题答案及匹配规则、主观题语音作答配置。
- 在已有幻灯片添加弹幕、投稿、点名、抢答入口。
- 检查 PPTX 中的题型、文字和雨课堂原生标签。

详细参数和宏入口见 [能力与调用契约](references/capabilities.md)。

## 安装

通过 Skills CLI 安装到支持 Agent Skills 的工具：

```bash
npx skills add Biaoo/rain-classroom-ppt-skill
```

也可将仓库内容放入所用 Agent 的技能目录，目录名使用 `rainclassroom-ppt`。技能入口是 [SKILL.md](SKILL.md)。

## 使用

向 Agent 提出包含雨课堂插件功能的 PPT 制作需求，例如：

> 使用 $rainclassroom-ppt，在这份课件中添加一道原生单选题和一道匿名投票，保留已有版式。

现有创建脚本使用 macOS PowerPoint 的 AppleScript 接口。运行时需已安装 PowerPoint 和雨课堂、加载 `RainClassroom.ppam` 并连接客户端。插件位置通过 PowerPoint 查询。

### 从题目 JSON 新建课件（macOS）

```bash
python3 scripts/create_deck.py references/questions.json /absolute/path/new-lesson.pptx
```

输入为题目 JSON 数组，输出路径所在目录需已存在，文件名使用新的 `.pptx` 路径。脚本完成插题、保存重开及页数和题型检查，并输出结构化结果。

为已有课件增加互动时，按参考文档定位文稿和目标页，调用 `Message.InsertProb`、`Message.EditProblem` 或 `Message.InsertInteraction`。

### 检查课件

```bash
python3 scripts/inspect_deck.py /absolute/path/lesson.pptx
```

检查脚本使用 Python 3 标准库读取 PPTX，可独立于 PowerPoint 运行。在装有 Python 启动器的 Windows 环境中，从技能目录运行：

```powershell
py -3 scripts/inspect_deck.py "课件.pptx"
```

## 验证

已在 macOS、PowerPoint 16.114.3、雨课堂 7.1.2 上完成原生题目及互动对象的生成或修改、保存重开和标签核对。验证覆盖本地课件制作与配置。

## 文件结构

```text
SKILL.md                     技能入口
agents/openai.yaml           技能展示信息
references/capabilities.md   宏接口与字段契约
references/questions.json    各题型输入示例
references/troubleshooting.md WorkBuddy 权限与 PowerPoint 无响应排查
scripts/create_deck.py       新建题目课件
scripts/inspect_deck.py      读取课件文字与标签
```
