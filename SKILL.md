---
name: rainclassroom-ppt
description: 当用户明确要求制作包含雨课堂插件功能的 PPT 时使用，与其他通用 PPT 制作 skill 配合使用。
---

# 雨课堂互动课件

先用通用 PPT skill 完成内容、页序和排版，保存为基础稿。再在基础稿的副本中插入雨课堂原生习题、投票和互动对象，保存为交付稿。

## 入口与用法

### 在已有课件中插入题页（Windows 和 macOS）

从技能目录运行以下命令（Windows 的 Python 启动器可用 `py -3`）。脚本使用 Python 3 标准库直接生成新的 PPTX。

```bash
python3 scripts/insert_questions.py "基础稿.pptx" "插题清单.json" "交付稿.pptx"
```

按[插题清单示例](references/insertion-plan.json)填写。每项的 `after` 指原稿页码，从 1 开始。同一页后的多道题按清单顺序插入。支持 PowerPoint 标准宽屏（13⅓ × 7½ 英寸），每题四个选项。

- `choice` 插入 2 分单选题，`answer` 为 A 到 D 中的一个，`remark` 为可选解析。
- `poll` 插入匿名投票，`limit` 为最多可选项数（1 到 4）。

保存后用 `scripts/inspect_deck.py` 对照清单核对页序、文字和标签。

### 通过雨课堂插件创建和编辑（macOS）

其他题型、题目编辑和页内互动使用已加载雨课堂插件的桌面版 PowerPoint。主要插题宏是 `Message.InsertProb`。

```applescript
tell application "Microsoft PowerPoint"
    run VB macro macro name "RainClassroom.ppam!Message.InsertProb" ¬
        list of parameters {operationId, questionJson, "0"}
end tell
```

`operationId` 用于关联本次操作的结果，`questionJson` 是题目 JSON，`"0"` 表示本地新建来源。每次调用在当前幻灯片之后插入新题目，并选中新页。

- 查题型、投票字段、编辑接口和例子时，阅读 [能力与调用契约](references/capabilities.md)。
- 在已有幻灯片添加弹幕、投稿、点名或抢答，使用 `Message.InsertInteraction`；其参数、原生标签和完整脚本见参考文档第 4 节。它添加当前页的互动对象，不新增题目页。
- 从题目 JSON 新建课件或制作小样，可运行 `python3 <技能目录>/scripts/create_deck.py <题目.json> <新课件.pptx>`。[样例 JSON](references/questions.json) 包含各支持题型。
- 检查已有课件，运行 `python3 <技能目录>/scripts/inspect_deck.py <课件.pptx>`。它输出每页文字与雨课堂标签，不依赖 PowerPoint。
- 修改既有课件，按参考文档调用 `Message.EditProblem` 定位题目，再调用相应编辑宏。
- 在 WorkBuddy 中调用 PowerPoint 报权限错误时，阅读 [故障排查](references/troubleshooting.md)。

## 插件宏的执行要点

1. 从当前 PowerPoint 查询实际加载的雨课堂插件及其位置，检查客户端连接。发现方法、实测版本和接口适配方法见参考文档第 6 节。
2. 调用前确认活动文稿完整路径和目标页。另存为后按新名称重新获取文稿引用；向前插页后重新定位后续目标，避免页码偏移。
3. 用 JSON 序列化和 AppleScript 字符串转义传参，执行脚本文件或通过标准输入传给 `osascript`。
4. 插入题目后核对页数是否增加一页，在原页添加互动入口后核对该页是否新增一个形状。不要用 AppleScript 遍历整页形状并逐个读取文本，这样做有导致 PowerPoint 无响应的风险。若检查卡住，见[遍历形状时无响应](references/troubleshooting.md#遍历形状时无响应)。
5. 保存交付稿，再用 `scripts/inspect_deck.py` 对照预期目标页逐项核对题干、选项及配置标签。投票核对匿名状态和可选上限，互动入口核对所在页及 `INTERACTIONTYPE`。需要确认插件识别时，重开课件，在目标题页调用 `Message.EditProblem` 让插件重新解析。

创建脚本使用 Python 标准库和已安装的本机应用。宏操作当前活动文稿，执行期间保持目标文稿处于活动状态。

## 交付

交付 `.pptx`，说明已添加的互动及配置。保留雨课堂原生标签；经过其他 PPT 工具重写后检查标签和插件识别结果。
