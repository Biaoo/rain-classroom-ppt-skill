---
name: rainclassroom-ppt
description: 当用户明确要求制作包含雨课堂插件功能的 PPT 时使用，与其他通用 PPT 制作 skill 配合使用。
---

# 雨课堂互动课件

通用 PPT skill 负责内容组织与版式，本技能负责雨课堂原生习题、投票和互动对象。沿用通用 skill 生成的课件，在主要排版完成后通过脚本调用雨课堂插件插入互动内容；交付以保存后仍能被雨课堂识别为准。

## 入口与用法

主要宏入口是 `Message.InsertProb`。本技能提供的脚本与以下调用示例使用 macOS PowerPoint 的 AppleScript 接口；其他平台需使用对应的自动化接口并核对宏契约。

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

## 执行要点

1. 从当前 PowerPoint 查询实际加载的雨课堂插件及其位置，检查客户端连接。发现方法、实测版本和接口适配方法见参考文档第 6 节。
2. 调用前确认活动文稿完整路径和目标页。另存为后按新名称重新获取文稿引用；向前插页后重新定位后续目标，避免页码偏移。
3. 用 JSON 序列化和 AppleScript 字符串转义传参，执行脚本文件或通过标准输入传给 `osascript`。
4. 宏可能吞掉内部错误。插入后核对页数、题干、选项及配置；中断后先检查实际结果，避免重复插题。
5. 保存并重开，使用 `Message.EditProblem` 让插件重新解析；结合标签检查和必要的界面查看验收。视觉操作用于脚本无法覆盖的环节和显示核验。

创建脚本使用 Python 标准库和已安装的本机应用。宏操作当前活动文稿，执行期间保持目标文稿处于活动状态。

## 交付

交付 `.pptx`，说明已添加的互动及配置。保留雨课堂原生标签；经过其他 PPT 工具重写后检查标签和插件识别结果。
