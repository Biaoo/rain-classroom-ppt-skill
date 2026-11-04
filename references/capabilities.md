# 能力与调用契约

本文介绍原生习题、投票和通用互动入口的宏调用方式。AppleScript 示例用于 macOS PowerPoint，入口来自实际加载的 `RainClassroom.ppam`；环境发现与版本依据见第 6 节。

## 1. 新建题目：Message.InsertProb

完整宏名：`RainClassroom.ppam!Message.InsertProb`。

参数依次为 `operationId`、`questionJson`、`"0"`。同一个操作标识再次调用仍可能重复插入。每次调用在当前页后增加一页，并选中新页。

### 支持的题型

| 能力 | ProblemType | 主要字段 | 保存结果 |
|---|---:|---|---|
| 单选题 | 1 | `Options`、`Answer`、`Score` | `MultipleChoice`，一个正确选项 |
| 多选题 | 2 | `Options`、`Answer`、`Score`、`HalfScore` | `MultipleChoiceMA`，多个正确选项及少选得分配置 |
| 单选或多选投票 | 3 | `Options`、`PollingCount`、`Anonymous` | `Polling`，最大选择数量和匿名状态 |
| 填空题 | 4 | `Body` 中的占位、`Blanks`、`OrderInsensitive` | `FillBlank`，每空答案、分值及匹配规则 |
| 主观题 | 5 | `Body`、`Score`、`HasRemark`、`Remark` | `ShortAnswer`，分值及解析 |
| 判断题 | 6 | `Body`、`Options: []`、`Answer`、`Score` | 自动转成 A/B 两项的原生单选题 |
| 信息页 | 0 | `Body` | 纯内容页，不作为可作答题目 |

匿名投票使用 `ProblemType: 3, Anonymous: true`。`PollingCount: 1` 是单选，设为不超过选项总数的更大正整数即为多选投票。

### 公共字段

| 字段 | 类型与含义 |
|---|---|
| `ProblemType` | 上表中的整数 |
| `Body` | 非空题干字符串；信息页使用它保存正文 |
| `Options` | `[{"key":"A","value":"选项文字"}, ...]`，选择/投票题至少两个选项 |
| `Answer` | 正确答案数组，例如 `["B"]` 或 `["A","C"]`；投票无需标准答案 |
| `Score` | 数值分数；填空题总分由各空分数构成 |
| `HalfScore` | 多选题少选得分数值 |
| `HasRemark` | 是否附答案解析，布尔值 |
| `Remark` | 解析字符串，配合 `HasRemark: true` |
| `PollingCount` | 投票最多可选项数，正整数 |
| `Anonymous` | 投票匿名状态，布尔值 |

主观题使用 `HasRemark: true` 时，`Answer` 数组中的文字会拼接到 `Remark` 前作为解析。完整解析也可直接放入 `Remark`。

### 单选题示例

```json
{
  "ProblemType": 1,
  "Body": "哪一个是演示文稿编辑软件？",
  "Options": [
    {"key": "A", "value": "PowerPoint"},
    {"key": "B", "value": "图片查看器"}
  ],
  "Answer": ["A"],
  "Score": 2,
  "HasRemark": true,
  "Remark": "PowerPoint 是演示文稿编辑软件。"
}
```

### 填空题示例

```json
{
  "ProblemType": 4,
  "Body": "常用演示文稿软件是#填空1#，互动插件是#填空2#。",
  "BlankKeyWord": "填空",
  "BlankDelimiter": "#",
  "Score": 4,
  "Blanks": [
    {"Num": 1, "Score": 2, "Answers": ["PowerPoint", "PPT"], "CaseSensitive": false, "FuzzyMatch": false},
    {"Num": 2, "Score": 2, "Answers": ["雨课堂"], "CaseSensitive": false, "FuzzyMatch": false}
  ],
  "OrderInsensitive": false
}
```

占位编号与 `Blanks[].Num` 对应，从 1 开始。`Answers` 是该空可接受的答案列表。字段名沿用示例大小写，尤其是 `Score`，以避免插件兼容旧数据的分值换算分支。正文插入后显示为 `[填空1]`、`[填空2]`。

各题型的完整输入见 [questions.json](questions.json)。它可直接交给创建脚本。

## 2. 定位并重新解析：Message.EditProblem

完整宏名：`RainClassroom.ppam!Message.EditProblem`。唯一参数为 **PowerPoint 原生 SlideID** 的字符串形式，通过目标页的 `slide ID` 属性取得。

```applescript
tell application "Microsoft PowerPoint"
    set p to active presentation
    -- 调用方先确认 full name of p 与目标文件完整路径一致
    go to slide (view of active window) number 2
    set targetId to (slide ID of slide 2 of p) as text
    run VB macro macro name "RainClassroom.ppam!Message.EditProblem" ¬
        list of parameters {targetId}
end tell
```

它让插件从目标页重建当前题目状态，并唤起编辑面板。执行下表的编辑宏前调用它，将 `CurrentProblem` 指向目标题。

## 3. 编辑题目：ThisAddIn

完整宏名前缀为 `RainClassroom.ppam!ThisAddIn.`。以下参数按表中顺序传递；AppleScript `list of parameters` 可使用文本形式的数值、`"True"` 或 `"False"`。

| 能力 | 宏名 | 参数示例 | 效果 |
|---|---|---|---|
| 改正确答案 | `SetAnswer` | `{"B"}`；多选 `{"AC"}` | 重新标记正确选项 |
| 改分值 | `SetScore` | `{"3"}` | 更新题目分值 |
| 改多选少选得分 | `SetHalfScore` | `{"1.5"}` | 更新少选得分配置 |
| 添加选项 | `AddOption` | `{}` | 新增一个带字母标识的选项占位，随后填入其文字 |
| 修改解析 | `SetRemark` | `{"True", "解析文字"}` | 设置是否附解析及内容 |
| 改投票最多可选数 | `SetPollingCount` | `{"3"}` | 上限须在 1 到当前选项总数之间 |
| 改投票匿名状态 | `SetPollingAnonymous` | `{"True"}` | 设置匿名投票 |
| 改某空分值 | `SetBlankScore` | `{"1", "3"}` | 第一空改为 3 分，同时更新总分 |
| 改某个备选答案 | `SetBlankAnswer` | `{"1", "1", "新答案"}` | 修改第一空已有的第一个答案 |
| 删除某个备选答案 | `DeleteBlankAnswer` | `{"1", "2"}` | 删除第一空的第二个答案 |
| 设置大小写匹配 | `SetCaseSensitive` | `{"1", "True"}` | 第一空区分大小写 |
| 设置模糊匹配 | `SetFuzzyMatch` | `{"2", "True"}` | 第二空启用模糊匹配 |
| 设置答案顺序无关 | `SetOrderInsensitive` | `{"True"}` | 更新填空题答案乱序配置 |
| 设置语音作答 | `SetVoiceAllowed` | `{"True"}` | 更新主观题允许音频作答的配置 |

例如：

```applescript
tell application "Microsoft PowerPoint"
    run VB macro macro name "RainClassroom.ppam!ThisAddIn.SetAnswer" list of parameters {"B"}
    run VB macro macro name "RainClassroom.ppam!ThisAddIn.SetScore" list of parameters {"3"}
    save active presentation
end tell
```

题干和选项文字是 PowerPoint 文本对象，可在保留形状标签的前提下用原生文本接口修改。新增占位选项后需补齐内容；插入多道新题可逐题调用 `InsertProb`，每次验证页数并重新确认位置。

## 4. 幻灯片上的互动入口：Message.InsertInteraction

完整宏名：`RainClassroom.ppam!Message.InsertInteraction`。在当前页插入原生互动对象，支持以下四种通用入口：

| 互动 | `type` |
|---|---:|
| 弹幕 | 1 |
| 投稿 | 2 |
| 点名 | 3 |
| 抢答 | 4 |

参数按以下顺序传递：

| 位置 | 参数 | 通用入口取值 |
|---|---|---|
| 1 | `type` | 上表的 1–4 |
| 2 | `envId` | `0` |
| 3 | `buttonStyle` | `0` |
| 4 | `src` | 本地按钮图片绝对路径；新建时空字符串生成灰色圆角对象 |
| 5、6 | `srcImageWidth`、`srcImageHeight` | 图像尺寸；本例为 `600`、`180`，插入后直接调整形状尺寸 |
| 7 | `label` | 通用入口传空字符串；此字段不负责通用按钮的可见文字 |
| 8 | `agentId` | `0` |
| 9 | `interactionId` | 空字符串 |
| 10 | `shapeId` | 新建传空字符串 |

以下完整示例在指定课件第一页新建弹幕对象。把 `targetPath` 改为目标文件的完整路径，先将该文稿置为活动文稿。示例使用空 `src`，再通过 PowerPoint 原生文本接口填入提示语，保留插件生成的标签。

```applescript
set targetPath to "/absolute/path/lesson.pptx"
tell application "Microsoft PowerPoint"
    if full name of active presentation is not targetPath then error "Wrong document"
    set p to active presentation
    go to slide (view of active window) number 1
    set n to count of shapes of slide 1 of p
    run VB macro macro name "RainClassroom.ppam!Message.InsertInteraction" ¬
        list of parameters {"1", "0", "0", "", "600", "180", "", "0", "", ""}
    if (count of shapes of slide 1 of p) is not (n + 1) then error "Insert failed"
    set s to shape (n + 1) of slide 1 of p
    set content of text range of text frame of s to "弹幕：说出一个疑问"
    set left position of s to 40
    set top of s to 140
    set width of s to 350
    set height of s to 55
    save p
end tell
```

插件通过形状的 `INTERACTIONTYPE`、`INTERACTIONBUTTONSTYLE`、`LOCALSHAPEID` 识别互动。保留这些标签，用 PowerPoint 原生接口调整文字、位置和尺寸，并核对 `INTERACTIONTYPE` 与预期类型一致。上述空 `src` 示例用于新建对象；更新已有图片对象时提供有效的图片路径。

小组讨论、智能体、试卷等互动还需要对应资源的配置和标识，参数与上述通用入口不同。

## 5. 保存与验收

信息页的原生初始文本框可能过窄；创建脚本会按页面宽度调整，并给定初始高度，后续可根据正文长度继续排版。

`scripts/create_deck.py` 是 macOS 上新建题目课件或小样的辅助脚本，接收题目 JSON 和新的输出路径，处理插题、保存后重新取引用、关闭重开及页数/题型检查。输出包含课件结构、PowerPoint 版本和已加载插件路径，便于对照输入核对答案及配置。在已有标准宽屏课件上插入四选项单选或匿名投票时，Windows 和 macOS 可用 `scripts/insert_questions.py` 和[插题清单](insertion-plan.json)。其他原生题型及页内互动按本文的 macOS 宏示例定位目标文稿并调用插件。

`scripts/inspect_deck.py` 沿 PPTX 中实际的关系引用读取每页及形状标签。关键字段：

| 检查项 | 位置 |
|---|---|
| 题型 | 幻灯片 `RAINPROBLEM` |
| 分值、少选得分 | 幻灯片 `PROBLEMSCORE`、`PROBLEMSCORE_HALF` |
| 答案解析 | 幻灯片 `PROBLEMHASREMARK`、`PROBLEMREMARK` |
| 正确选项 | 字母形状 `RAINBULLET=Correct`，结合该形状的文字 |
| 题干、选项文字 | 形状 `RAINPROBLEM=ProblemBody/ProblemItem` |
| 匿名投票 | 幻灯片 `ANONYMOUSPOLLING` |
| 投票可选上限 | `RAINPROBLEM=PollingAnswer` 形状的完整文本，例如“最多可选3项” |
| 填空答案、分值、匹配规则 | 幻灯片 `PROBLEMBLANK` 中的 JSON |
| 填空答案乱序 | 幻灯片 `PROBLEMBLANKORDER` |
| 语音作答配置 | 幻灯片 `PROBLEMVOICEALLOWED` |

`PROBLEMBLANK` 是 JSON 字符串，先用 JSON 解析器还原答案，再与输入比较。交付使用普通 PPTX，宏由授课电脑安装的插件提供。

## 6. 环境发现与接口维护

通过 PowerPoint 查询实际版本和已加载插件的路径：

```applescript
tell application "Microsoft PowerPoint"
    if not (loaded of add in "RainClassroom.ppam") then error "RainClassroom is not loaded"
    return version & linefeed & (full name of add in "RainClassroom.ppam")
end tell
```

返回两行：PowerPoint 版本、插件完整路径。查询不到加载项时，检查 PowerPoint 加载项列表与客户端连接。`RainClassroom.ppam!Message.InsertProb` 中的文件名用于解析已加载宏。

上述接口在 macOS、PowerPoint 16.114.3、雨课堂 7.1.2 上完成了生成或修改、保存重开和原生标签核对，验证覆盖本地课件制作与配置。适配其他版本时，用小样核对题型、配置及重开后的插件识别结果。

需要适配宏契约时，读取查询得到的插件文件；如受应用沙箱限制无法读取，可通过系统应用定位找到当前雨课堂客户端，在实际应用资源中查找同名插件，并核对其与已加载版本的对应关系。安装目录名可能沿用旧版本，版本信息应来自客户端自身或应用元数据。

适配这些内部宏接口时，可用 oletools 的 `VBA_Parser.extract_macros()` 读取定位后的插件，对照以下模块：

- `Message.bas`：`InsertProb`、`EditProblem`、`InsertInteraction`。
- `V_AIAssistantHelper.bas`、`M_Consts.bas`：互动对象创建、解析和标签。
- `LibQuestion.cls`、`KeyValuePair.cls`、`BlankInfo.cls`：输入字段与反序列化。
- `C_ImportProblem.bas`：按题型选择生成实现。
- `MultipleChoice.cls`、`MultipleChoiceMA.cls`、`Polling.cls`、`FillBlank.cls`、`OpenQuestion.cls`、`InfoSlide.cls`：原生页生成。
- `ThisAddIn.bas`、`PPTProblem.cls`：当前题目定位、编辑与标签读写。
- PowerPoint 自带 `Contents/Resources/PowerPoint.sdef`：当前版本的 AppleScript 命令字典。
