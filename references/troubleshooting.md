# 故障排查

## 遍历形状时无响应

在 macOS 上，用 AppleScript 遍历整页形状并逐个读取文本，有导致 PowerPoint 无响应的风险。查询超时时，插入宏可能已经执行。

1. PowerPoint 恢复响应后，确认打开的是交付稿副本，查看目标页或题目页数，保存希望保留的改动。
2. 用 `scripts/inspect_deck.py` 读取已保存的交付稿，核对目标页、题干及雨课堂标签。若仍无法保存，脚本只能检查上次保存的版本。
3. 根据核对结果继续，只补缺少的题目或互动入口。

## WorkBuddy：已授权自动化，仍报 `-10004`

在 macOS 上，WorkBuddy 调用 PowerPoint 时若报 `-10004`（权限违例），即使系统自动化权限已开启、对话已选择“允许完全访问”，仍可能受到 WorkBuddy 自身沙箱的限制。

此时可能需要关闭 **WorkBuddy → 设置 → 安全中心 → 沙箱安全**。“允许完全访问”与“沙箱安全”是两个独立设置。关闭沙箱会扩大命令的系统访问范围，应由用户决定是否关闭。

关闭后，在 WorkBuddy 中执行以下只读命令验证：

```bash
osascript -e 'tell application "Microsoft PowerPoint" to count presentations'
```

返回文稿数量且退出码为 `0`，说明 PowerPoint 查询已成功，可继续运行雨课堂脚本。
