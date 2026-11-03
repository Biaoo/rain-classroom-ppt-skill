# 故障排查

## WorkBuddy：已授权自动化，仍报 `-10004`

在 macOS 上，WorkBuddy 调用 PowerPoint 时若报 `-10004`（权限违例），即使系统自动化权限已开启、对话已选择“允许完全访问”，仍可能受到 WorkBuddy 自身沙箱的限制。

此时可能需要关闭 **WorkBuddy → 设置 → 安全中心 → 沙箱安全**。“允许完全访问”与“沙箱安全”是两个独立设置。关闭沙箱会扩大命令的系统访问范围，应由用户决定是否关闭。

关闭后，在 WorkBuddy 中执行以下只读命令验证：

```bash
osascript -e 'tell application "Microsoft PowerPoint" to count presentations'
```

返回文稿数量且退出码为 `0`，说明 PowerPoint 查询已成功，可继续运行雨课堂脚本。

已在 WorkBuddy 5.9.3、PowerPoint 16.114.3 上验证：相同查询在沙箱开启时返回 `-10004`，关闭后返回 `1`、退出码为 `0`，无需重新配置 macOS 自动化权限。此结果验证的是 PowerPoint 通信，课件仍按技能中的流程验收。
