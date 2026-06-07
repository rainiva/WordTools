# VM Live 全流程操作手册

在 **虚拟机快照** 上执行，不要在日常开发机上直接跑完整 Live 配置（会修改注册表并在最后卸载插件）。

## 前置条件

1. 快照环境：Word 64 + WPS 32（VM-03）
2. 已编译 `WordTools\bin\Release\WordTools.dll`
3. **以管理员身份** 打开 PowerShell
4. 完全关闭 Word / WPS 进程

## 编译插件（如需要）

```powershell
cd D:\Project\WordTools
msbuild WordTools.sln /p:Configuration=Release
```

## 阶段 1：Live 注册 + 验证 + COM 冒烟（不卸载）

先确认注册与加载正常，避免一次跑完就卸载：

```powershell
cd automation
python run_matrix_test.py --config configs/word64_wps32_live_register.json `
  --skip-phases unregister,verify_cleanup
```

期望结果：

| 阶段 | 期望 |
|------|------|
| host_probe | Word 64 + WPS 32 识别正确 |
| registration | Live 注册 Word64 成功 |
| verify_registration | `word_registry_path` 含 HKLM，`load_behavior_ok: true` |
| smoke (com_load) | `word.addin_loaded: true` |

## 阶段 2：完整 Live 闭环（含卸载）

确认阶段 1 通过后，**回滚快照**，再跑完整流程：

```powershell
python run_matrix_test.py --config configs/word64_wps32_live_register.json
```

期望额外通过：

| 阶段 | 期望 |
|------|------|
| unregister | Live 卸载 Word 成功 |
| verify_cleanup | `word_clean: true` |

## 常见问题

| 现象 | 处理 |
|------|------|
| `Live preflight failed` | 用管理员 PowerShell 重跑 |
| `verify_registration` 失败 | 确认 Installer.Core 写入 HKLM（已修复 HKCU-only 检测） |
| `com_load` 失败 | 先 Live 注册；关闭所有 Word 实例后重试 |
| `register` 失败 | 检查 Release DLL 是否存在、RegAsm 是否 x64 |

## 报告位置

`automation/reports/word64_wps32_live_register/matrix-report.json`
