# WPS AddinsWl 实验链设计规范

> 用途：本规范定义 WPS x86 COM 加载项写入契约的实验链设计。
> 所属计划：[四宿主适配计划](../plans/2026-05-24-four-host-adaptation-plan.zh-CN.md) Phase 2
> 日期：2026-05-26

## 一、目标与边界

### 目标

通过增量实验链，在真实 WPS x86 环境上验证 WordTools（COM 加载项）的 `AddinsWl` 写入契约：

- 找出使 WordTools 出现在 WPS **COM 加载项对话框**的最小必要写入条件
- 每个实验产出明确的"命中/未命中"结论
- 实验结束后 `AddinsWl` 完全恢复初始状态

### 非目标

- 不做原生 WPS 插件链（`setupplugin → pool → listV3`）的实验
- 不产出 WPS P0 功能验证
- 不修改 WordTools COM 加载项内部代码
- 不调整 Word x64 安装路径

### 硬边界

- 正式支持边界保持 `Word x64 only`，本实验只推进侦察
- `AddinsWl` 操作前必须先完整备份
- 任一实验导致 `AddinsWl` 无法恢复 → 后续实验全部中止
- 不把任何中间结果表述为"WPS 已支持"

## 二、前置假设与已确认状态

| 假设 | 状态 |
|------|------|
| WPS x86 已安装（`12.1.0.26375`） | ✅ 已侦察 |
| WordTools COM ProgId（`WordTools.ThisAddIn`）已在 WPS x86 机器注册 | ✅ shared core live register |
| `AddinsWl` 当前有 78 个条目，76 个空字符串，2 个非空（版本门控） | ✅ 已侦察 |
| `kshell.dll` 承载共享 COM 加载项对话框 | ✅ 已侦察 |
| `KxCOMAddinsDlg=TCOMAddinsDlg.UnicodeClass` 在三宿主一致 | ✅ 已侦察 |
| COM 加载项路径优于原生插件路径 | ✅ 决策确认 |
| 0/78 个 AddinsWl 条目解析到真实 ProgID | ✅ 已知 |

## 三、实验链

### 实验 1：空字符串写入

| 项目 | 内容 |
|------|------|
| **假设** | WPS 能识别 `AddinsWl` 中的 ProgID，空字符串即有效 |
| **操作** | 备份 → 写入 `WordTools.ThisAddIn = ""` → 重启 WPS → 打开 COM 加载项对话框 |
| **成功判定** | WordTools 出现在 COM 加载项列表中 |
| **命中** | 链路打通，记录完整契约证据，实验链终止 |
| **未命中** | 恢复 `AddinsWl`，进入实验 2 |

### 实验 2：版本门控写入

| 项目 | 内容 |
|------|------|
| **假设** | WPS 要求 `AddinsWl` 值具有版本门控格式（如 `>1.0.0.0`） |
| **操作** | 备份 → 写入 `WordTools.ThisAddIn = ">1.0.0.0"` → 重启 WPS → 观察 |
| **成功判定** | 同实验 1 |
| **命中** | 确认 AddinsWl 需要版本门控 |
| **未命中** | 恢复 `AddinsWl`，进入实验 3 |

### 实验 3：kcmddb 同步分析

| 项目 | 内容 |
|------|------|
| **假设** | `kcmddb` 是 COM 加载项 UI 的数据源，WordTools 需在此出现 |
| **操作** | 读取 `kcmddb_*/db/personal_cn/wps.db` 中 `COMAddIns` 相关数据；分析其与 `AddinsWl` 的关系；如可能，尝试注入 |
| **成功判定** | 确认 `kcmddb` 是否与 `AddinsWl` 联动，或是否需单独写入 |
| **结论** | 命中/未命中均产出分析证据，然后进入实验 4 |

### 实验 4：深度契约综合

| 项目 | 内容 |
|------|------|
| **假设** | `AddinsWl` 可能只是部分环节，加载链可能还涉及 `setupplugin.plg`、`listV3`、或宿主二进制中的硬编码 |
| **操作** | 基于前三轮实验结论，综合分析各层依赖关系 |
| **产出** | 最终契约假设文档（含已确认项 + 未确认风险点） |

### 实验顺序约束

```
实验 1 → 实验 2 → 实验 3 → 实验 4
   ↓        ↓        ↓        ↓
 命中=终止  命中=终止  无论命中/未命中都继续
```

## 四、基础设施变更

### 4.1 `Installer.Core.ps1` — 新增函数

```
Invoke-WpsAddinsWlExperiment
  参数：
    -ProgId          e.g. "WordTools.ThisAddIn"
    -ValuePayload    e.g. "" 或 ">1.0.0.0"
    -ExperimentId    e.g. "exp1"
    -EvidenceDir     证据输出目录
    -Restore         仅执行恢复操作
  内部步骤：
    1. 备份: reg export → evidenceDir/<date>-AddinsWl-backup.reg
    2. 写入前检查: ProgId 是否已存在
    3. 写入: Set-ItemProperty
    4. 验证: 读取确认
    5. 返回 JSON 结果
    6. Restore 模式: reg import backup.reg → 校验条目数
```

### 4.2 `Installer.SupportMatrix.json` — 新增状态

新增 `ui-experiment` 状态，区别于 `supported` / `planned` / `blocked`：

```json
{
  "WPS x86": {
    "status": "ui-experiment",
    "experimentPhase": "exp1",
    "lastExperimentDate": null,
    "lastExperimentResult": null,
    "notes": "AddinsWl 写入实验进行中"
  }
}
```

### 4.3 `WordTools.Tests/Program.cs` — 新增离线测试

- `TestInvokeWpsAddinsWlExperimentDefined` — 函数存在性检查
- `TestInvokeWpsAddinsWlExperimentBackupFlow` — backup.reg 生成逻辑
- `TestSupportMatrixHasUiExperimentState` — 矩阵新状态存在

## 五、证据数据模型

### 5.1 文件命名

```
docs/installer/evidence/
  CurrentMachine-WpsX86-AddinsWl-Exp{1-4}-{date}.json
  CurrentMachine-WpsX86-AddinsWl-Exp{1-4}-{date}.md
  CurrentMachine-WpsX86-AddinsWl-backup-{date}.reg
```

### 5.2 JSON 结构

```json
{
  "ExperimentId": "exp1",
  "Timestamp": "2026-05-26T...",
  "HostMachine": "CurrentMachine-WpsX86",
  "Step": "AddinsWl empty-string write",
  "Precondition": {
    "ComRegistered": true,
    "AddinsWlPreTotal": 78,
    "AddinsWlPreExistingWordTools": false
  },
  "Action": {
    "RegistryPath": "HKCU\\Software\\Kingsoft\\Office\\WPS\\AddinsWl",
    "WrittenName": "WordTools.ThisAddIn",
    "WrittenValue": ""
  },
  "PostWriteState": {
    "EntryPresent": true,
    "AddinsWlPostTotal": 79
  },
  "UIVerification": {
    "WpsRestarted": false,
    "ComAddinsDialogOpened": false,
    "WordToolsListed": null,
    "Notes": ""
  },
  "Conclusion": null,
  "NextExperiment": "exp2"
}
```

### 5.3 判断标准

| WordToolsListed | 含义 | 下一步 |
|-----------------|------|--------|
| `true` | 出现在 COM 加载项列表 | 实验链终止，记录完整契约 |
| `false` | 未出现 | 恢复 AddinsWl，进入下一实验 |
| `null` | UI 验证尚未执行 | 等待人工验证后填入 |

## 六、错误处理与回滚

### 6.1 写入安全流程

```
备份: reg export → 校验 backup.reg 存在且非空
  ↓ 失败 → 中止实验
写入: Set-ItemProperty
  ↓ 失败 → 中止（无需恢复）
验证: 读取确认
  ↓ 成功
人工 UI 验证
  ↓
恢复: reg import backup.reg → 校验条目数与实验前一致
  ↓ 不一致 → 告警 + 出具差异清单
```

### 6.2 熔断条件

| 条件 | 行为 |
|------|------|
| 备份失败 | 中止当前实验，禁止写入 |
| 写入失败 | 中止当前实验，不恢复 |
| 恢复后条目数异常 | 告警，保留手动恢复路径，中止后续实验 |
| WPS 崩溃/异常 | 记录，中止后续实验 |

## 七、测试策略

| 测试 | 类型 | 内容 |
|------|------|------|
| `TestInvokeWpsAddinsWlExperimentDefined` | 离线 | 函数存在于 `Installer.Core.ps1` |
| `TestInvokeWpsAddinsWlExperimentBackupFlow` | 离线 | backup.reg 生成逻辑源码审查 |
| `TestSupportMatrixHasUiExperimentState` | 离线 | 支持矩阵新状态字段 |
| 实验 1~4 | 在线 | 在真实 WPS x86 上执行，人工 UI 验证 |

## 八、涉及文件

| 文件 | 改动 |
|------|------|
| `Installer.Core.ps1` | 新增 `Invoke-WpsAddinsWlExperiment` |
| `Installer.SupportMatrix.json` | 新增 `ui-experiment` 状态 |
| `WordTools.Tests/Program.cs` | 新增 3 个离线测试 |
| `docs/installer/evidence/` | 新增实验证据文件 |

**不动的文件**：
- `Setup.iss` / `RegisterPlugin.ps1` / `RegisterPlugin.bat` — 不调整安装路径
- `WordTools/` 功能代码 — 不碰 COM 加载项内部实现

## 九、与四宿主计划的关系

- 本规范服务于 [四宿主适配计划](../plans/2026-05-24-four-host-adaptation-plan.zh-CN.md) **Phase 2**（WPS 入口验证）
- Phase 2 的完成定义包含"找到可重复验证的 WPS 原生插件写入契约"
- 本实验链的终点（实验 4 深度契约综合）应产出足以满足该定义的分析结论
- 正式支持边界保持 `Word x64 only`，WordTools 在 WPS UI 出现后也不自动升级为 `supported`——需先获得 Phase 3 的 P0 证据
