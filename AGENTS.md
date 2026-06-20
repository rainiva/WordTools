# Repository Rules / 仓库规则

## 中文与编码安全

本仓库含大量中文注释、UI 文案与资源文件，**编码损坏是最高优先级风险**。

**禁止的修改方式：**

- PowerShell / 脚本：`Get-Content` + `Set-Content`、`Out-File`、未指定 `-Encoding UTF8` 的读写
- Shell：`type`、`cat`、重定向、`echo` 等整文件覆盖
- 任何依赖终端/console 解码再写回的文件操作
- 未经验证编码的批量替换或脚本式改写

**必须遵守：**

- 不要把终端、控制台、diff 预览里显示的非 ASCII 文本当作文件真实内容
- 声称文件"乱码"或"损坏"前，先做字节级或显式编码核验（如 UTF-8 BOM / 无 BOM）
- 区分"文件真的坏了"与"只是显示层解码错误"，再决定是否修改
- 编码安全无法确认时，停止修改，改用更小、更稳的补丁

**含中文文件的编辑方式：**

| 环境 | 推荐工具 | 禁止 |
|------|----------|------|
| **Codex** | `Write` 精确补丁 | 整文件回写、脚本读写 |
| **Cursor** | `StrReplace` / `Write` 精确补丁 | PowerShell 改源码、整文件覆盖 |

对 `*.resx`、`*.Designer.cs`、含中文的 `*.xml` / `*.cs` 尤其谨慎；非必要不手改，必须改时只做最小变更。

## AI 协作

- 先读上下文，再动代码；修改前看周边实现与现有模式，贴合仓库写法
- 小改优先，不擅自扩需求；不把重构、格式化或额外优化混进当前任务
- 不覆盖、不回退你没创建的改动；与现有未提交修改冲突时先确认
- 对 Office interop、COM、UI 事件绑定、文档处理流程保持保守
- 不要把推测当需求，不要替用户补充未明确提出的功能

## 注释与生成

- 注释只写高价值信息：设计意图、边界条件、兼容性约束、Office/COM 坑点
- 不写复述代码表面的注释；改动行为时同步修正失真的注释
- 生成代码必须可直接落地，禁止占位实现、伪成功路径、空 `TODO`
- 遵循现有命名、判空、异常处理风格；无充分理由不引入新依赖或新层次

## 验证与交付

- 没有证据不声称"已修复""已完成""可安全发布"
- 能做构建、测试或定向验证时就做；做不了要说明原因
- 交付时区分：改了什么 / 验证了什么 / 哪些仍是未验证或风险点

## 项目上下文

- 主代码：`WordTools/`（COM 加载项，.NET Framework 4.8）
- 单元测试：`WordTools.Tests/`（`dotnet test`）
- 多宿主矩阵测试：`automation/`（pytest + PowerShell）
- 安装与注册：`build-installer.ps1`、`Setup.iss`、`RegisterPlugin.ps1`
- 规划与安装文档：`docs/superpowers/plans/`、`docs/installer/`
- 代码库分析：项目内 `.claude/skills/graphify/`

## 技能与工具（Codex / Cursor）

**优先级：** 用户明确指令 > 相关 skill > 默认系统行为。

**使用 skill：** 任务开始前检查是否有适用 skill（brainstorming、debugging、TDD、writing-plans 等）。有疑则读，读了就按 skill 执行；skill 内容与 AGENTS.md 冲突时以 AGENTS.md 为准。

| 环境 | 读 skill | 编辑代码 | 任务跟踪 |
|------|----------|----------|----------|
| **Codex** | skill 工具 | `Write` | 内置 todo |
| **Cursor** | `Read` 读 skill 文件 | `StrReplace` / `Write` | `TodoWrite`、`Task` |

"Add X" / "Fix Y" 不意味着跳过 brainstorming、debugging 等工作流。

## 优先级

以上规则优先于便利性和速度。目标是避免中文损坏、AI 误改、行为回归，以及未验证却宣称完成的交付。
