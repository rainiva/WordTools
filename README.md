# Word工具箱

Word工具箱是一个 Word COM 加载项，安装后在 Word 功能区新增"Word工具箱"选项卡，提供批量插图、表格编号刷新等办公辅助功能。

## 功能特性

| 功能组 | 功能 | 说明 |
|--------|------|------|
| **图片工具** | 批量插图 | 从文件夹或选择的图片文件批量插入到 Word 表格，支持自动编号、描述行、子文件夹递归 |
| **工具** | 刷新编号 | 一键刷新当前表格中的 SEQ 域编号 |
| **帮助** | 关于 | 显示版本信息 |

### 批量插图详细功能

- 支持按文件夹批量插入，或手动选择多个图片文件
- 自动编号：支持序号在描述前/后，左对齐/居中
- 描述行：可使用文件名或文件夹名作为描述
- 子文件夹递归：可选择是否包含子文件夹中的图片
- 图片高度自适应：可设置最小高度，保持比例缩放
- 进度显示：带进度条和详细日志的插入过程

## 系统要求

- Windows 10 / 11
- **64 位 Microsoft Word** 桌面版（2010 或更高版本）
- .NET Framework 4.8
- 管理员权限（安装时）

> **注意**：当前版本仅支持 64 位 Microsoft Word。32 位 Word、WPS、macOS、Office Online 均不支持。

## 安装方法

### 方式一：安装包（推荐）

从 [GitHub Releases](https://github.com/rainiva/WordTools/releases/latest) 下载 `WordToolbox_Setup_<version>_x64.exe`（当前 **v1.3.3**），或本地构建产物位于 `dist/` 目录。

1. 完全关闭 Microsoft Word（包括后台进程）
2. 运行 `WordToolbox_Setup_1.3.3_x64.exe`
3. 按安装向导完成安装
4. 重新打开 Word，确认功能区出现"Word工具箱"选项卡

安装后可在 Word 中检查：
`文件 → 选项 → 加载项 → 管理 COM 加载项 → 转到`，确认"Word工具箱"已启用。

### 方式二：手动注册（开发者）

以管理员身份运行命令提示符：

```batch
cd WordTools\bin\Release
regasm /codebase WordTools.dll
```

卸载：

```batch
regasm /u WordTools.dll
```

## 构建说明

### 使用 Visual Studio

1. 打开 `WordTools.sln`
2. 选择 `Release` 配置
3. 生成解决方案

### 使用命令行

```batch
msbuild WordTools.sln /p:Configuration=Release
```

### 运行测试

```batch
dotnet test
```

## 项目结构

```
WordTools/
├── WordTools.sln                 # 解决方案文件
├── README.md                     # 项目说明
├── INSTALLATION.md               # 安装指南
├── AGENTS.md                     # 仓库协作规则
├── WordTools/                    # 主插件项目
│   ├── WordTools.csproj          # 项目文件
│   ├── ThisAddIn.cs              # 插件入口（IDTExtensibility2）
│   ├── Ribbon.cs                 # Ribbon 回调实现
│   ├── Ribbon.xml                # Ribbon XML 定义
│   ├── Theme.cs                  # UI 主题与 DPI 缩放
│   ├── Forms/                    # 窗体
│   │   ├── InsertPhotosForm.cs   # 批量插图窗体
│   │   ├── ProgressForm.cs       # 进度显示窗体
│   │   └── FailureDetailsForm.cs # 失败详情窗体
│   ├── Services/                 # 业务服务层
│   │   ├── ImageService.cs       # 图片插入核心逻辑
│   │   ├── TableService.cs       # 表格处理与编号刷新
│   │   ├── FileService.cs        # 文件系统操作
│   │   ├── ConfigService.cs      # 配置持久化
│   │   ├── ProgressService.cs    # 带进度条的批量插入
│   │   └── BenchmarkLogService.cs     # 性能基准日志
│   └── Properties/               # 程序集信息
├── WordTools.Tests/              # 单元测试项目
└── dist/                         # 安装包输出目录（Inno Setup 构建产物）
```

## 版本号

单一来源：`version.json`（格式 **x.x.x**，当前 **1.3.3**）。

| 命令 | 说明 |
|------|------|
| `.\sync-version.ps1` | 将 `version.json` 同步到 `AssemblyInfo.cs`、`Setup.iss`、`WordTools.csproj` |
| `.\sync-version.ps1 -Bump Patch` | 修订号 +1（缺陷修复）→ x.x.**n+1** |
| `.\sync-version.ps1 -Bump Minor` | 次版本 +1（新功能/较大改动）→ x.**n+1**.0 |
| `.\sync-version.ps1 -Bump Major` | 主版本 +1（不兼容变更）→ **n+1**.0.0 |
| `.\build-installer.ps1 -Bump Patch` | 先 bump 再构建安装包到 `dist/` |

程序集 `AssemblyVersion` / `AssemblyFileVersion` 自动写为四段 `x.x.x.0`；Ribbon「关于」对话框与安装包显示 `AssemblyInformationalVersion`（三段 semver）。

## 更新日志

### v1.3.3

- Ribbon 入口统一经 `RibbonController` 转发（批量插图、刷新编号）
- 用户可见提示统一经 `INotificationService`（Orchestrator、Ribbon、插图窗体）
- Abstractions 接口标注 Phase 2 落地状态

### v1.3.2

- 移除 Excel 数据填充工具（功能区入口、窗体、服务与相关配置）
- 功能区保留：批量插图、刷新编号、日志设置、关于

### v1.3.1

- 分层 E2E 自动化（smoke / standard / full）与 COM 直连冒烟路径

## 开发环境要求

- Visual Studio 2022（Community/Professional/Enterprise）或 Build Tools for Visual Studio 2022
- 工作负载：**.NET 桌面开发** + **Office/SharePoint 开发**
- .NET Framework 4.8 目标包
- .NET 8.0 SDK（用于运行单元测试）
- [Inno Setup 6](https://jrsoftware.org/isinfo.php)（用于构建安装包）

## 脚本工具

仓库包含以下辅助脚本：

| 脚本 | 说明 |
|------|------|
| `build.bat` | 自动查找 MSBuild，构建 Release 配置，生成强名称密钥，执行 NGen 预编译 |
| `RegisterPlugin.ps1` | PowerShell 注册脚本，自动检测架构，仅支持 64 位 Word |
| `RegisterPlugin.bat` | `RegisterPlugin.ps1` 的批处理入口 |
| `build-installer.ps1` | 调用 Inno Setup 编译器构建安装包到 `dist/`（x64 为正式包，x86 仅用于不支持提示） |
| `sync-version.ps1` | 从 `version.json` 同步版本号，可选 `-Bump Patch|Minor|Major` |

## 已知限制

- 仅支持 **64 位 Microsoft Word**，32 位 Word 和 WPS 各版本均不支持
- 安装和注册需要管理员权限
- 安装/卸载前需完全关闭 Word（包括后台进程）

## 技术栈

- **插件架构**：COM 加载项（IDTExtensibility2 + IRibbonExtensibility）
- **目标框架**：.NET Framework 4.8
- **Office 集成**：Microsoft.Office.Interop.Word（PIA，嵌入互操作类型）
- **UI 框架**：Windows Forms + 自定义 Ribbon XML
- **测试框架**：.NET 8.0 + xUnit
- **安装包**：Inno Setup 6

## 许可证

MIT License
