# Word工具箱

这是一个Word COM 加载项项目，安装后会在 Word 功能区添加一个名为"Word工具箱"的自定义选项卡。

## 功能特性

安装后，Word 功能区将新增"Word工具箱"选项卡，包含以下功能：

- **图片工具** 组
  - 批量插图：打开批量插图工具窗口，从文件夹或选择的图片文件批量插入到Word表格中，支持自动编号和描述行

- **帮助** 组
  - 关于：显示版本信息和使用说明

## 系统要求

- Windows 7 或更高版本
- Microsoft Office 2010 或更高版本（Word）
- .NET Framework 4.8

## 构建说明

### 使用 Visual Studio

1. 打开 `WordTools.sln`
2. 选择 Debug 或 Release 配置
3. 按 **F5** 运行调试（会自动注册插件）
4. 或生成解决方案后手动注册

### 手动注册插件

构建成功后，以管理员身份运行命令提示符：

```batch
cd WordTools\bin\Debug
regasm /codebase WordTools.dll
```

### 使用命令行

运行 `build.bat` 批处理文件：

```batch
build.bat
```

## 项目结构

```
WordTools/
├── WordTools.sln          # 解决方案文件
├── build.bat              # 构建脚本
├── README.md              # 使用说明
└── WordTools/             # 插件项目
    ├── WordTools.csproj   # 项目文件
    ├── ThisAddIn.cs       # 插件入口（IDTExtensibility2）
    ├── Ribbon.cs          # 功能区实现
    ├── Ribbon.xml         # 功能区 XML 定义
    └── Properties/        # 程序集信息
```

## 卸载方法

以管理员身份运行命令提示符：

```batch
cd WordTools\bin\Debug
regasm /u WordTools.dll
```

或在 Word 中：
1. 文件 → 选项 → 加载项
2. 管理 COM 加载项 → 转到
3. 取消勾选"WordTools"

## 技术说明

本项目使用 **COM 加载项** 方式实现（IDTExtensibility2 接口），而不是 VSTO。
- 优点：不需要 VSTO Runtime，依赖更少
- 缺点：需要手动注册 COM 组件

## 许可证

MIT License
