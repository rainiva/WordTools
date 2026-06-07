# Word/WPS 多宿主矩阵自动化测试

基于 TDD 构建的配置驱动测试框架，按 **Office/WPS 位数组合** 设计环境，而不是按单宿主拆机。

## 测试分层

| 层级 | 脚本 | 占比建议 | 说明 |
|------|------|----------|------|
| 第 1 层 | `ps/Matrix.HostProbe.ps1` | 探测 | 识别 Word/WPS 安装路径、位数、注册表线索 |
| 第 2 层 | `ps/Matrix.Register.ps1` 等 | 70% | 一键注册/卸载/验证（优先 `Installer.Core.ps1`） |
| 第 3 层 | `ps/Matrix.SmokeInsertImage.ps1` | 10% | 冒烟（当前为宿主可启动性门禁，UI 插图待扩展） |

## 测试矩阵（按位数组合）

| ID | 配置 | Word | WPS | 状态 |
|----|------|------|-----|------|
| VM-01 | `word32_wps32.json` | 32 | 32 | 计划适配 |
| VM-02 | `word32_wps64.json` | 32 | 64 | 计划适配 |
| VM-03 | `word64_wps32.json` | **64** | 32 | **Word64 正式支持** + WPS32 实验 |
| VM-04 | `word64_wps64.json` | **64** | 64 | Word64 正式 + WPS64 待适配 |

索引文件：`configs/matrix_index.json`

### Word 64 专用配置

| 配置 | 用途 |
|------|------|
| `word64_only.json` | Word 64 主回归：探测→注册预览→COM 冒烟→卸载 |
| `word64_wps32.json` | VM-03 完整流水线（Word64+WPS32 共存） |
| `word64_wps32_com_smoke.json` | Word64 COM 加载项检查 |
| `word64_wps32_live_register.json` | VM 快照 Live 注册/卸载 |
| `word64_wps64.json` | VM-04 完整矩阵预览 |

## 运行方式

```powershell
cd automation
pip install -r requirements.txt
python -m pytest tests -v
python run_matrix_test.py --list-envs
python run_matrix_test.py --env VM-03
python run_matrix_test.py --config configs/word64_only.json
```

## 输出

- `reports/<env>/matrix-report.json` — 总报告
- `reports/<env>/host-probe.json` — 宿主探测详情

## 与 Installer.Core.ps1 的关系

若仓库根目录存在 `Installer.Core.ps1`（已从 `codex/DescriptionNumber` 引入），探测/注册/卸载会自动走共享安装核心；否则使用 `Matrix.*.ps1` 回退逻辑。

### 新增配置

| 配置 | 用途 |
|------|------|
| `word64_wps32_com_smoke.json` | 本机 Word64+WPS32，COM 加载项冒烟 |
| `word64_wps32_live_register.json` | **VM 快照专用**：Live 注册→验证→COM 冒烟→卸载 |

Live 配置会修改注册表，仅在虚拟机快照上运行。

## TDD 约定

1. 先写 `tests/` 中的 pytest 用例
2. 确认 RED（失败原因必须是功能缺失）
3. 写最小实现使 GREEN
4. 重构并保持全绿

## 虚拟机建议

在 VM 快照上运行 Live 注册/卸载；本机默认使用 `PreviewOnly` 避免污染注册表。
