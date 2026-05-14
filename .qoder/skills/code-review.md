# WordTools 代码质量审查提示词

> 适用于 WordTools VSTO Word 插件项目（.NET Framework 4.8 + C#）
> 审查范围：`WordTools/` 目录下所有 `.cs` 文件

---

## 使用说明

本提示词用于对 WordTools 项目的代码变更进行系统化、全面的质量审查。审查时应按以下维度逐项检查，发现问题时标注严重等级和问题类型：

### 严重等级定义

| 等级 | 标签    | 含义                 | 典型场景                                         |
| ---- | ------- | -------------------- | ------------------------------------------------ |
| **高** | P0 阻断 | 必须修复才能合并     | 插件崩溃、Word 无响应、COM 对象泄漏、数据损坏    |
| **高** | P1 重要 | 强烈建议修复         | 逻辑缺陷、异常吞没、资源未释放、用户体验问题     |
| **中** | P2 建议 | 改善质量但不阻止合并 | 命名不规范、文档缺失、风格偏差、轻微性能问题     |
| **低** | P3 微调 | 可选的改进建议       | C# 惯用写法优化、LINQ 替代循环、微优化           |

### 问题类型标签

- `[语法/编译]` — 语法错误、类型错误、编译失败、命名空间缺失
- `[逻辑/功能]` — 逻辑漏洞、边界条件错误、功能缺陷、COM 调用问题
- `[代码规范]` — 命名、注释、格式、重复代码
- `[性能]` — COM 调用冗余、内存泄漏、不必要的计算、GC 压力
- `[安全]` — 文件路径遍历、注入风险、敏感数据泄露
- `[可维护性]` — 模块耦合、函数过长、测试缺失、文档不全
- `[COM/Office]` — COM 对象生命周期、Word 互操作、VSTO 规范

### 审查原则

1. **每个问题**必须包含：严重程度 + 问题类型 + 具体位置（文件:行号）+ 修改建议
2. 对于 P0/P1 问题，**必须提供修复后的完整代码示例**
3. 若某维度无问题，**明确标注「✅ 通过」** 并简要说明检查内容
4. 审查应覆盖所有维度，不可跳过

---

## 一、COM 与 Office 互操作规范（P0 级）

### 1.1 COM 对象生命周期管理

- [ ] 是否正确使用 `try-finally` 或 `using` 确保 COM 资源释放？
- [ ] 是否避免在循环中重复获取相同的 COM 属性（如 `tbl.Rows.Count`）？
- [ ] `Marshal.ReleaseComObject()` 是否在必要处调用？（特别是动态创建的 COM 对象）
- [ ] 是否避免了 COM 对象的循环引用？

```csharp
// ❌ 错误：循环中重复访问 COM 属性，每次都会跨进程调用
for (int i = 1; i <= tbl.Rows.Count; i++) { ... }

// ✅ 正确：缓存 COM 属性值
int rowCount = tbl.Rows.Count;
for (int i = 1; i <= rowCount; i++) { ... }
```

### 1.2 Word 应用程序状态管理

- [ ] 修改 `ScreenUpdating`、`DisplayAlerts` 后是否在 `finally` 中恢复？
- [ ] 是否在高性能模式（关闭 ScreenUpdating）后正确恢复？
- [ ] `Application.DoEvents()` 调用是否必要且位置合理？（避免过度调用导致性能下降）
- [ ] 状态栏消息更新是否有意义，而非频繁无意义刷新？

```csharp
// ❌ 错误：未恢复 ScreenUpdating
app.ScreenUpdating = false;
// ... 操作 ...
// 忘记恢复！

// ✅ 正确：使用 try-finally 确保恢复
bool originalScreenUpdating = app.ScreenUpdating;
app.ScreenUpdating = false;
try
{
    // ... 操作 ...
}
finally
{
    app.ScreenUpdating = originalScreenUpdating;
}
```

### 1.3 Range 与 Selection 操作

- [ ] 对 `Range` 赋值 `.Text` 后是否重新获取 `Range` 引用？（赋值后原 Range 失效）
- [ ] `Range.SetRange()` 使用是否正确？是否处理了 end-of-cell marker（`\a`）？
- [ ] 是否避免在循环中频繁创建/销毁 Range 对象？
- [ ] 表格单元格访问是否做了行号/列号越界检查？

```csharp
// ❌ 错误：Text 赋值后继续使用原 Range
r.Text = "新内容";
r.ParagraphFormat.Alignment = ...; // r 已失效！

// ✅ 正确：重新获取 Range
r.Text = "新内容";
r = cell.Range;
r.SetRange(r.Start, r.End - 1);
r.ParagraphFormat.Alignment = ...;
```

### 1.4 VSTO 插件规范

- [ ] `ThisAddIn` 类是否正确实现了 `IDTExtensibility2` 接口？
- [ ] Ribbon 回调方法签名是否符合 `IRibbonExtensibility` 要求？
- [ ] COM 可见性标记 `[ComVisible(true)]` 是否正确配置？
- [ ] `Globals.ThisAddIn` 和 `Globals.Application` 的使用是否安全（空检查）？

---

## 二、语法与编译问题（P0 级）

### 2.1 C# 语法与类型检查

- [ ] 代码是否能在 .NET Framework 4.8 下编译通过？
- [ ] 是否存在未使用的 `using` 语句？
- [ ] 是否存在未使用的变量、字段或方法？（dead code）
- [ ] 可空类型（`int?`）在使用前是否检查了 `HasValue`？
- [ ] `out` 参数是否在调用前正确声明？
- [ ] 字符串插值 `$"..."` 是否优先于 `string.Format()`？（C# 6+ 特性在 .NET 4.8 可用）

### 2.2 异步与线程

- [ ] COM 对象操作是否在 STA 线程上执行？（避免跨线程访问）
- [ ] 是否误用了 `async/await` 处理 COM 操作？（VSTO 中通常应同步执行）
- [ ] 后台线程是否正确使用 `Invoke`/`BeginInvoke` 更新 UI？

---

## 三、逻辑与功能缺陷（P0-P1 级）

### 3.1 逻辑正确性

- [ ] 条件判断是否存在逻辑漏洞？（off-by-one、条件取反）
- [ ] 循环是否存在无限循环风险？终止条件是否正确？
- [ ] `switch` 分支是否处理了所有枚举值？是否有合理的 `default`？
- [ ] 返回值是否在所有代码路径上都有定义？
- [ ] `ref`/`out` 参数是否在所有分支都被正确赋值？

### 3.2 边界条件

- [ ] 空值处理：`null`、空字符串、空集合是否正确处理？
- [ ] 表格边界：行号/列号是否越界？是否检查了 `Rows.Count`/`Columns.Count`？
- [ ] 文件操作：路径为空、文件不存在、权限不足是否处理？
- [ ] 数值边界：除零、负数、溢出是否考虑？
- [ ] 字符串边界：空字符串、超长字符串、特殊字符（`\r\n\a`）是否正确？

```csharp
// ❌ 错误：未检查行号边界
string text = tbl.Cell(row, col).Range.Text;

// ✅ 正确：边界检查
if (row >= 1 && row <= tbl.Rows.Count && col >= 1 && col <= tbl.Columns.Count)
{
    string text = tbl.Cell(row, col).Range.Text;
}
```

### 3.3 异常处理

- [ ] 是否避免了裸 `catch { }` 吞掉所有异常？
- [ ] 异常处理是否记录了足够的信息（至少 `Debug.WriteLine`）？
- [ ] `try-catch` 块是否过大？（应精确包裹可能出错的代码）
- [ ] 是否在 `catch` 中正确恢复应用程序状态？
- [ ] `finally` 块是否释放了资源？

```csharp
// ❌ 错误：裸 catch 吞掉异常，无法调试
catch { }

// ✅ 正确：至少记录异常信息
catch (Exception ex)
{
    Debug.WriteLine($"[ClassName] Error in MethodName: {ex.Message}");
}
```

### 3.4 资源管理

- [ ] `IDisposable` 对象是否使用 `using` 语句？（如 `StreamReader`、`OleDbConnection`）
- [ ] 文件句柄是否在 `using` / `finally` 中正确关闭？
- [ ] 大型对象是否及时释放引用帮助 GC？

---

## 四、代码规范（P2 级）

### 4.1 命名规范

- [ ] 类名使用 `PascalCase`（如 `TableService`、`ProgressService`）
- [ ] 方法名使用 `PascalCase`（如 `RefreshTableNumbering`、`InsertImageToCell`）
- [ ] 私有字段使用 `_camelCase` 前缀下划线（如 `_application`、`_isCancelled`）
- [ ] 常量使用 `UPPER_CASE`（如 `CM_TO_POINTS`、`MAX_PREALLOCATE_ROWS`）
- [ ] 局部变量使用 `camelCase`（如 `rowIndex`、`filePath`）
- [ ] 布尔变量/属性是否有明确的前缀（`is`、`has`、`need`、`use`）？
- [ ] 方法名是否清晰表达意图？（避免 `DoSomething()` 等模糊命名）

### 4.2 文档注释

- [ ] 公开类/方法是否有 XML 文档注释（`/// <summary>`）？
- [ ] 复杂算法是否有行内注释说明**意图**（why），而非描述代码做了什么（what）？
- [ ] 注释是否过时？是否与代码实际行为一致？
- [ ] 配置项、魔法数字是否有注释说明含义？

```csharp
// ✅ 推荐
/// <summary>
/// 刷新表格编号（从光标处开始检查更新）
/// 自动检测光标位置，只处理光标行及之后的编号
/// </summary>
public static void RefreshTableNumbering(Table tbl, Document doc, int alignment = 2, 
    Action<string> progressCallback = null)
```

### 4.3 代码组织与格式

- [ ] 逻辑分段是否使用 `#region` 组织？
- [ ] `using` 语句是否按 系统 → 第三方 → 项目内部 排序？
- [ ] 缩进是否统一为 4 空格？
- [ ] 行长度是否控制在 120 字符以内？
- [ ] 单个文件行数是否超过 800 行？超过时是否有合理的拆分计划？
- [ ] 是否存在大量重复代码？是否可提取为辅助方法？

### 4.4 C# 语言特性最佳实践

- [ ] 是否使用了 `?.` 空条件运算符简化空检查？
- [ ] 是否使用了 `??` 空合并运算符提供默认值？
- [ ] 集合操作是否使用了 LINQ？（注意：COM 集合不支持 LINQ）
- [ ] 字符串拼接是否优先使用插值或 `StringBuilder`？
- [ ] `var` 关键字是否在类型明显时使用？
- [ ] 是否避免了魔法数字？（应提取为命名常量）

```csharp
// ❌ 非惯用
string result = "";
for (int i = 0; i < items.Length; i++)
{
    result += items[i];
    if (i < items.Length - 1) result += ", ";
}

// ✅ 惯用
string result = string.Join(", ", items);
```

### 4.5 错误处理模式

- [ ] 是否使用了防御式编程？（前置条件检查）
- [ ] 错误消息是否使用中文？（与用户界面语言一致）
- [ ] `MessageBox` 提示是否有明确的标题和图标？
- [ ] 是否区分了用户提示和调试日志？

---

## 五、性能优化（P1-P2 级）

### 5.1 COM 调用优化

- [ ] 是否缓存了频繁访问的 COM 属性？（`Rows.Count`、`Columns.Count`）
- [ ] 是否减少了 `Application.DoEvents()` 调用频率？
- [ ] `ScreenUpdating` 是否在批量操作期间关闭？
- [ ] 是否避免了在循环中创建不必要的 COM 对象？

### 5.2 内存管理

- [ ] 大型集合是否及时清空释放？
- [ ] GC 调用是否合理？（避免过度调用 `GC.Collect()`）
- [ ] 图片文件是否及时释放句柄？
- [ ] 事件订阅是否在窗体关闭时取消？（避免内存泄漏）

### 5.3 算法效率

- [ ] 循环嵌套深度是否合理？
- [ ] 字典查找是否替代了线性搜索？
- [ ] 正则表达式是否编译或缓存？（频繁使用时）

---

## 六、安全性审查（P0-P1 级）

### 6.1 文件系统安全

- [ ] 文件路径是否做了合法性验证？
- [ ] `OpenFileDialog`/`FolderBrowserDialog` 的返回值是否检查？
- [ ] 临时文件创建是否使用安全的方式？
- [ ] 是否防范了路径遍历攻击？

### 6.2 输入验证

- [ ] 用户输入的数字是否做了 `TryParse` 验证？
- [ ] 文本框输入是否限制了长度？
- [ ] Excel 数据读取是否处理了异常格式？

### 6.3 敏感信息

- [ ] 错误信息是否泄露了内部实现细节？
- [ ] 日志输出是否包含敏感数据（文件绝对路径）？
- [ ] 注册表操作是否限制了访问范围？

---

## 七、UI/UX 规范（P1-P2 级）

### 7.1 窗体设计

- [ ] 窗体是否应用了统一的 Theme 样式？
- [ ] DPI 缩放是否正确处理？
- [ ] 控件布局是否考虑了高 DPI 显示？
- [ ] 窗体关闭时是否保存了配置？

### 7.2 UI 布局一致性验证（代码修改影响评估）

> 以下检查项专门针对**代码修改可能导致 UI 布局异常**的场景，审查时必须对照修改前后的布局行为。

#### 7.2.1 DPI 缩放影响验证

- [ ] 修改 `Theme.Scale()` 调用或 `S()` 辅助方法时，是否所有控件的 Location/Size 都正确缩放？
- [ ] 新增控件是否都经过 `S()` 缩放？（硬编码像素值会导致高 DPI 下布局错乱）
- [ ] 修改 `Theme.Layout` 常量后，是否检查了所有引用点的布局是否仍然合理？

```csharp
// ❌ 错误：硬编码像素值，高 DPI 下会错位
btn.Location = new Point(100, 50);

// ✅ 正确：所有像素值都经过 DPI 缩放
btn.Location = new Point(S(100), S(50));
```

#### 7.2.2 控件尺寸与位置一致性

- [ ] 同一行内的控件高度是否一致？（Label、TextBox、Button 应使用相同的 `CTRL_HEIGHT`）
- [ ] 控件的垂直对齐是否正确？（`TextAlign = ContentAlignment.MiddleLeft` 与控件高度配合）
- [ ] 修改某个控件的 Size 后，相邻控件的 Location 是否需要同步调整？
- [ ] 窗体最终 `ClientSize` 是否包含所有控件且不留过大空隙或溢出？

```csharp
// ❌ 错误：同一行控件高度不一致，视觉上不对齐
lbl.Size = new Size(S(55), S(25));      // 高度 25
txt.Size = new Size(S(300), CTRL_HEIGHT); // 高度 30

// ✅ 正确：同一行使用统一的高度常量
lbl.Size = new Size(S(55), CTRL_HEIGHT);
txt.Size = new Size(S(300), CTRL_HEIGHT);
```

#### 7.2.3 动态布局计算验证

- [ ] 修改 `topPos` 累加逻辑时，是否所有行都正确递增？（漏掉一行会导致控件重叠）
- [ ] 条件分支（如 `if (needDescription)`）中创建/隐藏控件时，后续控件的 `topPos` 是否相应调整？
- [ ] 窗体 `ClientSize` 的最终计算是否基于 `topPos + 最后一行高度 + MARGIN`？
- [ ] 新增/删除控件行后，窗体高度是否需要重新计算？

```csharp
// ❌ 错误：新增行后忘记调整窗体高度
CreateNewRow(ref topPos);  // 新增了一行
topPos += LINE_SPACING;
// 忘记更新 this.ClientSize，导致窗体装不下或留大空白

// ✅ 正确：窗体高度在末尾统一计算
this.ClientSize = new Size(FORM_WIDTH, topPos + MARGIN);
```

#### 7.2.4 Panel 分组与 RadioButton 互斥

- [ ] 新增 RadioButton 时，是否放入正确的 Panel 容器？（不同组的 RadioButton 应分属不同 Panel）
- [ ] 修改 Panel 的 Location/Size 时，内部控件的相对位置是否仍然正确？
- [ ] 跨 Panel 的控件对齐（如 Panel 外的 CheckBox 与 Panel 内 RadioButton 的垂直对齐）是否一致？

```csharp
// ❌ 错误：RadioButton 未放入 Panel，导致与另一组互斥
optAlignLeft = new RadioButton { ... };  // 未放入 pnlAlignment
this.Controls.Add(optAlignLeft);

// ✅ 正确：同一组 RadioButton 放入独立 Panel
pnlAlignment.Controls.Add(optAlignLeft);
pnlAlignment.Controls.Add(optAlignCenter);
```

#### 7.2.5 主题颜色与字体一致性

- [ ] 新增控件是否应用了 `Theme.Fonts` 和 `Theme.Colors`？（硬编码字体/颜色会破坏主题一致性）
- [ ] 修改 `Theme.Colors` 值后，是否所有控件都正确反映新颜色？
- [ ] 禁用状态的控件是否使用了 `Theme.Colors.InputDisabled` / `Theme.Colors.TextDisabled`？
- [ ] 修改 `Theme.Fonts.Default` 后，是否检查了 `CenterTextVertically()` 的兼容性？

```csharp
// ❌ 错误：硬编码字体和颜色
txt.Font = new Font("宋体", 10);
txt.ForeColor = Color.Black;

// ✅ 正确：使用 Theme 统一配置
txt.Font = Theme.Fonts.Default;
txt.ForeColor = Theme.Colors.Text;
```

#### 7.2.6 布局审查清单（代码修改时必须执行）

| 检查项 | 验证方法 |
|--------|---------|
| 96 DPI (100%) 下布局正常 | 在 96 DPI 显示器上运行或设置系统缩放为 100% |
| 144 DPI (150%) 下布局正常 | 在 144 DPI 显示器上运行或设置系统缩放为 150% |
| 192 DPI (200%) 下布局正常 | 在 192 DPI 显示器上运行或设置系统缩放为 200% |
| 窗体高度包含所有控件 | 检查最底部控件是否完整可见，无被截断 |
| 无控件重叠 | 检查所有控件边界是否相交 |
| 同一行控件垂直对齐 | 检查 Label、TextBox、Button 的文字基线是否对齐 |
| 窗体无过大空白 | 检查窗体底部与最后控件之间的间距是否 ≈ MARGIN |
| 最小窗体尺寸合理 | 检查 `MinimumSize` 是否设置（如设置了 AutoScaleMode=None 则尤为重要）|

### 7.3 跨窗体 UI 风格一致性验证（P1-P2 级）

> 以下检查项确保**不同窗体/页面之间的控件在视觉风格、尺寸、位置等方面保持一致**，避免同一插件内出现多种 UI 风格。

#### 7.3.1 控件类型风格一致性

- [ ] **TextBox 样式**：所有窗体的 TextBox 是否统一使用 `Theme.CreateTextBox()` 创建？
  - 检查属性：`Multiline = true`、`WordWrap = false`、`BorderStyle = Fixed3D`、`Font = Theme.Fonts.Default`
- [ ] **Button 样式**：所有窗体的 Button 是否统一使用 `Theme.CreateButton()` 创建？
  - 检查属性：`FlatStyle = Flat`、`Font = Theme.Fonts.Bold`、`Cursor = Cursors.Hand`
- [ ] **Label 样式**：所有窗体的 Label 是否统一使用 `Theme.CreateLabel()` 或遵循相同规范？
  - 检查属性：`Font = Theme.Fonts.Bold`（标签）、`ForeColor = Theme.Colors.Text`、`TextAlign = MiddleLeft`
- [ ] **CheckBox/RadioButton 样式**：
  - 检查属性：`Font = Theme.Fonts.Default`、`ForeColor = Theme.Colors.Text`
  - 禁用状态：`ForeColor = Theme.Colors.TextDisabled`

```csharp
// ❌ 错误：不同窗体使用不同的创建方式，风格不一致
// InsertPhotosForm.cs
txtFolderPath = Theme.CreateTextBox();  // 使用 Theme

// ExcelDataFillerForm.cs
txtExcelPath = new TextBox();  // 直接 new，缺少统一样式

// ✅ 正确：所有窗体统一使用 Theme.CreateTextBox()
txtExcelPath = Theme.CreateTextBox();
```

#### 7.3.2 控件尺寸一致性

| 控件类型 | 标准高度 | 标准宽度 | 检查项 |
|---------|---------|---------|--------|
| TextBox | `CTRL_HEIGHT` | 根据内容 | 所有输入框高度是否一致？ |
| Button | `CTRL_HEIGHT` | `S(75)` ~ `S(110)` | 同功能按钮宽度是否一致？ |
| Label | `CTRL_HEIGHT` | 根据文字 | 标签高度是否与同行控件一致？ |
| CheckBox | `CTRL_HEIGHT` | 根据文字 | 复选框高度是否一致？ |
| RadioButton | `S(20)` | 根据文字 | 单选按钮尺寸是否一致？ |

- [ ] **按钮宽度一致性**：同类按钮在不同窗体中宽度是否一致？
  - "浏览..." 按钮：`S(75)`（`InsertPhotosForm`）vs `BUTTON_WIDTH`（`ExcelDataFillerForm`）
  - "执行/插入" 按钮：`S(100)` ~ `S(110)`
  - "取消" 按钮：`S(80)` ~ `S(110)`
- [ ] **文本框宽度一致性**：同类型文本框宽度是否遵循相同比例？
  - 短文本框（数字输入）：`S(50)` ~ `S(80)`
  - 中文本框（路径/字段名）：`S(300)` ~ `FORM_WIDTH - inputLeft - MARGIN`

```csharp
// ❌ 错误：同类按钮在不同窗体中宽度不一致
// InsertPhotosForm.cs
btnBrowseFolder.Size = new Size(S(75), CTRL_HEIGHT);

// ExcelDataFillerForm.cs
btnBrowse.Size = new Size(BUTTON_WIDTH, CTRL_HEIGHT);  // BUTTON_WIDTH = S(75)，但命名不同

// ✅ 正确：统一使用 Theme.Layout.ButtonWidth 常量
btnBrowse.Size = new Size(S(Theme.Layout.ButtonWidth), CTRL_HEIGHT);
```

#### 7.3.3 控件位置模式一致性

- [ ] **左侧边距**：所有窗体控件左侧起始位置是否为 `MARGIN` 或 `S(15)`？
- [ ] **标签位置**：标签是否统一位于 `X = MARGIN` 或 `X = S(15)`？
- [ ] **输入框位置**：输入框是否统一位于 `X = inputLeft = MARGIN + LABEL_WIDTH`？
- [ ] **按钮位置**：操作按钮是否统一位于左下角或底部居中？
- [ ] **行间距**：所有行之间的间距是否统一为 `LINE_SPACING`？

```csharp
// ❌ 错误：不同窗体使用不同的左边距
// InsertPhotosForm.cs
Location = new System.Drawing.Point(S(15), topPos)  // 左边距 15

// ExcelDataFillerForm.cs
Location = new Point(MARGIN, topPos)  // 左边距 MARGIN = S(20)

// ✅ 正确：统一使用 MARGIN 常量
Location = new Point(MARGIN, topPos)
```

#### 7.3.4 颜色使用一致性

| 用途 | 标准颜色 | 检查项 |
|------|---------|--------|
| 窗体背景 | `Theme.Colors.Background` | 所有窗体 `BackColor` 是否一致？ |
| 标签文字 | `Theme.Colors.Text` | 所有标签 `ForeColor` 是否一致？ |
| 提示文字 | `Theme.Colors.TextLight` | 提示标签 `ForeColor` 是否一致？ |
| 输入框背景 | `Theme.Colors.InputBackground` | 启用状态输入框背景是否一致？ |
| 禁用输入框 | `Theme.Colors.InputDisabled` | 禁用状态输入框背景是否一致？ |
| 只读输入框 | `Theme.Colors.InputReadonly` | 只读状态输入框背景是否一致？ |
| 主按钮 | `Theme.Colors.Primary` | "主要操作"按钮背景是否一致？ |
| 成功按钮 | `Theme.Colors.Success` | "执行/确认"按钮背景是否一致？ |
| 默认按钮 | `Theme.Colors.ButtonDefault` | "取消/辅助"按钮背景是否一致？ |

- [ ] 新增控件是否使用了上述标准颜色，而非硬编码 `Color` 值？
- [ ] 修改 `Theme.Colors` 后，是否所有窗体都正确反映新颜色？

```csharp
// ❌ 错误：硬编码颜色，不同窗体可能不一致
txtStatus.BackColor = Color.FromArgb(250, 250, 250);

// ✅ 正确：使用 Theme 标准颜色
txtStatus.BackColor = Theme.Colors.InputReadonly;
```

#### 7.3.5 字体使用一致性

| 用途 | 标准字体 | 检查项 |
|------|---------|--------|
| 普通文本 | `Theme.Fonts.Default` | 所有普通控件字体是否一致？ |
| 标签 | `Theme.Fonts.Bold` | 所有标签字体是否一致？ |
| 标题 | `Theme.Fonts.Title` | 区域标题字体是否一致？ |
| 提示 | `Theme.Fonts.Small` | 提示文字字体是否一致？ |
| 等宽 | `Theme.Fonts.Mono` | 状态/日志显示字体是否一致？ |

- [ ] 新增控件是否使用了上述标准字体？
- [ ] 修改 `Theme.Fonts` 后，是否所有窗体都正确反映新字体？

```csharp
// ❌ 错误：直接 new Font，不统一
lblStatusTitle.Font = new Font("微软雅黑", 10, FontStyle.Bold);

// ✅ 正确：使用 Theme 标准字体
lblStatusTitle.Font = Theme.Fonts.Title;
```

#### 7.3.6 分隔线使用一致性

- [ ] 分隔线是否统一使用 `Theme.CreateDivider()` 创建？
- [ ] 分隔线宽度是否为 `FORM_WIDTH - MARGIN * 2`？
- [ ] 分隔线前后间距是否统一？
  - 前间距：`S(Theme.Layout.DividerPaddingTop)` 或 `S(Theme.Layout.SectionSpacing)`
  - 后间距：`S(Theme.Layout.DividerPaddingBottom)` 或 `S(Theme.Layout.SectionSpacing)`

```csharp
// ❌ 错误：手动创建分隔线，样式可能不一致
Label separator = new Label();
separator.Size = new Size(FORM_WIDTH - MARGIN * 2, 1);
separator.BackColor = Theme.Colors.Border;

// ✅ 正确：使用 Theme.CreateDivider()
Label separator = Theme.CreateDivider(FORM_WIDTH - MARGIN * 2);
```

#### 7.3.7 跨窗体审查清单

| 检查项 | 验证方法 |
|--------|---------|
| 所有窗体使用相同的 `AutoScaleMode` | 检查 `Theme.ApplyFormDefaults()` 是否被调用 |
| 所有窗体 `FormBorderStyle` 一致 | 检查是否为 `FixedSingle` |
| 所有窗体 `StartPosition` 一致 | 检查是否为 `CenterScreen` |
| 按钮样式跨窗体一致 | 对比不同窗体的"执行"和"取消"按钮外观 |
| 输入框样式跨窗体一致 | 对比不同窗体的 TextBox 边框和高度 |
| 标签样式跨窗体一致 | 对比不同窗体的 Label 字体和颜色 |
| 颜色方案跨窗体一致 | 截图对比不同窗体的整体色调 |

### 7.4 用户反馈

- [ ] 长时间操作是否提供了进度提示？
- [ ] 是否支持取消操作？（如 ESC 键中断）
- [ ] 错误提示是否友好且包含解决建议？
- [ ] 成功/失败是否有明确的状态反馈？

### 7.5 配置持久化

- [ ] 用户配置是否正确保存到文档属性或注册表？
- [ ] 配置读取失败时是否有合理的默认值？
- [ ] 文档级配置和全局配置的优先级是否正确？

---

## 八、项目特定规范（P1-P2 级）

### 8.1 服务层规范

- [ ] `Services` 类是否为 `static`？（当前项目约定）
- [ ] 服务方法是否纯函数化？（避免副作用，或明确文档化副作用）
- [ ] 跨服务调用是否通过参数传递而非全局状态？

### 8.2 表格操作规范

- [ ] 修改单元格前是否检查了表格结构？（列数、合并单元格）
- [ ] 编号操作是否支持纯文本和 SEQ 域两种格式？
- [ ] 图片插入后是否调整了单元格尺寸？

### 8.3 Excel 数据填充规范

- [ ] OLEDB 连接字符串是否正确区分 `.xlsx` 和 `.xls`？
- [ ] 数据读取后是否释放了 `DataTable`？
- [ ] 模板检测是否覆盖了所有支持的表格结构？

---

## 审查输出模板

```
## 审查结果摘要

- **审查文件**: `FileName.cs`
- **审查行数**: 100 行
- **问题总数**: X 个（P0: X, P1: X, P2: X, P3: X）

---

### [P1][逻辑/功能] 问题描述
**位置**: `FileName.cs:行号`
**问题**: 具体问题描述
**建议**: 修改建议
**修复代码**:
```csharp
// 修复后的代码
```

### [P2][代码规范] 问题描述
**位置**: `FileName.cs:行号`
**问题**: 具体问题描述
**建议**: 修改建议

---

### 维度检查清单

| 维度 | 状态 | 说明 |
| ---- | ---- | ---- |
| COM 生命周期 | ✅/❌ | 说明 |
| Word 状态管理 | ✅/❌ | 说明 |
| 语法编译 | ✅/❌ | 说明 |
| 逻辑正确性 | ✅/❌ | 说明 |
| 异常处理 | ✅/❌ | 说明 |
| 代码规范 | ✅/❌ | 说明 |
| 性能优化 | ✅/❌ | 说明 |
| 安全性 | ✅/❌ | 说明 |
| UI/UX | ✅/❌ | 说明 |
```

---

## 附录：常见反模式

### 反模式 1：裸 catch 块
```csharp
// ❌ 禁止
catch { }

// ✅ 至少记录日志
catch (Exception ex)
{
    Debug.WriteLine($"[ClassName] Error: {ex.Message}");
}
```

### 反模式 2：忘记恢复 Word 状态
```csharp
// ❌ 禁止
app.ScreenUpdating = false;
// ... 操作 ...
// 未恢复

// ✅ 使用 try-finally
bool wasUpdating = app.ScreenUpdating;
app.ScreenUpdating = false;
try { ... }
finally { app.ScreenUpdating = wasUpdating; }
```

### 反模式 3：Range 赋值后复用
```csharp
// ❌ 禁止
range.Text = "新内容";
range.Bold = 1; // Range 已失效

// ✅ 重新获取
range.Text = "新内容";
range = cell.Range;
range.Bold = 1;
```

### 反模式 4：循环中重复 COM 调用
```csharp
// ❌ 禁止
for (int i = 1; i <= table.Rows.Count; i++) { ... }

// ✅ 缓存
int count = table.Rows.Count;
for (int i = 1; i <= count; i++) { ... }
```

### 反模式 5：未验证的文件操作
```csharp
// ❌ 禁止
var files = Directory.GetFiles(path);

// ✅ 验证路径
if (Directory.Exists(path))
{
    var files = Directory.GetFiles(path);
}
```

### 反模式 6：UI 布局修改导致的高 DPI 异常
```csharp
// ❌ 错误：混合使用缩放和未缩放的尺寸计算
// 在 200% DPI 下，S(5)=10，但 CTRL_HEIGHT 可能=60，导致垂直位置计算错误
Location = new System.Drawing.Point(S(5), (CTRL_HEIGHT - S(20)) / 2),

// ❌ 错误：Panel 内控件使用绝对位置而非相对布局
// 修改 Panel 高度后，内部控件可能不再居中
Location = new System.Drawing.Point(S(70), (S(Theme.Layout.CtrlHeight) - S(20)) / 2),

// ✅ 正确：统一使用已缩放的布局常量，且计算方式一致
int yOffset = (CTRL_HEIGHT - S(20)) / 2;  // 所有值都基于已缩放的 CTRL_HEIGHT
Location = new System.Drawing.Point(S(70), yOffset),
```

### 反模式 7：跨窗体 UI 风格不一致
```csharp
// ❌ 错误：不同窗体创建控件方式不同，风格不一致
// InsertPhotosForm.cs - 使用 Theme
var btn1 = Theme.CreateButton("插入", Theme.ButtonStyle.Success);

// ExcelDataFillerForm.cs - 直接 new，缺少样式
var btn2 = new Button();
btn2.Text = "执行填充";
btn2.BackColor = Color.Green;  // 硬编码颜色

// ❌ 错误：不同窗体使用不同的边距
// InsertPhotosForm.cs
Location = new Point(S(15), topPos);

// ExcelDataFillerForm.cs
Location = new Point(MARGIN, topPos);  // MARGIN = S(20)，不一致

// ✅ 正确：所有窗体统一使用 Theme 工厂方法
var btn = Theme.CreateButton("执行", Theme.ButtonStyle.Success);

// ✅ 正确：统一使用 MARGIN 常量
Location = new Point(MARGIN, topPos);
```
