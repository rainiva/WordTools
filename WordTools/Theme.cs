using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace WordTools
{
    /// <summary>
    /// 统一管理项目中所有窗体的 UI 样式
    /// </summary>
    public static class Theme
    {
        /// <summary>
        /// 颜色方案
        /// </summary>
        public static class Colors
        {
            /// <summary>背景色</summary>
            public static readonly Color Background = Color.FromArgb(245, 245, 245);

            /// <summary>主色调（按钮）</summary>
            public static readonly Color Primary = Color.FromArgb(74, 144, 217);

            /// <summary>成功/执行按钮</summary>
            public static readonly Color Success = Color.FromArgb(92, 184, 92);

            /// <summary>危险/错误</summary>
            public static readonly Color Danger = Color.FromArgb(217, 83, 79);

            /// <summary>主文本色</summary>
            public static readonly Color Text = Color.FromArgb(51, 51, 51);

            /// <summary>次要文本色</summary>
            public static readonly Color TextLight = Color.FromArgb(102, 102, 102);

            /// <summary>辅助文本色</summary>
            public static readonly Color TextSecondary = Color.FromArgb(85, 85, 85);

            /// <summary>边框/分隔线</summary>
            public static readonly Color Border = Color.FromArgb(204, 204, 204);

            /// <summary>默认按钮背景</summary>
            public static readonly Color ButtonDefault = Color.FromArgb(224, 224, 224);

            /// <summary>禁用输入框背景</summary>
            public static readonly Color InputDisabled = Color.FromArgb(240, 240, 240);

            /// <summary>禁用态文字</summary>
            public static readonly Color TextDisabled = Color.FromArgb(153, 153, 153);

            /// <summary>只读输入框背景</summary>
            public static readonly Color InputReadonly = Color.FromArgb(250, 250, 250);

            /// <summary>输入框默认背景</summary>
            public static readonly Color InputBackground = Color.White;

            /// <summary>禁用按钮背景</summary>
            public static readonly Color ButtonDisabled = Color.FromArgb(204, 204, 204);

            /// <summary>禁用按钮边框</summary>
            public static readonly Color ButtonDisabledBorder = Color.FromArgb(180, 180, 180);
        }

        /// <summary>
        /// 字体方案
        /// </summary>
        public static class Fonts
        {
            /// <summary>默认字体</summary>
            public static readonly Font Default = new Font("微软雅黑", 9);

            /// <summary>加粗字体</summary>
            public static readonly Font Bold = new Font("微软雅黑", 9, FontStyle.Bold);

            /// <summary>标题字体</summary>
            public static readonly Font Title = new Font("微软雅黑", 10, FontStyle.Bold);

            /// <summary>小字体</summary>
            public static readonly Font Small = new Font("微软雅黑", 8);

            /// <summary>等宽字体</summary>
            public static readonly Font Mono = new Font("Consolas", 9);
        }

        /// <summary>
        /// 布局常量（基准值，96 DPI 下的像素值）
        /// </summary>
        public static class Layout
        {
            /// <summary>窗体边距</summary>
            public const int Margin = 20;

            /// <summary>控件高度（输入框、按钮、复选框等）</summary>
            public const int CtrlHeight = 30;

            /// <summary>标签宽度</summary>
            public const int LabelWidth = 80;

            /// <summary>行间距（行顶到下一行顶的距离）</summary>
            public const int LineSpacing = 42;

            /// <summary>按钮宽度</summary>
            public const int ButtonWidth = 75;

            /// <summary>按钮高度</summary>
            public const int ButtonHeight = 32;

            /// <summary>控件间水平间距</summary>
            public const int Gap = 12;

            /// <summary>默认窗体宽度</summary>
            public const int FormWidth = 520;

            /// <summary>小窗体宽度</summary>
            public const int FormWidthSmall = 480;

            /// <summary>分隔线上方间距（前一行控件底部到分隔线）</summary>
            public const int DividerPaddingTop = 6;

            /// <summary>分隔线下方间距（分隔线到下一行控件顶部）</summary>
            public const int DividerPaddingBottom = 10;

            /// <summary>区域间距（主要区域分隔线前后的间距，比分隔线间距更大）</summary>
            public const int SectionSpacing = 18;

            /// <summary>区域标题到内容的间距</summary>
            public const int SectionTitleSpacing = 25;
        }

        /// <summary>
        /// 获取指定窗体的 DPI 缩放比例
        /// </summary>
        /// <param name="form">目标窗体</param>
        /// <returns>DPI 缩放比例</returns>
        public static float GetDpiScale(Form form)
        {
            using (Graphics g = form.CreateGraphics())
            {
                return g.DpiX / 96f;
            }
        }

        /// <summary>
        /// 按 DPI 缩放像素值
        /// </summary>
        /// <param name="value">原始像素值</param>
        /// <param name="dpiScale">DPI 缩放比例</param>
        /// <returns>缩放后的像素值</returns>
        public static int Scale(int value, float dpiScale)
        {
            return (int)(value * dpiScale);
        }

        /// <summary>
        /// 应用窗体默认样式
        /// </summary>
        /// <param name="form">目标窗体</param>
        /// <returns>DPI 缩放比例</returns>
        public static float ApplyFormDefaults(Form form)
        {
            form.AutoScaleMode = AutoScaleMode.None;
            form.BackColor = Colors.Background;
            form.Font = Fonts.Default;
            form.FormBorderStyle = FormBorderStyle.FixedSingle;
            form.MaximizeBox = false;
            form.ShowIcon = false;
            form.ShowInTaskbar = false;
            form.MinimizeBox = false;
            form.StartPosition = FormStartPosition.CenterScreen;
            return GetDpiScale(form);
        }

        /// <summary>
        /// 创建标准标签
        /// </summary>
        /// <param name="text">标签文本</param>
        /// <param name="bold">是否加粗</param>
        /// <returns>配置好的 Label 控件</returns>
        public static Label CreateLabel(string text, bool bold = false)
        {
            return new Label
            {
                Text = text,
                AutoSize = true,
                Font = bold ? Fonts.Bold : Fonts.Default,
                ForeColor = Colors.Text
            };
        }

        /// <summary>
        /// 创建标准文本框
        /// </summary>
        /// <param name="textAlign">文字对齐方式</param>
        /// <returns>配置好的 TextBox 控件</returns>
        public static TextBox CreateTextBox(HorizontalAlignment textAlign = HorizontalAlignment.Left)
        {
            var tb = new TextBox
            {
                Font = Fonts.Default,
                Multiline = true,       // 改为 true，允许显式控制高度
                WordWrap = false,       // 防止文字换行
                TextAlign = textAlign,
                BorderStyle = BorderStyle.Fixed3D  // 保持默认 3D 边框不变
            };
            tb.KeyPress += (s, e) => { if (e.KeyChar == '\r' || e.KeyChar == '\n') e.Handled = true; };
            return tb;
        }

        /// <summary>
        /// 输入框默认上下边距（用于精确控制高度）
        /// </summary>
        public const int TextBoxMargin = 3;

        /// <summary>
        /// 计算 TextBox 实际内容高度
        /// </summary>
        /// <param name="ctrlHeight">控件总高度</param>
        /// <returns>内容区域高度</returns>
        public static int GetTextBoxContentHeight(int ctrlHeight)
        {
            return ctrlHeight - TextBoxMargin * 2;
        }

        // Win32 API 用于调整 TextBox 文字垂直居中
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, ref RECT lParam);

        private const int EM_SETRECT = 0xB3;

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left, Top, Right, Bottom;
        }

        /// <summary>
        /// 设置 TextBox 文字垂直居中（需要在设置 Size 之后调用）
        /// </summary>
        public static void CenterTextVertically(TextBox tb)
        {
            if (tb == null || !tb.Multiline || tb.Height <= 0) return;

            int textHeight = System.Windows.Forms.TextRenderer.MeasureText("Ay中", tb.Font).Height;
            int topPadding = Math.Max(0, (tb.Height - textHeight) / 2);
            RECT rect = new RECT
            {
                Left = 2,
                Top = topPadding,
                Right = tb.Width - 2,
                Bottom = tb.Height - topPadding
            };
            // 需要确保 Handle 已创建
            var handle = tb.Handle;  // 强制创建 Handle
            SendMessage(handle, EM_SETRECT, IntPtr.Zero, ref rect);
        }

    }
}
