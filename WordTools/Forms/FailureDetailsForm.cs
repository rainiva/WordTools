using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace WordTools.Forms
{
    public class FailureDetailsForm : Form
    {
        private readonly List<(string fileName, string errorReason)> _failedFiles;
        private readonly List<int> _mergedCellRows;
        private readonly List<string> _overwriteWarnings;

        public FailureDetailsForm(
            List<(string fileName, string errorReason)> failedFiles,
            List<int> mergedCellRows = null,
            List<string> overwriteWarnings = null)
        {
            _failedFiles = failedFiles ?? new List<(string, string)>();
            _mergedCellRows = mergedCellRows ?? new List<int>();
            _overwriteWarnings = overwriteWarnings ?? new List<string>();
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            SuspendLayout();

            AutoScaleMode = AutoScaleMode.Font;
            Text = "详情";
            ClientSize = new Size(520, 420);
            StartPosition = FormStartPosition.CenterScreen;
            MinimizeBox = false;
            MaximizeBox = false;
            FormBorderStyle = FormBorderStyle.Sizable;
            MinimumSize = new Size(420, 320);

            Size buttonSize = GetButtonSize();
            int buttonPanelHeight = buttonSize.Height + 32;

            var txtDetails = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Both,
                Font = new Font("Consolas", 9.75F, FontStyle.Regular),
                Location = new Point(12, 12),
                Size = new Size(ClientSize.Width - 24, ClientSize.Height - buttonPanelHeight - 12),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                BackColor = SystemColors.Window,
                ForeColor = Color.FromArgb(60, 60, 60)
            };

            txtDetails.Text = BuildDetailsText();
            Controls.Add(txtDetails);

            var pnlButtons = new FlowLayoutPanel
            {
                Height = buttonPanelHeight,
                Dock = DockStyle.Bottom,
                Padding = new Padding(12, 12, 12, 12),
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false
            };

            var btnClose = new Button
            {
                Text = "关闭窗口",
                Size = buttonSize,
                Margin = new Padding(8, 0, 0, 0)
            };
            btnClose.Click += (s, e) => Close();
            pnlButtons.Controls.Add(btnClose);

            var btnCopy = new Button
            {
                Text = "复制详情",
                Size = buttonSize,
                Margin = new Padding(8, 0, 0, 0)
            };
            btnCopy.Click += (s, e) =>
            {
                try
                {
                    Clipboard.SetText(txtDetails.Text);
                    MessageBox.Show("已复制到剪贴板", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(string.Format("复制失败: {0}", ex.Message), "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
            pnlButtons.Controls.Add(btnCopy);

            Controls.Add(pnlButtons);

            ResumeLayout(false);
        }

        private Size GetButtonSize()
        {
            int minWidth = 132;
            int minHeight = 40;
            int width = Math.Max(minWidth, MeasureButtonWidth("复制详情"));
            width = Math.Max(width, MeasureButtonWidth("关闭窗口"));
            return new Size(width, minHeight);
        }

        private int MeasureButtonWidth(string text)
        {
            return TextRenderer.MeasureText(text, Font).Width + 28;
        }

        private string BuildDetailsText()
        {
            var sb = new StringBuilder();

            if (_failedFiles.Count > 0)
            {
                sb.AppendLine(string.Format("共 {0} 个文件插入失败：", _failedFiles.Count));
                sb.AppendLine();
                foreach (var item in _failedFiles)
                {
                    sb.AppendLine(string.Format("{0}: {1}", item.fileName, item.errorReason));
                }
            }

            if (_mergedCellRows.Count > 0)
            {
                AppendSectionSeparator(sb);
                sb.AppendLine(string.Format("共 {0} 处合并单元格已自动绕开：", _mergedCellRows.Count));
                sb.AppendLine();
                foreach (int row in _mergedCellRows)
                {
                    sb.AppendLine(string.Format("第 {0} 行", row));
                }
            }

            if (_overwriteWarnings.Count > 0)
            {
                AppendSectionSeparator(sb);
                sb.AppendLine(string.Format("共 {0} 处覆盖插图提示：", _overwriteWarnings.Count));
                sb.AppendLine();
                foreach (string warning in _overwriteWarnings)
                {
                    sb.AppendLine(warning);
                }
            }

            return sb.ToString();
        }

        private static void AppendSectionSeparator(StringBuilder sb)
        {
            if (sb.Length == 0)
            {
                return;
            }

            sb.AppendLine();
            sb.AppendLine("================================");
            sb.AppendLine();
        }
    }
}
