using System;
using System.Drawing;
using System.Windows.Forms;
using WordTools;

namespace WordTools.Forms
{
    /// <summary>
    /// 批量插图进度窗口（独立于 Word，不受 ScreenUpdating 影响）
    /// </summary>
    public partial class ProgressForm : Form
    {
        private Label lblStatus;
        private Label lblDetail;
        private ProgressBar progressBar;
        private Button btnCancel;

        public bool IsCancelled { get; private set; }

        public ProgressForm(int totalFiles)
        {
            IsCancelled = false;
            InitializeComponent(totalFiles);
        }

        private void InitializeComponent(int totalFiles)
        {
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "插入图片进度";
            this.BackColor = Theme.Colors.Background;
            // 使用 ClientSize 确保客户区足够大（不包含标题栏和边框）
            this.ClientSize = new Size(640, 235);
            // 保持进度窗口在最前面，避免被Word窗口遮挡
            this.TopMost = true;

            // 状态标签
            lblStatus = new Label
            {
                Location = new Point(20, 15),
                Size = new Size(600, 45),
                Font = new Font(Theme.Fonts.Default.FontFamily, 16, FontStyle.Bold),
                ForeColor = Theme.Colors.Text,
                TextAlign = ContentAlignment.MiddleCenter,
                Text = $"准备插入 {totalFiles} 张图片...",
                AutoSize = false
            };
            this.Controls.Add(lblStatus);

            // 进度条
            progressBar = new ProgressBar
            {
                Location = new Point(20, 70),
                Size = new Size(600, 25),
                Minimum = 0,
                Maximum = totalFiles,
                Value = 0
            };
            this.Controls.Add(progressBar);

            // 详情标签
            lblDetail = new Label
            {
                Location = new Point(20, 105),
                Size = new Size(600, 28),
                Font = Theme.Fonts.Default,
                ForeColor = Theme.Colors.TextLight,
                TextAlign = ContentAlignment.MiddleCenter,
                Text = "",
                AutoSize = false
            };
            this.Controls.Add(lblDetail);

            // 取消按钮
            btnCancel = new Button
            {
                Location = new Point(245, 148),
                Size = new Size(150, 40),
                Text = "取消",
                Font = new Font(Theme.Fonts.Default.FontFamily, 10, FontStyle.Bold),
                BackColor = Theme.Colors.Primary,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnCancel.FlatAppearance.BorderSize = 0;
            btnCancel.Click += (s, e) =>
            {
                IsCancelled = true;
                btnCancel.Enabled = false;
                btnCancel.Text = "正在取消...";
            };
            this.Controls.Add(btnCancel);
        }

        /// <summary>
        /// 更新进度（使用 BeginInvoke 避免阻塞插图线程）
        /// </summary>
        public void UpdateProgress(int current, int total, string currentFile, TimeSpan elapsed)
        {
            // 使用 BeginInvoke 异步更新 UI，避免阻塞插图线程
            this.BeginInvoke(new Action(() =>
            {
                int percent = total > 0 ? (int)((double)current / total * 100) : 0;
                double remainingSeconds = current > 0 ? elapsed.TotalSeconds / current * (total - current) : 0;

                lblStatus.Text = $"插入图片 {current}/{total} ({percent}%)";
                progressBar.Value = Math.Min(current, progressBar.Maximum);

                string shortFileName = currentFile.Length > 35
                    ? "..." + currentFile.Substring(currentFile.Length - 32)
                    : currentFile;
                lblDetail.Text = $"已用:{elapsed.TotalSeconds:F0}s 剩余:{remainingSeconds:F0}s | {shortFileName}";
            }));
        }

        /// <summary>
        /// 显示完成状态（使用 BeginInvoke 避免阻塞）
        /// </summary>
        public void ShowCompletion(int successCount, int failCount, double totalSeconds)
        {
            this.BeginInvoke(new Action(() =>
            {
                string timeInfo = totalSeconds >= 60
                    ? $"{(int)(totalSeconds / 60)}分{totalSeconds % 60:F1}秒"
                    : $"{totalSeconds:F1}秒";

                if (IsCancelled)
                {
                    lblStatus.Text = $"已取消 | 成功:{successCount} 失败:{failCount}";
                    lblStatus.ForeColor = Color.OrangeRed;
                }
                else
                {
                    lblStatus.Text = $"完成！成功:{successCount} 失败:{failCount}";
                    lblStatus.ForeColor = Color.ForestGreen;
                }

                lblDetail.Text = $"耗时: {timeInfo}";
                progressBar.Value = progressBar.Maximum;
                btnCancel.Text = "关闭";
                btnCancel.Enabled = true;
                btnCancel.Click -= btnCancel_Click;
                btnCancel.Click += (s, e) => this.Close();
            }));
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            // 由构造函数中的 lambda 处理
        }
    }
}
