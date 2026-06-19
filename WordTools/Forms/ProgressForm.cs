using System;
using System.Drawing;
using System.Windows.Forms;
using WordTools;
using WordTools.Services;

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
        private readonly ProgressFormStateController _stateController = new ProgressFormStateController();

        public bool IsCancelled
        {
            get { return _stateController.IsCancelled; }
        }

        /// <summary>
        /// 用户在操作进行期间尝试关闭窗口（点击 X），已转换为取消请求。
        /// </summary>
        public bool IsCloseRequestedByUser { get; private set; }

        /// <summary>
        /// 由服务代码主动触发关闭（而非用户点击 X）。
        /// </summary>
        public bool IsServiceClosing { get; set; }

        public ProgressForm(int totalFiles)
        {
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
            this.ClientSize = new Size(640, 235);
            this.TopMost = true;

            lblStatus = new Label
            {
                Location = new Point(20, 15),
                Size = new Size(600, 45),
                Font = new Font(Theme.Fonts.Default.FontFamily, 16, FontStyle.Bold),
                ForeColor = Theme.Colors.Text,
                TextAlign = ContentAlignment.MiddleCenter,
                Text = string.Format("准备插入 {0} 张图片...", totalFiles),
                AutoSize = false
            };
            this.Controls.Add(lblStatus);

            progressBar = new ProgressBar
            {
                Location = new Point(20, 70),
                Size = new Size(600, 25),
                Minimum = 0,
                Maximum = totalFiles,
                Value = 0
            };
            this.Controls.Add(progressBar);

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

            btnCancel = new Button
            {
                Location = new Point(245, 148),
                Size = new Size(150, 40),
                Text = _stateController.ButtonText,
                Font = new Font(Theme.Fonts.Default.FontFamily, 10, FontStyle.Bold),
                BackColor = Theme.Colors.Primary,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnCancel.FlatAppearance.BorderSize = 0;
            btnCancel.Click += btnCancel_Click;
            this.Controls.Add(btnCancel);

            this.FormClosing += ProgressForm_FormClosing;
        }

        /// <summary>
        /// 用户在操作进行期间点击右上角 X 时，视为取消请求，禁止直接关闭窗口。
        /// </summary>
        private void ProgressForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 服务代码主动关闭时不拦截，也不提示
            if (IsServiceClosing)
            {
                return;
            }

            if (e.CloseReason == CloseReason.UserClosing && !_stateController.IsCompleted)
            {
                e.Cancel = true;

                // 已经在取消过程中，仅提示用户
                if (_stateController.IsCancelled)
                {
                    MessageBox.Show("正在取消当前操作，请稍候...", "提示",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var result = MessageBox.Show("操作正在进行中，是否取消？", "确认取消",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result != DialogResult.Yes)
                {
                    return;
                }

                IsCloseRequestedByUser = true;
                _stateController.HandleButtonClick();
                btnCancel.Enabled = _stateController.IsButtonEnabled;
                btnCancel.Text = _stateController.ButtonText;
            }
        }

        /// <summary>
        /// 更新进度（使用 BeginInvoke 避免阻塞插图线程）
        /// </summary>
        public void UpdateProgress(int current, int total, string currentFile, TimeSpan elapsed)
        {
            this.BeginInvoke(new Action(() =>
            {
                int percent = total > 0 ? (int)((double)current / total * 100) : 0;
                double remainingSeconds = current > 0 ? elapsed.TotalSeconds / current * (total - current) : 0;

                lblStatus.Text = string.Format("插入图片 {0}/{1} ({2}%)", current, total, percent);
                progressBar.Value = Math.Min(current, progressBar.Maximum);

                string shortFileName = currentFile.Length > 35
                    ? "..." + currentFile.Substring(currentFile.Length - 32)
                    : currentFile;
                lblDetail.Text = string.Format("已用:{0:F0}s 剩余:{1:F0}s | {2}", elapsed.TotalSeconds, remainingSeconds, shortFileName);
            }));
        }

        /// <summary>
        /// 显示完成状态。使用 Invoke 同步执行，确保 MarkCompleted 在关闭窗口前生效，
        /// 避免代码触发的 FormClosing 误判为进行中关闭。
        /// </summary>
        public void ShowCompletion(int successCount, int failCount, double totalSeconds)
        {
            this.Invoke(new Action(() =>
            {
                string timeInfo = totalSeconds >= 60
                    ? string.Format("{0}分{1:F1}秒", (int)(totalSeconds / 60), totalSeconds % 60)
                    : string.Format("{0:F1}秒", totalSeconds);

                if (IsCancelled)
                {
                    lblStatus.Text = string.Format("已取消 | 成功:{0} 失败:{1}", successCount, failCount);
                    lblStatus.ForeColor = Color.OrangeRed;
                }
                else
                {
                    lblStatus.Text = string.Format("完成！成功:{0} 失败:{1}", successCount, failCount);
                    lblStatus.ForeColor = Color.ForestGreen;
                }

                lblDetail.Text = string.Format("耗时: {0}", timeInfo);
                progressBar.Value = progressBar.Maximum;
                _stateController.MarkCompleted();
                btnCancel.Text = _stateController.ButtonText;
                btnCancel.Enabled = _stateController.IsButtonEnabled;
            }));
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            var action = _stateController.HandleButtonClick();
            if (action == ProgressButtonAction.CancelRequested)
            {
                btnCancel.Enabled = _stateController.IsButtonEnabled;
                btnCancel.Text = _stateController.ButtonText;
                return;
            }

            if (action == ProgressButtonAction.CloseRequested)
            {
                this.Close();
            }
        }
    }
}
