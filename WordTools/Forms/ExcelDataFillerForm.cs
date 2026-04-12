using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using WordTools.Services;
using Theme = WordTools.Theme;

namespace WordTools.Forms
{
    /// <summary>
    /// Excel数据填充工具窗体
    /// </summary>
    public partial class ExcelDataFillerForm : Form
    {
        private EDF_DataFillerService fillerService;

        // 控件
        private TextBox txtExcelPath;
        private Button btnBrowse;
        private TextBox txtAnchorField;
        private TextBox txtTargetColumn;
        private CheckBox chkReplaceSampleSize;
        private TextBox txtSampleSizeColumn;
        private Label lblSampleSizeColumn;
        private TextBox txtStatus;
        private Button btnExecute;
        private Button btnCancel;

        // 布局常量（运行时按 DPI 缩放）
        private int MARGIN;
        private int CTRL_HEIGHT;
        private int LABEL_WIDTH;
        private int TEXTBOX_WIDTH;
        private int BUTTON_WIDTH;
        private int LINE_SPACING;
        private int FORM_WIDTH;
        private float dpiScale = 1f;

        public ExcelDataFillerForm()
        {
            InitializeComponent();
            fillerService = new EDF_DataFillerService();
        }

        private void InitializeComponent()
        {
            this.Text = "Excel数据填充工具";

            // 应用窗体默认样式并获取 DPI 缩放比例
            dpiScale = Theme.ApplyFormDefaults(this);
            this.CancelButton = btnCancel;

            // 按比例调整所有尺寸
            MARGIN = S(Theme.Layout.Margin);
            CTRL_HEIGHT = S(Theme.Layout.CtrlHeight);
            LABEL_WIDTH = S(Theme.Layout.LabelWidth);
            TEXTBOX_WIDTH = S(340);
            BUTTON_WIDTH = S(Theme.Layout.ButtonWidth);
            LINE_SPACING = S(Theme.Layout.LineSpacing);
            FORM_WIDTH = S(Theme.Layout.FormWidth);

            this.ClientSize = new Size(FORM_WIDTH, S(460));

            int topPos = MARGIN;
            int inputLeft = MARGIN + LABEL_WIDTH;
            int inputWidth = FORM_WIDTH - inputLeft - MARGIN;

            // ===== 第1行：Excel文件路径 =====
            Label lblExcelPath = new Label();
            lblExcelPath.Text = "Excel文件：";
            lblExcelPath.Location = new Point(MARGIN, topPos);
            lblExcelPath.Size = new Size(LABEL_WIDTH, CTRL_HEIGHT);
            lblExcelPath.Font = Theme.Fonts.Bold;
            lblExcelPath.ForeColor = Theme.Colors.Text;
            lblExcelPath.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.Controls.Add(lblExcelPath);

            txtExcelPath = Theme.CreateTextBox();
            txtExcelPath.Location = new Point(inputLeft, topPos);
            txtExcelPath.Size = new Size(inputWidth - BUTTON_WIDTH - S(10), CTRL_HEIGHT);
            Theme.CenterTextVertically(txtExcelPath);
            this.Controls.Add(txtExcelPath);

            // 浏览按钮：高度与输入框自动高度对齐
            btnBrowse = Theme.CreateButton("浏览...", Theme.ButtonStyle.Default);
            btnBrowse.Size = new Size(BUTTON_WIDTH, CTRL_HEIGHT);
            btnBrowse.Location = new Point(FORM_WIDTH - MARGIN - BUTTON_WIDTH, topPos);
            btnBrowse.Click += btnBrowse_Click;
            this.Controls.Add(btnBrowse);

            topPos += LINE_SPACING;

            // ===== 第2行：锚定字段 =====
            Label lblAnchorField = new Label();
            lblAnchorField.Text = "锚定字段：";
            lblAnchorField.Location = new Point(MARGIN, topPos);
            lblAnchorField.Size = new Size(LABEL_WIDTH, CTRL_HEIGHT);
            lblAnchorField.Font = Theme.Fonts.Bold;
            lblAnchorField.ForeColor = Theme.Colors.Text;
            lblAnchorField.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.Controls.Add(lblAnchorField);

            txtAnchorField = Theme.CreateTextBox();
            txtAnchorField.Location = new Point(inputLeft, topPos);
            txtAnchorField.Size = new Size(inputWidth, CTRL_HEIGHT);
            Theme.CenterTextVertically(txtAnchorField);
            this.Controls.Add(txtAnchorField);

            topPos += LINE_SPACING;

            // ===== 第3行：目标列 =====
            Label lblTargetColumn = new Label();
            lblTargetColumn.Text = "目标列：";
            lblTargetColumn.Location = new Point(MARGIN, topPos);
            lblTargetColumn.Size = new Size(LABEL_WIDTH, CTRL_HEIGHT);
            lblTargetColumn.Font = Theme.Fonts.Bold;
            lblTargetColumn.ForeColor = Theme.Colors.Text;
            lblTargetColumn.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.Controls.Add(lblTargetColumn);

            txtTargetColumn = Theme.CreateTextBox(HorizontalAlignment.Center);
            txtTargetColumn.Location = new Point(inputLeft, topPos);
            txtTargetColumn.Size = new Size(S(60), CTRL_HEIGHT);
            txtTargetColumn.Text = "2";
            txtTargetColumn.TextAlign = HorizontalAlignment.Center;
            Theme.CenterTextVertically(txtTargetColumn);
            this.Controls.Add(txtTargetColumn);

            Label lblTargetHint = new Label();
            lblTargetHint.Text = "(数据填充到Word表格的第几列)";
            lblTargetHint.Location = new Point(inputLeft + S(70), topPos);
            lblTargetHint.Size = new Size(S(200), CTRL_HEIGHT);
            lblTargetHint.ForeColor = Theme.Colors.TextLight;
            lblTargetHint.Font = Theme.Fonts.Small;
            lblTargetHint.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.Controls.Add(lblTargetHint);

            topPos += LINE_SPACING;

            // ===== 第4行：Sample Size选项 =====
            chkReplaceSampleSize = new CheckBox();
            chkReplaceSampleSize.Text = "替换Sample Size数量";
            chkReplaceSampleSize.Location = new Point(inputLeft, topPos);
            chkReplaceSampleSize.Size = new Size(S(160), CTRL_HEIGHT);
            chkReplaceSampleSize.ForeColor = Theme.Colors.Text;
            chkReplaceSampleSize.CheckedChanged += chkReplaceSampleSize_CheckedChanged;
            this.Controls.Add(chkReplaceSampleSize);

            // 第4行后半部分：所在列（留足间距）
            lblSampleSizeColumn = new Label();
            lblSampleSizeColumn.Text = "所在列：";
            lblSampleSizeColumn.Location = new Point(inputLeft + S(180), topPos);
            lblSampleSizeColumn.Size = new Size(S(65), CTRL_HEIGHT);
            lblSampleSizeColumn.Enabled = false;
            lblSampleSizeColumn.ForeColor = Theme.Colors.TextLight;
            lblSampleSizeColumn.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.Controls.Add(lblSampleSizeColumn);

            txtSampleSizeColumn = Theme.CreateTextBox(HorizontalAlignment.Center);
            txtSampleSizeColumn.Location = new Point(inputLeft + S(250), topPos);
            txtSampleSizeColumn.Size = new Size(S(50), CTRL_HEIGHT);
            txtSampleSizeColumn.Text = "B";
            txtSampleSizeColumn.Enabled = false;
            txtSampleSizeColumn.BackColor = Theme.Colors.InputDisabled;
            txtSampleSizeColumn.TextAlign = HorizontalAlignment.Center;
            Theme.CenterTextVertically(txtSampleSizeColumn);
            this.Controls.Add(txtSampleSizeColumn);

            topPos += S(Theme.Layout.CtrlHeight + Theme.Layout.SectionSpacing);

            // ===== 分隔线 =====
            Label separator = Theme.CreateDivider(FORM_WIDTH - MARGIN * 2);
            separator.Location = new Point(MARGIN, topPos);
            this.Controls.Add(separator);

            topPos += S(Theme.Layout.SectionSpacing);

            // ===== 状态显示区域标题 =====
            Label lblStatusTitle = new Label();
            lblStatusTitle.Text = "状态信息：";
            lblStatusTitle.Location = new Point(MARGIN, topPos);
            lblStatusTitle.AutoSize = true;
            lblStatusTitle.Font = Theme.Fonts.Title;
            lblStatusTitle.ForeColor = Theme.Colors.Primary;
            this.Controls.Add(lblStatusTitle);

            topPos += S(Theme.Layout.SectionTitleSpacing);

            // ===== 状态显示区域（多行文本框）=====
            txtStatus = new TextBox();
            txtStatus.Location = new Point(MARGIN, topPos);
            txtStatus.Size = new Size(FORM_WIDTH - MARGIN * 2, S(100));
            txtStatus.Multiline = true;
            txtStatus.ScrollBars = ScrollBars.Vertical;
            txtStatus.BackColor = Theme.Colors.InputReadonly;
            txtStatus.BorderStyle = BorderStyle.FixedSingle;
            txtStatus.Font = Theme.Fonts.Mono;
            txtStatus.Text = "等待执行...\r\n请选择Excel文件并点击执行填充";
            txtStatus.ReadOnly = true;
            this.Controls.Add(txtStatus);

            topPos += S(100) + S(Theme.Layout.Margin);

            // ===== 按钮区域（居中）=====
            int btnWidth = S(110);
            int btnHeight = S(Theme.Layout.CtrlHeight);
            int btnGap = S(Theme.Layout.Margin);
            int totalBtnWidth = btnWidth * 2 + btnGap;
            int btnStartX = (FORM_WIDTH - totalBtnWidth) / 2;

            btnExecute = Theme.CreateButton("执行填充", Theme.ButtonStyle.Success);
            btnExecute.Location = new Point(btnStartX, topPos);
            btnExecute.Size = new Size(btnWidth, btnHeight);
            btnExecute.Click += btnExecute_Click;
            this.Controls.Add(btnExecute);

            btnCancel = Theme.CreateButton("取消", Theme.ButtonStyle.Default);
            btnCancel.Location = new Point(btnStartX + btnWidth + btnGap, topPos);
            btnCancel.Size = new Size(btnWidth, btnHeight);
            btnCancel.DialogResult = DialogResult.Cancel;
            this.Controls.Add(btnCancel);

            // 根据实际内容调整窗体高度
            this.ClientSize = new Size(FORM_WIDTH, topPos + btnHeight + MARGIN);

            // 加载默认配置
            LoadDefaultConfig();
        }

        /// <summary>
        /// DPI 缩放辅助方法，将设计时像素值按当前 DPI 缩放
        /// </summary>
        private int S(int value)
        {
            return Theme.Scale(value, dpiScale);
        }

        /// <summary>
        /// 加载默认配置
        /// </summary>
        private void LoadDefaultConfig()
        {
            try
            {
                var doc = GetActiveDocument();
                txtExcelPath.Text = ConfigService.GetEdfExcelPath(doc);
                txtAnchorField.Text = ConfigService.GetEdfAnchorField(doc);
                txtTargetColumn.Text = ConfigService.GetEdfTargetColumn(doc);
                txtSampleSizeColumn.Text = ConfigService.GetEdfSampleSizeColumn(doc);
                chkReplaceSampleSize.Checked = ConfigService.GetEdfReplaceSampleSize(doc);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"LoadDefaultConfig error: {ex.Message}");
            }
        }

        /// <summary>
        /// 获取当前活动文档
        /// </summary>
        private Microsoft.Office.Interop.Word.Document GetActiveDocument()
        {
            try
            {
                return Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 浏览按钮点击事件
        /// </summary>
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "选择Excel文件";
            openFileDialog.Filter = "Excel文件|*.xlsx;*.xls";
            openFileDialog.Multiselect = false;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                txtExcelPath.Text = openFileDialog.FileName;
                AppendStatus("已选择Excel文件: " + openFileDialog.FileName);
            }
        }

        /// <summary>
        /// Sample Size复选框变化事件
        /// </summary>
        private void chkReplaceSampleSize_CheckedChanged(object sender, EventArgs e)
        {
            txtSampleSizeColumn.Enabled = chkReplaceSampleSize.Checked;
            lblSampleSizeColumn.Enabled = chkReplaceSampleSize.Checked;
            txtSampleSizeColumn.BackColor = chkReplaceSampleSize.Checked ? Theme.Colors.InputBackground : Theme.Colors.InputDisabled;
        }

        /// <summary>
        /// 执行按钮点击事件
        /// </summary>
        private void btnExecute_Click(object sender, EventArgs e)
        {
            // 验证输入
            if (!ValidateInput())
            {
                UpdateStatus("输入验证失败，请检查配置");
                return;
            }

            // 保存配置
            SaveCurrentConfig();

            // 收集参数
            string excelPath = txtExcelPath.Text.Trim();
            string anchorField = txtAnchorField.Text.Trim();
            int targetColumn;
            if (!int.TryParse(txtTargetColumn.Text.Trim(), out targetColumn))
            {
                MessageBox.Show("请输入有效的目标列数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string sampleSizeColumn = txtSampleSizeColumn.Text.Trim();
            bool replaceSampleSize = chkReplaceSampleSize.Checked;

            // 禁用按钮，防止重复点击
            btnExecute.Enabled = false;
            btnCancel.Enabled = false;

            UpdateStatus("开始执行...\r\n正在检测表格结构...");

            // 同步执行填充（COM对象必须在STA线程上访问）
            try
            {
                fillerService.ExecuteFilling(
                    excelPath,
                    anchorField,
                    targetColumn,
                    sampleSizeColumn,
                    replaceSampleSize,
                    (message) =>
                    {
                        AppendStatus(message);
                    });

                AppendStatus("填充完成！");
                btnExecute.Enabled = true;
                btnCancel.Enabled = true;
            }
            catch (Exception ex)
            {
                AppendStatus("错误: " + ex.Message);
                MessageBox.Show(string.Format("执行出错: {0}", ex.Message), "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                btnExecute.Enabled = true;
                btnCancel.Enabled = true;
            }
        }

        /// <summary>
        /// 验证输入
        /// </summary>
        private bool ValidateInput()
        {
            if (string.IsNullOrWhiteSpace(txtExcelPath.Text))
            {
                MessageBox.Show("请选择Excel文件！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtExcelPath.Focus();
                return false;
            }

            if (string.IsNullOrWhiteSpace(txtAnchorField.Text))
            {
                MessageBox.Show("请输入锚定字段！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtAnchorField.Focus();
                return false;
            }

            if (string.IsNullOrWhiteSpace(txtTargetColumn.Text) || !int.TryParse(txtTargetColumn.Text, out int _))
            {
                MessageBox.Show("请输入有效的目标列数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtTargetColumn.Focus();
                return false;
            }

            if (chkReplaceSampleSize.Checked && string.IsNullOrWhiteSpace(txtSampleSizeColumn.Text))
            {
                MessageBox.Show("请输入Sample Size所在列！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtSampleSizeColumn.Focus();
                return false;
            }

            return true;
        }

        /// <summary>
        /// 保存当前配置
        /// </summary>
        private void SaveCurrentConfig()
        {
            try
            {
                var doc = GetActiveDocument();
                ConfigService.SaveEdfExcelPath(txtExcelPath.Text.Trim(), doc);
                ConfigService.SaveEdfAnchorField(txtAnchorField.Text.Trim(), doc);
                ConfigService.SaveEdfTargetColumn(txtTargetColumn.Text.Trim(), doc);
                ConfigService.SaveEdfSampleSizeColumn(txtSampleSizeColumn.Text.Trim(), doc);
                ConfigService.SaveEdfReplaceSampleSize(chkReplaceSampleSize.Checked, doc);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"SaveCurrentConfig error: {ex.Message}");
            }
        }

        /// <summary>
        /// 更新状态显示
        /// </summary>
        private void UpdateStatus(string message)
        {
            txtStatus.Text = message;
            Application.DoEvents();
        }

        /// <summary>
        /// 追加状态信息
        /// </summary>
        private void AppendStatus(string message)
        {
            txtStatus.Text += "\r\n" + message;
            txtStatus.SelectionStart = txtStatus.Text.Length;
            txtStatus.ScrollToCaret();
            Application.DoEvents();
        }
    }
}
