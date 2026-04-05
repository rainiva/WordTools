using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using WordTools.Services;
using Application = Microsoft.Office.Interop.Word.Application;
using CheckBox = System.Windows.Forms.CheckBox;
using RadioButton = System.Windows.Forms.RadioButton;
using TextBox = System.Windows.Forms.TextBox;
using Button = System.Windows.Forms.Button;
using Label = System.Windows.Forms.Label;
using Theme = WordTools.Theme;

namespace WordTools.Forms
{
    /// <summary>
    /// 批量插图工具主窗体
    /// </summary>
    public partial class InsertPhotosForm : Form
    {
        private readonly Application _application;
        
        // 控件
        private TextBox txtFolderPath;
        private TextBox txtImageHeight;
        private Button btnBrowseFolder;
        private Button btnInsertFromFolder;
        private Button btnSelectFiles;
        private Button btnCancel;
        private RadioButton optNeedDescription;
        private RadioButton optNoDescription;
        private RadioButton optUseFilename;
        private CheckBox chkIncludeRoot;
        private CheckBox chkIncludeSubFolder;
        private CheckBox chkAutoNumbering;
        private RadioButton optAlignLeft;
        private RadioButton optAlignCenter;

        // DPI 缩放
        private float dpiScale;
        private int S(int value) { return Theme.Scale(value, dpiScale); }

        // 预计算布局常量
        private int MARGIN;
        private int CTRL_HEIGHT;
        private int LINE_SPACING;
        private int FORM_WIDTH;

        public InsertPhotosForm(Application application)
        {
            _application = application;
            InitializeComponent();
            LoadConfiguration();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // 应用 DPI 缩放和窗体默认样式
            dpiScale = Theme.ApplyFormDefaults(this);

            // 预计算布局常量
            MARGIN = S(Theme.Layout.Margin);
            CTRL_HEIGHT = S(Theme.Layout.CtrlHeight);
            LINE_SPACING = S(Theme.Layout.LineSpacing);
            FORM_WIDTH = S(Theme.Layout.FormWidthSmall);

            // 窗体属性
            this.Text = "批量插图工具";
            this.ClientSize = new Size(FORM_WIDTH, S(320));

            int currentTop = MARGIN;

            // 文件夹路径行
            CreateFolderPathRow(ref currentTop);

            // 图片高度和范围行
            CreateHeightAndScopeRow(ref currentTop);

            // 分隔线
            CreateDivider(ref currentTop);

            // 描述选项行
            CreateDescriptionOptionsRow(ref currentTop);

            // 编号对齐行
            CreateAlignmentOptionsRow(ref currentTop);

            // 分隔线
            CreateDivider(ref currentTop);

            // 操作按钮行
            CreateActionButtonsRow(ref currentTop);

            // 调整窗体高度
            this.ClientSize = new Size(FORM_WIDTH, currentTop + MARGIN);

            this.ResumeLayout(false);
        }

        #region 控件创建

        private void CreateFolderPathRow(ref int topPos)
        {
            // 标签（高度与输入框一致，文字垂直居中）
            var lblFolder = new Label
            {
                Text = "文件夹:",
                Location = new System.Drawing.Point(S(15), topPos),
                Size = new Size(S(55), CTRL_HEIGHT),
                Font = Theme.Fonts.Bold,
                ForeColor = Theme.Colors.Text,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            };
            this.Controls.Add(lblFolder);

            // 文本框（使用单行模式，恢复默认样式）
            txtFolderPath = Theme.CreateTextBox();
            txtFolderPath.Location = new System.Drawing.Point(S(75), topPos);
            txtFolderPath.Size = new Size(S(300), CTRL_HEIGHT);
            Theme.CenterTextVertically(txtFolderPath);
            this.Controls.Add(txtFolderPath);

            // 浏览按钮（高度与文本框自动高度对齐）
            btnBrowseFolder = Theme.CreateButton("浏览...", Theme.ButtonStyle.Default);
            btnBrowseFolder.Location = new System.Drawing.Point(S(385), topPos);
            btnBrowseFolder.Size = new Size(S(75), CTRL_HEIGHT);
            btnBrowseFolder.Click += BtnBrowseFolder_Click;
            this.Controls.Add(btnBrowseFolder);

            topPos += LINE_SPACING;
        }

        private void CreateHeightAndScopeRow(ref int topPos)
        {
            // 高度标签（高度与输入框一致，文字垂直居中）
            var lblHeight = new Label
            {
                Text = "高度:",
                Location = new System.Drawing.Point(S(15), topPos),
                Size = new Size(S(42), CTRL_HEIGHT),
                Font = Theme.Fonts.Bold,
                ForeColor = Theme.Colors.Text,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            };
            this.Controls.Add(lblHeight);

            // 高度文本框（使用单行模式，恢复默认样式）
            txtImageHeight = Theme.CreateTextBox(HorizontalAlignment.Center);
            txtImageHeight.Text = "";
            txtImageHeight.Location = new System.Drawing.Point(S(75), topPos);
            txtImageHeight.Size = new Size(S(80), CTRL_HEIGHT);
            txtImageHeight.TextAlign = HorizontalAlignment.Center;
            Theme.CenterTextVertically(txtImageHeight);
            this.Controls.Add(txtImageHeight);

            // 单位标签（高度与输入框一致，文字垂直居中）
            var lblUnit = new Label
            {
                Text = "cm",
                Location = new System.Drawing.Point(S(160), topPos),
                Size = new Size(S(25), CTRL_HEIGHT),
                Font = Theme.Fonts.Default,
                ForeColor = Theme.Colors.TextLight,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            };
            this.Controls.Add(lblUnit);

            // 范围标签（高度与输入框一致，文字垂直居中）
            var lblScope = new Label
            {
                Text = "范围:",
                Location = new System.Drawing.Point(S(200), topPos),
                Size = new Size(S(42), CTRL_HEIGHT),
                Font = Theme.Fonts.Bold,
                ForeColor = Theme.Colors.Text,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            };
            this.Controls.Add(lblScope);

            // 根目录复选框
            chkIncludeRoot = new CheckBox
            {
                Text = "根目录",
                Location = new System.Drawing.Point(S(248), topPos),
                Size = new Size(S(75), CTRL_HEIGHT),
                Checked = true,
                Font = Theme.Fonts.Default,
                ForeColor = Theme.Colors.Text
            };
            this.Controls.Add(chkIncludeRoot);

            // 子目录复选框
            chkIncludeSubFolder = new CheckBox
            {
                Text = "子目录",
                Location = new System.Drawing.Point(S(330), topPos),
                Size = new Size(S(75), CTRL_HEIGHT),
                Checked = true,
                Font = Theme.Fonts.Default,
                ForeColor = Theme.Colors.Text
            };
            this.Controls.Add(chkIncludeSubFolder);

            topPos += S(Theme.Layout.CtrlHeight + Theme.Layout.DividerPaddingTop);
        }

        private void CreateDivider(ref int topPos)
        {
            var divider = Theme.CreateDivider(FORM_WIDTH - MARGIN * 2);
            divider.Location = new System.Drawing.Point(MARGIN, topPos);
            this.Controls.Add(divider);
            topPos += 1 + S(Theme.Layout.DividerPaddingBottom);
        }

        private void CreateDescriptionOptionsRow(ref int topPos)
        {
            // 描述标签（高度与输入框一致，文字垂直居中）
            var lblDesc = new Label
            {
                Text = "描述:",
                Location = new System.Drawing.Point(S(15), topPos),
                Size = new Size(S(42), CTRL_HEIGHT),
                Font = Theme.Fonts.Bold,
                ForeColor = Theme.Colors.Text,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            };
            this.Controls.Add(lblDesc);

            // 用 Panel 包裹描述选项的 RadioButton
            var pnlDescription = new Panel
            {
                Location = new System.Drawing.Point(S(70), topPos),
                Size = new Size(S(210), CTRL_HEIGHT),
                BackColor = Theme.Colors.Background
            };
            this.Controls.Add(pnlDescription);

            // 手动描述
            optNeedDescription = new RadioButton
            {
                Text = "手动",
                Location = new System.Drawing.Point(S(5), (CTRL_HEIGHT - S(20)) / 2),
                Size = new Size(S(60), S(20)),
                Checked = true,
                Font = Theme.Fonts.Default,
                ForeColor = Theme.Colors.Text
            };
            optNeedDescription.CheckedChanged += DescriptionOption_CheckedChanged;
            pnlDescription.Controls.Add(optNeedDescription);

            // 无描述
            optNoDescription = new RadioButton
            {
                Text = "无",
                Location = new System.Drawing.Point(S(70), (S(Theme.Layout.CtrlHeight) - S(20)) / 2),
                Size = new Size(S(50), S(20)),
                Font = Theme.Fonts.Default,
                ForeColor = Theme.Colors.Text
            };
            optNoDescription.CheckedChanged += DescriptionOption_CheckedChanged;
            pnlDescription.Controls.Add(optNoDescription);

            // 文件名描述
            optUseFilename = new RadioButton
            {
                Text = "文件名",
                Location = new System.Drawing.Point(S(125), (S(Theme.Layout.CtrlHeight) - S(20)) / 2),
                Size = new Size(S(75), S(20)),
                Font = Theme.Fonts.Default,
                ForeColor = Theme.Colors.Text
            };
            optUseFilename.CheckedChanged += DescriptionOption_CheckedChanged;
            pnlDescription.Controls.Add(optUseFilename);

            // 自动编号复选框（不在 Panel 内，保持独立）
            chkAutoNumbering = new CheckBox
            {
                Text = "自动编号",
                Location = new System.Drawing.Point(S(285), topPos),
                Size = new Size(S(90), CTRL_HEIGHT),
                Enabled = true, // 手动模式下启用
                Font = Theme.Fonts.Default,
                ForeColor = Theme.Colors.Text
            };
            this.Controls.Add(chkAutoNumbering);

            topPos += LINE_SPACING;
        }

        private void CreateAlignmentOptionsRow(ref int topPos)
        {
            // 对齐标签（高度与输入框一致，文字垂直居中）
            var lblAlign = new Label
            {
                Text = "对齐:",
                Location = new System.Drawing.Point(S(15), topPos),
                Size = new Size(S(42), CTRL_HEIGHT),
                Font = Theme.Fonts.Bold,
                ForeColor = Theme.Colors.Text,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            };
            this.Controls.Add(lblAlign);

            // 用 Panel 包裹对齐选项，使其与描述选项的 RadioButton 互不影响
            var pnlAlignment = new Panel
            {
                Location = new System.Drawing.Point(S(70), topPos),
                Size = new Size(S(200), CTRL_HEIGHT),
                BackColor = Theme.Colors.Background
            };
            this.Controls.Add(pnlAlignment);

            // 靠左
            optAlignLeft = new RadioButton
            {
                Text = "靠左",
                Location = new System.Drawing.Point(S(5), (CTRL_HEIGHT - S(20)) / 2),
                Size = new Size(S(60), S(20)),
                Font = Theme.Fonts.Default,
                ForeColor = Theme.Colors.Text
            };
            pnlAlignment.Controls.Add(optAlignLeft);

            // 居中（默认选中）
            optAlignCenter = new RadioButton
            {
                Text = "居中",
                Location = new System.Drawing.Point(S(70), (S(Theme.Layout.CtrlHeight) - S(20)) / 2),
                Size = new Size(S(60), S(20)),
                Checked = true,
                Font = Theme.Fonts.Default,
                ForeColor = Theme.Colors.Text
            };
            pnlAlignment.Controls.Add(optAlignCenter);

            topPos += LINE_SPACING;
        }

        private void CreateActionButtonsRow(ref int topPos)
        {
            // 插入文件夹按钮
            btnInsertFromFolder = Theme.CreateButton("插入文件夹", Theme.ButtonStyle.Success);
            btnInsertFromFolder.Location = new System.Drawing.Point(S(15), topPos);
            btnInsertFromFolder.Size = new Size(S(100), CTRL_HEIGHT);
            btnInsertFromFolder.Click += BtnInsertFromFolder_Click;
            this.Controls.Add(btnInsertFromFolder);

            // 选择文件按钮
            btnSelectFiles = Theme.CreateButton("选择文件", Theme.ButtonStyle.Primary);
            btnSelectFiles.Location = new System.Drawing.Point(S(125), topPos);
            btnSelectFiles.Size = new Size(S(100), CTRL_HEIGHT);
            btnSelectFiles.Click += BtnSelectFiles_Click;
            this.Controls.Add(btnSelectFiles);

            // 取消按钮（靠右对齐）
            btnCancel = Theme.CreateButton("取消", Theme.ButtonStyle.Default);
            btnCancel.Location = new System.Drawing.Point(S(380), topPos);
            btnCancel.Size = new Size(S(80), CTRL_HEIGHT);
            btnCancel.DialogResult = DialogResult.Cancel;
            this.Controls.Add(btnCancel);

            topPos += LINE_SPACING;
        }

        #endregion

        #region 配置加载和保存

        private void LoadConfiguration()
        {
            try
            {
                var doc = _application != null ? _application.ActiveDocument : null;
                txtImageHeight.Text = ConfigService.GetLastImageHeightCM(doc);
                txtFolderPath.Text = ConfigService.GetLastFolderPath(doc);
                
                bool needDesc = ConfigService.GetNeedDescription(doc);
                bool useFilename = ConfigService.GetUseFilenameAsDescription(doc);
                
                if (useFilename)
                {
                    optUseFilename.Checked = true;
                }
                else if (needDesc)
                {
                    optNeedDescription.Checked = true;
                }
                else
                {
                    optNoDescription.Checked = true;
                }

                chkIncludeRoot.Checked = ConfigService.GetIncludeRootImages(doc);
                chkIncludeSubFolder.Checked = ConfigService.GetIncludeSubFolderImages(doc);
                chkAutoNumbering.Checked = ConfigService.GetAutoNumbering();

                int alignment = ConfigService.GetNumberAlignment();
                if (alignment == 1)
                {
                    optAlignLeft.Checked = true;
                }
                else
                {
                    optAlignCenter.Checked = true;
                }

                UpdateAutoNumberingState();
            }
            catch
            {
                // 使用默认值
            }
        }

        private void SaveConfiguration()
        {
            try
            {
                var doc = _application != null ? _application.ActiveDocument : null;

                // 保存高度（允许为空，空表示不限制）
                ConfigService.SaveLastImageHeightCM(txtImageHeight.Text.Trim(), doc);
                
                ConfigService.SaveLastFolderPath(txtFolderPath.Text, doc);
                ConfigService.SaveNeedDescription(optNeedDescription.Checked, doc);
                ConfigService.SaveUseFilenameAsDescription(optUseFilename.Checked, doc);
                ConfigService.SaveIncludeRootImages(chkIncludeRoot.Checked, doc);
                ConfigService.SaveIncludeSubFolderImages(chkIncludeSubFolder.Checked, doc);
                ConfigService.SaveAutoNumbering(chkAutoNumbering.Checked);
                ConfigService.SaveNumberAlignment(optAlignCenter.Checked ? 2 : 1);
            }
            catch
            {
                // 忽略保存错误
            }
        }

        #endregion

        #region 事件处理

        private void DescriptionOption_CheckedChanged(object sender, EventArgs e)
        {
            UpdateAutoNumberingState();
        }

        private void UpdateAutoNumberingState()
        {
            // 无描述模式下禁用自动编号
            if (optNoDescription.Checked)
            {
                chkAutoNumbering.Enabled = false;
                chkAutoNumbering.Checked = false;
                chkAutoNumbering.ForeColor = Theme.Colors.TextDisabled;
            }
            else
            {
                chkAutoNumbering.Enabled = true;
                chkAutoNumbering.ForeColor = Theme.Colors.Text;
            }
        }

        private void BtnBrowseFolder_Click(object sender, EventArgs e)
        {
            string lastPath = ConfigService.GetLastFolderPath();
            if (!string.IsNullOrEmpty(txtFolderPath.Text))
            {
                lastPath = txtFolderPath.Text;
            }

            string folderPath = FileService.SelectFolder("请选择文件夹...", lastPath);
            if (!string.IsNullOrEmpty(folderPath))
            {
                txtFolderPath.Text = folderPath;
                ConfigService.SaveLastFolderPath(folderPath);
            }
        }

        private void BtnInsertFromFolder_Click(object sender, EventArgs e)
        {
            try
            {
                string folderPath = txtFolderPath.Text;

                if (string.IsNullOrEmpty(folderPath))
                {
                    MessageBox.Show("请选择文件夹！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                float minHeight;
                if (!ImageService.ValidateAndConvertHeight(txtImageHeight.Text, out minHeight))
                {
                    MessageBox.Show("输入的高度无效，请输入大于 0 的数字。", "提示", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                SaveConfiguration();

                bool needDescription = optNeedDescription.Checked;
                bool useFileNameAsDescription = optUseFilename.Checked;
                bool includeRootImages = chkIncludeRoot.Checked;
                bool includeSubFolderImages = chkIncludeSubFolder.Checked;
                bool needAutoNumbering = chkAutoNumbering.Checked;
                int numberAlignment = optAlignCenter.Checked ? 2 : 1; // 2=居中, 1=靠左

                this.Hide();
                System.Windows.Forms.Application.DoEvents();

                var progressService = new ProgressService(_application);
                progressService.InsertPhotosWithProgress(
                    folderPath, minHeight, needDescription,
                    useFileNameAsDescription, includeRootImages, includeSubFolderImages,
                    needAutoNumbering, numberAlignment);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("操作失败: {0}", ex.Message), "错误", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnSelectFiles_Click(object sender, EventArgs e)
        {
            try
            {
                string lastPath = ConfigService.GetLastFolderPath();
                var selectedFiles = FileService.SelectImageFiles("请选择图片文件...", lastPath);

                if (selectedFiles == null || selectedFiles.Length == 0)
                {
                    return;
                }

                // 保存最后的文件夹路径
                ConfigService.SaveLastFolderPath(FileService.GetParentFolder(selectedFiles[0]));

                float minHeight;
                if (!ImageService.ValidateAndConvertHeight(txtImageHeight.Text, out minHeight))
                {
                    MessageBox.Show("输入的高度无效，请输入大于 0 的数字。", "提示", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                SaveConfiguration();

                bool needDescription = optNeedDescription.Checked;
                bool useFileNameAsDescription = optUseFilename.Checked;
                bool needAutoNumbering = chkAutoNumbering.Checked;
                int numberAlignment = optAlignCenter.Checked ? 2 : 1; // 2=居中, 1=靠左

                this.Hide();
                System.Windows.Forms.Application.DoEvents();

                var progressService = new ProgressService(_application);
                progressService.InsertSelectedPhotosWithProgress(
                    selectedFiles, minHeight, needDescription,
                    useFileNameAsDescription, needAutoNumbering, numberAlignment);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("操作失败: {0}", ex.Message), "错误", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion
    }
}
