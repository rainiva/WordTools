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

        // 颜色方案
        private static readonly Color COLOR_BG = Color.FromArgb(245, 245, 245);
        private static readonly Color COLOR_PRIMARY = Color.FromArgb(74, 120, 217);
        private static readonly Color COLOR_SUCCESS = Color.FromArgb(92, 184, 92);
        private static readonly Color COLOR_TEXT_PRIMARY = Color.FromArgb(51, 51, 51);
        private static readonly Color COLOR_TEXT_SECONDARY = Color.FromArgb(85, 85, 85);
        private static readonly Color COLOR_DIVIDER = Color.FromArgb(204, 204, 204);

        // 布局常量
        private const int MARGIN = 15;
        private const int GAP = 12;
        private const int LINE_HEIGHT = 25;
        private const int BUTTON_HEIGHT = 30;

        public InsertPhotosForm(Application application)
        {
            _application = application;
            InitializeComponent();
            LoadConfiguration();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // 启用 DPI 缩放
            this.AutoScaleMode = AutoScaleMode.Dpi;
            this.AutoScaleDimensions = new SizeF(96F, 96F);

            // 窗体属性
            this.Text = "批量插图工具";
            this.Size = new Size(500, 320);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = COLOR_BG;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ShowIcon = false;
            this.ShowInTaskbar = false;

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
            this.ClientSize = new Size(480, currentTop + MARGIN);

            this.ResumeLayout(false);
        }

        #region 控件创建

        private void CreateFolderPathRow(ref int topPos)
        {
            // 标签
            var lblFolder = new Label
            {
                Text = "文件夹:",
                Location = new System.Drawing.Point(15, topPos + 4),
                Size = new Size(55, 18),
                Font = new System.Drawing.Font("Microsoft YaHei UI", 9F, FontStyle.Bold),
                ForeColor = COLOR_TEXT_SECONDARY
            };
            this.Controls.Add(lblFolder);

            // 文本框
            txtFolderPath = new TextBox
            {
                Location = new System.Drawing.Point(75, topPos),
                Size = new Size(300, LINE_HEIGHT),
                Font = new System.Drawing.Font("Microsoft YaHei UI", 9F)
            };
            this.Controls.Add(txtFolderPath);

            // 浏览按钮
            btnBrowseFolder = new Button
            {
                Text = "浏览...",
                Location = new System.Drawing.Point(385, topPos),
                Size = new Size(75, LINE_HEIGHT),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(224, 224, 224),
                ForeColor = COLOR_TEXT_PRIMARY,
                Font = new System.Drawing.Font("Microsoft YaHei UI", 9F),
                Cursor = Cursors.Hand
            };
            btnBrowseFolder.FlatAppearance.BorderColor = COLOR_DIVIDER;
            btnBrowseFolder.Click += BtnBrowseFolder_Click;
            this.Controls.Add(btnBrowseFolder);

            topPos += 38;
        }

        private void CreateHeightAndScopeRow(ref int topPos)
        {
            // 高度标签
            var lblHeight = new Label
            {
                Text = "高度:",
                Location = new System.Drawing.Point(15, topPos + 4),
                Size = new Size(42, 18),
                Font = new System.Drawing.Font("Microsoft YaHei UI", 9F, FontStyle.Bold),
                ForeColor = COLOR_TEXT_SECONDARY
            };
            this.Controls.Add(lblHeight);

            // 高度文本框
            txtImageHeight = new TextBox
            {
                Text = "",
                Location = new System.Drawing.Point(75, topPos),
                Size = new Size(55, LINE_HEIGHT),
                Font = new System.Drawing.Font("Microsoft YaHei UI", 9F),
                TextAlign = HorizontalAlignment.Center
            };
            this.Controls.Add(txtImageHeight);

            // 单位标签
            var lblUnit = new Label
            {
                Text = "cm",
                Location = new System.Drawing.Point(135, topPos + 4),
                Size = new Size(25, 18),
                Font = new System.Drawing.Font("Microsoft YaHei UI", 9F),
                ForeColor = Color.FromArgb(119, 119, 119)
            };
            this.Controls.Add(lblUnit);

            // 范围标签
            var lblScope = new Label
            {
                Text = "范围:",
                Location = new System.Drawing.Point(185, topPos + 4),
                Size = new Size(42, 18),
                Font = new System.Drawing.Font("Microsoft YaHei UI", 9F, FontStyle.Bold),
                ForeColor = COLOR_TEXT_SECONDARY
            };
            this.Controls.Add(lblScope);

            // 根目录复选框
            chkIncludeRoot = new CheckBox
            {
                Text = "根目录",
                Location = new System.Drawing.Point(230, topPos + 1),
                Size = new Size(75, 20),
                Checked = true,
                Font = new System.Drawing.Font("Microsoft YaHei UI", 9F),
                ForeColor = COLOR_TEXT_PRIMARY
            };
            this.Controls.Add(chkIncludeRoot);

            // 子目录复选框
            chkIncludeSubFolder = new CheckBox
            {
                Text = "子目录",
                Location = new System.Drawing.Point(315, topPos + 1),
                Size = new Size(75, 20),
                Checked = true,
                Font = new System.Drawing.Font("Microsoft YaHei UI", 9F),
                ForeColor = COLOR_TEXT_PRIMARY
            };
            this.Controls.Add(chkIncludeSubFolder);

            topPos += 35;
        }

        private void CreateDivider(ref int topPos)
        {
            var divider = new Label
            {
                Location = new System.Drawing.Point(MARGIN, topPos),
                Size = new Size(450, 1),
                BackColor = COLOR_DIVIDER,
                BorderStyle = BorderStyle.None
            };
            this.Controls.Add(divider);
            topPos += GAP;
        }

        private void CreateDescriptionOptionsRow(ref int topPos)
        {
            // 描述标签
            var lblDesc = new Label
            {
                Text = "描述:",
                Location = new System.Drawing.Point(15, topPos + 2),
                Size = new Size(42, 18),
                Font = new System.Drawing.Font("Microsoft YaHei UI", 9F, FontStyle.Bold),
                ForeColor = COLOR_TEXT_SECONDARY
            };
            this.Controls.Add(lblDesc);

            // 用 Panel 包裹描述选项的 RadioButton
            var pnlDescription = new Panel
            {
                Location = new System.Drawing.Point(70, topPos - 2),
                Size = new Size(210, 24),
                BackColor = COLOR_BG
            };
            this.Controls.Add(pnlDescription);

            // 手动描述
            optNeedDescription = new RadioButton
            {
                Text = "手动",
                Location = new System.Drawing.Point(5, 2),
                Size = new Size(60, 20),
                Checked = true,
                Font = new System.Drawing.Font("Microsoft YaHei UI", 9F),
                ForeColor = COLOR_TEXT_PRIMARY
            };
            optNeedDescription.CheckedChanged += DescriptionOption_CheckedChanged;
            pnlDescription.Controls.Add(optNeedDescription);

            // 无描述
            optNoDescription = new RadioButton
            {
                Text = "无",
                Location = new System.Drawing.Point(70, 2),
                Size = new Size(50, 20),
                Font = new System.Drawing.Font("Microsoft YaHei UI", 9F),
                ForeColor = COLOR_TEXT_PRIMARY
            };
            optNoDescription.CheckedChanged += DescriptionOption_CheckedChanged;
            pnlDescription.Controls.Add(optNoDescription);

            // 文件名描述
            optUseFilename = new RadioButton
            {
                Text = "文件名",
                Location = new System.Drawing.Point(125, 2),
                Size = new Size(75, 20),
                Font = new System.Drawing.Font("Microsoft YaHei UI", 9F),
                ForeColor = COLOR_TEXT_PRIMARY
            };
            optUseFilename.CheckedChanged += DescriptionOption_CheckedChanged;
            pnlDescription.Controls.Add(optUseFilename);

            // 自动编号复选框（不在 Panel 内，保持独立）
            chkAutoNumbering = new CheckBox
            {
                Text = "自动编号",
                Location = new System.Drawing.Point(285, topPos + 1),
                Size = new Size(90, 20),
                Enabled = true, // 手动模式下启用
                Font = new System.Drawing.Font("Microsoft YaHei UI", 9F),
                ForeColor = COLOR_TEXT_PRIMARY
            };
            this.Controls.Add(chkAutoNumbering);

            topPos += 32;
        }

        private void CreateAlignmentOptionsRow(ref int topPos)
        {
            // 对齐标签
            var lblAlign = new Label
            {
                Text = "对齐:",
                Location = new System.Drawing.Point(15, topPos + 2),
                Size = new Size(42, 18),
                Font = new System.Drawing.Font("Microsoft YaHei UI", 9F, FontStyle.Bold),
                ForeColor = COLOR_TEXT_SECONDARY
            };
            this.Controls.Add(lblAlign);

            // 用 Panel 包裹对齐选项，使其与描述选项的 RadioButton 互不影响
            var pnlAlignment = new Panel
            {
                Location = new System.Drawing.Point(70, topPos - 2),
                Size = new Size(200, 24),
                BackColor = COLOR_BG
            };
            this.Controls.Add(pnlAlignment);

            // 靠左
            optAlignLeft = new RadioButton
            {
                Text = "靠左",
                Location = new System.Drawing.Point(5, 2),
                Size = new Size(60, 20),
                Font = new System.Drawing.Font("Microsoft YaHei UI", 9F),
                ForeColor = COLOR_TEXT_PRIMARY
            };
            pnlAlignment.Controls.Add(optAlignLeft);

            // 居中（默认选中）
            optAlignCenter = new RadioButton
            {
                Text = "居中",
                Location = new System.Drawing.Point(70, 2),
                Size = new Size(60, 20),
                Checked = true,
                Font = new System.Drawing.Font("Microsoft YaHei UI", 9F),
                ForeColor = COLOR_TEXT_PRIMARY
            };
            pnlAlignment.Controls.Add(optAlignCenter);

            topPos += 32;
        }

        private void CreateActionButtonsRow(ref int topPos)
        {
            // 插入文件夹按钮
            btnInsertFromFolder = new Button
            {
                Text = "插入文件夹",
                Location = new System.Drawing.Point(15, topPos),
                Size = new Size(100, BUTTON_HEIGHT),
                FlatStyle = FlatStyle.Flat,
                BackColor = COLOR_SUCCESS,
                ForeColor = Color.White,
                Font = new System.Drawing.Font("Microsoft YaHei UI", 9F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnInsertFromFolder.FlatAppearance.BorderSize = 0;
            btnInsertFromFolder.Click += BtnInsertFromFolder_Click;
            this.Controls.Add(btnInsertFromFolder);

            // 选择文件按钮
            btnSelectFiles = new Button
            {
                Text = "选择文件",
                Location = new System.Drawing.Point(125, topPos),
                Size = new Size(100, BUTTON_HEIGHT),
                FlatStyle = FlatStyle.Flat,
                BackColor = COLOR_PRIMARY,
                ForeColor = Color.White,
                Font = new System.Drawing.Font("Microsoft YaHei UI", 9F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnSelectFiles.FlatAppearance.BorderSize = 0;
            btnSelectFiles.Click += BtnSelectFiles_Click;
            this.Controls.Add(btnSelectFiles);

            // 取消按钮（靠右对齐）
            btnCancel = new Button
            {
                Text = "取消",
                Location = new System.Drawing.Point(380, topPos),
                Size = new Size(80, BUTTON_HEIGHT),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(224, 224, 224),
                ForeColor = COLOR_TEXT_PRIMARY,
                Font = new System.Drawing.Font("Microsoft YaHei UI", 9F, FontStyle.Bold),
                Cursor = Cursors.Hand,
                DialogResult = DialogResult.Cancel
            };
            btnCancel.FlatAppearance.BorderColor = COLOR_DIVIDER;
            this.Controls.Add(btnCancel);

            topPos += BUTTON_HEIGHT + 10;
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
                chkAutoNumbering.ForeColor = Color.FromArgb(153, 153, 153);
            }
            else
            {
                chkAutoNumbering.Enabled = true;
                chkAutoNumbering.ForeColor = COLOR_TEXT_PRIMARY;
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
