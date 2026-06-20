using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using WordTools.Services;
using Application = Microsoft.Office.Interop.Word.Application;
using Button = System.Windows.Forms.Button;
using CheckBox = System.Windows.Forms.CheckBox;
using DrawingPoint = System.Drawing.Point;
using Label = System.Windows.Forms.Label;
using RadioButton = System.Windows.Forms.RadioButton;
using TextBox = System.Windows.Forms.TextBox;
using Theme = WordTools.Theme;

namespace WordTools.Forms
{
    public enum InsertPhotosRequestMode
    {
        Folder,
        SelectedFiles
    }

    public sealed class InsertPhotosRequest
    {
        public InsertPhotosRequestMode Mode { get; set; }
        public string FolderPath { get; set; }
        public string[] SelectedFiles { get; set; }
        public float MinHeight { get; set; }
        public bool NeedDescription { get; set; }
        public bool UseFileNameAsDescription { get; set; }
        public bool UseFolderNameAsDescription { get; set; }
        public bool IncludeRootImages { get; set; }
        public bool IncludeSubFolderImages { get; set; }
        public bool NeedAutoNumbering { get; set; }
        public int NumberAlignment { get; set; }
        public int NumberPosition { get; set; }
    }

    /// <summary>
    /// 批量插图工具主窗体
    /// </summary>
    public partial class InsertPhotosForm : Form
    {
        private readonly Application _application;
        public InsertPhotosRequest PendingRequest { get; private set; }

        private TextBox txtFolderPath;
        private TextBox txtImageHeight;
        private Button btnBrowseFolder;
        private Button btnInsertFromFolder;
        private Button btnSelectFiles;
        private Button btnCancel;
        private RadioButton optNeedDescription;
        private RadioButton optNoDescription;
        private RadioButton optUseFilename;
        private RadioButton optUseFolderName;
        private CheckBox chkIncludeRoot;
        private CheckBox chkIncludeSubFolder;
        private CheckBox chkAutoNumbering;
        private RadioButton optNumberBeforeDesc;
        private RadioButton optNumberAfterDesc;
        private RadioButton optAlignLeft;
        private RadioButton optAlignCenter;

        private float dpiScale;
        private int MARGIN;
        private int CTRL_HEIGHT;
        private int LINE_SPACING;
        private int FORM_WIDTH;

        private int S(int value)
        {
            return Theme.Scale(value, dpiScale);
        }

        public InsertPhotosForm(Application application)
        {
            _application = application;
            InitializeComponent();
            LoadConfiguration();
        }

        private void InitializeComponent()
        {
            SuspendLayout();

            dpiScale = Theme.ApplyFormDefaults(this);
            MARGIN = S(Theme.Layout.Margin);
            CTRL_HEIGHT = S(Theme.Layout.CtrlHeight);
            LINE_SPACING = S(Theme.Layout.LineSpacing);
            FORM_WIDTH = S(Theme.Layout.FormWidthSmall);

            Text = "批量插图工具";
            Name = "InsertPhotosForm";
            ClientSize = new Size(FORM_WIDTH, S(400));

            int currentTop = MARGIN;
            CreateFolderPathRow(ref currentTop);
            CreateHeightAndScopeRow(ref currentTop);
            CreateDivider(ref currentTop);
            CreateDescriptionOptionsRow(ref currentTop);
            CreateNumberPositionRow(ref currentTop);
            CreateAlignmentOptionsRow(ref currentTop);
            CreateDivider(ref currentTop);
            CreateActionButtonsRow(ref currentTop);

            ClientSize = new Size(FORM_WIDTH, currentTop + MARGIN);
            ResumeLayout(false);
        }

        #region Layout

        private void CreateFolderPathRow(ref int topPos)
        {
            var lblFolder = new Label
            {
                Text = "文件夹:",
                Location = new DrawingPoint(S(15), topPos),
                Size = new Size(S(55), CTRL_HEIGHT),
                Font = Theme.Fonts.Bold,
                ForeColor = Theme.Colors.Text,
                TextAlign = ContentAlignment.MiddleLeft
            };
            Controls.Add(lblFolder);

            txtFolderPath = Theme.CreateTextBox();
            txtFolderPath.Location = new DrawingPoint(S(75), topPos);
            txtFolderPath.Size = new Size(S(300), CTRL_HEIGHT);
            Theme.CenterTextVertically(txtFolderPath);
            Controls.Add(txtFolderPath);

            btnBrowseFolder = UiToolkit.CreateButton("浏览...", UiToolkit.ButtonStyle.Default);
            btnBrowseFolder.Location = new DrawingPoint(S(385), topPos);
            btnBrowseFolder.Size = new Size(S(75), CTRL_HEIGHT);
            btnBrowseFolder.Click += BtnBrowseFolder_Click;
            Controls.Add(btnBrowseFolder);

            topPos += LINE_SPACING;
        }

        private void CreateHeightAndScopeRow(ref int topPos)
        {
            var lblHeight = new Label
            {
                Text = "高度:",
                Location = new DrawingPoint(S(15), topPos),
                Size = new Size(S(42), CTRL_HEIGHT),
                Font = Theme.Fonts.Bold,
                ForeColor = Theme.Colors.Text,
                TextAlign = ContentAlignment.MiddleLeft
            };
            Controls.Add(lblHeight);

            txtImageHeight = Theme.CreateTextBox(HorizontalAlignment.Center);
            txtImageHeight.Location = new DrawingPoint(S(75), topPos);
            txtImageHeight.Size = new Size(S(80), CTRL_HEIGHT);
            txtImageHeight.TextAlign = HorizontalAlignment.Center;
            Theme.CenterTextVertically(txtImageHeight);
            Controls.Add(txtImageHeight);

            var lblUnit = new Label
            {
                Text = "cm",
                Location = new DrawingPoint(S(160), topPos),
                Size = new Size(S(25), CTRL_HEIGHT),
                Font = Theme.Fonts.Default,
                ForeColor = Theme.Colors.TextLight,
                TextAlign = ContentAlignment.MiddleLeft
            };
            Controls.Add(lblUnit);

            var lblScope = new Label
            {
                Text = "范围:",
                Location = new DrawingPoint(S(200), topPos),
                Size = new Size(S(42), CTRL_HEIGHT),
                Font = Theme.Fonts.Bold,
                ForeColor = Theme.Colors.Text,
                TextAlign = ContentAlignment.MiddleLeft
            };
            Controls.Add(lblScope);

            chkIncludeRoot = new CheckBox
            {
                Text = "根目录",
                Location = new DrawingPoint(S(248), topPos),
                Size = new Size(S(75), CTRL_HEIGHT),
                Checked = true,
                Font = Theme.Fonts.Default,
                ForeColor = Theme.Colors.Text,
                BackColor = Theme.Colors.Background
            };
            Controls.Add(chkIncludeRoot);

            chkIncludeSubFolder = new CheckBox
            {
                Text = "子目录",
                Location = new DrawingPoint(S(330), topPos),
                Size = new Size(S(75), CTRL_HEIGHT),
                Checked = true,
                Font = Theme.Fonts.Default,
                ForeColor = Theme.Colors.Text,
                BackColor = Theme.Colors.Background
            };
            Controls.Add(chkIncludeSubFolder);

            topPos += S(Theme.Layout.CtrlHeight + Theme.Layout.DividerPaddingTop);
        }

        private void CreateDivider(ref int topPos)
        {
            var divider = UiToolkit.CreateDivider(FORM_WIDTH - MARGIN * 2);
            divider.Location = new DrawingPoint(MARGIN, topPos);
            Controls.Add(divider);
            topPos += 1 + S(Theme.Layout.DividerPaddingBottom);
        }

        private void CreateDescriptionOptionsRow(ref int topPos)
        {
            var lblDesc = new Label
            {
                Text = "描述:",
                Location = new DrawingPoint(S(15), topPos),
                Size = new Size(S(42), CTRL_HEIGHT),
                Font = Theme.Fonts.Bold,
                ForeColor = Theme.Colors.Text,
                TextAlign = ContentAlignment.MiddleLeft
            };
            Controls.Add(lblDesc);

            var pnlDescription = new Panel
            {
                Location = new DrawingPoint(S(70), topPos),
                Size = new Size(S(300), CTRL_HEIGHT),
                BackColor = Theme.Colors.Background
            };
            Controls.Add(pnlDescription);

            optNeedDescription = new RadioButton
            {
                Text = "手动",
                Location = new DrawingPoint(S(5), (CTRL_HEIGHT - S(20)) / 2),
                Size = new Size(S(60), S(20)),
                Checked = true,
                Font = Theme.Fonts.Default,
                ForeColor = Theme.Colors.Text,
                BackColor = Theme.Colors.Background
            };
            optNeedDescription.CheckedChanged += DescriptionOption_CheckedChanged;
            pnlDescription.Controls.Add(optNeedDescription);

            optNoDescription = new RadioButton
            {
                Text = "无",
                Location = new DrawingPoint(S(70), (CTRL_HEIGHT - S(20)) / 2),
                Size = new Size(S(50), S(20)),
                Font = Theme.Fonts.Default,
                ForeColor = Theme.Colors.Text,
                BackColor = Theme.Colors.Background
            };
            optNoDescription.CheckedChanged += DescriptionOption_CheckedChanged;
            pnlDescription.Controls.Add(optNoDescription);

            optUseFilename = new RadioButton
            {
                Text = "文件名",
                Location = new DrawingPoint(S(125), (CTRL_HEIGHT - S(20)) / 2),
                Size = new Size(S(75), S(20)),
                Font = Theme.Fonts.Default,
                ForeColor = Theme.Colors.Text,
                BackColor = Theme.Colors.Background
            };
            optUseFilename.CheckedChanged += DescriptionOption_CheckedChanged;
            pnlDescription.Controls.Add(optUseFilename);

            optUseFolderName = new RadioButton
            {
                Text = "文件夹名",
                Location = new DrawingPoint(S(205), (CTRL_HEIGHT - S(20)) / 2),
                Size = new Size(S(90), S(20)),
                Font = Theme.Fonts.Default,
                ForeColor = Theme.Colors.Text,
                BackColor = Theme.Colors.Background
            };
            optUseFolderName.CheckedChanged += DescriptionOption_CheckedChanged;
            pnlDescription.Controls.Add(optUseFolderName);

            topPos += LINE_SPACING;
        }

        private void CreateNumberPositionRow(ref int topPos)
        {
            chkAutoNumbering = new CheckBox
            {
                Text = "自动编号",
                Location = new DrawingPoint(S(15), topPos),
                Size = new Size(S(90), CTRL_HEIGHT),
                Font = Theme.Fonts.Default,
                ForeColor = Theme.Colors.Text,
                BackColor = Theme.Colors.Background
            };
            chkAutoNumbering.CheckedChanged += AutoNumbering_CheckedChanged;
            Controls.Add(chkAutoNumbering);

            var pnlNumberPos = new Panel
            {
                Location = new DrawingPoint(S(110), topPos),
                Size = new Size(S(200), CTRL_HEIGHT),
                BackColor = Theme.Colors.Background
            };
            Controls.Add(pnlNumberPos);

            optNumberBeforeDesc = new RadioButton
            {
                Text = "编号在前",
                Location = new DrawingPoint(S(5), (CTRL_HEIGHT - S(20)) / 2),
                Size = new Size(S(85), S(20)),
                Checked = true,
                Font = Theme.Fonts.Default,
                ForeColor = Theme.Colors.Text,
                BackColor = Theme.Colors.Background
            };
            pnlNumberPos.Controls.Add(optNumberBeforeDesc);

            optNumberAfterDesc = new RadioButton
            {
                Text = "编号在后",
                Location = new DrawingPoint(S(95), (CTRL_HEIGHT - S(20)) / 2),
                Size = new Size(S(85), S(20)),
                Font = Theme.Fonts.Default,
                ForeColor = Theme.Colors.Text,
                BackColor = Theme.Colors.Background
            };
            pnlNumberPos.Controls.Add(optNumberAfterDesc);

            topPos += LINE_SPACING;
        }

        private void CreateAlignmentOptionsRow(ref int topPos)
        {
            var lblAlign = new Label
            {
                Text = "对齐:",
                Location = new DrawingPoint(S(15), topPos),
                Size = new Size(S(42), CTRL_HEIGHT),
                Font = Theme.Fonts.Bold,
                ForeColor = Theme.Colors.Text,
                TextAlign = ContentAlignment.MiddleLeft
            };
            Controls.Add(lblAlign);

            var pnlAlignment = new Panel
            {
                Location = new DrawingPoint(S(70), topPos),
                Size = new Size(S(200), CTRL_HEIGHT),
                BackColor = Theme.Colors.Background
            };
            Controls.Add(pnlAlignment);

            optAlignLeft = new RadioButton
            {
                Text = "靠左",
                Location = new DrawingPoint(S(5), (CTRL_HEIGHT - S(20)) / 2),
                Size = new Size(S(60), S(20)),
                Font = Theme.Fonts.Default,
                ForeColor = Theme.Colors.Text,
                BackColor = Theme.Colors.Background
            };
            pnlAlignment.Controls.Add(optAlignLeft);

            optAlignCenter = new RadioButton
            {
                Text = "居中",
                Location = new DrawingPoint(S(70), (CTRL_HEIGHT - S(20)) / 2),
                Size = new Size(S(60), S(20)),
                Checked = true,
                Font = Theme.Fonts.Default,
                ForeColor = Theme.Colors.Text,
                BackColor = Theme.Colors.Background
            };
            pnlAlignment.Controls.Add(optAlignCenter);

            topPos += LINE_SPACING;
        }

        private void CreateActionButtonsRow(ref int topPos)
        {
            btnInsertFromFolder = UiToolkit.CreateButton("插入文件夹", UiToolkit.ButtonStyle.Success);
            btnInsertFromFolder.Location = new DrawingPoint(S(15), topPos);
            btnInsertFromFolder.Size = new Size(S(100), CTRL_HEIGHT);
            btnInsertFromFolder.Click += BtnInsertFromFolder_Click;
            Controls.Add(btnInsertFromFolder);

            btnSelectFiles = UiToolkit.CreateButton("选择文件", UiToolkit.ButtonStyle.Primary);
            btnSelectFiles.Name = "btnSelectFiles";
            btnSelectFiles.AccessibleName = "btnSelectFiles";
            btnSelectFiles.Location = new DrawingPoint(S(125), topPos);
            btnSelectFiles.Size = new Size(S(100), CTRL_HEIGHT);
            btnSelectFiles.Click += BtnSelectFiles_Click;
            Controls.Add(btnSelectFiles);

            btnCancel = UiToolkit.CreateButton("取消", UiToolkit.ButtonStyle.Default);
            btnCancel.Location = new DrawingPoint(S(380), topPos);
            btnCancel.Size = new Size(S(80), CTRL_HEIGHT);
            btnCancel.DialogResult = DialogResult.Cancel;
            Controls.Add(btnCancel);

            topPos += LINE_SPACING;
        }

        #endregion

        #region Configuration

        private void LoadConfiguration()
        {
            try
            {
                var doc = _application != null ? _application.ActiveDocument : null;
                txtImageHeight.Text = ConfigService.GetLastImageHeightCM(doc);
                string lastFolderPath = ConfigService.GetLastFolderPath(doc);
                txtFolderPath.Text = Directory.Exists(lastFolderPath) ? lastFolderPath : string.Empty;

                bool needDesc = ConfigService.GetNeedDescription(doc);
                bool useFilename = ConfigService.GetUseFilenameAsDescription(doc);
                bool useFolderName = ConfigService.GetUseFolderNameAsDescription(doc);

                if (useFolderName)
                {
                    optUseFolderName.Checked = true;
                }
                else if (useFilename)
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
                optAlignLeft.Checked = ConfigService.GetNumberAlignment() == 1;
                optAlignCenter.Checked = !optAlignLeft.Checked;
                optNumberAfterDesc.Checked = ConfigService.GetNumberPosition() == 2;
                optNumberBeforeDesc.Checked = !optNumberAfterDesc.Checked;

                UpdateAutoNumberingState();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[InsertPhotosForm] LoadConfiguration error: {ex.Message}");
            }
        }

        private void SaveConfiguration()
        {
            try
            {
                var doc = _application != null ? _application.ActiveDocument : null;
                ConfigService.SaveLastImageHeightCM(txtImageHeight.Text.Trim(), doc);
                string folderPath = txtFolderPath.Text.Trim();
                if (Directory.Exists(folderPath))
                {
                    ConfigService.SaveLastFolderPath(folderPath, doc);
                }
                ConfigService.SaveNeedDescription(optNeedDescription.Checked, doc);
                ConfigService.SaveUseFilenameAsDescription(optUseFilename.Checked, doc);
                ConfigService.SaveUseFolderNameAsDescription(optUseFolderName.Checked, doc);
                ConfigService.SaveNumberPosition(optNumberAfterDesc.Checked ? 2 : 1);
                ConfigService.SaveIncludeRootImages(chkIncludeRoot.Checked, doc);
                ConfigService.SaveIncludeSubFolderImages(chkIncludeSubFolder.Checked, doc);
                ConfigService.SaveAutoNumbering(chkAutoNumbering.Checked);
                ConfigService.SaveNumberAlignment(optAlignCenter.Checked ? 2 : 1);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[InsertPhotosForm] SaveConfiguration error: {ex.Message}");
            }
        }

        #endregion

        #region Events

        private void DescriptionOption_CheckedChanged(object sender, EventArgs e)
        {
            UpdateAutoNumberingState();
        }

        private void AutoNumbering_CheckedChanged(object sender, EventArgs e)
        {
            UpdateAutoNumberingState();
        }

        private void UpdateAutoNumberingState()
        {
            if (optNoDescription.Checked)
            {
                chkAutoNumbering.Enabled = false;
                chkAutoNumbering.Checked = false;
                chkAutoNumbering.ForeColor = Theme.Colors.TextDisabled;
                optNumberBeforeDesc.Enabled = false;
                optNumberAfterDesc.Enabled = false;
                optNumberBeforeDesc.ForeColor = Theme.Colors.TextDisabled;
                optNumberAfterDesc.ForeColor = Theme.Colors.TextDisabled;
                return;
            }

            chkAutoNumbering.Enabled = true;
            chkAutoNumbering.ForeColor = Theme.Colors.Text;

            bool numberingEnabled = chkAutoNumbering.Checked;
            optNumberBeforeDesc.Enabled = numberingEnabled;
            optNumberAfterDesc.Enabled = numberingEnabled;
            optNumberBeforeDesc.ForeColor = numberingEnabled ? Theme.Colors.Text : Theme.Colors.TextDisabled;
            optNumberAfterDesc.ForeColor = numberingEnabled ? Theme.Colors.Text : Theme.Colors.TextDisabled;
        }

        private void BtnBrowseFolder_Click(object sender, EventArgs e)
        {
            string lastPath = ConfigService.GetLastFolderPath();
            if (!string.IsNullOrEmpty(txtFolderPath.Text))
            {
                lastPath = txtFolderPath.Text;
            }

            string folderPath = SelectFolder("请选择文件夹...", lastPath);
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
                string folderPath = txtFolderPath.Text.Trim();
                if (string.IsNullOrEmpty(folderPath))
                {
                    MessageBox.Show("请选择文件夹！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (!Directory.Exists(folderPath))
                {
                    MessageBox.Show("所选文件夹不存在，可能已被移动、删除，或当前不可访问。请重新选择有效文件夹。", "提示",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtFolderPath.Focus();
                    txtFolderPath.SelectAll();
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
                bool useFolderNameAsDescription = optUseFolderName.Checked;
                bool includeRootImages = chkIncludeRoot.Checked;
                bool includeSubFolderImages = chkIncludeSubFolder.Checked;
                bool needAutoNumbering = chkAutoNumbering.Checked;
                int numberAlignment = optAlignCenter.Checked ? 2 : 1;
                int numberPosition = optNumberAfterDesc.Checked ? 2 : 1;

                PendingRequest = new InsertPhotosRequest
                {
                    Mode = InsertPhotosRequestMode.Folder,
                    FolderPath = folderPath,
                    MinHeight = minHeight,
                    NeedDescription = needDescription,
                    UseFileNameAsDescription = useFileNameAsDescription,
                    UseFolderNameAsDescription = useFolderNameAsDescription,
                    IncludeRootImages = includeRootImages,
                    IncludeSubFolderImages = includeSubFolderImages,
                    NeedAutoNumbering = needAutoNumbering,
                    NumberAlignment = numberAlignment,
                    NumberPosition = numberPosition
                };

                DialogResult = DialogResult.OK;
                Close();
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
                string[] selectedFiles = null;
                if (InsertPhotosAutomationGate.IsEnabled
                    && InsertPhotosAutomationGate.TryGetPresetSelectedFiles(out var presetFiles))
                {
                    selectedFiles = presetFiles;
                }
                else
                {
                    string lastPath = ConfigService.GetLastFolderPath();
                    selectedFiles = SelectImageFiles("请选择图片文件...", lastPath);
                }

                if (selectedFiles == null || selectedFiles.Length == 0)
                {
                    return;
                }

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
                bool useFolderNameAsDescription = optUseFolderName.Checked;
                bool needAutoNumbering = chkAutoNumbering.Checked;
                int numberAlignment = optAlignCenter.Checked ? 2 : 1;
                int numberPosition = optNumberAfterDesc.Checked ? 2 : 1;

                PendingRequest = new InsertPhotosRequest
                {
                    Mode = InsertPhotosRequestMode.SelectedFiles,
                    SelectedFiles = selectedFiles,
                    MinHeight = minHeight,
                    NeedDescription = needDescription,
                    UseFileNameAsDescription = useFileNameAsDescription,
                    UseFolderNameAsDescription = useFolderNameAsDescription,
                    NeedAutoNumbering = needAutoNumbering,
                    NumberAlignment = numberAlignment,
                    NumberPosition = numberPosition
                };

                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("操作失败: {0}", ex.Message), "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string SelectFolder(string dialogTitle = "请选择文件夹...", string initialPath = "")
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = dialogTitle;
                dialog.ShowNewFolderButton = false;

                if (!string.IsNullOrEmpty(initialPath) && Directory.Exists(initialPath))
                {
                    dialog.SelectedPath = initialPath;
                }

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    return dialog.SelectedPath;
                }
            }
            return string.Empty;
        }

        private string[] SelectImageFiles(string dialogTitle = "请选择图片文件...", string initialPath = "")
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Title = dialogTitle;
                dialog.Multiselect = true;
                dialog.Filter = "图片文件|*.jpg;*.jpeg;*.png|所有文件|*.*";
                dialog.FilterIndex = 1;

                if (!string.IsNullOrEmpty(initialPath) && Directory.Exists(initialPath))
                {
                    dialog.InitialDirectory = initialPath;
                }

                if (dialog.ShowDialog() == DialogResult.OK && dialog.FileNames.Length > 0)
                {
                    return dialog.FileNames;
                }
            }
            return null;
        }

        #endregion
    }
}

