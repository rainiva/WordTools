using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Extensibility;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using WordTools.Forms;
using WordTools.Services;

namespace WordTools
{
    [ComVisible(true)]
    [Guid("A1B2C3D4-E5F6-7890-ABCD-EF1234567890")]
    [ProgId("WordTools.ThisAddIn")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public partial class ThisAddIn : IDTExtensibility2, Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbonUI;
        private readonly RibbonController _ribbonController = new RibbonController();
        private static readonly string _ribbonXml = LoadRibbonXml();

        /// <summary>
        /// 从嵌入资源中加载并缓存 Ribbon XML
        /// </summary>
        private static string LoadRibbonXml()
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare("WordTools.Ribbon.xml", resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (var stream = asm.GetManifestResourceStream(resourceNames[i]))
                    using (var reader = new StreamReader(stream))
                    {
                        return reader.ReadToEnd();
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// COM 类必须有无参数构造函数
        /// </summary>
        public ThisAddIn()
        {
        }
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region IDTExtensibility2 Members

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref System.Array custom)
        {
            this.Application = (Word.Application)Application;

            if (Globals.ThisAddIn == null)
            {
                Globals.ThisAddIn = this;
            }
            if (Globals.Application == null)
            {
                Globals.Application = this.Application;
            }

            if (ConnectMode == ext_ConnectMode.ext_cm_Startup || ConnectMode == ext_ConnectMode.ext_cm_External)
            {
                ThisAddIn_Startup(this, EventArgs.Empty);
            }
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref System.Array custom)
        {
            ThisAddIn_Shutdown(this, EventArgs.Empty);
        }

        public void OnAddInsUpdate(ref System.Array custom)
        {
        }

        public void OnStartupComplete(ref System.Array custom)
        {
        }

        public void OnBeginShutdown(ref System.Array custom)
        {
        }

        #endregion

        #region IRibbonExtensibility Members

        public string GetCustomUI(string RibbonID)
        {
            return _ribbonXml;
        }

        #endregion

        #region Ribbon 回调方法 - 直接在此类中实现

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbonUI = ribbonUI;
            _ribbonController.OnRibbonLoad(ribbonUI);
        }

        public void InvalidateRibbon()
        {
            _ribbonController.InvalidateRibbon();
        }

        public bool GetDetailedLoggingPressed(Office.IRibbonControl control)
        {
            return _ribbonController.GetDetailedLoggingPressed(control);
        }

        public bool GetBenchmarkLoggingPressed(Office.IRibbonControl control)
        {
            return _ribbonController.GetBenchmarkLoggingPressed(control);
        }

        public bool GetBenchmarkLoggingEnabled(Office.IRibbonControl control)
        {
            return _ribbonController.GetBenchmarkLoggingEnabled(control);
        }

        #region 按钮点击回调

        /// <summary>
        /// 批量插图按钮点击
        /// </summary>
        public void OnInsertPhotosClick(Office.IRibbonControl control)
        {
            ShowInsertPhotosForm();
        }

        public void OnShowLoggingSettingsSummary(Office.IRibbonControl control)
        {
            _ribbonController.OnShowLoggingSettingsSummary(control);
        }

        public void OnToggleDetailedLogging(Office.IRibbonControl control, bool pressed)
        {
            _ribbonController.OnToggleDetailedLogging(control, pressed);
        }

        public void OnToggleBenchmarkLogging(Office.IRibbonControl control, bool pressed)
        {
            _ribbonController.OnToggleBenchmarkLogging(control, pressed);
        }

        /// <summary>
        /// 关于按钮点击
        /// </summary>
        public void OnAboutClick(Office.IRibbonControl control)
        {
            _ribbonController.OnAboutClick(control);
        }

        /// <summary>
        /// 刷新编号按钮点击
        /// </summary>
        public void OnRefreshNumberingClick(Office.IRibbonControl control)
        {
            try
            {
                if (Application == null || Application.ActiveDocument == null)
                {
                    MessageBox.Show("请先打开一个 Word 文档", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var doc = Application.ActiveDocument;
                var selection = Application.Selection;

                if (!Services.TableService.IsSelectionInTable(selection))
                {
                    MessageBox.Show("请先将光标放在需要刷新编号的表格中", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var tbl = Services.TableService.GetCurrentTable(selection);
                if (tbl == null)
                {
                    MessageBox.Show("无法获取当前表格", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 刷新表格编号（带进度提示，ScreenUpdating 由 TableService 内部控制）
                Application.StatusBar = "正在刷新编号...";
                System.Windows.Forms.Application.DoEvents();

                Services.TableNumberingService.RefreshTableNumbering(tbl, doc, 2, (status) =>
                {
                    try
                    {
                        Application.StatusBar = status;
                        System.Windows.Forms.Application.DoEvents();
                    }
                    catch (COMException)
                    {
                        // Word 可能处于繁忙或关闭状态，忽略状态栏更新失败
                    }
                });

                // 确保域代码不可见
                try
                {
                    if (doc.ActiveWindow.View.ShowFieldCodes)
                        doc.ActiveWindow.View.ShowFieldCodes = false;
                }
                catch (COMException)
                {
                    // 视图对象可能暂时不可用，不影响编号结果
                }

                Application.StatusBar = "";
                // 确保UI完全刷新后再显示提示框
                // 使用后台线程延迟，避免阻塞UI线程
                var t = new System.Threading.Thread(() =>
                {
                    System.Threading.Thread.Sleep(300);
                });
                t.Start();
                while (t.IsAlive)
                {
                    System.Windows.Forms.Application.DoEvents();
                    System.Threading.Thread.Sleep(10);
                }
                MessageBox.Show("表格编号已刷新完成！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("刷新编号失败: {0}", ex.Message), "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Excel数据填充按钮点击
        /// </summary>
        public void OnExcelDataFillerClick(Office.IRibbonControl control)
        {
            try
            {
                var appContext = new Services.Adapters.WordApplicationContext(Globals.Application);
                var documentContext = new Services.Adapters.WordDocumentContext(appContext);
                var notificationService = new Services.Adapters.MessageBoxNotificationService();
                using (var form = new Forms.ExcelDataFillerForm(documentContext, notificationService))
                {
                    form.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("打开Excel数据填充工具失败: {0}", ex.Message),
                               "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region 窗体显示

        /// <summary>
        /// 显示批量插图主窗体
        /// </summary>
        public void ShowInsertPhotosForm()
        {
            try
            {
                InsertPhotosRequest pendingRequest = null;
                using (var form = new InsertPhotosForm(Globals.Application))
                {
                    if (form.ShowDialog() == DialogResult.OK)
                    {
                        pendingRequest = form.PendingRequest;
                    }
                }

                if (pendingRequest != null)
                {
                    ExecuteInsertPhotosRequestDeferred(pendingRequest);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "打开窗体失败: " + ex.Message,
                    "错误",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void ExecuteInsertPhotosRequest(InsertPhotosRequest request)
        {
            if (request == null)
            {
                return;
            }

            var appContext = new Services.Adapters.WordApplicationContext(Globals.Application);
            var notificationService = new Services.Adapters.MessageBoxNotificationService();
            var failureDetailsPresenter = new Services.Adapters.FailureDetailsFormAdapter();

            int totalFiles = request.Mode == InsertPhotosRequestMode.Folder
                ? Services.FileService.CountTotalImageFiles(
                    request.FolderPath,
                    request.IncludeRootImages,
                    request.IncludeSubFolderImages)
                : (request.SelectedFiles?.Length ?? 0);

            var progressReporter = new Services.Adapters.ProgressFormAdapter(totalFiles);
            var progressService = new Services.ProgressService(
                appContext,
                progressReporter,
                failureDetailsPresenter,
                notificationService);

            if (request.Mode == InsertPhotosRequestMode.Folder)
            {
                progressService.InsertPhotosWithProgress(
                    request.FolderPath,
                    request.MinHeight,
                    request.NeedDescription,
                    request.UseFileNameAsDescription,
                    request.UseFolderNameAsDescription,
                    request.IncludeRootImages,
                    request.IncludeSubFolderImages,
                    request.NeedAutoNumbering,
                    request.NumberAlignment,
                    request.NumberPosition);
                return;
            }

            progressService.InsertSelectedPhotosWithProgress(
                request.SelectedFiles,
                request.MinHeight,
                request.NeedDescription,
                request.UseFileNameAsDescription,
                request.UseFolderNameAsDescription,
                request.NeedAutoNumbering,
                request.NumberAlignment,
                request.NumberPosition);
        }

        private void ExecuteInsertPhotosRequestDeferred(InsertPhotosRequest request)
        {
            if (request == null)
            {
                return;
            }

            // 使用一次性 Timer 在消息循环的下一周期调度插入操作。
            // 这样可以确保模态对话框已完全关闭，避免在对话框句柄仍有效时操作 Word。
            var timer = new System.Windows.Forms.Timer { Interval = 1 };
            timer.Tick += (sender, e) =>
            {
                timer.Stop();
                timer.Dispose();

                try
                {
                    ExecuteInsertPhotosRequest(request);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        "批量插图失败: " + ex.Message,
                        "错误",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            };

            timer.Start();
        }

        #endregion

        #endregion
    }
}
