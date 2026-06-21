using System;
using System.IO;
using System.Reflection;
using Word = Microsoft.Office.Interop.Word;
using Extensibility;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using WordTools.Services;
using WordTools.Interop;

namespace WordTools
{
    [ComVisible(true)]
    [Guid("A1B2C3D4-E5F6-7890-ABCD-EF1234567890")]
    [ProgId("WordTools.ThisAddIn")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public partial class ThisAddIn : IDTExtensibility2, Office.IRibbonExtensibility, IRequestComAddInAutomationService
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

            if (AddInInst is Office.COMAddIn comAddIn)
            {
                comAddIn.Object = this;
            }

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

        public object GetComAddInAutomationService()
        {
            return this;
        }

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
            new InsertPhotosOrchestrator(Globals.Application).ShowFormAndExecuteIfConfirmed();
        }

        public void Automation_ExecuteFromConfig()
        {
            new InsertPhotosOrchestrator(Globals.Application).ExecuteFromAutomationConfig();
        }

        /// <summary>
        /// UI 自动化入口：等价于 Ribbon「批量插图」，需 WORDTOOLS_UI_AUTOMATION=1。
        /// </summary>
        public void Automation_ShowInsertPhotosForm()
        {
            InsertPhotosAutomationGate.EnsureEnabled();
            new InsertPhotosOrchestrator(Globals.Application).ShowFormAndExecuteIfConfirmed();
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
            var appContext = new Services.Adapters.WordApplicationContext(Globals.Application);
            var notificationService = new Services.Adapters.MessageBoxNotificationService();
            var numberingRefreshService = new Services.NumberingRefreshService(appContext, notificationService);
            numberingRefreshService.RefreshFromCurrentSelection();
        }

        /// <summary>
        /// Excel数据填充按钮点击
        /// </summary>
        public void OnExcelDataFillerClick(Office.IRibbonControl control)
        {
            new ExcelDataFillerOrchestrator(Globals.Application).ShowForm();
        }

        #endregion

        #endregion
    }
}
