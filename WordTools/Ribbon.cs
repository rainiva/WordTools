using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace WordTools
{
    [ComVisible(true)]
    public partial class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon()
        {
        }

        #region IRibbonExtensibility 成员

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("WordTools.Ribbon.xml");
        }

        #endregion

        #region Ribbon 回调

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void OnHelloButtonClick(Office.IRibbonControl control)
        {
            try
            {
                MessageBox.Show("你好！欢迎使用 Word 插件演示程序！", "Hello", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("错误: {0}", ex.Message), "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnInsertTextClick(Office.IRibbonControl control)
        {
            try
            {
                if (Globals.ThisAddIn == null || Globals.ThisAddIn.Application == null || Globals.ThisAddIn.Application.ActiveDocument == null)
                {
                    MessageBox.Show("请先打开一个 Word 文档", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                var doc = Globals.ThisAddIn.Application.ActiveDocument;
                var selection = Globals.ThisAddIn.Application.Selection;
                selection.TypeText("【这是由插件插入的文本】");
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("插入文本失败: {0}", ex.Message), "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnShowInfoClick(Office.IRibbonControl control)
        {
            try
            {
                if (Globals.ThisAddIn == null || Globals.ThisAddIn.Application == null || Globals.ThisAddIn.Application.ActiveDocument == null)
                {
                    MessageBox.Show("请先打开一个 Word 文档", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                var app = Globals.ThisAddIn.Application;
                var doc = app.ActiveDocument;
                string info = string.Format("文档名称: {0}\n文档路径: {1}\n段落数: {2}\n字数: {3}",
                    doc.Name,
                    doc.Path,
                    doc.Paragraphs.Count,
                    doc.ComputeStatistics(Microsoft.Office.Interop.Word.WdStatistic.wdStatisticWords));
                MessageBox.Show(info, "文档信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("获取文档信息失败: {0}", ex.Message), "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnAboutClick(Office.IRibbonControl control)
        {
            try
            {
                MessageBox.Show(
                    "Word工具箱 v1.0\n\n" +
                    "功能：批量插图、Excel数据填充、自动编号\n\n" +
                    "© 2026 WordTools",
                    "关于",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("错误: {0}", ex.Message), "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnExcelDataFillerClick(Office.IRibbonControl control)
        {
            try
            {
                // 打开Excel数据填充窗体
                Forms.ExcelDataFillerForm form = new Forms.ExcelDataFillerForm();
                form.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("打开Excel数据填充工具失败: {0}", ex.Message), "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public string GetLabel(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "tabWordToolbox": return "Word工具箱";
                case "grpMainFunctions": return "图片工具";
                case "grpTools": return "工具";
                case "grpHelp": return "帮助";
                case "btnInsertPhotos": return "批量插图";
                case "btnExcelDataFiller": return "Excel数据填充";
                case "btnAbout": return "关于";
                default: return "";
            }
        }

        public string GetDescription(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "btnInsertPhotos": return "批量插入图片到表格";
                case "btnExcelDataFiller": return "Excel数据批量填充到Word表格";
                case "btnAbout": return "显示版本信息和使用说明";
                default: return "";
            }
        }

        public string GetSupertip(Office.IRibbonControl control)
        {
            return GetDescription(control);
        }

        public string GetScreentip(Office.IRibbonControl control)
        {
            return GetLabel(control);
        }

        #endregion

        #region 帮助程序

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
