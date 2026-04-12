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

        public void OnRefreshNumberingClick(Office.IRibbonControl control)
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
                var selection = app.Selection;

                // 检查是否在表格中
                if (!Services.TableService.IsSelectionInTable(selection))
                {
                    MessageBox.Show("请先将光标放在需要刷新编号的表格中", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 获取当前表格
                var tbl = Services.TableService.GetCurrentTable(selection);
                if (tbl == null)
                {
                    MessageBox.Show("无法获取当前表格", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 刷新表格编号（带进度提示）
                app.StatusBar = "正在刷新编号...";
                System.Windows.Forms.Application.DoEvents();

                Services.TableService.RefreshTableNumbering(tbl, doc, 2, (msg) =>
                {
                    try
                    {
                        app.StatusBar = msg;
                        System.Windows.Forms.Application.DoEvents();
                    }
                    catch { }
                });

                // 确保域代码不可见
                try
                {
                    if (doc.ActiveWindow.View.ShowFieldCodes)
                        doc.ActiveWindow.View.ShowFieldCodes = false;
                }
                catch { }

                app.StatusBar = "";
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
                    using (var stream = asm.GetManifestResourceStream(resourceNames[i]))
                    using (var resourceReader = new StreamReader(stream))
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
