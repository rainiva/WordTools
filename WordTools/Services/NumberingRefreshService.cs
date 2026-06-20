using System;
using System.Runtime.InteropServices;
using System.Threading;
using WordTools.Services.Abstractions;

namespace WordTools.Services
{
    /// <summary>
    /// Ribbon「刷新编号」操作的编排服务。
    /// </summary>
    public sealed class NumberingRefreshService
    {
        private readonly IWordApplicationContext _appContext;
        private readonly INotificationService _notification;

        public NumberingRefreshService(IWordApplicationContext appContext, INotificationService notification)
        {
            _appContext = appContext ?? throw new ArgumentNullException(nameof(appContext));
            _notification = notification ?? throw new ArgumentNullException(nameof(notification));
        }

        public void RefreshFromCurrentSelection()
        {
            try
            {
                var application = _appContext.Application;
                if (application == null || application.ActiveDocument == null)
                {
                    _notification.ShowWarning("请先打开一个 Word 文档", "提示");
                    return;
                }

                var doc = application.ActiveDocument;
                var selection = application.Selection;

                if (!TableService.IsSelectionInTable(selection))
                {
                    _notification.ShowWarning("请先将光标放在需要刷新编号的表格中", "提示");
                    return;
                }

                var tbl = TableService.GetCurrentTable(selection);
                if (tbl == null)
                {
                    _notification.ShowError("无法获取当前表格", "错误");
                    return;
                }

                _appContext.SetStatusBar("正在刷新编号...");
                _appContext.DoEvents();

                TableNumberingService.RefreshTableNumbering(tbl, doc, 2, status =>
                {
                    try
                    {
                        _appContext.SetStatusBar(status);
                        _appContext.DoEvents();
                    }
                    catch (COMException)
                    {
                        // Word 可能处于繁忙或关闭状态，忽略状态栏更新失败
                    }
                });

                try
                {
                    if (doc.ActiveWindow.View.ShowFieldCodes)
                    {
                        doc.ActiveWindow.View.ShowFieldCodes = false;
                    }
                }
                catch (COMException)
                {
                    // 视图对象可能暂时不可用，不影响编号结果
                }

                _appContext.SetStatusBar("");
                WaitForUiSettle();
                _notification.ShowInformation("表格编号已刷新完成！", "成功");
            }
            catch (Exception ex)
            {
                _notification.ShowError(string.Format("刷新编号失败: {0}", ex.Message), "错误");
            }
        }

        private void WaitForUiSettle()
        {
            var t = new Thread(() => Thread.Sleep(300));
            t.Start();
            while (t.IsAlive)
            {
                _appContext.DoEvents();
                Thread.Sleep(10);
            }
        }
    }
}
