using System;
using Microsoft.Office.Interop.Word;

namespace WordTools.Services.Abstractions
{
    /// <summary>
    /// 表格编号管理抽象接口
    /// 包含编号刷新、清除、添加等编号相关方法
    /// </summary>
    /// <remarks>规划抽象（Phase 2）。运行时仍使用 static TableNumberingService，尚无 Adapter 实现。</remarks>
    public interface ITableNumberingService
    {
        void RefreshTableNumbering(Table tbl, Document doc, int alignment = 2,
            Action<string> progressCallback = null);

        int ClearTableNumbering(Table tbl, int startRow = 1,
            Action<string> progressCallback = null);

        void AddNumberingToDescriptionRows(Table tbl, Document doc,
            int startRow = 1, int alignment = 1, bool needAutoNumbering = false,
            Action<string> progressCallback = null);
    }
}
