using System.Collections.Generic;

namespace WordTools.Services.Abstractions
{
    /// <summary>
    /// 失败/警告详情弹窗抽象。返回 true 表示用户请求复制详情。
    /// </summary>
    public interface IFailureDetailsPresenter
    {
        bool ShowDetails(
            string summary,
            List<(string fileName, string errorReason)> failures,
            List<int> mergedCellRows,
            List<string> overwriteWarnings);
    }
}
