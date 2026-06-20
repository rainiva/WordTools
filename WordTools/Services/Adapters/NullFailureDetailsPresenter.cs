using System.Collections.Generic;
using WordTools.Services.Abstractions;

namespace WordTools.Services.Adapters
{
    /// <summary>
    /// 不弹窗的失败详情展示器，供无人值守 E2E 使用。
    /// </summary>
    public sealed class NullFailureDetailsPresenter : IFailureDetailsPresenter
    {
        public bool ShowDetails(
            string summary,
            List<(string fileName, string errorReason)> failures,
            List<int> mergedCellRows,
            List<string> overwriteWarnings)
        {
            return false;
        }
    }
}
