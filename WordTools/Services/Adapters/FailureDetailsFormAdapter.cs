using System.Collections.Generic;
using System.Windows.Forms;
using WordTools.Forms;
using WordTools.Services.Abstractions;

namespace WordTools.Services.Adapters
{
    /// <summary>
    /// IFailureDetailsPresenter 的 WinForms 实现，封装 FailureDetailsForm。
    /// </summary>
    public sealed class FailureDetailsFormAdapter : IFailureDetailsPresenter
    {
        public bool ShowDetails(
            string summary,
            List<(string fileName, string errorReason)> failures,
            List<int> mergedCellRows,
            List<string> overwriteWarnings)
        {
            using (var form = new FailureDetailsForm(
                failures ?? new List<(string, string)>(),
                mergedCellRows ?? new List<int>(),
                overwriteWarnings ?? new List<string>()))
            {
                // FailureDetailsForm 通过 DialogResult 区分“复制详情”(Yes) 与“关闭”(No/Cancel)
                return form.ShowDialog() == DialogResult.Yes;
            }
        }
    }
}
