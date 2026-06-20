using Word = Microsoft.Office.Interop.Word;

namespace WordTools.Services.Abstractions
{
    /// <summary>
    /// Word Application 上下文抽象。封装屏幕更新、状态栏、消息泵等全局状态操作。
    /// </summary>
    public interface IWordApplicationContext
    {
        Word.Application Application { get; }
        bool ScreenUpdating { get; set; }
        void SetStatusBar(string text);
        void DoEvents();
    }
}
