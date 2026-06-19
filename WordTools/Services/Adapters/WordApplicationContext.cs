using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using WordTools.Services.Abstractions;

namespace WordTools.Services.Adapters
{
    /// <summary>
    /// IWordApplicationContext 的 Word COM 实现。
    /// </summary>
    public sealed class WordApplicationContext : IWordApplicationContext
    {
        public WordApplicationContext(Word.Application application)
        {
            Application = application;
        }

        public Word.Application Application { get; }

        public bool ScreenUpdating
        {
            get => Application.ScreenUpdating;
            set => Application.ScreenUpdating = value;
        }

        public void SetStatusBar(string text)
        {
            Application.StatusBar = text;
        }

        public void DoEvents()
        {
            System.Windows.Forms.Application.DoEvents();
        }
    }
}
