using Word = Microsoft.Office.Interop.Word;
using WordTools.Services.Abstractions;

namespace WordTools.Services.Adapters
{
    /// <summary>
    /// IDocumentContext 的 Word COM 实现。
    /// </summary>
    public sealed class WordDocumentContext : IDocumentContext
    {
        private readonly IWordApplicationContext _appContext;

        public WordDocumentContext(IWordApplicationContext appContext)
        {
            _appContext = appContext;
        }

        public bool HasActiveDocument => _appContext.Application?.ActiveDocument != null;

        public Word.Document ActiveDocument => _appContext.Application?.ActiveDocument;
    }
}
