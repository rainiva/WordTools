using Word = Microsoft.Office.Interop.Word;

namespace WordTools.Services.Abstractions
{
    /// <summary>
    /// 当前 Word 文档上下文抽象。替代服务层读取 Globals.ThisAddIn.Application。
    /// </summary>
    public interface IDocumentContext
    {
        bool HasActiveDocument { get; }
        Word.Document ActiveDocument { get; }
    }
}
