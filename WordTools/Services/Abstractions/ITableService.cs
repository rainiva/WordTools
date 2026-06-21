using System.Collections.Generic;
using Microsoft.Office.Interop.Word;

namespace WordTools.Services.Abstractions
{
    /// <summary>
    /// 表格结构操作抽象接口
    /// 包含表格验证、单元格可用性检查、行列操作等结构相关方法
    /// </summary>
    /// <remarks>规划抽象（Phase 2）。运行时仍使用 static TableService，尚无 Adapter 实现。</remarks>
    public interface ITableService
    {
        // 表格验证
        bool IsSelectionInTable(Selection selection);
        bool IsSelectionInFirstColumn(Selection selection);
        Table GetCurrentTable(Selection selection);

        // 单元格可用性检查
        bool IsCellSuitableForImage(Cell targetCell);
        ImageCellAvailability GetCellAvailability(Cell targetCell);
        ImageCellAvailability GetCellAvailability(Cell targetCell, ImageInsertionBatchContext context);

        // 查找适合插入的位置
        bool FindNextSuitableCell(Table tbl, int startRow, out int foundRow, out int foundCol, int preferredCol = 1);
        ImageRowAvailability GetImageRowAvailability(Table tbl, int row);
        ImageRowAvailability GetImageRowAvailability(Table tbl, int row, ImageInsertionBatchContext context);
        bool FindNextSuitableImageRow(Table tbl, int startRow, out int foundRow, List<int> mergedCellRows = null);
        bool FindNextSuitableImageRow(Table tbl, int startRow, out int foundRow, List<int> mergedCellRows, ImageInsertionBatchContext context);

        // 表格结构操作
        bool ShouldTreatCellCountMismatchAsMerged(int requestedColumn, int visibleCellCount, int totalColumnCount);
        int GetImageRowSearchEndRow(int startRow, int lastExistingRow);
        void EnsureRowExists(Table tbl, int rowIndex, ref int cachedRowCount);
        void EnsureRowExists(Table tbl, int rowIndex);
        void AdjustTableColumns(Table tbl, int targetColCount);
        bool IsTableFixedColumnWidth(Table tbl);
        void SetTableFixedColumnWidth(Table tbl);

        // 标题行和描述行
        void CreateTitleRow(Table tbl, ref int rowIndex, string titleText);
        void InsertDescriptionRow(Table tbl, ref int rowIndex);
        void InsertFileNameDescriptionRow(Table tbl, ref int rowIndex, string[] descriptions, bool isFilePath = true);
        void FillEmptyCellsWithNA(Table tbl);
    }
}
