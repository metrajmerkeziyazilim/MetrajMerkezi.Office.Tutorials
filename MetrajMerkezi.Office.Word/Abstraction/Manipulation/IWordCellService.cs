using MetrajMerkezi.Office.Word.Enum;
using Microsoft.Office.Interop.Word;

namespace MetrajMerkezi.Office.Word.Abstraction.Manipulation
{
    public interface IWordCellService
    {
        void MergeCells(Table table, int startRow, int startColumn, int endRow, int endColumn);
        void MergeAndSetTextToSpesicifCell(Table contentTable,
            int startRow, int startColumn,
            int endRow, int endColumn,
            int cellRow, int cellColumn,
            string cellContent, float fontSize,
            HorizontalTextAlignment HtextAlignment,
            VerticalTextAlignment VtextAlignment,
            bool isBold = true);

        void SetTextToCell(Table table, int cellRow, int cellColumn, string text, float fontSize, bool isBold, WdParagraphAlignment alignment);
        void InsertImageToCell(Table table, int cellRow, int cellColumn, string imagePath, float maxWidth, float maxHeight);

    }
}
