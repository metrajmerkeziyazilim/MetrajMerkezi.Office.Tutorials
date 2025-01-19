using MetrajMerkezi.Office.Word.Abstraction.Manipulation;
using MetrajMerkezi.Office.Word.Enum;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;

namespace MetrajMerkezi.Office.Word.Services.Manipulation
{
    public class WordCellService : IWordCellService
    {

        public void MergeCells(Table table,
        int startRow,
        int startColumn,
        int endRow,
        int endColumn)
        {
            var startCell = table.Cell(startRow, startColumn);
            var endCell = table.Cell(endRow, endColumn);
            startCell.Merge(endCell);
        }
        public void MergeAndSetTextToSpesicifCell(Table contentTable,
            int startRow, int startColumn,
            int endRow, int endColumn,
            int cellRow, int cellColumn,
            string cellContent, float fontSize,
            HorizontalTextAlignment HtextAlignment,
            VerticalTextAlignment VtextAlignment,
            bool isBold = true)
        {
            MergeCells(contentTable, startRow, startColumn, endRow, endColumn);
            Cell mergedCell=contentTable.Cell(cellRow,cellColumn);

            WdParagraphAlignment horizontalAlignment= HtextAlignment switch
            {
                HorizontalTextAlignment.Left=>WdParagraphAlignment.wdAlignParagraphLeft,
                HorizontalTextAlignment.Center=>WdParagraphAlignment.wdAlignParagraphCenter,
                HorizontalTextAlignment.Right=>WdParagraphAlignment.wdAlignParagraphRight,
                HorizontalTextAlignment.Distribute=>WdParagraphAlignment.wdAlignParagraphDistribute,
                _=> WdParagraphAlignment.wdAlignParagraphCenter,
            };

            WdCellVerticalAlignment verticalAlignmnet= VtextAlignment switch
            {
                VerticalTextAlignment.Top=> WdCellVerticalAlignment.wdCellAlignVerticalTop,
                VerticalTextAlignment.Center=> WdCellVerticalAlignment.wdCellAlignVerticalCenter,
                VerticalTextAlignment.Bottom=> WdCellVerticalAlignment.wdCellAlignVerticalBottom,
                _=> WdCellVerticalAlignment.wdCellAlignVerticalCenter
            };

            var cell=contentTable.Cell(cellRow,cellColumn);
            var cellRange=cell.Range;

            cellRange.Text = cellContent;
            cellRange.Font.Bold = isBold ? 1 : 0;
            cellRange.Font.Size = fontSize;
            cellRange.Font.Name = "Arial";

            cellRange.ParagraphFormat.Alignment = horizontalAlignment;

            cell.VerticalAlignment = verticalAlignmnet;

        }

        public void SetTextToCell(Table table,
            int cellRow, int cellColumn,
            string text, float fontSize,
            bool isBold, WdParagraphAlignment alignment)
        {
            Cell cell=table.Cell(cellRow,cellColumn);
            cell.Range.Text = string.Empty;
            cell.Range.Text = text;

            Range cellRange=cell.Range;

            cellRange.Font.Bold = isBold ? 1 : 0;
            cellRange.ParagraphFormat.Alignment = alignment;

            cellRange.Font.Name = "Arial";
            cellRange.Font.Size = fontSize;
        }

        public void InsertImageToCell(Table table, int cellRow, int cellColumn, string imagePath, float maxWidth, float maxHeight)
        {
            Cell cell=table.Cell(cellRow,cellColumn);

            var cellRange=cell.Range;

            InlineShape inlineShape=cellRange.InlineShapes.AddPicture(
                imagePath,
                false,
                true);


            float originalWidth=inlineShape.Width;
            float originalHeight=inlineShape.Height;

            float aspectRatio=originalWidth/originalHeight;

            float cellWidth=cell.Width;
            float cellHeight=cell.Height;

            if (originalWidth > cellWidth || originalHeight > cellHeight)
            {
                if (aspectRatio > 1)
                {
                    inlineShape.Width = cellWidth;
                    inlineShape.Height = cellWidth / aspectRatio;

                    if (inlineShape.Height > cellHeight)
                    {
                        inlineShape.Height = cellHeight;
                        inlineShape.Width = cellHeight * aspectRatio;
                    }
                }
                else
                {
                    inlineShape.Height = cellHeight;
                    inlineShape.Width = cellHeight * aspectRatio;

                    if (inlineShape.Width > cellWidth)
                    {
                        inlineShape.Width = cellWidth;
                        inlineShape.Height = cellWidth / aspectRatio;
                    }
                }
            }

            cell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            cellRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

        }
    }
}
