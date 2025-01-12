using MetrajMerkezi.Office.Word.Abstraction.Object._Table;
using Microsoft.Office.Interop.Word;

namespace MetrajMerkezi.Office.Word.Services.Object._Table
{
    public class WordTableService : IWordTableService
    {
        public Table CreateTable(Document doc, int rows, int columns)
        {
            var table = doc.Tables.Add(doc.Range(),rows,columns);

            table.Borders.Enable = 1;

            table.Rows.Alignment = WdRowAlignment.wdAlignRowLeft;

            return table;
        }
    }
}
