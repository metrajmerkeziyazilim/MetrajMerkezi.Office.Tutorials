using Microsoft.Office.Interop.Word;

namespace MetrajMerkezi.Office.Word.Abstraction.Object._Table
{
    public interface IWordTableService
    {
        Table CreateTable(Document doc, int rows, int columns);
    }
}
