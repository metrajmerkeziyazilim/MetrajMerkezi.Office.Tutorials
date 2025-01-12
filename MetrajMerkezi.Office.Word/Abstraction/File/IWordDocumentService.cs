using Microsoft.Office.Interop.Word;

namespace MetrajMerkezi.Office.Word.Abstraction.File
{
    public class WordDocumentResult
    {
        public Application Application { get; set; }
        public Document Document { get; set; }
    }

    public interface IWordDocumentService
    {
        WordDocumentResult CreateDocument();
        void SaveDocument(Document doc, string filePath);
        void SaveDocumentWithAutoIncrement(IWordDocumentService wordDocumentService, Document doc, string folderPath, string baseFileName);
    }
}
