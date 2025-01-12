using MetrajMerkezi.Office.Word.Abstraction.File;
using Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;

namespace MetrajMerkezi.Office.Word.Services.File
{
    public class WordDocumentService : IWordDocumentService
    {
        public WordDocumentResult CreateDocument()
        {
            var app= new Application();
            var doc= app.Documents.Add();
            return new WordDocumentResult { Application = app, Document = doc };
        }

        public void SaveDocument(Document doc, string filePath)
        {
            if (Directory.Exists(Path.GetDirectoryName(filePath)))
            {
                doc.SaveAs2(filePath);
            }
        }

        public void SaveDocumentWithAutoIncrement(IWordDocumentService wordDocumentService, Document doc, string folderPath, string baseFileName)
        {
            try
            {
                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                }

                int counter=1;
                string filePath;

                do
                {
                    filePath = Path.Combine(folderPath, $"{baseFileName}_{counter}.docx");
                    counter++;
                } while (System.IO.File.Exists(filePath));

                wordDocumentService.SaveDocument(doc, filePath);

            }
            finally
            {
                if(doc!=null)
                {
                    Marshal.FinalReleaseComObject(doc);
                    doc = null;
                }
            }
        }
    }
}
