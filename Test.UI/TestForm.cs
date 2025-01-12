using MetrajMerkezi.Office.Word.Abstraction.File;
using MetrajMerkezi.Office.Word.Abstraction.Object._Table;
using MetrajMerkezi.Office.Word.DependencyResolver;
using Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;

namespace Test.UI
{
    public partial class TestForm : Form
    {
        private readonly IWordDocumentService _wordDocumentService;
        private readonly IWordTableService _wordTableService;

        public TestForm()
        {
            InitializeComponent();
            _wordDocumentService = InstanceFactory.GetInstance<IWordDocumentService>();
            _wordTableService = InstanceFactory.GetInstance<IWordTableService>();
        }


        private void CreateDocumentBtn_Click(object sender, EventArgs e)
        {
            CreateTitlePage();
        }



        private void CreateTitlePage()
        {
            Microsoft.Office.Interop.Word.Application application=null;
            Document document=null;

            try
            {
                var result=_wordDocumentService.CreateDocument();
                application = result.Application;
                document = result.Document;

                int totalRows=16;
                int totalColumns=21;

                Table contentTable=_wordTableService.CreateTable(document,totalRows,totalColumns);

                string filePath=@"C:\Users\Metraj Merkezi\Desktop\WordYoutubeDenemeler";
                _wordDocumentService.SaveDocumentWithAutoIncrement(_wordDocumentService, document, filePath, "Antet_Deneme");
                MessageBox.Show("Word Dokümaný Baþarýyla Oluþturdu.");
            }
            finally
            {
                try
                {
                    if(document!=null)
                    {
                        Marshal.FinalReleaseComObject(document);
                        document = null;
                    }

                    try
                    {
                        if (application != null)
                        {
                            var version=application.Version;
                            application.Quit();
                            Marshal.FinalReleaseComObject(application);
                        }
                    }
                    catch (COMException)
                    {
      
                    }
                    finally
                    {
                        application = null;
                    }
                }
                catch (Exception ex)
                {

                    MessageBox.Show($"Hata Meydana Geldi! {ex.Message}");
                }
                finally
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
        }
    }
}
